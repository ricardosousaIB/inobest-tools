import streamlit as st
import io, zipfile, re, csv, os, time, requests, base64, hashlib, secrets
import xml.etree.ElementTree as ET
import pandas as pd
from typing import Tuple, List, Dict, Any
from io import BytesIO

st.set_page_config(layout="wide", page_title="Ferramentas de Ficheiros")

st.title("Ferramentas de Ficheiros")

# --- Função para o Agregador de Excel (Seu código atual) ---
def excel_aggregator_app():
    st.header("Agregador de Ficheiros Excel (via ZIP)")
    st.write("Carregue um ficheiro ZIP contendo os seus ficheiros Excel (.xls ou .xlsx) para agregá-los num único ficheiro CSV, que será depois comprimido e devolvido num novo ZIP.")
    st.write("Esta versão está otimizada para lidar com um grande número de ficheiros, processando-os de forma iterativa para poupar memória e gerando um CSV comprimido para maior velocidade.")

    # Inicializa o session_state para armazenar o resultado e evitar reprocessamento
    # Usamos chaves específicas para esta função para evitar conflitos com outras partes da app
    if 'excel_processed_data_zip' not in st.session_state:
        st.session_state.excel_processed_data_zip = None
    if 'excel_processed_data_csv_preview' not in st.session_state:
        st.session_state.excel_processed_data_csv_preview = None
    if 'excel_processing_done' not in st.session_state:
        st.session_state.excel_processing_done = False
    if 'excel_arquivos_com_erro_state' not in st.session_state:
        st.session_state.excel_arquivos_com_erro_state = []
    if 'excel_last_uploaded_file_id' not in st.session_state: # Para detetar novo upload
        st.session_state.excel_last_uploaded_file_id = None

    # Widget para carregamento de um único ficheiro ZIP
    uploaded_zip_file = st.file_uploader(
        "Arraste e largue o seu ficheiro ZIP aqui ou clique para procurar",
        type=["zip"],
        accept_multiple_files=False,
        key="zip_uploader_excel" # Chave única para este uploader
    )

    # Lógica para resetar o estado quando um NOVO ficheiro é carregado
    if uploaded_zip_file is not None:
        current_file_id = uploaded_zip_file.file_id # Streamlit 1.26+
        if current_file_id != st.session_state.excel_last_uploaded_file_id:
            st.session_state.excel_last_uploaded_file_id = current_file_id
            st.session_state.excel_processing_done = False
            st.session_state.excel_processed_data_zip = None
            st.session_state.excel_processed_data_csv_preview = None
            st.session_state.excel_arquivos_com_erro_state = []
            st.rerun() # Força um rerun para iniciar o processamento do novo ficheiro
    elif uploaded_zip_file is None and st.session_state.excel_last_uploaded_file_id is not None:
        # Se o uploader foi limpo (ficheiro removido), resetar tudo
        st.session_state.excel_last_uploaded_file_id = None
        st.session_state.excel_processing_done = False
        st.session_state.excel_processed_data_zip = None
        st.session_state.excel_processed_data_csv_preview = None
        st.session_state.excel_arquivos_com_erro_state = []
        st.rerun()


    # Se o processamento ainda não foi feito e há um ficheiro ZIP para processar
    if uploaded_zip_file is not None and not st.session_state.excel_processing_done:
        excel_files_in_zip = []
        arquivos_com_erro = []

        temp_csv_buffer = StringIO()
        header_written = False

        st.info(f"A iniciar o processamento do ficheiro ZIP: **{uploaded_zip_file.name}**...")

        try:
            with zipfile.ZipFile(uploaded_zip_file, 'r') as zf:
                for file_info in zf.infolist():
                    if not file_info.is_dir() and (file_info.filename.endswith('.xls') or file_info.filename.endswith('.xlsx')):
                        excel_files_in_zip.append(file_info.filename)

                if not excel_files_in_zip:
                    st.warning("Nenhum ficheiro Excel (.xls ou .xlsx) encontrado dentro do ficheiro ZIP. Por favor, verifique o conteúdo do ZIP.")
                    # Não usar st.stop() aqui, apenas resetar o estado para permitir novo upload
                    st.session_state.excel_processing_done = False
                    st.session_state.excel_processed_data_zip = None
                    st.session_state.excel_processed_data_csv_preview = None
                    st.session_state.excel_arquivos_com_erro_state = []
                    st.session_state.excel_last_uploaded_file_id = None # Resetar o ID para permitir novo upload
                    st.rerun() # Força um rerun para limpar a interface

                st.write(f"Encontrados **{len(excel_files_in_zip)}** ficheiro(s) Excel no ZIP. A iniciar leitura...")
                progress_text = st.empty()
                progress_bar = st.progress(0)

                for i, filename_in_zip in enumerate(excel_files_in_zip):
                    progress_text.text(f"A processar ficheiro {i+1}/{len(excel_files_in_zip)}: **{filename_in_zip}**")

                    try:
                        with zf.open(filename_in_zip) as excel_file_in_zip:
                            excel_content = BytesIO(excel_file_in_zip.read())

                            temp_dfs_from_file = []
                            excel_reader = pd.ExcelFile(excel_content)
                            for folha in excel_reader.sheet_names:
                                df_sheet = pd.read_excel(excel_content, sheet_name=folha)
                                temp_dfs_from_file.append(df_sheet)

                            if temp_dfs_from_file:
                                df_current_file = pd.concat(temp_dfs_from_file, ignore_index=True)

                                df_current_file.to_csv(temp_csv_buffer, sep=';', mode='a', header=not header_written, index=False, encoding='utf-8', quoting=csv.QUOTE_MINIMAL)
                                if not header_written:
                                    header_written = True
                            else:
                                st.warning(f"O ficheiro '{filename_in_zip}' não contém dados em nenhuma folha ou está vazio.")

                    except Exception as e:
                        arquivos_com_erro.append(f"{filename_in_zip} ({e})")
                        st.error(f"Erro ao processar '{filename_in_zip}' dentro do ZIP: {e}")

                    progress_bar.progress((i + 1) / len(excel_files_in_zip))

                progress_text.text("Todos os ficheiros foram processados. A finalizar...")

        except zipfile.BadZipFile:
            st.error("O ficheiro carregado não é um ficheiro ZIP válido ou está corrompido. Por favor, verifique e tente novamente.")
            # Resetar estado para permitir novo upload
            st.session_state.excel_processing_done = False
            st.session_state.excel_processed_data_zip = None
            st.session_state.excel_processed_data_csv_preview = None
            st.session_state.excel_arquivos_com_erro_state = []
            st.session_state.excel_last_uploaded_file_id = None
            st.rerun()
        except Exception as e:
            st.error(f"Ocorreu um erro inesperado ao processar o ficheiro ZIP: {e}")
            # Resetar estado para permitir novo upload
            st.session_state.excel_processing_done = False
            st.session_state.excel_processed_data_zip = None
            st.session_state.excel_processed_data_csv_preview = None
            st.session_state.excel_arquivos_com_erro_state = []
            st.session_state.excel_last_uploaded_file_id = None
            st.rerun()

        # Se algum dado foi processado, armazena no session_state
        if header_written:
            st.session_state.excel_processing_done = True
            st.session_state.excel_arquivos_com_erro_state = arquivos_com_erro

            with st.spinner("A carregar dados agregados e a preparar para download..."):
                temp_csv_buffer.seek(0)

                try:
                    df_final_preview = pd.read_csv(temp_csv_buffer, sep=';', encoding='utf-8')
                    st.session_state.excel_processed_data_csv_preview = df_final_preview.head()
                    temp_csv_buffer.seek(0)
                except Exception as e:
                    st.warning(f"Não foi possível gerar a pré-visualização dos dados agregados devido a um erro: {e}. O download ainda estará disponível.")
                    st.session_state.excel_processed_data_csv_preview = None
                    temp_csv_buffer.seek(0)

            with st.spinner("A comprimir o CSV num ficheiro ZIP..."):
                zip_output_buffer = BytesIO()
                with zipfile.ZipFile(zip_output_buffer, 'w', zipfile.ZIP_DEFLATED) as zf_out:
                    zf_out.writestr('resultado_agregado.csv', temp_csv_buffer.getvalue())
                zip_output_buffer.seek(0)
                st.session_state.excel_processed_data_zip = zip_output_buffer.getvalue()

            st.rerun() # Força um rerun para exibir os resultados e o botão de download

        else: # Se nenhum dado válido pôde ser processado
            st.error("Nenhum dado válido pôde ser processado a partir dos ficheiros Excel no ZIP. Verifique se os ficheiros não estão vazios ou corrompidos.")
            st.session_state.excel_processing_done = False
            st.session_state.excel_processed_data_zip = None
            st.session_state.excel_processed_data_csv_preview = None
            st.session_state.excel_arquivos_com_erro_state = []
            st.session_state.excel_last_uploaded_file_id = None # Resetar o ID para permitir novo upload


    # Exibir resultados e botão de download se o processamento estiver concluído
    if st.session_state.excel_processing_done:
        st.success("Todos os ficheiros Excel válidos foram lidos e agregados!")

        if st.session_state.excel_processed_data_csv_preview is not None:
            st.subheader("Pré-visualização dos Dados Agregados (CSV):")
            st.dataframe(st.session_state.excel_processed_data_csv_preview)

        if st.session_state.excel_processed_data_zip is not None:
            st.download_button(
                label="Descarregar Resultado Agregado (resultado_agregado.zip)",
                data=st.session_state.excel_processed_data_zip,
                file_name="resultado_agregado.zip",
                mime="application/zip"
            )

        if st.session_state.excel_arquivos_com_erro_state:
            st.warning("Alguns ficheiros dentro do ZIP tiveram erros e não foram incluídos no resultado:")
            for erro in st.session_state.excel_arquivos_com_erro_state:
                st.write(f"- {erro}")

    elif uploaded_zip_file is None: # Mensagem inicial se nenhum ficheiro foi carregado
        st.info("A aguardar o carregamento de um ficheiro ZIP contendo os ficheiros Excel...")



def parse_saft_xml_bytes(xml_bytes: bytes) -> Tuple[str, str, str]:
    """
    Recebe bytes XML (encoding desconhecido) e retorna
    (base_name, customers_csv_text, invoices_csv_text) em encoding latin-1 como str.
    """
    # tenta decodificar; fallback para latin-1
    try:
        xmlstring = xml_bytes.decode('utf-8')
    except Exception:
        xmlstring = xml_bytes.decode('latin-1', errors='replace')

    # remover namespace padrão para simplificar XPath
    xmlstring = re.sub(r'\sxmlns="[^"]+"', '', xmlstring, count=0)
    root = ET.fromstring(xmlstring)

    # preparar CSVs em memória
    customers_buf = io.StringIO()
    invoices_buf = io.StringIO()
    csvwritercustomer = csv.writer(customers_buf, lineterminator='\n')
    csvwriterinvoices = csv.writer(invoices_buf, lineterminator='\n')

    # Customers
    Customer_head = []
    count = 0
    for Customer in root.findall('./MasterFiles/Customer'):
        customer = []
        if count == 0:
            # cabeçalhos (tags)
            if Customer.find('CustomerID') is not None:
                Customer_head.append(Customer.find('CustomerID').tag)
            if Customer.find('CustomerTaxID') is not None:
                Customer_head.append(Customer.find('CustomerTaxID').tag)
            if Customer.find('CompanyName') is not None:
                Customer_head.append(Customer.find('CompanyName').tag)
            if Customer.find('./BillingAddress/Country') is not None:
                Customer_head.append(Customer.find('./BillingAddress/Country').tag)
            csvwritercustomer.writerow(Customer_head)
            count += 1
        else:
            ID = Customer.find('CustomerID').text if Customer.find('CustomerID') is not None else ''
            TaxID = Customer.find('CustomerTaxID').text if Customer.find('CustomerTaxID') is not None else ''
            Name = Customer.find('CompanyName').text if Customer.find('CompanyName') is not None else ''
            Country = Customer.find('./BillingAddress/Country').text if Customer.find('./BillingAddress/Country') is not None else ''
            csvwritercustomer.writerow([ID, TaxID, Name, Country])

    # Invoices
    Invoices_head = []
    first = True
    for Invoice in root.findall('./SourceDocuments/SalesInvoices/Invoice'):
        if first:
            # construir cabeçalhos com segurança (se existir)
            def tag_of(path, default=None):
                el = Invoice.find(path)
                return el.tag if el is not None and el.tag is not None else (default or path)
            try:
                Invoice.find('InvoiceNo')  # probe
                Invoices_head.append(tag_of('InvoiceNo'))
                Invoices_head.append(tag_of('./DocumentStatus/InvoiceStatus'))
                # Period pode não existir
                p = Invoice.find('Period')
                Invoices_head.append(p.tag if p is not None else 'Period')
                Invoices_head.append(tag_of('InvoiceDate'))
                Invoices_head.append(tag_of('InvoiceType'))
                Invoices_head.append(tag_of('CustomerID'))
                # para linhas, usamos tags da primeira Line (se existir)
                first_line = Invoice.find('./Line')
                if first_line is not None:
                    Invoices_head.append(first_line.find('ProductCode').tag if first_line.find('ProductCode') is not None else 'ProductCode')
                    Invoices_head.append(first_line.find('ProductDescription').tag if first_line.find('ProductDescription') is not None else 'ProductDescription')
                    Invoices_head.append(first_line.find('Quantity').tag if first_line.find('Quantity') is not None else 'Quantity')
                    Invoices_head.append(first_line.find('UnitOfMeasure').tag if first_line.find('UnitOfMeasure') is not None else 'UnitOfMeasure')
                    Invoices_head.append(first_line.find('UnitPrice').tag if first_line.find('UnitPrice') is not None else 'UnitPrice')
                    Invoices_head.append(first_line.find('Description').tag if first_line.find('Description') is not None else 'Description')
                else:
                    # tags padrão
                    Invoices_head += ["ProductCode","ProductDescription","Quantity","UnitOfMeasure","UnitPrice","Description"]
                Invoices_head += ["Amount","TaxAmount","TaxCountryRegion","Reference","Reason"]
            except Exception:
                pass
            csvwriterinvoices.writerow(Invoices_head)
            first = False

        # cada linha da fatura vira uma linha no CSV
        for Line in Invoice.findall('./Line'):
            row = []
            row.append(Invoice.find('InvoiceNo').text if Invoice.find('InvoiceNo') is not None else '')
            row.append(Invoice.find('./DocumentStatus/InvoiceStatus').text if Invoice.find('./DocumentStatus/InvoiceStatus') is not None else '')
            try:
                row.append(Invoice.find('Period').text)
            except:
                invoice_date = Invoice.find('InvoiceDate').text if Invoice.find('InvoiceDate') is not None else ''
                row.append(invoice_date[5:7] if invoice_date else '')
            row.append(Invoice.find('InvoiceDate').text if Invoice.find('InvoiceDate') is not None else '')
            row.append(Invoice.find('InvoiceType').text if Invoice.find('InvoiceType') is not None else '')
            row.append(Invoice.find('CustomerID').text if Invoice.find('CustomerID') is not None else '')

            row.append(Line.find('ProductCode').text if Line.find('ProductCode') is not None else '')
            row.append(Line.find('ProductDescription').text if Line.find('ProductDescription') is not None else '')
            row.append(Line.find('Quantity').text if Line.find('Quantity') is not None else '')
            row.append(Line.find('UnitOfMeasure').text if Line.find('UnitOfMeasure') is not None else '')
            row.append(Line.find('UnitPrice').text if Line.find('UnitPrice') is not None else '')
            row.append(Line.find('Description').text if Line.find('Description') is not None else '')

            # Amount: DebitAmount negative, else CreditAmount
            debit = Line.find('DebitAmount')
            credit = Line.find('CreditAmount')
            if debit is not None and debit.text:
                row.append("-" + debit.text)
            elif credit is not None and credit.text:
                row.append(credit.text)
            else:
                row.append('')

            taxamt_el = Line.find('./Tax/TaxAmount')
            row.append(taxamt_el.text if taxamt_el is not None else '0')

            taxcr_el = Line.find('./Tax/TaxCountryRegion')
            row.append(taxcr_el.text if taxcr_el is not None else '')

            ref_el = Line.find('./References/Reference')
            row.append(ref_el.text if ref_el is not None else '')

            reason_el = Line.find('./References/Reason')
            row.append(reason_el.text if reason_el is not None else '')

            csvwriterinvoices.writerow(row)

    # Obter textos (em latin-1 para manter compatibilidade com o resto do ecossistema)
    customers_text = customers_buf.getvalue()
    invoices_text = invoices_buf.getvalue()
    customers_buf.close()
    invoices_buf.close()

    # base name usado para o ficheiro ZIP
    base_name = "saft"  # fallback
    # tentar extrair um nome mais amigável do XML (por exemplo raiz/CompanyID) — opcional
    try:
        # se existir <Header>/<AuditFileVersion> ou similar, não assumimos sempre a mesma tag
        base_name = "saft_export"
    except Exception:
        pass

    return base_name, customers_text, invoices_text


def saf_t_tab():
    st.header("SAF-T Faturação → CSV")
    uploaded = st.file_uploader("Escolha um ficheiro .xml ou um .zip contendo .xml", type=["xml", "zip"])
    if uploaded is None:
        st.info("Faça upload de um ficheiro SAF-T (.xml) ou um .zip que contenha um .xml.")
        return

    file_bytes = uploaded.read()
    filename = uploaded.name

    # Se for zip, abrir e perguntar qual xml usar (se houver vários)
    xml_bytes = None
    xml_name = None
    if filename.lower().endswith('.zip'):
        with io.BytesIO(file_bytes) as bio:
            with zipfile.ZipFile(bio) as z:
                xml_files = [n for n in z.namelist() if n.lower().endswith('.xml')]
                if not xml_files:
                    st.error("O ZIP não contém ficheiros .xml válidos.")
                    return
                if len(xml_files) == 1:
                    xml_name = xml_files[0]
                else:
                    xml_name = st.selectbox("Selecione o ficheiro XML dentro do ZIP", xml_files)
                xml_bytes = z.read(xml_name)
    else:
        xml_bytes = file_bytes
        xml_name = filename

    if st.button("Processar SAF-T"):
        try:
            base_name, customers_csv, invoices_csv = parse_saft_xml_bytes(xml_bytes)
        except Exception as e:
            st.error(f"Erro ao analisar o ficheiro XML: {e}")
            return

        # mostrar um preview dos CSVs
        st.subheader("Preview Customers (primeiras 2000 chars)")
        st.code(customers_csv[:2000], language='text')
        st.subheader("Preview Invoices (primeiras 2000 chars)")
        st.code(invoices_csv[:2000], language='text')

        # criar ZIP em memória com os dois CSVs (codificados em latin-1)
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            customers_filename = xml_name.rsplit('.', 1)[0] + "-Customers.csv"
            invoices_filename = xml_name.rsplit('.', 1)[0] + "-Invoices.csv"
            zf.writestr(customers_filename, customers_csv.encode('latin-1', errors='replace'))
            zf.writestr(invoices_filename, invoices_csv.encode('latin-1', errors='replace'))
        zip_buffer.seek(0)

        st.download_button(
            label="Descarregar CSVs (ZIP)",
            data=zip_buffer.getvalue(),
            file_name=f"{xml_name.rsplit('.',1)[0]}-CSVs.zip",
            mime="application/zip"
        )

# ======= OrangeHRM Pivot Tab (auto-contido) =======
def _pkce_generate_code_verifier(n_bytes: int = 64) -> str:
    # Gera um code_verifier url-safe (43-128 chars). 64 bytes -> 86 chars base64-url.
    return base64.urlsafe_b64encode(os.urandom(n_bytes)).rstrip(b"=").decode("ascii")

def _pkce_code_challenge(verifier: str) -> str:
    digest = hashlib.sha256(verifier.encode("ascii")).digest()
    return base64.urlsafe_b64encode(digest).rstrip(b"=").decode("ascii")

def _oauth_token_exchange_with_pkce(token_url: str, client_id: str, code: str, redirect_uri: str, code_verifier: str):
    data = {
        "grant_type": "authorization_code",
        "client_id": client_id,
        "code": code,
        "redirect_uri": redirect_uri,
        "code_verifier": code_verifier,
    }
    # Importante: sem Authorization header
    r = requests.post(token_url, data=data, timeout=30)
    ctype = r.headers.get("Content-Type", "")
    body = r.json() if "application/json" in ctype else {"raw": r.text}
    return r.status_code, body

def _get_setting(key: str, default: str = "") -> str:
    # Lê ORANGEHRM_* do ambiente; fallback para st.secrets['orangehrm'][key]
    env_key = f"ORANGEHRM_{key.upper()}"
    v = os.getenv(env_key)
    if v:
        return v
    try:
        return st.secrets["orangehrm"].get(key, default)
    except Exception:
        return default

# Cliente com refresh automático
class _OrangeHRMClient:
    def __init__(self, client_id: str, refresh_token: str, token_url: str, api_base: str):
        self.client_id = client_id
        self.token_url = token_url
        self.api_base = api_base

        if "orange_access_token" not in st.session_state:
            st.session_state["orange_access_token"] = ""
        if "orange_expires_at" not in st.session_state:
            st.session_state["orange_expires_at"] = 0.0
        if "orange_refresh_token" not in st.session_state:
            st.session_state["orange_refresh_token"] = refresh_token

    @property
    def access_token(self) -> str:
        return st.session_state["orange_access_token"]

    @property
    def refresh_token(self) -> str:
        return st.session_state["orange_refresh_token"]

    @property
    def expires_at(self) -> float:
        return st.session_state["orange_expires_at"]

    def _save_tokens(self, token_resp: dict):
        st.session_state["orange_access_token"] = token_resp["access_token"]
        if "refresh_token" in token_resp and token_resp["refresh_token"]:
            st.session_state["orange_refresh_token"] = token_resp["refresh_token"]
        expires_in = int(token_resp.get("expires_in", 3600))
        st.session_state["orange_expires_at"] = time.time() + max(expires_in - 60, 0)

    def _needs_refresh(self) -> bool:
        return (not self.access_token) or time.time() >= self.expires_at

    def _refresh(self) -> bool:
        if not self.refresh_token:
            return False
        data = {
            "grant_type": "refresh_token",
            "client_id": self.client_id,
            "refresh_token": self.refresh_token,
        }
        r = requests.post(self.token_url, data=data, timeout=30)
        if r.ok:
            self._save_tokens(r.json())
            return True
        else:
            st.warning(f"Falha ao renovar token ({r.status_code}): {r.text}")
            return False

    def _ensure_token(self):
        if self._needs_refresh():
            ok = self._refresh()
            if not ok:
                raise RuntimeError("Não foi possível obter token de acesso (refresh falhou).")

    def request(self, method: str, path: str, retry_on_401: bool = True, **kwargs):
        self._ensure_token()
        url = path if path.startswith("http") else self.api_base + path.lstrip("/")
        headers = kwargs.pop("headers", {})
        headers = {"Authorization": f"Bearer {self.access_token}", **headers}
        resp = requests.request(method, url, headers=headers, timeout=60, **kwargs)
        if resp.status_code == 401 and retry_on_401:
            if self._refresh():
                headers["Authorization"] = f"Bearer {self.access_token}"
                resp = requests.request(method, url, headers=headers, timeout=60, **kwargs)
        resp.raise_for_status()
        ctype = resp.headers.get("Content-Type", "")
        return resp.json() if "application/json" in ctype else resp.text

# Endpoints
_PATH_LIST_EMPLOYEE_TIMESHEETS = "time/employees/{empNumber}/timesheets"
_PATH_GET_TIMESHEET_ENTRIES    = "time/employees/timesheets/{timesheetId}/entries"

# Funções de dados
def _list_all_employees(client: _OrangeHRMClient, limit: int = 200) -> List[Dict[str, Any]]:
    results, offset = [], 0
    while True:
        params = {"limit": limit, "offset": offset}
        data = client.request("GET", "pim/employees", params=params)
        rows = []
        if isinstance(data, dict) and "data" in data and isinstance(data["data"], list):
            rows = data["data"]
        elif isinstance(data, list):
            rows = data
        else:
            break
        results.extend(rows)
        if len(rows) < limit:
            break
        offset += limit
    return results

def _full_name_from_employee_row(e: Dict[str, Any]) -> str:
    first  = (e.get("firstName")  or "").strip()
    middle = (e.get("middleName") or "").strip()
    last   = (e.get("lastName")   or "").strip()
    parts = [p for p in [first, middle, last] if p]
    if parts:
        return " ".join(parts)
    if isinstance(e.get("name"), str) and e["name"].strip():
        return e["name"].strip()
    return str(e.get("empNumber") or e.get("employeeNumber") or e.get("code") or "")

def _build_empnumber_to_name_map(employees: List[Dict[str, Any]]) -> Dict[str, str]:
    mapping = {}
    for e in employees:
        emp_no = e.get("empNumber") or e.get("employeeNumber") or e.get("code")
        if emp_no is not None:
            mapping[str(emp_no)] = _full_name_from_employee_row(e)
    return mapping

def _list_employee_timesheets(client: _OrangeHRMClient, emp_number: str, from_date: str = None, to_date: str = None, limit: int = 50, offset: int = 0) -> List[Dict[str, Any]]:
    path = _PATH_LIST_EMPLOYEE_TIMESHEETS.format(empNumber=emp_number)
    params = {"limit": limit, "offset": offset}
    if from_date: params["fromDate"] = from_date
    if to_date:   params["toDate"]   = to_date
    data = client.request("GET", path, params=params)
    if isinstance(data, dict) and "data" in data: return data["data"]
    if isinstance(data, list): return data
    return []

def _get_timesheet_entries(client: _OrangeHRMClient, timesheet_id: str) -> List[Dict[str, Any]]:
    path = _PATH_GET_TIMESHEET_ENTRIES.format(timesheetId=timesheet_id)
    data = client.request("GET", path)
    if isinstance(data, dict) and "data" in data: return data["data"]
    if isinstance(data, list): return data
    return []

def _sum_timesheet_hours(entries: List[Dict[str, Any]]) -> float:
    total_hours = 0.0
    for e in entries:
        t = e.get("total", {})
        h = t.get("hours")
        m = t.get("minutes")
        if isinstance(h, (int, float)):
            add = float(h)
            if isinstance(m, (int, float)) and m:
                add += float(m) / 60.0
            total_hours += add
            continue
        dates = e.get("dates", {})
        for _, d in dates.items():
            dur = d.get("duration")
            if isinstance(dur, str) and ":" in dur:
                hh, mm = dur.split(":", 1)
                try:
                    total_hours += int(hh) + int(mm) / 60.0
                except:
                    pass
            elif isinstance(dur, (int, float)):
                total_hours += float(dur)
    return round(total_hours, 2)

def _get_totals_by_employee_and_timesheet(client: _OrangeHRMClient, emp_numbers: List[str], empname_map: Dict[str, str], from_date: str = None, to_date: str = None) -> List[Dict[str, Any]]:
    rows = []
    for emp in emp_numbers:
        offset = 0
        while True:
            sheets = _list_employee_timesheets(client, emp, from_date=from_date, to_date=to_date, limit=50, offset=offset)
            if not sheets:
                break
            for ts in sheets:
                ts_id = str(ts.get("id") or ts.get("timesheetId"))
                period_start = ts.get("startDate") or ts.get("fromDate")
                period_end = ts.get("endDate") or ts.get("toDate")
                if not ts_id:
                    continue
                entries = _get_timesheet_entries(client, ts_id)
                total_hours = _sum_timesheet_hours(entries)
                rows.append({
                    "empNumber": emp,
                    "empName": empname_map.get(str(emp), str(emp)),
                    "timesheetId": ts_id,
                    "periodStart": period_start,
                    "periodEnd": period_end,
                    "totalHours": total_hours
                })
            if len(sheets) < 50:
                break
            offset += 50
    return rows

def _pivot_hours_by_employee_and_start(rows: List[Dict[str, Any]]) -> pd.DataFrame:
    if not rows:
        return pd.DataFrame()
    df = pd.DataFrame(rows)
    index_col = "empName" if "empName" in df.columns else "empNumber"
    pivot = pd.pivot_table(
        df,
        index=index_col,
        columns="periodStart",
        values="totalHours",
        aggfunc="sum",
        fill_value=0.0,
    )
    try:
        ordered_cols = sorted(pivot.columns, key=pd.to_datetime)
        pivot = pivot[ordered_cols]
    except Exception:
        pass
    return pivot.sort_index()

# Cache sem fechar sobre objetos não-hasháveis
@st.cache_data(ttl=600, show_spinner=False)
def _cached_employees_and_map(domain: str, client_id: str, refresh_token: str):
    api_base = domain.rstrip("/") + "/api/v2/"
    token_url = domain.rstrip("/") + "/oauth2/token"
    client = _OrangeHRMClient(client_id, refresh_token, token_url, api_base)
    emps = _list_all_employees(client, limit=200)
    emp_map = _build_empnumber_to_name_map(emps)
    emp_numbers = list(emp_map.keys())
    return emps, emp_map, emp_numbers

def _sum_entry_hours(entry: Dict[str, Any]) -> float:
    # Soma horas de uma entrada, suportando tanto 'total' (hours/minutes) como 'dates' com 'HH:MM'
    t = entry.get("total", {})
    h = t.get("hours")
    m = t.get("minutes")
    if isinstance(h, (int, float)):
        add = float(h)
        if isinstance(m, (int, float)) and m:
            add += float(m) / 60.0
        return add

    total = 0.0
    dates = entry.get("dates", {})
    for _day, d in dates.items():
        dur = d.get("duration")
        if isinstance(dur, str) and ":" in dur:
            try:
                hh, mm = dur.split(":", 1)
                total += int(hh) + int(mm) / 60.0
            except:
                pass
        elif isinstance(dur, (int, float)):
            total += float(dur)
    return total


def _get_hours_by_employee_and_project(
    client: "_OrangeHRMClient",
    emp_numbers: List[str],
    empname_map: Dict[str, str],
    from_date: str = None,
    to_date: str = None,
) -> List[Dict[str, Any]]:
    """
    Devolve linhas agregadas por colaborador × projeto (formato longo).
    Cada linha: { empNumber, empName, projectId, projectName, totalHours }
    """
    acc: Dict[tuple, float] = {}
    proj_names: Dict[str, str] = {}

    for emp in emp_numbers:
        offset = 0
        while True:
            sheets = _list_employee_timesheets(client, emp, from_date=from_date, to_date=to_date, limit=50, offset=offset)
            if not sheets:
                break
            for ts in sheets:
                ts_id = str(ts.get("id") or ts.get("timesheetId") or "")
                if not ts_id:
                    continue
                entries = _get_timesheet_entries(client, ts_id)
                for e in entries:
                    hours = _sum_entry_hours(e)
                    # Extrair projeto de forma robusta
                    proj = e.get("project") if isinstance(e.get("project"), dict) else {}
                    project_id = proj.get("id") or e.get("projectId") or e.get("project_id")
                    project_name = (
                        proj.get("name")
                        or e.get("projectName")
                        or e.get("project_name")
                        or "Sem Projeto"
                    )
                    # Chave de agregação usa emp + project_id se existir; senão emp + nome
                    key = (str(emp), str(project_id) if project_id is not None else f"name::{project_name}")
                    acc[key] = acc.get(key, 0.0) + hours
                    # Guardar nome do projeto por chave (prioriza nome visto)
                    proj_names[key] = project_name
            if len(sheets) < 50:
                break
            offset += 50

    rows: List[Dict[str, Any]] = []
    for key, total in acc.items():
        emp_key, proj_key = key
        rows.append({
            "empNumber": emp_key,
            "empName": empname_map.get(emp_key, emp_key),
            "projectId": proj_key if not proj_key.startswith("name::") else None,
            "projectName": proj_names.get(key, "Sem Projeto"),
            "totalHours": round(total, 2),
        })
    return rows
    
def _pivot_hours_by_employee_and_project(rows: List[Dict[str, Any]]) -> pd.DataFrame:
    if not rows:
        return pd.DataFrame()
    df = pd.DataFrame(rows)
    index_col = "empName" if "empName" in df.columns else "empNumber"
    col_col = "projectName" if "projectName" in df.columns else "projectKey"
    pivot = pd.pivot_table(
        df,
        index=index_col,
        columns=col_col,
        values="totalHours",
        aggfunc="sum",
        fill_value=0.0,
    )
    # Ordena alfabeticamente os projetos para leitura mais fácil
    try:
        pivot = pivot[sorted(pivot.columns, key=lambda x: (str(x).lower()))]
    except Exception:
        pass
    return pivot.sort_index()

def render_orangehrm_oauth_bootstrap_tab():
    st.header("Configuração OAuth (Admin) — Obter Refresh Token")

    # Lê definições atuais (env > secrets)
    domain = _get_setting("domain") or "https://rh.inobest.com/web/index.php/"
    client_id = _get_setting("client_id") or ""
    default_redirect = domain.rstrip("/")  # usa o mesmo domínio por defeito
    token_url = domain.rstrip("/") + "/oauth2/token"
    auth_base = domain.rstrip("/") + "/oauth2/authorize"

    st.markdown("Use esta aba apenas uma vez para gerar um refresh_token e colocá-lo no stack.env. Depois pode remover esta aba.")

    with st.form("oauth_bootstrap_form"):
        c1, c2 = st.columns([1, 1])
        with c1:
            client_id_in = st.text_input("Client ID", value=client_id, help="O Client ID registado no OrangeHRM (OAuth).")
            redirect_uri = st.text_input("Redirect URI", value=default_redirect, help="Tem de corresponder ao Redirect URI registado no OAuth.")
        with c2:
            st.write(" ")

        # Gerar/verificar PKCE
        # Mantemos na session_state para que o mesmo verifier seja usado na autorização e na troca
        if "pkce_code_verifier" not in st.session_state or st.form_submit_button("Gerar novo code_verifier"):
            st.session_state["pkce_code_verifier"] = _pkce_generate_code_verifier()
        code_verifier = st.session_state["pkce_code_verifier"]
        code_challenge = _pkce_code_challenge(code_verifier)

        st.text_input("Code Verifier (guarde até concluir)", value=code_verifier, disabled=True)
        st.text_input("Code Challenge (S256)", value=code_challenge, disabled=True)

        # Montar URL de autorização
        params = {
            "response_type": "code",
            "code_challenge_method": "S256",
            "code_challenge": code_challenge,
            "client_id": client_id_in,
            "redirect_uri": redirect_uri,
        }
        from urllib.parse import urlencode
        auth_url = auth_base + "?" + urlencode(params)

        st.write("1) Clique no link abaixo, autentique-se e autorize a aplicação. Será redirecionado para o redirect_uri com um parâmetro 'code'.")
        st.write(f"[Abrir e autorizar no OrangeHRM]({auth_url})")

        auth_code = st.text_input("2) Cole aqui o 'code' devolvido no redirect", value="", help="Exemplo: https://.../?code=ABC...XYZ -> cole apenas o valor do code.")

        submit_exchange = st.form_submit_button("3) Trocar code por tokens")
        if submit_exchange:
            if not auth_code:
                st.error("Cole o 'code' devolvido pelo OrangeHRM.")
            elif not client_id_in or not redirect_uri:
                st.error("Preencha o Client ID e o Redirect URI.")
            else:
                st.info(f"A enviar pedido ao token endpoint: {token_url}")
                status, body = _oauth_token_exchange_with_pkce(
                    token_url=token_url,
                    client_id=client_id_in,
                    code=auth_code,
                    redirect_uri=redirect_uri,
                    code_verifier=code_verifier,
                )
                if status == 200 and isinstance(body, dict) and "refresh_token" in body:
                    st.success("Tokens obtidos com sucesso.")
                    st.json({
                        "access_token_prefix": body.get("access_token", "")[:16] + "...",
                        "expires_in": body.get("expires_in"),
                        "refresh_token_prefix": body.get("refresh_token", "")[:12] + "...",
                        "token_type": body.get("token_type"),
                        "scope": body.get("scope"),
                    })
                    st.code(body.get("refresh_token", ""), language="text")
                    st.warning("Copie o refresh_token acima e atualize o ficheiro stack.env no Portainer: ORANGEHRM_REFRESH_TOKEN=<este valor>. Depois redeploy.")
                else:
                    st.error(f"Falha ao obter tokens (HTTP {status}).")
                    st.write("Resposta do servidor:")
                    st.json(body)

def render_orangehrm_pivot_tab():
    # Ler configurações
    domain = _get_setting("domain") or "https://rh.inobest.com/web/index.php/"
    client_id = _get_setting("client_id") or ""
    refresh_token = _get_setting("refresh_token") or ""

    if not client_id or not refresh_token:
        st.error("Credenciais em falta. Defina ORANGEHRM_CLIENT_ID e ORANGEHRM_REFRESH_TOKEN no ambiente (stack.env) ou em st.secrets['orangehrm'].")
        return

    api_base = domain.rstrip("/") + "/api/v2/"
    token_url = domain.rstrip("/") + "/oauth2/token"
    client = _OrangeHRMClient(client_id, refresh_token, token_url, api_base)

    st.header("Folhas de Horas — Tabela Dinâmica por Colaborador")

    # Filtros no topo da aba
    with st.container():
        c1, c2, c3 = st.columns([1, 1, 0.6])
        with c1:
            from_date = st.date_input("De (início do período)", None, format="YYYY-MM-DD", key="ts_from_date")
        with c2:
            to_date = st.date_input("Até (fim do período)", None, format="YYYY-MM-DD", key="ts_to_date")
        with c3:
            run_btn = st.button("Gerar Tabela Dinâmica", key="run_pivot_btn")

    from_date_str = from_date.isoformat() if from_date else None
    to_date_str   = to_date.isoformat() if to_date else None

    # Carregar colaboradores (em cache)
    with st.spinner("A carregar colaboradores..."):
        employees, emp_map, all_emp_numbers = _cached_employees_and_map(domain, client_id, refresh_token)

    # Filtro de colaboradores nesta aba
    emp_display = {k: emp_map.get(k, k) for k in all_emp_numbers}
    emp_choices = st.multiselect(
        "Filtrar colaboradores (opcional)",
        options=all_emp_numbers,
        default=all_emp_numbers[: min(len(all_emp_numbers), 20)],
        format_func=lambda k: emp_display.get(k, k),
        key="emp_filter_multiselect",
    )

    if run_btn:
        if not emp_choices:
            st.warning("Selecione pelo menos um colaborador.")
            return

        with st.spinner("A obter folhas de horas e a calcular totais..."):
            rows = _get_totals_by_employee_and_timesheet(
                client,
                emp_choices,
                emp_map,
                from_date=from_date_str,
                to_date=to_date_str,
            )
            pivot_df = _pivot_hours_by_employee_and_start(rows)

        st.subheader("Horas por Colaborador × Data de Início da Folha de Horas")
        if pivot_df.empty:
            st.info("Não foram encontrados dados para os filtros selecionados.")
        else:
            st.dataframe(pivot_df, use_container_width=True)

            # Descarregar Excel (inclui folha 'Pivot' e folha 'Dados')
            xlsx_buffer = BytesIO()
            with pd.ExcelWriter(xlsx_buffer, engine="openpyxl") as writer:
                pivot_df.to_excel(writer, sheet_name="Pivot", index=True)
                pd.DataFrame(rows).to_excel(writer, sheet_name="Dados", index=False)
            st.download_button(
                "Descarregar Excel",
                data=xlsx_buffer.getvalue(),
                file_name="folhas_horas_pivot.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_xlsx_button",
            )
        # ——————————————————————————————————————————————
        # Nova secção: Tabela por Cliente com Projetos em linhas
        # ——————————————————————————————————————————————
        st.subheader("Projetos por Cliente — Totais de Horas")
        
        with st.spinner("A obter entradas e a agregar por cliente e projeto..."):
            client_proj_rows = _get_hours_by_client_and_project(
                client,
                emp_choices,
                from_date=from_date_str,
                to_date=to_date_str,
            )
            client_proj_df = pd.DataFrame(client_proj_rows)
        
        if client_proj_df.empty:
            st.info("Não foram encontrados dados por cliente/projeto para os filtros selecionados.")
        else:
            # Ordenar por cliente e horas desc
            client_proj_df = client_proj_df.sort_values(by=["clientName", "totalHours"], ascending=[True, False])
        
            # Mostrar tabela por cliente (um bloco por cliente)
            for customer, g in client_proj_df.groupby("clientName", sort=False):
                st.markdown(f"### {customer}")
                show = g[["projectName", "totalHours"]].rename(columns={
                    "projectName": "Projeto",
                    "totalHours": "Total de Horas"
                })
                st.dataframe(show.reset_index(drop=True), use_container_width=True)
        
            # Exportação apenas em Excel (sem CSV)
            client_proj_xlsx_buffer = BytesIO()
            with pd.ExcelWriter(client_proj_xlsx_buffer, engine="openpyxl") as writer:
                # Folha detalhada (formato longo)
                client_proj_df.rename(columns={
                    "clientName": "Cliente",
                    "projectName": "Projeto",
                    "totalHours": "TotalHoras"
                })[["Cliente", "Projeto", "TotalHoras"]].to_excel(writer, sheet_name="Projetos_por_Cliente", index=False)
        
                # Opcional: uma folha por cliente
                for customer, g in client_proj_df.groupby("clientName", sort=False):
                    g.rename(columns={"projectName": "Projeto", "totalHours": "TotalHoras"})[["Projeto", "TotalHoras"]].to_excel(
                        writer, sheet_name=(str(customer)[:28] or "Sem_Cliente"), index=False
                    )
        
            st.download_button(
                "Descarregar Excel (Projetos por Cliente)",
                data=client_proj_xlsx_buffer.getvalue(),
                file_name="folhas_horas_projetos_por_cliente.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_xlsx_proj_por_cliente",
            )
# ======= Fim do bloco OrangeHRM Pivot Tab =======


# --- Lógica Principal da Aplicação (Usando Abas) ---
tab1, tab2, tab3, tab4 = st.tabs(["Agregador de Excel", "SAF-T Faturação → CSV", "Configuração OAuth (Admin)", "Timesheets Pivot"])
with tab1:
    excel_aggregator_app()
with tab2:
    saf_t_tab()
with tab3:
    render_orangehrm_oauth_bootstrap_tab()
with tab4:
    render_orangehrm_pivot_tab()
