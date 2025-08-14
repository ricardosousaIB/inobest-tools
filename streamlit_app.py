import streamlit as st
import io, zipfile, re, csv, os, time, requests, base64, hashlib, json, tempfile
from urllib.parse import urlencode
import xml.etree.ElementTree as ET
import pandas as pd
from typing import Optional, List, Dict, Any, Tuple
from io import BytesIO, StringIO

# Configure page early
st.set_page_config(layout="wide", page_title="Ferramentas Inobest — O365 + OrangeHRM")

# =============================
# O365 AUTH (Global App Login)
# =============================

def _o365_get_setting(key: str) -> str:
    """Read O365 settings from st.secrets['o365'] with env fallback O365_*"""
    try:
        val = st.secrets["o365"].get(key)
    except Exception:
        val = None
    return (val or os.environ.get(f"O365_{key.upper()}", "")).strip()

def _o365_authority() -> str:
    tenant = _o365_get_setting("tenant_id") or "common"  # fallback keeps login working in dev
    return f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0"

def _o365_auth_and_token_urls() -> Tuple[str, str]:
    base = _o365_authority()
    return base + "/authorize", base + "/token"

def _o365_begin_login() -> Optional[str]:
    client_id = _o365_get_setting("client_id")
    redirect_uri = _o365_get_setting("redirect_uri")
    missing = []
    if not client_id:
        missing.append("O365_CLIENT_ID")
    if not redirect_uri:
        missing.append("O365_REDIRECT_URI")
    if missing:
        st.error("Configuração O365 em falta: " + ", ".join(missing) + ". Defina no stack.env (ou secrets) e faça redeploy.")
        return None
    auth_url, _ = _o365_auth_and_token_urls()
    scopes = ["openid", "profile", "email", "User.Read"]
    params = {
        "client_id": client_id,
        "response_type": "code",
        "redirect_uri": redirect_uri,
        "response_mode": "query",
        "scope": " ".join(scopes),
    }
    return f"{auth_url}?{urlencode(params)}"

def _o365_exchange_code_for_tokens(code: str) -> dict:
    _, token_url = _o365_auth_and_token_urls()
    client_id = _o365_get_setting("client_id")
    client_secret = _o365_get_setting("client_secret")
    redirect_uri = _o365_get_setting("redirect_uri")
    data = {
        "client_id": client_id,
        "client_secret": client_secret,
        "grant_type": "authorization_code",
        "code": code,
        "redirect_uri": redirect_uri,
    }
    r = requests.post(token_url, data=data, timeout=30)
    r.raise_for_status()
    return r.json()

def _o365_refresh_tokens(refresh_token: str) -> Optional[dict]:
    _, token_url = _o365_auth_and_token_urls()
    client_id = _o365_get_setting("client_id")
    client_secret = _o365_get_setting("client_secret")
    data = {
        "client_id": client_id,
        "client_secret": client_secret,
        "grant_type": "refresh_token",
        "refresh_token": refresh_token,
        "scope": "openid profile email User.Read",
    }
    r = requests.post(token_url, data=data, timeout=30)
    if not r.ok:
        return None
    return r.json()

def _o365_parse_id_token(id_token: str) -> dict:
    """Parse JWT without signature verification — only to extract non-sensitive claims in frontend."""
    try:
        payload_b64 = id_token.split(".")[1] + "=="
        payload_json = base64.urlsafe_b64decode(payload_b64.encode("utf-8")).decode("utf-8")
        return json.loads(payload_json)
    except Exception:
        return {}

def _safe_rerun():
    try:
        st.rerun()
    except Exception:
        # compatibility with older Streamlit
        try:
            st.experimental_rerun()  # type: ignore[attr-defined]
        except Exception:
            pass

def ensure_o365_login() -> bool:
    """Global guard. If not logged in via O365, prompt user to sign in; else refresh tokens as needed."""
    qp = st.query_params
    code = qp.get("code", [None])[0] if isinstance(qp.get("code"), list) else qp.get("code")

    if "o365_auth" not in st.session_state:
        st.session_state["o365_auth"] = {}
    sess = st.session_state["o365_auth"]

    # Finish sign-in if code present
    if code and not sess.get("access_token"):
        with st.spinner("A concluir início de sessão Microsoft..."):
            try:
                tokens = _o365_exchange_code_for_tokens(code)
                sess["access_token"] = tokens.get("access_token")
                sess["refresh_token"] = tokens.get("refresh_token")
                sess["expires_at"] = time.time() + int(tokens.get("expires_in", 3600)) - 30
                sess["id_token"] = tokens.get("id_token", "")
                claims = _o365_parse_id_token(sess["id_token"]) if sess.get("id_token") else {}
                sess["email"] = (claims.get("preferred_username") or claims.get("email") or "").strip()
                sess["name"] = (claims.get("name") or "").strip()
                sess["oid"] = claims.get("oid") or ""
                st.query_params.clear()
                _safe_rerun()
                return True
            except Exception as e:
                st.error(f"Falha a concluir sessão O365: {e}")
                st.query_params.clear()
                return False

    # Refresh if expired
    if sess.get("access_token") and time.time() >= sess.get("expires_at", 0):
        tokens = _o365_refresh_tokens(sess.get("refresh_token", ""))
        if tokens:
            sess["access_token"] = tokens.get("access_token")
            sess["refresh_token"] = tokens.get("refresh_token", sess.get("refresh_token"))
            sess["expires_at"] = time.time() + int(tokens.get("expires_in", 3600)) - 30
        else:
            st.session_state.pop("o365_auth", None)

    # If still not authenticated, render sign-in
    if not sess.get("access_token"):
        st.info("Para continuar, autentique-se com a sua conta Microsoft 365.")
        if st.button("Iniciar sessão com Microsoft", key="o365_login_btn"):
            login_url = _o365_begin_login()
            if login_url:
                st.markdown(f"[Clique aqui para iniciar sessão]({login_url})")
        return False

    return True

def o365_logout_button():
    sess = st.session_state.get("o365_auth", {})
    user_label = sess.get("name") or sess.get("email") or "Utilizador"
    c1, c2 = st.columns([1, 0.3])
    with c1:
        st.caption(f"Autenticado como: {user_label}")
    with c2:
        if st.button("Terminar sessão", key="o365_logout_btn"):
            st.session_state.pop("o365_auth", None)
            _safe_rerun()


def _is_admin_o365() -> bool:
    admins: List[str] = []
    try:
        admins = st.secrets["o365"].get("admin_emails") or []
    except Exception:
        admins = []
    if not admins:
        env_admins = os.environ.get("O365_ADMIN_EMAILS", "")
        admins = [a.strip() for a in env_admins.split(",") if a.strip()]
    email = st.session_state.get("o365_auth", {}).get("email", "").lower()
    return email in [a.lower() for a in admins]

# Enforce global login
if not ensure_o365_login():
    st.stop()

# Header with logout
o365_logout_button()

# ========================================
# Admin-only password gate for OAuth tab
# ========================================

def _get_oauth_admin_password() -> str:
    # Priority: st.secrets, then env var
    pwd = None
    try:
        pwd = st.secrets.get("oauth_admin_password")
    except Exception:
        pwd = None
    if not pwd:
        pwd = os.environ.get("OAUTH_ADMIN_PASSWORD")
    return pwd or ""

def _read_shared_refresh_token(path: str, fallback: str) -> str:
    try:
        if path and os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            rt = (data or {}).get("refresh_token", "")
            if rt and isinstance(rt, str):
                return rt.strip()
    except Exception:
        pass
    return (fallback or "").strip()

def _write_shared_refresh_token(path: str, refresh_token: str) -> None:
    if not path or not refresh_token:
        return
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        tmp = f"{path}.tmp-{int(time.time()*1000)}"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump({"refresh_token": refresh_token}, f)
        os.replace(tmp, path)  # atomic on same filesystem
    except Exception:
        pass

def _ensure_oauth_admin() -> str:
    """
    Protege a aba OAuth Admin.
    - Se o utilizador constar nos emails de admin O365, dá acesso sem password.
    - Caso contrário, permite acesso com password (se configurada).
    - Não chama st.stop(); devolve um estado para a aba decidir o que renderizar.

    Returns: "ok" | "needs-auth" | "no-access"
    """
    # Auto-grant para administradores definidos por email
    if _is_admin_o365():
        return "ok"

    # Fallback: password gate
    if "oauth_admin_ok" not in st.session_state:
        st.session_state["oauth_admin_ok"] = False

    if st.session_state["oauth_admin_ok"]:
        col1, col2 = st.columns([1, 0.25])
        with col2:
            if st.button("Terminar sessão", key="oauth_admin_logout"):
                st.session_state["oauth_admin_ok"] = False
                _safe_rerun()
        return "ok"

    admin_pwd = _get_oauth_admin_password()
    if not admin_pwd:
        st.error("Área restrita. Contacte um administrador para aceder a esta secção (password não configurada).")
        return "no-access"

    st.info("Área restrita. Introduza a password de administrador para continuar.")
    with st.form("oauth_admin_login"):
        pwd = st.text_input("Password de administrador", type="password")
        ok = st.form_submit_button("Entrar")
    if ok:
        if pwd == admin_pwd:
            st.session_state["oauth_admin_ok"] = True
            st.success("Autenticação bem-sucedida.")
            _safe_rerun()
            return "ok"
        else:
            st.error("Password incorreta.")
            return "no-access"

    return "needs-auth"

# ========================================
# Excel Aggregator (ZIP of Excel -> single CSV in ZIP)
# ========================================

def excel_aggregator_app():
    st.header("Agregador de Ficheiros Excel (via ZIP)")
    st.write("Carregue um ficheiro ZIP contendo os seus ficheiros Excel (.xls ou .xlsx) para agregá-los num único ficheiro CSV, depois comprimido num novo ZIP.")

    # session state init
    if 'excel_processed_data_zip' not in st.session_state:
        st.session_state.excel_processed_data_zip = None
    if 'excel_processed_data_csv_preview' not in st.session_state:
        st.session_state.excel_processed_data_csv_preview = None
    if 'excel_processing_done' not in st.session_state:
        st.session_state.excel_processing_done = False
    if 'excel_arquivos_com_erro_state' not in st.session_state:
        st.session_state.excel_arquivos_com_erro_state = []
    if 'excel_last_uploaded_file_id' not in st.session_state:
        st.session_state.excel_last_uploaded_file_id = None

    uploaded_zip_file = st.file_uploader(
        "Arraste e largue o seu ficheiro ZIP aqui ou clique para procurar",
        type=["zip"],
        accept_multiple_files=False,
        key="zip_uploader_excel"
    )

    # Detect new upload to reset state
    if uploaded_zip_file is not None:
        try:
            current_file_id = getattr(uploaded_zip_file, 'file_id', None) or f"{uploaded_zip_file.name}:{uploaded_zip_file.size}"
        except Exception:
            current_file_id = uploaded_zip_file.name
        if current_file_id != st.session_state.excel_last_uploaded_file_id:
            st.session_state.excel_last_uploaded_file_id = current_file_id
            st.session_state.excel_processing_done = False
            st.session_state.excel_processed_data_zip = None
            st.session_state.excel_processed_data_csv_preview = None
            st.session_state.excel_arquivos_com_erro_state = []
            _safe_rerun()
    elif st.session_state.excel_last_uploaded_file_id is not None:
        st.session_state.excel_last_uploaded_file_id = None
        st.session_state.excel_processing_done = False
        st.session_state.excel_processed_data_zip = None
        st.session_state.excel_processed_data_csv_preview = None
        st.session_state.excel_arquivos_com_erro_state = []
        _safe_rerun()

    if uploaded_zip_file is not None and not st.session_state.excel_processing_done:
        excel_files_in_zip: List[str] = []
        arquivos_com_erro: List[str] = []

        temp_csv_buffer = StringIO()
        header_written = False

        st.info(f"A iniciar o processamento do ficheiro ZIP: **{uploaded_zip_file.name}**...")

        try:
            with zipfile.ZipFile(uploaded_zip_file, 'r') as zf:
                for file_info in zf.infolist():
                    if not file_info.is_dir() and (file_info.filename.lower().endswith('.xls') or file_info.filename.lower().endswith('.xlsx')):
                        excel_files_in_zip.append(file_info.filename)

                if not excel_files_in_zip:
                    st.warning("Nenhum ficheiro Excel (.xls ou .xlsx) encontrado dentro do ficheiro ZIP.")
                    st.session_state.excel_processing_done = False
                    st.session_state.excel_processed_data_zip = None
                    st.session_state.excel_processed_data_csv_preview = None
                    st.session_state.excel_arquivos_com_erro_state = []
                    st.session_state.excel_last_uploaded_file_id = None
                    _safe_rerun()

                st.write(f"Encontrados **{len(excel_files_in_zip)}** ficheiro(s) Excel no ZIP. A iniciar leitura...")
                progress_text = st.empty()
                progress_bar = st.progress(0)

                for i, filename_in_zip in enumerate(excel_files_in_zip):
                    progress_text.text(f"A processar ficheiro {i+1}/{len(excel_files_in_zip)}: **{filename_in_zip}**")
                    try:
                        with zf.open(filename_in_zip) as excel_file_in_zip:
                            excel_content = BytesIO(excel_file_in_zip.read())
                            # Use ExcelFile to iterate sheets, then parse via excel_reader.parse
                            excel_reader = pd.ExcelFile(excel_content)
                            temp_dfs_from_file: List[pd.DataFrame] = []
                            for folha in excel_reader.sheet_names:
                                try:
                                    df_sheet = excel_reader.parse(folha)
                                    temp_dfs_from_file.append(df_sheet)
                                except Exception as e:
                                    arquivos_com_erro.append(f"{filename_in_zip}:{folha} ({e})")
                            if temp_dfs_from_file:
                                df_current_file = pd.concat(temp_dfs_from_file, ignore_index=True)
                                df_current_file.to_csv(
                                    temp_csv_buffer,
                                    sep=';',
                                    mode='a',
                                    header=not header_written,
                                    index=False,
                                    encoding='utf-8',
                                    quoting=csv.QUOTE_MINIMAL,
                                )
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
            st.error("O ficheiro carregado não é um ZIP válido ou está corrompido.")
            st.session_state.excel_processing_done = False
            st.session_state.excel_processed_data_zip = None
            st.session_state.excel_processed_data_csv_preview = None
            st.session_state.excel_arquivos_com_erro_state = []
            st.session_state.excel_last_uploaded_file_id = None
            _safe_rerun()
        except Exception as e:
            st.error(f"Ocorreu um erro inesperado ao processar o ZIP: {e}")
            st.session_state.excel_processing_done = False
            st.session_state.excel_processed_data_zip = None
            st.session_state.excel_processed_data_csv_preview = None
            st.session_state.excel_arquivos_com_erro_state = []
            st.session_state.excel_last_uploaded_file_id = None
            _safe_rerun()

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
                    st.warning(f"Não foi possível gerar a pré-visualização: {e}.")
                    st.session_state.excel_processed_data_csv_preview = None
                    temp_csv_buffer.seek(0)

            with st.spinner("A comprimir o CSV num ficheiro ZIP..."):
                zip_output_buffer = BytesIO()
                with zipfile.ZipFile(zip_output_buffer, 'w', zipfile.ZIP_DEFLATED) as zf_out:
                    zf_out.writestr('resultado_agregado.csv', temp_csv_buffer.getvalue())
                zip_output_buffer.seek(0)
                st.session_state.excel_processed_data_zip = zip_output_buffer.getvalue()

            _safe_rerun()
        else:
            st.error("Nenhum dado válido pôde ser processado dos ficheiros Excel no ZIP.")
            st.session_state.excel_processing_done = False
            st.session_state.excel_processed_data_zip = None
            st.session_state.excel_processed_data_csv_preview = None
            st.session_state.excel_arquivos_com_erro_state = []
            st.session_state.excel_last_uploaded_file_id = None

    if st.session_state.excel_processing_done:
        st.success("Todos os ficheiros Excel válidos foram lidos e agregados!")
        if st.session_state.excel_processed_data_csv_preview is not None:
            st.subheader("Pré-visualização dos Dados Agregados (CSV):")
            st.dataframe(st.session_state.excel_processed_data_csv_preview, use_container_width=True)
        if st.session_state.excel_processed_data_zip is not None:
            st.download_button(
                label="Descarregar Resultado Agregado (resultado_agregado.zip)",
                data=st.session_state.excel_processed_data_zip,
                file_name="resultado_agregado.zip",
                mime="application/zip"
            )
        if st.session_state.excel_arquivos_com_erro_state:
            st.warning("Alguns ficheiros dentro do ZIP tiveram erros:")
            for erro in st.session_state.excel_arquivos_com_erro_state:
                st.write(f"- {erro}")
    elif uploaded_zip_file is None:
        st.info("A aguardar o carregamento de um ficheiro ZIP contendo os ficheiros Excel...")

# ========================================
# SAF-T Faturação -> CSV
# ========================================

def parse_saft_xml_bytes(xml_bytes: bytes) -> Tuple[str, str, str]:
    """Recebe bytes XML e retorna (base_name, customers_csv_text, invoices_csv_text)."""
    try:
        xmlstring = xml_bytes.decode('utf-8')
    except Exception:
        xmlstring = xml_bytes.decode('latin-1', errors='replace')

    # remove default namespace to simplify XPath
    xmlstring = re.sub(r'\sxmlns="[^"]+"', '', xmlstring, count=0)
    root = ET.fromstring(xmlstring)

    customers_buf = io.StringIO()
    invoices_buf = io.StringIO()
    csvwritercustomer = csv.writer(customers_buf, lineterminator='\n')
    csvwriterinvoices = csv.writer(invoices_buf, lineterminator='\n')

    Customer_head: List[str] = []
    count = 0
    for Customer in root.findall('./MasterFiles/Customer'):
        if count == 0:
            if Customer.find('CustomerID') is not None: Customer_head.append('CustomerID')
            if Customer.find('CustomerTaxID') is not None: Customer_head.append('CustomerTaxID')
            if Customer.find('CompanyName') is not None: Customer_head.append('CompanyName')
            if Customer.find('./BillingAddress/Country') is not None: Customer_head.append('Country')
            csvwritercustomer.writerow(Customer_head)
            count += 1
        else:
            ID = Customer.find('CustomerID').text if Customer.find('CustomerID') is not None else ''
            TaxID = Customer.find('CustomerTaxID').text if Customer.find('CustomerTaxID') is not None else ''
            Name = Customer.find('CompanyName').text if Customer.find('CompanyName') is not None else ''
            Country = Customer.find('./BillingAddress/Country').text if Customer.find('./BillingAddress/Country') is not None else ''
            csvwritercustomer.writerow([ID, TaxID, Name, Country])

    Invoices_head: List[str] = []
    first = True
    for Invoice in root.findall('./SourceDocuments/SalesInvoices/Invoice'):
        if first:
            Invoices_head += [
                'InvoiceNo', 'InvoiceStatus', 'Period', 'InvoiceDate', 'InvoiceType', 'CustomerID',
                'ProductCode', 'ProductDescription', 'Quantity', 'UnitOfMeasure', 'UnitPrice', 'Description',
                'Amount', 'TaxAmount', 'TaxCountryRegion', 'Reference', 'Reason'
            ]
            csvwriterinvoices.writerow(Invoices_head)
            first = False
        for Line in Invoice.findall('./Line'):
            row: List[str] = []
            row.append(Invoice.find('InvoiceNo').text if Invoice.find('InvoiceNo') is not None else '')
            row.append(Invoice.find('./DocumentStatus/InvoiceStatus').text if Invoice.find('./DocumentStatus/InvoiceStatus') is not None else '')
            try:
                row.append(Invoice.find('Period').text)
            except Exception:
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
            debit = Line.find('DebitAmount'); credit = Line.find('CreditAmount')
            if debit is not None and debit.text:
                row.append("-" + debit.text)
            elif credit is not None and credit.text:
                row.append(credit.text)
            else:
                row.append('')
            taxamt_el = Line.find('./Tax/TaxAmount'); row.append(taxamt_el.text if taxamt_el is not None else '0')
            taxcr_el = Line.find('./Tax/TaxCountryRegion'); row.append(taxcr_el.text if taxcr_el is not None else '')
            ref_el = Line.find('./References/Reference'); row.append(ref_el.text if ref_el is not None else '')
            reason_el = Line.find('./References/Reason'); row.append(reason_el.text if reason_el is not None else '')
            csvwriterinvoices.writerow(row)

    customers_text = customers_buf.getvalue(); invoices_text = invoices_buf.getvalue()
    customers_buf.close(); invoices_buf.close()
    base_name = "saft_export"
    return base_name, customers_text, invoices_text


def saf_t_tab():
    st.header("SAF-T Faturação → CSV")
    uploaded = st.file_uploader("Escolha um ficheiro .xml ou um .zip contendo .xml", type=["xml", "zip"])
    if uploaded is None:
        st.info("Faça upload de um ficheiro SAF-T (.xml) ou um .zip que contenha um .xml.")
        return
    file_bytes = uploaded.read(); filename = uploaded.name
    xml_bytes = None; xml_name = None
    if filename.lower().endswith('.zip'):
        with io.BytesIO(file_bytes) as bio:
            with zipfile.ZipFile(bio) as z:
                xml_files = [n for n in z.namelist() if n.lower().endswith('.xml')]
                if not xml_files:
                    st.error("O ZIP não contém ficheiros .xml válidos.")
                    return
                xml_name = xml_files[0] if len(xml_files) == 1 else st.selectbox("Selecione o ficheiro XML dentro do ZIP", xml_files)
                xml_bytes = z.read(xml_name)
    else:
        xml_bytes = file_bytes; xml_name = filename

    if st.button("Processar SAF-T"):
        try:
            base_name, customers_csv, invoices_csv = parse_saft_xml_bytes(xml_bytes)
        except Exception as e:
            st.error(f"Erro ao analisar o ficheiro XML: {e}")
            return
        st.subheader("Preview Customers (primeiras 2000 chars)")
        st.code(customers_csv[:2000], language='text')
        st.subheader("Preview Invoices (primeiras 2000 chars)")
        st.code(invoices_csv[:2000], language='text')
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

# ========================================
# OrangeHRM Timesheets Pivot
# ========================================

# PKCE helpers (used in OrangeHRM OAuth bootstrap)
def _pkce_generate_code_verifier(n_bytes: int = 64) -> str:
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
    r = requests.post(token_url, data=data, timeout=30)
    ctype = r.headers.get("Content-Type", "")
    body = r.json() if "application/json" in ctype else {"raw": r.text}
    return r.status_code, body

# OrangeHRM settings (env first, then secrets)
def _get_setting(key: str, default: str = "") -> str:
    env_key = f"ORANGEHRM_{key.upper()}"
    v = os.getenv(env_key)
    if v is not None and isinstance(v, str) and v.strip():
        return v.strip()
    try:
        sv = st.secrets.get("orangehrm", {}).get(key, default)
        return sv.strip() if isinstance(sv, str) else sv
    except Exception:
        return default
        
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
            st.session_state["orange_refresh_token"] = ""

        # Novo: refresh_token partilhado por ficheiro (com fallback às envs)
        refresh_token_env = refresh_token  # vem das envs
        shared_path = os.getenv("ORANGEHRM_REFRESH_TOKEN_FILE", "")
        self._shared_rt_path = shared_path
        rt = _read_shared_refresh_token(shared_path, refresh_token_env)
        st.session_state["orange_refresh_token"] = rt

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
            # Novo: escrever no armazenamento partilhado
            _write_shared_refresh_token(self._shared_rt_path, token_resp["refresh_token"])
        expires_in = int(token_resp.get("expires_in", 3600))
        st.session_state["orange_expires_at"] = time.time() + max(expires_in - 60, 0)

    def _needs_refresh(self) -> bool:
        return (not self.access_token) or time.time() >= self.expires_at

    def _refresh(self) -> bool:
        # Novo: re-sync com armazenamento partilhado antes de chamar o token endpoint
        path = getattr(self, "_shared_rt_path", "")
        current_rt = _read_shared_refresh_token(path, st.session_state.get("orange_refresh_token", ""))
        if current_rt and current_rt != st.session_state.get("orange_refresh_token"):
            st.session_state["orange_refresh_token"] = current_rt

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

_PATH_LIST_EMPLOYEE_TIMESHEETS = "time/employees/{empNumber}/timesheets"
_PATH_GET_TIMESHEET_ENTRIES    = "time/employees/timesheets/{timesheetId}/entries"

# Employees helpers
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
        h = t.get("hours"); m = t.get("minutes")
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
                except Exception:
                    pass
            elif isinstance(dur, (int, float)):
                total_hours += float(dur)
    return round(total_hours, 2)

def _get_totals_by_employee_and_timesheet(client: _OrangeHRMClient, emp_numbers: List[str], empname_map: Dict[str, str], from_date: str = None, to_date: str = None) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
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
                    "totalHours": total_hours,
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
    pivot = pd.pivot_table(df, index=index_col, columns="periodStart", values="totalHours", aggfunc="sum", fill_value=0.0)
    try:
        ordered_cols = sorted(pivot.columns, key=pd.to_datetime)
        pivot = pivot[ordered_cols]
    except Exception:
        pass
    return pivot.sort_index()

@st.cache_data(ttl=600, show_spinner=False)
def _cached_employees_and_map(domain: str, client_id: str, refresh_token: str):
    api_base = domain.rstrip("/") + "/api/v2/"
    token_url = domain.rstrip("/") + "/oauth2/token"
    client = _OrangeHRMClient(client_id, refresh_token, token_url, api_base)
    emps = _list_all_employees(client, limit=200)
    emp_map = _build_empnumber_to_name_map(emps)
    emp_numbers = list(emp_map.keys())
    return emps, emp_map, emp_numbers

# Sum a single entry hours
def _sum_entry_hours(entry: Dict[str, Any]) -> float:
    t = entry.get("total", {})
    h = t.get("hours"); m = t.get("minutes")
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
            except Exception:
                pass
        elif isinstance(dur, (int, float)):
            total += float(dur)
    return total

# Resolve project and customer by project id with cache
def _resolve_project_names_and_customer(
    client: _OrangeHRMClient,
    project_id: Any,
    cache: Dict[str, Dict[str, str]],
) -> Tuple[str, str]:
    pid = str(project_id)
    if pid in cache:
        d = cache[pid]
        return d.get("projectName") or "Sem Projeto", d.get("customerName") or "Sem Cliente"
    project_name, customer_name = "Sem Projeto", "Sem Cliente"
    endpoints = [f"time/projects/{pid}", f"projects/{pid}"]
    for ep in endpoints:
        try:
            data = client.request("GET", ep)
            if isinstance(data, dict):
                proj = data.get("data") or data
                if isinstance(proj, dict):
                    project_name = proj.get("name") or proj.get("projectName") or project_name
                    cust = proj.get("customer")
                    if isinstance(cust, dict):
                        customer_name = cust.get("name") or customer_name
                    else:
                        customer_name = proj.get("customerName") or customer_name
                    break
        except Exception:
            pass
    cache[pid] = {"projectName": project_name, "customerName": customer_name}
    return project_name, customer_name

# Aggregate hours by employee x client x project
def _get_hours_by_employee_client_project(
    client: _OrangeHRMClient,
    emp_numbers: List[str],
    empname_map: Dict[str, str],
    from_date: str = None,
    to_date: str = None,
) -> List[Dict[str, Any]]:
    acc: Dict[tuple, float] = {}
    proj_cust_cache: Dict[str, Dict[str, str]] = {}
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
                    proj = e.get("project") if isinstance(e.get("project"), dict) else {}
                    customer = proj.get("customer") if isinstance(proj.get("customer"), dict) else {}
                    client_name = (
                        customer.get("name")
                        or (proj.get("customer").get("name") if isinstance(proj.get("customer"), dict) else None)
                        or proj.get("customerName")
                        or (e.get("customer").get("name") if isinstance(e.get("customer"), dict) else None)
                        or e.get("customerName")
                        or e.get("customer_name")
                    )
                    project_name = (
                        proj.get("name")
                        or (e.get("project").get("name") if isinstance(e.get("project"), dict) else None)
                        or e.get("projectName")
                        or e.get("project_name")
                    )
                    project_id = proj.get("id") or e.get("projectId") or e.get("project_id")
                    if (not client_name or not project_name) and project_id is not None:
                        p_name, c_name = _resolve_project_names_and_customer(client, project_id, proj_cust_cache)
                        project_name = project_name or p_name
                        client_name = client_name or c_name
                    client_name = client_name or "Sem Cliente"
                    project_name = project_name or "Sem Projeto"
                    key = (str(emp), str(client_name), str(project_name))
                    acc[key] = acc.get(key, 0.0) + hours
            if len(sheets) < 50:
                break
            offset += 50
    rows: List[Dict[str, Any]] = []
    for (emp_key, client_name, project_name), total in acc.items():
        rows.append({
            "empNumber": emp_key,
            "empName": empname_map.get(emp_key, emp_key),
            "clientName": client_name,
            "projectName": project_name,
            "totalHours": round(total, 2),
        })
    return rows

# Map O365 email -> OrangeHRM employee number
def _email_to_username(email: str) -> str:
    try:
        local = (email or "").strip().split("@", 1)[0]
        return local.strip().lower()
    except Exception:
        return ""

def _try_map_by_admin_users(client: _OrangeHRMClient, username: str) -> Optional[str]:
    """
    Mapeia via /api/v2/admin/users tentando vários filtros comuns.
    Requer permissões para admin/users. Extrai empNumber de u["employee"].
    """
    if not username:
        return None

    def _extract_emp_no(u: dict) -> Optional[str]:
        emp = u.get("employee") if isinstance(u.get("employee"), dict) else {}
        emp_no = emp.get("empNumber") or u.get("empNumber") or None
        return str(emp_no) if emp_no else None

    # 1) Filtro direto por userName
    try:
        data = client.request("GET", f"admin/users?limit=50&userName={username}")
        arr = data.get("data") if isinstance(data, dict) else (data if isinstance(data, list) else [])
        for u in arr:
            emp_no = _extract_emp_no(u)
            if emp_no:
                return emp_no
    except Exception:
        pass

    # 2) Tentativa por employeeEmail (se a tua API suportar este filtro)
    try:
        # Se soubermos o email completo podemos tentar aqui também, mas como recebemos só username,
        # este passo é opcional e pode ser removido. Mantemos apenas por compatibilidade nalgumas instâncias.
        # data = client.request("GET", f"admin/users?limit=50&employeeEmail={username}@inobest.com")
        # ...
        pass
    except Exception:
        pass

    # 3) Paginar admin/users sem filtro e procurar username localmente
    #    (só se o token tiver permissão, e pode ser pesado; limitamos a algumas páginas)
    try:
        limit, offset, scanned = 100, 0, 0
        max_scan = 1000  # salvaguarda
        while scanned < max_scan:
            data = client.request("GET", f"admin/users?limit={limit}&offset={offset}")
            arr = data.get("data") if isinstance(data, dict) else (data if isinstance(data, list) else [])
            if not arr:
                break
            for u in arr:
                scanned += 1
                u_name = (u.get("userName") or u.get("username") or "").strip().lower()
                if u_name == username:
                    emp_no = _extract_emp_no(u)
                    if emp_no:
                        return emp_no
            if len(arr) < limit:
                break
            offset += limit
    except Exception:
        pass

    return None

def _map_email_to_empnumber(client: _OrangeHRMClient, email: str) -> Optional[str]:
    if not email:
        return None
    email_l = email.strip().lower()

    # 1) Tentativa por e-mail (workEmail) via PIM
    try:
        data = client.request("GET", f"pim/employees?limit=50&email={email_l}")
        if isinstance(data, dict):
            arr = data.get("data") or data.get("employees") or []
            for emp in arr:
                work = (emp.get("workEmail") or emp.get("email") or "").strip().lower()
                if work == email_l and emp.get("empNumber"):
                    return str(emp["empNumber"])
    except Exception:
        pass

    # 2) Tentativa por username = prefixo do e-mail (ex.: ricardo.sousa)
    username = _email_to_username(email_l)
    emp_no = _try_map_by_admin_users(client, username)
    if emp_no:
        return emp_no

    # 3) Fallback: listar todos via PIM e comparar pelo workEmail (caso 1 não resulte)
    try:
        employees, _emp_map, _all = _cached_employees_and_map(
            _get_setting("domain") or "", _get_setting("client_id") or "", _get_setting("refresh_token") or ""
        )
        for emp in employees:
            work = (emp.get("workEmail") or emp.get("email") or "").strip().lower()
            if work == email_l and emp.get("empNumber"):
                return str(emp["empNumber"])
    except Exception:
        pass

    return None

# =============================
# OrangeHRM OAuth Bootstrap Tab
# =============================

def render_orangehrm_oauth_bootstrap_tab():
    st.header("Configuração OAuth (Admin) — Obter Refresh Token")

    status = _ensure_oauth_admin()
    # Só renderiza o conteúdo da aba se houver acesso
    if status != "ok":
        # "needs-auth" mostra o formulário acima; "no-access" mostra erro.
        # Em ambos os casos, apenas não renderizamos o resto desta aba.
        return

    # (continua daqui para baixo o conteúdo atual da aba: leitura de domain, client_id, PKCE, etc.)
    domain = _get_setting("domain") or "https://rh.inobest.com/web/index.php/"
    client_id = _get_setting("client_id") or ""
    default_redirect = domain.rstrip("/")
    token_url = domain.rstrip("/") + "/oauth2/token"
    auth_base = domain.rstrip("/") + "/oauth2/authorize"

    st.markdown("Use esta aba apenas para gerar um refresh_token e colocá-lo no stack.env.")

    with st.form("oauth_bootstrap_form"):
        c1, c2 = st.columns([1, 1])
        with c1:
            client_id_in = st.text_input("Client ID", value=client_id, disabled=True)
            redirect_uri = st.text_input("Redirect URI", value=default_redirect, disabled=True)
        with c2:
            st.write(" ")

        #if "pkce_code_verifier" not in st.session_state or st.form_submit_button("Gerar novo code_verifier"):
        #    st.session_state["pkce_code_verifier"] = _pkce_generate_code_verifier()
        code_verifier = st.session_state["pkce_code_verifier"]
        code_challenge = _pkce_code_challenge(code_verifier)

        st.text_input("Code Verifier", value=code_verifier, disabled=True)
        st.text_input("Code Challenge (S256)", value=code_challenge, disabled=True)

        params = {
            "response_type": "code",
            "code_challenge_method": "S256",
            "code_challenge": code_challenge,
            "client_id": client_id_in,
            "redirect_uri": redirect_uri,
        }
        auth_url = auth_base + "?" + urlencode(params)
        st.write("1) Clique no link abaixo, autentique e autorize. Será redirecionado com ?code=...")
        st.write(f"[Abrir e autorizar no OrangeHRM]({auth_url})")

        auth_code = st.text_input("2) Cole aqui o 'code' devolvido", value="")
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
                        "access_token_prefix": (body.get("access_token", "")[:16] + "...") if body.get("access_token") else None,
                        "expires_in": body.get("expires_in"),
                        "refresh_token_prefix": (body.get("refresh_token", "")[:12] + "...") if body.get("refresh_token") else None,
                        "token_type": body.get("token_type"),
                        "scope": body.get("scope"),
                    })
                    st.code(body.get("refresh_token", ""), language="text")
                    st.warning("Copie o refresh_token acima e configure ORANGEHRM_REFRESH_TOKEN no stack.env. Depois redeploy.")
                    shared_path = os.getenv("ORANGEHRM_REFRESH_TOKEN_FILE", "")
                    rt_new = body.get("refresh_token", "")
                    if shared_path and rt_new:
                        try:
                            _write_shared_refresh_token(shared_path, rt_new)
                            st.success(f"Refresh token gravado no ficheiro partilhado: {shared_path}")
                        except Exception as e:
                            st.warning(f"Não foi possível gravar no ficheiro partilhado ({shared_path}): {e}")
                    else:
                        if not shared_path:
                            st.info("Defina ORANGEHRM_REFRESH_TOKEN_FILE para gravar automaticamente o refresh token num ficheiro partilhado.")    
                else:
                    st.error(f"Falha ao obter tokens (HTTP {status}).")
                    st.write("Resposta do servidor:")
                    st.json(body)

# =============================
# Timesheets Pivot Tab
# =============================

def render_orangehrm_pivot_tab():
    domain = _get_setting("domain") or "https://rh.inobest.com/web/index.php/"
    client_id = _get_setting("client_id") or ""
    refresh_token = _get_setting("refresh_token") or ""
    api_base = domain.rstrip("/") + "/api/v2/"
    token_url = domain.rstrip("/") + "/oauth2/token"

    st.header("Folhas de Horas — Tabela Dinâmica por Colaborador")

    c1, c2, c3 = st.columns([1, 1, 0.6])
    with c1:
        from_date = st.date_input("De (início do período)", None, format="YYYY-MM-DD", key="ts_from_date")
    with c2:
        to_date = st.date_input("Até (fim do período)", None, format="YYYY-MM-DD", key="ts_to_date")
    with c3:
        run_btn = st.button("Gerar Tabela Dinâmica", key="run_pivot_btn")
    from_date_str = from_date.isoformat() if from_date else None
    to_date_str = to_date.isoformat() if to_date else None

    if not (client_id and refresh_token):
        st.error("Credenciais do serviço em falta (ORANGEHRM_CLIENT_ID/REFRESH_TOKEN).")
        return

    service_client = _OrangeHRMClient(client_id, refresh_token, token_url, api_base)

    if _is_admin_o365():
        with st.spinner("A carregar colaboradores..."):
            employees, emp_map, all_emp_numbers = _cached_employees_and_map(domain, client_id, refresh_token)
        emp_display = {k: emp_map.get(k, k) for k in all_emp_numbers}
        emp_choices = st.multiselect(
            "Filtrar colaboradores (opcional)",
            options=all_emp_numbers,
            default=all_emp_numbers[: min(len(all_emp_numbers), 20)],
            format_func=lambda k: emp_display.get(k, k),
            key="emp_filter_multiselect",
        )
        client_for_calls = service_client
    else:
        email = st.session_state.get("o365_auth", {}).get("email", "")
        if not email:
            st.error("Não foi possível identificar o seu e‑mail.")
            return
        with st.spinner("A identificar o seu número de colaborador..."):
            emp_number = _map_email_to_empnumber(service_client, email)
        if not emp_number:
            st.error("Não foi possível associar o seu e‑mail ao registo no OrangeHRM. Contacte o administrador.")
            return
        emp_choices = [str(emp_number)]
        emp_map = {str(emp_number): st.session_state.get("o365_auth", {}).get("name") or "Eu"}
        client_for_calls = service_client

    if run_btn:
        with st.spinner("A obter folhas de horas e a calcular totais..."):
            rows = _get_totals_by_employee_and_timesheet(
                client_for_calls, emp_choices, emp_map, from_date=from_date_str, to_date=to_date_str
            )
            pivot_df = _pivot_hours_by_employee_and_start(rows)
        st.subheader("Horas por Colaborador × Data de Início da Folha de Horas")
        if pivot_df.empty:
            st.info("Não foram encontrados dados para os filtros selecionados.")
        else:
            st.dataframe(pivot_df, use_container_width=True)
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

        st.subheader("Cliente/Projeto por Colaborador — Totais de Horas")
        with st.spinner("A obter entradas e a agregar por cliente e projeto..."):
            ecp_rows = _get_hours_by_employee_client_project(
                client_for_calls,
                emp_choices,
                emp_map,
                from_date=from_date_str,
                to_date=to_date_str,
            )
            ecp_df = pd.DataFrame(ecp_rows)
        if ecp_df.empty:
            st.info("Não foram encontrados dados por cliente/projeto para os filtros selecionados.")
        else:
            ecp_df = ecp_df.sort_values(by=["empName", "clientName", "totalHours"], ascending=[True, True, False])
            for emp, g in ecp_df.groupby("empName", sort=False):
                st.markdown(f"### {emp}")
                show = g[["clientName", "projectName", "totalHours"]].rename(columns={
                    "clientName": "Cliente",
                    "projectName": "Projeto",
                    "totalHours": "Total de Horas",
                })
                st.dataframe(show.reset_index(drop=True), use_container_width=True)
            ecp_xlsx_buffer = BytesIO()
            with pd.ExcelWriter(ecp_xlsx_buffer, engine="openpyxl") as writer:
                ecp_df.rename(columns={
                    "empName": "Colaborador",
                    "clientName": "Cliente",
                    "projectName": "Projeto",
                    "totalHours": "TotalHoras",
                })[["Colaborador", "Cliente", "Projeto", "TotalHoras"]].to_excel(writer, sheet_name="Cliente_Projeto_por_Colab", index=False)
                for emp, g in ecp_df.groupby("empName", sort=False):
                    g.rename(columns={
                        "clientName": "Cliente",
                        "projectName": "Projeto",
                        "totalHours": "TotalHoras",
                    })[["Cliente", "Projeto", "TotalHoras"]].to_excel(
                        writer, sheet_name=(str(emp)[:28] or "Sem_Nome"), index=False
                    )
            st.download_button(
                "Descarregar Excel (Cliente/Projeto por Colaborador)",
                data=ecp_xlsx_buffer.getvalue(),
                file_name="folhas_horas_cliente_projeto_por_colaborador.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_xlsx_cliente_projeto_por_colab",
            )

# =============================
# Tabs
# =============================

tab1, tab2, tab3, tab4 = st.tabs(["Agregador de Excel", "SAF-T Faturação → CSV", "Configuração OAuth (Admin)", "Timesheets Pivot"])
with tab1:
    excel_aggregator_app()
with tab2:
    saf_t_tab()
with tab3:
    render_orangehrm_oauth_bootstrap_tab()
with tab4:
    render_orangehrm_pivot_tab()
