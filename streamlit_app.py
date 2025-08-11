import streamlit as st
import io
import zipfile
import xml.etree.ElementTree as ET
import re
import csv
from typing import Tuple

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
    st.header("SAF-T → CSV (in-memory)")
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


# --- Lógica Principal da Aplicação (Usando Abas) ---
tab1, tab2 = st.tabs(["App principal", "SAF-T → CSV"])
with tab2:
    saf_t_tab()
