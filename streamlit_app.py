import streamlit as st
import pandas as pd
import os
from io import BytesIO, StringIO # <--- Esta linha já importa BytesIO e StringIO do módulo io
import zipfile
import csv
import fitz
from PyPDF2 import PdfReader, PdfWriter
import io # <--- ADICIONA ESTA LINHA SE AINDA NÃO ESTIVER LÁ, OU SE ESTIVER, GARANTE QUE ESTÁ NO TOPO

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

# --- Função para o Redutor de PDF (COM A CORREÇÃO DO ARGUMENTO 'compress') ---
def reduce_pdf_size(input_file, output_path=None):
    try:
        # Se não for especificado um caminho de saída, criar um baseado no original
        if output_path is None:
            file_name = os.path.basename(input_file.name)
            base_name, ext = os.path.splitext(file_name)
            output_path = f"{base_name}_reduced{ext}"
        
        # Ler o arquivo de entrada
        pdf_data = input_file.read()
        
        # Abrir o documento com PyMuPDF
        doc = fitz.open(stream=pdf_data, filetype="pdf")
        
        # Aplicar compressão a cada página
        for page in doc:
            # Limpar conteúdo redundante
            page.clean_contents()
            
            # Comprimir imagens na página
            xref = page.get_images(full=True)
            for img in xref:
                xref_obj = img[0]
                try:
                    # Tentar comprimir a imagem se possível
                    pix = fitz.Pixmap(doc, xref_obj)
                    if pix.colorspace and pix.colorspace.n >= 3:  # Se for colorida
                        if pix.alpha:  # Se tiver canal alpha
                            pix = fitz.Pixmap(fitz.csRGB, pix)
                        # Reduzir a qualidade da imagem
                        pix = fitz.Pixmap(pix, 0.5)  # Reduz para 50% da qualidade
                        doc.update_stream(xref_obj, pix.tobytes())
                except Exception:
                    pass  # Ignorar erros em imagens específicas
        
        # Salvar com opções de compressão
        doc.save(output_path, garbage=4, deflate=True, clean=True)
        doc.close()
        
        # Obter tamanhos para comparação
        original_size = len(pdf_data)
        with open(output_path, "rb") as f:
            reduced_size = len(f.read())
        
        reduction_percentage = ((original_size - reduced_size) / original_size) * 100
        
        return {
            "success": True,
            "output_path": output_path,
            "original_size": original_size,
            "reduced_size": reduced_size,
            "reduction_percentage": reduction_percentage
        }
    except Exception as e:
        return {
            "success": False,
            "error": str(e)
        }


# --- Lógica Principal da Aplicação (Usando Abas) ---
tab1, tab2 = st.tabs(["Agregador de Excel", "Redutor de PDF"])

with tab1:
    excel_aggregator_app()

with tab2:
    pdf_reducer_app()
