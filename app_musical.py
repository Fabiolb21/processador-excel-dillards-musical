import streamlit as st
import pandas as pd
import numpy as np
import os
import xlwt
from io import BytesIO

st.set_page_config(page_title="Processador Musical", layout="wide")

st.title("Processador de Planilhas - Dillard's Musical")

st.markdown("""
Esta aplicação processa uma planilha Excel para extrair e transformar dados, gerando uma planilha formatada.

**Instruções:**
1. Faça o upload do seu arquivo Excel (formato `.xlsx`).
2. Aguarde o processamento dos dados.
3. Faça o download da planilha gerada.
""")

st.divider()


# Função para processar o arquivo
def process_excel(uploaded_file):
    try:
        # Ler o arquivo Excel
        df = pd.read_excel(uploaded_file, engine='openpyxl')
        # st.success(f"✓ Arquivo Excel lido com sucesso! ({len(df)} linhas encontradas)")

        # Identificar as colunas
        primeira_coluna = df.columns[0]  # PO 3173441000
        cod_barras_col = df.columns[1]   # CÓD DE BARRAS
        marca_col = df.columns[2]        # Unnamed: 2 (terceira coluna - MARCA)
        modelo_col = df.columns[3]       # MODELO
        style_col = df.columns[4]        # STYLE
        cor_col = df.columns[5]          # COR
        dpto_col = df.columns[6]         # DPTO
        mic_col = df.columns[7]          # MIC
        size_col = df.columns[9]         # SIZE
        quant_ped_col = df.columns[10]   # QUANT DO PED

        # Criar lista para armazenar as linhas expandidas
        linhas_expandidas = []
        num_etq_counter = 1  # Contador para NUM DA ETQ

        # Processar cada linha
        for idx, row in df.iterrows():
            # Obter valores da linha
            # Garantir 13 caracteres com zero na frente
            primeira_col_valor = str(int(row[primeira_coluna])).zfill(13) if pd.notna(row[primeira_coluna]) else ''
            cod_barras = str(int(row[cod_barras_col])).zfill(13) if pd.notna(row[cod_barras_col]) else ''
            marca = row[marca_col] if pd.notna(row[marca_col]) else ''
            modelo = row[modelo_col] if pd.notna(row[modelo_col]) else ''
            style = row[style_col] if pd.notna(row[style_col]) else ''
            cor = row[cor_col] if pd.notna(row[cor_col]) else ''
            dpto = row[dpto_col] if pd.notna(row[dpto_col]) else ''
            mic = row[mic_col] if pd.notna(row[mic_col]) else ''
            size = row[size_col] if pd.notna(row[size_col]) else ''
            qtd_ped = row[quant_ped_col] if pd.notna(row[quant_ped_col]) else 0
            
            # Calcular QTD EXTRA (QTD DO PED + 1)
            qtd_extra = int(qtd_ped) + 1 if pd.notna(qtd_ped) else 1
            
            # Calcular PREFIXO DA EMP (7 primeiros caracteres do COD DE BARRAS de 13 chars)
            prefixo_emp = cod_barras[:7] if len(cod_barras) >= 7 else cod_barras
            
            # Calcular ITEM DE REF (do 8º ao 12º caractere + zero na frente)
            if len(cod_barras) >= 13:
                item_ref = '0' + cod_barras[7:12]
            else:
                item_ref = ''
            
            # Gerar linhas de acordo com QTD EXTRA
            for i in range(qtd_extra):
                nova_linha = {
                    primeira_coluna: primeira_col_valor,
                    'COD DE BARRAS': cod_barras,
                    'MARCA': marca,
                    'MODELO': modelo,
                    'STYLE': style,
                    'COR': cor,
                    'DPTO': dpto,
                    'MIC': mic,
                    'SIZE': size,
                    'QTD DO PED': int(qtd_ped) if pd.notna(qtd_ped) else '',
                    'QTD EXTRA': qtd_extra,
                    'NUM DA ETQ': str(num_etq_counter).zfill(10),  # Formato 0000000001
                    'VALOR DO FILTRO': 1,
                    'PREFIXO DA EMP': prefixo_emp,
                    'ITEM DE REF': item_ref,
                    'SERIAL': ''
                }
                linhas_expandidas.append(nova_linha)
                num_etq_counter += 1

        # Criar DataFrame com as linhas expandidas
        df_resultado = pd.DataFrame(linhas_expandidas)
        
        # st.success(f"✓ Total de linhas processadas: {len(df_resultado)}")

        # Gerar planilha única
        output_dir = "output_files"
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        # Converter todas as colunas para string
        df_resultado_str = df_resultado.copy()
        for col in df_resultado_str.columns:
            df_resultado_str[col] = df_resultado_str[col].apply(
                lambda x: '' if pd.isna(x) or str(x).lower() == 'nan' else str(x)
            )

        # Criar arquivo .xls usando xlwt
        output_filename = os.path.join(output_dir, "planilha_musical.xls")
        
        workbook = xlwt.Workbook()
        worksheet = workbook.add_sheet('Dados')
        
        # Criar estilo de texto para forçar formato de texto
        text_style = xlwt.XFStyle()
        text_style.num_format_str = '@'  # @ é o código para formato de texto no Excel

        # Escrever cabeçalhos
        for col_idx, col_name in enumerate(df_resultado_str.columns):
            worksheet.write(0, col_idx, col_name)

        # Escrever dados (todos como texto)
        for row_idx, row in enumerate(df_resultado_str.values, start=1):
            for col_idx, value in enumerate(row):
                # Garantir que tudo seja escrito como texto
                val_str = str(value) if value != '' and str(value).lower() != 'nan' else ''
                worksheet.write(row_idx, col_idx, val_str, text_style)

        workbook.save(output_filename)
        # st.success(f"✓ Planilha gerada com sucesso!")
        
        return output_filename, df_resultado

    except Exception as e:
        st.error(f"❌ Ocorreu um erro: {e}")
        import traceback
        st.code(traceback.format_exc())
        return None, None


# Interface do Streamlit
col1, col2 = st.columns([2, 1])

with col1:
    uploaded_file = st.file_uploader("📁 Escolha um arquivo Excel", type=["xlsx"])

with col2:
    st.markdown("### Informações")
    st.markdown("""
    **Formato aceito:** `.xlsx`
    
    **Processamento:** Expansão de linhas conforme quantidade
    """)

if uploaded_file is not None:
    st.divider()
    
    if st.button("🚀 Processar Arquivo", type="primary", use_container_width=True):
        with st.spinner('Processando...'):
            output_file, df_resultado = process_excel(uploaded_file)

            if output_file:
                st.divider()
                st.subheader("📦 Download do Resultado")
                
                # Download do arquivo
                with open(output_file, "rb") as fp:
                    st.download_button(
                        label="⬇️ Download Planilha (.xls)",
                        data=fp,
                        file_name="planilha_musical.xls",
                        mime="application/vnd.ms-excel",
                        type="primary",
                        use_container_width=True
                    )
                
                # Mostrar preview dos dados
                with st.expander("👁️ Visualizar Preview dos Dados Processados"):
                    st.dataframe(df_resultado.head(100), use_container_width=True)

st.divider()

