import streamlit as st
import pandas as pd
import re
from datetime import datetime, timedelta
import io

# Funções auxiliares
def limpar_cnpj_cpf(valor):
    if pd.isna(valor):
        return ''
    return re.sub(r'\D', '', str(valor)).zfill(14)

def formatar_data(data):
    if pd.isna(data):
        return ''
    if isinstance(data, (int, float)):
        data = (datetime(1899, 12, 30) + timedelta(days=int(data))).date()
    return data.strftime('%d%m%Y') if isinstance(data, datetime) else str(data)

def formatar_valor(valor):
    if pd.isna(valor):
        return '0,00'
    return f"{valor:.2f}".replace('.', ',')

def determinar_grupo_pagamento(fornecedor):
    if 'BEBIDAS' in str(fornecedor).upper() or 'VINHO' in str(fornecedor).upper():
        return '1106020000'
    return '1106010000'

def limpar_numero_documento(valor):
    if pd.isna(valor):
        return ''
    valor_str = str(valor).strip()
    # Tenta remover a parte de hora se existir
    if ' 00:00:00' in valor_str:
        valor_str = valor_str.split(' ')[0]
    # Se for uma data no formato yyyy-mm-dd, converte para dd-mm-yyyy
    if re.match(r'\d{4}-\d{2}-\d{2}', valor_str):
        partes = valor_str.split('-')
        valor_str = f"{partes[2]}-{partes[1]}-{partes[0]}"
    return valor_str

def main():
    st.title("Conversor de Planilha")

    uploaded_file = st.file_uploader("Escolha uma planilha", type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file, header=6)
            st.success("Planilha carregada com sucesso!")

            # Mapear as colunas
            colunas = {
                'CNPJ/CPF': df.columns[17],  # Coluna R
                'FORNECEDOR': df.columns[16],  # Coluna Q
                'VALOR': df.columns[8],  # Coluna I
                'DATA DA ENTRADA': df.columns[6],  # Coluna G
                'VENCTO': df.columns[4],  # Coluna E
                'Nº DOCTO': df.columns[2],  # Coluna C
            }

            st.write("Colunas mapeadas:", colunas)

            if st.button("Processar Planilha"):
                # Remover a última linha se for a linha de total
                if pd.isna(df.iloc[-1][colunas['CNPJ/CPF']]) or 'TOTAL' in str(df.iloc[-1][colunas['FORNECEDOR']]).upper():
                    df = df.iloc[:-1]

                # Criar um dicionário para o DataFrame de saída
                dados_saida = {
                    'A': ['PP'] * len(df),  # Identificação do tipo de integração de título
                    'B': df[colunas['CNPJ/CPF']].apply(limpar_cnpj_cpf),  # Codigo do Fornecedor
                    'C': df[colunas['Nº DOCTO']].apply(limpar_numero_documento),  # Numero do Titulo
                    'D': df[colunas['Nº DOCTO']].apply(limpar_numero_documento),  # Documento Fiscal
                    'E': ['0001'] * len(df),  # Empresa Emitente
                    'F': ['0001'] * len(df),  # Codigo da Filial
                    'G': ['0001'] * len(df),  # Empresa Pagadora
                    'H': ['55'] * len(df),  # Tipo de Titulo
                    'I': df[colunas['DATA DA ENTRADA']].apply(formatar_data),  # Data de Emissao do Titulo
                    'J': df[colunas['VENCTO']].apply(formatar_data),  # Data de Vencimento do Titulo
                    'K': df[colunas['VENCTO']].apply(formatar_data),  # Data de Programacao do Titulo
                    'L': ['BRL'] * len(df),  # Codigo da Moeda
                    'M': ['CA'] * len(df),  # Tipo de Cobranca
                }

                # Adicionar as colunas específicas
                dados_saida['CE'] = df[colunas['FORNECEDOR']].apply(determinar_grupo_pagamento)  # Grupo de Pagamento
                dados_saida['CG'] = df[colunas['VALOR']].apply(formatar_valor)  # Valor do Grupo de Pagamento
                dados_saida['CJ'] = ['01'] * len(df)  # Codigo do fluxo de caixa

                # Criar o DataFrame de saída
                df_saida = pd.DataFrame(dados_saida)

                st.write("Preview dos dados convertidos:")
                st.dataframe(df_saida)

                # Botão para download
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_saida.to_excel(writer, index=False)
                st.download_button(
                    label="Download planilha convertida",
                    data=output.getvalue(),
                    file_name="planilha_convertida.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"Erro ao processar a planilha: {str(e)}")
            st.write(f"Erro detalhado: {e}")

if __name__ == "__main__":
    main()
