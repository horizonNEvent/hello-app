import streamlit as st
import pandas as pd
import re
from datetime import datetime, timedelta
import io
import string

# Funções auxiliares (mantenha as funções existentes)
def limpar_cnpj_cpf(valor):
    if pd.isna(valor) or valor == '':
        return None
    valor_limpo = re.sub(r'\D', '', str(valor))
    valor_formatado = valor_limpo.zfill(14)
    return valor_formatado[-14:]

def formatar_data(data):
    if pd.isna(data):
        return None
    if isinstance(data, (int, float)):
        try:
            return (datetime(1899, 12, 30) + timedelta(days=int(data))).strftime('%d%m%Y')
        except ValueError:
            return str(data)
    elif isinstance(data, str):
        try:
            return pd.to_datetime(data).strftime('%d%m%Y')
        except:
            return data
    elif isinstance(data, pd.Timestamp) or isinstance(data, datetime):
        return data.strftime('%d%m%Y')
    else:
        return str(data)

def formatar_valor(valor):
    if pd.isna(valor) or valor == '':
        return None
    return str(valor).replace('.', ',')

def determinar_grupo_pagamento(fornecedor):
    if isinstance(fornecedor, str) and ('BEBIDAS' in fornecedor.upper() or 'VINHO' in fornecedor.upper()):
        return '1106020000'
    return '1106010000'

def limpar_numero_documento(valor):
    if pd.isna(valor) or valor == '':
        return None
    valor_str = str(valor)
    
    if re.match(r'\d{2}/\d{2}/\d{4}', valor_str):
        return valor_str
    
    if ' 00:00:00' in valor_str:
        valor_str = valor_str.split(' ')[0]
    
    if re.match(r'\d{4}-\d{2}-\d{2}', valor_str):
        partes = valor_str.split('-')
        return f"{partes[2]}/{partes[1]}/{partes[0]}"
    
    return valor_str

# Função principal do Streamlit
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

                # Criar o DataFrame de saída com 100 colunas vazias
                colunas_saida = [f'Coluna_{i}' for i in range(100)]
                df_saida = pd.DataFrame(columns=colunas_saida)

                # Preencher as colunas nas posições corretas
                df_saida['Coluna_0'] = ['PP'] * len(df)  # Identificação do tipo de integração de título
                df_saida['Coluna_1'] = df[colunas['CNPJ/CPF']].apply(limpar_cnpj_cpf)  # Codigo do Fornecedor
                df_saida['Coluna_2'] = df[colunas['Nº DOCTO']].apply(limpar_numero_documento)  # Numero do Titulo
                df_saida['Coluna_3'] = df[colunas['Nº DOCTO']].apply(limpar_numero_documento)  # Documento Fiscal
                df_saida['Coluna_4'] = ['0001'] * len(df)  # Empresa Emitente
                df_saida['Coluna_5'] = ['0001'] * len(df)  # Codigo da Filial
                df_saida['Coluna_6'] = ['0001'] * len(df)  # Empresa Pagadora
                df_saida['Coluna_7'] = ['55'] * len(df)  # Tipo de Titulo
                df_saida['Coluna_8'] = df[colunas['DATA DA ENTRADA']].apply(formatar_data)  # Data de Emissao do Titulo
                df_saida['Coluna_9'] = df[colunas['VENCTO']].apply(formatar_data)  # Data de Vencimento do Titulo
                df_saida['Coluna_10'] = df[colunas['VENCTO']].apply(formatar_data)  # Data de Programacao do Titulo
                df_saida['Coluna_11'] = ['BRL'] * len(df)  # Codigo da Moeda
                df_saida['Coluna_12'] = ['CA'] * len(df)  # Tipo de Cobranca

                # Adicionar as colunas nas posições específicas
                df_saida['Coluna_82'] = df[colunas['FORNECEDOR']].apply(determinar_grupo_pagamento)  # CE: Grupo de Pagamento
                df_saida['Coluna_84'] = df[colunas['VALOR']].apply(formatar_valor)  # CG: Valor do Grupo de Pagamento
                df_saida['Coluna_87'] = ['01'] * len(df)  # CJ: Codigo do fluxo de caixa

                # Renomear as colunas para letras
                letras = list(string.ascii_uppercase) + [f'A{letra}' for letra in string.ascii_uppercase] + [f'B{letra}' for letra in string.ascii_uppercase]
                df_saida.columns = letras[:len(df_saida.columns)]

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
