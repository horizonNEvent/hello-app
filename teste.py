import streamlit as st
import pandas as pd
import re
from datetime import datetime, timedelta
import io

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

                # Criar o DataFrame de saída com as colunas na ordem correta
                colunas_saida = [
                    'Identificação do tipo de integração de título',
                    'Codigo do Fornecedor',
                    'Numero do Titulo',
                    'Documento Fiscal',
                    'Empresa Emitente',
                    'Codigo da Filial',
                    'Empresa Pagadora',
                    'Tipo de Titulo',
                    'Data de Emissao do Titulo',
                    'Data de Vencimento do Titulo',
                    'Data de Programacao do Titulo',
                    'Codigo da Moeda',
                    'Tipo de Cobranca',
                ]

                df_saida = pd.DataFrame(columns=colunas_saida)

                # Preencher as colunas
                df_saida['Identificação do tipo de integração de título'] = ['PP'] * len(df)
                df_saida['Codigo do Fornecedor'] = df[colunas['CNPJ/CPF']].apply(limpar_cnpj_cpf)
                df_saida['Numero do Titulo'] = df[colunas['Nº DOCTO']].apply(limpar_numero_documento)
                df_saida['Documento Fiscal'] = df[colunas['Nº DOCTO']].apply(limpar_numero_documento)
                df_saida['Empresa Emitente'] = ['0001'] * len(df)
                df_saida['Codigo da Filial'] = ['01'] * len(df)
                df_saida['Empresa Pagadora'] = ['0001'] * len(df)
                df_saida['Tipo de Titulo'] = ['55'] * len(df)
                df_saida['Data de Emissao do Titulo'] = df[colunas['DATA DA ENTRADA']].apply(formatar_data)
                df_saida['Data de Vencimento do Titulo'] = df[colunas['VENCTO']].apply(formatar_data)
                df_saida['Data de Programacao do Titulo'] = df[colunas['VENCTO']].apply(formatar_data)
                df_saida['Codigo da Moeda'] = ['BRL'] * len(df)
                df_saida['Tipo de Cobranca'] = ['CA'] * len(df)

                # Adicionar as colunas nas posições específicas
                df_saida.insert(loc=df_saida.columns.get_loc('Codigo da Moeda') + 1, 
                                column='Grupo de Pagamento', 
                                value=df[colunas['FORNECEDOR']].apply(determinar_grupo_pagamento))
                
                df_saida.insert(loc=df_saida.columns.get_loc('Grupo de Pagamento') + 1, 
                                column='Valor do Grupo de Pagamento', 
                                value=df[colunas['VALOR']].apply(formatar_valor))
                
                df_saida.insert(loc=df_saida.columns.get_loc('Valor do Grupo de Pagamento') + 1, 
                                column='Codigo do fluxo de caixa', 
                                value=['01'] * len(df))

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
