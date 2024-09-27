import sys
import pandas as pd
import re
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, QFileDialog, 
                             QVBoxLayout, QHBoxLayout, QWidget, QTableWidget, QTableWidgetItem, QMessageBox)
from PyQt5.QtCore import Qt
from collections import OrderedDict
from datetime import datetime, timedelta

class PlanilhaLeitor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Conversor de Planilha")
        self.setGeometry(100, 100, 1200, 800)

        layout = QVBoxLayout()

        botoes_layout = QHBoxLayout()
        self.botao_selecionar = QPushButton("Selecionar Planilha")
        self.botao_selecionar.clicked.connect(self.selecionar_planilha)
        botoes_layout.addWidget(self.botao_selecionar)

        self.botao_preview = QPushButton("Preview")
        self.botao_preview.clicked.connect(self.preview_conversao)
        self.botao_preview.setEnabled(False)
        botoes_layout.addWidget(self.botao_preview)

        self.botao_converter = QPushButton("Converter e Salvar")
        self.botao_converter.clicked.connect(self.converter_e_salvar)
        self.botao_converter.setEnabled(False)
        botoes_layout.addWidget(self.botao_converter)

        layout.addLayout(botoes_layout)

        self.tabela = QTableWidget()
        layout.addWidget(self.tabela)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

        self.df = None
        self.df_saida = None

    def limpar_cnpj_cpf(self, valor):
        if pd.isna(valor) or valor == '':
            return None
        # Remove caracteres não numéricos
        valor_limpo = re.sub(r'\D', '', str(valor))
        # Preenche com zeros à esquerda se necessário para manter 14 dígitos
        valor_formatado = valor_limpo.zfill(14)
        # Garante que o valor tenha 14 dígitos, mesmo se já tiver mais
        return valor_formatado[-14:]

    def formatar_data(self, data):
        if pd.isna(data):
            return None
        if isinstance(data, (int, float)):
            # Converter número para data assumindo que é um timestamp do Excel
            try:
                return (datetime(1899, 12, 30) + timedelta(days=int(data))).strftime('%d%m%Y')
            except ValueError:
                return str(data)  # Retorna o valor original se não puder ser convertido
        elif isinstance(data, str):
            # Tentar converter string para data
            try:
                return pd.to_datetime(data).strftime('%d%m%Y')
            except:
                return data
        elif isinstance(data, pd.Timestamp) or isinstance(data, datetime):
            return data.strftime('%d%m%Y')
        else:
            return str(data)

    def formatar_valor(self, valor):
        if pd.isna(valor) or valor == '':
            return None
        # Converte para string e substitui ponto por vírgula, se necessário
        return str(valor).replace('.', ',')

    def selecionar_planilha(self):
        opcoes = QFileDialog.Options()
        arquivo_entrada, _ = QFileDialog.getOpenFileName(self, "Selecionar Planilha", "", 
                                                         "Arquivos Excel (*.xlsx *.xls);;Todos os Arquivos (*)", options=opcoes)
        if not arquivo_entrada:
            return

        try:
            # Ler a planilha de entrada com o cabeçalho na linha 7 (índice 6)
            self.df = pd.read_excel(arquivo_entrada, header=6)
            
            print("Primeiras linhas do DataFrame original:")
            print(self.df.head())
            
            print("Colunas encontradas:", self.df.columns.tolist())

            # Função auxiliar para encontrar colunas por letra
            def encontrar_coluna_por_letra(letra):
                indice = ord(letra.upper()) - ord('A')
                if 0 <= indice < len(self.df.columns):
                    return self.df.columns[indice]
                return None

            # Mapear as colunas necessárias usando as letras especificadas
            self.coluna_cnpj_cpf = encontrar_coluna_por_letra('R')
            self.coluna_fornecedor = encontrar_coluna_por_letra('Q')
            self.coluna_valor = encontrar_coluna_por_letra('I')
            self.coluna_data_entrada = encontrar_coluna_por_letra('G')
            self.coluna_vencto = encontrar_coluna_por_letra('E')
            self.coluna_n_docto = encontrar_coluna_por_letra('C')

            # Verificar se todas as colunas necessárias foram encontradas
            colunas_necessarias = {
                'CNPJ/CPF': self.coluna_cnpj_cpf,
                'FORNECEDOR': self.coluna_fornecedor,
                'VALOR': self.coluna_valor,
                'DATA DA ENTRADA': self.coluna_data_entrada,
                'VENCTO': self.coluna_vencto,
                'Nº DOCTO': self.coluna_n_docto
            }
            colunas_faltantes = [nome for nome, coluna in colunas_necessarias.items() if coluna is None]
            
            if colunas_faltantes:
                colunas_faltantes_str = ", ".join(colunas_faltantes)
                QMessageBox.warning(self, "Erro", f"As seguintes colunas necessárias não foram encontradas: {colunas_faltantes_str}")
                return

            print("Colunas mapeadas:")
            for nome, coluna in colunas_necessarias.items():
                print(f"{nome}: {coluna}")

            self.botao_preview.setEnabled(True)
            QMessageBox.information(self, "Sucesso", "Planilha carregada com sucesso.")

        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao carregar a planilha: {str(e)}")
            print(f"Erro detalhado: {e}")  # Para debug

    def determinar_grupo_pagamento(self, fornecedor):
        if isinstance(fornecedor, str) and ('BEBIDAS' in fornecedor.upper() or 'VINHO' in fornecedor.upper()):
            return '1106020000'
        return '1106010000'

    def limpar_numero_documento(self, valor):
        if pd.isna(valor) or valor == '':
            return None
        valor_str = str(valor)
        
        # Verifica se é uma data no formato DD/MM/YYYY
        if re.match(r'\d{2}/\d{2}/\d{4}', valor_str):
            return valor_str  # Retorna a data no formato original
        
        # Remove o '00:00:00' se estiver presente
        if ' 00:00:00' in valor_str:
            valor_str = valor_str.split(' ')[0]
        
        # Verifica se é uma data no formato YYYY-MM-DD e converte para DD/MM/YYYY
        if re.match(r'\d{4}-\d{2}-\d{2}', valor_str):
            partes = valor_str.split('-')
            return f"{partes[2]}/{partes[1]}/{partes[0]}"
        
        # Para outros formatos, mantém o formato original
        return valor_str

    def preview_conversao(self):
        if self.df is None:
            QMessageBox.warning(self, "Erro", "Nenhuma planilha carregada.")
            return

        try:
            # Remover a última linha se for a linha de total
            df_processado = self.df.copy()
            ultima_linha = df_processado.iloc[-1]
            
            # Verifica se a última linha é a linha de total
            if (pd.isna(ultima_linha[self.coluna_cnpj_cpf]) or 
                ultima_linha[self.coluna_cnpj_cpf] == '' or 
                'TOTAL' in str(ultima_linha[self.coluna_fornecedor]).upper()):
                df_processado = df_processado.iloc[:-1]
            
            print(f"Número de linhas após remoção do total: {len(df_processado)}")

            # Criar um novo DataFrame mantendo a ordem original
            self.df_saida = pd.DataFrame({
                'Identificação do tipo de integração de título': ['PP'] * len(df_processado),
                'Codigo do Fornecedor': df_processado[self.coluna_cnpj_cpf].apply(self.limpar_cnpj_cpf),
                'Numero do Titulo': df_processado[self.coluna_n_docto].apply(self.limpar_numero_documento),
                'Documento Fiscal': df_processado[self.coluna_n_docto].apply(self.limpar_numero_documento),
                'Empresa Emitente': ['0001'] * len(df_processado),
                'Codigo da Filial': ['01'] * len(df_processado),
                'Empresa Pagadora': ['0001'] * len(df_processado),
                'Tipo de Titulo': ['55'] * len(df_processado),
                'Data de Emissao do Titulo': df_processado[self.coluna_data_entrada].apply(self.formatar_data),
                'Data de Vencimento do Titulo': df_processado[self.coluna_vencto].apply(self.formatar_data),
                'Data de Programacao do Titulo': df_processado[self.coluna_vencto].apply(self.formatar_data),
                'Codigo da Moeda': ['BRL'] * len(df_processado),
                'Tipo de Cobranca': ['CA'] * len(df_processado),
                'Grupo de Pagamento': df_processado[self.coluna_fornecedor].apply(self.determinar_grupo_pagamento),
                'Valor do Grupo de Pagamento': df_processado[self.coluna_valor].apply(self.formatar_valor),
                'Codigo do fluxo de caixa': ['01'] * len(df_processado),
            })

            print("Amostra de dados convertidos:")
            print(self.df_saida.head())
            print(f"Número de linhas no DataFrame final: {len(self.df_saida)}")

            self.exibir_resultado(self.df_saida)
            self.botao_converter.setEnabled(True)

        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro na conversão: {str(e)}")
            print(f"Erro detalhado na conversão: {e}")  # Para debug

    def converter_e_salvar(self):
        if self.df_saida is None:
            QMessageBox.warning(self, "Erro", "Faça o preview antes de salvar.")
            return

        # Selecionar arquivo de saída
        opcoes = QFileDialog.Options()
        arquivo_saida, _ = QFileDialog.getSaveFileName(self, "Salvar Arquivo Convertido", "", 
                                                       "Arquivos Excel (*.xlsx);;Todos os Arquivos (*)", options=opcoes)

        if arquivo_saida:
            # Adicionar a extensão .xlsx se não foi fornecida
            if not arquivo_saida.endswith('.xlsx'):
                arquivo_saida += '.xlsx'

            # Salvar o DataFrame no novo arquivo Excel
            self.df_saida.to_excel(arquivo_saida, index=False)
            QMessageBox.information(self, "Sucesso", f"Arquivo convertido e salvo com sucesso: {arquivo_saida}")

    def exibir_resultado(self, df):
        self.tabela.setRowCount(df.shape[0])
        self.tabela.setColumnCount(df.shape[1])
        self.tabela.setHorizontalHeaderLabels(df.columns)

        for row in range(df.shape[0]):
            for col in range(df.shape[1]):
                valor = df.iloc[row, col]
                item = QTableWidgetItem(str(valor) if pd.notna(valor) else '')
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)  # Torna a célula não editável
                self.tabela.setItem(row, col, item)

        self.tabela.resizeColumnsToContents()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    janela = PlanilhaLeitor()
    janela.show()
    sys.exit(app.exec_())
