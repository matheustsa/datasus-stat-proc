from openpyxl import Workbook
import argparse
from gerenciador_biblioteca_openpyxl import *

class PlanilhaProcessor:
    """
    Classe responsável por processar planilhas Excel, adicionando colunas com
    cálculos de porcentagens de dados obtidos e ausentes.

    Attributes:
        nome_arquivo_entrada (str): Nome do arquivo Excel de entrada.
        nome_arquivo_saida (str): Nome do arquivo Excel de saída.
        lista_planilhas_nao_processadas (list): Lista de planilhas que não foram processadas.
        arquivo_xlsx (Workbook): Objeto que representa o arquivo Excel aberto.
    
    Examples:
        ### Exemplo de execução do script sobre arquivo "planilhas.xlsx"
        ```$> python trata_planilhas.py planilhas.xlsx planilhas_tratadas.xlsx
    """

    def __init__(self, nome_arquivo_entrada, nome_arquivo_saida):
        """
        Inicializa o Processador de Planilhas com o nome do arquivo de entrada e saída.

        Args:
            nome_arquivo_entrada (str): Nome do arquivo Excel de entrada.
            nome_arquivo_saida (str): Nome do arquivo Excel de saída.
        """

        self.cabecalho_periodo = 'Período'
        self.cabecalho_dados_obtidos = 'DADOS_OBTIDOS (%)'
        self.cabecalho_dados_ausentes = 'DADOS_AUSENTES (%)'
        self.cabecalho_procent_completa = 'PORCENT_COMPLETA (%)'
        self.cabecalho_procent_ausente = 'PORCENT_AUSENTE (%)'

        self.nome_arquivo_entrada = nome_arquivo_entrada
        self.nome_arquivo_saida = nome_arquivo_saida
        self.lista_planilhas_nao_processadas = []
        self.arquivo_xlsx = self.abrir_arquivo()

    def abrir_arquivo(self):
        """
        Abre o arquivo Excel para processamento.

        Returns:
            Workbook: O arquivo Excel aberto.
        """
        return abrir_arquivo(self.nome_arquivo_entrada)

    def processar_planilha(self, planilha):
        """
        Processa uma única planilha adicionando colunas de porcentagem de dados obtidos
        e dados ausentes, além de calcular as porcentagens gerais.

        Args:
            planilha (Worksheet): Planilha a ser processada.
        """
        print(f'Processando a planilha: {planilha.title}...')

        ultima_linha_planilha, ultima_coluna_planilha = obter_limites_planilha(planilha)

        # Verifica se a planilha é válida
        if ultima_linha_planilha < 3:
            self.lista_planilhas_nao_processadas.append(planilha.title)
            return

        celula_periodo = buscar_celula_por_texto(planilha, self.cabecalho_periodo)
        if not celula_periodo:
            self.lista_planilhas_nao_processadas.append(planilha.title)
            return

        # Definindo o intervalo de dados e as colunas adicionais
        linha_cabecalhos = celula_periodo.row + 1
        inicio_dados_tabela = planilha.cell(row=linha_cabecalhos + 1, column=2)
        fim_dados_tabela = planilha.cell(row=ultima_linha_planilha - 1, column=ultima_coluna_planilha)

        # Verifica se já existe a coluna 'dados faltantes'
        celula_dados_ausentes = buscar_celula_por_texto(planilha, 'dados faltantes')
        if celula_dados_ausentes:
            celula_dados_ausentes.value = self.cabecalho_dados_ausentes
            celula_dados_ausentes.offset(0,-1).value = self.cabecalho_dados_obtidos
        else:
            col_dados_obtidos, col_dados_ausentes = self.adicionar_colunas_dados_obtidos_ausentes(planilha, linha_cabecalhos, ultima_coluna_planilha)
            self.calcular_porcentagens_dados_tabela(planilha, inicio_dados_tabela, fim_dados_tabela, col_dados_obtidos, col_dados_ausentes)
        
        self.calcular_porcentagens_gerais(planilha, inicio_dados_tabela, fim_dados_tabela, ultima_linha_planilha)

    def adicionar_colunas_dados_obtidos_ausentes(self, planilha, linha_cabecalhos, ultima_coluna):
        """
        Adiciona colunas de 'DADOS_OBTIDOS' e 'DADOS_AUSENTES' à planilha.

        Args:
            planilha (Worksheet): A planilha onde as colunas serão adicionadas.
            linha_cabecalhos (int): Linha onde os cabeçalhos estão localizados.
            ultima_coluna (int): A última coluna com dados na planilha.

        Returns:
            tuple: Colunas onde 'DADOS_OBTIDOS' e 'DADOS_AUSENTES' foram inseridos.
        """
        col_dados_obtidos = adicionar_celula(planilha, linha=linha_cabecalhos, coluna=ultima_coluna + 1, valor=self.cabecalho_dados_obtidos)
        col_dados_ausentes = adicionar_celula(planilha, linha=linha_cabecalhos, coluna=ultima_coluna + 2, valor=self.cabecalho_dados_ausentes)
        return col_dados_obtidos, col_dados_ausentes

    def calcular_porcentagens_dados_tabela(self, planilha, inicio_dados_tabela, fim_dados_tabela, col_dados_obtidos, col_dados_ausentes):
        """
        Calcula as porcentagens de dados obtidos e ausentes para cada linha da tabela.

        Args:
            planilha (Worksheet): A planilha onde os dados serão calculados.
            inicio_dados_tabela (Cell): Célula que marca o início dos dados.
            fim_dados_tabela (Cell): Célula que marca o fim dos dados.
            col_dados_obtidos (Cell): Coluna onde a porcentagem de dados obtidos será inserida.
            col_dados_ausentes (Cell): Coluna onde a porcentagem de dados ausentes será inserida.
        """
        for linha in range(inicio_dados_tabela.row, fim_dados_tabela.row + 1):
            intervalo = f'{inicio_dados_tabela.column_letter}{linha}:{fim_dados_tabela.column_letter}{linha}'
            
            porcent_obtidos = adicionar_celula(planilha, linha=linha, coluna=col_dados_obtidos.col_idx, valor=f'=COUNT({intervalo}) / COUNTA({intervalo})', formatacao='0.00%')
            adicionar_celula(planilha, linha=linha, coluna=col_dados_ausentes.col_idx, valor=f'=1-{porcent_obtidos.column_letter}{linha}', formatacao='0.00%')

    def calcular_porcentagens_gerais(self, planilha, inicio_dados_tabela, fim_dados_tabela, ultima_linha_planilha):
        """
        Calcula as porcentagens gerais de dados obtidos e ausentes para a tabela inteira.

        Args:
            planilha (Worksheet): A planilha onde os cálculos serão feitos.
            inicio_dados_tabela (Cell): Célula que marca o início dos dados.
            fim_dados_tabela (Cell): Célula que marca o fim dos dados.
            ultima_linha_planilha (int): A última linha da planilha.
        """
        adicionar_celula(planilha, linha=ultima_linha_planilha + 1, coluna=1, valor='-------------')
        adicionar_celula(planilha, linha=ultima_linha_planilha + 2, coluna=1, valor=self.cabecalho_procent_completa)
        intervalo_dados = f'{inicio_dados_tabela.coordinate}:{fim_dados_tabela.coordinate}'
        valor_porcent_completa = adicionar_celula(planilha, linha=ultima_linha_planilha + 2, coluna=2, valor=f'=COUNT({intervalo_dados}) / COUNTA({intervalo_dados})', formatacao='0.00%')

        adicionar_celula(planilha, linha=ultima_linha_planilha + 3, coluna=1, valor=self.cabecalho_procent_ausente)
        adicionar_celula(planilha, linha=ultima_linha_planilha + 3, coluna=2, valor=f'=1-{valor_porcent_completa.column_letter}{valor_porcent_completa.row}', formatacao='0.00%')

    def salvar_planilhas_nao_processadas(self):
        """
        Salva as planilhas que não foram processadas em um arquivo de texto.
        """
        if self.lista_planilhas_nao_processadas:
            with open('planilhas_nao_processadas.txt', 'w', encoding='utf-8') as file:
                file.write('Planilhas não processadas:\n')
                for planilha in self.lista_planilhas_nao_processadas:
                    file.write(f'{planilha}\n')
            
            print(f'Lista de planilhas não processadas ({self.lista_planilhas_nao_processadas.__len__()}) salva em "planilhas_nao_processadas.txt"')
        else:
            print('Todas as planilhas foram processadas com sucesso!')

    def processar(self):
        """
        Processa todas as planilhas do arquivo Excel.
        """
        print('Iniciando o processamento das planilhas...')
        for planilha in self.arquivo_xlsx.worksheets:
            self.processar_planilha(planilha)

        print('Salvando planilha gerada...')
        self.arquivo_xlsx.save(self.nome_arquivo_saida)
        print('Salvando lista de planilhas não processadas...')
        self.salvar_planilhas_nao_processadas()
        print('Processamento concluído.')

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Processar planilhas Excel adicionando porcentagens de dados.')
    parser.add_argument('nome_arquivo_entrada', help='Nome do arquivo Excel de entrada')
    parser.add_argument('nome_arquivo_saida', help='Nome do arquivo Excel de saída')
    
    args = parser.parse_args()
    
    processor = PlanilhaProcessor(args.nome_arquivo_entrada, args.nome_arquivo_saida)
    processor.processar()
