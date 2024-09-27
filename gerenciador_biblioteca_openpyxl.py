import openpyxl

def abrir_arquivo(arquivo):
    """
    Abre um arquivo XLSX com o nome especificado.

    Args:
        arquivo (str): Nome do arquivo a ser aberto.

    Returns:
        workbook: Objetos de trabalho da planilha.
        
    Example:
        ### Abre arquivo "planilhas.xlsx"
        ```planilha = abrir_arquivo('planilhas.xlsx')
    """

    # Tenta criar um objeto de trabalho para o arquivo
    try:
        workbook = openpyxl.load_workbook(arquivo)
        
        # Caso tenha sucesso, retorna o objeto de trabalho
        return workbook
    
    except PermissionError:
        print(f"Erro: Não é possível abrir o arquivo '{arquivo}'. Ele pode estar em uso ou não existir.")
    
    except FileNotFoundError:
        print(f"Erro: O arquivo '{arquivo}' não foi encontrado. Certifique-se que ele está na mesma pasta do script ou especifique um caminho completo.")

def obter_limites_planilha(planilha):
    """
    Retorna a posição da última linha e coluna da planilha.

    Args:
        ws (worksheet): Planilha a ser consumida.
    
    Returns:
        ultima_linha: Última linha com dados na tabela.
        ultima_coluna: Última coluna com dados na tabela.
        
    Example:
        ### Pegando limites da tabela planilhas.xlsx
        ```wb = openpyxl.load_workbook('planilhas.xlsx')
        for ws in wb.worksheets:
            linhas, colunas = obter_limites_planilha(ws)
            print(f"Tamanho da planilha '{ws.title}': {linhas} x {colunas} (linhas x colunas)")
    """
    ultima_linha = planilha.max_row
    ultima_coluna = planilha.max_column
    return ultima_linha, ultima_coluna

def buscar_celula_por_texto(planilha, texto):
    """
    Busca pelo texto especificado numa célula da planilha.

    Args:
        texto (str): Texto a ser buscado.
    
    Returns:
        cell: Primeira célula da planilha onde o texto foi encontrado.
        
    Example:
        ### Buscando texto "PERIODO"
        ```wb = openpyxl.load_workbook('planilhas.xlsx')
        texto = 'PERIODO'
        for ws in wb.worksheets:
            celula_encontrada = buscar_celula_por_texto(ws, texto) 
            
            if celula_encontrada:
                print(f'Celula encontrada: {celula_encontrada.coordinate}, valor: {celula_encontrada.value}')
            else:
                print(f"Célula com o texto {texto} não foi encontrada.")
    """
    
    texto_lower = texto.lower()
    for row in planilha.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str) and texto_lower in cell.value.lower():
                return cell
    return None

from openpyxl import Workbook

def adicionar_celula(planilha, linha, coluna, valor=None, formatacao=None):
    """
    Adiciona uma célula a uma planilha do Excel.

    Args:
        planilha (Worksheet): A planilha onde a célula será adicionada.
        linha (int): O número da linha onde a célula será inserida.
        coluna (int): O número da coluna onde a célula será inserida.
        valor (any, optional): O valor a ser inserido na célula. Padrão é None.
        formatacao (str, optional): O formato numérico da célula (ex: '0.00%', '#,##0.00'). Padrão é None.

    Returns:
        Cell: A célula criada ou modificada.

    Example:
        ### Criando uma nova planilha e adicionando uma célula formatada
        ```wb = Workbook()
        ws = wb.active
        adicionar_celula(ws, linha=1, coluna=1, valor=0.1234, formatacao='0.00%')
        wb.save('exemplo.xlsx')
    """
    celula = planilha.cell(row=linha, column=coluna, value=valor)
    if formatacao:
        celula.number_format = formatacao
    return celula

