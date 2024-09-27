import xlsxwriter

def abrir_planilha(arquivo):
    """
    Abre um arquivo XLSX com o nome especificado.

    Args:
        arquivo (str): Nome do arquivo a ser aberto.

    Returns:
        workbook: Objetos de trabalho da planilha.
        
    Example: planilha = abrir_planilha('planilhas.xlsx')
    """

    # Tenta criar um objeto de trabalho para o arquivo
    try:
        workbook = xlsxwriter.Workbook(arquivo)
        
        # Caso tenha sucesso, retorna o objeto de trabalho
        return workbook
    
    except PermissionError:
        print(f"Erro: Não é possível abrir o arquivo '{arquivo}'. Ele pode estar em uso ou não existir.")
    
    except FileNotFoundError:
        print(f"Erro: O arquivo '{arquivo}' não foi encontrado. Certifique-se que ele está na mesma pasta do script ou especifique um caminho completo.")
