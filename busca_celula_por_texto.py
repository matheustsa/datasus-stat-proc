import openpyxl

def buscar_celula_por_texto(ws, texto):
    """
    Busca pelo texto especificado numa célula da planilha.

    Args:
        texto (str): Texto a ser buscado.
    
    Returns:
        cell: Primeira célula da planilha onde o texto foi encontrado.
    """
    
    texto_lower = texto.lower()
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str) and texto_lower in cell.value.lower():
                return cell
    return None

# Exemplo de uso
wb = openpyxl.load_workbook('planilhas.xlsx')
texto = 'PERIODO'
for ws in wb.worksheets:
    celula_encontrada = buscar_celula_por_texto(ws, texto) 
    
    if celula_encontrada:
        print(f'Celula encontrada: {celula_encontrada.coordinate}, valor: {celula_encontrada.value}')
    else:
        print(f"Célula com o texto {texto} não foi encontrada.")