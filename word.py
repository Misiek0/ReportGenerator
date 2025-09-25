import os
from docx import Document


def open_docx(filename):
    curr_dir = os.getcwd()
    filepath = os.path.join(curr_dir,"templates",filename)
    opened_doc = Document(filepath)
    return opened_doc

def find_col_index(target_colname, table):
    header_row = table.rows[0]
    header_cells = [cell.text.strip() for cell in header_row.cells]
    if target_colname not in header_cells:
        raise ValueError(f"Nie znaleziono kolumny o nazwie {target_colname}")
    col_index = header_cells.index(target_colname)
    return col_index

def find_row_index(target_automat_id, table):
    col_index = find_col_index("Nr seryjny\nautomatu biletowego", table)
    for index, row in enumerate(table.rows):
        cell_value = row.cells[col_index].text.strip()
        if target_automat_id == cell_value:
            return index
    return None



