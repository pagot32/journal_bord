from openpyxl import load_workbook, Workbook
from openpyxl.styles import numbers
from datetime import datetime
from typing import Tuple

titles = ["Bateau", "Date", "Départ", "Arrivée", "Lat. Départ", "Long. Départ", "Lat. Arrivée",
          "Long. Arrivée", "Commentaire"]


def get_trip(val: str) -> Tuple[str,str]:
    if val.count("_") > 0:
        splt = val.split("_")
    else:
        splt = val.split("-")
    if len(splt) == 0:
        raise ValueError("Error in trip value")
    elif len(splt) == 1:
        return [(splt[0],)]
    elif len(splt) == 2:
        return [(splt[0],splt[1])]
    elif len(splt) == 3:
        return[(splt[0], splt[1]), (splt[1], splt[2])]
    else:
        raise ValueError("Trip contains more than 3 places")


def adapt_file(file: str, dest: str):
    wb = load_workbook(filename=file)
    ws = Workbook()
    for i_col, title in enumerate(titles):
        ws.active.cell(row=1, column=i_col+1).value = title
    curr_row = 2
    for sheet in wb.get_sheet_names():
        max_row = wb[sheet].max_row
        max_col = wb[sheet].max_column
        for i_col, col in enumerate(wb[sheet].iter_cols(min_row=2, min_col=2,
                                                        max_col=max_col, max_row=max_row)):
            for i_row, cell in enumerate(col):
                if cell.value is not None:
                    date = wb[sheet].cell(row=1, column=i_col + 2).value
                    date_with_day = datetime(year=date.year, month=date.month, day=i_row + 1)
                    trips = get_trip(cell.value)
                    for trip in trips:
                        ws.active.cell(row=curr_row, column=1).value = sheet
                        ws.active.cell(row=curr_row, column=2).number_format = \
                            numbers.FORMAT_DATE_DDMMYY
                        ws.active.cell(row=curr_row, column=2).value = date_with_day
                        ws.active.cell(row=curr_row, column=3).value = trip[0]
                        if len(trip) > 1:
                            ws.active.cell(row=curr_row, column=4).value = trip[1]
                        if cell.comment is not None:
                            ws.active.cell(row=curr_row, column=9).value = cell.comment.text
                        curr_row += 1
    ws.save(dest)


if __name__ == '__main__':  # pragma: no cover
    adapt_file("./journaux de bord.xlsx", "./journaux de bord adapt.xlsx")