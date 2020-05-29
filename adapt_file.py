from openpyxl import load_workbook, Workbook
from openpyxl.styles import numbers
from datetime import datetime
from typing import Tuple, NamedTuple
from geopy.geocoders import Nominatim

titles = ["Bateau", "Date", "Départ", "Arrivée", "Lat. Départ", "Long. Départ", "Lat. Arrivée",
          "Long. Arrivée", "Commentaire"]


class Trip(object):
    def __init__(self, boat: str, date: datetime, start: str = '', end: str = '',
                 comment: str = ''):
        """ A structure for trip management"""
        self.boat = boat
        """Trip date as a datetime"""
        self.date = date
        """Start location"""
        self.start = start
        """End location"""
        self.end = end
        """Start position"""
        self.start_pos = None
        """End position"""
        self.end_pos = None
        """Comment"""
        self.comment = comment
        """Locator"""
        self._geolocator = Nominatim(user_agent="osm")

    def add_positions(self):
        location = self._geolocator.geocode(self.start, country_codes=['gb', 'fr', 'es', 'pt'])
        print(location)
        return


def get_trip(val: str) -> Tuple[str, str]:
    if val.count("_") > 0:
        splt = val.split("_")
    else:
        splt = val.split("-")
    if len(splt) == 0:
        raise ValueError("Error in trip value")
    elif len(splt) == 1:
        return [(splt[0],)]
    elif len(splt) == 2:
        return [(splt[0], splt[1])]
    elif len(splt) == 3:
        return [(splt[0], splt[1]), (splt[1], splt[2])]
    else:
        raise ValueError("Trip contains more than 3 places")


def write_trip(ws: Workbook, trip: Trip, row: int):
    ws.active.cell(row=row, column=1).value = trip.boat
    ws.active.cell(row=row, column=2).number_format = \
        numbers.FORMAT_DATE_DDMMYY
    ws.active.cell(row=row, column=2).value = trip.date
    ws.active.cell(row=row, column=3).value = trip.start
    ws.active.cell(row=row, column=4).value = trip.end
    if trip.start_pos is not None:
        ws.active.cell(row=row, column=5).value = trip.start_pos[0]
        ws.active.cell(row=row, column=6).value = trip.start_pos[1]
    if trip.end_pos is not None:
        ws.active.cell(row=row, column=7).value = trip.end_pos[0]
        ws.active.cell(row=row, column=8).value = trip.end_pos[1]
    ws.active.cell(row=row, column=9).value = trip.comment


def adapt_file(file: str, dest: str):
    wb = load_workbook(filename=file)
    ws = Workbook()
    for i_col, title in enumerate(titles):
        ws.active.cell(row=1, column=i_col + 1).value = title
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
                    places = get_trip(cell.value)

                    for place in places:
                        current_trip = Trip(boat=sheet, date=date_with_day, start=place[0])
                        if len(place) > 1:
                            current_trip.end = place[1]
                        if cell.comment is not None:
                            current_trip.comment = cell.comment.text
                        current_trip.add_positions()
                        write_trip(ws, current_trip, row=curr_row)
                        curr_row += 1
    ws.save(dest)


if __name__ == '__main__':  # pragma: no cover
    adapt_file("./journaux de bord.xlsx", "./journaux de bord adapt.xlsx")
