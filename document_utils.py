import os
import pandas as pd
import numpy as np
from datetime import datetime as dt
import time
import pytz


class Document:

    DEFAULT_FONT_SIZE = 12
    DEFAULT_HEADER_FONT_SIZE = 18
    DEFAULT_BORDER = 1
    DEFAULT_MARGIN_BOTTOM = 5
    DEFAULT_ROW_HEIGHT = 15
    DEFAULT_HEADER_ROW_HEIGHT = 30

    TIMEZONE = pytz.timezone("Europe/Berlin")

    def __init__(self, buffer):
        self.writer = pd.ExcelWriter(buffer, engine='xlsxwriter')
        self.workbook = self.writer.book
        pd.DataFrame().to_excel(self.writer, sheet_name='Tabelle 1')
        self.worksheet = self.writer.sheets["Tabelle 1"]
        self.default_format = self.workbook.add_format({
            'border': self.DEFAULT_BORDER,
            'valign': 'top',
            'text_wrap': True,
            'font_size': self.DEFAULT_FONT_SIZE,
            'bottom': 1,
            'top': 0,
            'left': 0,
            'right': 0})

        self.highlight_format = self.workbook.add_format({
            'bold': True,
            'border': self.DEFAULT_BORDER,
            'valign': 'top',
            'text_wrap': True,
            'font_size': self.DEFAULT_FONT_SIZE,
            'bottom': 1,
            'top': 0,
            'left': 0,
            'right': 0})

        self.header_format = self.workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'bottom',
            'font_size': self.DEFAULT_HEADER_FONT_SIZE,
            'border': self.DEFAULT_BORDER})

        self.worksheet.set_default_row(self.DEFAULT_ROW_HEIGHT)

    def _set_column(self, col, width):
        self.worksheet.set_column(f"{col}:{col}", width, self.default_format)

    def _merge_cells(self, df):
        weekday_date = [list(x) for x in set(tuple(x) for x in df[['Wochentag', 'Datum']].values)]
        for weekday, date in weekday_date:
            # find indices and add one to account for header
            u = df.loc[(df['Wochentag'] == weekday) & (df['Datum'] == date)].index.values + 1
            if len(u) < 2:
                pass  # do not merge cells if there is only one car name
            else:
                # merge cells using the first and last indices
                self.worksheet.merge_range(u[0], 0, u[-1], 0, df.loc[u[0], 'Wochentag'], self.default_format)
                self.worksheet.merge_range(u[0], 1, u[-1], 1, df.loc[u[0], 'Datum'], self.default_format)

    def _highlight_rows(self, df, rows):
        for row_num in rows:
            cell_data = df.iloc[row_num, 3]
            cell_value = str(cell_data)
            if '\n' in cell_value:
                lines = cell_value.split('\n')
                # Write the first line with bold formatting
                self.worksheet.write_rich_string(row_num + 1, 3, self.highlight_format, lines[0], self.default_format, "\n" + "\n".join(lines[1:]))
            else:
                self.worksheet.write(row_num + 1, 3, cell_value, self.highlight_format)

    def _adjust_row_heights(self, df):
        for row_num, (index, row) in enumerate(df.iterrows()):
            max_lines = 1
            for cell in row:
                if pd.notna(cell):
                    # Calculate the required height based on the length of the cell's value
                    num_lines = len(str(cell).split('\n'))
                    max_lines = max(max_lines, num_lines)
            new_height = max(self.DEFAULT_ROW_HEIGHT, max_lines * self.DEFAULT_ROW_HEIGHT)
            # Set the row height, ensuring it is not less than the minimum height
            self.worksheet.set_row(row_num + 1, new_height)  # +1 because the first row is the header

    def write(self, df, col_widths, header, header_row, highlight_rows, landscape=False, with_header=True):
        df.to_excel(self.writer, sheet_name='Tabelle 1', index=False, header=False, startrow=1)
        self.worksheet.write(0, 0, header_row, self.header_format)
        self.worksheet.set_row(0, self.DEFAULT_HEADER_ROW_HEIGHT)
        self._merge_cells(df)
        self._highlight_rows(df, highlight_rows)
        self._adjust_row_heights(df)
        for col, width in col_widths:
            self._set_column(col, width)
        self.worksheet.set_header(header)
        if landscape:
            self.worksheet.set_landscape()
        print_area_col = chr(ord('@') + len(col_widths))
        print_area_row = len(df)
        self.worksheet.print_area(f'A1:{print_area_col}{print_area_row}')
        self.writer.close()
        # outdir = os.path.dirname(self.filename)
        # os.system(f"libreoffice --headless --convert-to pdf:calc_pdf_Export --outdir {outdir} {self.filename}")
        # time.sleep(0.1)


def dump_calendar(df, header, highlight_rows, output_buffer):
    col_widths = [
        ("A", 13),  # Wochentag
        ("B", 15),  # Datum
        ("C", 7),  # Uhrzeit
        ("D", 40),  # Termin
        # ("E", 30),  # Untertitel?
        # ("F", 30),  # Kalender
    ]

    doc = Document(output_buffer)
    doc.write(df, col_widths, "Kalender", header_row=header, highlight_rows=highlight_rows, with_header=False)


def dump_services(df, output_buffer):
    col_widths = [
        ("A", 7),  # Datum
        ("B", 30),  # Termin
        ("C", 30),  # Untertitel?
        ("D", 30),  # Kalender
        ("E", 30),  # Kalender
        ("F", 30),  # Kalender
        ("G", 30),  # Kalender
        ("H", 30),  # Kalender
        ("I", 30),  # Kalender
    ]

    doc = Document(output_buffer)
    doc.write(df, col_widths, "Dienste", with_header=True)



def dump_registrations(data, filename, date, service_type=None, with_seats=True):

    if isinstance(data, list):
        df = pd.DataFrame(data)
    else:
        df = pd.read_sql(data.statement, data.session.bind)

    if with_seats:
        df = df.replace(np.nan, -1, regex=True)
        df["assigned_row"] = df["assigned_row"].astype('int')
        df["assigned_row"] = df["assigned_row"].astype('str')
        df = df.replace("-1", '', regex=True)
        df = df.replace(-1, '', regex=True)
        df["Sitzplatz"] = df[["assigned_row", "assigned_seat"]].agg(''.join, axis=1)
    else:
        df["Sitzplatz"] = ""
    df = df.rename(columns={"last_name": "Nachname",
                            "first_name": "Vorname"})
    df["Adresse"] = "bekannt"
    df["Telefon"] = "bekannt"
    df["Anw."] = ""
    df = df[["Nachname", "Vorname", "Adresse", "Telefon", "Sitzplatz", "Anw."]]

    col_widths = [
        ("A", Document.DEFAULT_WIDTH_LAST_NAME),
        ("B", Document.DEFAULT_WIDTH_FIRST_NAME),
        ("C", Document.DEFAULT_WIDTH_ADDRESS),
        ("D", Document.DEFAULT_WIDTH_PHONE),
        ("E", Document.DEFAULT_WIDTH_SEAT),
        ("F", Document.DEFAULT_WIDTH_ATTENDANCE),
    ]

    doc = Document(filename)
    doc.write(df, col_widths, date, service_type)


def dump_members_by_group(members, filename):
    df = pd.read_sql(members.statement, members.session.bind)

    df = df.replace(np.nan, -1, regex=True)
    df["assigned_row"] = df["assigned_row"].astype('int')
    df["assigned_row"] = df["assigned_row"].astype('str')
    df = df.replace("-1", '', regex=True)
    df = df.replace(-1, '', regex=True)
    df["Zugew. Sitzplatz"] = df[["assigned_row", "assigned_seat"]].agg(''.join, axis=1)
    df = df.rename(columns={"last_name": "Nachname",
                            "first_name": "Vorname",
                            "group": "Gruppe"})

    df = df[["Nachname", "Vorname", "Gruppe", "Zugew. Sitzplatz"]]

    col_widths = [
        ("A", Document.DEFAULT_WIDTH_LAST_NAME),
        ("B", Document.DEFAULT_WIDTH_FIRST_NAME),
        ("C", Document.DEFAULT_WIDTH_GROUP),
        ("D", Document.DEFAULT_WIDTH_SEAT)
        ]

    doc = Document(filename)
    doc.write(df, col_widths)


def create_filename(date=None, type=None, groups=None, ext=".xlsx"):
    date_suffix = ""
    type_suffix = ""
    groups_suffix = ""
    if date is not None:
        date_suffix = f"_{date}"
    if type is not None:
        type_suffix = str(type)[0]
        if type_suffix == "R":
            type_suffix = ""
        elif type_suffix == "Ü":
            type_suffix = "U"
        type_suffix = "_" + type_suffix
    if groups is not None:
        if groups == "1":
            groups_suffix = "_A"
        elif groups == "2":
            groups_suffix = "_B"
        elif groups == "3":
            groups_suffix = "_C"
        elif groups == "4":
            groups_suffix = "_A-C"
        elif groups == "5":
            groups_suffix = "_B-C"
        elif groups == "6":
            groups_suffix = "_ohne_Gruppe"
        else:
            groups_suffix = ""

    return f"Teilnehmer{date_suffix}{type_suffix}{groups_suffix}{ext}"


if __name__ == '__main__':
    registrations = RegistrationJoin.query_for_export_2(0)
    dump_registrations(registrations)
