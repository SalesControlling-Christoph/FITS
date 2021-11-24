import PySimpleGUI as sg
import sys
import xlwings as xw
import pandas as pd
import datetime as dt
import sqlalchemy as sa
from tqdm import tqdm


class Insert:

    def __init__(self):
        print("Verbindung mit der Datenbank wird hergestellt.")
        try:
            params = r"Driver={ODBC Driver 17 for SQL Server};Server=tcp:scioexcel2python.database.windows.net,1433;" \
                     r"Database=TestDB;Uid=Daniel;pwd=SCIO-123!#;Encrypt=yes;TrustServerCertificate=no;Connection " \
                     r"Timeout=30 "
            engine = sa.create_engine("mssql+pyodbc:///?odbc_connect=%s" % params, echo=False, fast_executemany=True)
            self.connection = engine.connect()
        except Exception as e:
            print(e)
            sg.popup(
                "Die Verbindung zur Datenbank konnte nicht hergestellt werden. "
                "\n Bitte versuchen Sie es später erneut.", title="Keine Verbindung", custom_text="OK")
            print("Verbindung mit der Datenbank ist fehlgeschlagen.")
            sys.exit()

    def chunker(self, seq, size):
        return (seq[pos:pos + size] for pos in range(0, len(seq), size))

    def insert_with_progress(self, df):
        chunksize = int(len(df) / 10)
        with tqdm(total=len(df)) as pbar:
            for i, cdf in enumerate(self.chunker(df, chunksize)):
                replace = "replace" if i == 0 else "append"
                cdf.to_sql(name="t_FITS_Aufträge_Angebote", con=self.connection, if_exists="append", index=False)
                pbar.update(chunksize)
                tqdm._instances.clear()
        print("Ihre Daten wurden in die Datenbank geladen. \n "
              "Schließen Sie das Programm noch nicht, da die Daten noch formatiert werden müssen")

    def select_files(self):

        # Set the window look and interface
        sg.change_look_and_feel('Dark Blue 3')
        layout = [
            [sg.Input(key='_FILES_'), sg.FilesBrowse()],
            [sg.OK(), sg.Cancel()]
        ]
        window = sg.Window("Datei Hochladen", layout)

        # Get the input from the .Browse
        # Verify the user input if it is .xlsx
        while True:
            event, values = window.read()
            path = ""

            if values['_FILES_'] != "":
                path = values['_FILES_'].split(';')

            if event == "OK":
                if len(path) > 0:
                    valid = [index for index, value in enumerate(path) if
                             value.lower().endswith(('.xlsx', '.xlsm', '.xls'))]

                    if len(valid) == len(path):
                        window.close()
                        return path

                    else:
                        sg.popup("Bitte wählen Sie ein passendes Dateiformat (.xlsx)", title="Falsches Dateiformat",
                                 custom_text="OK")

                elif path == "":
                    sg.popup("Bitte wählen Sie die Dokumente aus, \n die Sie in die Datenbhank laden wollen",
                             title="Dokument wählen", custom_text="OK")

            elif event == "Cancel":
                sys.exit()

    def get_sheet(self, file, wb):
        sheet = []

        for s in wb.sheets:
            sheet.append(s.name)

        if len(sheet) == 1:
            return wb.sheets[0], ""
        else:
            layout = [
                [sg.Listbox(sheet, size=(max(map(len, sheet)) + 50, 10), key='LISTBOX')],
                [sg.OK(), sg.Cancel()]
            ]
            window = sg.Window(file, layout, finalize=True)

            while True:
                event, values = window.read(timeout=500)
                if event == sg.WIN_CLOSED:
                    break
                if event == "OK":
                    for i, s in enumerate(wb.sheets):
                        if s.name == values['LISTBOX'][0]:
                            window.close()
                            return wb.sheets[i], s.name

                if event == "Cancel":
                    window.close()
                    self.store_data()

    def store_data(self):
        files = self.select_files()

        for index, file in enumerate(files):
            xw.App(visible=False)
            wb = xw.Book(file)
            file_name = file.split("/")
            file_name = file_name[len(file_name) - 1]
            sheet, sheet_name = self.get_sheet(file_name, wb)

            sg.popup(
                "Folgende Datei wird hochgeladen: \n Mappe: %s - Arbeitsblatt: %s \n das kann einige Minuten dauern." %
                (str(file_name), str(sheet_name)), title="Übersicht", custom_text="OK")
            data = sheet.used_range.value
            data = pd.DataFrame(data)

            # Header in eine Liste schreiben
            first_row = data.iloc[0]
            first_row = first_row.values.tolist()

            # Durch Liste iterieren und Header setzen
            for i in range(len(first_row)):
                data = data.rename(columns={i: first_row[i]})

            # Erste Zeile entfernen
            data = data.drop(data.index[0])

            data['timestamp'] = dt.datetime.now()
            try:
                self.insert_with_progress(data)
            except:
                sg.popup("Die Daten konnten nicht in die Datenbank geladen werden", title="Falsche Datensätze",
                         custom_text="OK")
                sys.exit()

        self.connection.execute("exec dbo.sp_dwh")
        self.connection.close()
        sg.popup("Die Daten aus Ihrem Dokument \n wurden erfolgreich in die Datenbank geladen",
                 title="Daten Hochgeladen", custom_text="OK")
        sys.exit()



if __name__ == '__main__':
    print("Bitte schließen Sie dieses Fenster nicht.")
    i = Insert()
    print("Beim Hochladen Ihrer Dokumente wird Ihnen der Ladefortschritt angezeigt.")
    i.store_data()


