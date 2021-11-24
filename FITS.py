import xlwings as xw
import os
import pandas as pd
import sqlalchemy as sa
import datetime as dt
import PySimpleGUI as sg
from Update import Update

class User:

    def __init__(self):
        self.update = Update()
        params = r"Driver={ODBC Driver 17 for SQL Server};Server=tcp:scioexcel2python.database.windows.net,1433;" \
                 r"Database=TestDB;Uid=Daniel;pwd=SCIO-123!#;Encrypt=yes;TrustServerCertificate=no;Connection Timeout=30"
        self.engine = sa.create_engine("mssql+pyodbc:///?odbc_connect=%s" % params, echo=False, fast_executemany=True)
        self.connection = self.engine.connect()

    def get_new_values(self):
        # Liest die aktuellen Daten aus dem Excel File
        wb = xw.Book.caller()
        sheet = wb.sheets[0].used_range.value
        df = pd.DataFrame(sheet)

        # Strukturiert das Pandas DataFrame
        first_row = df.iloc[0]
        first_row = first_row.values.tolist()

        # Durch Liste iterieren und Header setzen
        for i in range(len(first_row)):
            df = df.rename(columns={i: first_row[i]})

        # Erste Zeile entfernen
        df = df.drop(df.index[0])

        return df

    def get_old_values(self):
        table = sa.Table("t_FITS_Abweichungsanalyse_Kommentare", sa.MetaData(), autoload=True, autoload_with=self.engine)
        query = sa.select([table])
        result_proxy = self.connection.execute(query)
        data = result_proxy.fetchall()
        return data

    def set_timestamp_to(self, row):
        table = sa.Table("t_FITS_Abweichungsanalyse_Kommentare", sa.MetaData(), autoload=True,
                         autoload_with=self.engine)
        query = sa.update().where(table.id == row).values(timestamp_to = dt.datetime.now())
        self.engine.execute(query)


    # Fügt kommentare in die Datenbank ein (inklusive Username und Timestamp)
    def insert_comments(self):
        new = self.get_new_values()
        old = self.get_old_values()

        df, row = self.update.get_update(old, new.values.tolist())

        df = pd.DataFrame(df)
        df.columns=("Planungsjahr", "Auftraggeber", "Name1", "Erlösvertragsnummer", "Bezeichnung", "Fakturadatum",
                    "PlanIst Kennzeichen", "PLA", "FC1", "FC2", "FC3", "timestamp", "Kommentar", "user")
        print(df)
        # Werte für Timestamp und Username setzen
        df['timestamp'] = dt.datetime.now()
        df['user'] = os.getlogin()

        df = df[["Planungsjahr","Auftraggeber","Erlösvertragsnummer","Fakturadatum",
                 "Kommentar","PlanIst Kennzeichen","timestamp","user"]]\
                .where(pd.notnull(df["Kommentar"])).dropna()

        df.to_sql('t_FITS_Abweichungsanalyse_Kommentare', con=self.connection, if_exists='append', index=False)
        self.connection.close()


def main():
    user = User()
    user.insert_comments()


if __name__ == "__main__":
    xw.Book("FITS.xlsm").set_mock_caller()
    main()

    # try:
    #     xw.Book("FITS.xlsm").set_mock_caller()
    #     main()
    # except Exception as e:
    #     sg.popup("Es ist ein Fehler aufgetreten \n"
    #              "Versuchen Sie es später erneut.", title="Unerwarteter Fehler", custom_text="OK")
    #     print(e)