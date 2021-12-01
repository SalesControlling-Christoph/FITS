import sqlalchemy as sa
import datetime
import PySimpleGUI as sg
import sys
import pandas as pd
import xlwings as xw


class Update:

    def get_update(self, old, new):
        old = old
        new = new

        outdated = []
        updated = []
        row = []

        if len(old) == 0:
            updated = new
            for index, value in enumerate(updated):
                value.append(str(datetime.datetime.now()))

        elif len(old) < len(new):
            for index, value in enumerate(new):
                for i, v, in enumerate(old):
                    if value != v:
                        print(value, v)
                        print(i)

        elif len(new) < len(old):
            pass

        else:
            for index, value in enumerate(new):
                for i, v, in enumerate(old):
                    if index == i:
                        if value != v:
                            outdated.append(v)
                            updated.append(value)
                            row.append(i)
                        else:
                            print(str(i) + " ist gleich geblieben")
                    if index > len(old)-1 and index not in row:
                        updated.append(value)
                        outdated.append([])
                        row.append(index)

        for index, value in enumerate(updated):
            for i, v in enumerate(outdated):
                if i == index:
                    v.append(str(datetime.datetime.now()))
                    value.append(str(datetime.datetime.now()))

                    print()
                    print("Outdated: " + str(v))
                    print("Updated: " + str(value))
                    print("Row: " + str(row[index]))

        return updated, row


if __name__ == '__main__':
    u = Update()
    old = [["Hallo"], ["Hallo"]]
    new = [["Hallo"], ["Hallo"], ["Hallo"]]
    updated, row = u.get_update(old, new)
    print(updated, row)
