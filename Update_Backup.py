import sqlalchemy as sa
import datetime
import PySimpleGUI as sg
import sys
import pandas as pd
import xlwings as xw


class Update:

    def prepare_data(self, old, new):
        old = old
        new = new


    def get_added(self, old, new):
        old = old
        new = new

        added = []
        row = []
        old_ids = []
        new_ids = []

        for index, value in enumerate(new):
            if value[0] == '':
                added.append(value)
                row.append(index)

        counter = 0

        for n in range(row[0], len(new)):
            for r in row:
                if new[n][0] != '' and new[n][0] + counter == r:
                    counter += 1
            if new[n][0] != '':
                old_ids.append(new[n][0])
                new_ids.append(new[n][0] + counter)

        # Adds the IDs to the new values
        for index, value in enumerate(added):
            value[0] = row[index]

        return added, old_ids, new_ids

    def get_deleted(self, old, new):

        deleted = []
        new_ids = []

        for old_id in old:
            for new_id in new:
                if old_id[0] == new_id[0]:
                    new_ids.append(new_id[0])

        for id in old:
            if id[0] not in new_ids:
                deleted.append(id[0])

        return deleted

    def get_updated(self, old, new):

        old = old
        new = new

        outdated = []
        updated = []
        row = []

        for index, value in enumerate(new):
            for i, v, in enumerate(old):
                if value[0] == v[0]:
                    if value != v:
                        outdated.append(v)
                        updated.append(value)
                        row.append(i)
                    else:
                        print(str(i) + " ist gleich geblieben")


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

    old = [[0, "Das erste"], [1, "Das zweite (dann 3.)"], [2, "Test"], [3, "Test"]]
    new = [[0, "Das erste"], ["", "Text"], [1, "Das dritte (Dann 5.)"], ["", "Text"]]

    print("Updated")
    updated, rows = u.get_updated(old, new)
    print("Updated values: " + str(updated))
    print("Updated rows: " + str(rows))

    print("Deleted")
    deleted = u.get_deleted(old, new)
    print("Deleted values: " + str(deleted))

    print("Added")
    added, old_ids, new_ids = u.get_added(old, new)
    print("Added values: " + str(added))
    print("IDs to update: " + str(old_ids))
    print("Update old IDs to: " + str(new_ids))




