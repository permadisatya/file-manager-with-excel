#!/usr/bin/env python

from genericpath import exists
import os
import pathlib
from openpyxl import load_workbook, Workbook
from posixpath import join

fpath = r"C:\Users\permadisatya\Documents\DT Documents\Personal\Log"
txt_filename = "FOLDER.TXT"
xlsx_filename = os.path.join(fpath, "LOG.XLSX")
txt = os.path.join(fpath, txt_filename)
xlsx = load_workbook(xlsx_filename)

sheet = xlsx["tbl_update"]



# function to read all file information in certain folder
def listFile(pathfolder):
    a = []
    b = []
    for entry in pathlib.Path(pathfolder).iterdir():
        if entry.is_file():
            a.append("".join(["ID", str(os.path.getctime(entry)).replace(".", "")]))
            b.append(entry.name)
    return a, b

def main():
    file = open(txt).read().splitlines()
    folder = []
    for p in file:
        folder.append(p.lstrip())
    for fp in folder:
        # check status existing excel sheet
        file_id, file_name = listFile(fp)
        for x, y, z in zip(sheet["A"], sheet["B"], sheet["C"]):
            if x.value != "id_file" and y.value != "current_file_name":
                if z.value != fp:
                    pass
                else:
                    if x.value in file_id and y.value in file_name:
                        sheet.cell(row = x.row, column = 4).value = "Exist"
                        sheet.cell(row = x.row, column = 3).value = fp
                    elif x.value in file_id and y.value not in file_name:
                        sheet.cell(row = x.row, column = 4).value = "Renamed"
                        sheet.cell(row = x.row, column = 5).value = sheet.cell(row = x.row, column = 2).value
                        sheet.cell(row = x.row, column = 2).value = file_name[file_id.index(x.value)]
                        sheet.cell(row = x.row, column = 3).value = fp
                    elif x.value not in file_id:
                        sheet.cell(row = x.row, column = 4).value = "Missing"
                        sheet.cell(row = x.row, column = 3).value = fp

        # check file in excel sheet
        for id in file_id:
            if id not in [c.value for c in sheet["A"]]:
                row = []
                for i in sheet["A"]:
                    if i.value != "" and i.value != None:
                        row.append(i.row)
                sheet.cell(row = max(row)+1, column = 1).value = id
                sheet.cell(row = max(row)+1, column = 2).value = file_name[file_id.index(id)]
                sheet.cell(row = max(row)+1, column = 3).value = fp
                sheet.cell(row = max(row)+1, column = 4).value = "New"

        # # bulk rename
        # for x in sheet["G"]:
        #     if x.value == "ok":
        #         id = sheet.cell(row = x.row, column = 1).value
        #         name = file_name[file_id.index(id)]
        #         src = "\\".join([str(folder), str(name)])
        #         pnm, ext = os.path.splitext(src)
        #         new = "\\".join([str(folder), "".join([str(sheet.cell(row = x.row, column = 6).value), ext.upper()])])
        #         os.rename(src, new)
        #         sheet.cell(row = x.row, column = 6).value = None
        #         sheet.cell(row = x.row, column = 7).value = None
    
    xlsx.save(xlsx_filename)

if __name__ == "__main__":
    main()