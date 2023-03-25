# Objectives

This project aims to simplify the management file process related to filenames and attributes using an excel spreadsheet.

Users can construct the filename using excel and bulk rename it.

Each file is associated with an ID file from creation time to maintain relational between the file and attribute.

Users can log which files that missing, renamed, etc., and can have all information in a graph.

**No need to copy and paste all the time if the filename has changed.**

# User Target

Have basic knowledge of python and excel, and want to manage filenames and attributes using excel.

# Usage

![image info](images/f1.png)

For inspect all files and update it into spreadsheet:

```
app.py -i
```

For renaming selected files that user selected in spreadsheet:

```
app.py -r
```
