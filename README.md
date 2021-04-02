# VBA
Collection of VBA classes, modules and functions.

You will find a description of each function below


## SQL API CLASS
**Purpose**
Allow the user to connect to a SQL database, and use the class read/write methods.
**Usage**
Import the class by right clicking in th eproject explorer and say import.
Alternative, create a new class and paste in the text. 
Remember to reference "Microsoft ActiveX Data Objects x.x Library" in the tools menu. 

Use the class as such:
    Dim log As SQL_API
    Set log = New SQL_API
    log.OpenSQL
    'Create a SQL string and get data
    strSQL = "SELECT * FROM test_table"
    log(strSQL)
    'Output data to range
    rngData = worksheets("test").range("A1")
    log.readRsToExcel()
    log.CloseSQL

    What the project does
    Why the project is useful
    How users can get started with the project
    Where users can get help with your project
    Who maintains and contributes to the project

