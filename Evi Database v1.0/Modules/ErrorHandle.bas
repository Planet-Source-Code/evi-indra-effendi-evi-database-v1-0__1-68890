Attribute VB_Name = "ErrorHandle"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Evi Database v1.0                                                      '
'   Welcome to evi technologi software. This evi database is freeware      '
'   please dont sale                                                       '
'   if you found bug you can contact me.                                   '
'                                                                          '
'   For more information you can contact me on 6281395840904               '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Enum TypeErrorObject
     [Database Not Found] = -15000
     [File Not Found] = -15100
     [Error Open Database] = -15110
     [Error Create Database File] = -15140
     [Unknow Format Database File] = -15150
     [Database File is Corrupted] = -15160
     [Record Is Empty] = -15170
     [Record Is Not Found] = -15180
     [Error Add New Table] = -15190
     [Table is Already Exists] = -15200
     [Database Is Closed] = -15210
     [Database Is Opened] = -15220
     [Error Add Column On Table] = -15230
     [Column is Already Exists] = -15240
     [Database is Already Exists] = -15250
     [Invalid Password Database] = -15260
     [File Already Exists] = -15270
     [Copy File Error] = -15280
     [Delete File Error] = -15290
     [Column Not Found] = -15291
     [Not Allow Column Name is Empty] = -15292
     [Not Allow Table Name is Empty] = -15293
     [No Row On Column] = -15294
     [Value Not Found] = -15295
     [Value Must Be Key] = -15296
     [Value Cant be Empty] = -15297
     [New Value Same The Key] = -15298
     [Cant Remove Table] = -15299
     [Cant save append] = -15300
     [No Append on this section] = -15301
     [Table is open] = -15301
     [Table is close] = -15302
End Enum

Private Description As String
Private Num As Long

Public Function Raise(Optional Number As TypeErrorObject)
Select Case Number
        Case -15000: Description = "Database not found!"
        Case -15100: Description = "File not found!"
        Case -15110: Description = "Cant open database file!"
        Case -15140: Description = "Cant create new database file!"
        Case -15150: Description = "Unknow format database file!"
        Case -15160: Description = "Database file is corrupted!"
        Case -15170: Description = "Record is empty!"
        Case -15180: Description = "Record is not found!"
        Case -15190: Description = "Cant add new Record!"
        Case -15200: Description = "Record is already exists!"
        Case -15210: Description = "Database is closed!"
        Case -15220: Description = "Database is opened! Cant Open Database!"
        Case -15230: Description = "Error Add New Field On Record!"
        Case -15240: Description = "Field is Already Exists!"
        Case -15250: Description = "Database is Already Exists!"
        Case -15260: Description = "Invalid password database! Cant open database object!"
        Case -15270: Description = "File is already exists!"
        Case -15280: Description = "Error when copy file!"
        Case -15290: Description = "Error when delete file!"
        Case -15291: Description = "Field not found on Record!"
        Case -15292: Description = "Not allow Field name is empty!"
        Case -15293: Description = "Not allow Record name is empty!"
        Case -15294: Description = "No Value on Field!"
        Case -15295: Description = "Value not found on Field!"
        Case -15296: Description = "Value must be key!"
        Case -15297: Description = "Value cant be empty!"
        Case -15298: Description = "Cant save new value if new value same the key on Field!"
        Case -15299: Description = "Cant remove Record! Record is open!"
        Case -15300: Description = "Cant save append!"
        Case -15301: Description = "Append is empty!"
        Case -15301: Description = "Record is open!"
        Case -15302: Description = "Record is close!"
End Select
Num = Number
Err.Raise Number, , Description
End Function

Public Function ErrNumber() As Long
ErrNumber = Num
End Function

Public Function ErrDescription() As String
ErrDescription = Description
End Function
