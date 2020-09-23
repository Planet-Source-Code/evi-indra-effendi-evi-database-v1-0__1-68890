Attribute VB_Name = "mdlUtilities"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Evi Database v1.0                                                      '
'   Welcome to evi technologi software. This evi database is freeware      '
'   please dont sale                                                       '
'   if you found bug you can contact me.                                   '
'                                                                          '
'   For more information you can contact me on 6281395840904               '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Declare Function StrFormatByteSize Lib "shlwapi" Alias _
"StrFormatByteSizeA" (ByVal dw As Long, ByVal pszBuf As String, ByRef _
cchBuf As Long) As String

Public Function IsExists(Optional FileName As String) As Boolean
On Error GoTo err_handler
    Call FileLen(FileName)
    IsExists = True
    Exit Function
err_handler:
    IsExists = False
End Function

Public Function Calculate(Target As String, Template As String) As Long
Dim Pos1 As Long
Dim Pos2 As Long
Dim Count As Long
If Len(Target) = 0 Or Len(Template) = 0 Or Len(Template) > Len(Target) Then
    Calculate = -1
    Exit Function
End If
Count = 0
Pos2 = 1
Do
    Pos1 = InStr(Pos2, Target, Template, vbTextCompare)
    If Pos1 > 0 Then
        Count = Count + 1
        Pos2 = Pos1 + 1
    End If
Loop Until Pos1 = 0
Calculate = Count
End Function

Public Function FormatKB(ByVal Amount As Long) As String
Dim Buffer As String
Dim Result As String
Buffer = Space$(255)
Result = StrFormatByteSize(Amount, Buffer, Len(Buffer))
If InStr(Result, vbNullChar) > 1 Then
    FormatKB = Left$(Result, InStr(Result, vbNullChar) - 1)
End If
End Function

Public Function GetTemp(Optional Password As String) As String
GetTemp = MyProductName & SplitProperties & MyVersion & SplitProperties & _
          MyCopyright & SplitProperties & MyLicense & SplitProperties & _
          Password & SplitProperties & "NONE" & SplitProperties
End Function

Public Function SaveDatabase(Optional DatabaseName As String, Optional _
Value As String, Optional Encrypt As Boolean)
Dim EncryptDatabaseWhenSavinG As New clsEncrypt
If Len(Value) = 0 Then
    Raise [Unknow Format Database File]
    Exit Function
End If
Open DatabaseName For Output As #1
     Print #1, Trim$(Value$)
Close #1
If Encrypt = True Then
    Set EncryptDatabaseWhenSavinG = New clsEncrypt
    EncryptDatabaseWhenSavinG.Encrypt DatabaseName, "Evi-Indra-Effendi-Cakep-Banget-Euy-Ya-Kan"
    Set EncryptDatabaseWhenSavinG = Nothing
End If
End Function

Sub Main()
If IsExists(App.Path & "\EviDatabase.dll") = False Then
    Err.Raise 95, , "Ilegal rename file activex from ""EviDatabase.dll"""
    Exit Sub
ElseIf App.FileDescription <> "Create and read evi database" Then
    Err.Raise 95, , "Ilegal file description!"
    Exit Sub
ElseIf App.Title <> "Evi Database v1.0" Then
    Err.Raise 95, , "Ilegal file title!"
    Exit Sub
ElseIf App.Comments <> "Evi Database v1.0" Then
    Err.Raise 95, , "Ilegal file comments!"
    Exit Sub
ElseIf App.Major <> 1 And App.Minor <> 0 And App.Revision <> 0 Then
    Err.Raise 95, , "Invalid version!"
    Exit Sub
ElseIf App.CompanyName <> "Evi Indra Effendi" Then
    Err.Raise 95, , "Ilegal company name!"
    Exit Sub
ElseIf App.LegalCopyright <> MyCopyright Then
    Err.Raise 95, , "Ilegal copyright!"
    Exit Sub
ElseIf App.LegalTrademarks <> "Evi Database v1.0" Then
    Err.Raise 95, , "Ilegal trademarks!"
    Exit Sub
End If
End Sub

Public Function ReadAllDatabase(Optional DatabaseName As String, Optional _
Decrypt As Boolean) As String
Dim DecryptDatabaseWhenReadAll As New clsEncrypt
If Decrypt = True Then
    Set DecryptDatabaseWhenReadAll = New clsEncrypt
    ReadAllDatabase = DecryptDatabaseWhenReadAll.Decrypt(DatabaseName, "Evi-Indra-Effendi-Cakep-Banget-Euy-Ya-Kan")
    Set DecryptDatabaseWhenReadAll = Nothing
Else
    Open DatabaseName For Input As #1
            Line Input #1, ReadAllDatabase
    Close #1
End If
If Len(ReadAllDatabase) = 0 Then
    Raise [Unknow Format Database File]
End If
End Function

Public Function RefreshDatabase(Optional DatabaseName As String, Optional _
CheckConnected As Boolean, Optional CheckRecord As Boolean, Optional _
CheckTable As Boolean, Optional TableName As String) As Boolean
If IsExists(DatabaseName) = False Then
    RefreshDatabase = False
    Raise [Database Not Found]
    Exit Function
Else
    If CheckConnected = True Then
        If Connected = False Then
            RefreshDatabase = False
            Raise [Database Is Closed]
            Exit Function
        End If
    End If
    If CheckRecord = True Then
        If TableOpen = False Then
            RefreshDatabase = False
            Raise [Table is close]
            Exit Function
        End If
    End If
    Reader = ReadAllDatabase(DatabaseName, True)
    If Len(Reader) = 0 Then
        RefreshDatabase = False
        Raise [Unknow Format Database File]
        Exit Function
    Else
        CountProperties = Calculate(Reader, SplitProperties)
        If CountProperties < 5 Then
            RefreshDatabase = False
            Raise [Unknow Format Database File]
            Exit Function
        Else
            m_Properties = Split(Reader, SplitProperties)
            If m_Properties(0) <> MyProductName Then
                RefreshDatabase = False
                Raise [Unknow Format Database File]
                Exit Function
            ElseIf m_Properties(2) <> MyCopyright Then
                RefreshDatabase = False
                Raise [Unknow Format Database File]
                Exit Function
            End If
            CountTable = Calculate(m_Properties(6), SplitTable)
            If CountTable > 0 Then
                m_Table = Split(m_Properties(6), SplitTable)
            End If
            If CheckTable = True Then
                If CountTable <= 0 Then
                    RefreshDatabase = False
                    Raise [Record Is Empty]
                    Exit Function
                End If
                If Len(TableName) = 0 Then
                    RefreshDatabase = False
                    Raise [Not Allow Table Name is Empty]
                    Exit Function
                End If
                For i = 0 To CountTable
                    If i <> CountTable Then
                        m_Column = Split(m_Table(i), SplitGlobalColumn)
                        If m_Column(0) = TableName Then Exit For
                    End If
                    If i = CountTable Then
                        RefreshDatabase = False
                        Raise [Record Is Not Found]
                        Exit Function
                    End If
                Next i
                CountColumn = Calculate(m_Table(i), SplitGlobalColumn)
                MyTableName = TableName
            End If
            RefreshDatabase = True
        End If
    End If
End If
End Function
