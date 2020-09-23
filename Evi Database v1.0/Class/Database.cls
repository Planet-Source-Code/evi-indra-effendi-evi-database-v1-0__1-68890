VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Database"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Recordset"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Evi Database v1.0                                                      '
'   Welcome to evi technologi software. This evi database is freeware      '
'   please dont sale                                                       '
'   if you found bug you can contact me.                                   '
'                                                                          '
'   For more information you can contact me on 6281395840904               '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Public Recordset As New Recordset
Public Collections As New Collections

Enum ReplaceEnum
    [No] = 0
    [Yes] = 1
End Enum

Event DatabaseConnected()
Event DatabaseDisconnected()

Public Property Get ConnectionString() As String
Attribute ConnectionString.VB_UserMemId = 0
ConnectionString = Constring
End Property

Public Property Let ConnectionString(ByVal New_Connection As String)
Constring = New_Connection
End Property

Public Property Get ProductName() As String
ProductName = MyProductName
End Property

Public Property Get Version() As String
Version = MyVersion
End Property

Public Property Get Copyright() As String
Copyright = MyCopyright
End Property

Public Sub NewDatabase(Optional DatabaseName As String = "", Optional Passwd _
As String = "", Optional Replace As ReplaceEnum = 0)
m_Temp = Kosong
If Replace = No Then
    If IsExists(DatabaseName) = True Then
        Raise [Database is Already Exists]
        Exit Sub
    End If
End If
If Len(Passwd) = 0 Then
    m_Temp = GetTemp("NONE")
Else
    m_Temp = GetTemp(Passwd)
End If
SaveDatabase DatabaseName, m_Temp, True
End Sub

Public Sub Connect(Optional Passwd As String = "")
If Connected = True Then
    Raise [Database Is Opened]
Else
    If RefreshDatabase(Constring, False, False, False, "") = False Then Exit Sub
    If m_Properties(4) = "NONE" Then
        If Len(Passwd) > 0 Then
            Raise [Invalid Password Database]
            Exit Sub
        End If
    Else
        If Passwd <> m_Properties(4) Then
            Raise [Invalid Password Database]
            Exit Sub
        End If
    End If
    Connected = True
    If Len(Passwd) = 0 Then
        MyPwd = "NONE"
    Else
        MyPwd = Passwd
    End If
    RaiseEvent DatabaseConnected
End If
End Sub

Public Sub Disconnect()
If RefreshDatabase(ConnectionString, True, False, False, "") = False Then Exit Sub
Connected = False
MyPwd = Kosong
TableOpen = False
MyTableName = Kosong
RaiseEvent DatabaseDisconnected
End Sub

Private Sub Class_Initialize()
ClearCon
Set Recordset = New Recordset
Set Collections = New Collections
End Sub

Private Sub Class_Terminate()
ClearCon
Set Recordset = Nothing
Set Collections = Nothing
End Sub

Private Sub ClearCon()
Connected = False
Constring = Kosong
MyPwd = Kosong
CounterRow = 0
MyTableName = Kosong
End Sub

Public Property Get Password() As String
If RefreshDatabase(ConnectionString, True, False, False, "") = False Then Exit _
Property
If m_Properties(4) = "NONE" Then Exit Property
Password = m_Properties(4)
End Property

Public Property Let Password(ByVal New_Password As String)
If RefreshDatabase(ConnectionString, True, False, False, "") = False Then Exit _
Property
m_Temp = Kosong
If Len(New_Password) = 0 Then
    m_Temp = GetTemp("NONE")
Else
    m_Temp = GetTemp(New_Password)
End If
If CountTable > 0 Then
    m_Temp = m_Temp & m_Properties(6)
End If
SaveDatabase ConnectionString, m_Temp, True
If Len(New_Password) = 0 Then
    MyPwd = "NONE"
Else
    MyPwd = New_Password
End If
End Property

Public Property Get Status() As Variant
If Connected = False Then
    Status = "Not connect"
Else
    Status = "Connect to database"
End If
End Property

Public Sub Delete(Optional FileName As String = "")
If IsExists(FileName) = False Then
    Raise [File Not Found]
Else
    Kill FileName
End If
End Sub

Public Sub Copy(Optional Source As String = "", Optional Destination As String _
= "")
If IsExists(Source) = False Then
    Raise [File Not Found]
Else
    If IsExists(Destination) = True Then
        Raise [File Already Exists]
    Else
        FileCopy Source, Destination
    End If
End If
End Sub

Public Property Get Size() As String
If RefreshDatabase(Constring, True, False, False, "") = False Then Exit Property
Size = FormatKB(Len(Reader))
End Property

Public Sub About()
Form1.Show vbModal
End Sub
