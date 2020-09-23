VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Recordset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Field"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Evi Database v1.0                                                      '
'   Welcome to evi technologi software. This evi database is freeware      '
'   please dont sale                                                       '
'   if you found bug you can contact me.                                   '
'                                                                          '
'   For more information you can contact me on 6281395840904               '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Public Field As New Fields

Event RecordOpen()
Event RecordClose()
Event RecordRefresh()
Event RecordRemove()
Event RecordSaving()

Public Property Get RecordCount() As Long
If RefreshDatabase(Constring, True, True, False, "") = False Then Exit Property
If CountTable <= 0 Then
    RecordCount = 0
Else
    RecordCount = CountTable
End If
End Property

Private Sub Class_Initialize()
Set Field = New Fields
TableOpen = False
MyTableName = Kosong
CounterRow = 0
MVN = False
End Sub

Private Sub Class_Terminate()
Set Field = Nothing
TableOpen = False
MyTableName = Kosong
CounterRow = 0
MVN = False
End Sub

Public Property Get RecordSize() As String
If RefreshDatabase(Constring, True, True, True, MyTableName) = False Then Exit _
Property
RecordSize = FormatKB(Len(m_Table(i)))
End Property

Public Sub AddRecord(Optional RecordName As String = "")
If RefreshDatabase(Constring, True, False, False, "") = False Then Exit Sub
m_Temp = ""
If Len(RecordName) = 0 Then
    Raise [Not Allow Table Name is Empty]
    Exit Sub
End If
If CountTable <= 0 Then
    m_Temp = GetTemp(MyPwd)
    m_Temp = m_Temp & RecordName & SplitGlobalColumn & Now & _
             SplitGlobalColumn & SplitTable
    SaveDatabase Constring, m_Temp, True
Else
    For a = 0 To CountTable
        If a <> CountTable Then
            m_Column = Split(m_Table(a), SplitGlobalColumn)
            If m_Column(0) = RecordName Then
                Raise [Table is Already Exists]
                Exit Sub
            End If
        End If
    Next a
    m_Temp = GetTemp(MyPwd)
    For a = 0 To CountTable
        If a <> CountTable Then
            m_Temp = m_Temp & m_Table(a) & SplitTable
        End If
    Next a
    m_Temp = m_Temp & RecordName & SplitGlobalColumn _
             & Now & SplitGlobalColumn & SplitTable
    SaveDatabase Constring, m_Temp, True
End If
End Sub

Public Sub RemoveRecord(Optional RecordName As String = "")
If RefreshDatabase(Constring, True, False, False, "") = False Then Exit Sub
If Len(RecordName) = 0 Then
    Raise [Not Allow Table Name is Empty]
    Exit Sub
Else
    m_Temp = Kosong
    m_Temp = GetTemp(MyPwd)
    If TableOpen = True Then
        If RecordName = MyTableName Then
            Raise [Cant Remove Table]
            Exit Sub
        End If
    End If
    For a = 0 To CountTable
        If a <> CountTable Then
            tMpColumn = Split(m_Table(a), SplitGlobalColumn)
            If tMpColumn(0) = RecordName Then Exit For
        End If
        If a = CountTable Then
            Raise [Record Is Not Found]
            Exit Sub
        End If
    Next a
    For b = 0 To CountTable
        If b <> CountTable Then
            If b <> a Then
                m_Temp = m_Temp & m_Table(b) & SplitTable
            End If
        End If
    Next b
    SaveDatabase Constring, m_Temp, True
    RaiseEvent RecordRemove
End If
End Sub

Public Sub RenameRecord(Optional OldRecordName As String = "", Optional _
NewRecordName As String = "")
If RefreshDatabase(Constring, True, False, False, "") = False Then _
Exit Sub
If Len(OldRecordName) = 0 Or Len(NewRecordName) = 0 Then
    Raise [Not Allow Table Name is Empty]
Else
    m_Temp = Kosong
    m_Temp = GetTemp(MyPwd)
    If TableOpen = True Then
        If OldRecordName = MyTableName Then
            Raise [Cant Remove Table]
            Exit Sub
        End If
    End If
    For a = 0 To CountTable
        If a <> CountTable Then
            tMpColumn = Split(m_Table(a), SplitGlobalColumn)
            If tMpColumn(0) = OldRecordName Then Exit For
        End If
        If a = CountTable Then
            Raise [Record Is Not Found]
            Exit Sub
        End If
    Next a
    For b = 0 To CountTable
        If b <> CountTable Then
            If b <> a Then
                tMpColumn = Split(m_Table(b), SplitGlobalColumn)
                If tMpColumn(0) = NewRecordName Then
                    Raise [Table is Already Exists]
                    Exit Sub
                End If
            End If
        End If
    Next b
    For b = 0 To CountTable
        If b <> CountTable Then
            If b = a Then
                tMpColumn = Split(m_Table(b), SplitGlobalColumn)
                m_Temp = m_Temp & NewRecordName & SplitGlobalColumn & _
                         tMpColumn(1) & SplitGlobalColumn
                CountColumn = Calculate(m_Table(b), SplitGlobalColumn)
                If CountColumn <= 2 Then
                    m_Temp = m_Temp & SplitTable
                Else
                    For c = 0 To CountColumn
                        If c <> CountColumn Then
                            If c > 1 Then
                                m_Temp = m_Temp & tMpColumn(c) & SplitGlobalColumn
                            End If
                        End If
                    Next c
                    CountRow = Calculate(tMpColumn(CountColumn), SplitRow)
                    If CountRow <= 0 Then
                        m_Temp = m_Temp & SplitTable
                    Else
                        m_Temp = m_Temp & tMpColumn(CountColumn) & SplitTable
                    End If
                End If
            Else
                m_Temp = m_Temp & m_Table(b) & SplitTable
            End If
        End If
    Next b
    SaveDatabase Constring, m_Temp, True
End If
End Sub

Public Sub Opened(Optional RecordName As String = "")
If TableOpen = True Then
    Raise [Table is open]
Else
    If RefreshDatabase(Constring, True, False, True, RecordName) = False _
    Then Exit Sub
    TableOpen = True
    RaiseEvent RecordOpen
End If
End Sub

Public Sub Closed()
If RefreshDatabase(Constring, True, True, True, MyTableName) = False Then _
Exit Sub
TableOpen = False
MyTableName = Kosong
RaiseEvent RecordClose
End Sub

Public Property Get Status() As String
If TableOpen = False Then
    Status = "Record is close."
Else
    Status = "Record is open table " & MyTableName & "."
End If
End Property

Public Property Get ActiveRecordset() As String
If RefreshDatabase(Constring, True, True, True, MyTableName) = False Then Exit _
Property
ActiveRecordset = MyTableName
End Property

Public Sub Refresh()
If RefreshDatabase(Constring, True, True, True, MyTableName) = False Then _
Exit Sub
RaiseEvent RecordRefresh
End Sub

Public Property Get CreateDate() As Variant
If RefreshDatabase(Constring, True, True, True, MyTableName) = False Then _
Exit Property
CreateDate = m_Column(1)
End Property

Public Sub MoveFirst()
If RefreshDatabase(Constring, True, True, True, MyTableName) = False Then Exit Sub
If CountColumn <= 2 Then
    Raise [Column Not Found]
Else
    CountRow = Calculate(m_Column(CountColumn), SplitRow)
    If CountRow <= 0 Then
        Raise [No Row On Column]
    Else
        CounterRow = 0
    End If
End If
End Sub

Public Sub MoveNext()
If RefreshDatabase(Constring, True, True, True, MyTableName) = False Then Exit Sub
If CountColumn <= 2 Then
    Raise [Column Not Found]
Else
    CountRow = Calculate(m_Column(CountColumn), SplitRow)
    If CountRow <= 0 Then
        Raise [No Row On Column]
    Else
        MVN = True
        If CounterRow >= EOF Then Exit Sub
        CounterRow = CounterRow + 1
    End If
End If
End Sub

Public Property Get EOF() As Long
If RefreshDatabase(Constring, True, True, True, MyTableName) = False Then Exit _
Property
If CountColumn <= 2 Then
    Raise [Column Not Found]
Else
    CountRow = Calculate(m_Column(CountColumn), SplitRow)
    If CountRow <= 0 Then
        EOF = 0
    Else
        EOF = CountRow - 1
    End If
End If
End Property

Public Sub Save()
If RefreshDatabase(Constring, True, True, True, MyTableName) = False Then Exit Sub
If CountColumn <= 2 Then
    Raise [Column Not Found]
Else
    m_Temp = Kosong
    m_Temp = GetTemp(MyPwd)
        CountRow = Calculate(m_Column(CountColumn), SplitRow)
        For a = 0 To CountTable
            If a <> CountTable Then
                If a = i Then
                    m_Temp = m_Temp & m_Column(0) & SplitGlobalColumn & _
                             m_Column(1) & SplitGlobalColumn
                    For b = 0 To CountColumn
                        If b <> CountColumn Then
                            If b > 1 Then
                                m_Temp = m_Temp & m_Column(b) & SplitGlobalColumn
                            End If
                        End If
                    Next b
                    CountRow = Calculate(m_Column(CountColumn), SplitRow)
                    If CountRow <= 0 Then
                        For b = 0 To CountColumn
                            If b <> CountColumn Then
                                If b > 1 Then
                                    m_Temp = m_Temp & IndxRecord(b - 2) & SplitDataRow
                                End If
                            End If
                        Next b
                        m_Temp = m_Temp & SplitRow & SplitTable
                    Else
                        m_Row = Split(m_Column(CountColumn), SplitRow)
                        For b = 0 To CountRow
                            If b <> CountRow Then
                                m_Temp = m_Temp & m_Row(b) & SplitRow
                            End If
                        Next b
                        For b = 0 To CountColumn
                            If b <> CountColumn Then
                                If b > 1 Then
                                    m_Temp = m_Temp & IndxRecord(b - 2) & SplitDataRow
                                End If
                            End If
                        Next b
                        m_Temp = m_Temp & SplitRow & SplitTable
                    End If
                Else
                    m_Temp = m_Temp & m_Table(a) & SplitTable
                End If
            End If
        Next a
        SaveDatabase Constring, m_Temp, True
        RaiseEvent RecordSaving
End If
End Sub

Public Sub Update(Optional KeyValue As String = "")
If RefreshDatabase(Constring, True, True, True, MyTableName) = False Then Exit Sub
If Len(KeyValue) = 0 Then
    Raise [Value Cant be Empty]
    Exit Sub
End If
If CountColumn <= 2 Then
    Raise [Column Not Found]
    Exit Sub
End If
    For a = 0 To CountColumn
        If a <> CountColumn Then
            If a > 1 Then
                tMpColumn = Split(m_Column(a), SplitTypeColumn)
                If tMpColumn(3) = 1 Then Exit For
            End If
        End If
        If a = CountColumn Then
            Err.Raise 95, , "No key on the field!"
            Exit Sub
        End If
    Next a
    CountRow = Calculate(m_Column(CountColumn), SplitRow)
    If CountRow <= 0 Then
        Raise [No Row On Column]
        Exit Sub
    Else
        m_Row = Split(m_Column(CountColumn), SplitRow)
        For b = 0 To CountRow
            If b <> CountRow Then
                CountDataRow = Calculate(m_Row(b), SplitDataRow)
                If CountDataRow > 0 Then
                    m_DataRow = Split(m_Row(b), SplitDataRow)
                    For c = 0 To CountDataRow
                        If c <> CountDataRow Then
                            If c = a - 2 Then
                                If m_DataRow(c) = KeyValue Then GoTo FoundKey
                            End If
                        End If
                    Next c
                End If
            End If
            If b = CountRow Then
                Raise [Value Not Found]
                Exit Sub
            End If
        Next b
    End If
Exit Sub
FoundKey:
m_Temp = Kosong
m_Temp = GetTemp(MyPwd)
For d = 0 To CountTable
    If d <> CountTable Then
        If d = i Then
            m_Temp = m_Temp & m_Column(0) & SplitGlobalColumn & m_Column(1) _
                     & SplitGlobalColumn
            For e = 0 To CountColumn
                If e > 1 Then
                    m_Temp = m_Temp & m_Column(e) & SplitGlobalColumn
                End If
            Next e
            For e = 0 To CountRow
                If e <> CountRow Then
                    If e = b Then
                        For f = 0 To CountColumn
                            If f <> CountColumn Then
                                If f > 1 Then
                                    If Len(IndxRecord(f - 2)) > 0 Then
                                        m_Temp = m_Temp & IndxRecord(f - 2) & SplitDataRow
                                    End If
                                End If
                            End If
                            If f = CountColumn Then
                                m_Temp = m_Temp & SplitRow
                            End If
                        Next f
                    Else
                        m_Temp = m_Temp & m_Row(e) & SplitRow
                    End If
                End If
            Next e
            m_Temp = m_Temp & SplitTable
        Else
            m_Temp = m_Temp & m_Table(d) & SplitTable
        End If
    End If
Next d
SaveDatabase Constring, m_Temp, True
End Sub

Public Sub Delete(Optional KeyValue As String = "")
If RefreshDatabase(Constring, True, True, True, MyTableName) = False Then Exit Sub
If CountColumn <= 2 Then
    Raise [Column Not Found]
Else
    If Len(KeyValue) = 0 Then
        Raise [Value Cant be Empty]
        Exit Sub
    End If
    For a = 0 To CountColumn
        If a <> CountColumn Then
            If a > 1 Then
                tMpColumn = Split(m_Column(a), SplitTypeColumn)
                If tMpColumn(3) = 1 Then Exit For
            End If
        End If
        If a = CountColumn Then
            Err.Raise 95, , "No key on the field!"
            Exit Sub
        End If
    Next a
    CountRow = Calculate(m_Column(CountColumn), SplitRow)
    If CountRow <= 0 Then
        Raise [No Row On Column]
        Exit Sub
    Else
        m_Row = Split(m_Column(CountColumn), SplitRow)
        For b = 0 To CountRow
            If b <> CountRow Then
                CountDataRow = Calculate(m_Row(b), SplitDataRow)
                If CountDataRow > 0 Then
                    m_DataRow = Split(m_Row(b), SplitDataRow)
                    For c = 0 To CountDataRow
                        If c <> CountDataRow Then
                            If c = a - 2 Then
                                If m_DataRow(c) = KeyValue Then GoTo FoundKey
                            End If
                        End If
                    Next c
                End If
            End If
            If b = CountRow Then
                Raise [Value Not Found]
                Exit Sub
            End If
        Next b
    End If
End If
Exit Sub
FoundKey:
m_Temp = Kosong
m_Temp = GetTemp(MyPwd)
For d = 0 To CountTable
    If d <> CountTable Then
        If d = i Then
            m_Temp = m_Temp & m_Column(0) & SplitGlobalColumn & m_Column(1) _
                     & SplitGlobalColumn
            For e = 0 To CountColumn
                If e <> CountColumn Then
                    If e > 1 Then
                        m_Temp = m_Temp & m_Column(e) & SplitGlobalColumn
                    End If
                End If
            Next e
            For e = 0 To CountRow
                If e <> CountRow Then
                    If e <> b Then
                        m_Temp = m_Temp & m_Row(e) & SplitRow
                    End If
                End If
            Next e
            m_Temp = m_Temp & SplitTable
        Else
            m_Temp = m_Temp & m_Table(d) & SplitTable
        End If
    End If
Next d
SaveDatabase Constring, m_Temp, True
End Sub
