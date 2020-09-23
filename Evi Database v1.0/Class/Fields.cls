VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Fields"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
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

Enum DataTypeFieldObjectEnum
     [Integers] = 0
     [Long Integer] = 1
     [Doubles] = 2
     [Text] = 3
     '[Encrypt] = 4
     [Dates] = 4
End Enum

Enum TypeIndexFieldEnum
     [None] = 0
     [Primary Key] = 1
End Enum

Event FieldRefresh()

Public Property Get Count() As Long
If RefreshDatabase(Constring, True, True, True, MyTableName) = False Then _
Exit Property
If CountColumn <= 2 Then
    Count = 0
Else
    Count = CountColumn - 2
End If
End Property

Public Property Get Item(Optional Index As Variant) As Field
Attribute Item.VB_UserMemId = 0
Dim clsField As New Field
Set clsField = New Field
m_IndexField = Index
Set Item = clsField
Set clsField = Nothing
End Property

Public Sub AddField(Optional FieldName As String = "", Optional _
DataTypeField As DataTypeFieldObjectEnum = 3, Optional DefinedSize As _
Long = 0, Optional TypeIndexField As TypeIndexFieldEnum = 0)
If RefreshDatabase(Constring, True, True, True, MyTableName) = False Then Exit Sub
If Len(FieldName) = 0 Then
    Raise [Not Allow Column Name is Empty]
    Exit Sub
ElseIf DefinedSize < 0 Then
    Raise [Error Add Column On Table]
    Exit Sub
End If
m_Temp = ""
m_Temp = GetTemp(MyPwd)
For a = 0 To CountTable
    If a <> CountTable Then
        If a = i Then
            m_Temp = m_Temp & m_Column(0) & SplitGlobalColumn & m_Column(1) _
                     & SplitGlobalColumn
            If CountColumn <= 2 Then
                m_Temp = m_Temp & FieldName & SplitTypeColumn & DataTypeField _
                         & SplitTypeColumn & DefinedSize & SplitTypeColumn _
                         & TypeIndexField & SplitTypeColumn & SplitGlobalColumn _
                         & SplitTable
            Else
                For b = 0 To CountColumn
                    If b <> CountColumn Then
                        If b > 1 Then
                            tMpColumn = Split(m_Column(b), SplitTypeColumn)
                            If tMpColumn(0) = FieldName Then
                                Raise [Column is Already Exists]
                                Exit Sub
                            End If
                        End If
                    End If
                Next b
                For b = 0 To CountColumn
                    If b <> CountColumn Then
                        If b > 1 Then
                            tMpColumn = Split(m_Column(b), SplitTypeColumn)
                            If TypeIndexField = [Primary Key] Then
                                If tMpColumn(3) = 1 Then
                                    m_Temp = m_Temp & tMpColumn(0) & SplitTypeColumn _
                                             & tMpColumn(1) & SplitTypeColumn & tMpColumn(2) & _
                                             SplitTypeColumn & 0 & SplitTypeColumn & _
                                             SplitGlobalColumn
                                Else
                                    m_Temp = m_Temp & m_Column(b) & SplitGlobalColumn
                                End If
                            Else
                                m_Temp = m_Temp & m_Column(b) & SplitGlobalColumn
                            End If
                        End If
                    End If
                Next b
                m_Temp = m_Temp & FieldName & SplitTypeColumn & DataTypeField _
                         & SplitTypeColumn & DefinedSize & SplitTypeColumn _
                         & TypeIndexField & SplitTypeColumn & SplitGlobalColumn
                CountRow = Calculate(m_Column(CountColumn), SplitRow)
                If CountRow <= 0 Then
                    m_Temp = m_Temp & SplitTable
                Else
                    m_Row = Split(m_Column(CountColumn), SplitRow)
                    For b = 0 To CountRow
                        If b <> CountRow Then
                            CountDataRow = Calculate(m_Row(b), SplitDataRow)
                            If CountDataRow > 0 Then
                                m_DataRow = Split(m_Row(b), SplitDataRow)
                                For c = 0 To CountDataRow
                                    If c <> CountDataRow Then
                                        m_Temp = m_Temp & m_DataRow(c) _
                                                 & SplitDataRow
                                    End If
                                    If c = CountDataRow Then
                                        m_Temp = m_Temp & "" & SplitDataRow _
                                                 & SplitRow
                                    End If
                                Next c
                            End If
                        End If
                    Next b
                    m_Temp = m_Temp & SplitTable
                End If
            End If
        Else
            m_Temp = m_Temp & m_Table(a) & SplitTable
        End If
    End If
Next a
SaveDatabase Constring, m_Temp, True
End Sub

Public Sub Refresh()
If RefreshDatabase(Constring, True, True, True, MyTableName) = False Then Exit Sub
RaiseEvent FieldRefresh
End Sub

Public Sub RemoveField(Optional FieldName As String = "")
If RefreshDatabase(Constring, True, True, True, MyTableName) = False Then Exit Sub
If CountColumn <= 2 Then
    Raise [Column Not Found]
Else
    If Len(FieldName) = 0 Then
        Raise [Not Allow Column Name is Empty]
        Exit Sub
    End If
    m_Temp = Kosong
    m_Temp = GetTemp(MyPwd)
    For a = 0 To CountColumn
        If a <> CountColumn Then
            If a > 1 Then
                tMpColumn = Split(m_Column(a), SplitTypeColumn)
                If tMpColumn(0) = FieldName Then Exit For
            End If
        End If
        If a = CountColumn Then
            Raise [Column Not Found]
            Exit Sub
        End If
    Next a
    For b = 0 To CountTable
        If b <> CountTable Then
            If b = i Then
                m_Temp = m_Temp & m_Column(0) & SplitGlobalColumn & m_Column(1) _
                         & SplitGlobalColumn
                For c = 0 To CountColumn
                    If c <> CountColumn Then
                        If c > 1 Then
                            If c <> a Then
                                m_Temp = m_Temp & m_Column(c) & SplitGlobalColumn
                            End If
                        End If
                    End If
                Next c
                CountRow = Calculate(m_Column(CountColumn), SplitRow)
                If CountRow <= 0 Then
                    m_Temp = m_Temp & SplitTable
                Else
                    m_Row = Split(m_Column(CountColumn), SplitRow)
                    For c = 0 To CountRow
                        If c <> CountRow Then
                            CountDataRow = Calculate(m_Row(c), SplitDataRow)
                            If CountDataRow > 0 Then
                                m_DataRow = Split(m_Row(c), SplitDataRow)
                                For d = 0 To CountDataRow
                                    If d <> CountDataRow Then
                                        If d <> a - 2 Then
                                            m_Temp = m_Temp & m_DataRow(d) _
                                                     & SplitDataRow
                                        End If
                                    End If
                                    If d = CountDataRow Then
                                        m_Temp = m_Temp & SplitRow
                                    End If
                                Next d
                            End If
                        End If
                    Next c
                    m_Temp = m_Temp & SplitTable
                End If
            Else
                m_Temp = m_Temp & m_Table(b) & SplitTable
            End If
        End If
    Next b
    SaveDatabase Constring, m_Temp, True
End If
End Sub

Private Sub Class_Initialize()
m_IndexField = Kosong
End Sub

Private Sub Class_Terminate()
m_IndexField = Kosong
End Sub
