VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Field"
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

Public Property Get Name() As String
If RefreshDatabase(Constring, True, True, True, MyTableName) = False Then _
Exit Property
If CountColumn <= 2 Then
    Raise [Column Not Found]
Else
    For a = 0 To CountColumn
        If a <> CountColumn Then
            If a > 1 Then
                tMpColumn = Split(m_Column(a), SplitTypeColumn)
                If tMpColumn(0) = m_IndexField Then Exit For
                If a - 2 = m_IndexField Then Exit For
            End If
        End If
        If a = CountColumn Then
            Raise [Column Not Found]
            Exit Property
        End If
    Next a
    Name = tMpColumn(0)
End If
End Property

Public Property Get DataType() As String
If RefreshDatabase(Constring, True, True, True, MyTableName) = False Then Exit _
Property
If CountColumn <= 2 Then
    Raise [Column Not Found]
Else
    For a = 0 To CountColumn
        If a <> CountColumn Then
            If a > 1 Then
                tMpColumn = Split(m_Column(a), SplitTypeColumn)
                If tMpColumn(0) = m_IndexField Then Exit For
                If a - 2 = m_IndexField Then Exit For
            End If
        End If
        If a = CountColumn Then
            Raise [Column Not Found]
            Exit Property
        End If
    Next a
    If tMpColumn(1) = 0 Then
        DataType = "Integer"
    ElseIf tMpColumn(1) = 1 Then
        DataType = "Long Integer"
    ElseIf tMpColumn(1) = 2 Then
        DataType = "Double"
    ElseIf tMpColumn(1) = 3 Then
        DataType = "Text"
    ElseIf tMpColumn(1) = 4 Then
        DataType = "Date"
    End If
End If
End Property

Public Property Get DefSize() As Long
If RefreshDatabase(Constring, True, True, True, MyTableName) = False Then Exit _
Property
If CountColumn <= 2 Then
    Raise [Column Not Found]
Else
    For a = 0 To CountColumn
        If a <> CountColumn Then
            If a > 1 Then
                tMpColumn = Split(m_Column(a), SplitTypeColumn)
                If tMpColumn(0) = m_IndexField Then Exit For
                If a - 2 = m_IndexField Then Exit For
            End If
        End If
        If a = CountColumn Then
            Raise [Column Not Found]
            Exit Property
        End If
    Next a
    DefSize = tMpColumn(2)
End If
End Property

Public Property Get TypeIndex() As String
If RefreshDatabase(Constring, True, True, True, MyTableName) = False Then Exit _
Property
If CountColumn <= 2 Then
    Raise [Column Not Found]
Else
    For a = 0 To CountColumn
        If a <> CountColumn Then
            If a > 1 Then
                tMpColumn = Split(m_Column(a), SplitTypeColumn)
                If tMpColumn(0) = m_IndexField Then Exit For
                If a - 2 = m_IndexField Then Exit For
            End If
        End If
        If a = CountColumn Then
            Raise [Column Not Found]
            Exit Property
        End If
    Next a
    If tMpColumn(3) = 0 Then
        TypeIndex = "None"
    ElseIf tMpColumn(3) = 1 Then
        TypeIndex = "Primary Key"
    End If
End If
End Property

Public Property Get Value() As String
Attribute Value.VB_UserMemId = 0
If RefreshDatabase(Constring, True, True, True, MyTableName) = False Then Exit _
Property
If CountColumn <= 2 Then
    Raise [Column Not Found]
Else
    For a = 0 To CountColumn
        If a <> CountColumn Then
            If a > 1 Then
                tMpColumn = Split(m_Column(a), SplitTypeColumn)
                If tMpColumn(0) = m_IndexField Then Exit For
                If a - 2 = m_IndexField Then Exit For
            End If
        End If
        If a = CountColumn Then
            Raise [Column Not Found]
            Exit Property
        End If
    Next a
    CountRow = Calculate(m_Column(CountColumn), SplitRow)
    If CountRow <= 0 Then
        Raise [No Row On Column]
        Exit Property
    Else
        m_Row = Split(m_Column(CountColumn), SplitRow)
        If CounterRow > CountRow - 1 Then
            Err.Raise 95, , "Value is end of"
            MVN = False
            CounterRow = 0
            Exit Property
        End If
        If MVN = False Then
            m_DataRow = Split(m_Row(0), SplitDataRow)
            Value = m_DataRow(a - 2)
        Else
            m_DataRow = Split(m_Row(CounterRow), SplitDataRow)
            Value = m_DataRow(a - 2)
        End If
    End If
End If
End Property

Public Property Let Value(ByRef New_Value As String)
If RefreshDatabase(Constring, True, True, True, MyTableName) = False Then
    Exit Property
End If
If CountColumn <= 2 Then
    Raise [Column Not Found]
Else
    For a = 0 To CountColumn
        If a <> CountColumn Then
            If a > 1 Then
                tMpColumn = Split(m_Column(a), SplitTypeColumn)
                If tMpColumn(0) = m_IndexField Then Exit For
                If a - 2 = m_IndexField Then Exit For
            End If
        End If
        If a = CountColumn Then
            Raise [Column Not Found]
            Exit Property
        End If
    Next a
    If tMpColumn(1) = 0 Then
        ForInteger = New_Value
        If tMpColumn(2) > 0 Then
            If Len(New_Value) > tMpColumn(2) Then
                Err.Raise 95, , "Defined size is " & tMpColumn(2)
                Exit Property
            End If
        End If
    ElseIf tMpColumn(1) = 1 Then
        ForLongInteger = New_Value
        If tMpColumn(2) > 0 Then
            If Len(New_Value) > tMpColumn(2) Then
                Err.Raise 95, , "Defined size is " & tMpColumn(2)
                Exit Property
            End If
        End If
    ElseIf tMpColumn(1) = 2 Then
        ForDouble = New_Value
        If tMpColumn(2) > 0 Then
            If Len(New_Value) > tMpColumn(2) Then
                Err.Raise 95, , "Defined size is " & tMpColumn(2)
                Exit Property
            End If
        End If
    ElseIf tMpColumn(1) = 3 Then
        ForText = New_Value
        If tMpColumn(2) > 0 Then
            If Len(New_Value) > tMpColumn(2) Then
                Err.Raise 95, , "Defined size is " & tMpColumn(2)
                Exit Property
            End If
        End If
    ElseIf tMpColumn(1) = 4 Then
        ForDate = New_Value
    End If
    If tMpColumn(3) = 1 Then
        If Len(New_Value) <= 0 Then
            Raise [Value Cant be Empty]
            Exit Property
        End If
        CountRow = Calculate(m_Column(CountColumn), SplitRow)
        If CountRow > 0 Then
            m_Row = Split(m_Column(CountColumn), SplitRow)
            For b = 0 To CountRow
                If b <> CountRow Then
                    CountDataRow = Calculate(m_Row(b), SplitDataRow)
                    If CountDataRow > 0 Then
                        m_DataRow = Split(m_Row(b), SplitDataRow)
                        For c = 0 To CountDataRow
                            If c <> CountDataRow Then
                                If c = a - 2 Then
                                    If m_DataRow(c) = New_Value Then
                                        Raise [New Value Same The Key]
                                        Exit Property
                                    End If
                                End If
                            End If
                        Next c
                    End If
                End If
            Next b
        End If
    End If
    IndxRecord(a - 2) = New_Value
End If
End Property

Private Sub Class_Initialize()
If RefreshDatabase(Constring, True, True, True, MyTableName) = True Then
    If CountColumn > 2 Then
        ReDim Preserve IndxRecord(0 To CountColumn - 2)
    End If
End If
End Sub
