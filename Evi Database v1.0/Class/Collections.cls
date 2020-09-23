VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Collections"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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
Dim m_Data() As String, Cnt As Long
Public Sub Add(Optional Items As Variant)
Cnt = Cnt + 1: ReDim Preserve m_Data(1 To Cnt): m_Data(Cnt) = Items
End Sub
Public Sub Clear()
Cnt = 0
End Sub
Public Property Get Item(Optional Index As Variant) As Variant
Attribute Item.VB_UserMemId = 0
Dim Counter As Long
On Error GoTo Error
Counter = Counter + 1
Do While Counter <= Cnt
    If Index = Counter Then: Item = m_Data(Counter): Exit Property
    If Index = m_Data(Counter) Then: Item = m_Data(Counter): Exit Property
    Counter = Counter + 1
Loop
Error:
Err.Raise 5, , "Item not found"
End Property
Public Property Get Count() As Long
Count = Cnt
End Property
Public Sub Remove(Optional Index As Variant)
Dim Counter As Long
Dim TempCNT As Long
Dim Hit As Long
Dim Tmp() As String
Hit = Hit + 1: TempCNT = Count - 1
If TempCNT = 0 Or TempCNT < 0 Then: Err.Raise 5, , "Item not found": Exit Sub
ReDim Tmp(TempCNT)
For Counter = 1 To Count + 1
    If Counter <> Count + 1 Then
        If m_Data(Counter) = Index Then Exit For
        If Counter = Index Then Exit For
    End If
    If Counter = Count + 1 Then: Err.Raise 5, , "Item not found": Exit Sub
Next Counter
For Counter = 1 To Count
    If Counter = Index Then GoTo MyLooping
    If m_Data(Counter) = Index Then GoTo MyLooping
    Tmp(Hit) = m_Data(Counter): Hit = Hit + 1
MyLooping:
Next Counter
Cnt = Cnt - 1
If Cnt > 0 Then
    ReDim Preserve m_Data(1 To Cnt)
    For Counter = 1 To Cnt
        m_Data(Counter) = Tmp(Counter)
    Next Counter
End If
End Sub
Private Sub Class_Initialize()
Clear
End Sub
Private Sub Class_Terminate()
Clear
End Sub
