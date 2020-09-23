Attribute VB_Name = "mdlMakingDatabase"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Evi Database v1.0                                                      '
'   Welcome to evi technologi software. This evi database is freeware      '
'   please dont sale                                                       '
'   if you found bug you can contact me.                                   '
'                                                                          '
'   For more information you can contact me on 6281395840904               '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public m_Properties() As String
Public m_Table() As String
Public m_Column() As String
Public m_Row() As String
Public m_DataRow() As String

Public CountProperties As Long
Public CountTable As Long
Public CountColumn As Long
Public CountRow As Long
Public CountDataRow As Long

Public m_Temp As String
Public tMpColumn() As String

Public i As Long
Public a As Long
Public b As Long
Public c As Long
Public d As Long
Public e As Long
Public f As Long

Public m_IndexField As Variant
Public IndxRecord() As String

Public ForInteger As Integer
Public ForLongInteger As Long
Public ForDouble As Double
Public ForText As String
Public ForDate As Date
