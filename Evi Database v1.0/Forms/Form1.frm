VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Evi Database v1.0"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5250
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "For more information please contact me on +6281395840904"
      Height          =   495
      Left            =   2160
      TabIndex        =   3
      Top             =   960
      Width           =   3015
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright (c) Evi Indra Effendi"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Evi Database v1.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   2010
      Left            =   0
      Picture         =   "Form1.frx":038A
      Top             =   0
      Width           =   1965
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
With Me
    .Caption = "About Evi Database v1.0"
    .Label1.Caption = "Evi Database v1.0"
    .Label2.Caption = "Copyright (c) Evi Indra Effendi"
    .Label3.Caption = "For more information please contact me on +6281395840904"
    .Command1.Caption = "&Close"
End With
End Sub
