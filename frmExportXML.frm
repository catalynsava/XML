VERSION 5.00
Begin VB.Form frmExportXML 
   Caption         =   "Form1"
   ClientHeight    =   2340
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   3624
   LinkTopic       =   "Form1"
   ScaleHeight     =   2340
   ScaleWidth      =   3624
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1212
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2052
   End
End
Attribute VB_Name = "frmExportXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call exportCap0_12
End Sub
