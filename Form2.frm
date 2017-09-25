VERSION 5.00
Begin VB.Form Form2 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3624
   ClientLeft      =   5880
   ClientTop       =   3420
   ClientWidth     =   5268
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3624
   ScaleWidth      =   5268
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   3585
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form2.frx":0000
      Top             =   0
      Width           =   5265
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_Click()
Unload Form2
End Sub
