VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test"
   ClientHeight    =   1200
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   2190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   2190
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
      Begin VB.CommandButton Command1 
         Caption         =   "Good-Bye!"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Please do-not modify any of this code!
'As in resulting the offsets to be changed...explination;PATCH WONT WORK!
'Thank-You
'Just Compile IT!
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    MsgBox "Test Message to get rid off!", vbInformation, "Message #1"
End Sub
