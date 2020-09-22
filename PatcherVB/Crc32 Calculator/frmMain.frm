VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Crc32 Calculation!"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   3540
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   3540
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmdCrc 
      Caption         =   "Ca&lculate Crc32"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   3255
      Begin VB.TextBox txtCRC 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   855
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "..."
      Height          =   255
      Left            =   3120
      TabIndex        =   1
      ToolTipText     =   "Open File to calculate CRC32!"
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Calculate the CRC32 your compiled test.exe!"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : CRC Calc
'    Project    : prjCRC32Calc
'    Author     : Dazz(punk_dude_daz@hotmail.com)
'    Description: calculates CRC32 of a file
'
'--------------------------------------------------------------------------------
'Ok this is a brief explanation of why file patchers use a CRC32 check!
'Alright here we go....
'a crc check can be used to validate if a file has been patched already or maybe damaged
'maybe also if the file version is different the crc will change aswell!
'Installers use the crc check to validate file's Data integrity to see if they might be damaged or whatever
'Winzip also use's this method aswell to find corrupt files
'So in other words using a crc check in a filepatcher is a time saver!
'Instead of adding code to validate versions
'or calculating if the areas needed to be patched are patched or not
'bleh in summary...MUCH,MUCH EASIER WITH A CRC CHECK
'it changes as bytes are changed!
'
'If you have any questions/comments email me or add me to your msn buddy list!
Private m_CRC As clsCRC
Dim lastfilename As Long

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdOpen_Click()

    Dim sFile As String, sFilter As String
    sFilter = "" 'clear the filter
    sFilter = AddFilterItem(sFilter, "AllFiles", "*.*") 'add a new filter for all file types!
    
    'open file dialog!
    Call DlgInitMoveSystem(MM_PARENT_CENTER, 30)
    sFile = ShowFileOpenSave(Me, True, "Add a file..", lastfilename, sFilter, 1, "*.*")
    
    If sFile <> "" Then

        txtFile.Text = sFile 'put the filename into the text box

    End If

End Sub

Private Sub Form_Load()
    MsgBox "No need to really compile this project not necessary..Just run from the IDE!", vbInformation, "Message from Dazz!"
    Set m_CRC = New clsCRC 'Create the CRC object
    m_CRC.Algorithm = CRC32 'set the crc algorithym

End Sub

Private Sub cmdCrc_Click()

    If txtFile.Text <> "" Then

        txtCRC.Text = Hex$(m_CRC.CalculateFile(txtFile.Text)) 'this gets the hex values of the crc32

    Else 'no file has been opened!

        MsgBox "Open a file first...Then retry ;)", vbCritical, "Bleh!"
        Exit Sub 'get outta here!

    End If

End Sub

