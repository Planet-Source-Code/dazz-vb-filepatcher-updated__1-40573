VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Patcher 1.0 "
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4875
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
   ScaleHeight     =   4680
   ScaleWidth      =   4875
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRun 
      Caption         =   "&Run File"
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   3960
      Width           =   4575
   End
   Begin VB.CheckBox chkBackup 
      Caption         =   "Make a  backup when possible"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3120
      Value           =   1  'Checked
      Width           =   4215
   End
   Begin VB.TextBox txtInfo 
      Height          =   735
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Text            =   "form1.frx":0000
      Top             =   480
      Width           =   4575
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About..."
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3480
      Width           =   1335
   End
   Begin VB.PictureBox picStatus 
      Align           =   2  'Align Bottom
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   4815
      TabIndex        =   6
      Top             =   4425
      Width           =   4875
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ready to patch..."
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   4815
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton cmdPatch 
      Caption         =   "&Start"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "..."
      Height          =   255
      Left            =   4320
      TabIndex        =   3
      Top             =   2760
      Width           =   375
   End
   Begin VB.TextBox txtFile 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   4095
   End
   Begin VB.Label lblFileDate 
      Caption         =   "%FileDate%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   16
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label lblFileSize 
      Caption         =   "%FileSize%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   15
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label lblFileName 
      Caption         =   "%FileName%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   14
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      Caption         =   "FileDate:"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   13
      Top             =   2160
      Width           =   645
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      Caption         =   "FileSize:"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   12
      Top             =   1800
      Width           =   585
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      Caption         =   "FileName:"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   11
      Top             =   1440
      Width           =   705
   End
   Begin VB.Label Label1 
      Caption         =   "Target File:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label lblPatchName 
      Caption         =   "%PatchName%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'--------------------------------------------------------------------------------
'    Project    : PatcherVB
'    Created By : Dazz(punk_dude_daz@hotmail.com)
'    Date       : 10-11-2002
'    Description: Example of how to create a patcher
'                 and demonstation use!
'
'--------------------------------------------------------------------------------
Private m_CRC As clsCRC 'make the crc class
Private Offset(999) As Long 'all the offsets
Private Data(999) As Byte 'all the data
Private Filename As String 'the filename
Private NumberToPatch As Long 'number of offsets to patch
Private FileSize As Integer 'the file size
Private FileCRC As String '"        " crc
Private FileDate As String '"       " date
Dim CRC As String 'the crc calculated of the open dialog file
Dim flen As Long 'the filelen of the open dialog file
Dim sFilename As String 'the filename
Dim lastfilename As String 'last opened filename

Private Sub cmdAbout_Click()
    MsgBox "Created by Dazz 10-11-2002" + vbNewLine + "Thanks To-" + vbNewLine + "Fredrik Qvarfort for CRC Class" + vbNewLine + "And" + "J.-C. Stritt for the common dialog modules", vbInformation, "About..."
End Sub

Private Sub cmdOpen_Click()
        
        On Error GoTo cmdOpen_Click_Err
   

        Dim sFile As String, sFilter As String
100     sFilter = ""
102     sFilter = AddFilterItem(sFilter, Filename, Filename)
    
        'open file dialog!
104     Call DlgInitMoveSystem(MM_PARENT_CENTER, 30)
106     sFile = ShowFileOpenSave(Me, True, "Add a file..", lastfilename, sFilter, 1, "*.*")
    
108     If sFile <> "" Then

110         txtFile.Text = sFile
112         CRC = Hex$(m_CRC.CalculateFile(sFile))
114         flen = FileLen(sFile)
116         sFilename = sFile

        Else
    
        End If

        
        Exit Sub

cmdOpen_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in PatcherVB.Form1.cmdOpen_Click " & _
               "at line " & Erl
        Resume Next
        
End Sub

Private Sub cmdPatch_Click()

    PatchFile

End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdRun_Click()
    Shell txtFile.Text, vbNormalFocus
End Sub

Private Sub Form_Load()
       
        On Error GoTo Form_Load_Err
        

100     Set m_CRC = New clsCRC 'Create the CRC object
        '** Patcher Settings **
102     Filename = "prjTest.exe"
        ' Offset & Data here
104
        Offset(1) = &H1D57: Data(1) = &H90 'nops the offset @ 1D57
        Offset(2) = &H1D58: Data(2) = &H90 'nops the offset @ 1D58
        Offset(3) = &H1D59: Data(3) = &H90 'nops the offset @ 1D59
        Offset(4) = &H1D5A: Data(4) = &H90 'nops the offset @ 1D5A
        Offset(5) = &H1D5B: Data(5) = &H90 'nops the offset @ 1D5B
        
106     NumberToPatch = 5 'total number of offsets to be patched
        '** File Information **
108     FileCRC = "Goes Here" 'Enter the CRC of your compiled prjTest.exe
110     FileDate = "12/10/2002" 'The Files to be patched creation date!
112     FileSize = "20573" 'The files length in bytes
        '** Setting all the captions **
114     lblFileDate.Caption = FileDate
116     lblFileSize.Caption = FileSize & " Bytes"
118     lblFileName.Caption = Filename
        '** Initialization **
120     lblPatchName.Caption = "Patch For Test.exe"
122     lblStatus.Caption = "Please locate " & Filename & "..."
124     m_CRC.Algorithm = CRC32 'set the algo
    
126     cmdOpen_Click 'show the open dialog
   
128     lblStatus = "Ready to patch..."

    
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in PatcherVB.Form1.Form_Load " & _
               "at line " & Erl
        Resume Next
        
End Sub

Private Sub Form_Unload(Cancel As Integer)

    End

End Sub

Public Sub PatchFile()

    If CRC = FileCRC Then 'Crc is good

        If flen = FileSize Then 'file size is good

            '** Open filename for output

            If chkBackup = 1 Then 'wanna make a backup?
            
                FileCopy sFilename, sFilename & ".bak" 'copy the file to be patched and add the .bak extension...very simple
                lblStatus.Caption = "BackUp Created...~!" 'update the status
                Open Filename For Binary As #1
            
                '** Start patching
            
                For i = 1 To NumberToPatch
            
                    Put #1, Offset(i) + 1, Data(i)  '** Offset(i) + 1, not just Offset(i)
            
                Next i
            
                Close #1
            
                lblStatus.Caption = "Done... " & NumberToPatch & " bytes written sucessfully!"
            
            Else 'no backup wanted
            
                Open Filename For Binary As #1
            
                '** Start patching
            
                For i = 1 To NumberToPatch
            
                    Put #1, Offset(i) + 1, Data(i)  '** Offset(i) + 1, not just Offset(i)
            
                Next i
            
                Close #1
            
                lblStatus.Caption = "Done !" & NumberToPatch & " bytes written"
            
            End If

        Else 'file size is bad

            MsgBox "File is already patched or wrong version.(File Size Difference!)", vbInformation, "Error"
            Exit Sub

        End If

    Else 'crc is bad

        MsgBox "File is already patched or wrong version.(Crc Check)", vbInformation, "Error"
        Exit Sub

    End If

End Sub

