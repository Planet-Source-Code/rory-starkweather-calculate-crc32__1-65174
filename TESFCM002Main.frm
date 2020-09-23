VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "CRC32 Calculator"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog cdlg001 
      Left            =   120
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select File"
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label lblCRCValue 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label lblCRC32 
      Caption         =   "CRC32: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   840
      Width           =   855
   End
   Begin VB.Label lblStatus 
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      Height          =   255
      Left            =   3240
      TabIndex        =   3
      Top             =   2880
      Width           =   1335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pstrFileSpec As String

Private Sub cmdQuit_Click()

   Dim frmForm As Form
   
   For Each frmForm In Forms
      Call Unload(frmForm)
      Set frmForm = Nothing
   Next ' frmForm
   
End Sub

Private Sub cmdRun_Click()

   Dim curCRC As Currency
   '*** This is the string that will hold
   '*** the whole file . . . hopefully.
   Dim strFile As String
   
   If pstrFileSpec <> vbNullString Then
      Me.lblStatus.Caption = "Running . . ."
      DoEvents
      strFile = GetFileQuick(pstrFileSpec)
      
      curCRC = CRC32(strFile)
      
      Me.lblStatus.Caption = "Done"
      Me.lblCRCValue.Caption = Hex(curCRC)
      DoEvents
   Else
      MsgBox "No file selected."
   End If
   
   Me.cmdRun.Enabled = False
   
End Sub

Private Sub cmdSelect_Click()

   Dim strFileSpec As String
   
   Me.lblCRCValue.Caption = vbNullString
   
   '*** Set filters.
   cdlg001.Filter = "All Files (*.*)|*.*"
   cdlg001.FileName = vbNullString
   '*** Display the Open dialog box.
   cdlg001.ShowOpen
   strFileSpec = cdlg001.FileName
   '### MsgBox "strFileSpec: " & strFileSpec
   
   If strFileSpec <> vbNullString Then
      If FileExists(strFileSpec) Then
         Me.cmdRun.Enabled = True
         pstrFileSpec = strFileSpec
      Else
         MsgBox "File not found."
      End If
   Else
      MsgBox "No file selected."
   End If
   
End Sub

Private Sub Form_Load()

   Call CRC32Setup

   Me.lblVersion.Caption = "Version: " & _
                           App.Major & "." & _
                           App.Minor & "." & _
                           App.Revision
                           
End Sub

