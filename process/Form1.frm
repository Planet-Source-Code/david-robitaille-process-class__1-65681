VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Test 
   Caption         =   "Test Process"
   ClientHeight    =   9285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   ScaleHeight     =   9285
   ScaleWidth      =   7485
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtInput 
      Height          =   375
      Left            =   2040
      TabIndex        =   11
      Top             =   7680
      Width           =   2655
   End
   Begin VB.CommandButton btnStdInput 
      Caption         =   "StdInput("""")"
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   7560
      Width           =   1815
   End
   Begin VB.CommandButton btninError 
      Caption         =   "inError"
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   2040
      Width           =   1455
   End
   Begin RichTextLib.RichTextBox RichTextBoxOutput 
      Height          =   2535
      Left            =   120
      TabIndex        =   7
      Top             =   4800
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   4471
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton btnClean 
      Caption         =   "Clean()"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   8400
      Width           =   1815
   End
   Begin VB.CommandButton btnOutput 
      Caption         =   "strGetOutput(True)"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton btnFinished 
      Caption         =   "Finished"
      Height          =   495
      Left            =   2160
      TabIndex        =   3
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton btnRun 
      Caption         =   "Run"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton btnInit 
      Caption         =   "setProcess()"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   2415
   End
   Begin VB.TextBox txtCommandLine 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
   Begin RichTextLib.RichTextBox RichTextBoxError 
      Height          =   1575
      Left            =   120
      TabIndex        =   12
      Top             =   2640
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   2778
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":0080
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblInError 
      Height          =   495
      Left            =   1680
      TabIndex        =   9
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblFinished 
      Height          =   255
      Left            =   4080
      TabIndex        =   6
      Top             =   1560
      Width           =   735
   End
End
Attribute VB_Name = "Test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim testProcess As New Process

Private Sub btnClean_Click()
    testProcess.Clean
End Sub

Private Sub btnFinished_Click()
    lblFinished.Caption = CStr(testProcess.Finished())
End Sub

Private Sub btninError_Click()
    lblInError.Caption = testProcess.inError
End Sub

Private Sub btnInit_Click()
    testProcess.setProcess (txtCommandLine.Text)
End Sub

Private Sub btnOutput_Click()
    RichTextBoxOutput.Text = testProcess.strGetOutput(True)
    RichTextBoxError = testProcess.strError
End Sub

Private Sub btnRun_Click()
    testProcess.Run
End Sub

Private Sub btnStdInput_Click()
    
    testProcess.StdInput txtInput.Text + vbCrLf
End Sub

Private Sub txtCommandLine_Change()
    btnInit.Caption = "setProcess(" & txtCommandLine.Text & ")"
End Sub

Private Sub txtInput_Change()
    btnStdInput.Caption = "stdInput(""" & txtInput.Text & """)"
End Sub
