VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   4695
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Show Report"
         Height          =   495
         Left            =   360
         TabIndex        =   6
         ToolTipText     =   "view the report created with TTX"
         Top             =   360
         Width           =   4000
      End
      Begin VB.CommandButton cmdTut 
         Caption         =   "Tutorial (HTML)"
         Height          =   495
         Left            =   360
         TabIndex        =   5
         ToolTipText     =   "tutorial.htm"
         Top             =   960
         Width           =   4000
      End
      Begin VB.CommandButton cmdPDF 
         Caption         =   "Official Guide (PDF format)"
         Height          =   495
         Left            =   360
         TabIndex        =   4
         ToolTipText     =   "CR_Howto_TTX.pdf"
         Top             =   1800
         Width           =   4000
      End
      Begin VB.Label Label4 
         Caption         =   "Read the tutorial for more information, step by step"
         ForeColor       =   &H00404000&
         Height          =   255
         Left            =   405
         TabIndex        =   8
         Top             =   1500
         Width           =   3855
      End
      Begin VB.Label Label1 
         Caption         =   "Crystal Desicion Official Guide to TTX Available at http://support.crystaldecisions.com/communityCS/"
         ForeColor       =   &H00404000&
         Height          =   495
         Left            =   400
         TabIndex        =   7
         Top             =   2400
         Width           =   3975
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "About"
      Height          =   495
      Left            =   3600
      TabIndex        =   2
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "XTREME.MDB which ships with Crystal Reports is used as the example database"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "This Example shows how to create a report and display it without worrying about the Database Location"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As Connection
Dim rs1 As Recordset 'recordset to create TTX file

'2 declerations below are to display tutorial file.
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOW = 5


Private Sub cmdPDF_Click()
'display the Crystal decisions official guide
Dim hBrowse As Long
'open PDF reader
   On Error GoTo cmdPDF_Click_Error

hBrowse = ShellExecute(Me.hwnd, "open", "http://support.crystaldecisions.com/communityCS/TechnicalPapers/scr8_ttxado.pdf.asp", "", "", SW_SHOW)

   Exit Sub

cmdPDF_Click_Error:

    MsgBox "Error: " & Err.Description

End Sub

Private Sub cmdPrint_Click()
'First we execute the CreateFieldDefFile function & create the customer.ttx file
'format : function recordset,filepath,1 is to overwrite
CreateFieldDefFile rs1, App.path & "\customer.ttx", 1
Form1.Show
End Sub

Private Sub cmdTut_Click()
'display tutorial
Dim hBrowse As Long
'Open the default browser
   On Error GoTo cmdTut_Click_Error

hBrowse = ShellExecute(Me.hwnd, "open", App.path & "\tutorial.htm", "", "", SW_SHOW)

   Exit Sub

cmdTut_Click_Error:

    MsgBox "Error: " & Err.Description

End Sub



Private Sub Command1_Click()
MsgBox "Hi! I'm Murshid, Hope u enjoyed the program" & vbCrLf & "pls vote!!!!!", vbInformation
End Sub

Private Sub Form_Load()
Set cn = New Connection
Set rs1 = New Recordset
Dim path As String
path = App.path & "\db1.mdb"
cn.Open "provider=microsoft.jet.oledb.4.0;data source=" & path & ";jet oledb:database password=NIM"
cn.CursorLocation = adUseClient

rs1.Open "select * from customer", cn, adOpenDynamic, adLockOptimistic

End Sub

