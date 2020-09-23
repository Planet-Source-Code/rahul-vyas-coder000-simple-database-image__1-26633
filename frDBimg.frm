VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDbimg 
   Caption         =   "DB Image"
   ClientHeight    =   5325
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   7065
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Top             =   4800
      Width           =   495
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   4800
      Width           =   495
   End
   Begin MSComDlg.CommonDialog CDlg 
      Left            =   3720
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdGet 
      Caption         =   "Get Image"
      Height          =   495
      Left            =   4320
      TabIndex        =   2
      Top             =   4800
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Height          =   4575
      Left            =   120
      ScaleHeight     =   4515
      ScaleWidth      =   6795
      TabIndex        =   1
      Top             =   120
      Width           =   6855
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "Save Image"
      Height          =   495
      Left            =   5760
      TabIndex        =   0
      Top             =   4800
      Width           =   1215
   End
End
Attribute VB_Name = "frmDbimg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New Connection
Dim rs As New Recordset

' In the following module I am creating a byte array
' getting the data from rs into the array
Private Sub cmdGet_Click()
    Dim arr() As Byte, i As Long
    arr = rs![img]
    Open "c:\temp.xxx" For Binary As 1
    Put #1, , arr
    Close #1
    Picture1.Picture = LoadPicture("c:\temp.xxx")
    Kill "c:\temp.xxx"
End Sub

Private Sub cmdNext_Click()
    On Error GoTo err_handler
    rs.MoveNext
    cmdGet_Click
    Exit Sub
err_handler:
    If Err.Number = 3021 Then
        MsgBox "No records left!", vbExclamation
        rs.MoveLast
    End If
End Sub

Private Sub cmdPrev_Click()
    On Error GoTo err_handler
    rs.MovePrevious
    cmdGet_Click
    Exit Sub
err_handler:
    If Err.Number = 3021 Then
        MsgBox "This is the first record!", vbExclamation
        rs.MoveFirst
    End If
End Sub

' In this I am reading the file into a byte array and
' saving the array in the access database

' NOTE - The data-type of the field 'img' is OLEData

Private Sub cmdSet_Click()
    Dim arr() As Byte, i As Long
    CDlg.ShowOpen
    If CDlg.FileName = "" Then Exit Sub
    Open CDlg.FileName For Binary As 1
    ReDim arr(LOF(1))
    Do While Not EOF(1)
        Get #1, , arr
        i = i + 1
    Loop
    Close #1
    rs.AddNew
    rs![img] = arr
    rs.Update
    MsgBox "Saved", vbInformation
End Sub

Private Sub Form_Load()
    con.Open "provider=microsoft.jet.oledb.4.0;data source=" & App.Path & "\pic.mdb"
    rs.Open "Table1", con, adOpenDynamic, adLockOptimistic, adCmdTable
    CDlg.Filter = "Bitmap Images|*.bmp|GIF Images|*.gif|JPG Images|*.jpg|JPEG Images|*.jpeg|ICONS|*.ico|CURSORS|*.cur"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If con.State = adStateOpen Then con.Close
    If rs.State = adStateOpen Then rs.Close
    Set con = Nothing
    Set rs = Nothing
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MsgBox "You may contact me at - rahulreceive@hotmail.com", vbInformation, "Hi!"
End Sub
