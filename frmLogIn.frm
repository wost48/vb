VERSION 5.00
Begin VB.Form frmLogIn 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database II Project"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   3870
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Administrator 
      Caption         =   "LOGIN"
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   3375
      Begin VB.CommandButton cmdClearA 
         Caption         =   "Clear"
         Height          =   375
         Left            =   960
         TabIndex        =   7
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cmdLogIn 
         Caption         =   "Log In"
         Height          =   375
         Left            =   2160
         TabIndex        =   5
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtPWAdmin 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1080
         PasswordChar    =   "l"
         TabIndex        =   3
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox txtUserAdmin 
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
         Left            =   1080
         TabIndex        =   1
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "PASSWORD"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "USERNAME"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000D&
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmLogIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strName

Private Sub cmdClearA_Click()
txtUserAdmin.Text = ""
txtPWAdmin.Text = ""
End Sub

Private Sub cmdClearT_Click()
txtUserTeller.Text = ""
    txtPWTeller.Text = ""
End Sub


Private Sub cmdLogIn_Click()
If txtUserAdmin.Text = "anam" And txtPWAdmin.Text = "anam" Then
txtUserAdmin.Text = ""
txtPWAdmin.Text = ""
Menu.Show
Me.Hide
Else
MsgBox "Username or Password did not match. Try Again", vbExclamation, "Notice"
End If
End Sub


Private Sub cmdLogInTeller_Click()
On Error GoTo err_CmdAdd_Click
Dim db_file As String
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim csrc As String
Dim csrcc As String

  ' Get the database name.
    db_file = App.Path
    If Right$(db_file, 1) <> "\" Then db_file = db_file & _
        "\"
    db_file = db_file & "OurDB.mdb"

    ' Open a connection.
    Set conn = New ADODB.Connection
    conn.ConnectionString = _
        "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=" & db_file & ";" & _
        "Persist Security Info=False"
    conn.Open

If txtUserTeller.Text = "" And txtPWTeller.Text = "" Then
    MsgBox "User ID and Password must be filled.", vbExclamation, "Notice"
    Exit Sub
  End If
    csrc = txtUserTeller.Text
    csrcc = txtPWTeller.Text
    Set rs = conn.Execute("SELECT *FROM tblteller WHERE TellerId=" & csrc & " AND PW='" & csrcc & "'")
    If rs.BOF = False Or rs.EOF = False Then
    If rs.RecordCount <> 0 Then
    txtPWTeller.Text = ""
    
    strName = txtUserTeller.Text
   frmBill.Show
Unload Me
   End If
   Else
    MsgBox "User ID and Password did not match.", vbCritical, "Notice"
    
   End If
    rs.Close
    Set rs = Nothing
exit_err_CmdAdd_Click:
    Exit Sub
 
err_CmdAdd_Click:
    If Err = -2147467259 Then
        MsgBox "Ask the Admin for the User ID and Password", vbCritical, "Error"
    Else
        MsgBox "Ask the Admin for the User ID and Password", vbCritical, "Error"
    End If


End Sub


Private Sub txtPWAdmin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdLogIn_Click
End If
End Sub

Private Sub txtPWTeller_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdLogInTeller_Click
End If
End Sub

Private Sub txtUserTeller_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Then
ElseIf KeyAscii = 8 Then
Else
KeyAscii = 0
End If
End Sub
