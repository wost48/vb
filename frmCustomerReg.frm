VERSION 5.00
Begin VB.Form frmCustomerReg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database II Project"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   5415
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtLName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   3
      Top             =   2880
      Width           =   3015
   End
   Begin VB.TextBox txtMName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   2280
      Width           =   3015
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   11
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2760
      TabIndex        =   6
      Top             =   4320
      Width           =   2175
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   5
      ToolTipText     =   "Fill out the above fields then press this button to SAVE"
      Top             =   4320
      Width           =   2175
   End
   Begin VB.TextBox txtNos 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   4
      Top             =   3480
      Width           =   1815
   End
   Begin VB.TextBox txtFName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   1680
      Width           =   3015
   End
   Begin VB.TextBox txtId 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   7
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "Last Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   13
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Middle Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Contact Nos."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "First Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   8
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Register Customer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "frmCustomerReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub cntRec()
Dim db_file As String
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim num_records As Integer

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
    
    ' See how many records the table contains.
    Set rs = conn.Execute("SELECT MAX(CustId) FROM tblcust")
  If IsNull(rs.Fields(0)) = True Then
  txtId.Text = 10
  Else
  txtId.Text = rs.Fields(0) + 1
  End If
End Sub
Private Sub cmdAdd_Click()
Dim db_file As String
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim num_records As Integer

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
 
If txtId.Text = "" Or txtFName.Text = "" Or txtNos.Text = "" Or txtMName.Text = "" Or txtLName.Text = "" Then
    MsgBox "All fields are required!", vbExclamation, "Error"
    Exit Sub
 
Else

  ' Populate the table.
   conn.Execute "INSERT INTO tblcust VALUES('" & txtId.Text & "','" & txtFName.Text & "','" & txtMName.Text & "','" & txtLName.Text & "','" & _
           txtNos.Text & "')"
    MsgBox "Succesfully Registered!"
    txtFName.Text = ""
txtMName.Text = ""
txtLName.Text = ""
txtNos.Text = ""
    cntRec
End If
    
End Sub

Private Sub cmdClear_Click()
 cntRec
txtFName.Text = ""
txtMName.Text = ""
txtLName.Text = ""
txtNos.Text = ""
End Sub
Private Sub cmdMenu_Click()
Unload Me
frmBill.Show
End Sub
Private Sub Form_Load()
cntRec
End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Then
ElseIf KeyAscii = 8 Then
Else
KeyAscii = 0
End If
End Sub

Private Sub txtNos_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Then
ElseIf KeyAscii = 8 Then
Else
KeyAscii = 0
End If
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Then
ElseIf KeyAscii = 8 Then
Else
KeyAscii = 0
End If
End Sub


