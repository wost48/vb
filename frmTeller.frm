VERSION 5.00
Begin VB.Form frmTeller 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database II Project"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   4905
   StartUpPosition =   2  'CenterScreen
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
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   2040
      Width           =   2655
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "LIST"
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
      Left            =   3000
      TabIndex        =   17
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtPW 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "SEARCH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   11
      Top             =   6960
      Width           =   1335
   End
   Begin VB.TextBox txtSearch 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Text            =   "enter teller id..."
      Top             =   6960
      Width           =   3255
   End
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
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   2640
      Width           =   2655
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "DELETE"
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
      Left            =   2280
      TabIndex        =   9
      Top             =   5640
      Width           =   2535
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
      Left            =   120
      TabIndex        =   7
      Top             =   5640
      Width           =   2055
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "UPDATE"
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
      Left            =   2280
      TabIndex        =   8
      Top             =   4440
      Width           =   2535
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
      Left            =   120
      TabIndex        =   6
      Top             =   4440
      Width           =   2055
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "MENU"
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
      Left            =   3960
      TabIndex        =   14
      Top             =   120
      Width           =   855
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
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   1560
      Width           =   2655
   End
   Begin VB.TextBox txtID 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label6 
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
      Height          =   615
      Left            =   120
      TabIndex        =   18
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label Label5 
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   16
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Last Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   15
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "First Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   13
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   12
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "TELLER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "frmTeller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
On Error GoTo err_CmdAdd_Click
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
    
If txtId.Text = "" Or txtFName.Text = "" Or txtLName.Text = "" Then
    MsgBox "All fields are required!", vbExclamation, "Error"
    Exit Sub
  End If

  ' Populate the table.
   conn.Execute "INSERT INTO tblteller VALUES('" & txtId.Text & "','" & txtFName.Text & "','" & txtMName.Text & "','" & _
           txtLName.Text & "','" & txtPW.Text & "')"
           
    ' See how many records the table contains.
    Set rs = conn.Execute("SELECT COUNT (*) FROM tblteller")
    num_records = rs.Fields(0)
     MsgBox "Teller Added: " & num_records & " Teller in Total", _
        vbInformation, "Done"


exit_err_CmdAdd_Click:
    Exit Sub
 
err_CmdAdd_Click:
    If Err = -2147467259 Then
        MsgBox "ID Taken!"
    Else
        MsgBox "ID Taken!"
     End If
End Sub

Private Sub cmdClear_Click()
txtFName.Text = ""
txtMName.Text = ""
txtLName.Text = ""
txtPW.Text = ""
cntRec
End Sub
Private Sub cmdEdit_Click()
Dim db_file As String
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset

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
If txtId.Text = "" Or txtFName.Text = "" Or txtLName.Text = "" Then
    MsgBox "All fields are required!", vbExclamation, "Error"
    Exit Sub
  End If
   If txtSearch.Text = "" Then
  MsgBox "Search the Item First", vbCritical, "Error"
  Exit Sub
  End If
If MsgBox("This action will modify the selected record.  Proceed?", vbYesNo, "Update") = vbYes Then
    Set rs = conn.Execute("UPDATE tblteller SET TellerId='" & txtId & "', FName='" & txtFName & "',MName='" & txtMName & "',LName='" & txtLName & "',PW='" & txtPW & "'" & _
          "WHERE TellerID=" & txtSearch.Text & "")
           MsgBox "Edited Sucessfully", _
        vbInformation, "Done"
 Else
    Cancel = True
  End If

   
End Sub
Private Sub cmdView_Click()
frmViewTeller.Show
End Sub



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
    Set rs = conn.Execute("SELECT MAX(TellerID) FROM tblteller")
  If IsNull(rs.Fields(0)) = True Then
  txtId.Text = 1
  Else
  txtId.Text = rs.Fields(0) + 1
  End If
End Sub
Private Sub Form_Load()
On Error GoTo err_CmdAdd_Click
Dim db_file As String
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset


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
   conn.Execute _
        "CREATE TABLE tblteller(" & _
            "TellerID INTEGER NOT NULL," & _
            "FName   VARCHAR(40)  NOT NULL," & _
            "MName   VARCHAR(40)  NOT NULL," & _
            "LName   VARCHAR(40)  NOT NULL," & _
            "PW     VARCHAR(40)   NOT NULL," & _
            "CONSTRAINT pk PRIMARY KEY(TellerID))"
 cntRec

exit_err_CmdAdd_Click:
    Exit Sub
 
err_CmdAdd_Click:
    If Err = -2147467259 Then
        MsgBox "Table Created!"
    Else
        MsgBox "Table Exists!"
        End If

End Sub

Private Sub cmdDel_Click()
Dim ans As String
Dim db_file As String
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset

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


If txtSearch.Text = "" Or txtId.Text = "" Then
    MsgBox "Nothing to Delete.", vbExclamation, "Error"
  
  Else
ans = MsgBox("Do you Want to Delete This Records", vbYesNo + vbQuestion, _
"Delete")
If ans = vbYes Then MsgBox "Succesfully Deleted", vbInformation, "Done"
If ans = vbNo Then Exit Sub


    Set rs = conn.Execute("DELETE *FROM tblteller WHERE TellerID=" & txtSearch.Text & "")
    End If
    
Set rs = Nothing
End Sub


Private Sub cmdGO_Click()
Dim db_file As String
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset

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

If txtSearch.Text = "" Then
    MsgBox "Nothing to Search", vbExclamation, "Error"
    Exit Sub
  End If
    csrc = txtSearch.Text
    Set rs = conn.Execute("SELECT *FROM tblteller WHERE TellerID=" & txtSearch.Text & "")
    If rs.BOF = False Or rs.EOF = False Then
    If rs.RecordCount <> 0 Then
    MsgBox "Item Found.", vbInformation, "Result"
   txtId.Text = rs.Fields("TellerID").Value
   txtFName.Text = rs.Fields("FName").Value
   txtMName.Text = rs.Fields("MName").Value
   txtLName.Text = rs.Fields("LName").Value
   txtPW.Text = rs.Fields("PW").Value
   End If
   Else
    MsgBox "Item Not Found!", vbCritical, "Result"
    
   End If
    rs.Close
    Set rs = Nothing

End Sub

Private Sub cmdMenu_Click()
Menu.Show
Unload Me
End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Then
ElseIf KeyAscii = 8 Then
Else
KeyAscii = 0
End If
End Sub


Private Sub txtSearch_GotFocus()
txtSearch.Text = ""
txtSearch.ForeColor = &H0
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Then
ElseIf KeyAscii = 8 Then
Else
KeyAscii = 0
End If
End Sub
