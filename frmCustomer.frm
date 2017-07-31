VERSION 5.00
Begin VB.Form frmCustomer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database II Project"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   5490
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
      TabIndex        =   4
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
      TabIndex        =   3
      Top             =   2280
      Width           =   3015
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
      Left            =   4320
      TabIndex        =   16
      Top             =   120
      Width           =   855
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
      Left            =   3360
      TabIndex        =   15
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdGO 
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
      Left            =   3840
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
      Left            =   240
      TabIndex        =   10
      Text            =   "type customer id..."
      Top             =   6960
      Width           =   3495
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
      Left            =   2520
      TabIndex        =   9
      Top             =   5760
      Width           =   2655
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
      Left            =   240
      TabIndex        =   8
      Top             =   5760
      Width           =   2175
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
      Left            =   2520
      TabIndex        =   7
      Top             =   4560
      Width           =   2655
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
      TabIndex        =   6
      ToolTipText     =   "Fill out the above fields then press this button to SAVE"
      Top             =   4560
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
      TabIndex        =   5
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
      TabIndex        =   2
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
      TabIndex        =   1
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
      TabIndex        =   18
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
      TabIndex        =   17
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
      TabIndex        =   14
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
      TabIndex        =   13
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
      TabIndex        =   12
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "CUSTOMER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "frmCustomer"
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
    
If txtId.Text = "" Or txtFName.Text = "" Or txtNos.Text = "" Or txtMName.Text = "" Or txtLName.Text = "" Then
    MsgBox "All fields are required!", vbExclamation, "Error"
    Exit Sub
  End If


  ' Populate the table.
   conn.Execute "INSERT INTO tblcust VALUES('" & txtId.Text & "','" & txtFName.Text & "','" & txtMName.Text & "','" & txtLName.Text & "','" & _
           txtNos.Text & "')"
           
    ' See how many records the table contains.
    Set rs = conn.Execute("SELECT COUNT (*) FROM tblcust")
    num_records = rs.Fields(0)
     MsgBox "Item Added: " & num_records & " Customer in Total", _
        vbInformation, "Done"

exit_err_CmdAdd_Click:
    Exit Sub
 
err_CmdAdd_Click:
    If Err = -2147467259 Then
        MsgBox "ID Taken. Use Another", vbCritical, "Error"
    Else
        MsgBox Err.Description, vbInformation, "Proceed"
    End If
    
End Sub

Private Sub cmdClear_Click()
txtFName.Text = ""
txtMName.Text = ""
txtLName.Text = ""
txtNos.Text = ""
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
If txtId.Text = "" Or txtFName.Text = "" Or txtNos.Text = "" Or txtMName.Text = "" Or txtLName.Text = "" Then
    MsgBox "All fields are required!", vbExclamation, "Error"
    Exit Sub
  End If
  
  If txtSearch.Text = "" Then
  MsgBox "Search the Item First", vbCritical, "Error"
  Exit Sub
  End If
If MsgBox("This action will modify the selected record.  Proceed?", vbYesNo, "Update") = vbYes Then
    Set rs = conn.Execute("UPDATE tblcust SET CustId='" & txtId & "', FName='" & txtFName & "',MName='" & txtMName & "',LName='" & txtLName & "',Contact_Nos='" & txtNos & "'" & _
          "WHERE CustId=" & txtSearch.Text & "")
           MsgBox "Edited Sucessfully", _
        vbInformation, "Done"
 Else
    Cancel = True
  End If

   
End Sub

Private Sub cmdView_Click()
frmCustView.Show
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
    Set rs = conn.Execute("SELECT MAX(CustId) FROM tblcust")
  If IsNull(rs.Fields(0)) = True Then
  txtId.Text = 10
  Else
  txtId.Text = rs.Fields(0) + 1
  End If
End Sub
Private Sub Form_Load()
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

   
    conn.Execute _
        "CREATE TABLE tblcust(" & _
            "CustID INTEGER NOT NULL," & _
            "FName   VARCHAR(40)  NOT NULL," & _
            "MName   VARCHAR(40)  NOT NULL," & _
            "LName   VARCHAR(40)  NOT NULL," & _
            "Contact_Nos   INTEGER  NOT NULL," & _
            "CONSTRAINT pk PRIMARY KEY(CustID))"
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
csrc = txtSearch.Text
    Set rs = conn.Execute("DELETE *FROM tblcust WHERE CustID=" & txtSearch.Text & "")
If ans = vbNo Then Exit Sub

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
    
    Set rs = conn.Execute("SELECT *FROM tblcust WHERE CustID=" & txtSearch.Text & "")
    If rs.BOF = False Or rs.EOF = False Then
    If rs.RecordCount <> 0 Then
    MsgBox "Item Found.", vbInformation, "Result"
   txtId.Text = rs.Fields("CustID").Value
   txtFName.Text = rs.Fields("FName").Value
   txtMName.Text = rs.Fields("MName").Value
   txtLName.Text = rs.Fields("LName").Value
   txtNos.Text = rs.Fields("Contact_Nos").Value
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

Private Sub txtNos_KeyPress(KeyAscii As Integer)
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


