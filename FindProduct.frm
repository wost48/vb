VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FindProduct 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "v"
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   15030
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton menu 
      Caption         =   "menu"
      Height          =   495
      Left            =   12480
      TabIndex        =   26
      Top             =   360
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4575
      Left            =   6960
      TabIndex        =   25
      Top             =   1320
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   8070
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text6 
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
      Left            =   2520
      TabIndex        =   24
      Top             =   5880
      Width           =   1215
   End
   Begin VB.TextBox Text5 
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
      Left            =   2520
      TabIndex        =   23
      Top             =   6600
      Width           =   1215
   End
   Begin VB.TextBox Text4 
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
      Left            =   2520
      TabIndex        =   22
      Top             =   7320
      Width           =   1215
   End
   Begin VB.TextBox Text3 
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
      Left            =   2520
      TabIndex        =   18
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox Text2 
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
      Left            =   2520
      TabIndex        =   17
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox Text1 
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
      Left            =   2520
      TabIndex        =   14
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "CLEAR"
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
      Left            =   12240
      TabIndex        =   6
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "UPDATE"
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
      Left            =   10200
      TabIndex        =   5
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "ADD"
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
      Left            =   10200
      TabIndex        =   4
      ToolTipText     =   "Fill out the above fields then press this button to SAVE"
      Top             =   7800
      Width           =   1455
   End
   Begin VB.TextBox txtQnty 
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
      Left            =   2520
      TabIndex        =   3
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox txtPrice 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """Php""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   13321
         SubFormatType   =   2
      EndProperty
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
      Left            =   2520
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtDesc 
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
      Left            =   2520
      TabIndex        =   1
      Top             =   1680
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
      Height          =   405
      Left            =   2520
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "DELETE"
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
      Left            =   12240
      TabIndex        =   7
      Top             =   7800
      Width           =   1455
   End
   Begin VB.Label Label11 
      Caption         =   "harga"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   21
      Top             =   5880
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "satuan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   20
      Top             =   6600
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "subtotal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   19
      Top             =   7320
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "warna"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   16
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "ukuran"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   15
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "tanggal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   13
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "barang"
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
      Left            =   6720
      TabIndex        =   12
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "stok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   11
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "jenis"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   10
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "nama barang"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   9
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "id barang"
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
      Left            =   960
      TabIndex        =   8
      Top             =   960
      Width           =   1335
   End
End
Attribute VB_Name = "FindProduct"
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
    
If txtId.Text = "" Or txtDesc.Text = "" Or txtPrice.Text = "" Or txtQnty.Text = "" Then
    MsgBox "All fields are required!", vbExclamation, "Error"
    Exit Sub
  End If


  ' Populate the table.
   conn.Execute "INSERT INTO tblprod VALUES('" & txtId.Text & "','" & txtDesc.Text & "','" & _
           txtPrice.Text & "','" & txtQnty.Text & "')"
           
    ' See how many records the table contains.
    Set rs = conn.Execute("SELECT COUNT (*) FROM tblprod")
    num_records = rs.Fields(0)
     MsgBox "Item Added: " & num_records & " Item in Total", _
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
txtDesc.Text = ""
txtPrice.Text = ""
txtQnty.Text = ""
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
If txtId.Text = "" Or txtDesc.Text = "" Or txtPrice.Text = "" Or txtQnty.Text = "" Then
    MsgBox "All fields are required!", vbExclamation, "Error"
    Exit Sub
  End If
 
  If txtSearch.Text = "" Then
  MsgBox "Search the Item First", vbCritical, "Error"
  Exit Sub
  End If
If MsgBox("This action will modify the selected record.  Proceed?", vbYesNo, "Update") = vbYes Then
    Set rs = conn.Execute("UPDATE tblprod SET ProdID='" & txtId & "', Description='" & txtDesc & "',Price='" & txtPrice & "',Qnty='" & txtQnty & "'" & _
          "WHERE ProdID=" & txtSearch.Text & "")
           MsgBox "Edited Sucessfully", _
        vbInformation, "Done"
 Else
    Cancel = True
  End If

   
End Sub

Private Sub cmdView_Click()
frmViewProduct.Show
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
    Set rs = conn.Execute("SELECT MAX(ProdId) FROM tblprod")
  If IsNull(rs.Fields(0)) = True Then
  txtId.Text = 100
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
        "CREATE TABLE tblprod(" & _
            "ProdID INTEGER NOT NULL," & _
            "Description   VARCHAR(40)  NOT NULL," & _
            "Price   CURRENCY  NOT NULL," & _
            "Qnty      INTEGER   NOT NULL," & _
            "CONSTRAINT pk PRIMARY KEY(ProdID))"
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


If txtSearch.Text = "" Or txtId.Text = "" Then
    MsgBox "Nothing to Delete.", vbExclamation, "Error"
  
  Else
ans = MsgBox("Do you Want to Delete This Records", vbYesNo + vbQuestion, _
"Delete")
If ans = vbYes Then MsgBox "Succesfully Deleted", vbInformation, "Done"
If ans = vbNo Then Exit Sub

    Set rs = conn.Execute("DELETE *FROM tblprod WHERE ProdID=" & txtSearch.Text & "")
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
       Set rs = conn.Execute("SELECT *FROM tblprod WHERE ProdID=" & txtSearch.Text & "")
    If rs.BOF = False Or rs.EOF = False Then
    If rs.RecordCount <> 0 Then
    MsgBox "Item Found.", vbInformation, "Result"
   txtId.Text = rs.Fields("ProdID").Value
   txtDesc.Text = rs.Fields("Description").Value
   txtPrice.Text = rs.Fields("Price").Value
   txtQnty.Text = rs.Fields("Qnty").Value
   End If
   Else
    MsgBox "Item Not Found!", vbCritical, "Result"
    
   End If
    rs.Close
    Set rs = Nothing

End Sub

Private Sub cmdMenu_Click()
menu.Show
Unload Me
End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Then
ElseIf KeyAscii = 8 Then
Else
KeyAscii = 0
End If
End Sub

Private Sub txtQnty_KeyPress(KeyAscii As Integer)
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
