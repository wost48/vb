VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBill 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database II Project"
   ClientHeight    =   8940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8940
   ScaleWidth      =   6480
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReg 
      Caption         =   "New Customer?"
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
      Left            =   3240
      TabIndex        =   34
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
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
      Left            =   5160
      TabIndex        =   8
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Compute"
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
      Left            =   3720
      TabIndex        =   7
      Top             =   7680
      Width           =   1335
   End
   Begin VB.TextBox txtgrand 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """Php""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   13321
         SubFormatType   =   2
      EndProperty
      Enabled         =   0   'False
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
      Left            =   3720
      TabIndex        =   32
      Top             =   6960
      Width           =   2655
   End
   Begin VB.CommandButton cmdDEL 
      Caption         =   "DEL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   5160
      Width           =   1095
   End
   Begin MSComctlLib.ListView Item 
      Height          =   1815
      Left            =   1440
      TabIndex        =   31
      Top             =   5040
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   3201
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Product ID"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Quantity Purchase"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Total"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.ComboBox cmbIDCust 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "frmBill.frx":0000
      Left            =   2040
      List            =   "frmBill.frx":0002
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtOR 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
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
      Left            =   5040
      TabIndex        =   29
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox txtDate 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Left            =   2400
      TabIndex        =   28
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "LOG OUT"
      Height          =   495
      Left            =   5160
      TabIndex        =   27
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtProdPrice 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """Php""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   13321
         SubFormatType   =   2
      EndProperty
      Enabled         =   0   'False
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
      Left            =   2400
      TabIndex        =   26
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton cmdCE 
      Caption         =   "New Transaction"
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
      Left            =   3720
      TabIndex        =   9
      ToolTipText     =   "Clears the form for another transaction"
      Top             =   8280
      Width           =   2655
   End
   Begin VB.TextBox txtQntyPurchased 
      BackColor       =   &H80000018&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
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
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Top             =   4320
      Width           =   1095
   End
   Begin VB.TextBox txtChange 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """Php""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   13321
         SubFormatType   =   2
      EndProperty
      Enabled         =   0   'False
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
      Left            =   1800
      TabIndex        =   21
      Top             =   8280
      Width           =   1695
   End
   Begin VB.TextBox txtTotal 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """Php""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   13321
         SubFormatType   =   2
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   4920
      TabIndex        =   19
      Top             =   4320
      Width           =   1455
   End
   Begin VB.TextBox txtAmnt 
      BackColor       =   &H80000018&
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
      Height          =   495
      Left            =   1800
      TabIndex        =   6
      Top             =   7680
      Width           =   1695
   End
   Begin VB.TextBox txtProdQnty 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
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
      Left            =   4800
      TabIndex        =   17
      Top             =   3720
      Width           =   1575
   End
   Begin VB.TextBox txtProdDesc 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Left            =   3720
      TabIndex        =   15
      Top             =   3120
      Width           =   2655
   End
   Begin VB.ComboBox cmbProdID 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "frmBill.frx":0004
      Left            =   2400
      List            =   "frmBill.frx":0006
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox txtCustName 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Left            =   3360
      TabIndex        =   13
      Top             =   2520
      Width           =   3015
   End
   Begin VB.TextBox txtTellerID 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Left            =   2400
      TabIndex        =   11
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label13 
      Caption         =   "Grand Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   2400
      TabIndex        =   33
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Label Label12 
      Caption         =   "OR No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   30
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label11 
      Caption         =   "Price"
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
      Left            =   1680
      TabIndex        =   25
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label Label10 
      Caption         =   "Qnty Purchased"
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
      Left            =   840
      TabIndex        =   24
      Top             =   4440
      Width           =   2055
   End
   Begin VB.Label Label9 
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   23
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "Change"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   840
      TabIndex        =   22
      Top             =   8400
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Amnt Received"
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
      Left            =   120
      TabIndex        =   20
      Top             =   7800
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Total Cost"
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
      Left            =   3720
      TabIndex        =   18
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Qnty."
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
      Left            =   4080
      TabIndex        =   16
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Product ID"
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
      Left            =   1200
      TabIndex        =   14
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Customer ID"
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
      Left            =   600
      TabIndex        =   12
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Teller ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   10
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Water-Refilling Station"
      BeginProperty Font 
         Name            =   "Mufferaw Free"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   3855
   End
End
Attribute VB_Name = "frmBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub updateQnty()
Dim nq As Integer
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

    nq = txtProdQnty - txtQntyPurchased
    Set rs = conn.Execute("UPDATE tblprod SET Qnty=" & nq & " WHERE ProdID=" & cmbProdID.Text & " ")
End Sub

Private Sub cmbIDCust_Click()
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
     Set rs = conn.Execute("SELECT *FROM tblcust WHERE CustID=" & cmbIDCust.Text & " ")
   txtCustName.Text = rs.Fields("FName").Value + " " + rs.Fields("MName").Value + " " + rs.Fields("LName").Value
   rs.Close
    Set rs = Nothing
cmbProdID.Enabled = True
End Sub
Private Sub cmbProdID_Click()
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
     Set rs = conn.Execute("SELECT *FROM tblprod WHERE ProdID=" & cmbProdID.Text & " ")
   txtProdDesc.Text = rs.Fields("Description").Value
   txtProdQnty.Text = rs.Fields("Qnty").Value
   txtProdPrice.Text = FormatCurrency(rs.Fields("Price").Value)
   rs.Close
    Set rs = Nothing
    txtQntyPurchased.Enabled = True
End Sub
Private Sub cmdAdd_Click()
Dim db_file As String
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim qnty As Currency

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
       Set rs = conn.Execute("SELECT Qnty FROM tblprod WHERE ProdID=" & cmbProdID.Text & "")
    qnty = rs.Fields(0)
    
If qnty < txtQntyPurchased Then
MsgBox "Not enough quantity!", vbCritical, "Error"
Else
insertTbl
updateQnty
grandtotal
Item.ListItems.Clear
PurItem
txtAmnt.Enabled = True
cmdReg.Enabled = False
End If
End Sub

Private Sub cmdCE_Click()
cntRec
Unload Me
frmBill.Show
End Sub

Private Sub cmdDel_Click()
On Error GoTo err_CmdAdd_Click
Dim db_file As String
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim rt As ADODB.Recordset
Dim orig, taken, left As Integer

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

Set rs = conn.Execute("SELECT Qnty FROM tblprod WHERE ProdID = " & Item.ListItems.Item(Item.SelectedItem.Index) & "")
 orig = rs.Fields("Qnty").Value
    
Set rt = conn.Execute("SELECT PurQnty FROM tblsales  WHERE ProdID = " & Item.ListItems.Item(Item.SelectedItem.Index) & " AND RId=" & txtOR.Text & "")
 taken = rt.Fields("PurQnty").Value
left = orig + taken

If MsgBox("Are you sure you want to delete this?", vbYesNo + vbQuestion, "Delete?") = vbNo Then
Exit Sub
Else

With conn
.BeginTrans
.Execute "DELETE FROM tblsales WHERE ProdID = " & Item.ListItems.Item(Item.SelectedItem.Index) & " AND RId=" & txtOR.Text & ""
.Execute "UPDATE tblprod SET Qnty= " & left & "  WHERE ProdID = " & Item.ListItems.Item(Item.SelectedItem.Index) & " "
.CommitTrans
End With
End If
Item.ListItems.Clear
PurItem
grandtotal

exit_err_CmdAdd_Click:
    Exit Sub
 
err_CmdAdd_Click:
    If Err = -2147467259 Then
       MsgBox "Please Add Item!", vbInformation, "Notice"
    Else
       MsgBox "Please Add Item!", vbInformation, "Notice"
       End If
End Sub

Private Sub cmdExit_Click()
Dim ans As String

ans = MsgBox("Do you want to quit?", vbYesNo + vbQuestion, _
"Confirm")
If ans = vbYes Then MsgBox "Bye Teller"
If ans = vbNo Then Exit Sub
frmLogIn.Show
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim db_file As String
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim change As Currency
Dim total As Currency

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
   If MsgBox("Is that all the item you want to sell?", vbYesNo + vbQuestion, "Delete?") = vbNo Then
Exit Sub
Else
    Set rs = conn.Execute("SELECT SUM(Total)FROM tblsales WHERE RId=" & txtOR.Text & "")
    total = rs.Fields(0)
    If txtAmnt < total Then
    MsgBox " Not Enough Money", vbCritical, "Error!"
    cmdPrint.Enabled = False
    Else
    change = txtAmnt - total
    txtChange.Text = FormatCurrency(change)
    cmdOK.Enabled = False
    txtAmnt.Enabled = False
    End If
End If
End Sub

Private Sub cmdPrint_Click()
cmdCE.Enabled = True
cmdOK.Enabled = False
frmBill.Show
PrintOR.Show 1, frmBill
End Sub


Private Sub cmdReg_Click()
frmCustomerReg.Show
Unload Me
End Sub

Private Sub Form_Load()
Dim db_file As String
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim rt As ADODB.Recordset
Dim i As Integer
Dim j As Integer

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
cntRec
txtTellerID.Text = frmLogIn.strName
txtDate.Text = Date
txtQntyPurchased.Enabled = False
txtAmnt.Enabled = False
cmdOK.Enabled = False
cmdAdd.Enabled = False
cmdDel.Enabled = False
cmdCE.Enabled = False
cmdPrint.Enabled = False
cmbProdID.Enabled = False
Set rs = conn.Execute("SELECT ProdID FROM tblprod")
    While Not rs.EOF
cmbProdID.List(i) = rs.Fields!ProdID
rs.MoveNext
i = i + 1
Wend
rs.Close
    Set rs = Nothing
    
    Set rt = conn.Execute("SELECT CustID FROM tblcust")
    While Not rt.EOF
cmbIDCust.List(j) = rt.Fields!CustID
rt.MoveNext
j = j + 1
Wend
rt.Close
    Set rt = Nothing
End Sub
Private Sub txtAmnt_Change()
cmdOK.Enabled = Not (txtAmnt.Text = vbNullString)
End Sub

Private Sub txtAmnt_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Then
ElseIf KeyAscii = 8 Then
ElseIf KeyAscii = 13 Then
cmdOK_Click
Else
KeyAscii = 0
End If
End Sub
Sub grandtotal()
Dim db_file As String
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim total As Currency

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

    Set rs = conn.Execute("SELECT SUM(Total)FROM tblsales tblsales WHERE RId=" & txtOR.Text & " AND Det=#" & txtDate & "# ")
    total = rs.Fields(0)
    txtgrand.Text = FormatCurrency(total)
    
    
End Sub

Private Sub txtChange_Change()
cmdAdd.Enabled = (txtChange.Text = vbNullString)
cmdDel.Enabled = (txtChange.Text = vbNullString)
cmbProdID.Enabled = (txtChange.Text = vbNullString)
cmbIDCust.Enabled = (txtChange.Text = vbNullString)
txtQntyPurchased.Enabled = (txtChange.Text = vbNullString)
cmdPrint.Enabled = Not (txtChange.Text = vbNullString)
End Sub
Private Sub txtgrand_Change()
cmdDel.Enabled = Not (txtgrand.Text = vbNullString)
End Sub

Private Sub txtOR_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Then
Else
KeyAscii = 0
End If
End Sub

Private Sub txtProdPrice_Change()
Dim total As Currency
If txtQntyPurchased.Text = "" Then
       MsgBox "Input Quantity Now!"
       Exit Sub
    Else
    total = txtProdPrice * txtQntyPurchased
    txtTotal.Text = FormatCurrency(total)
End If
End Sub

Private Sub txtQntyPurchased_Change()
Dim total As Currency
cmdAdd.Enabled = Not (txtQntyPurchased.Text = vbNullString)
    If txtQntyPurchased.Text = "" Then
       MsgBox "Input Needed!"
       Exit Sub
    Else
    total = txtProdPrice * txtQntyPurchased
    txtTotal.Text = FormatCurrency(total)
End If
End Sub

Private Sub txtQntyPurchased_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Then
ElseIf KeyAscii = 8 Then
ElseIf KeyAscii = 13 Then
cmdAdd_Click
Else
KeyAscii = 0
End If
End Sub

Sub insertTbl()
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
    
  
  ' Populate the table.
   conn.Execute "INSERT INTO tblsales VALUES('" & txtOR.Text & "','" & txtDate.Text & "','" & _
           txtTellerID.Text & "','" & cmbIDCust.Text & "','" & cmbProdID.Text & "','" & _
          txtQntyPurchased.Text & "','" & txtTotal.Text & "')"
          
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
    Set rs = conn.Execute("SELECT MAX(RId) FROM tblsales")
  If IsNull(rs.Fields(0)) = True Then
  txtOR.Text = 100
  Else
  txtOR.Text = rs.Fields(0) + 1
  End If
End Sub
Sub PurItem()
Dim db_file As String
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim Panoy As ListItem

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
    

With Item
 Set rs = conn.Execute("SELECT ProdID,PurQnty,Total FROM tblsales WHERE RId=" & txtOR.Text & "")
        Do While Not rs.EOF
            Set Panoy = .ListItems.Add(, , rs!ProdID)
                Panoy.SubItems(1) = rs.Fields("PurQnty").Value
                Panoy.SubItems(2) = FormatCurrency(rs.Fields("Total").Value)
            rs.MoveNext
           
        Loop
End With

End Sub


