VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSalesReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database II Project"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
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
      Height          =   495
      Left            =   8160
      TabIndex        =   1
      Top             =   6240
      Width           =   2175
   End
   Begin VB.CommandButton cmdYearly 
      Caption         =   "Annual"
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
      Left            =   7920
      TabIndex        =   14
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton cmdMOntly 
      Caption         =   "Monthly"
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
      Left            =   6600
      TabIndex        =   13
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Daily"
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
      Left            =   5400
      TabIndex        =   12
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "CLEAR"
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
      Left            =   9360
      TabIndex        =   4
      Top             =   5640
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "SELECT"
      Height          =   1815
      Left            =   840
      TabIndex        =   5
      Top             =   5400
      Width           =   9735
      Begin VB.ComboBox cmbMonth 
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
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cmbYear 
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
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cmbDay 
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
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lbldate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   7320
         TabIndex        =   17
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "*As of "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6720
         TabIndex        =   16
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Total Profit"
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
         Left            =   5760
         TabIndex        =   15
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "YEAR"
         Height          =   375
         Left            =   3240
         TabIndex        =   11
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "MONTH"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "DAY"
         Height          =   375
         Left            =   1800
         TabIndex        =   9
         Top             =   720
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "MENU"
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
      Left            =   9960
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin MSComctlLib.ListView SalesList 
      Height          =   4455
      Left            =   840
      TabIndex        =   2
      Top             =   720
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   7858
      SortKey         =   1
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "OR ID"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date"
         Object.Width           =   2470
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Teller ID"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Customer ID"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Product ID"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Qnty"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Total Sales"
         Object.Width           =   2822
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "SALES REPORT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   0
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "frmSalesReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Matt As ListItem
Sub yearload()
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
    
With SalesList
 Set rs = conn.Execute("SELECT * FROM tblsales WHERE year(Det)='" & cmbYear.Text & "';  ")
        Do While Not rs.EOF
            Set Matt = .ListItems.Add(, , rs!RId)
                Matt.SubItems(1) = rs.Fields("Det").Value
                Matt.SubItems(2) = rs.Fields("TellerId").Value
                Matt.SubItems(3) = rs.Fields("CustID").Value
                Matt.SubItems(4) = rs.Fields("ProdId").Value
                Matt.SubItems(5) = rs.Fields("PurQnty").Value
                Matt.SubItems(6) = FormatCurrency(rs.Fields("Total").Value)
            rs.MoveNext
           
        Loop
End With

End Sub

Sub monthload()
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
    
With SalesList
 Set rs = conn.Execute("SELECT * FROM tblsales WHERE month(Det)='" & cmbMonth.Text & "' AND year(Det)='" & cmbYear.Text & "';  ")
        Do While Not rs.EOF
            Set Matt = .ListItems.Add(, , rs!RId)
                Matt.SubItems(1) = rs.Fields("Det").Value
                Matt.SubItems(2) = rs.Fields("TellerId").Value
                Matt.SubItems(3) = rs.Fields("CustID").Value
                Matt.SubItems(4) = rs.Fields("ProdId").Value
                Matt.SubItems(5) = rs.Fields("PurQnty").Value
                Matt.SubItems(6) = FormatCurrency(rs.Fields("Total").Value)
            rs.MoveNext
           
        Loop
End With
End Sub

Private Sub cmdClear_Click()
Dim cbo As Control

For Each cbo In Me
If TypeOf cbo Is ComboBox Then
cbo.ListIndex = -1
End If
Next cbo

SalesList.ListItems.Clear
txtTotal.Text = ""
lblDate.Caption = ""
Label3.Visible = False
End Sub
Sub dailyload()
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
    
intYear = cmbYear.Text
intMonth = cmbMonth.Text
intDay = cmbDay.Text
dtmNewDate = DateSerial(intYear, intMonth, intDay)
With SalesList
 Set rs = conn.Execute("SELECT * FROM tblsales WHERE Det=#" & dtmNewDate & "# ")
        Do While Not rs.EOF
            Set Matt = .ListItems.Add(, , rs!RId)
                Matt.SubItems(1) = rs.Fields("Det").Value
                Matt.SubItems(2) = rs.Fields("TellerId").Value
                Matt.SubItems(3) = rs.Fields("CustID").Value
                Matt.SubItems(4) = rs.Fields("ProdId").Value
                Matt.SubItems(5) = rs.Fields("PurQnty").Value
                Matt.SubItems(6) = FormatCurrency(rs.Fields("Total").Value)
            rs.MoveNext
           
        Loop
End With
exit_err_CmdAdd_Click:
    Exit Sub
 
err_CmdAdd_Click:
    If Err = -2147467259 Then
       MsgBox "Valid Dates Needed!", vbExclamation, "Confirm"
    Else
        MsgBox "Valid Dates Needed!", vbExclamation, "Confirm"
    End If
End Sub
Private Sub cmdMenu_Click()
Menu.Show
Unload Me
End Sub

Private Sub cmdMOntly_Click()
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

Label3.Visible = True
lblDate.Caption = cmbMonth.Text + " " + cmbYear.Text
SalesList.ListItems.Clear
 monthload
    Set rs = conn.Execute("SELECT SUM(Total)FROM tblsales WHERE month(Det)='" & cmbMonth.Text & "' AND year(Det)='" & cmbYear.Text & "';  ")
    
    
If IsNull(rs.Fields(0)) = False Then
    total = rs.Fields(0)
   txtTotal.Text = FormatCurrency(total)
Else
  MsgBox "Record Not Found!"
  txtTotal.Text = ""
End If
  
End Sub

Private Sub cmdTotal_Click()
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

SalesList.ListItems.Clear
dailyload
Label3.Visible = True
lblDate.Caption = cmbMonth.Text + " " + cmbDay.Text + " " + cmbYear.Text
    Set rs = conn.Execute("SELECT SUM(Total)FROM tblsales WHERE month(Det)='" & cmbMonth.Text & "' AND year(Det)='" & cmbYear.Text & "' AND day(Det)='" & cmbDay.Text & "';  ")
    

If IsNull(rs.Fields(0)) = False Then
    total = rs.Fields(0)
   txtTotal.Text = FormatCurrency(total)
Else
 txtTotal.Text = ""
  MsgBox "Record Not Found!"
End If
End Sub
Sub LoadDB()
Dim db_file As String
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim Matt As ListItem

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
    

With SalesList
 Set rs = conn.Execute("SELECT * FROM tblsales")
        Do While Not rs.EOF
            Set Matt = .ListItems.Add(, , rs!RId)
                Matt.SubItems(1) = rs.Fields("Det").Value
                Matt.SubItems(2) = rs.Fields("TellerId").Value
                Matt.SubItems(3) = rs.Fields("CustID").Value
                Matt.SubItems(4) = rs.Fields("ProdId").Value
                Matt.SubItems(5) = rs.Fields("PurQnty").Value
                Matt.SubItems(6) = FormatCurrency(rs.Fields("Total").Value)
            rs.MoveNext
           
        Loop
End With

End Sub
Sub createTbl()
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
             "CREATE TABLE tblsales(" & _
            "RId INTEGER NOT NULL," & _
            "Det  DATE  NOT NULL," & _
            "TellerID   INTEGER  NOT NULL," & _
             "CustID   INTEGER  NOT NULL," & _
             "ProdID   INTEGER  NOT NULL," & _
            "PurQnty  INTEGER  NOT NULL," & _
            "Total      CURRENCY   NOT NULL)"
          
exit_err_CmdAdd_Click:
    Exit Sub
 
err_CmdAdd_Click:
    If Err = -2147467259 Then
        MsgBox "Table Created!"
    Else
        MsgBox "Table Exists!"
    End If
                
End Sub
Private Sub cmdView_Click()
SalesList.ListItems.Clear
dailyload
End Sub

Private Sub cmdYearly_Click()
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


Label3.Visible = True
lblDate.Caption = cmbYear.Text
SalesList.ListItems.Clear
yearload
    Set rs = conn.Execute("SELECT SUM(Total)FROM tblsales WHERE year(Det)='" & cmbYear.Text & "';  ")
    
    
If IsNull(rs.Fields(0)) = False Then
    total = rs.Fields(0)
    txtTotal.Text = FormatCurrency(total)
Else
  MsgBox "Record Not Found!"
  txtTotal.Text = ""
End If
  
End Sub

Private Sub Form_Load()
Dim a As Integer
Dim b As Integer
Dim c As Integer
createTbl
LoadDB
Label3.Visible = False
a = 0
Do While a < 31
a = a + 1
cmbDay.AddItem a
Loop


b = 0
Do While b < 12
b = b + 1
cmbMonth.AddItem b
Loop


c = 2010
Do While c < 2015
c = c + 1
cmbYear.AddItem c
Loop


End Sub
