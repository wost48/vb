VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form PrintOR 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Official Receipt"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "EXIT"
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
      Left            =   360
      TabIndex        =   18
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "PRINT"
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
      Left            =   360
      TabIndex        =   17
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Caption         =   "Purchased  Products Breakdown"
      Height          =   3015
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   4455
      Begin MSComctlLib.ListView ListView1 
         Height          =   2175
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   3836
         SortKey         =   1
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   1076
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Width           =   1411
         EndProperty
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000A&
         X1              =   0
         X2              =   4440
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label10 
         BackColor       =   &H80000005&
         Caption         =   "Total(PhP)"
         Height          =   255
         Left            =   3480
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000005&
         Caption         =   "Qnty."
         Height          =   255
         Left            =   2760
         TabIndex        =   12
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000005&
         Caption         =   "Price(PhP)"
         Height          =   255
         Left            =   1920
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000005&
         Caption         =   "Description"
         Height          =   255
         Left            =   840
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000005&
         Caption         =   "ID"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Label lblCustID 
      BackColor       =   &H80000005&
      Height          =   255
      Left            =   1080
      TabIndex        =   25
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblORID 
      BackColor       =   &H80000005&
      Height          =   255
      Left            =   840
      TabIndex        =   24
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label14 
      BackColor       =   &H80000005&
      Caption         =   "Customer ID:"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblChange 
      BackColor       =   &H8000000E&
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
      Left            =   3480
      TabIndex        =   21
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Label lblAmnt 
      BackColor       =   &H8000000E&
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
      Left            =   3480
      TabIndex        =   20
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label lblGrandTotal 
      BackColor       =   &H8000000E&
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
      Left            =   3480
      TabIndex        =   19
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label Label13 
      BackColor       =   &H8000000E&
      Caption         =   "Change:"
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
      Left            =   2640
      TabIndex        =   16
      Top             =   5880
      Width           =   735
   End
   Begin VB.Label Label12 
      BackColor       =   &H80000005&
      Caption         =   "Amnt Received:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   15
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000005&
      Caption         =   "Grand Total:"
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
      Left            =   2280
      TabIndex        =   14
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label lblTellerID 
      BackColor       =   &H80000005&
      Height          =   255
      Left            =   3840
      TabIndex        =   7
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000005&
      Caption         =   "OR No:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000005&
      Caption         =   "Teller ID:"
      Height          =   255
      Left            =   3120
      TabIndex        =   5
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000005&
      Caption         =   "Time:"
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
      Caption         =   "Date:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   375
   End
   Begin VB.Label lblTime 
      BackColor       =   &H80000005&
      Height          =   255
      Left            =   3600
      TabIndex        =   2
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblDate 
      BackColor       =   &H80000005&
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      X1              =   240
      X2              =   4320
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "Water-Refilling Station Receipt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "PrintOR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
PurItem
Me.Top = (Screen.Height - Me.Height) / 2
Me.left = (Screen.Width - Me.Width) / 2
lblDate.Caption = Format$(Date, "m/d/yyyy")
lblTime.Caption = Format$(Time, "h:nn AM/PM")

lblGrandTotal.Caption = frmBill.txtgrand.Text
lblORID.Caption = frmBill.txtOR.Text
lblTellerID.Caption = frmBill.txtTellerID.Text
lblAmnt.Caption = FormatCurrency(frmBill.txtAmnt.Text)
lblChange.Caption = frmBill.txtChange.Text
lblCustID.Caption = frmBill.cmbIDCust.Text
End Sub
Private Sub cmdPrint_Click()
cmdPrint.Visible = False
cmdExit.Visible = False
PrintForm
cmdPrint.Visible = True
cmdExit.Visible = True
End Sub
Private Sub cmdExit_Click()
Unload Me
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
    

With ListView1
 Set rs = conn.Execute("SELECT tblsales.ProdID,PurQnty,Total,tblprod.Price,tblprod.Description FROM tblprod,tblsales WHERE RId=" & frmBill.txtOR.Text & " AND Det=#" & frmBill.txtDate.Text & "#  AND tblsales.ProdID = tblprod.ProdID")
        Do While Not rs.EOF
            Set Panoy = .ListItems.Add(, , rs!ProdID)
                Panoy.SubItems(1) = rs.Fields("Description").Value
                Panoy.SubItems(2) = rs.Fields("Price").Value
                Panoy.SubItems(3) = rs.Fields("PurQnty").Value
                Panoy.SubItems(4) = rs.Fields("Total").Value
            rs.MoveNext
           
        Loop
End With

End Sub



