VERSION 5.00
Begin VB.Form Menu 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database II Project"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   4470
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Caption         =   "Choose Below"
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   4215
      Begin VB.CommandButton cmsCustomer 
         Caption         =   "CUSTOMER"
         Height          =   615
         Left            =   1200
         TabIndex        =   4
         Top             =   2040
         Width           =   1815
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "LOG OUT"
         Height          =   615
         Left            =   1200
         TabIndex        =   5
         Top             =   3720
         Width           =   1815
      End
      Begin VB.CommandButton cmdTeller 
         Caption         =   "TELLER"
         Height          =   495
         Left            =   1200
         TabIndex        =   2
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CommandButton cmdReport 
         Caption         =   "SALES REPORT"
         Height          =   615
         Left            =   1200
         TabIndex        =   3
         Top             =   2880
         Width           =   1815
      End
      Begin VB.CommandButton cmdProduct 
         BackColor       =   &H80000009&
         Caption         =   "PRODUCT"
         Height          =   495
         Left            =   1200
         MaskColor       =   &H8000000C&
         TabIndex        =   1
         Top             =   600
         Width           =   1815
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "Water-Relling Station"
      BeginProperty Font 
         Name            =   "@Vyper Expanded"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub createDB(Path As String)
Dim OurDB As Database
Set OurDB = CreateDatabase(Path, dbLangGeneral)
OurDB.Close
Set OurDB = Nothing

End Sub
Private Sub cmdExit_Click()
Dim ans As String

ans = MsgBox("Do you want to log out?", vbYesNo + vbQuestion, _
"Confirm")
If ans = vbYes Then MsgBox "Bye Admin", vbInformation, "Done"
If ans = vbNo Then Exit Sub
frmLogIn.Show
Unload Me
End Sub
Private Sub cmdProduct_Click()
FindProduct.Show
Me.Hide
End Sub
Private Sub cmdReport_Click()
frmSalesReport.Show
Me.Hide
End Sub
Private Sub cmdTeller_Click()
frmTeller.Show
Me.Hide
End Sub
Private Sub cmsCustomer_Click()
frmCustomer.Show
Me.Hide
End Sub
Private Sub Form_Load()

If (Dir("" & App.Path & "\OurDB.mdb") = "") Then
createDB "" & App.Path & "\OurDB.mdb"
MsgBox "Welcome Admin, Please Proceed.", vbInformation, "Message"
Else
MsgBox "Welcome Admin, Please Proceed.", vbInformation, "Message"
End If
End Sub

