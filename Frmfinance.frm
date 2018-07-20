VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmfinance 
   Caption         =   "Finance Form"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14055
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   14055
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   120
      Top             =   2640
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=schooldb.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=schooldb.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Finance"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      Height          =   8175
      Left            =   8520
      TabIndex        =   21
      Top             =   360
      Width           =   5175
      Begin VB.CommandButton Command80 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Exit "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   28
         Top             =   3480
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "New Purchase"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   27
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Save "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   26
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Clear "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3600
         TabIndex        =   25
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3600
         TabIndex        =   24
         Top             =   2880
         Width           =   1335
      End
      Begin VB.CommandButton cmdupdate 
         BackColor       =   &H00C0C0C0&
         Caption         =   "<< View Previous"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   23
         Top             =   2880
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "View Next >>"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   22
         Top             =   2880
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Height          =   8175
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   7695
      Begin VB.TextBox Text7 
         DataField       =   "middle name"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   3000
         TabIndex        =   19
         Top             =   1560
         Width           =   2895
      End
      Begin VB.TextBox Text6 
         DataField       =   "Current Balance"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   3000
         TabIndex        =   17
         Top             =   6600
         Width           =   2895
      End
      Begin VB.TextBox Text5 
         DataField       =   "Semester Fees"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   3000
         TabIndex        =   12
         Top             =   5880
         Width           =   2895
      End
      Begin VB.TextBox Text4 
         DataField       =   "other names"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   3000
         TabIndex        =   11
         Top             =   2280
         Width           =   2895
      End
      Begin VB.TextBox Text2 
         DataField       =   "Paid Amount"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   3000
         TabIndex        =   10
         Top             =   7320
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         DataField       =   "registration no"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   3000
         TabIndex        =   9
         Top             =   3720
         Width           =   2895
      End
      Begin VB.TextBox Text16 
         DataField       =   "Course Title"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   3000
         TabIndex        =   4
         Top             =   5160
         Width           =   2895
      End
      Begin VB.TextBox Text15 
         DataField       =   "Gender"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   3000
         TabIndex        =   3
         Top             =   4440
         Width           =   2895
      End
      Begin VB.TextBox Text13 
         DataField       =   "date of admission"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   3000
         TabIndex        =   2
         Top             =   3000
         Width           =   2895
      End
      Begin VB.TextBox Text3 
         DataField       =   "surname"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   3000
         TabIndex        =   1
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Middle Name"
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
         Left            =   1320
         TabIndex        =   20
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Gender"
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
         TabIndex        =   18
         Top             =   4440
         Width           =   2055
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Current Balance"
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
         Left            =   960
         TabIndex        =   16
         Top             =   6600
         Width           =   1935
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Other Names"
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
         TabIndex        =   15
         Top             =   2400
         Width           =   2055
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Paid Amount"
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
         TabIndex        =   14
         Top             =   7440
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Semester Fees"
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
         TabIndex        =   13
         Top             =   5880
         Width           =   2055
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Course Title"
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
         Left            =   960
         TabIndex        =   8
         Top             =   5280
         Width           =   1935
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Admission"
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
         TabIndex        =   7
         Top             =   3000
         Width           =   2055
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Registration No"
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
         TabIndex        =   6
         Top             =   3720
         Width           =   2055
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Surname"
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
         TabIndex        =   5
         Top             =   960
         Width           =   2055
      End
   End
End
Attribute VB_Name = "Frmfinance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub wipe()
Text3.Text = ""
Text7.Text = ""
Text4.Text = ""
Text13.Text = ""
Text6.Text = ""
Text15.Text = ""
Text16.Text = ""
Text5.Text = ""
Text1.Text = ""
Text2.Text = ""

End Sub

Private Sub cmdupdate_Click()
Command10.Enabled = True
Command5.Enabled = True

If Adodc1.Recordset.BOF Then
Exit Sub
Else
Adodc1.Recordset.MovePrevious
End If
End Sub

Private Sub Command1_Click()
Command1.Enabled = False
Command4.Enabled = True
Command10.Enabled = True
Command5.Enabled = True
Me.wipe
End Sub

Private Sub Command10_Click()
Me.wipe
Command10.Enabled = False
Command5.Enabled = False
Command4.Enabled = True
Command1.Enabled = True
End Sub

Private Sub Command2_Click()
Data1.Recordset.Update
End Sub

Private Sub Command3_Click()
Command4.Enabled = True
Command10.Enabled = True

Command5.Enabled = True
If Adodc1.Recordset.EOF Then
Exit Sub
Exit Sub
Else
Adodc1.Recordset.MoveNext
End If
End Sub

Private Sub Command4_Click()
Command4.Enabled = False

Command10.Enabled = True
Command1.Enabled = True
Command5.Enabled = True
If Text1.Text = "" Then
Exit Sub
Text1.SetFocus
Exit Sub
Else
Adodc1.Recordset.AddNew
'Adodc1.Recordset.Update
End If
End Sub

Private Sub Command5_Click()
Command1.Enabled = True
If Text1.Text = "" Then
Exit Sub
Else
Adodc1.Recordset.Delete
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF Then
If Adodc1.Recordset.BOF Then
Exit Sub
Else
Adodc1.Recordset.MoveLast
End If
End If
End If
Command5.Enabled = False

End Sub

Private Sub Command6_Click()
Data1.Recordset.MoveLast
If Adodc1.Recordset.EOF = True Then
MsgBox "This is the Last Record", vbInformation
End If
End Sub

Private Sub Command7_Click()

End Sub

Private Sub Command8_Click()
Adodc1.Recordset.MoveFirst
If Data1.Recordset.BOF Then
Exit Sub
End If
End Sub

Private Sub Command9_Click()
Command10.Enabled = True
Command5.Enabled = True

End Sub

Private Sub Command80_Click()
Unload Me
MDIForm1.Show
End Sub

Private Sub Form_Load()
'Command4.Enabled = False
'Command10.Enabled = False
'Command5.Enabled = False
End Sub


