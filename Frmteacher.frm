VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmteacher 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Teacher's Form"
   ClientHeight    =   9735
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13530
   LinkTopic       =   "Form1"
   ScaleHeight     =   9735
   ScaleWidth      =   13530
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   0
      Top             =   3240
      Width           =   2295
      _ExtentX        =   4048
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
      RecordSource    =   "teacher"
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Height          =   8535
      Left            =   1080
      TabIndex        =   0
      Top             =   600
      Width           =   10815
      Begin VB.ComboBox Combo2 
         DataField       =   "employment category"
         DataSource      =   "Adodc1"
         Height          =   315
         ItemData        =   "Frmteacher.frx":0000
         Left            =   2400
         List            =   "Frmteacher.frx":000A
         TabIndex        =   24
         Top             =   5520
         Width           =   2895
      End
      Begin VB.ComboBox Combo1 
         DataField       =   "Gender"
         DataSource      =   "Adodc1"
         Height          =   315
         ItemData        =   "Frmteacher.frx":001C
         Left            =   2400
         List            =   "Frmteacher.frx":0026
         TabIndex        =   23
         Top             =   2520
         Width           =   2895
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFC0C0&
         Height          =   6615
         Left            =   5520
         TabIndex        =   15
         Top             =   720
         Width           =   5055
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
            TabIndex        =   22
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
            TabIndex        =   21
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
            TabIndex        =   20
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
            TabIndex        =   19
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
            TabIndex        =   18
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
            TabIndex        =   17
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
            TabIndex        =   16
            Top             =   2880
            Width           =   1335
         End
      End
      Begin VB.TextBox Text3 
         DataField       =   "Surname"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   2400
         TabIndex        =   6
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox Text13 
         DataField       =   "date of employment"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   2400
         TabIndex        =   5
         Top             =   3960
         Width           =   2895
      End
      Begin VB.TextBox Text15 
         DataField       =   "office no"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   2400
         TabIndex        =   4
         Top             =   4680
         Width           =   2895
      End
      Begin VB.TextBox Text16 
         DataField       =   "lecture hours"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   2400
         TabIndex        =   3
         Top             =   6360
         Width           =   2895
      End
      Begin VB.TextBox Text4 
         DataField       =   "OtherNames"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   2400
         TabIndex        =   2
         Top             =   1680
         Width           =   2895
      End
      Begin VB.TextBox Text6 
         DataField       =   "Staff No"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   2400
         TabIndex        =   1
         Top             =   3360
         Width           =   2895
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
         Left            =   240
         TabIndex        =   14
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Staff No"
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
         Left            =   240
         TabIndex        =   13
         Top             =   3360
         Width           =   2055
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Employment"
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
         Left            =   240
         TabIndex        =   12
         Top             =   3960
         Width           =   2055
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Lecture Hours per Week"
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
         Left            =   360
         TabIndex        =   11
         Top             =   6480
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
         Left            =   240
         TabIndex        =   10
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Employment Category"
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
         Left            =   360
         TabIndex        =   9
         Top             =   5520
         Width           =   1935
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Office No"
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
         Left            =   240
         TabIndex        =   8
         Top             =   4680
         Width           =   2055
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
         Left            =   240
         TabIndex        =   7
         Top             =   2520
         Width           =   2055
      End
   End
End
Attribute VB_Name = "Frmteacher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub wipe()
Text3.Text = ""
Text4.Text = ""
Text13.Text = ""
Text6.Text = ""
Text15.Text = ""
Text16.Text = ""
'Text1.Text = ""
'Text2.Text = ""
Combo1.Text = ""
Combo2.Text = ""
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
If Text6.Text = "" Then
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
If Text6.Text = "" Then
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



