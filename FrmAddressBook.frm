VERSION 5.00
Begin VB.Form FrmAddressBook 
   BackColor       =   &H00800000&
   Caption         =   "Address Book"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7335
   Icon            =   "FrmAddressBook.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   7335
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FramePersonalInformation 
      BackColor       =   &H00800000&
      Caption         =   "Personal Information:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   3135
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   7095
      Begin VB.TextBox txtName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   360
         Left            =   1800
         TabIndex        =   1
         Top             =   240
         Width           =   5175
      End
      Begin VB.TextBox txtAddress 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   360
         Left            =   1800
         TabIndex        =   2
         Top             =   720
         Width           =   5175
      End
      Begin VB.TextBox txtCompany 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   360
         Left            =   1800
         TabIndex        =   3
         Text            =   "----------"
         Top             =   1200
         Width           =   5175
      End
      Begin VB.TextBox txtTelephone01 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   360
         Left            =   1800
         TabIndex        =   4
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtTelephone02 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   360
         Left            =   6000
         TabIndex        =   5
         Text            =   "0000000"
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtCelular 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   360
         Left            =   1800
         TabIndex        =   6
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox txtFax 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   360
         Left            =   6000
         TabIndex        =   7
         Text            =   "0000000"
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox txtEmail 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   360
         Left            =   1800
         TabIndex        =   8
         Top             =   2640
         Width           =   5175
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   240
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   690
      End
      Begin VB.Label lbAddress 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         Caption         =   "Address:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   240
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   945
      End
      Begin VB.Label lbCompany 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         Caption         =   "Company:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   240
         Left            =   120
         TabIndex        =   22
         Top             =   1200
         Width           =   1050
      End
      Begin VB.Label lbTelephone01 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         Caption         =   "Telephone 01:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   240
         Left            =   120
         TabIndex        =   21
         Top             =   1680
         Width           =   1500
      End
      Begin VB.Label lbTelephone02 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         Caption         =   "Telephone 02:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   240
         Left            =   4320
         TabIndex        =   20
         Top             =   1680
         Width           =   1500
      End
      Begin VB.Label lbCelular 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         Caption         =   "Celular Num:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   240
         Left            =   120
         TabIndex        =   19
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label lbFax 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         Caption         =   "Fax Number:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   240
         Left            =   4320
         TabIndex        =   18
         Top             =   2160
         Width           =   1320
      End
      Begin VB.Label lbEmail 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         Caption         =   "E-Mail:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   240
         Left            =   120
         TabIndex        =   17
         Top             =   2640
         Width           =   735
      End
   End
   Begin VB.Frame FrameContacts 
      BackColor       =   &H00800000&
      Caption         =   "Contacts:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   2295
      Left            =   120
      TabIndex        =   15
      Top             =   3360
      Width           =   7095
      Begin VB.ListBox LstContacts 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   1740
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   12
         Top             =   360
         Width           =   6855
      End
   End
   Begin VB.Frame FrameButtons 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   5760
      Width           =   7095
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         Height          =   255
         Left            =   1560
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   255
         Left            =   3000
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdModify 
         Caption         =   "&Modify"
         Height          =   255
         Left            =   4560
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "&Exit"
         Height          =   255
         Left            =   6000
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "FrmAddressBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mydata As Database
Dim myrecord As Recordset
Dim SQL As String
Private Sub cmdDelete_Click()
    Set mydata = OpenDatabase(App.Path + "\" + "DBEnglishEspañol.mdb")
    SQL = "SELECT * FROM TblContacts"
    Set myrecord = mydata.OpenRecordset(SQL)
        Do Until myrecord.EOF
            If LstContacts.Text = myrecord!CmpNames Then
                myrecord.Delete
                LstContacts.RemoveItem (LstContacts.ListIndex)
            End If
            myrecord.MoveNext
        Loop
    Call cmdNew_Click
End Sub
Private Sub cmdSave_Click()
    Set mydata = OpenDatabase(App.Path + "\" + "DBEnglishEspañol.mdb")
    Set myrecord = mydata.OpenRecordset("TblContacts")
    With myrecord
        .AddNew
        !CmpNames = Trim(txtName.Text)
        !CmpAddress = Trim(txtAddress.Text)
        !CmpCompany = Trim(txtCompany.Text)
        !CmpTelephone01 = Trim(txtTelephone01.Text)
        !CmpTelephone02 = Trim(txtTelephone02.Text)
        !CmpCelular = Trim(txtCelular.Text)
        !CmpFax = Trim(txtFax.Text)
        !CmpEmail = Trim(txtEmail.Text)
        .Update
    End With
    mydata.Close
    LstContacts.AddItem txtName.Text
    Call cmdNew_Click
End Sub
Private Sub cmdModify_Click()
    Set mydata = OpenDatabase(App.Path + "\" + "DBEnglishEspañol.mdb")
    Set myrecord = mydata.OpenRecordset("TblContacts")
        Do Until myrecord.EOF
            If LstContacts.Text = myrecord!CmpNames Then
                LstContacts.RemoveItem (LstContacts.ListIndex)
                LstContacts.AddItem txtName.Text
                With myrecord
                    .Edit
                    !CmpNames = Trim(txtName.Text)
                    !CmpAddress = Trim(txtAddress.Text)
                    !CmpCompany = Trim(txtCompany.Text)
                    !CmpTelephone01 = Trim(txtTelephone01.Text)
                    !CmpTelephone02 = Trim(txtTelephone02.Text)
                    !CmpCelular = Trim(txtCelular.Text)
                    !CmpFax = Trim(txtFax.Text)
                    !CmpEmail = Trim(txtEmail.Text)
                    .Update
                End With
            End If
            myrecord.MoveNext
        Loop
End Sub
Private Sub cmdNew_Click()
    txtName.Text = ""
    txtAddress.Text = ""
    txtCompany.Text = ""
    txtTelephone01.Text = ""
    txtTelephone02.Text = ""
    txtCelular.Text = ""
    txtFax.Text = ""
    txtEmail.Text = ""
    txtName.SetFocus
End Sub
Private Sub cmdExit_Click()
    Unload Me
    End
End Sub
Private Sub Form_Load()
    Set mydata = OpenDatabase(App.Path + "\" + "DBEnglishEspañol.mdb")
    Set myrecord = mydata.OpenRecordset("TblContacts")
    myrecord.MoveFirst
    Do Until myrecord.EOF
        LstContacts.AddItem myrecord.Fields("CmpNames")
        myrecord.MoveNext
    Loop
    mydata.Close
End Sub
Private Sub LstContacts_Click()
    On Error Resume Next
        Set mydata = OpenDatabase(App.Path + "\" + "DBEnglishEspañol.mdb")
        Set myrecord = mydata.OpenRecordset("TblContacts")
        myrecord.MoveFirst
        Do Until myrecord.EOF
            If LstContacts.Text = myrecord!CmpNames Then
                txtName.Text = myrecord!CmpNames
                txtAddress.Text = myrecord!CmpAddress
                txtCompany.Text = myrecord!CmpCompany
                txtTelephone01.Text = myrecord!CmpTelephone01
                txtTelephone02.Text = myrecord!CmpTelephone02
                txtCelular.Text = myrecord!CmpCelular
                txtFax.Text = myrecord!CmpFax
                txtEmail.Text = myrecord!CmpEmail
            End If
            myrecord.MoveNext
        Loop
End Sub

