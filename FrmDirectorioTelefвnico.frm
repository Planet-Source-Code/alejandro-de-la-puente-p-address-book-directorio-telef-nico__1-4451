VERSION 5.00
Begin VB.Form FrmDirectorioTelefónico 
   BackColor       =   &H00800000&
   Caption         =   "Directorio Telefónico"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7335
   ForeColor       =   &H00800000&
   Icon            =   "FrmDirectorioTelefónico.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   7335
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
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
      TabIndex        =   24
      Top             =   5760
      Width           =   7095
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         Height          =   255
         Left            =   6000
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "&Modificar"
         Height          =   255
         Left            =   4560
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   255
         Left            =   3000
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "&Nuevo"
         Height          =   255
         Left            =   1560
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "&Guardar"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame FrameContactos 
      BackColor       =   &H00800000&
      Caption         =   "Contactos:"
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
      TabIndex        =   23
      Top             =   3360
      Width           =   7095
      Begin VB.ListBox LstContactos 
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
   Begin VB.Frame FrameDatosPersonales 
      BackColor       =   &H00800000&
      Caption         =   "Datos Personales:"
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
      TabIndex        =   0
      Top             =   120
      Width           =   7095
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
      Begin VB.TextBox txtTeléfono02 
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
      Begin VB.TextBox txtTeléfono01 
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
      Begin VB.TextBox txtCompañia 
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
      Begin VB.TextBox txtDirección 
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
      Begin VB.TextBox txtNombre 
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
         TabIndex        =   22
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label lbFax 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         Caption         =   "No. de Fax:"
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
         TabIndex        =   21
         Top             =   2160
         Width           =   1200
      End
      Begin VB.Label lbCelular 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         Caption         =   "No. de Celular:"
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
         TabIndex        =   20
         Top             =   2160
         Width           =   1560
      End
      Begin VB.Label lbTeléfono02 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         Caption         =   "Teléfono 02:"
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
         TabIndex        =   19
         Top             =   1680
         Width           =   1305
      End
      Begin VB.Label lbTeléfono01 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         Caption         =   "Teléfono 01:"
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
         TabIndex        =   18
         Top             =   1680
         Width           =   1305
      End
      Begin VB.Label lbCompañia 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         Caption         =   "Compañia:"
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
         Top             =   1200
         Width           =   1125
      End
      Begin VB.Label lbDirección 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         Caption         =   "Dirección:"
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
         TabIndex        =   16
         Top             =   720
         Width           =   1065
      End
      Begin VB.Label lbNombre 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         Caption         =   "Nombre:"
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
         TabIndex        =   15
         Top             =   240
         Width           =   900
      End
   End
End
Attribute VB_Name = "FrmDirectorioTelefónico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim midata As Database
Dim mirecord As Recordset
Dim SQL As String
Private Sub cmdEliminar_Click()
    Set midata = OpenDatabase(App.Path + "\" + "DBEnglishEspañol.mdb")
    SQL = "SELECT * FROM TblContactos"
    Set mirecord = midata.OpenRecordset(SQL)
        Do Until mirecord.EOF
            If LstContactos.Text = mirecord!CmpNombres Then
                mirecord.Delete
                LstContactos.RemoveItem (LstContactos.ListIndex)
            End If
            mirecord.MoveNext
        Loop
    Call cmdNuevo_Click
End Sub
Private Sub cmdGuardar_Click()
    Set midata = OpenDatabase(App.Path + "\" + "DBEnglishEspañol.mdb")
    Set mirecord = midata.OpenRecordset("TblContactos")
    With mirecord
        .AddNew
        !CmpNombres = Trim(txtNombre.Text)
        !CmpDirección = Trim(txtDirección.Text)
        !CmpCompañia = Trim(txtCompañia.Text)
        !CmpTeléfono01 = Trim(txtTeléfono01.Text)
        !CmpTeléfono02 = Trim(txtTeléfono02.Text)
        !CmpCelular = Trim(txtCelular.Text)
        !CmpFax = Trim(txtFax.Text)
        !CmpEmail = Trim(txtEmail.Text)
        .Update
    End With
    midata.Close
    LstContactos.AddItem txtNombre.Text
    Call cmdNuevo_Click
End Sub
Private Sub cmdModificar_Click()
    Set midata = OpenDatabase(App.Path + "\" + "DBEnglishEspañol.mdb")
    Set mirecord = midata.OpenRecordset("TblContactos")
        Do Until mirecord.EOF
            If LstContactos.Text = mirecord!CmpNombres Then
                LstContactos.RemoveItem (LstContactos.ListIndex)
                LstContactos.AddItem txtNombre.Text
                With mirecord
                    .Edit
                    !CmpNombres = Trim(txtNombre.Text)
                    !CmpDirección = Trim(txtDirección.Text)
                    !CmpCompañia = Trim(txtCompañia.Text)
                    !CmpTeléfono01 = Trim(txtTeléfono01.Text)
                    !CmpTeléfono02 = Trim(txtTeléfono02.Text)
                    !CmpCelular = Trim(txtCelular.Text)
                    !CmpFax = Trim(txtFax.Text)
                    !CmpEmail = Trim(txtEmail.Text)
                    .Update
                End With
            End If
            mirecord.MoveNext
            'LstContactos.AddItem txtNombre.Text
        Loop
End Sub
Private Sub cmdNuevo_Click()
    txtNombre.Text = ""
    txtDirección.Text = ""
    txtCompañia.Text = ""
    txtTeléfono01.Text = ""
    txtTeléfono02.Text = ""
    txtCelular.Text = ""
    txtFax.Text = ""
    txtEmail.Text = ""
    txtNombre.SetFocus
End Sub
Private Sub cmdSalir_Click()
    Unload Me
    End
End Sub
Private Sub Form_Load()
    Set midata = OpenDatabase(App.Path + "\" + "DBEnglishEspañol.mdb")
    Set mirecord = midata.OpenRecordset("TblContactos")
    mirecord.MoveFirst
    Do Until mirecord.EOF
        LstContactos.AddItem mirecord.Fields("CmpNombres")
        mirecord.MoveNext
    Loop
    midata.Close
End Sub
Private Sub LstContactos_Click()
    On Error Resume Next
        Set midata = OpenDatabase(App.Path + "\" + "DBEnglishEspañol.mdb")
        Set mirecord = midata.OpenRecordset("TblContactos")
        mirecord.MoveFirst
        Do Until mirecord.EOF
            If LstContactos.Text = mirecord!CmpNombres Then
                txtNombre.Text = mirecord!CmpNombres
                txtDirección.Text = mirecord!CmpDirección
                txtCompañia.Text = mirecord!CmpCompañia
                txtTeléfono01.Text = mirecord!CmpTeléfono01
                txtTeléfono02.Text = mirecord!CmpTeléfono02
                txtCelular.Text = mirecord!CmpCelular
                txtFax.Text = mirecord!CmpFax
                txtEmail.Text = mirecord!CmpEmail
            End If
            mirecord.MoveNext
        Loop
End Sub
