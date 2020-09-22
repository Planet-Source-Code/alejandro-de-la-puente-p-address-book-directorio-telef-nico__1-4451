VERSION 5.00
Begin VB.Form FrmEnglishEspañol 
   BackColor       =   &H00800000&
   Caption         =   "English - Español"
   ClientHeight    =   1335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4815
   ForeColor       =   &H00800000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   4815
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEspañol 
      Caption         =   "&Español"
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdEnglish 
      Caption         =   "&English"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lbEspañol 
      BackColor       =   &H00800000&
      Caption         =   "Por favor haz clic aquí si hablas español:"
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
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lbEnglish 
      BackColor       =   &H00800000&
      Caption         =   "Please click here if you speak english:"
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
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "FrmEnglishEspañol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEnglish_Click()
    FrmAddressBook.Show
    Unload FrmEnglishEspañol
    Unload FrmDirectorioTelefónico
End Sub
Private Sub cmdEspañol_Click()
    FrmDirectorioTelefónico.Show
    Unload FrmEnglishEspañol
    Unload FrmAddressBook
End Sub
