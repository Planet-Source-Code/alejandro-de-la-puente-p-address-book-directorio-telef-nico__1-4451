VERSION 5.00
Begin VB.Form FrmEnglishEspa�ol 
   BackColor       =   &H00800000&
   Caption         =   "English - Espa�ol"
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
   Begin VB.CommandButton cmdEspa�ol 
      Caption         =   "&Espa�ol"
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
   Begin VB.Label lbEspa�ol 
      BackColor       =   &H00800000&
      Caption         =   "Por favor haz clic aqu� si hablas espa�ol:"
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
Attribute VB_Name = "FrmEnglishEspa�ol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEnglish_Click()
    FrmAddressBook.Show
    Unload FrmEnglishEspa�ol
    Unload FrmDirectorioTelef�nico
End Sub
Private Sub cmdEspa�ol_Click()
    FrmDirectorioTelef�nico.Show
    Unload FrmEnglishEspa�ol
    Unload FrmAddressBook
End Sub
