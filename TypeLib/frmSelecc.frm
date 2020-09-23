VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSelecc 
   Caption         =   "Select DLL to use"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   ScaleHeight     =   4590
   ScaleWidth      =   5280
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3960
      TabIndex        =   4
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdSeleccionar 
      Caption         =   "OK"
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   3960
      Width           =   1215
   End
   Begin VB.ListBox lstDLLs 
      Height          =   2985
      ItemData        =   "frmSelecc.frx":0000
      Left            =   120
      List            =   "frmSelecc.frx":0002
      TabIndex        =   2
      Top             =   840
      Width           =   5055
   End
   Begin VB.TextBox txtDir 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Text            =   "C:\Components\"
      Top             =   120
      Width           =   4455
   End
   Begin VB.CommandButton cmdSeleccFile 
      Caption         =   "..."
      Height          =   495
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
   Begin MSComDlg.CommonDialog cdlgArch 
      Left            =   120
      Top             =   100
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmSelecc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSeleccFile_Click()
    lFlags = BIF_RETURNONLYFSDIRS

    sDir = BrowseForFolder(Me.hWnd, "Seleccionar Directorio", , lFlags)
    
    If Right(sDir, 1) <> "\" Then
        sDir = sDir & "\"
    End If
    
    txtDir.Text = sDir
End Sub
