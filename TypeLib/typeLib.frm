VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmObjBrowser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Object Browser"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   5175
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "Documentation"
      Height          =   1455
      Left            =   120
      TabIndex        =   7
      Top             =   3840
      Width           =   4935
      Begin VB.TextBox txtDocumentation 
         Height          =   1095
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Method/Property"
      Height          =   2415
      Left            =   2640
      TabIndex        =   5
      Top             =   1320
      Width           =   2415
      Begin VB.ListBox List2 
         Height          =   2010
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Class Relationed"
      Height          =   2415
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   2415
      Begin VB.ListBox List1 
         Height          =   2010
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Analyze"
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog cdlgArch 
      Left            =   240
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   495
      Left            =   4560
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Text            =   "C:\Components\eAgent.dll"
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmObjBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Given a VB form with two ListBox controls, a TextBox, a Label,
'and a Command Button, load a type library based on Text1,
'show types in List1 and members in List2.  The PrototypeMember
'function is defined in the linked topic.
Option Explicit
Private m_TLInf As TypeLibInfo

Private Sub Command2_Click()
    cdlgArch.ShowOpen
    
    If cdlgArch.Flags = 0 Then
        End
    End If
    
    Text1.Text = cdlgArch.FileName
    Command1_Click
End Sub

Private Sub Form_Load()
  Set m_TLInf = New TypeLibInfo
  m_TLInf.AppObjString = "<Unqualified>"
  Inicializar
End Sub
Private Sub Command1_Click()
  On Error Resume Next
  m_TLInf.ContainingFile = Text1
  If Err Then Beep: Exit Sub
  List2.Clear
  With List1
    .Clear
     m_TLInf.GetTypesDirect .hWnd
    If .ListCount Then .ListIndex = 0
  End With
End Sub

Private Sub List1_Click()
  With List1
    List2.Clear
    txtDocumentation.Text = ""
    'Retrieve the SearchData from the ItemData property
    m_TLInf.GetMembersDirect .ItemData(.ListIndex), List2.hWnd, tliWtListBox, tliIdtInvokeKinds
  End With
End Sub
Private Sub List2_Click()
Dim InvKinds As TLI.InvokeKinds
    
    With List2
        InvKinds = .ItemData(.ListIndex)
        txtDocumentation.Text = PrototypeMember(m_TLInf, _
                                 List1.ItemData(List1.ListIndex), _
                                 InvKinds, , .[_Default])
    End With
End Sub


Sub Inicializar()
Dim strFilter  As String
    strFilter = "DLL Files (*.dll)|*.dll|"
    strFilter = strFilter & "OCX Files (*.ocx)|*.ocx|"
    strFilter = strFilter & "All Files (*.*)|*.*"
    cdlgArch.Filter = strFilter
End Sub


