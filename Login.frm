VERSION 5.00
Begin VB.Form frmLogin 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Identification"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   Icon            =   "Login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Abandonner"
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connexion"
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox motdepass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   "waechter"
      Top             =   840
      Width           =   2535
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   300
      Width           =   2535
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Mot de passe"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   960
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nom d'utilisateur"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1155
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Noyau.AquerirEmployes
    If (Combo1.ListIndex <> -1 And motdepass.Text <> "") Then
        Noyau.Login Combo1.List(Combo1.ListIndex), motdepass.Text
    Else
        MsgBox "Informattions invalides", vbCritical + vbOKOnly, Me.Caption
    End If
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
Noyau.AquerirEmployes
Employes.Show
End Sub

Private Sub Form_Load()
On Error GoTo Oups
    Dim g_connData As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim fld As ADODB.Field, alignment As Integer
    Dim recCount As Long, i As Long, fldName As String
    g_connData.Open "Driver={SQL Server};Server=TOUR-PATRICE\SQLEXPRESS;Database=WebGRB;Trusted_Connection=Yes;"
    rs.Open "GrbEmploye", g_connData, adOpenForwardOnly, adLockReadOnly
    Combo1.Clear
    rs.MoveFirst
    Do Until rs.EOF
        recCount = recCount + 1
        Combo1.AddItem rs.Fields("employe")
        If recCount = MaxRecords Then Exit Do
        rs.MoveNext
    Loop
    Combo1.ListIndex = recount - 1

    Exit Sub
Oups:
    MsgBox Err.Description + vbCrLf + Err.Source, vbCritical, Me.Caption
End Sub
