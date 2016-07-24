VERSION 5.00
Begin VB.Form frmaccmanage 
   Caption         =   "Account Manager"
   ClientHeight    =   1905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   ScaleHeight     =   1905
   ScaleWidth      =   5220
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmpass 
      Caption         =   "Password Form"
      Height          =   1695
      Left            =   3240
      TabIndex        =   9
      Top             =   120
      Width           =   1935
      Begin VB.TextBox txtnewpass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtpass1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label lblnewpass 
         Caption         =   "Enter new password:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblpass1 
         Caption         =   "Re-Type Password:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   960
         Width           =   1575
      End
   End
   Begin VB.TextBox txtdomain 
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Top             =   330
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.OptionButton optacctype 
      Caption         =   "Domain Account"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
   Begin VB.OptionButton optacctype 
      Caption         =   "Local Account"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox txtusername 
      Height          =   285
      Left            =   1080
      TabIndex        =   7
      Top             =   1080
      Width           =   1935
   End
   Begin VB.ComboBox cbofunction 
      Height          =   315
      ItemData        =   "frmaccmanage.frx":0000
      Left            =   120
      List            =   "frmaccmanage.frx":0002
      TabIndex        =   3
      Text            =   "Choose an Option"
      Top             =   720
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "COMMIT CHANGE"
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label lbldomain 
      Caption         =   "Enter Domain:"
      Height          =   255
      Left            =   1800
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblusername 
      Caption         =   "User Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1080
      Width           =   975
   End
End
Attribute VB_Name = "frmaccmanage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim comtcommand As Double
Dim comstring As String

Private Sub cbofunction_Click()
Refresh
If cbofunction.Text = "Change Password" Then
frmaccmanage.Width = 5500
Else
frmaccmanage.Width = 3300
End If
End Sub


Private Sub Command1_Click()
RunFunction
End Sub

Private Sub Form_Load()
frmaccmanage.Width = 3300

MsgBox ("This utility will work only if you are logged in with admin priviledges")

With cbofunction
    .AddItem ("Disable Account")
    .AddItem ("Enable Account")
    .AddItem ("Change Password")
End With
End Sub

Private Sub optacctype_Click(Index As Integer)
If optacctype(1) Then
lbldomain.Visible = True
txtdomain.Visible = True
Else
lbldomain.Visible = False
txtdomain.Visible = False
End If
End Sub
Private Sub RunFunction()

If cbofunction.Text <> "Choose an Option" Then
    comstring = "net user "
    If txtusername.Text <> "" Then
        comstring = comstring & txtusername.Text
    Else
        MsgBox ("You did not enter a username.")
        GoTo endcmd
    End If

    If cbofunction.Text = "Change Password" Then
        If txtnewpass.Text <> "" And txtpass1.Text <> "" Then
            If txtnewpass.Text = txtpass1.Text Then
                comstring = comstring & " " & txtnewpass.Text
            Else
                MsgBox ("Passwords do not match")
                GoTo endcmd
            End If
        Else
            MsgBox ("Check password fields one of both may be empty")
        End If
    End If
    
    If optacctype(1) Then
        comstring = comstring & " /" & txtdomain.Text
            If txtdomain.Text = "" Then
                MsgBox ("You did not enter a Domain.")
                GoTo endcmd
            End If
    End If
    
    If cbofunction.Text = "Disable Account" Then
        comstring = comstring & " /active:no "
    End If
    
    If cbofunction.Text = "Enable Account" Then
        comstring = comstring & " /active:yes "
    End If
    
    If MsgBox("Are you sure you want to commit this change?", vbYesNo, "WARNING") = vbYes Then
        comtcommand = Shell(comstring, vbHide)
        MsgBox ("Option committed.")
        Call writetofile
    Else
        MsgBox ("You have cancelled the action.")
        GoTo endcmd
    End If
    
Else
    MsgBox ("You must first choose an option")
End If

endcmd:
End Sub

Private Sub writetofile()
Dim fsize As Long
Dim fso As FileSystemObject
Dim changefile As File

On Error GoTo filenotfound

fsize = 0

fsize = FileLen(App.Path & "\changes.txt")

filenotfound:

Updated = Now() & vbTab & cbofunction.Text & vbTab & txtusername.Text & vbTab
If optacctype(1) Then
    Updated = Updated & txtdomain.Text
Else
    Updated = Updated & "localhost"
End If

If fsize = 0 Then
    Const forwriting = 2
    Set fso = New FileSystemObject
    fso.CreateTextFile App.Path & "\changes.txt", True
    Set changefile = fso.GetFile(App.Path & "\changes.txt")
    
    Open App.Path & "\changes.txt" For Append As #2
    Print #2, Updated
    Close #2
Else
    Open App.Path & "\changes.txt" For Append As #2
    Print #2, Updated
    Close #2
End If
End Sub
