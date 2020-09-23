VERSION 5.00
Begin VB.Form frmLogIn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System Security"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   Icon            =   "frmSystemSecurity.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   4695
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   4455
      Begin VB.TextBox txtUsername 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1560
         TabIndex        =   0
         Text            =   "admin"
         Top             =   360
         Width           =   2580
      End
      Begin VB.TextBox txtPassword 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   1
         Text            =   "admin"
         Top             =   720
         Width           =   2580
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2880
         TabIndex        =   3
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdLogIn 
         Appearance      =   0  'Flat
         Caption         =   "&Log In"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1560
         TabIndex        =   2
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Username"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   810
      End
   End
End
Attribute VB_Name = "frmLogIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsUser As Recordset

Private Sub cmdCancel_Click()
    End
End Sub

Private Sub cmdLogIn_Click()
    'checks if username is blank then exit this codes/sub procedure
    If Trim(Me.txtUsername.Text) = "" Then
        MsgBox "Blank User Name!", , "Login"
        Exit Sub
    End If

    'checks if username exist
    openRS rsUser, "SELECT * FROM User WHERE usrName =" & Chr$(34) & Trim(Me.txtUsername.Text) & Chr$(34), dbOpenSnapshot
    
    If Not rsUser.EOF Then 'if username is found
        'check for correct password
        If Me.txtPassword.Text = rsUser!usrPassWord Then
            usrLogName = rsUser!usrFullNm
            usrLogLevel = rsUser!usrLevel
            frmMain.Show
            Unload Me
        Else
            MsgBox "Invalid Password, try again!", , "Login"
            Me.txtPassword.SetFocus
            SendKeys "{Home}+{End}"
        End If
        
    Else    'if username does not exist
            MsgBox "Unauthorized User Name!", , "Login"
            Me.txtUsername.SetFocus
            SendKeys "{Home}+{End}"
    End If

End Sub
