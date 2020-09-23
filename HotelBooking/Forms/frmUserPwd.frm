VERSION 5.00
Begin VB.Form frmUserPwd 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Change Password"
   ClientHeight    =   3225
   ClientLeft      =   2760
   ClientTop       =   3630
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAddSave 
      Caption         =   "&Change"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2430
      TabIndex        =   3
      Top             =   2355
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3645
      TabIndex        =   4
      Top             =   2355
      Width           =   1215
   End
   Begin VB.TextBox txtConPass 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   1890
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1515
      Width           =   2805
   End
   Begin VB.TextBox txtNewPass 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   1890
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   930
      Width           =   2805
   End
   Begin VB.TextBox txtOldPass 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   1890
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   360
      Width           =   2805
   End
   Begin VB.Label lblConPass 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   315
      TabIndex        =   7
      Top             =   1605
      Width           =   780
   End
   Begin VB.Label lblNewPass 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Password:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   315
      TabIndex        =   6
      Top             =   1020
      Width           =   1305
   End
   Begin VB.Label lblOldPass 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Old Password:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   315
      TabIndex        =   5
      Top             =   450
      Width           =   1230
   End
End
Attribute VB_Name = "frmUserPwd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsUserP As Recordset    'declares variable for user table

Private Sub cmdAddSave_Click()
    'this search the user by its usrID tagged in txtOldPass
    Set rsUserP = db.OpenRecordset("SELECT * FROM User WHERE usrID =" & Val(Me.Tag))
    
    'checks if current password does not match to the one typed in Old Password
    If rsUserP!usrPassWord <> Me.txtOldPass.Text Then
        MsgBox "Incorrect password.", vbCritical
        Me.txtOldPass.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
    
    'checks if New Password is the same with Confirm
    If Me.txtNewPass.Text <> Me.txtConPass.Text Then
        MsgBox "Confirm Password does not match with New Password.", vbCritical
        Me.txtConPass.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
    
    'changes the password
    db.Execute "UPDATE User SET usrPassWord =" & Chr$(34) & Trim(Me.txtNewPass.Text) & Chr$(34) & " WHERE usrID =" & Val(Me.Tag)
    
    'prompts a success message
    MsgBox "Password successfully changed.", vbInformation
    
    Unload Me
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

