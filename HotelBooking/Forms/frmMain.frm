VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "The Sports Front Hotel"
   ClientHeight    =   8310
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   1920
      Top             =   1560
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1376
      ButtonWidth     =   1535
      ButtonHeight    =   1376
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ImageList2"
      HotImageList    =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Rm Status"
            Key             =   "Status"
            Object.ToolTipText     =   "Room Status"
            ImageIndex      =   12
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Single"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Double"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reserve"
            Key             =   "Reserve"
            Object.ToolTipText     =   "Reserve Room"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Check-In"
            Key             =   "In"
            Object.ToolTipText     =   "Check-In"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Confirm"
            Key             =   "Confirm"
            Object.ToolTipText     =   "Confirm Reservation"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Check-Out"
            Key             =   "Out"
            Object.ToolTipText     =   "Check-Out"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Clients"
            Key             =   "Clients"
            Object.ToolTipText     =   "Clients Master List"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Rooms"
            Key             =   "Rooms"
            Object.ToolTipText     =   "Room Details"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "User"
            Key             =   "Settings"
            Object.ToolTipText     =   "User Settings"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Help"
            Key             =   "Help"
            Object.ToolTipText     =   "Help Topics and About"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Key             =   "Exit"
            Object.ToolTipText     =   "Exit Program"
            ImageIndex      =   6
         EndProperty
      EndProperty
      Begin VB.PictureBox picDate 
         BorderStyle     =   0  'None
         Height          =   600
         Left            =   9360
         ScaleHeight     =   600
         ScaleWidth      =   2415
         TabIndex        =   1
         Top             =   90
         Width           =   2415
         Begin VB.Label lblDate 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Current Date"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   195
            Left            =   720
            TabIndex        =   5
            Top             =   0
            Width           =   1605
         End
         Begin VB.Label lblTime 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Current Time"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   195
            Left            =   720
            TabIndex        =   4
            Top             =   360
            Width           =   1620
         End
         Begin VB.Label lblDateTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
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
            Left            =   120
            TabIndex        =   3
            Top             =   0
            Width           =   405
         End
         Begin VB.Label lblTimeTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Time"
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
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Width           =   420
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   0
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A82
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":235E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2C3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3516
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3DF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":46CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4FAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5E84
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6D5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7C38
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8B12
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":99EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A8C6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mTransaction 
      Caption         =   "&Transaction"
      Begin VB.Menu mReserve 
         Caption         =   "&Reserve"
      End
      Begin VB.Menu mCheckIn 
         Caption         =   "Check-&In"
      End
      Begin VB.Menu mConfirm 
         Caption         =   "Confir&m Reservation"
      End
      Begin VB.Menu mCheckOut 
         Caption         =   "Check-&Out"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu rStat 
         Caption         =   "Room Status"
      End
      Begin VB.Menu rInv 
         Caption         =   "Room Inventory"
      End
   End
   Begin VB.Menu mFile 
      Caption         =   "&File"
      Begin VB.Menu mRoom 
         Caption         =   "&Rooms"
      End
      Begin VB.Menu mClients 
         Caption         =   "&Clients"
      End
      Begin VB.Menu mUser 
         Caption         =   "&User Settings"
      End
   End
   Begin VB.Menu mReports 
      Caption         =   "&Reports"
      Begin VB.Menu mRepRegisterGuest 
         Caption         =   "List of Registered Guest"
      End
      Begin VB.Menu mRepReserve 
         Caption         =   "List of Reservations"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mRepClient 
         Caption         =   "List of Clients"
      End
   End
   Begin VB.Menu mLogOff 
      Caption         =   "&Log Off"
   End
   Begin VB.Menu mQuit 
      Caption         =   "&Quit"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mCheckIn_Click()
    Load frmCheckInOut
    frmCheckInOut.dtpIn.Value = Now
    frmCheckInOut.dtpIn.MaxDate = Now
    frmCheckInOut.dtpOut.Value = DateAdd("d", 1, Now)
    frmCheckInOut.dispIn
    frmCheckInOut.Show vbModal
End Sub

Private Sub mCheckOut_Click()
    Load frmCheckInOut
    frmCheckInOut.Caption = "Check-Out"
    frmCheckInOut.lblTitle.Caption = "CHECK-OUT"
    frmCheckInOut.dtpIn.Value = Now
    frmCheckInOut.dtpIn.Enabled = False
    frmCheckInOut.dtpOut.MaxDate = Now
    frmCheckInOut.dtpOut.Value = Now
    frmCheckInOut.dispOut
    frmCheckInOut.Show vbModal
End Sub

Private Sub mClients_Click()
    On Error Resume Next
    frmClientList.Show
    frmClientList.WindowState = vbMaximized
    frmClientList.SetFocus
    On Error GoTo 0
End Sub

Private Sub mConfirm_Click()
    Load frmCheckInOut
    frmCheckInOut.Caption = "Confirm Reservation"
    frmCheckInOut.lblTitle.Caption = "CONFIRM RESERVATION"
    frmCheckInOut.dtpIn.Value = Now
    frmCheckInOut.dtpIn.MaxDate = Now
    frmCheckInOut.dtpOut.Value = DateAdd("d", 1, Now)
    frmCheckInOut.dispConfirm
    frmCheckInOut.Show vbModal
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        If vbNo = MsgBox("You are about to quit this application. Are you sure?", vbQuestion + vbYesNo + vbDefaultButton2) Then
            Cancel = 1
            Exit Sub
        Else
            End
        End If
    Else
        Unload Me
    End If
End Sub

Private Sub MDIForm_Resize()
    On Error Resume Next
    Me.picDate.Left = Me.ScaleWidth - Me.picDate.Width - 100
    On Error GoTo 0
End Sub

Private Sub mLogOff_Click()
    If vbNo = MsgBox("Are you sure you want to Log Off?", vbQuestion + vbYesNo + vbDefaultButton2) Then Exit Sub
    MDIForm_QueryUnload 0, 1
    main
End Sub

Private Sub mQuit_Click()
    MDIForm_QueryUnload 0, 0
End Sub

Private Sub mRepClient_Click()
    On Error Resume Next
    repClient.Show
    repClient.WindowState = vbMaximized
    repClient.SetFocus
    On Error GoTo 0
End Sub

Private Sub mRepRegisterGuest_Click()
    On Error Resume Next
    repRegGuest.Show
    repRegGuest.WindowState = vbMaximized
    repRegGuest.SetFocus
    On Error GoTo 0
End Sub

Private Sub mRepReserve_Click()
    On Error Resume Next
    repReserve.Show
    repReserve.WindowState = vbMaximized
    repReserve.SetFocus
    On Error GoTo 0
End Sub

Private Sub mReserve_Click()
    Load frmCheckInOut
    frmCheckInOut.Caption = "Reserve"
    frmCheckInOut.lblTitle.Caption = "RESERVE"
    frmCheckInOut.dtpIn.Value = Now
    frmCheckInOut.dtpIn.MinDate = Now
    frmCheckInOut.dtpOut.Value = Now
    frmCheckInOut.dtpOut.MinDate = Now
    frmCheckInOut.dispIn
    frmCheckInOut.Show vbModal
End Sub

Private Sub mRoom_Click()
    On Error Resume Next
    frmRoomList.Show
    frmRoomList.WindowState = vbMaximized
    frmRoomList.SetFocus
    On Error GoTo 0
End Sub

Private Sub mUser_Click()
    On Error Resume Next
    frmUserList.Show
    frmUserList.WindowState = vbMaximized
    frmUserList.SetFocus
End Sub

Private Sub rInv_Click()
    On Error Resume Next
    frmRoomInventory.Show
    frmRoomInventory.WindowState = vbMaximized
    frmRoomInventory.SetFocus
    On Error GoTo 0
End Sub

Private Sub rStat_Click()
    On Error Resume Next
        frmroomstatus.Show
        frmroomstatus.WindowState = vbMaximized
        frmroomstatus.SetFocus
    On Error GoTo 0
End Sub

Private Sub Timer2_Timer()
lblDate.Caption = Format(Date, "dd MMM yyyy")
lblTime.Caption = Format(Time, "hh:dd:ss am/pm")
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case UCase(Button.Key)
    Case "STATUS"
        rStat_Click
    Case "RESERVE"
        mReserve_Click
    Case "IN"
        mCheckIn_Click
    Case "OUT"
        mCheckOut_Click
    Case "CONFIRM"
        mConfirm_Click
    Case "CLIENTS"
        mClients_Click
    Case "ROOMS"
        mRoom_Click
    Case "SETTINGS"
        mUser_Click
    Case "EXIT"
        MDIForm_QueryUnload 0, 0
End Select
End Sub
