VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRoomInventory 
   Caption         =   "Room Daily Inventory"
   ClientHeight    =   8145
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12315
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8145
   ScaleWidth      =   12315
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView lvw 
      Height          =   7035
      Left            =   90
      TabIndex        =   0
      Top             =   765
      Width           =   12075
      _ExtentX        =   21299
      _ExtentY        =   12409
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList imgIcon 
      Left            =   315
      Top             =   7740
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRoomInventory.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRoomInventory.frx":0EDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRoomInventory.frx":1DB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRoomInventory.frx":2C8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRoomInventory.frx":3B68
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRoomInventory.frx":4A42
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12315
      _ExtentX        =   21722
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgIcon"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Close"
            ImageIndex      =   6
         EndProperty
      EndProperty
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   1710
         ScaleHeight     =   555
         ScaleWidth      =   5460
         TabIndex        =   2
         Top             =   45
         Width           =   5460
         Begin MSComCtl2.DTPicker dtpMonth 
            Height          =   420
            Left            =   2340
            TabIndex        =   4
            Top             =   90
            Width           =   2445
            _ExtentX        =   4313
            _ExtentY        =   741
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "MMMM yyyy"
            Format          =   16515075
            CurrentDate     =   40500
         End
         Begin VB.Label lblDate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "For the Month of:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   180
            TabIndex        =   3
            Top             =   135
            Width           =   2115
         End
      End
   End
End
Attribute VB_Name = "frmRoomInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim invRs As Recordset, rmInRs As Recordset
Dim itm As ListItem

Private Sub dtpMonth_CloseUp()
    Dim inDtS As Date, inDtE As Date, tiCol%, lvCtr%
    
    'gets the beggining date of the month
    inDtS = DateSerial(Year(Me.dtpMonth.Value), Month(Me.dtpMonth.Value), 1)
    
    'gets the end date of the month
    inDtE = DateSerial(Year(Me.dtpMonth.Value), Month(Me.dtpMonth.Value) + 1, 1) - 1
    
    'total days
    tiCol = DateDiff("d", inDtS, inDtE)
    
    'List rooms and room inventory per day
    Me.lvw.ListItems.Clear
    Me.lvw.ColumnHeaders.Clear
    With Me.lvw.ColumnHeaders
        .Add , , "RM NO", 1700
        .Add , , "ROOM TYPE", 2500
        .Add , , "RATE", 1400, vbCenter
        .Add , , "OCCUPIED", 1200, vbCenter
        .Add , , "RESERVE", 1200, vbCenter
        .Add , , "VACANT", 1200, vbCenter
        
        For lvCtr = 1 To tiCol + 1
            .Add , , Format(DateAdd("d", (lvCtr - 1), inDtS), "d"), 550, vbCenter
        Next
    
    End With
    
    openRS rmInRs, "SELECT * FROM Rooms ORDER BY rmName", dbOpenSnapshot
    While Not rmInRs.EOF
        Set itm = Me.lvw.ListItems.Add(, , rmInRs!rmName)
        itm.SubItems(1) = rmInRs!rmType
        itm.SubItems(2) = Format(rmInRs!rmRate, "P #,##0.00")
        
        For lvCtr = 6 To tiCol + 6
            openRS invRs, "SELECT * FROM Accomodation WHERE rmID =" & rmInRs!rmID & " AND #" & Format(DateAdd("d", lvCtr - 6, inDtS), "MM/dd/yyyy") & "# BETWEEN DATEVALUE(FORMAT(accDtIN,'MM/dd/yyyy')) AND DATEVALUE(FORMAT(accDtOut,'MM/dd/yyyy'))", dbOpenSnapshot
            If invRs.RecordCount > 0 Then
                Select Case invRs!accStatus
                Case "IN", "OUT"
                    itm.SubItems(3) = Val(itm.SubItems(3)) + 1
                    itm.SubItems(lvCtr) = "O"
                Case "RESERVE"
                    itm.SubItems(4) = Val(itm.SubItems(4)) + 1
                    itm.SubItems(lvCtr) = "R"
                End Select
                itm.SubItems(lvCtr) = IIf(invRs!accStatus = "IN" Or invRs!accStatus = "OUT", "O", IIf(invRs!accStatus = "RESERVE", "R", ""))
            Else
                itm.SubItems(5) = Val(itm.SubItems(5)) + 1
            End If
            closeRs invRs
        Next
    
        rmInRs.MoveNext
    Wend
    closeRs rmInRs
    
'    Me.lblDate.Caption = "For the Month of: " & Format(Now, "MMMM yyyy")
End Sub

Private Sub Form_Load()
    Me.dtpMonth.Value = Now
    Me.dtpMonth.MaxDate = Now
    dtpMonth_CloseUp
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Me.lvw.Height = Me.ScaleHeight - Me.lvw.Top - 100
    Me.lvw.Width = Me.ScaleWidth - (Me.lvw.Left * 2)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.ToolTipText
    Case "Refresh"
        Form_Load
    Case "Close"
        Unload Me
    End Select
End Sub
