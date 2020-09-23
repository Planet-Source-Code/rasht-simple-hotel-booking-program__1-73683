VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmClientList 
   Caption         =   "Client List"
   ClientHeight    =   7380
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7380
   ScaleWidth      =   10905
   Begin MSComctlLib.ImageList imgIcon 
      Left            =   315
      Top             =   6300
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
            Picture         =   "frmClientList.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientList.frx":0EDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientList.frx":1DB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientList.frx":2C8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientList.frx":3B68
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientList.frx":4A42
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
      Width           =   10905
      _ExtentX        =   19235
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
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\My Documents\caLLMe\fccT2\dbHotel97.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      EOFAction       =   2  'Add New
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1575
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   $"frmClientList.frx":591C
      Top             =   6525
      Visible         =   0   'False
      Width           =   1860
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmClientList.frx":59FE
      Height          =   5370
      Left            =   135
      OleObjectBlob   =   "frmClientList.frx":5A12
      TabIndex        =   0
      Top             =   765
      Width           =   10365
   End
End
Attribute VB_Name = "frmClientList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    'List Clients
    Data1.DatabaseName = dbPath
    Data1.RecordSource = "SELECT Client.cusName AS Name, Client.cusAdd AS Address, Client.cusCont AS [Contact Number], Client.cusNational AS Nationality, Client.cusCountry AS Country, Client.cusCompany AS Company FROM Client ORDER BY Client.cusName"
End Sub

Private Sub Form_Resize()
    'Resize dbGrid
    Me.DBGrid1.Width = Me.ScaleWidth - (Me.DBGrid1.Left * 2)
    Me.DBGrid1.Height = Me.ScaleHeight - Me.DBGrid1.Top - 100
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case UCase(Button.ToolTipText)
    Case "REFRESH"
        Data1.DatabaseName = dbPath
        Data1.RecordSource = "SELECT Client.cusName AS Name, Client.cusAdd AS Address, Client.cusCont AS [Contact Number], Client.cusNational AS Nationality, Client.cusCountry AS Country, Client.cusCompany AS Company FROM WHERE cusID =0 Client ORDER BY Client.cusName"
        Me.Data1.Refresh
        
        With DBGrid1
          For i = 0 To Me.Data1.Recordset.RecordCount - 1
            For j = 0 To 14
              .Row = i
              .Col = j
              .Text = ""
            Next j
          Next i
        End With
        
        Me.DBGrid1.ReBind
        Me.DBGrid1.Refresh
    Case "CLOSE"
        Unload Me
    End Select
End Sub
