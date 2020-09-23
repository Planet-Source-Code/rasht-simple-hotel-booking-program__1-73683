VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmRoomStatus 
   Caption         =   "Room Status"
   ClientHeight    =   6990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11670
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6990
   ScaleWidth      =   11670
   Begin VB.ComboBox cmbView 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "frmRoomStatus.frx":0000
      Left            =   135
      List            =   "frmRoomStatus.frx":0007
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   180
      Width           =   3390
   End
   Begin VB.ComboBox cmbStatus 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "frmRoomStatus.frx":0016
      Left            =   3600
      List            =   "frmRoomStatus.frx":002C
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   180
      Width           =   2940
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\My Documents\caLLMe\fccT2\dbHotel97.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      EOFAction       =   1  'EOF
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1575
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   $"frmRoomStatus.frx":0062
      Top             =   6525
      Visible         =   0   'False
      Width           =   1860
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmRoomStatus.frx":0470
      Height          =   5145
      Left            =   180
      OleObjectBlob   =   "frmRoomStatus.frx":0484
      TabIndex        =   0
      Top             =   810
      Width           =   11265
   End
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
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRoomStatus.frx":25E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRoomStatus.frx":34C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRoomStatus.frx":439B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRoomStatus.frx":5275
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRoomStatus.frx":614F
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRoomStatus.frx":7029
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRoomStatus.frx":7F03
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   630
      Left            =   6615
      TabIndex        =   1
      Top             =   45
      Width           =   3120
      _ExtentX        =   5503
      _ExtentY        =   1111
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      ImageList       =   "imgIcon"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modify"
            Object.Tag             =   "Modify"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Delete Accomodation"
            Object.Tag             =   "Delete"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Refresh"
            Object.Tag             =   "Refresh"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Print Summary of Bills"
            Object.Tag             =   "Print"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Close"
            Object.Tag             =   "Close"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmroomstatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rmRs As Recordset

Private Sub cmbStatus_Click()
Dim i%, j%, strSel$

'fields to be displayed in dbGrid
strSel = "SELECT Accomodation.accID AS [ID], Accomodation.accRFNo AS [RF Number], Rooms.rmName AS [Room No], Rooms.rmType AS [Room Type], Client.cusName AS Guest, " & _
        "IIf(IsNull([accID]),'',DateValue(Format([accDtIn],'MMM-d-yyyy'))) AS Arrival, IIf(IsNull([accID]),'',DateValue(Format([accDtOut],'MMM-d-yyyy'))) AS Departure, " & _
        "IIf(IsNull([accStatus]),'AVAILABLE',[accStatus]) AS Status, IIf(IsNull([accID]),0,DateDiff('d',[accDtIn],[accDtOut])+1) AS Days, IIf(IsNull([accID]),'',Format([accDtRecord],'MMM-d-yyyy')) AS [Date Recorded], " & _
        "IIf(IsNull([accID]),0,[accHeadA]+[accHeadC]) AS [Head Count], IIf(IsNull([accID]),0,[accRate]) AS Rate, IIf(IsNull([accID]),0,[accCharge]) AS [Other Charges], " & _
        "IIf(IsNull([accID]),0,([accRate]*(DateDiff('d',[accDtIn],[accDtOut])+1))+[accCharge]) AS [Total Bill], IIf(IsNull([accID]),'',[accFOIn]) AS [Check-In By], IIf(IsNull([accID]),'',[accFOOut]) AS [Check-Out By] " & _
        "FROM (Accomodation LEFT JOIN Client ON Accomodation.cusID = Client.cusID) RIGHT JOIN Rooms ON Accomodation.rmID = Rooms.rmID "

'clear dbGrid
Me.Data1.DatabaseName = dbPath
Me.Data1.RecordSource = strSel & " WHERE Accomodation.accID =0"
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


Select Case Me.cmbView.Text
Case "ALL ROOMS"
    Select Case Me.cmbStatus.Text
    Case "CURRENT STATUS"   'List all rooms and their current status
        Data1.RecordSource = strSel & " WHERE accStatus = 'IN' OR DATEVALUE(FORMAT(NOW, 'MM/dd/yyyy')) BETWEEN DATEVALUE(IIf(IsNull([accID]),Now(),FORMAT(accDtIn,'MM/dd/yyyy'))) AND DATEVALUE(IIf(IsNull([accID]),Now(),FORMAT(accDtOut, 'MM/dd/yyyy')))" & _
                                        " OR(NOT ISNULL(accID) AND DATEVALUE(FORMAT(NOW, 'MM/dd/yyyy')) <> DATEVALUE(IIf(IsNull([accID]),Now(),FORMAT(accDtIn,'MM/dd/yyyy'))) AND DATEVALUE(FORMAT(NOW, 'MM/dd/yyyy')) <> DATEVALUE(IIf(IsNull([accID]),Now(),FORMAT(accDtOut, 'MM/dd/yyyy'))) AND accStatus <> 'OUT') " & _
                                        " ORDER BY Rooms.rmName, DateValue(Format(IIf(IsNull([accID]),Now(),[accDtIn]),'MMM-d-yyyy'))"
    Case "AVAILABLE"        'List available rooms only
        Data1.RecordSource = strSel & " WHERE ( ISNULL(accID) AND DATEVALUE(FORMAT(NOW, 'MM/dd/yyyy')) BETWEEN DATEVALUE(IIf(IsNull([accID]),Now(),FORMAT(accDtIn,'MM/dd/yyyy'))) AND DATEVALUE(IIf(IsNull([accID]),Now(),FORMAT(accDtOut, 'MM/dd/yyyy'))))" & _
                                        " ORDER BY Rooms.rmName, DateValue(Format(IIf(IsNull([accID]),Now(),[accDtIn]),'MMM-d-yyyy'))"
    
    Case "ALL"              'List all accomodation record
        Data1.RecordSource = strSel & " ORDER BY Rooms.rmName, DateValue(Format(IIf(IsNull([accID]),Now(),[accDtIn]),'MMM-d-yyyy'))"
    Case Else               'List rooms by status selected in status combobox
        Data1.RecordSource = strSel & " WHERE accStatus = " & Chr$(34) & Me.cmbStatus.Text & Chr$(34) & " ORDER BY Rooms.rmName, DateValue(Format([accDtIn],'MMM-d-yyyy'))"
    End Select
Case Else
    Select Case Me.cmbStatus.Text
    Case "CURRENT STATUS"   'List specified rooms and its current status
        Data1.RecordSource = strSel & " WHERE (rmName =" & Chr$(34) & Me.cmbView.Text & Chr$(34) & " AND DATEVALUE(FORMAT(NOW, 'MM/dd/yyyy')) BETWEEN DATEVALUE(IIf(IsNull([accID]),Now(),FORMAT(accDtIn,'MM/dd/yyyy'))) AND DATEVALUE(IIf(IsNull([accID]),Now(),FORMAT(accDtOut, 'MM/dd/yyyy')))) AND (rmName =" & Chr$(34) & Me.cmbView.Text & Chr$(34) & " OR accStatus = 'IN') ORDER BY Rooms.rmName, DateValue(Format(IIf(IsNull([accID]),Now(),[accDtIn]),'MMM-d-yyyy'))"
    Case "AVAILABLE"        'List if room is available
        Data1.RecordSource = strSel & " WHERE (rmName =" & Chr$(34) & Me.cmbView.Text & Chr$(34) & " AND ISNULL(accID) AND DATEVALUE(FORMAT(NOW, 'MM/dd/yyyy')) BETWEEN DATEVALUE(IIf(IsNull([accID]),Now(),FORMAT(accDtIn,'MM/dd/yyyy'))) AND DATEVALUE(IIf(IsNull([accID]),Now(),FORMAT(accDtOut, 'MM/dd/yyyy'))))" & _
                                        " ORDER BY Rooms.rmName, DateValue(Format(IIf(IsNull([accID]),Now(),[accDtIn]),'MMM-d-yyyy'))"
    Case "ALL"              'List all accomodation record of specified room
        Data1.RecordSource = strSel & " WHERE rmName = " & Chr$(34) & Me.cmbView.Text & Chr$(34) & " ORDER BY DateValue(Format(IIf(IsNull([accID]),Now(),[accDtIn]),'MMM-d-yyyy'))"
    Case Else               'List the room by status selected in status combobox
        Data1.RecordSource = strSel & " WHERE rmName = " & Chr$(34) & Me.cmbView.Text & Chr$(34) & " AND accStatus = " & Chr$(34) & Me.cmbStatus.Text & Chr$(34) & " ORDER BY Rooms.rmName, DateValue(Format([accDtIn],'MMM-d-yyyy'))"
    End Select
End Select

'HAVING PROBLEM IN REFRESHING LIST IN DBGRID ..
'this is my first time to use dbGrid, and my last resort to have easy codes..
Me.Data1.Refresh
Me.DBGrid1.ReBind
Me.DBGrid1.Refresh


End Sub

Private Sub cmbView_Click()
    If Me.cmbStatus.Text = "" Then
        Me.cmbStatus.Text = "CURRENT STATUS"
    Else
        cmbStatus_Click
    End If
End Sub


Private Sub Form_Load()
    
    'List rooms in combobox
    openRS rmRs, "SELECT rmName FROM Rooms ORDER BY rmName", dbOpenSnapshot
    Me.cmbView.Clear
    Me.cmbView.AddItem "ALL ROOMS"
    While Not rmRs.EOF
        Me.cmbView.AddItem rmRs!rmName
        rmRs.MoveNext
    Wend
    closeRs rmRs
    
    Me.cmbView.Text = "ALL ROOMS"
    
End Sub

Private Sub Form_Resize()
    Me.DBGrid1.Width = Me.ScaleWidth - (Me.DBGrid1.Left * 2)
    Me.DBGrid1.Height = Me.ScaleHeight - Me.DBGrid1.Top - 100
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case UCase(Button.Tag)
    Case "MODIFY"
        'Modifies accomodation details
        If Me.DBGrid1.Columns(0) = "" Or Me.DBGrid1.Columns(7) <> "IN" Then Exit Sub
        
        Load frmCheckInOut
        frmCheckInOut.lblTitle.Caption = "MODIFY"
        frmCheckInOut.dispOut
        frmCheckInOut.Tag = Me.DBGrid1.Columns(0)
        frmCheckInOut.cmbRoom.Text = Me.DBGrid1.Columns(2)
        frmCheckInOut.dispMod
        frmCheckInOut.Show vbModal
        
        cmbStatus_Click
        cmbStatus_Click
    Case "DELETE"
        'Deletes accomodation
        If Me.DBGrid1.Columns(0) = "" Then Exit Sub
        
        If vbNo = MsgBox("Are you sure you want to delete this accomodation record?", vbYesNo + vbDefaultButton2 + vbQuestion) Then Exit Sub
        
        db.Execute "DELETE * FROM Accomodation WHERE accID =" & Me.DBGrid1.Columns(0)
        
        cmbStatus_Click
        
    Case "REFRESH"
        cmbStatus_Click
    Case "PRINT"
        'Prints bill of selected accomodations
        If Me.DBGrid1.Columns(0) = "" Then Exit Sub
        
        If devReports.rsBill.State = adStateOpen Then
            devReports.rsBill.Close
        End If
        
        On Error GoTo errMess
        devReports.rsBill.Open "SELECT Format(Now(),'MMM-dd-yyyy') AS [DATE], Client.cusName AS GUEST, Client.cusAdd AS ADDRESS, Client.cusCont AS [CONTACT NO], Client.cusCompany AS COMPANY, Format([accDtIn],'MMM-dd-yyyy') AS ARRIVAL, Format([accDtOut],'MMM-dd-yyyy') AS DEPARTURE, Accomodation.accRFNo AS [RF NO], Rooms.rmName AS ROOM, Format([accRate],'P #,##0.00') AS RATE, DateDiff('d',[accDtIn],[accDtOut])+1 AS [LENGTH OF STAY], Format(((DateDiff('d',[accDtIn],[accDtOut])+1)*[accRate]),'P #,##0.00') AS ACCOMODATION, Format(((DateDiff('d',[accDtIn],[accDtOut])+1)*[accRate])+[accCharge],'P #,##0.00') AS BALANCE, Format([accCharge],'P #,##0.00') AS CHARGES FROM (Accomodation LEFT JOIN Client ON Accomodation.cusID = Client.cusID) LEFT JOIN Rooms ON Accomodation.rmID = Rooms.rmID WHERE accID =" & Val(Me.DBGrid1.Columns(0))
        
        repBills.Show
        repBills.WindowState = vbMaximized
        repBills.SetFocus
        On Error GoTo 0
        
    Case "CLOSE"
        Unload Me
    End Select

Exit Sub
errMess:
MsgBox "An error occured. Please try again.", vbCritical
End Sub


