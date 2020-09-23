VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCheckInOut 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Check-In"
   ClientHeight    =   6975
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNationality 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4950
      TabIndex        =   8
      Top             =   3150
      Width           =   2580
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
      Left            =   6075
      TabIndex        =   17
      Top             =   5985
      Width           =   1455
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
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
      Left            =   4635
      TabIndex        =   16
      Top             =   5985
      Width           =   1455
   End
   Begin VB.TextBox txtChild 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2835
      TabIndex        =   12
      Top             =   4725
      Width           =   960
   End
   Begin VB.TextBox txtAdult 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1620
      TabIndex        =   11
      Text            =   "1"
      Top             =   4725
      Width           =   960
   End
   Begin MSMask.MaskEdBox mebCharges 
      Height          =   360
      Left            =   4950
      TabIndex        =   13
      Top             =   4680
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "P #,##0.00;(P #,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mebRate 
      Height          =   360
      Left            =   1620
      TabIndex        =   14
      Top             =   5220
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "P #,##0.00;(P #,##0.00)"
      PromptChar      =   "_"
   End
   Begin VB.ComboBox cmbFO 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1620
      TabIndex        =   10
      Top             =   4050
      Width           =   5910
   End
   Begin VB.TextBox txtContact 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1620
      TabIndex        =   7
      Top             =   3150
      Width           =   2220
   End
   Begin VB.ComboBox cmbCompany 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1620
      TabIndex        =   9
      Top             =   3600
      Width           =   5910
   End
   Begin VB.TextBox txtAddress 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1620
      TabIndex        =   6
      Top             =   2700
      Width           =   5910
   End
   Begin VB.ComboBox cmbGuest 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1620
      TabIndex        =   5
      Top             =   2250
      Width           =   5910
   End
   Begin VB.TextBox txtDays 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4185
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1305
      Width           =   690
   End
   Begin MSComCtl2.DTPicker dtpOut 
      Height          =   360
      Left            =   4950
      TabIndex        =   2
      Top             =   1305
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "MMMM d, yyyy"
      Format          =   16384003
      CurrentDate     =   40250
   End
   Begin MSComCtl2.DTPicker dtpIn 
      Height          =   360
      Left            =   1620
      TabIndex        =   0
      Top             =   1305
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "MMMM d, yyyy"
      Format          =   16384003
      CurrentDate     =   40250
   End
   Begin VB.TextBox txtRF 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4950
      TabIndex        =   4
      Top             =   1755
      Width           =   2580
   End
   Begin VB.ComboBox cmbRoom 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1620
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1755
      Width           =   2490
   End
   Begin MSMask.MaskEdBox mebTotal 
      Height          =   360
      Left            =   4950
      TabIndex        =   15
      Top             =   5175
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "P #,##0.00;(P #,##0.00)"
      PromptChar      =   "_"
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nationality:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4005
      TabIndex        =   35
      Top             =   3195
      Width           =   960
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4005
      TabIndex        =   34
      Top             =   5265
      Width           =   555
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rate:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   33
      Top             =   5310
      Width           =   465
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Charges:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4005
      TabIndex        =   32
      Top             =   4770
      Width           =   780
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Child"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2835
      TabIndex        =   31
      Top             =   4455
      Width           =   420
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Adult"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1620
      TabIndex        =   30
      Top             =   4455
      Width           =   435
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Head Count:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   29
      Top             =   4815
      Width           =   1065
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "CHECK-IN"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   315
      TabIndex        =   28
      Top             =   225
      Width           =   7305
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Frontdesk:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   27
      Top             =   4095
      Width           =   915
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   26
      Top             =   3645
      Width           =   870
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact No:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   25
      Top             =   3195
      Width           =   1005
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   24
      Top             =   2745
      Width           =   765
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Guest Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   23
      Top             =   2295
      Width           =   1110
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RF No.:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4230
      TabIndex        =   22
      Top             =   1800
      Width           =   645
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Room No:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   21
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Departure"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4950
      TabIndex        =   20
      Top             =   1035
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Arrival"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1620
      TabIndex        =   19
      Top             =   1035
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   18
      Top             =   1350
      Width           =   465
   End
End
Attribute VB_Name = "frmCheckInOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsGuest As Recordset, rsRoom As Recordset

Public Sub dispIn()
    dtpIn_CloseUp
End Sub

Public Sub dispOut()
    'Lists occupied rooms for Check-out
    openRS rsRoom, "SELECT Rooms.rmName FROM Accomodation LEFT JOIN Rooms ON Accomodation.rmID = Rooms.rmID WHERE accStatus ='IN' ORDER BY Rooms.rmName", dbOpenSnapshot
    Me.cmbRoom.Clear
    While Not rsRoom.EOF
        Me.cmbRoom.AddItem rsRoom!rmName
        rsRoom.MoveNext
    Wend
    closeRs rsRoom
End Sub

Public Sub dispConfirm()
    'Lists reserved rooms for confirm check-in
    openRS rsRoom, "SELECT Rooms.rmName FROM Accomodation LEFT JOIN Rooms ON Accomodation.rmID = Rooms.rmID WHERE accStatus ='RESERVE' AND DATEVALUE(FORMAT(NOW,'MM/dd/yyyy')) >= DATEVALUE(FORMAT(accDtIn,'MM/dd/yyyy')) ORDER BY Rooms.rmName", dbOpenSnapshot
    Me.cmbRoom.Clear
    While Not rsRoom.EOF
        Me.cmbRoom.AddItem rsRoom!rmName
        rsRoom.MoveNext
    Wend
    closeRs rsRoom
End Sub

Public Sub dispMod()
    'Prepares form for modification purpose
    Me.dtpIn.Enabled = False
    Me.dtpOut.Enabled = False
    Me.cmbRoom.Locked = True
    
    cmbRoom_Click
End Sub

Private Sub computeT()
    'Computes total days and total charges
    Me.txtDays.Text = DateDiff("d", Me.dtpIn.Value, Me.dtpOut.Value) + 1
    Me.mebTotal.Text = Val(Me.mebRate.Text) * (DateDiff("d", Me.dtpIn.Value, Me.dtpOut.Value) + 1) + Val(Me.mebCharges.Text)
End Sub

Private Sub cmbGuest_Click()
    cmbGuest_LostFocus
End Sub

Private Sub cmbGuest_LostFocus()
    'Displays customer's pre-recorded details
    openRS rsGuest, "SELECT * FROM Client WHERE cusName =" & Chr$(34) & Trim(Me.cmbGuest.Text) & Chr$(34), dbOpenSnapshot
    If rsGuest.RecordCount > 0 Then
        Me.txtAddress.Text = rsGuest!cusAdd
        Me.txtContact.Text = rsGuest!cusCont
        Me.txtNationality.Text = rsGuest!cusNational
        Me.cmbCompany.Text = rsGuest!cusCompany
    End If
    closeRs rsGuest
End Sub

Private Sub cmbRoom_Click()
    'Fill data in form except during check-in
    If UCase(Me.lblTitle.Caption) = "CHECK-OUT" Or UCase(Me.lblTitle.Caption) = "CONFIRM RESERVATION" Or UCase(Me.lblTitle.Caption) = "MODIFY" Then
        If UCase(Me.lblTitle.Caption) = "CONFIRM RESERVATION" Then
            openRS rsRoom, "SELECT Accomodation.*, Rooms.rmName, Client.cusName, Client.cusAdd, Client.cusCont, Client.cusNational, Client.cusCountry, Client.cusCompany FROM (Accomodation LEFT JOIN Rooms ON Accomodation.rmID = Rooms.rmID) INNER JOIN Client ON Accomodation.cusID = Client.cusID WHERE rmName =" & Chr$(34) & Me.cmbRoom.Text & Chr$(34) & " AND accStatus ='RESERVE'", dbOpenSnapshot
        Else
            openRS rsRoom, "SELECT Accomodation.*, Rooms.rmName, Client.cusName, Client.cusAdd, Client.cusCont, Client.cusNational, Client.cusCountry, Client.cusCompany FROM (Accomodation LEFT JOIN Rooms ON Accomodation.rmID = Rooms.rmID) INNER JOIN Client ON Accomodation.cusID = Client.cusID WHERE rmName =" & Chr$(34) & Me.cmbRoom.Text & Chr$(34) & " AND accStatus ='IN'", dbOpenSnapshot
        End If
        
        If Not rsRoom.EOF Then
            On Error Resume Next
            Me.Tag = rsRoom!accID
            Me.txtRF.Text = rsRoom!accRfNo
            Me.cmbGuest.Text = rsRoom!cusName
            Me.dtpIn.Value = rsRoom!accDtIn
            Me.dtpOut.Value = rsRoom!accDtOut
            Me.txtAddress.Text = rsRoom!cusAdd
            Me.txtContact.Text = rsRoom!cusCont
            Me.txtNationality.Text = rsRoom!cusNational
            Me.cmbCompany.Text = rsRoom!cusCompany
            Me.cmbFO.Text = rsRoom!accFOIn
            Me.txtAdult.Text = rsRoom!accHeadA
            Me.txtChild.Text = rsRoom!accHeadC
            Me.mebCharges.Text = rsRoom!accCharge
            Me.mebRate.Text = rsRoom!accRate
            
            On Error GoTo 0
            
            dtpIn_CloseUp
        
        End If
        
        closeRs rsRoom
    Else    'executes if check-in
        openRS rsRoom, "SELECT rmRate FROM Rooms WHERE rmName =" & Chr$(34) & Me.cmbRoom.Text & Chr$(34), dbOpenSnapshot
        Me.mebRate.Text = rsRoom!rmRate
        closeRs rsRoom
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDone_Click()
    Dim guestID As Long, roomID As Long
    Dim intArr As Long, intDepa As Long
    
    'Verify total days and must not be zero
    If Val(Me.txtDays.Text) <= 0 Then MsgBox "Invalid Dates.", vbCritical: Exit Sub
    
    'Room must not be blank
    If Trim(Me.cmbRoom.Text) = "" Then MsgBox "Invalid Room.", vbCritical: Exit Sub

    'Converts date to integer, its easier for me to query in tables than dates
    intArr = Format(Me.dtpIn.Value, "yyyyMMdd")
    intDepa = Format(Me.dtpOut.Value, "yyyyMMdd")
    
    'query for roomID (room's primary key)
    openRS rsRoom, "SELECT rmID FROM Rooms WHERE rmName =" & Chr$(34) & Me.cmbRoom.Text & Chr$(34), dbOpenSnapshot
    roomID = rsRoom!rmID
    closeRs rsRoom
    
    openRS rsGuest, "SELECT * FROM Accomodation WHERE accID <> " & Val(Me.Tag) & " AND rmID =" & roomID & _
                      " AND ((" & intArr & " <= VAL(FORMAT(accDtIn,'yyyyMMdd')) AND " & intDepa & " > VAL(FORMAT(accDtIn,'yyyyMMdd')))" & _
                      " OR  (" & intArr & " < VAL(FORMAT(accDtOut,'yyyyMMdd')) AND " & intDepa & " > VAL(FORMAT(accDtOut,'yyyyMMdd')))" & _
                      " OR  (" & intArr & " < VAL(FORMAT(accDtIn,'yyyyMMdd')) AND " & intDepa & " > VAL(FORMAT(accDtOut,'yyyyMMdd')))" & _
                      " OR  (" & intArr & " >= VAL(FORMAT(accDtIn,'yyyyMMdd')) AND " & intDepa & " <= VAL(FORMAT(accDtOut,'yyyyMMdd'))))" & _
                      " ORDER BY accDtIn DESC", dbOpenSnapshot
    
    If Not rsGuest.EOF Then MsgBox "Date conflicts with previous accomodation.", vbCritical: closeRs rsGuest: Exit Sub
    closeRs rsGuest
    
    'RF number must not be blank. RF Number is the reference # from the registration form where the guest fills-up
    If Me.lblTitle.Caption <> "RESERVE" Then If Trim(Me.txtRF.Text) = "" Then MsgBox "Invalid RF Number.", vbCritical: Exit Sub
    
    'RF number must not duplicate
    If Trim(Me.txtRF.Text) <> "" Then
        openRS rsRoom, "SELECT accRFNo FROM Accomodation WHERE accID <> " & Val(Me.Tag) & " AND accRFNo =" & Chr$(34) & Me.txtRF.Text & Chr$(34), dbOpenSnapshot
        If rsRoom.RecordCount > 0 Then MsgBox "RF Number already exist.", vbCritical: closeRs rsRoom: Exit Sub
        closeRs rsRoom
    End If
    
    'Guest name must not be blank
    If Trim(Me.cmbGuest.Text) = "" Then Exit Sub
    
    'Checks for guest count
    If Val(Me.txtAdult.Text) <= 0 Then MsgBox "Please specify number of Adult.", vbCritical: Exit Sub
    'Room rate must not be zero or less
    If Val(Me.mebRate.Text) <= 0 Then MsgBox "Invalid Room Rate.", vbCritical: Exit Sub
    
    'records new guest in Client table
    openRS rsGuest, "SELECT * FROM Client WHERE cusName =" & Chr$(34) & Trim(Me.cmbGuest.Text) & Chr$(34)
    If Not rsGuest.EOF Then
        guestID = rsGuest!cusID
    Else
        rsGuest.AddNew
        guestID = rsGuest!cusID
        rsGuest!cusName = UCase(Trim(Me.cmbGuest.Text))
        rsGuest!cusAdd = Trim(Me.txtAddress.Text)
        rsGuest!cusCont = Trim(Me.txtContact.Text)
        rsGuest!cusNational = Trim(Me.txtNationality.Text)
        rsGuest!cusCompany = UCase(Trim(Me.cmbCompany.Text))
        rsGuest.Update
    End If
    closeRs rsGuest
    
    'Records accomodation details
    openRS rsRoom, "SELECT * FROM Accomodation WHERE accID =" & Val(Me.Tag)
    Select Case UCase(Me.lblTitle.Caption)
    Case "CHECK-IN"
        rsRoom.AddNew
        rsRoom!accStatus = "IN"
        rsRoom!accFOIn = Me.cmbFO.Text
    Case "CHECK-OUT"
        rsRoom.Edit
        rsRoom!accStatus = "OUT"
        rsRoom!accFOOut = Me.cmbFO.Text
    Case "RESERVE"
        rsRoom.AddNew
        rsRoom!accStatus = "RESERVE"
        rsRoom!accFOIn = Me.cmbFO.Text
    Case "CONFIRM RESERVATION"
        rsRoom.Edit
        rsRoom!accStatus = "IN"
        rsRoom!accFOIn = Me.cmbFO.Text
    Case "MODIFY"
        rsRoom.Edit
        rsRoom!accFOIn = Me.cmbFO.Text
    End Select

    rsRoom!cusID = guestID
    rsRoom!rmID = roomID
    rsRoom!accRfNo = Trim(Me.txtRF.Text)
    rsRoom!accDtIn = Me.dtpIn.Value
    rsRoom!accDtOut = Me.dtpOut.Value
    rsRoom!accHeadA = Val(Me.txtAdult.Text)
    rsRoom!accHeadC = Val(Me.txtChild.Text)
    rsRoom!accRate = Val(Me.mebRate.Text)
    rsRoom!accCharge = Val(Me.mebCharges.Text)
    rsRoom.Update
    
    closeRs rsRoom
    
    DoEvents
    Unload Me
End Sub

Private Sub dtpIn_CloseUp()
    computeT
    If Not (UCase(Me.lblTitle.Caption) = "CHECK-OUT" Or UCase(Me.lblTitle.Caption) = "CONFIRM RESERVATION" Or UCase(Me.lblTitle.Caption) = "MODIFY") Then
'    Else
        'FOR RESERVATION: Lists available room for reservation
        Dim intArr As Long, intDepa As Long
        
        intArr = Format(Me.dtpIn.Value, "yyyyMMdd")
        intDepa = Format(Me.dtpOut.Value, "yyyyMMdd")
        
        openRS rsRoom, "SELECT rmID, rmName FROM Rooms ORDER BY rmName", dbOpenSnapshot
        Me.cmbRoom.Clear
        While Not rsRoom.EOF
    
            openRS rsGuest, "SELECT * FROM Accomodation WHERE (accStatus = 'IN' AND rmID =" & rsRoom!rmID & ") OR (accID <> " & Val(Me.Tag) & " AND rmID =" & rsRoom!rmID & _
                              " AND ((" & intArr & " <= VAL(FORMAT(accDtIn,'yyyyMMdd')) AND " & intDepa & " > VAL(FORMAT(accDtIn,'yyyyMMdd')))" & _
                              " OR  (" & intArr & " < VAL(FORMAT(accDtOut,'yyyyMMdd')) AND " & intDepa & " > VAL(FORMAT(accDtOut,'yyyyMMdd')))" & _
                              " OR  (" & intArr & " < VAL(FORMAT(accDtIn,'yyyyMMdd')) AND " & intDepa & " > VAL(FORMAT(accDtOut,'yyyyMMdd')))" & _
                              " OR  (" & intArr & " >= VAL(FORMAT(accDtIn,'yyyyMMdd')) AND " & intDepa & " <= VAL(FORMAT(accDtOut,'yyyyMMdd'))))" & _
                              ") ORDER BY accDtIn DESC", dbOpenSnapshot
            
            If rsGuest.EOF Then Me.cmbRoom.AddItem rsRoom!rmName
            closeRs rsGuest
            
            rsRoom.MoveNext
        Wend
        closeRs rsRoom
    End If
    
End Sub

Private Sub dtpOut_CloseUp()
    dtpIn_CloseUp
End Sub

Private Sub Form_Load()
    openRS rsGuest, "SELECT cusName FROM Client ORDER BY cusName", dbOpenSnapshot
    Me.cmbGuest.Clear
    While Not rsGuest.EOF
        Me.cmbGuest.AddItem rsGuest!cusName
        rsGuest.MoveNext
    Wend
    closeRs rsGuest
    
    openRS rsGuest, "SELECT DISTINCT cusCompany FROM Client ORDER BY cusCompany", dbOpenSnapshot
    Me.cmbCompany.Clear
    While Not rsGuest.EOF
        Me.cmbCompany.AddItem rsGuest!cusCompany
        rsGuest.MoveNext
    Wend
    closeRs rsGuest
    
    openRS rsGuest, "SELECT DISTINCT usrName FROM User ORDER BY usrName", dbOpenSnapshot
    Me.cmbFO.Clear
    While Not rsGuest.EOF
        Me.cmbFO.AddItem rsGuest!usrName
        rsGuest.MoveNext
    Wend
    closeRs rsGuest
    
    Me.cmbFO.Text = usrLogName
End Sub

Private Sub mebCharges_Change()
    computeT
End Sub

Private Sub mebRate_Change()
    computeT
End Sub
