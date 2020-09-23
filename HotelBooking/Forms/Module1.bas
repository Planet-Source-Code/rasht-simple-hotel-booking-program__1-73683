Attribute VB_Name = "Module1"
Option Explicit
Public db As Database
Public rs As Recordset
Public dbPath As String
Public usrLogLevel As String
Public usrLogName As String

Sub main()
    dbPath = App.Path & "\dbHotel97.mdb"
    Set db = OpenDatabase(dbPath)
    On Error Resume Next
    devReports.Connection1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\dbHotel97.mdb;Persist Security Info=False"
    On Error GoTo 0
    frmLogIn.Show
End Sub

Public Sub openRS(dbRs As Recordset, strSql As String, Optional opType As RecordsetTypeEnum = dbOpenDynaset)

    Set dbRs = db.OpenRecordset(strSql, opType)

End Sub

Public Sub closeRs(ByVal dbRs As Recordset)
    On Error Resume Next
        dbRs.Close
        Set dbRs = Nothing
    On Error GoTo 0
End Sub

