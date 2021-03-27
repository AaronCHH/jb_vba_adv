Attribute VB_Name = "Module1"
Option Explicit


'-----------------------------------------------------------
'範例107
'在Excel工作表中列出Access的資料表
'-----------------------------------------------------------

Sub ADOSample1()
    Dim myConnect As New ADODB.Connection
    Dim myRcdSet As New ADODB.Recordset
    Dim myProvider As String, mySource As String
    
    myProvider = "Provider=Microsoft.Jet.OLEDB.4.0;"
    mySource = "Data Source=C:\Excel2003VBA應用篇\會員管理.mdb;"
    
    myConnect.Open myProvider & mySource

    myRcdSet.Open "T_會員名單", myConnect
    
    Range("A6").CopyFromRecordset myRcdSet

    myRcdSet.Close
    myConnect.Close
    
    Set myRcdSet = Nothing
    Set myConnect = Nothing
End Sub


'-----------------------------------------------------------
'範例108
'連接資料連結檔案
'-----------------------------------------------------------

Sub ADOSample2()
    Dim myConnect As New ADODB.Connection
    Dim myRcdSet As New ADODB.Recordset
    
    myConnect.Open "File Name=C:\Excel2003VBA應用篇\Test.udl;"

    myRcdSet.Open "T_會員名單", myConnect
    
    Range("A6").CopyFromRecordset myRcdSet

    myRcdSet.Close
    myConnect.Close
    
    Set myRcdSet = Nothing
    Set myConnect = Nothing
End Sub


'-----------------------------------------------------------
'範例109
'使用SQL篩選資料
'-----------------------------------------------------------

Sub ADOSample3()
    Dim myConnect As New ADODB.Connection
    Dim myRcdSet As New ADODB.Recordset
    
    myConnect.Open "File Name=C:\Excel2003VBA應用篇\Test.udl;"

    With myRcdSet
        .ActiveConnection = myConnect
        .Source = "SELECT * FROM T_申請 WHERE 課程NO='C001'"
        .Open
    End With
    
    Range("A6").CopyFromRecordset myRcdSet

    myRcdSet.Close
    myConnect.Close
    
    Set myRcdSet = Nothing
    Set myConnect = Nothing
End Sub
