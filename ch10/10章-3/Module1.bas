Attribute VB_Name = "Module1"
Option Explicit


'-----------------------------------------------------------
'�d��107
'�bExcel�u�@���C�XAccess����ƪ�
'-----------------------------------------------------------

Sub ADOSample1()
    Dim myConnect As New ADODB.Connection
    Dim myRcdSet As New ADODB.Recordset
    Dim myProvider As String, mySource As String
    
    myProvider = "Provider=Microsoft.Jet.OLEDB.4.0;"
    mySource = "Data Source=C:\Excel2003VBA���νg\�|���޲z.mdb;"
    
    myConnect.Open myProvider & mySource

    myRcdSet.Open "T_�|���W��", myConnect
    
    Range("A6").CopyFromRecordset myRcdSet

    myRcdSet.Close
    myConnect.Close
    
    Set myRcdSet = Nothing
    Set myConnect = Nothing
End Sub


'-----------------------------------------------------------
'�d��108
'�s����Ƴs���ɮ�
'-----------------------------------------------------------

Sub ADOSample2()
    Dim myConnect As New ADODB.Connection
    Dim myRcdSet As New ADODB.Recordset
    
    myConnect.Open "File Name=C:\Excel2003VBA���νg\Test.udl;"

    myRcdSet.Open "T_�|���W��", myConnect
    
    Range("A6").CopyFromRecordset myRcdSet

    myRcdSet.Close
    myConnect.Close
    
    Set myRcdSet = Nothing
    Set myConnect = Nothing
End Sub


'-----------------------------------------------------------
'�d��109
'�ϥ�SQL�z����
'-----------------------------------------------------------

Sub ADOSample3()
    Dim myConnect As New ADODB.Connection
    Dim myRcdSet As New ADODB.Recordset
    
    myConnect.Open "File Name=C:\Excel2003VBA���νg\Test.udl;"

    With myRcdSet
        .ActiveConnection = myConnect
        .Source = "SELECT * FROM T_�ӽ� WHERE �ҵ{NO='C001'"
        .Open
    End With
    
    Range("A6").CopyFromRecordset myRcdSet

    myRcdSet.Close
    myConnect.Close
    
    Set myRcdSet = Nothing
    Set myConnect = Nothing
End Sub
