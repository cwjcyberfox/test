Sub TestImportFileFunction()

Dim filepath As String
filepath = "C:\Users\vaio\Downloads\201408190820424849\导出CSV示例\测试表.csv"
Call importFile("test", filepath)

End Sub


Private Function importFile(ByVal tblName As String, FileFullPath As String, Optional FieldDelimiter As String = ",", Optional RecordDelimiter As String = vbCrLf) As Boolean
    
    On Error GoTo BACKERR:
 
    Dim objDB As DAO.Workspace
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim iFileNum As Integer
    Dim sFileContents As String         ' csv文件数据
    Dim sTableSplit() As String         ' csv以行为单位的数据组
    Dim sRecordSplit() As String        ' csv每行内以(,)分隔开后的数据组
    Dim lCtr As Integer
    Dim iCtr As Integer
    Dim lRecordCount As Long            ' csv文件内的行数
    Dim tmp As String
    Dim sql As String
    
    ' csv文件数据的取得
    iFileNum = FreeFile
    Open FileFullPath For Binary As #iFileNum
    sFileContents = Space(LOF(iFileNum))
    Get #iFileNum, , sFileContents
    Close #iFileNum
    
    sTableSplit = Split(sFileContents, RecordDelimiter)
    ' 每行数据的项目数取得
    lRecordCount = UBound(sTableSplit)
    ' 对应项目名取得
    sRecordFields = Split(sTableSplit(0), FieldDelimiter)
     
    Set objDB = DBEngine.Workspaces(0)
    Set db = CurrentDb
    
    Call RenameTable("test2", "test")
     
    objDB.BeginTrans
    For lCtr = 1 To lRecordCount - 1
        ' 每一行的数据取得
        tmp = Replace(CStr(sTableSplit(lCtr)), Chr(34), "")
       ' 去除双引号(")
        sRecordSplit = Split(tmp, FieldDelimiter)
         
        'Set rs = db.OpenRecordset(tblName, dbOpenDynaset)
        'rs.FindFirst "ID= '1-" & sRecordSplit(12) & "-" & sRecordSplit(13) & "'"
        
        sql = "INSERT INTO test([编号],[姓名],[性别],[年龄]) VALUES('" & sRecordSplit(0) & "','" & sRecordSplit(1) & "','" & sRecordSplit(2) & "','" & sRecordSplit(3) & "')"
        
        CurrentDb.Execute (sql)
                         
        'If rs.NoMatch = False Then
        '    rs.Edit
        '    rs!modified = Now
        '    rs!modifieduser = Comm.getUserId
        '    rs.Update
        'End If
    Next lCtr
     
   objDB.CommitTrans
     
   Call RenameTable("test", "test2")
     
   'rs.Close
   db.Close
   objDB.Close
   Set rs = Nothing
   Set db = Nothing
   Set objDB = Nothing
   importFile = True
   Exit Function
   
BACKERR:
   objDB.Rollback
   If Not rs Is Nothing Then
        rs.Close
   End If
   db.Close
   objDB.Close
   Set rs = Nothing
   Set db = Nothing
   Set objDB = Nothing
   importFile = False
     
    MsgBox Err.Number & "->" & Err.Description
End Function
    
Public Function RenameTable(ByVal targetTableName As String, ByVal newTableName)
    Dim tbl As TableDef
    Dim dbs As Database
    Dim nbl As String
    Dim obl As String
    nbl = "Newtable"
    Set dbs = CurrentDb
    For Each tbl In dbs.TableDefs
        obl = tbl.Name
        If obl = targetTableName Then
            tbl.Name = newTableName
        MsgBox "修改成功", vbInformation, "提示"
        End If
      
    Next
End Function
