Attribute VB_Name = "Database_SQL"
Option Compare Database
Option Explicit
'LIST OF SQL STRING WOULD BE USE FOR DB.EXECUTE
'------------------------------------------------------------------------------------
'UI SUPPORT SQL
Public Function SQL_CurrentDB_UITable_Deactivate(ByVal IDNumber As String, _
                                                 ByVal IDFieldName As String, _
                                                 ByVal UITable As String) As String
    Dim strSQL As String
    strSQL = " UPDATE " & UITable & _
             "  SET IsActive = FALSE " & _
             "  WHERE " & IDFieldName & " = " & IDNumber & " ;"
    
    SQL_CurrentDB_UITable_Deactivate = strSQL
End Function

Public Function SQL_CurrentDB_UITable_UpdateChangeDate(ByVal IDNumber As String, _
                                                       ByVal IDFieldName As String, _
                                                       ByVal UITable As String) As String
    Dim strSQL As String
    strSQL = " UPDATE " & UITable & _
             "  SET ChangeDate = Now() " & _
             "  WHERE " & IDFieldName & " = " & IDNumber & " ;"
    
    SQL_CurrentDB_UITable_UpdateChangeDate = strSQL
End Function



'------------------------------------------------------------------------------------
'CURRENTDB_TableScheme: ADD COLLUM
Public Function SQL_CurrentDB_AddIsActive(ByVal TableName As String) As String
    Dim strSQL As String
    strSQL = " ALTER TABLE " & TableName & _
             " ADD COLUMN IsActive YESNO DEFAULT '-1'; "
    
    SQL_CurrentDB_AddIsActive = strSQL
End Function
Public Function SQL_CurrentDB_AddInsertDate(ByVal TableName As String) As String
    Dim strSQL As String
    strSQL = " ALTER TABLE " & TableName & _
             "     ADD COLUMN InsertDate DATETIME;"
    
    SQL_CurrentDB_AddInsertDate = strSQL
End Function
Public Function SQL_CurrentDB_AddChangeDate(ByVal TableName As String) As String
    Dim strSQL As String
    strSQL = " ALTER TABLE " & TableName & _
             "     ADD COLUMN ChangeDate DATETIME;"
    
    SQL_CurrentDB_AddChangeDate = strSQL
End Function
'------------------------------------------------------------------------------------
'TEMPDB TABLE STRUCTURES AND FEATURES
Public Function SQL_TempDatabase_DropTable(ByVal TableName As String) As String
    Dim strSQL As String
    strSQL = " DROP TABLE " & TableName & ";"
    
    SQL_TempDatabase_DropTable = strSQL
End Function

Public Function SQL_TempDatabase_AddBeforeHash(ByVal TableName As String) As String
    Dim strSQL As String
    strSQL = " ALTER TABLE " & TableName & _
             "     ADD COLUMN BeforeHash CHAR(64);"
    
    SQL_TempDatabase_AddBeforeHash = strSQL
End Function

Public Function SQL_TempDatabase_AddAfterHash(ByVal TableName As String) As String
    Dim strSQL As String
    strSQL = " ALTER TABLE " & TableName & _
             "     ADD COLUMN AfterHash CHAR(64); "
    
    SQL_TempDatabase_AddAfterHash = strSQL
End Function

'FEATURES
Public Function SQL_TempDatabase_SelectIntoFromQuery(ByVal TempTableName As String, _
                                                     ByVal QueryName As String) As String
'GENERATE DATA SOURCE
    Dim TargetDBMSPath As String
    TargetDBMSPath = AccessTempDatabase.TempDatabase.Name
    
    Dim strSQL As String
    strSQL = " SELECT * " & _
             " INTO " & TempTableName & " IN " & """" & TargetDBMSPath & """" & _
             " FROM " & QueryName
    
    SQL_TempDatabase_SelectIntoFromQuery = strSQL

End Function

'CALCULATE HASH
Public Function SQL_CurrentDB_UpdateBeforeHash(ByVal TableName As String) As String
    
    'FILTER OUT UNWANTED ROW
    Dim FieldSCollection As VBA.Collection
    Set FieldSCollection = New VBA.Collection        'CONTAINER OF ONLY SELECTED OBJECTS
    
    Dim db As DAO.Database
    Set db = CurrentDb
    
    Dim fld As DAO.Field
    For Each fld In db.TableDefs(TableName).Fields
        Select Case fld.Name
            Case "AfterHash", "BeforeHash"
            Case "ChangeDate", "InsertDate"  ',"IsActive"
            Case "RowNum"
            Case Else
                Call FieldSCollection.Add(fld)
        End Select
    Next fld
    
    'BUILD SQL STRING
    Dim HashString As String
    HashString = BuildHashStringFromFields(FieldSCollection)
    
    Dim strSQL As String
    strSQL = " UPDATE " & TableName & _
             "   SET BeforeHash = SHA256_CNG(" & HashString & ") ; "
             
    SQL_CurrentDB_UpdateBeforeHash = strSQL
End Function
    
Public Function SQL_CurrentDB_UpdateAfterHash(ByVal TableName As String) As String
'CALCULATE HASH

    'FILTER OUT UNWANTED ROW
    Dim FieldSCollection As VBA.Collection
    Set FieldSCollection = New VBA.Collection        'CONTAINER OF ONLY SELECTED OBJECTS
    
    Dim db As DAO.Database
    Set db = CurrentDb
    
    Dim fld As DAO.Field
    For Each fld In db.TableDefs(TableName).Fields
        Select Case fld.Name
            Case "ChangeDate", "InsertDate" ',"IsActive"
            Case "AfterHash", "BeforeHash"
            Case "RowNum"
            Case Else
                Call FieldSCollection.Add(fld)
        End Select
    Next fld
    
    'BUILD SQL STRING
    Dim HashString As String
    HashString = BuildHashStringFromFields(FieldSCollection)
    
    Dim strSQL As String
    strSQL = " UPDATE " & TableName & _
             "   SET AfterHash = SHA256_CNG(" & HashString & ") ; "
             
    SQL_CurrentDB_UpdateAfterHash = strSQL
End Function
                                       
'--------------------------------------------------------------------------------------
'SQL - DATA MANIPULATING OPERATIONS
Public Function SQL_CurrentDB_UpdateTable(ByVal TargetUpdateTable As String, _
                                                ByVal DataSourceTable As String, _
                                                ByVal FieldName As String, _
                                                ByVal IDFieldName As String) As String
'UPDATE ONLY CHANGES VALUES
    'EXPLAIN CONDITION:
    '   *BEFOREHASH <> AFTERHASH                            = Select Rows has changed value
    '   *Nz(Table1.FieldName,"") <> Nz(Table2.FieldName,"") = Select Fields has changed value
    '   *NULL is filter with out with comparision
    Dim strSQL As String
    strSQL = " UPDATE " & TargetUpdateTable & _
         "  INNER JOIN " & DataSourceTable & _
         "  ON " & DataSourceTable & "." & IDFieldName & " = " & TargetUpdateTable & "." & IDFieldName & _
         " SET " & TargetUpdateTable & "." & FieldName & " = " & DataSourceTable & "." & FieldName & _
         " WHERE BeforeHash <> AfterHash " & _
         " AND  Nz(" & TargetUpdateTable & "." & FieldName & ","""")" & " <> Nz(" & DataSourceTable & "." & FieldName & ","""")" & _
         " ;"
    
    SQL_CurrentDB_UpdateTable = strSQL
End Function

Public Function SQL_CurrentDB_InsertIntoTable(ByVal TargetUpdateTable As String, _
                                                    ByVal DataSourceTable As String, _
                                                    ByVal ListOfFields As String, _
                                                    ByVal ListOfValues) As String

'INSERT ALL NEW ROWS
    'EXPLAIN CONDITION:
    '   * BEFOREHASH is null -> New Insert Data
    
    Dim strSQL As String
    strSQL = " INSERT INTO " & TargetUpdateTable & "(" & ListOfFields & ")" & _
             "  SELECT " & ListOfValues & "" & _
             "  FROM " & DataSourceTable & _
             "  WHERE BeforeHash IS NULL ;"
    
    SQL_CurrentDB_InsertIntoTable = strSQL
End Function

Public Function SQL_CurrentDB_InsertParentTable(ByVal TargetUpdateTable As String, _
                                                      ByVal DataSourceTable As String, _
                                                      ByVal ListOfFields As String, _
                                                      ByVal ListOfValues As String) As String
    
    Dim strSQL As String
    strSQL = " INSERT INTO " & TargetUpdateTable & "(" & ListOfFields & ")" & _
             "  SELECT " & ListOfValues & "" & _
             "  FROM " & DataSourceTable & _
             "  WHERE BeforeHash <> AfterHash OR BeforeHash IS NULL ;"
    
    SQL_CurrentDB_InsertParentTable = strSQL
End Function

'Private Function SQL_CurrentDB_UpdateIsActive(ByVal TargetUpdateTable As String, _
'                                                    ByVal DataSourceTable As String, _
'                                                    ByVal IDFieldName As String) As String
''UPDATE ONLY CHANGES VALUES
'    'EXPLAIN CONDITION:
'    '   *BEFOREHASH <> AFTERHASH                            = Select Rows has changed value
'    '   *Nz(Table1.FieldName,"") <> Nz(Table2.FieldName,"") = Select Fields has changed value
'    '   *NULL is filter with out with comparision
'    Dim strSQL As String
'    strSQL = " UPDATE " & TargetUpdateTable & _
'         "  INNER JOIN " & DataSourceTable & _
'         "  ON " & DataSourceTable & "." & IDFieldName & " = " & TargetUpdateTable & "." & IDFieldName & _
'         " SET " & TargetUpdateTable & "." & "IsActive" & " = " & DataSourceTable & "." & "IsActive" & _
'         " WHERE " & TargetUpdateTable & "." & "IsActive" & " <> " & DataSourceTable & "." & "IsActive"
'
'    SQL_CurrentDB_UpdateIsActive = strSQL
'End Function

'FORM SPECIFIC
Public Function SQL_CurrentDB_frmProductList_UpdateImagePath(ByVal IDNumber As String, _
                                                             ByVal Value As String, _
                                                             ByVal UITable As String) As String
'    Dim strSQL As String
'    strSQL = " UPDATE " & UITable & _
'             "  SET ImagePath = Now() " & _
'             "  WHERE " & IDFieldName & " = " & IDNumber & " ;"
             
    Dim strSQL As String
    strSQL = " UPDATE " & UITable & _
             "  SET ImagePath = " & """" & Value & """" & _
             "  WHERE ProductID " & " = " & IDNumber & " ;"
    
    SQL_CurrentDB_frmProductList_UpdateImagePath = strSQL
End Function



