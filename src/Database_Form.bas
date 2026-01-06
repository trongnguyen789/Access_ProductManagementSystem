Attribute VB_Name = "Database_Form"
Option Compare Database
Option Explicit

'COMMON FEATURES
Public Function RefreshData(ByVal TableName_TempDB As String, ByVal Query_GenerateDataSource As String)
    Call SpeedCheck("RefreshData")
    'DROP TABLE ON TEMP TABLE
    If AccessTempDatabase.TableExistsInTempDB(TableName_TempDB) Then
        Call AccessTempDatabase.TempDatabase.Execute(SQL_TempDatabase_DropTable(TableName_TempDB))
    End If
    
    'GENERATE DATA SOURCE
    Call CurrentDb.Execute(SQL_TempDatabase_SelectIntoFromQuery(TableName_TempDB, Query_GenerateDataSource))
    
    'MAKE TABLE STRUCTURE CHANGES
    Call AccessTempDatabase.TempDatabase.Execute(SQL_TempDatabase_AddBeforeHash(TableName_TempDB))
    Call AccessTempDatabase.TempDatabase.Execute(SQL_TempDatabase_AddAfterHash(TableName_TempDB))
    AccessTempDatabase.RelinkTables
    
    'MAKE ISACTIVE DEFAULT
    AccessTempDatabase.TempDatabase.TableDefs(TableName_TempDB).Fields("IsActive").DefaultValue = "True"
    'DEBUG
    Call SpeedCheck("RefreshData")
    Call DEBUG_ProcedureResult("RefreshData", "Success")
End Function


Public Function TrackChanges(ByRef DataChanged As Boolean, _
                             ByRef NewData As Boolean, _
                             ByVal Table_LinkedTable As String) As Boolean
'DEFINED DATA CHANGED BY CALCULATE HASH-BEFORE AND HASH-AFTER
    Call SpeedCheck("TrackChangesAndNotify")

    'CALCULATE AFTER SNAPSHOT
    Call CurrentDb.Execute(SQL_CurrentDB_UpdateAfterHash(Table_LinkedTable))
    
    'GET STATE -> WRITE DEBUG FOR THIS
    DataChanged = IsDataChanged(Table_LinkedTable)
    NewData = IsThereNewData(Table_LinkedTable)

    If DataChanged Or NewData Then TrackChanges = True
    'DEBUGING
    Call SpeedCheck("TrackChanges")
    Call DEBUG_ProcedureResult("TrackChanges", "Success")
End Function

Public Function DataChangeNotify() As Integer
    'CONFIRM TO MAKE CHANGES
    Call SpeedCheck("DataChangeNotify")
    Dim UserConfirm As Integer
    UserConfirm = MsgBox("  Do you want to save edited data?    ", vbYesNo, _
                         " THERE IS NEW DATA OR DATA IS CHANGED !!! ")
    DataChangeNotify = UserConfirm
    
    'DEBUGING
    Call SpeedCheck("DataChangeNotify")
    Call DEBUG_ProcedureResult("DataChangeNotify", "Success")
End Function








