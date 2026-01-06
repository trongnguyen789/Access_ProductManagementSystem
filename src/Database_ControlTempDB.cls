VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Database_ControlTempDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'
Private TempDB As DAO.Database
Private TempDBPath As String
Private TempConnection As ADODB.Connection
Private m_ControlTempDB As Database_ControlTempDB

'====================================================================
' PROPER PROPERTIES
'====================================================================

Public Property Get TempDatabase() As DAO.Database
    Set TempDatabase = TempDB
End Property


'====================================================================
' CLASS EVENTS
'====================================================================
Private Sub Class_Initialize()
    
    'SET THE TEMPDB PATH
    TempDBPath = CurrentProject.Path & "\TempDatabase.accdb"
    
    'KILL THE FILE
    If IsFileExist(TempDBPath, vbHidden) Then
        Call SetAttr(TempDBPath, vbNormal)
        Call Kill(TempDBPath)
    End If
    
    'CREATE DATABASE / CONNECTION
    Call GenerateTempDatabase(TempDB)
    Call ClearUnkownLinkedTable
    
    'SET IT HIDDEN
    Call SetAttr(TempDBPath, vbHidden)
    
End Sub

Private Sub Class_Terminate()
    
    ' KILL THE TEMP TABLE
    If IsFileExist(TempDBPath, vbHidden) Then
        Call SetAttr(TempDBPath, vbNormal)
        Set TempDB = Nothing
        Call Kill(TempDBPath)
    End If
    
End Sub

'====================================================================
' PUBLIC METHODS
'====================================================================

Public Sub RelinkTables()
    ''REFRESH TABLEDEFS
    Dim MyDB As DAO.Database:        Set MyDB = CurrentDb:     Call MyDB.TableDefs.Refresh
    Dim TemporaryDB As DAO.Database: Set TemporaryDB = TempDB: Call TemporaryDB.TableDefs.Refresh

    'LOOP TABLEDEF
    Dim curdb_tbdef As DAO.TableDef, curdb_tbdef_name As String
    
    Dim TableDefObject As TableDef
    For Each TableDefObject In TemporaryDB.TableDefs
        Select Case True
            Case TableDefObject.Name Like "MSys*"
            Case TableDefObject.Name Like "~*"
            Case Else
                
                'BUILD LINKED TABLE NAME
                Dim LinkedTableName As String
                LinkedTableName = TableDefObject.Name & "_Linked"
                
                Dim TableDef_Linked As DAO.TableDef
                
                If IsTableDefExist(MyDB.TableDefs, LinkedTableName) Then
                'LINKED TABLE EXISTS?
                    Set TableDef_Linked = MyDB.TableDefs(LinkedTableName)
                
                    If TableDef_Linked.SourceTableName = TableDefObject.Name Then
                    'IF LINKED TABLE SOURCE IS CORRECT -> REFRESH
                        TableDef_Linked.Connect = ";DATABASE=" & TemporaryDB.Name
                        TableDef_Linked.RefreshLink
                    Else
                    'IF SOURCENAME IS WRONG -> RE-CREATE TABLEDEF + NEW SOURCE
                        Call MyDB.TableDefs.Delete(LinkedTableName)                 'DEL OBJECT
                        GoSub CreateLinkedTable:
                    End If
            
                Else
                'IF TABLE NO EXIST -> CREATE TABLEDEF
                    GoSub CreateLinkedTable:
                End If
        End Select
    Next TableDefObject
    
    'EXIT
    Call ClearUnkownLinkedTable
    RefreshDatabaseWindow
    Exit Sub

CreateLinkedTable:
    Set TableDef_Linked = MyDB.CreateTableDef(LinkedTableName)
    TableDef_Linked.SourceTableName = TableDefObject.Name
    TableDef_Linked.Connect = ";DATABASE=" & TemporaryDB.Name
    Call MyDB.TableDefs.Append(TableDef_Linked)
    Return
End Sub

Public Sub ClearUnkownLinkedTable()
    'REFRESH DB
    Dim MyDB As DAO.Database:       Set MyDB = CurrentDb:        Call MyDB.TableDefs.Refresh
    Dim TemporDB As DAO.Database:   Set TemporDB = TempDatabase: Call TemporDB.TableDefs.Refresh
      
    'CHECK CURRENT DB FOR TABLE WITH "_LINK" AND DELETE IT IN TEMPDB
    Dim TableDef_Linked As DAO.TableDef, TableDef_TemporDB As DAO.TableDef, TableName As String
    For Each TableDef_Linked In MyDB.TableDefs
        If TableDef_Linked.Name Like "*_Linked" Then
            'EXTRACT NAME
            TableName = Left(TableDef_Linked.Name, Len(TableDef_Linked.Name) - Len("_Linked"))
            'DELETE LINKED-TABLE IF NOT EXISTS
            If IsTableDefExist(TemporDB.TableDefs, TableName) = False Then
                Call MyDB.TableDefs.Delete(TableDef_Linked.Name)
            End If
        End If
    Next TableDef_Linked
    
    'EXIT DATABASE
    RefreshDatabaseWindow
End Sub


Public Function GenerateLinkedTable(ByVal TableName_TempDB As String) As DAO.TableDef


    'DELETE IF EXIST
    Dim TemporDB As DAO.Database: Set TemporDB = TempDB
    If IsTableDefExist(TemporDB.TableDefs, TableName_TempDB) Then
        Call TemporDB.TableDefs.Delete(TableName_TempDB)
    End If
    
    'CREATE TEMP TABLE
    Dim TableDefObject_TempDB As DAO.TableDef
    Set TableDefObject_TempDB = TemporDB.CreateTableDef(TableName_TempDB)
    
    Dim FieldObject As DAO.Field
    Set FieldObject = TableDefObject_TempDB.CreateField("ID", dbInteger)
    Call TableDefObject_TempDB.Fields.Append(FieldObject)
    Call TemporDB.TableDefs.Append(TableDefObject_TempDB)
    
    'REFERENCE CURRENT DATABASE
    Dim MyDB As DAO.Database: Set MyDB = CurrentDb

    'CREATE LINKED TABLE
    Dim TableDefObject_MyDB As DAO.TableDef
    Set TableDefObject_MyDB = MyDB.CreateTableDef(TableDefObject_TempDB.Name & "_Linked")
    
    With TableDefObject_MyDB
        .SourceTableName = TableDefObject_TempDB.Name
        .Connect = ";DATABASE=" & TempDBPath
    End With
    
    'DELETE OBJECT IF EXIST
    If IsTableDefExist(MyDB.TableDefs, TableDefObject_MyDB.Name) Then
        Call MyDB.TableDefs.Delete(TableDefObject_MyDB.Name)
    End If
    
    'CREATE NEW CURRENTDB OBJECT
    Call MyDB.TableDefs.Append(TableDefObject_MyDB)
    
    'EXIT
    RefreshDatabaseWindow
End Function



Public Sub DeleteLinkedTable(ByVal TableName_TempDB As String)
    'GET MyDB AND TemporDB
    Dim MyDB As DAO.Database:       Set MyDB = CurrentDb
    Dim TemporDB As DAO.Database:   Set TemporDB = TempDB
    
    'DELETE IT ON TABLEDEF
    If IsTableDefExist(TemporDB.TableDefs, TableName_TempDB) Then
        Call TemporDB.TableDefs.Delete(TableName_TempDB)
        
        'DELETE TABLEDEF ON THIS CURRENT DATABASE
        If IsTableDefExist(MyDB.TableDefs, TableName_TempDB & "_Linked") Then
            Call MyDB.TableDefs.Delete(TableName_TempDB & "_Linked")
        End If
    End If
    
    'EXIT
    RefreshDatabaseWindow
End Sub

Public Function TableExistsInTempDB(ByVal TableDefName As String) As Boolean
'TableExistsInTempDB
    On Error Resume Next

    Dim FindingTable As DAO.TableDef
    Set FindingTable = TempDB.TableDefs(TableDefName)
    
    Dim Exists As Boolean
    Exists = (Err.Number = 0)
    
    TableExistsInTempDB = Exists
    Err.Clear
    On Error GoTo 0

End Function


'============================= IMPLEMENTS ======================
Private Sub GenerateTempDatabase(ByRef DatabaseObject As DAO.Database)
    'GENERATE DATABASE FOR
'    Dim MyTempDatabasePath As String
'    MyTempDatabasePath = TempDBPath
'    Dim WSpace As DAO.Workspace, db As DAO.Database
'    Set WSpace = DBEngine.Workspaces(0)
'    Set db = WSpace.CreateDatabase(MyTempDatabasePath, dbLangGeneral)
'    Set DBObject = db
    
    'version2:
    Dim MyTempDatabasePath As String
    MyTempDatabasePath = TempDBPath
    
    Dim db As DAO.Database
    Set db = DBEngine.CreateDatabase(MyTempDatabasePath, dbLangGeneral)
    
    Set DatabaseObject = db
End Sub

'====================================================================
' PRIVATE HELPERS
'====================================================================
Private Function IsFileExist(ByVal FilePath As String, _
                             ByVal FileAttribute As VbFileAttribute) As Boolean
    Dim bln As Boolean
    bln = (Dir(FilePath, FileAttribute) <> "")
    
    IsFileExist = bln
End Function

Private Function IsTableDefExist(ByRef TableDefsCollection As DAO.TableDefs, _
                                 ByVal TableDefName As String) As Boolean
'    Call TableDefsCollection.Refresh
'
'    Dim Exists As Boolean: Exists = False
'
'    Dim TableDefObject As DAO.TableDef
'    For Each TableDefObject In TableDefsCollection
'        If TableDefObject.Name = TableDefName Then Exists = True
'    Next TableDefObject
'
'    IsTableDefExist = Exists
    
    'Version 2:
    On Error Resume Next
    Call TableDefsCollection.Refresh

    Dim FindingTable As DAO.TableDef
    Set FindingTable = TableDefsCollection(TableDefName)
    
    Dim Exists As Boolean
    Exists = (Err.Number = 0)
    
    IsTableDefExist = Exists
    Err.Clear
    On Error GoTo 0
End Function






