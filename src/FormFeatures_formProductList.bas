Attribute VB_Name = "FormFeatures_formProductList"
Option Compare Database
Option Explicit

Public Function frmProductList_RefreshData()
'REGENERATE DATA SOURCE AND MAKE BEFORE SNAPSHOT AND MAKE TABLE CHANGES
    
    With Form_frmProductList
        'GENERATE DATA SOURCE
        Call RefreshData(.Table_TempDatabase, .Query_GenerateDataSource)
        ' SNAPSHOT
        Call CurrentDb.Execute(SQL_CurrentDB_UpdateBeforeHash(.Table_Linked))
        
        'CUSTOMS:
            'MAKE BARCODE UNIQUE
            AccessTempDatabase.TempDatabase.Execute _
            (" CREATE UNIQUE INDEX idx_Barcode ON " & .Table_TempDatabase & " (Barcode) ")
    End With
    
End Function

Public Function frmProductList_TrackChanges()
    Dim DataIsChanged As Boolean, IsChanged As Boolean, IsNew As Boolean
    With Form_frmProductList
        DataIsChanged = TrackChanges(IsChanged, IsNew, .Table_Linked)
        .DataChanged = IsChanged                        'GET STATE
        .NewData = IsNew                                'GET STATE
    End With
    
    frmProductList_TrackChanges = DataIsChanged
    'DEBUG
    Call DEBUG_PrintVariableChanges("Form_frmProductList.DataChanged", Form_frmProductList.DataChanged, "TrackChanges")
    Call DEBUG_PrintVariableChanges("Form_frmProductList.NewData", Form_frmProductList.NewData, "TrackChanges")
End Function

Public Function frmProductList_DataChangeNotify()
    Dim UserConfirm As Integer
    UserConfirm = DataChangeNotify
    
    frmProductList_DataChangeNotify = UserConfirm
End Function

Public Sub frmProductList_MakeChanges()
On Error GoTo ErrHandler:
    Call SpeedCheck("frmProductList_MakeChanges")
    DBEngine.BeginTrans
    With Form_frmProductList
        'INSERT/UPDATE PARENT TABLES
        
        If .DataChanged Or .NewData Then
            Call frmProductList_MakeChanges_UpdateParents(.Table_Linked)
        End If
        'UPDATE CHANGED DATA
        If .DataChanged Then
            Call frmProductList_MakeChanges_UpdateProducts(.Table_Database, .Table_Linked)
        End If
    
        'INSERT NEW PRODUCT
        If .NewData Then
            Call frmProductList_MakeChanges_InsertProducts(.Table_Database, .Table_Linked)
        End If
        
        'UPDATE BEFOREHASH AFTER DATA CHANGED
        If .NewData Or .DataChanged Then
            Call CurrentDb.Execute(SQL_CurrentDB_UpdateBeforeHash(.Table_Linked))
        End If
    End With
    DBEngine.CommitTrans
    
    'DEBUG
    Call SpeedCheck("frmProductList_MakeChanges")
    Call DEBUG_ProcedureResult("frmProductList_MakeChanges", "Success")
    Exit Sub
ErrHandler:
    DAO.Rollback
    Call DEBUG_ProcedureResult("frmProductList_MakeChanges", "Fail")
End Sub

Private Function frmProductList_MakeChanges_UpdateParents(ByVal CurrentDB_LinkedTableName As String) As String
On Error GoTo ErrHandler:
    Call SpeedCheck("frmProductList_MakeChanges_UpdateProducts")

    Call DBEngine.BeginTrans
    'SET UP
    Dim db As DAO.Database: Set db = CurrentDb
    'LOOP FIELDS -> UPDATE FIELD BY FIELD
    Dim FieldObject As DAO.Field
    For Each FieldObject In db.TableDefs(CurrentDB_LinkedTableName).Fields
        Select Case FieldObject.Name
'            Case "MarketingCategoryID"                              'Logic: Different update
'                Call db.Execute(SQL_CurrentDB_InsertParentTable("MarketingCategories", _
'                                                                        CurrentDB_LinkedTableName, _
'                                                                        "MarketingCategoryID", _
'                                                                        "MarketingCategoryID"))
            Case "MarketingLineID"                                  'Logic: Different update
                Call db.Execute(SQL_CurrentDB_InsertParentTable("MarketingLines", _
                                                                       CurrentDB_LinkedTableName, _
                                                                       "MarketingLineID", _
                                                                       "MarketingLineID"))
                                                                       
            Case "PackageTypeID"                                    'Logic: Different update
                Call db.Execute(SQL_CurrentDB_InsertParentTable("PackageTypes", _
                                                                      CurrentDB_LinkedTableName, _
                                                                      "PackageTypeID", _
                                                                      "PackageTypeID"))
            Case "CapacityUnitID"
                Call db.Execute(SQL_CurrentDB_InsertParentTable("CapacityUnits", _
                                                                      CurrentDB_LinkedTableName, _
                                                                      "CapacityUnitID", _
                                                                      "CapacityUnitID"))
        End Select
    Next FieldObject
    
    'COMMIT CHANGES
    Call DBEngine.CommitTrans
    frmProductList_MakeChanges_UpdateParents = "Update Database:" & Space(10) & "Success"
    'DEBUGING
    Call SpeedCheck("frmProductList_MakeChanges_UpdateParents")
    Call DEBUG_ProcedureResult("frmProductList_MakeChanges_UpdateParents", "Success")
    Exit Function
ErrHandler:
    Call DBEngine.Rollback
    frmProductList_MakeChanges_UpdateParents = "Update Database:" & Space(10) & "Fail"
End Function

Private Function frmProductList_MakeChanges_UpdateProducts(ByVal CurrentDB_DataTableName As String, _
                                                           ByVal CurrentDB_LinkedTableName As String) As String
On Error GoTo ErrHandler:
    Call SpeedCheck("frmProductList_MakeChanges_UpdateProducts")

    Call DBEngine.BeginTrans
    'SET UP
    Dim db As DAO.Database: Set db = CurrentDb
    'LOOP FIELDS -> UPDATE FIELD BY FIELD
    Dim FieldObject As DAO.Field
    For Each FieldObject In db.TableDefs(CurrentDB_LinkedTableName).Fields
        If IsFieldExist(FieldObject.Name, CurrentDB_DataTableName) Then
            Select Case FieldObject.Name
                Case "ProductID", "InsertDate"                          'Logic: Skip ID
                Case Else                                               'Logic: Update field by field
                    'UPDATE ONLY CHANGE VALUE
                    Call CurrentDb.Execute(SQL_CurrentDB_UpdateTable(CurrentDB_DataTableName, _
                                                                           CurrentDB_LinkedTableName, _
                                                                           FieldObject.Name, "ProductID"))
            End Select
        End If
    Next FieldObject
    
    'COMMIT CHANGES
    Call DBEngine.CommitTrans
    frmProductList_MakeChanges_UpdateProducts = "Update Database:" & Space(10) & "Success"
    'DEBUGING
    Call SpeedCheck("frmProductList_MakeChanges_UpdateProducts")
    Call DEBUG_ProcedureResult("frmProductList_MakeChanges_UpdateProducts", "Success")
    Exit Function
ErrHandler:
    Call DBEngine.Rollback
    frmProductList_MakeChanges_UpdateProducts = "Update Database:" & Space(10) & "Fail"
End Function

Private Function frmProductList_MakeChanges_InsertProducts(ByVal CurrentDB_DataTableName As String, _
                                                           ByVal CurrentDB_LinkedTableName As String) As String
'NOTE: This also insert Time of Insert
On Error GoTo ErrHandler:
    Call SpeedCheck("frmProductList_MakeChanges_InsertProducts")
    
    'BEGIN TRANS
    Call DBEngine.BeginTrans
    'SET UP
    Dim db As DAO.Database:    Set db = CurrentDb
    Dim i As Integer:          i = 0
    
    Dim FieldObject As DAO.Field
    Dim ArrayOfFieldNames() As String, ArrayOfValueNames() As String
    'DYNAMICALLY BUILD ARRAY
        '* Replace InsertDate field = Now()
    For Each FieldObject In db.TableDefs(CurrentDB_LinkedTableName).Fields
        If IsFieldExist(FieldObject.Name, CurrentDB_DataTableName) Then
            Select Case FieldObject.Name
                Case "ProductID", "IsActive", "ChangeDate"              'SKIP THESE FIELDS
                Case Else
                    'BUILD FIELDS ARRAY
                    ReDim Preserve ArrayOfFieldNames(i)
                    ArrayOfFieldNames(i) = FieldObject.Name
                    'BUILD VALUES ARRAY
                    ReDim Preserve ArrayOfValueNames(i)
                    If FieldObject.Name = "InsertDate" Then             'INSERT TODAY TIME
                        ArrayOfValueNames(i) = "Now()"
                    Else
                        ArrayOfValueNames(i) = FieldObject.Name
                    End If
                    'NEXT OBJECT
                    i = i + 1
            End Select
        End If
    Next FieldObject
    'BUILD STRING
    Dim ListOfFields As String, ListOfValues As String
    ListOfFields = Join(ArrayOfFieldNames, ",")
    ListOfValues = Join(ArrayOfValueNames, ",")
    
    'RUN SQL
    Call CurrentDb.Execute(SQL_CurrentDB_InsertIntoTable(CurrentDB_DataTableName, _
                                                               CurrentDB_LinkedTableName, _
                                                               ListOfFields, _
                                                               ListOfValues))
    
    'COMMIT THE TRANS
    Call DBEngine.CommitTrans
    frmProductList_MakeChanges_InsertProducts = "Insert Database:" & Space(10) & "Success"
    'DEBUGING
    Call SpeedCheck("frmProductList_MakeChanges_InsertProducts")
    Call DEBUG_ProcedureResult("frmProductList_MakeChanges_InsertProducts", "Success")
    Exit Function
ErrHandler:
    Call DBEngine.Rollback
    MsgBox DBEngine.Errors(0).Number & " " & DBEngine.Errors(0).Description & " " & DBEngine.Errors(0).Source
    frmProductList_MakeChanges_InsertProducts = "Insert Database:" & Space(10) & "Fail"
End Function


'------------------------------------------------------------------------------------------------
'UI FUNCTIONS
Public Function frmProductList_AfterDelConfirm_UpdateIsActive(ByRef FormObject As Access.Form)
    'On Error GoTo ErrHandler:
    Call SpeedCheck("frmProductList_AfterDelConfirm_UpdateIsActive")
    DBEngine.BeginTrans
    'LOGIC
    
    With Form_frmProductList
        'USER CONFIRM
        Dim UserConfirm As Integer:
        UserConfirm = MsgBox("Do you want to delete " & FormObject.SelHeight & " record?", vbYesNo)
        
        If UserConfirm = vbYes Then
            'BUILD ID LIST
            Dim SelectingRows As Long:          SelectingRows = FormObject.SelHeight
            Dim rs As DAO.Recordset:            Set rs = FormObject.Recordset
            
            With rs
                .MoveFirst
                .Move (FormObject.SelTop - 1)
                If (.EOF Or .BOF) Then Exit Function
                
                Dim i As Long, IDList() As Long
                For i = 0 To (SelectingRows - 1)
                    ReDim Preserve IDList(i)
                    IDList(i) = FormObject!ProductID
                    .MoveNext
                Next i
            End With
            'UPDATE MULTIPLE
            For i = LBound(IDList) To UBound(IDList)
                Call CurrentDb.Execute(SQL_CurrentDB_UITable_Deactivate(IDList(i), "ProductID", .Table_Linked))
            Next i
        End If
    End With
    
    DBEngine.CommitTrans
    'DEBUG
    Call SpeedCheck("frmProductList_AfterDelConfirm_UpdateIsActive")
    Call DEBUG_ProcedureResult("frmProductList_AfterDelConfirm_UpdateIsActive", "Success")
    Exit Function
ErrHandler:
    DBEngine.Rollback
End Function


Public Function frmProductList_AfterUpdate_UpdateChangeDate(ByRef FormObject As Access.Form)
    On Error GoTo ErrHandler:
    Call SpeedCheck("frmProductList_AfterUpdate_UpdateChangeDate")
    DBEngine.BeginTrans
    
    With Form_frmProductList
        Call CurrentDb.Execute(SQL_CurrentDB_UITable_UpdateChangeDate(FormObject!ProductID, "ProductID", .Table_Linked))
    End With
    
    DBEngine.CommitTrans
    'DEBUG
    Call SpeedCheck("frmProductList_AfterUpdate_UpdateChangeDate")
    Call DEBUG_ProcedureResult("frmProductList_AfterUpdate_UpdateChangeDate", "Success")
    Exit Function
ErrHandler:
    DBEngine.Rollback
End Function

