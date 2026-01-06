VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmProductList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'DB OBJECTS
Private Const c_Table_Database = "Products"
Private Const c_Table_TempDatabase As String = "Products"
Private Const c_Table_Linked As String = "Products_Linked"
Private Const c_Query_GenerateDataSource = "GUI_frmProductList_GenerateDataSource"
Private Const c_Query_GUI = "GUI_frmProductList"
'STATES
Private m_DataChanged As Boolean
Private m_NewData As Boolean
Private m_IsClosing As Boolean
'DATA
Private m_ImagePath As String
'====================================================================
' EVENTS
'====================================================================
Public Event HasImage(ByVal strImagePath As String)

'====================================================================
' STATES
'====================================================================
Public Property Get NewData() As Boolean
    NewData = m_NewData
End Property
Public Property Let NewData(ByVal Value As Boolean)
    m_NewData = Value
End Property

Public Property Get DataChanged() As Boolean
    DataChanged = m_DataChanged
End Property

Public Property Let DataChanged(ByVal Value As Boolean)
    m_DataChanged = Value
End Property

'====================================================================
' OTHER INFORMATION
'====================================================================
' TABLE
Public Property Get Table_Database() As String
    Table_Database = c_Table_Database
    'UPDATE TABLE
End Property

Public Property Get Table_Linked() As String
    Table_Linked = c_Table_Linked
End Property

Public Property Get Table_TempDatabase() As String
    Table_TempDatabase = c_Table_TempDatabase
End Property

'QUERY
Public Property Get Query_GUI() As String
    Query_GUI = c_Query_GUI
End Property
Public Property Get Query_GenerateDataSource() As String
    Query_GenerateDataSource = c_Query_GenerateDataSource
End Property

'DATA
Public Property Get CurrentImagePath() As String
    CurrentImagePath = m_ImagePath
End Property
Public Property Let CurrentImagePath(ByVal Value As String)
    m_ImagePath = Value
End Property

'====================================================================
' OPEN SETTING
'====================================================================
Private Sub Form_Load()
On Error GoTo ErrHandler:
    Call SpeedCheck("Form_Load")
    
    
    With Form_frmProductList
        'REFRESH DATA & MAKE SNAPSHOT
        Call frmProductList_RefreshData
        'LINK TEMP DATABASE
        Me.RecordSource = .Query_GUI
        'BUILD ROW STRING
        Dim strTXTROWNUM, Q As String:
        Q = """"
        strTXTROWNUM = "=IIf(IsNull([ProductID]),Null,DCount(" & Q & "*" & Q & "," & _
                                                            Q & .Query_GUI & Q & "," & _
                                                            Q & "ProductID <" & Q & " & [ProductID])+1)"
        Me.txtRowNum.ControlSource = strTXTROWNUM
    End With
    'DEBUG
    Call SpeedCheck("Form_Load")
    Call DEBUG_ProcedureResult("Form_Load", "Success")
    Exit Sub
ErrHandler:
    Dim ErrString As String
    ErrString = "Object: frmProductList" & vbCrLf & _
                "Error Number: " & Err.Number & vbCrLf & _
                "Error Describtion: " & Err.Description
    Call MsgBox(ErrString)
End Sub

'====================================================================
' SETTING BEFORE CLOSE
'====================================================================
Private Sub Form_Unload(Cancel As Integer)
    Call SpeedCheck("Form_Unload")
    
    With Form_frmProductList
        'TRACK CHANGES
        Dim DataIsChanged As Boolean: DataIsChanged = frmProductList_TrackChanges
        
        'NOTIFY CHANGES
        Dim UserConfirm As Integer: If DataIsChanged Then UserConfirm = frmProductList_DataChangeNotify
        
        'PERFORM DATA CHANGE OPERATION IF ANY
        If UserConfirm = vbYes Then Call frmProductList_MakeChanges
    End With

    'RELEASE
    Me.txtRowNum.ControlSource = ""
    m_IsClosing = True
    'DEBUG
    Call SpeedCheck("Form_Unload")
    Call DEBUG_ProcedureResult("Form_Unload", "Success")
    Call DEBUG_PrintVariableChanges("m_IsClosing", m_IsClosing, "FormProductList_Unload")
End Sub

Private Sub Form_Close()
    Me.RecordSource = ""
    'REALEASE
    Form_frmProductList.CurrentImagePath = ""
    
    Call DEBUG_PrintVariableChanges("ImagePath", Form_frmProductList.CurrentImagePath, "frmProductList")
End Sub


'====================================================================
' DELETE FEATURE/ DELETE EVENTS
'====================================================================
Private Sub Form_BeforeDelConfirm(Cancel As Integer, Response As Integer)
    Cancel = True
End Sub

Private Sub Form_AfterDelConfirm(Status As Integer)
    'On Error GoTo ErrHandler:
    Call SpeedCheck("Form_AfterDelConfirm")

    'UPDATE DELETE DATA
    Call frmProductList_AfterDelConfirm_UpdateIsActive(Me)
    
    'UPDATE DELETE DATA V2

    'RESET DATA
    Me.RecordSource = ""
    Me.RecordSource = Form_frmProductList.Query_GUI
    'DEBUG
    Call SpeedCheck("Form_AfterDelConfirm")
    Call DEBUG_ProcedureResult("Form_AfterDelConfirm", "Success")
    Exit Sub
ErrHandler:
End Sub

'====================================================================
' UPDATE EVENTS
'====================================================================
Private Sub Form_AfterUpdate()
    'On Error GoTo ErrHandler:
    Call SpeedCheck("Form_AfterUpdate")

    'LOGIC
    Call frmProductList_AfterUpdate_UpdateChangeDate(Me)

    'DEBUG
    Call SpeedCheck("Form_AfterUpdate")
    Call DEBUG_ProcedureResult("Form_AfterUpdate", "Success")
    Exit Sub
ErrHandler:
End Sub


Private Sub Form_Current()
    If Not Me.Recordset Is Nothing Then
        Dim imgPath As String
        imgPath = Nz(Me!ImagePath, "")
        Form_frmProductList.CurrentImagePath = imgPath
        RaiseEvent HasImage(imgPath)
        Call DEBUG_PrintVariableChanges("m_ImagePath", m_ImagePath, "Form_frmProductList_OnCurrent")
    End If
End Sub

'====================================================================
' CUSTOMS FUNCTIONS
'====================================================================

