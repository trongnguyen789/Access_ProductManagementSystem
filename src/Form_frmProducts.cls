VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private WithEvents Listener_Form_frmProductList As Form_frmProductList
Attribute Listener_Form_frmProductList.VB_VarHelpID = -1

Private Sub Form_Load()
    
    Set Listener_Form_frmProductList = Me.sfProductList.Form
    
    If Me.cmdLanguage.Caption = "EN" Then
        Call SwitchLanguage("VN")
    ElseIf Me.cmdLanguage.Caption = "VN" Then
        Call SwitchLanguage("EN")
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set Listener_Form_frmProductList = Nothing
    Dim rs As DAO.Recordset2
End Sub

Private Sub Listener_Form_frmProductList_HasImage(ByVal strImagePath As String)
    If Me.cmdViewImage.Caption = "HIDE IMAGE" Then
        If Nz(strImagePath, "") = "" Then
            Me.LabelNoImage.Visible = True
        Else
            Me.LabelNoImage.Visible = False
        End If
    End If
End Sub



Private Sub cmdAddImage_Click()
    
    Dim IDNumber As Long
    IDNumber = Nz(Me.sfProductList.Form!ProductID, 0)
    
    Dim dialog As Office.FileDialog
    Set dialog = Application.FileDialog(msoFileDialogFilePicker)
    With dialog
        'SHOW IT
        .Filters.Add "Images", "*.gif; *.jpg; *.jpeg; *.png ", 1
        .Show
    End With
    
    If dialog.SelectedItems.Count <> 0 Then
        Call CurrentDb.Execute(SQL_CurrentDB_frmProductList_UpdateImagePath( _
                            IDNumber, dialog.SelectedItems(1), Form_frmProductList.Table_Linked))
        Me.Refresh
    End If
                            
                            
                            
End Sub

Private Sub cmdLanguage_Click()
    Call SwitchLanguage(Me.cmdLanguage.Caption)
    Select Case Me.cmdLanguage.Caption
        Case "VN"
            Me.cmdLanguage.Caption = "EN"
        Case "EN"
            Me.cmdLanguage.Caption = "VN"
    End Select
End Sub

Private Sub cmdRefreshData_Click()
    Call SpeedCheck("cmdRefreshData_Click")
    
    With Form_frmProductList
        'MAKE SNAPSHOT AND IDENTIFY CHANGES
        Call frmProductList_TrackChanges

        'MAKE CHANGES IF ANY
            'CONFIRM?
        Dim UserConfirm As Integer: UserConfirm = 0
        If .DataChanged Or .NewData Then UserConfirm = frmProductList_DataChangeNotify
            'RUN CHANGE
        If UserConfirm = vbYes Then Call frmProductList_MakeChanges

        'REFRESH DATA & 'MAKE SNAPSHOT
        Me.sfProductList.Form.RecordSource = ""
        
        'REGENERATE DATA SOURCE AND MAKE BEFORE SNAPSHOT AND MAKE TABLE CHANGES
        Call frmProductList_RefreshData
        
        Me.sfProductList.Form.RecordSource = .Query_GUI
    End With
    
    'DEBUG
    Call SpeedCheck("cmdRefreshData_Click")
    Call DEBUG_ProcedureResult("cmdRefreshData_Click", "Success")
    
End Sub

Private Sub cmdRenew_Click()

    Me.sfProductList.Form.OrderBy = ""
    
End Sub

Private Sub cmdSaveProducts_Click()
    Call SpeedCheck("cmdRenew_Click")
    
    With Form_frmProductList
        Call frmProductList_TrackChanges
        Call frmProductList_MakeChanges
    End With
    'DEBUG
    Call SpeedCheck("cmdRenew_Click")
    Call DEBUG_ProcedureResult("cmdRefreshData_Click", "Success")
End Sub



Private Sub cmdViewImage_Click()
    With Me
        Select Case cmdViewImage.Caption
            Case "VIEW IMAGE"
                .imgProduct.Visible = True
                .cmdAddImage.Visible = True
                .cmdClearImage.Visible = True
                .sfProductList.Width = 6.6 * TWIPTS_PER_INCH
                .imgProduct.Requery
                .cmdViewImage.Caption = "HIDE IMAGE"
                'OTHER
                If Form_frmProductList.CurrentImagePath = "" Then .LabelNoImage.Visible = True
            Case "HIDE IMAGE"
                .imgProduct.Visible = False
                .cmdAddImage.Visible = False
                .cmdClearImage.Visible = False
                .sfProductList.Width = 10.9 * TWIPTS_PER_INCH
                .imgProduct.Requery
                .cmdViewImage.Caption = "VIEW IMAGE"
                'OTHER
                .LabelNoImage.Visible = False
        End Select
    End With
End Sub




'------------------------------------------------------------------
'FEATURES
Private Sub SwitchLanguage(ByVal LanguageToSwitch As String)
    With Me.sfProductList.Form
        Select Case LanguageToSwitch
            Case "VN"
                .Controls("LabelProductName").Caption = "Product Name (VN)"
                .Controls("txtProductName").ControlSource = "ProductName_vn"
                
                .Controls("LabelPrice").Caption = "Price (VND)"
                .Controls("txtPrice").ControlSource = "Price_vnd"
    
                .Controls("LabelUSP").Caption = "USP (VN)"
                .Controls("txtUSP").ControlSource = "USP_vn"
    
                .Controls("LabelHowToUse").Caption = "How To Use (VN)"
                .Controls("txtHowToUse").ControlSource = "HowToUse_vn"
            
            Case "EN"
                .Controls("LabelProductName").Caption = "Product Name (EN)"
                .Controls("txtProductName").ControlSource = "ProductName_en"

                .Controls("LabelPrice").Caption = "Price (EN)"
                .Controls("txtPrice").ControlSource = "Price_usd"

                .Controls("LabelUSP").Caption = "USP (EN)"
                .Controls("txtUSP").ControlSource = "USP_en"

                .Controls("LabelHowToUse").Caption = "How To Use (EN)"
                .Controls("txtHowToUse").ControlSource = "HowToUse_en"
        End Select
    End With
End Sub
