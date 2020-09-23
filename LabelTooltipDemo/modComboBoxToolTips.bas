Attribute VB_Name = "modComboBoxToolTips"
  '********************************************************************************************'
  '********************************************************************************************'
  
  'module by tom l (dvlp@tjl-enterprises.net), may 23,2009
  
  'this module is depending on the CToolTip.cls being part of the project
  
  '********************************************************************************************'
  '********************************************************************************************'
Option Explicit
    
    Private colComboBoxToolTips As New Collection
            
    Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
    End Type
    
    Private Type PCOMBOBOXINFO
        cbSize As Long
        rcItem As RECT
        rcButton As RECT
        stateButton As Long
        hwndCombo As Long
        hwndItem As Long
        hwndList As Long
    End Type
    
    Private Declare Function GetComboBoxInfo Lib "user32.dll" (ByVal hwndCombo As Long, ByRef pcbi As PCOMBOBOXINFO) As Long

Public Sub CreateComboBoxToolTip(cbo As ComboBox, ParentFormName As String, ToolTipTitle As String, ToolTipText As String)
On Error GoTo errHandler

    Dim ToolTipParentHandle As Long
    Dim NewTT As CTooltip
    
    ToolTipParentHandle = cbo.hWnd
    
    Set NewTT = New CTooltip
        NewTT.Style = TTBalloon
        NewTT.Icon = TTIconInfo
        NewTT.Title = ToolTipTitle
        NewTT.TipText = ToolTipText
        
        NewTT.Create ToolTipParentHandle
        
        colComboBoxToolTips.Add NewTT, ParentFormName & "_" & cbo.Name & "_1"
    Set NewTT = Nothing
    
    If cbo.Style = 0 Then
        Dim ComboInfo As PCOMBOBOXINFO
        ComboInfo.cbSize = Len(ComboInfo)

        GetComboBoxInfo cbo.hWnd, ComboInfo
        ToolTipParentHandle = ComboInfo.hwndItem
        
        Set NewTT = New CTooltip
            NewTT.Style = TTBalloon
            NewTT.Icon = TTIconInfo
            NewTT.Title = ToolTipTitle
            NewTT.VisibleTime = 9000
            NewTT.TipText = ToolTipText
            NewTT.Create ToolTipParentHandle
            
            colComboBoxToolTips.Add NewTT, ParentFormName & "_" & cbo.Name & "_2"
        Set NewTT = Nothing
    End If


Exit Sub
errHandler:
    If Err.Number = 0 Then
        'nothing here
    Else
        MsgBox "CreateComboBoxToolTip @ " & Erl & vbNewLine & "unexpected error has occurred" & vbNewLine & _
        "error number : " & Err.Number & vbNewLine & _
        "description : " & Err.Description, vbCritical, "ERROR"
    End If

End Sub


Public Sub DeleteComboBoxToolTips(cbo As ComboBox, ParentFormName As String)
On Error GoTo errHandler

    Dim tempTT As CTooltip

    Set tempTT = colComboBoxToolTips(ParentFormName & "_" & cbo.Name & "_1")
    tempTT.Destroy
    Set tempTT = Nothing
    colComboBoxToolTips.Remove (ParentFormName & "_" & cbo.Name & "_1")
    
    If cbo.Style = 0 Then
        Set tempTT = colComboBoxToolTips(ParentFormName & "_" & cbo.Name & "_2")
        tempTT.Destroy
        Set tempTT = Nothing
        colComboBoxToolTips.Remove (ParentFormName & "_" & cbo.Name & "_2")
    End If

Exit Sub
errHandler:
    If Err.Number = 5 Or Err.Number = 91 Then
        Resume Next
    Else
        MsgBox "DeleteComboBoxToolTips @ " & Erl & vbNewLine & "unexpected error has occurred" & vbNewLine & _
        "error number : " & Err.Number & vbNewLine & _
        "description : " & Err.Description, vbCritical, "ERROR"
    End If

End Sub
