VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    ' This ScreenUpdating = False, is necessary to hide Excel file this is being opened
    Application.ScreenUpdating = False
    Checklist_Processor_Form.Checklist_Folder_TextBox = Utility.Get_Checklist_Folder
    Checklist_Processor_Form.Checklist_CSV_Folder_TextBox = Utility.Get_Checklist_CSV_Folder
    Checklist_Processor_Form.Show vbModeless
End Sub

