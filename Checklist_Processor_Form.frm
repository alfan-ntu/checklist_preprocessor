VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Checklist_Processor_Form 
   Caption         =   "�ˮ֪�e�B�z�u��"
   ClientHeight    =   7692
   ClientLeft      =   84
   ClientTop       =   360
   ClientWidth     =   11388
   OleObjectBlob   =   "Checklist_Processor_Form.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Checklist_Processor_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' Project : �ˮ֪� Checklist Preprocessing Tool
' Author : Al Fan
' Version : 1.0
' Coding date : 2020/12/24
'
Const Application_Version As String = "1.0"
Const Coding_Date = "2020/12/24"
'
'
'
Private Sub About_Button_Click()
    Const About_Title                  As String = "�ˮ֪�e�B�z�u��"
    MsgBox "�����s�� : ver. " & Application_Version & vbNewLine & "�{�ǽs�g��� : " & Coding_Date, vbInformation, About_Title
End Sub
'
' Select Checklist File Folder as the source of �ˮ֪�
'
Private Sub Checklist_Folder_Select_Button_Click()
    Dim Directory_String As String
    
    If Checklist_Folder_TextBox.Text <> "" Then
        Directory_String = Utility.GetFolder(Checklist_Folder_TextBox.Text)
    Else
        Directory_String = Utility.GetFolder("C:\")
    End If
    Checklist_Folder_TextBox.Text = Directory_String
    '
    ' Default directory to store CSV file is the same as that stores checklist(�ˮ֪�) files
    '
    Checklist_CSV_Folder_TextBox.Text = Directory_String
    
    Call SaveSetting(Constant.ApplicationName, Constant.RegistrySectionName, _
                    Constant.RegistryChecklistFolder, Directory_String)

    Call Utility.Write_Log(Utility.Log_Type_Info, "�ˮ֪��Ƨ�:" & Directory_String, True)
    Checklist_Processor_Form.Repaint
    
End Sub
'
' Select CSV File Folder for storing the consolidated CSV file; Default CSV folder is the Checklist File Folder
'
Private Sub CSV_Folder_Select_Button_Click()
    Dim Directory_String As String
    
    If Checklist_CSV_Folder_TextBox.Text <> "" Then
        Directory_String = Utility.GetFolder(Checklist_CSV_Folder_TextBox.Text)
    Else
        Directory_String = Utility.GetFolder("C:\")
    End If
    Checklist_CSV_Folder_TextBox.Text = Directory_String
    
    Call SaveSetting(Constant.ApplicationName, Constant.RegistrySectionName, _
                    Constant.RegistryCVSOutputFolder, Directory_Setting)

    Call Utility.Write_Log(Utility.Log_Type_Info, "CSV��Ƨ�:" & Directory_String, True)
    Checklist_Processor_Form.Repaint
End Sub
'
' Exit of this Form
'
Private Sub Exit_Button_Click()
    Unload Me
    Application.ScreenUpdating = True
End Sub
'
' Subject: Generate_CSV_Button_Click() traverses the selected Checklist File Folder, parses each checklist file
'          extracts data from the checklist file to compose consolidated CSV file for uploading
'
Private Sub Generate_CSV_Button_Click()
    Dim ExcelFile, ExcelFileFolder                  As String
    Dim sourceWorkbook                              As Workbook
    Dim recordCount, totalRecordCount               As Integer
    Dim completePercentage                          As Double
    Dim CSVFileName                                 As String
    Dim objCSVFile                                  As Object       ' A file system object
    Dim streamCSVFile                               As TextStream
   
    '
    ' Prepare the target CSV file for consolidating all the checklist files
    '
    Set objCSVFile = CreateObject("Scripting.FileSystemObject")
    CSVFileName = Checklist_Processor_Form.Checklist_CSV_Folder_TextBox.Text & "\�ˮ֪�J��_" & Format(Date, "yyyymmdd") & ".csv"
    Call Utility.Write_Log(Utility.Log_Type_Info, "CSV �ɮצW��:" & CSVFileName, True)
    If objCSVFile.FileExists(CSVFileName) Then
        Debug.Print "Target CSV file existis!"
        Kill CSVFileName
        ' Set streamCSVFile = objCSVFile.OpenTextFile(CSVFileName, F, True, TristateTrue)
        Set streamCSVFile = objCSVFile.CreateTextFile(CSVFileName, True, True)
    Else
        Debug.Print "Target CSV file does not exist, create one!"
        Set streamCSVFile = objCSVFile.CreateTextFile(CSVFileName, True, True)
    End If
    '
    ' Traverse all the checklist Excel files
    '
    If Checklist_Folder_TextBox.Text = "" Then
        MsgBox "�Ы��w�ˮ֪��Ƨ�"
        Exit Sub
    Else
        ExcelFileFolder = Checklist_Folder_TextBox.Text
        recordCount = 0
        totalRecordCount = Utility.Get_Number_Of_Excel_Files(ExcelFileFolder)
        
        ExcelFile = Dir(ExcelFileFolder & "\*.xls?")
        ' Traverses ExcelFileFolder
        Call ProgressBar_Form.Init_Progress_Bar
        
        Do While ExcelFile <> ""
            Application.ScreenUpdating = False
            Set sourceWorkbook = Workbooks.Open(ExcelFileFolder & "\" & ExcelFile)
            recordCount = recordCount + 1
            Call Utility.Write_Log(Utility.Log_Type_Info, "�ˮ֪�:" & ExcelFile, False)
            Checklist_Processor_Form.Repaint
            Call Utility.Extract_And_Save(ExcelFile, sourceWorkbook, streamCSVFile)
            sourceWorkbook.Close
            Set sourceWorkbook = Nothing
            Application.ScreenUpdating = True

            completePercentage = CDbl(recordCount) / CDbl(totalRecordCount)
            ProgressBar_Form.Set_Progress_Percentage (completePercentage)
            
            ExcelFile = Dir()
        Loop
        Unload ProgressBar_Form
        
        Call Utility.Write_Log(Utility.Log_Type_Info, "�`�@�פJ " & CStr(recordCount) & " ���ˮ֪�", False)
    End If
    Set streamCSVFile = Nothing
    Set objCSVFile = Nothing
End Sub


