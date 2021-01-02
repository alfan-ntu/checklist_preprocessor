VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Checklist_Processor_Form 
   Caption         =   "�ˮ֪�e�B�z�u��"
   ClientHeight    =   8268.001
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

    Call Utility.Write_Log(Utility.Log_Type_Info, "����ˮ֪��Ƨ�:" & Directory_String, True)
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

    Call Utility.Write_Log(Utility.Log_Type_Info, "���CSV��Ƨ�:" & Directory_String, True)
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
'
'
Private Sub Export_Log_Button_Click()
    Dim LogFileName                                 As String
    Dim objLogFile                                  As Object       ' A file system object
    Dim streamLogFile                               As TextStream
    
    Set objLogFile = CreateObject("Scripting.FileSystemObject")
    LogFileName = Checklist_Processor_Form.Checklist_CSV_Folder_TextBox.Text & "\Log_" & Format(Date, "yyyymmdd") & ".txt"
    Call Utility.Write_Log(Utility.Log_Type_Info, "�ާ@��x�ɮ�:" & LogFileName, True)
    If objLogFile.FileExists(LogFileName) Then
        Set streamLogFile = objLogFile.OpenTextFile(LogFileName, ForAppending, True, TristateTrue)
    Else
        Set streamLogFile = objLogFile.CreateTextFile(LogFileName, True, True)
    End If
    streamLogFile.Write (Checklist_Processor_Form.Log_TextBox.Text)
    
    Set objLogFile = Nothing
    Set streamLogFile = Nothing
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
        Debug.Print "Target CSV file existis! Delete the existed one!"
        Kill CSVFileName
    Else
        Debug.Print "Target CSV file does not exist, create one!"
    End If
    Set streamCSVFile = objCSVFile.CreateTextFile(CSVFileName, True, True)
    '
    ' Traverse all the checklist Excel files
    '
    If Checklist_Folder_TextBox.Text = "" Then
        MsgBox "�Ы��w�ˮ֪��Ƨ�"
        GoTo HouseKeeping
    Else
        ExcelFileFolder = Checklist_Folder_TextBox.Text
        recordCount = 0
        totalRecordCount = Utility.Get_Number_Of_Excel_Files(ExcelFileFolder)
        ExcelFile = Dir(ExcelFileFolder & "\*.xls?")
        If Checklist_Processor_Form.Header_CheckBox.Value = True Then
            Call Utility.Add_Header_Row(streamCSVFile)
        End If
        ' Traverses ExcelFileFolder
        Call ProgressBar_Form.Init_Progress_Bar
        Do While ExcelFile <> ""
            Application.ScreenUpdating = False
            Set sourceWorkbook = Workbooks.Open(ExcelFileFolder & "\" & ExcelFile)
            recordCount = recordCount + 1
            Call Utility.Write_Log(Utility.Log_Type_Info, "�פJ�ˮ֪�:" & ExcelFile, False)
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
        Call Utility.Write_Log(Utility.Log_Type_Info, "�J�`�ˮ֪��x�s�b:" & CSVFileName, False)
    End If

HouseKeeping:
    Set streamCSVFile = Nothing
    Set objCSVFile = Nothing
End Sub


