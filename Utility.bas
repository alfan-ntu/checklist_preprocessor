Attribute VB_Name = "Utility"
'
' Constants
'
Public Const Log_Type_Error          As Integer = 0
Public Const Log_Type_Warning        As Integer = 1
Public Const Log_Type_Info           As Integer = 2
Public Const Log_Type_Verbose        As Integer = 255
Private Const Log_Type_Default      As Integer = Log_Type_Info

'
' Pop out directory select dialog and return selected folder
'
Function GetFolder(strPath As String) As String
    Dim fldr As FileDialog
    Dim sItem As String

    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        .InitialFileName = strPath
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With

NextCode:
    GetFolder = sItem
    Set fldr = Nothing
End Function

'
' Get_Checklist_Folder() returns default or stored file of ���Ӫ�
'
Public Function Get_Checklist_Folder() As String
    Dim Checklist_Folder As String
    
    Checklist_Folder = GetSetting(Constant.ApplicationName, Constant.RegistrySectionName, Constant.RegistryChecklistFolder, _
                        "C:\")
    Get_Checklist_Folder = Checklist_Folder
End Function
'
'
'
Public Function Get_Checklist_CSV_Folder() As String
    Dim Checklist_CSV_Folder As String
    
    Checklist_CSV_Folder = GetSetting(Constant.ApplicationName, Constant.RegistrySectionName, Constant.RegistryChecklistFolder, _
                        "")
    Get_Checklist_CSV_Folder = Checklist_CSV_Folder
End Function
'
' Log_Type_Error = 0
' Log_Type_Warning = 1
' Log_Type_Info = 2
' Log_Type_Verbose = 255
'
Public Sub Write_Log(LogMsgType As Integer, LogMsg As String, ToConsole As Boolean)
    Dim LogString   As String
    LogString = Checklist_Processor_Form.Log_TextBox.Text
    
    If ToConsole = True Then
        Debug.Print (LogMsg)
    End If
    If LogMsgType <= Log_Type_Default Then
        Checklist_Processor_Form.Log_TextBox.SetFocus
        If LogMsgType = Log_Type_Info Then
            Checklist_Processor_Form.Log_TextBox.Text = LogString & "[INFO-" & Date & Time & "]:" & LogMsg & vbNewLine
        ElseIf LogMsgType = Log_Type_Warning Then
            Checklist_Processor_Form.Log_TextBox.Text = LogString & "[WARNING-" & Date & Time & "]:" & LogMsg & vbNewLine
        ElseIf LogMsgType = Log_Type_Error Then
            Checklist_Processor_Form.Log_TextBox.Text = LogString & "[ERROR-" & Date & Time & "]:" & LogMsg & vbNewLine
        Else
            Checklist_Processor_Form.Log_TextBox.Text = LogString & "[DEBUG-" & Date & Time & "]:" & LogMsg & vbNewLine
        End If
    End If
End Sub
'
' Extract_And_Save() opens the sourceWorkbook, extracts data from the worksheet and stores them to the targetStream
'
Public Sub Extract_And_Save(ByVal sourceFileName As String, sourceWorkbook As Workbook, targetStream As TextStream)
    Dim tws         As Worksheet
    Dim appRecord   As String
    
    Set tws = sourceWorkbook.ActiveSheet
    appRecord = Extract_Case_ID(sourceFileName)                                                     ' �ˮ֪��ɮ׽s��
    appRecord = appRecord & ","
    appRecord = appRecord & sourceFileName                                                          ' �ˮ֪��ɮצW��
    appRecord = appRecord & tws.Range(Constant.Dealer_Range).Text & ","                             ' �g�P��
    appRecord = appRecord & tws.Range(Constant.Dealer_Contact_Range).Text & ","                     ' �g�P�өӿ�H
    appRecord = appRecord & tws.Range(Constant.Submit_Date_Range).Text & ","                        ' �g�P�Ӱe���
    If bankAccountVer = True Then
        appRecord = appRecord & tws.Range(Constant.Cause_to_Refund_Range_2021a).Text & ","          ' �h�|��]
        appRecord = appRecord & tws.Range(Constant.Cheque_Payee_Range_2021a).Text & ","             ' �h�|���ڤH
        appRecord = appRecord & tws.Range(Constant.Cheque_Payee_ID_Range_2021a).Text & ","          ' ���ڤH�����Ҧr��
        appRecord = appRecord & tws.Range(Constant.Bank_Range_2021a).Text & ","                     ' ���ڻȦ�
        appRecord = appRecord & tws.Range(Constant.Bank_Code_Range_2021a).Text & ","                ' ���ڻȦ�N�X
        appRecord = appRecord & tws.Range(Constant.Branch_Range_2021a).Text & ","                   ' ���ڻȦ����
        appRecord = appRecord & tws.Range(Constant.Branch_Code_Range_2021a).Text & ","              ' ���ڻȦ����N�X
        appRecord = appRecord & tws.Range(Constant.Bank_Account_Range_2021a).Text & ","             ' ���ڤH�Ȧ�b��
    Else
        appRecord = appRecord & tws.Range(Constant.Cause_to_Refund_Range).Text & ","
        appRecord = appRecord & tws.Range(Constant.Cheque_Payee_Range).Text & ","
        appRecord = appRecord & ","
        appRecord = appRecord & ","
        appRecord = appRecord & ","
        appRecord = appRecord & ","
        appRecord = appRecord & ","
        appRecord = appRecord & ","
    End If
    
    appRecord = appRecord & tws.Range(Constant.New_Vehicle_Brand_Range).Text & ","                  ' �s���~�P
    appRecord = appRecord & tws.Range(Constant.New_Vehicle_Model_Range).Text & ","                  ' �s������
    appRecord = appRecord & tws.Range(Constant.New_Vehicle_Factory_Date_Range).Text & ","           ' �s���X�t�~��
    appRecord = appRecord & tws.Range(Constant.Old_Vehicle_Brand_Range).Text & ","                  ' �¨��~�P
    appRecord = appRecord & tws.Range(Constant.New_Vehicle_Owner_Name_Range).Text & ","             ' �s�����D
    appRecord = appRecord & tws.Range(Constant.New_Vehicle_Owner_ID_Range).Text & ","               ' �s�����D�����Ҧr��
    appRecord = appRecord & tws.Range(Constant.New_Vehicle_Type_Range).Text & ","                   ' �s�����O
    appRecord = appRecord & tws.Range(Constant.New_Vehicle_Plate_ID_Range).Text & ","               ' �s���P�Ӹ��X
    appRecord = appRecord & tws.Range(Constant.New_Vehicle_Engine_ID_Range).Text & ","              ' �s������/�����X
    appRecord = appRecord & tws.Range(Constant.New_Vehicle_Registration_Date_Range).Text & ","      ' �s����P���
    appRecord = appRecord & "C,"                                                                    ' �㨮�h�|�`�� C
    appRecord = appRecord & tws.Range(Constant.Old_Vehicle_Owner_Name_Range).Text & ","             ' �¨����D
    appRecord = appRecord & tws.Range(Constant.Old_Vehicle_Owner_ID_Range).Text & ","               ' �¨����D�����Ҧr��
    appRecord = appRecord & tws.Range(Constant.Vehicle_Owner_Relation_Range).Text & ","             ' �s�¨��D���Y
    appRecord = appRecord & tws.Range(Constant.Old_Vehicle_Type_Range).Text & ","                   ' �¨����O
    appRecord = appRecord & tws.Range(Constant.Old_Vehicle_Plate_ID_Range).Text & ","               ' �¨��P�Ӹ��X
    If tws.Range(Constant.Old_Vehicle_Body_ID_Range).Text <> "" Then                                ' �¨��Y�������X�B�e�����X�A�Y�S���e�����X
        appRecord = appRecord & tws.Range(Constant.Old_Vehicle_Body_ID_Range).Text & ","
    Else
        appRecord = appRecord & tws.Range(Constant.Old_Vehicle_Engine_ID_Range).Text & ","
    End If
    appRecord = appRecord & tws.Range(Constant.Old_Vehicle_Factory_Date_Range).Text & ","           ' �¨��X�t���
    appRecord = appRecord & tws.Range(Constant.Old_Vehicle_Registration_Date_Range).Text & ","      ' �¨��n�O���
    appRecord = appRecord & tws.Range(Constant.Old_Vehicle_Recycle_Control_ID_Range).Text & ","     ' �¨��^���ި��p�渹�X
    appRecord = appRecord & tws.Range(Constant.Old_Vehicle_Customs_Date_Range).Text & ","           ' �¨��X�f������
    appRecord = appRecord & tws.Range(Constant.Old_Vehicle_Recycle_Date_Range).Text & ","           ' �¨��^�����
    appRecord = appRecord & tws.Range(Constant.Old_Vehicle_Scrapped_Date_Range).Text                ' �¨����o���
    
    targetStream.WriteLine (appRecord)
End Sub
'
'   �ˮ֪�W�� naming convention : �ץ�s��_�~�P_�g�P��_�����X_���j���f���|.xlsx
'
Function Extract_Case_ID(ByVal sourceFileName As String) As String
    Dim firstUnderscorePos  As Integer
    Dim strCaseID           As String
    
    firstUnderscorePos = InStr(1, sourceFileName, "_", vbTextCompare)
    strCaseID = Left(sourceFileName, firstUnderscorePos - 1)

    Extract_Case_ID = strCaseID
End Function
'
'   Traverses the specified folder and returns the number of Excel files within it
'
Public Function Get_Number_Of_Excel_Files(targetFolder As String) As Integer
    Dim fileName As String
    Dim numberOfFiles As Integer
    
    numberOfFile = 0
    fileName = Dir(targetFolder & "\*.xls?")
    Do While fileName <> ""
        numberOfFile = numberOfFile + 1
        fileName = Dir()
    Loop
    Get_Number_Of_Excel_Files = numberOfFile
End Function
