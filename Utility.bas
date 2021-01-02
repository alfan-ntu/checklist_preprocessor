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
' Get_Checklist_Folder() returns default or stored file of 明細表
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
' Adds a header row in case the user ticks "Add Header"
'
Public Sub Add_Header_Row(targetStream As TextStream)
    Dim hRow    As String
    
    Debug.Print "User ticks Add Header"
    hRow = "檔案編號,"
    hRow = hRow & "檔案名稱,"
    hRow = hRow & "經銷商,"
    hRow = hRow & "承辦人員,"
    hRow = hRow & "收件日期,"
    hRow = hRow & "退稅原因,"
    hRow = hRow & "退稅支票受款人,"
    hRow = hRow & "受款人身分字號,"
    hRow = hRow & "受款銀行,"
    hRow = hRow & "受款銀行代碼,"
    hRow = hRow & "受款銀行分行,"
    hRow = hRow & "受款銀行分行代碼,"
    hRow = hRow & "受款帳號,"
    hRow = hRow & "新車品牌,"
    hRow = hRow & "新車車型,"
    hRow = hRow & "新車出廠年月,"
    hRow = hRow & "舊車品牌,"
    hRow = hRow & "新車車主,"
    hRow = hRow & "新車車主身分證/統一編號,"
    hRow = hRow & "新車車別,"
    hRow = hRow & "新車牌照號碼,"
    hRow = hRow & "新車車身碼,"
    hRow = hRow & "新車領牌日期,"
    hRow = hRow & "備註,"
    hRow = hRow & "舊車車主,"
    hRow = hRow & "舊車車主身分證/統一編號,"
    hRow = hRow & "新舊車車主關係,"
    hRow = hRow & "舊車車別,"
    hRow = hRow & "舊車牌照號碼,"
    hRow = hRow & "舊車車身碼,"
    hRow = hRow & "舊車出廠日期,"
    hRow = hRow & "舊車登記日期,"
    hRow = hRow & "舊車回收管制聯單編號,"
    hRow = hRow & "舊車出口報單日期,"
    hRow = hRow & "舊車回收日期,"
    hRow = hRow & "舊車報廢日期"
    
    targetStream.WriteLine (hRow)
End Sub
'
' Extract_And_Save() opens the sourceWorkbook, extracts data from the worksheet and stores them to the targetStream
'
Public Sub Extract_And_Save(ByVal sourceFileName As String, sourceWorkbook As Workbook, targetStream As TextStream)
    Dim tws             As Worksheet
    Dim appRecord       As String
    Dim tempDate        As String
    Dim bankAccountVer As Boolean
    
    Set tws = sourceWorkbook.ActiveSheet
    '
    ' Note: this is not a reliable way to determine the version of 檢核表
    '
    If tws.Range("N7").Text = "分行代碼" Then
        bankAccountVer = True
    Else
        bankAccountVer = False
    End If
    
    appRecord = Extract_Case_ID(sourceFileName)                                                     ' 檢核表檔案編號
    appRecord = appRecord & ","
    appRecord = appRecord & sourceFileName & ","                                                    ' 檢核表檔案名稱
    appRecord = appRecord & tws.Range(Constant.Dealer_Range).Text & ","                             ' 經銷商
    appRecord = appRecord & tws.Range(Constant.Dealer_Contact_Range).Text & ","                     ' 經銷商承辦人
    tempDate = Validated_Date_Format(tws.Range(Constant.Submit_Date_Range).Text)                    ' 經銷商送件日
    appRecord = appRecord & tempDate & ","
    If bankAccountVer = True Then
        appRecord = appRecord & tws.Range(Constant.Cause_to_Refund_Range_2021a).Text & ","          ' 退稅原因
        appRecord = appRecord & tws.Range(Constant.Cheque_Payee_Range_2021a).Text & ","             ' 退稅受款人
        appRecord = appRecord & tws.Range(Constant.Cheque_Payee_ID_Range_2021a).Text & ","          ' 受款人身份證字號
        appRecord = appRecord & tws.Range(Constant.Bank_Range_2021a).Text & ","                     ' 受款銀行
        appRecord = appRecord & tws.Range(Constant.Bank_Code_Range_2021a).Text & ","                ' 受款銀行代碼
        appRecord = appRecord & tws.Range(Constant.Branch_Range_2021a).Text & ","                   ' 受款銀行分行
        appRecord = appRecord & tws.Range(Constant.Branch_Code_Range_2021a).Text & ","              ' 受款銀行分行代碼
        appRecord = appRecord & tws.Range(Constant.Bank_Account_Range_2021a).Text & ","             ' 受款人銀行帳號
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
    
    appRecord = appRecord & tws.Range(Constant.New_Vehicle_Brand_Range).Text & ","                  ' 新車品牌
    appRecord = appRecord & tws.Range(Constant.New_Vehicle_Model_Range).Text & ","                  ' 新車車型
    tempDate = Validated_Date_Format(tws.Range(Constant.New_Vehicle_Factory_Date_Range).Text)       ' 新車出廠年月
    appRecord = appRecord & tempDate & ","
    appRecord = appRecord & tws.Range(Constant.Old_Vehicle_Brand_Range).Text & ","                  ' 舊車品牌
    appRecord = appRecord & tws.Range(Constant.New_Vehicle_Owner_Name_Range).Text & ","             ' 新車車主
    appRecord = appRecord & tws.Range(Constant.New_Vehicle_Owner_ID_Range).Text & ","               ' 新車車主身份證字號
    appRecord = appRecord & tws.Range(Constant.New_Vehicle_Type_Range).Text & ","                   ' 新車車別
    appRecord = appRecord & tws.Range(Constant.New_Vehicle_Plate_ID_Range).Text & ","               ' 新車牌照號碼
    appRecord = appRecord & tws.Range(Constant.New_Vehicle_Engine_ID_Range).Text & ","              ' 新車引擎/車身碼
    tempDate = Validated_Date_Format(tws.Range(Constant.New_Vehicle_Registration_Date_Range).Text)  ' 新車領牌日期
    appRecord = appRecord & tempDate & ","
    appRecord = appRecord & "C,"                                                                    ' 整車退稅常數 C
    appRecord = appRecord & tws.Range(Constant.Old_Vehicle_Owner_Name_Range).Text & ","             ' 舊車車主
    appRecord = appRecord & tws.Range(Constant.Old_Vehicle_Owner_ID_Range).Text & ","               ' 舊車車主身份證字號
    appRecord = appRecord & tws.Range(Constant.Vehicle_Owner_Relation_Range).Text & ","             ' 新舊車主關係
    appRecord = appRecord & tws.Range(Constant.Old_Vehicle_Type_Range).Text & ","                   ' 舊車車別
    appRecord = appRecord & tws.Range(Constant.Old_Vehicle_Plate_ID_Range).Text & ","               ' 舊車牌照號碼
    If tws.Range(Constant.Old_Vehicle_Body_ID_Range).Text <> "" Then                                ' 舊車若有車身碼、送車身碼，若沒有送引擎碼
        appRecord = appRecord & tws.Range(Constant.Old_Vehicle_Body_ID_Range).Text & ","
    Else
        appRecord = appRecord & tws.Range(Constant.Old_Vehicle_Engine_ID_Range).Text & ","
    End If
    tempDate = Validated_Date_Format(tws.Range(Constant.Old_Vehicle_Factory_Date_Range).Text)       ' 舊車出廠日期
    appRecord = appRecord & tempDate & ","
    tempDate = Validated_Date_Format(tws.Range(Constant.Old_Vehicle_Registration_Date_Range).Text)  ' 舊車登記日期
    appRecord = appRecord & tempDate & ","
    tempDate = Validated_Date_Format(tws.Range(Constant.Old_Vehicle_Recycle_Control_ID_Range).Text) ' 舊車回收管制聯單號碼
    appRecord = appRecord & tempDate & ","
    tempDate = Validated_Date_Format(tws.Range(Constant.Old_Vehicle_Customs_Date_Range).Text)       ' 舊車出口報單日期
    appRecord = appRecord & tempDate & ","
    tempDate = Validated_Date_Format(tws.Range(Constant.Old_Vehicle_Recycle_Date_Range).Text)       ' 舊車回收日期
    appRecord = appRecord & tempDate & ","
    tempDate = Validated_Date_Format(tws.Range(Constant.Old_Vehicle_Scrapped_Date_Range).Text)      ' 舊車報廢日期
    appRecord = appRecord & tempDate
        
    targetStream.WriteLine (appRecord)
End Sub
'
'   檢核表名稱 naming convention : 案件編號_品牌_經銷商_車身碼_中古車貨物稅.xlsx
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
'
' Date of wrong date format is corrected in this subroutine
'   Wrong format:
'       1. yyyy.mm.dd => yyyy/mm/dd
'
Function Validated_Date_Format(dateUnformated As String) As String
    Dim tempDateString As String
    
    Validated_Date_Format = Replace(dateUnformated, ".", "/")
End Function

