Imports System.IO
Imports System.Diagnostics.Process
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Text.RegularExpressions
Imports System.Text
Imports System.Net
Imports System.Xml
Imports Microsoft.VisualBasic
Imports System
Imports System.Threading

Public Class clsImport
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oStaticText As SAPbouiCOM.StaticText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oGrid As SAPbouiCOM.Grid
    Private dtTemp As SAPbouiCOM.DataTable
    Private dtResult As SAPbouiCOM.DataTable
    Private oMode As SAPbouiCOM.BoFormMode
    Private oItem As SAPbobsCOM.Items
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBase As DocumentType
    Private InvBaseDocNo As String
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False
    Dim strFileName As String
    Dim strSelectedFilepath, sPath, strSelectedFolderPath As String
    Dim dtDatatable As SAPbouiCOM.DataTable
    Dim blnErrorflag As Boolean = False
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

#Region "Methods"
    Private Sub LoadForm()
        oApplication.Utilities.LoadForm(xml_Import, frm_Import)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.DataSources.UserDataSources.Add("path", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oCombobox = oForm.Items.Item("4").Specific
        oCombobox.ValidValues.Add("UDT", "Import All files")
        'oCombobox.ValidValues.Add("SKU", "SKU")
        ' oCombobox.ValidValues.Add("BP", "Business Partner")
        oCombobox.ValidValues.Add("SHP", "Invoice Import")
        oCombobox.ValidValues.Add("ASN", "Receipt Import")
        'oCombobox.ValidValues.Add("ADJ", "Adjustment Import")
        'oCombobox.ValidValues.Add("HOLD", "Hold Import")
        oCombobox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
        oForm.PaneLevel = 1
        AddChooseFromList(oForm)
        oForm.Items.Item("4").DisplayDesc = True
        oEditText = oForm.Items.Item("6").Specific
        oEditText.DataBind.SetBound(True, "", "path")

    End Sub

#Region "Add Choose From List"
    Private Sub AddChooseFromList(ByVal objForm As SAPbouiCOM.Form)
        Try

            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            oCFLs = objForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFL = oCFLs.Item("CFL_2")
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "Status"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "R"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

#End Region

    


#Region "Browse File"

    '*****************************************************************
    'Type               : Procedure    
    'Name               : BrowseFile
    'Parameter          : Form
    'Return Value       : 
    'Author             :  Senthil Kumar B
    'Created Date       : 
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To Browse a  File
    '******************************************************************
    Private Sub BrowseFile(ByVal Form As SAPbouiCOM.Form)
        'ShowFileDialog(Form)
    End Sub
#End Region

#Region "ShowFileDialog"

    '*****************************************************************
    'Type               : Procedure
    'Name               : ShowFileDialog
    'Parameter          :
    'Return Value       :
    'Author             : Senthil Kumar B 
    'Created Date       : 
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To open a File Browser
    '******************************************************************

    Private Sub fillopen()
        Dim mythr As New System.Threading.Thread(AddressOf ShowFileDialog)
        mythr.SetApartmentState(Threading.ApartmentState.STA)
        mythr.Start()
        mythr.Join()

    End Sub

    Private Sub ShowFileDialog()
        Dim oDialogBox As New OpenFileDialog
        Dim strFileName, strMdbFilePath As String
        Dim oEdit As SAPbouiCOM.EditText
        Dim oProcesses() As Process
        Try
            oProcesses = Process.GetProcessesByName("SAP Business One")
            If oProcesses.Length <> 0 Then
                For i As Integer = 0 To oProcesses.Length - 1
                    Dim MyWindow As New WindowWrapper(oProcesses(i).MainWindowHandle)
                    If oDialogBox.ShowDialog(MyWindow) = DialogResult.OK Then
                        strMdbFilePath = oDialogBox.FileName
                        strSelectedFilepath = oDialogBox.FileName
                        strFileName = strSelectedFilepath
                        strSelectedFolderPath = strFileName
                        If strSelectedFolderPath.EndsWith("\") Then
                            strSelectedFolderPath = strSelectedFilepath.Substring(0, strSelectedFolderPath.Length - 1)
                        End If
                    Else
                    End If
                Next
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
        End Try
    End Sub
#End Region


#Region "Write into ErrorLog File"
    Private Sub WriteErrorlog(ByVal aMessage As String, ByVal aPath As String)
        Dim aSw As System.IO.StreamWriter
        If File.Exists(aPath) Then
        End If
        aSW = New StreamWriter(sPath, True)
        aSW.WriteLine(aMessage)
        aSW.Flush()
        aSW.Close()
    End Sub
#End Region

#Region "Import"
    Private Sub Import(ByVal aForm As SAPbouiCOM.Form)
        Dim strvalue, strTime, strFileName1 As String
        Dim stpath As String
        oCombobox = aForm.Items.Item("4").Specific
        strvalue = oCombobox.Selected.Value
        If strvalue = "" Then
            oApplication.Utilities.Message("Select the Document Type", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If
        stpath = oApplication.Utilities.getEdittextvalue(oForm, "6")
        If stpath = "" Then
            oApplication.Utilities.Message("Folder path missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If
        If Directory.Exists(stpath) = False Then
            oApplication.Utilities.Message("Folder does not exists", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If
        strTime = Now.ToShortTimeString.Replace(":", "")
        strFilename1 = Now.Date.ToString("ddMMyyyy")
        strFileName1 = strFileName1 & strTime
        strImportErrorLog = System.Windows.Forms.Application.StartupPath & "\ImportLog"
        If Directory.Exists(strImportErrorLog) = False Then
            Directory.CreateDirectory(strImportErrorLog)
        End If
        strImportErrorLog = strImportErrorLog & "\Import_" & strFileName1 & ".txt"
        Try
            'If ReadImportFiles(aForm) = False Then
            '    Exit Sub
            'End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End Try
        oApplication.Utilities.WriteErrorHeader(strImportErrorLog, "Import Reading files Processing...")

        oApplication.Utilities.WriteErrorHeader(strImportErrorLog, "Import Reading files Process Completed....")
        oApplication.Utilities.WriteErrorHeader(strImportErrorLog, "Document Creation Processing...")
        Select Case strvalue
            Case "ASN"
                oApplication.Utilities.ImportASNFiles(stpath)
                ' oApplication.Utilities.ImportASNSTFiles(stpath)
            Case "ADJ"
                'oApplication.Utilities.ImportADJFiles(stpath)
            Case "SHP"
                oApplication.Utilities.ImportSOFiles(stpath)
                ' oApplication.Utilities.ImportSOTFiles(stpath)
            Case "HOLD"
                'oApplication.Utilities.ImportHOLDFiles(stpath)
            Case "UDT"
                oApplication.Utilities.ImportASNFiles(stpath)
                ' oApplication.Utilities.ImportASNSTFiles(stpath)
                ' oApplication.Utilities.ImportADJFiles(stpath)
                oApplication.Utilities.ImportSOFiles(stpath)
                ' oApplication.Utilities.ImportSOTFiles(stpath)
                'oApplication.Utilities.ImportHOLDFiles(stpath)
        End Select
        oApplication.Utilities.WriteErrorHeader(strImportErrorLog, "Document Creation Process Completed....")
        If 1 = 1 Then
            Dim x As System.Diagnostics.ProcessStartInfo
            x = New System.Diagnostics.ProcessStartInfo
            x.UseShellExecute = True
            sPath = strImportErrorLog ' System.Windows.Forms.Application.StartupPath & "\ImportLog.txt"
            If File.Exists(sPath) Then
                x.FileName = sPath
                System.Diagnostics.Process.Start(x)
                x = Nothing
            End If
        End If
        oApplication.Utilities.Message("Export process completed successfully.....", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
    End Sub
#End Region

#Region "Read Payroll Interface file"


#Region "Read Import files"
    Private Function ReadImportFiles(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Try
            Dim strvalue As String
            Dim stpath, strImpLogFolder As String
            oCombobox = aForm.Items.Item("4").Specific
            strvalue = oCombobox.Selected.Value
            stpath = oApplication.Utilities.getEdittextvalue(oForm, "6")
            strImpLogFolder = System.Windows.Forms.Application.StartupPath & "\ImportLog.txt"
            strImpLogFolder = strImportErrorLog

            If stpath = "" Then
                oApplication.Utilities.Message("Import folder path is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            sPath = System.Windows.Forms.Application.StartupPath & "\test.txt"
            If File.Exists(sPath) Then
                File.Delete(sPath)
            End If
            If validateFolderPaths(stpath, oCombobox.Selected.Value) = False Then
                Return False
            End If


            Select Case oCombobox.Selected.Value
                Case "SHP"
                    readSOImport(stpath & "\Import\XSO_Export", aForm, sPath)
                Case "ASN"
                    readASNImport(stpath & "\Import\XASN_Export", aForm, sPath)
                Case "ADJ"
                    readADJImport(stpath & "\Import\XINV_Export", aForm, sPath)
                Case "HOLD"
                    readHOLImport(stpath & "\Import\XHOL_Export", aForm, sPath)
                Case "UDT"
                    readSOImport(stpath & "\Import\XSO_Export", aForm, sPath)
                    readASNImport(stpath & "\Import\XASN_Export", aForm, sPath)
                    readADJImport(stpath & "\Import\XINV_Export", aForm, sPath)
                    readHOLImport(stpath & "\Import\XHOL_Export", aForm, sPath)
            End Select
            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
        Return True
    End Function
#End Region

#Region "Validate Folder path"
    Private Function validateFolderPaths(ByVal aPath As String, ByVal choice As String) As Boolean
        Dim strFolder As String
        Select Case choice
            Case "SHP"
                strFolder = aPath & "\Import\XSO_Export"
                If Directory.Exists(aPath & "\Import\XSO_Export") = False Then
                    oApplication.Utilities.Message("Folder does not exist: " & strFolder, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            Case "ASN"
                strFolder = aPath & "\Import\XASN_Export"
                If Directory.Exists(strFolder) = False Then
                    oApplication.Utilities.Message("Folder does not exist: " & strFolder, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            Case "ADJ"
                strFolder = aPath & "\Import\XINV_Export"
                If Directory.Exists(strFolder) = False Then
                    oApplication.Utilities.Message("Folder does not exist: " & strFolder, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            Case "HOLD"
                strFolder = aPath & "\Import\XHOL_Export"
                If Directory.Exists(strFolder) = False Then
                    oApplication.Utilities.Message("Folder does not exist: " & strFolder, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            Case "UDT"
                strFolder = aPath & "\Import\XSO_Export"
                If Directory.Exists(aPath & "\Import\XSO_Export") = False Then
                    oApplication.Utilities.Message("Folder does not exist: " & strFolder, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
                strFolder = aPath & "\Import\XASN_Export"
                If Directory.Exists(strFolder) = False Then
                    oApplication.Utilities.Message("Folder does not exist: " & strFolder, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
                strFolder = aPath & "\Import\XINV_Export"
                If Directory.Exists(strFolder) = False Then
                    oApplication.Utilities.Message("Folder does not exist: " & strFolder, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
                strFolder = aPath & "\Import\XHOL_Export"
                If Directory.Exists(strFolder) = False Then
                    oApplication.Utilities.Message("Folder does not exist: " & strFolder, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
        End Select
        Return True
    End Function
#End Region
#Region "Read SO Import"

    Private Sub readSOImport(ByVal aFolderpath As String, ByVal aform As SAPbouiCOM.Form, ByVal aPath As String)
        Dim di As New IO.DirectoryInfo(aFolderpath)
        Dim aryFi As IO.FileInfo() = di.GetFiles("*.csv")
        Dim fi As IO.FileInfo
        Dim strStorekey, strSokey, strSuccessFile, strErrorFile, strsuccessfolder, strErrorfolder, strLineno, strImpDocType, strType, strdate, strOrderKey, strShipdate, strSKU, strQty, strbatch, strmfgdate, strexpdate
        Dim dtShipdate, dtMfrDate, dtExpDate As Date
        Dim sr As IO.StreamReader
        Dim YEAR, MONTH, DAY, DATE1, strFilename, linje As String
        Dim oDelrec As SAPbobsCOM.Recordset
        Try
            oApplication.Utilities.Message("Reading Shipment files...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            oStaticText = aform.Items.Item("9").Specific
            oStaticText.Caption = "Processing Reading Shipment file..."
            oApplication.Utilities.WriteErrorlog("Reading shipment files...", strImportErrorLog)
            strsuccessfolder = aFolderpath
            strsuccessfolder = aFolderpath & "\Success"
            strErrorfolder = aFolderpath & "\Error"
            If Directory.Exists(strsuccessfolder) = False Then
                Directory.CreateDirectory(strsuccessfolder)
            End If
            If Directory.Exists(strErrorfolder) = False Then
                Directory.CreateDirectory(strErrorfolder)
            End If
            For Each fi In aryFi
                strFilename = fi.FullName
                strSuccessFile = strsuccessfolder & "\" & fi.Name
                strErrorFile = strErrorfolder & "\" & fi.Name
                sr = New StreamReader(fi.FullName, System.Text.Encoding.Default) 'IO.File.OpenText(fil)
                sPath = aPath
                Dim strLIneStrin As String()
                Try
                    oApplication.Utilities.WriteErrorlog("Reading Shipment File Processing...File Name : " & fi.Name, sPath)
                    oApplication.Utilities.WriteErrorlog("Reading Shipment File Processing...File Name : " & fi.Name, strImportErrorLog)
                    'oApplication.Utilities.WriteErrorlog("File Name : " & fi.Name, sPath)
                    Dim oRec, oRecUpdate As SAPbobsCOM.Recordset
                    oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRecUpdate = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRec.DoQuery("Select * from [@Z__XSO] where U_Z_FileName='" & fi.Name & "' and U_Z_Imported='N'")
                    If oRec.RecordCount > 0 Then
                        oRec.DoQuery("Delete  from [@Z__XSO] where U_Z_FileName='" & fi.Name & "' and U_Z_Imported='N'")
                    End If
                    Do While (sr.Peek <> -1)
                        linje = ""
                        linje = sr.ReadLine()
                        strLIneStrin = linje.Split(vbTab)
                        If strLIneStrin.Length > 0 Then
                            strStorekey = strLIneStrin.GetValue(0)
                            strSokey = strLIneStrin.GetValue(1)
                            strType = strLIneStrin.GetValue(2)
                            If strType = "R" Then
                                strImpDocType = "R"
                            Else
                                strImpDocType = "INVTRN"

                            End If
                            strOrderKey = strLIneStrin.GetValue(3)
                            strShipdate = strLIneStrin.GetValue(4)
                            strSKU = strLIneStrin.GetValue(5)
                            strQty = strLIneStrin.GetValue(6)
                            strbatch = strLIneStrin.GetValue(7)
                            strmfgdate = strLIneStrin.GetValue(8)
                            strexpdate = strLIneStrin.GetValue(9)
                            strLineno = strLIneStrin.GetValue(10)
                            strdate = strShipdate
                            strdate = strdate.ToString.Replace("-", "")
                            DAY = strdate.Substring(0, 2)
                            MONTH = strdate.Substring(2, 2)
                            YEAR = strdate.Substring(4, 4)
                            DATE1 = DAY & MONTH & YEAR
                            dtShipdate = GetDateTimeValue(DATE1)
                            strdate = strmfgdate
                            If strdate <> "" Then

                                strdate = strdate.ToString.Replace("-", "")
                                DAY = strdate.Substring(0, 2)
                                MONTH = strdate.Substring(2, 2)
                                YEAR = strdate.Substring(4, 4)
                                DATE1 = DAY & MONTH & YEAR
                                dtMfrDate = GetDateTimeValue(DATE1)
                            End If
                            strdate = strexpdate
                            If strdate <> "" Then
                                strdate = strdate.ToString.Replace("-", "")
                                DAY = strdate.Substring(0, 2)
                                MONTH = strdate.Substring(2, 2)
                                YEAR = strdate.Substring(4, 4)
                                DATE1 = DAY & MONTH & YEAR
                                dtExpDate = GetDateTimeValue(DATE1)
                            End If
                            Dim oUsertable As SAPbobsCOM.UserTable
                            Dim strsql, sCode, strUpdateQuery As String
                            strsql = oApplication.Utilities.getMaxCode("@Z__XSO", "CODE")
                            oUsertable = oApplication.Company.UserTables.Item("Z__XSO")
                            oUsertable.Code = strsql
                            oUsertable.Name = strsql & "M"
                            ' oUsertable.UserFields.Fields.Item("U_Z_DocType").Value = "SO"
                            oUsertable.UserFields.Fields.Item("U_Z_StoreKey").Value = strStorekey
                            oUsertable.UserFields.Fields.Item("U_Z_SAPDocKey").Value = strSokey
                            oUsertable.UserFields.Fields.Item("U_Z_Type").Value = strType
                            oUsertable.UserFields.Fields.Item("U_Z_ImpDocType").Value = strImpDocType
                            oUsertable.UserFields.Fields.Item("U_Z_OrderKey").Value = strOrderKey
                            oUsertable.UserFields.Fields.Item("U_Z_Receiptdate").Value = dtShipdate
                            oUsertable.UserFields.Fields.Item("U_Z_SKU").Value = strSKU
                            oUsertable.UserFields.Fields.Item("U_Z_Quantity").Value = CDbl(strQty)
                            oUsertable.UserFields.Fields.Item("U_Z_BatchNo").Value = strbatch
                            oUsertable.UserFields.Fields.Item("U_Z_MfrDate").Value = dtMfrDate
                            oUsertable.UserFields.Fields.Item("U_Z_ExpDate").Value = dtExpDate
                            oUsertable.UserFields.Fields.Item("U_Z_LineNo").Value = strLineno
                            oUsertable.UserFields.Fields.Item("U_Z_FileName").Value = fi.Name
                            oUsertable.UserFields.Fields.Item("U_Z_Imported").Value = "N"
                            oUsertable.UserFields.Fields.Item("U_Z_ImpMethod").Value = "M"
                            If oUsertable.Add <> 0 Then
                                MsgBox(oApplication.Company.GetLastErrorDescription)
                                oApplication.Utilities.WriteErrorlog("Error --> " & oApplication.Company.GetLastErrorDescription & " File Name : " & fi.Name, sPath)
                                oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oDelrec.DoQuery("Delete from [@Z__XSO] where Name like '%M' and U_Z_Filename='" & fi.Name & "'")
                            End If
                        End If
                    Loop
                    oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oDelrec.DoQuery("Update [@Z__XSO] set Name=code where name like '%M' and U_Z_Filename='" & fi.Name & "'")
                    sr.Close()
                    If File.Exists(strSuccessFile) Then
                        File.Delete(strSuccessFile)
                    End If
                    File.Move(fi.FullName, strSuccessFile)
                    oApplication.Utilities.WriteErrorlog("Reading Process Completed: Filename-->" & fi.Name & " Moved to success folder", sPath)
                    oApplication.Utilities.WriteErrorlog("Reading Process Completed: Filename-->" & fi.Name & " Moved to success folder", strImportErrorLog)

                    'Return True
                Catch ex As Exception

                    oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oApplication.Utilities.WriteErrorlog("Reading SO File Failed...File Name : " & fi.Name, strImportErrorLog)
                    oApplication.Utilities.WriteErrorlog("Error -> " & ex.Message, strImportErrorLog)
                    sr.Close()
                    If File.Exists(strErrorFile) Then
                        File.Delete(strErrorFile)
                    End If
                    File.Move(fi.FullName, strErrorFile)
                    oApplication.Utilities.WriteErrorlog("Reading SO file failed: Filename : " & fi.Name & " Moved to Error folder", strImportErrorLog)
                    ' Return False
                End Try
            Next

            oApplication.Utilities.Message("Reading Shipment file completed...", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            oApplication.Utilities.WriteErrorlog("Reading Shipment file completed", strImportErrorLog)
            oStaticText = aform.Items.Item("9").Specific
            oStaticText.Caption = ""
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub readASNImport(ByVal aFolderpath As String, ByVal aform As SAPbouiCOM.Form, ByVal apath As String)
        Dim di As New IO.DirectoryInfo(aFolderpath)
        Dim aryFi As IO.FileInfo() = di.GetFiles("*.csv")
        Dim fi As IO.FileInfo
        Dim strStorekey, Desgfolder, strsokey, strOrderKey, strSuccessFile, strErrorFile, strsuccessfolder, strErrorfolder, strLineno, strImpDocType, strType, strdate, strShipdate, strSKU, strQty, strbatch, strmfgdate, strexpdate, strSusr1, strSur2, strholdcode As String
        Dim dtShipdate, dtMfrDate, dtExpDate As Date
        Dim sr As IO.StreamReader
        Dim YEAR, MONTH, DAY, DATE1, strFilename, linje As String
        Dim oDelrec As SAPbobsCOM.Recordset
        Try
            oApplication.Utilities.Message("Reading ASN files...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            oStaticText = aform.Items.Item("9").Specific
            oStaticText.Caption = "Processing Reading ASN file..."
            oApplication.Utilities.WriteErrorlog("Reading ASN Files...", strImportErrorLog)
            strsuccessfolder = aFolderpath
            strsuccessfolder = aFolderpath & "\Success"
            strErrorfolder = aFolderpath & "\Error"
            If Directory.Exists(strsuccessfolder) = False Then
                Directory.CreateDirectory(strsuccessfolder)
            End If
            If Directory.Exists(strErrorfolder) = False Then
                Directory.CreateDirectory(strErrorfolder)
            End If

            For Each fi In aryFi
                strFilename = fi.FullName
                strSuccessFile = strsuccessfolder & "\" & fi.Name
                strErrorFile = strErrorfolder & "\" & fi.Name
                sr = New StreamReader(fi.FullName, System.Text.Encoding.Default) 'IO.File.OpenText(fil)
                sPath = apath
                'If File.Exists(sPath) Then
                '    File.Delete(sPath)
                'End If
                Dim strLIneStrin As String()
                Try
                    Dim oRec, oRecUpdate As SAPbobsCOM.Recordset
                    oApplication.Utilities.WriteErrorlog("Reading ASN File Processing...File Name : " & fi.Name, sPath)
                    oApplication.Utilities.WriteErrorlog("Reading ASN File Processing...File Name : " & fi.Name, strImportErrorLog)
                    oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRecUpdate = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRec.DoQuery("Select * from [@Z_XASN] where U_Z_FileName='" & fi.Name & "' and U_Z_Imported='N'")
                    If oRec.RecordCount > 0 Then
                        oRec.DoQuery("Delete from [@Z_XASN] where U_Z_FileName='" & fi.Name & "' and U_Z_Imported='N'")
                    End If
                    Do While (sr.Peek <> -1)
                        linje = ""
                        linje = sr.ReadLine()
                        strLIneStrin = linje.Split(vbTab)
                        If strLIneStrin.Length > 0 Then
                            strStorekey = strLIneStrin.GetValue(0)
                            strsokey = strLIneStrin.GetValue(1)
                            strType = strLIneStrin.GetValue(2)
                            If strType = "" Then
                                strImpDocType = ""
                            End If
                            strImpDocType = "ST"
                            Select Case strType.ToUpper
                                Case "NORMAL"
                                    strImpDocType = "GRPO"
                                Case "I"
                                    strImpDocType = "GRPO"
                                Case "RETRUN ORDER"
                                    strImpDocType = "RETURNS"
                                Case "OR"
                                    strImpDocType = "RETURNS"
                                Case "RETURN INVOICE"
                                    strImpDocType = "ARCR"
                                Case "IR"
                                    strImpDocType = "ARCR"
                                Case "TRN"
                                    strImpDocType = "ST"
                                Case "TRS"
                                    strImpDocType = "ST"
                            End Select

                            strShipdate = strLIneStrin.GetValue(3)
                            strSKU = strLIneStrin.GetValue(4)
                            strQty = strLIneStrin.GetValue(5)
                            strbatch = strLIneStrin.GetValue(6)
                            strmfgdate = strLIneStrin.GetValue(7)
                            strexpdate = strLIneStrin.GetValue(8)
                            strSusr1 = strLIneStrin.GetValue(9)
                            strSur2 = strLIneStrin.GetValue(10)
                            strholdcode = strLIneStrin.GetValue(11)
                            strLineno = strLIneStrin.GetValue(12)

                            strdate = strShipdate.Replace("-", "")
                            If strdate <> "" Then
                                DAY = strdate.Substring(0, 2)
                                MONTH = strdate.Substring(2, 2)
                                YEAR = strdate.Substring(4, 4)
                                DATE1 = DAY & MONTH & YEAR
                                dtShipdate = GetDateTimeValue(DATE1)

                            End If

                            strdate = strmfgdate.Replace("-", "")
                            If strdate <> "" Then
                                DAY = strdate.Substring(0, 2)
                                MONTH = strdate.Substring(2, 2)
                                YEAR = strdate.Substring(4, 4)
                                DATE1 = DAY & MONTH & YEAR
                                dtMfrDate = GetDateTimeValue(DATE1)
                            End If

                            strdate = strexpdate.Replace("-", "")
                            If strdate <> "" Then
                                DAY = strdate.Substring(0, 2)
                                MONTH = strdate.Substring(2, 2)
                                YEAR = strdate.Substring(4, 4)
                                DATE1 = DAY & MONTH & YEAR
                                dtExpDate = GetDateTimeValue(DATE1)
                            End If
                            Dim oUsertable As SAPbobsCOM.UserTable
                            Dim strsql As String
                            strsql = oApplication.Utilities.getMaxCode("@Z_XASN", "CODE")
                            oUsertable = oApplication.Company.UserTables.Item("Z_XASN")
                            oUsertable.Code = strsql
                            oUsertable.Name = strsql & "M"
                            'oUsertable.UserFields.Fields.Item("U_Z_DocType").Value = "ASN"
                            oUsertable.UserFields.Fields.Item("U_Z_StoreKey").Value = strStorekey
                            oUsertable.UserFields.Fields.Item("U_Z_SAPDocKey").Value = strsokey
                            oUsertable.UserFields.Fields.Item("U_Z_Type").Value = strType
                            oUsertable.UserFields.Fields.Item("U_Z_ImpDocType").Value = strImpDocType
                            oUsertable.UserFields.Fields.Item("U_Z_Receiptdate").Value = dtShipdate
                            oUsertable.UserFields.Fields.Item("U_Z_SKU").Value = strSKU
                            oUsertable.UserFields.Fields.Item("U_Z_Quantity").Value = CDbl(strQty)
                            oUsertable.UserFields.Fields.Item("U_Z_BatchNo").Value = strbatch
                            oUsertable.UserFields.Fields.Item("U_Z_MfrDate").Value = dtMfrDate
                            oUsertable.UserFields.Fields.Item("U_Z_ExpDate").Value = dtExpDate
                            oUsertable.UserFields.Fields.Item("U_Z_LineNo").Value = strLineno
                            oUsertable.UserFields.Fields.Item("U_Z_Susr").Value = strSusr1
                            oUsertable.UserFields.Fields.Item("U_Z_Susr2").Value = strSur2
                            oUsertable.UserFields.Fields.Item("U_Z_HoldCode").Value = strholdcode
                            oUsertable.UserFields.Fields.Item("U_Z_FileName").Value = fi.Name
                            oUsertable.UserFields.Fields.Item("U_Z_Imported").Value = "N"

                            oUsertable.UserFields.Fields.Item("U_Z_ImpMethod").Value = "M"
                            If oUsertable.Add <> 0 Then
                                MsgBox(oApplication.Company.GetLastErrorDescription)
                                oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oDelrec.DoQuery("Delete from [@Z_XASN] where Name like '%M' and U_Z_Filename='" & fi.Name & "'")
                            End If
                        End If


                    Loop
                    oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oDelrec.DoQuery("Update [@Z_XASN] set Name=code where name like '%M' and U_Z_Filename='" & fi.Name & "'")

                    sr.Close()
                    File.Move(fi.FullName, strSuccessFile)
                    oApplication.Utilities.WriteErrorlog("Reading Process Completed: Filename-->" & fi.Name & " Moved to success folder", sPath)
                    oApplication.Utilities.WriteErrorlog("Reading Process Completed: Filename-->" & fi.Name & " Moved to success folder", strImportErrorLog)

                    'Return True
                Catch ex As Exception
                    oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oApplication.Utilities.WriteErrorlog("Reading ADN File Failed...File Name : " & fi.Name, sPath)
                    oApplication.Utilities.WriteErrorlog("Error -> " & ex.Message, sPath)
                    oApplication.Utilities.WriteErrorlog("Reading ADN file Failed...File Name : " & fi.Name, strImportErrorLog)
                    oApplication.Utilities.WriteErrorlog("Error -> " & ex.Message, strImportErrorLog)
                    sr.Close()
                    If File.Exists(strErrorFile) Then
                        File.Delete(strErrorFile)
                    End If
                    File.Move(fi.FullName, strErrorFile)
                    oApplication.Utilities.WriteErrorlog("Reading SO file failed: Filename : " & fi.Name & " Moved to Error folder", strImportErrorLog)

                    ' Return False
                End Try
            Next
            oApplication.Utilities.Message("Reading ASN Import completed...", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            oApplication.Utilities.WriteErrorlog("Reading ADN File Completed", strImportErrorLog)

            oStaticText = aform.Items.Item("9").Specific
            oStaticText.Caption = ""
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub readADJImport(ByVal aFolderpath As String, ByVal aform As SAPbouiCOM.Form, ByVal apath As String)
        Dim di As New IO.DirectoryInfo(aFolderpath)
        Dim aryFi As IO.FileInfo() = di.GetFiles("*.csv")
        Dim fi As IO.FileInfo
        Dim strStorekey, strsokey, strwhs, strImpDocType, strSuccessFile, strErrorFile, strsuccessfolder, strErrorfolder, strremarks, strType, strdate, strOrderKey, strShipdate, strSKU, strQty, strbatch, strmfgdate, strexpdate, strSusr1, strSur2, strholdcode As String
        Dim dtShipdate, dtMfrDate, dtExpDate As Date
        Dim sr As IO.StreamReader
        Dim YEAR, MONTH, DAY, DATE1, strFilename, linje As String
        Dim oDelrec As SAPbobsCOM.Recordset
        Try
            oApplication.Utilities.Message("Reading ADJ files...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            oStaticText = aform.Items.Item("9").Specific
            oStaticText.Caption = "Processing Reading ADJ file..."
            oApplication.Utilities.WriteErrorlog("Reading ADJ Files processing..", strImportErrorLog)
            strsuccessfolder = aFolderpath
            strsuccessfolder = aFolderpath & "\Success"
            strErrorfolder = aFolderpath & "\Error"
            If Directory.Exists(strsuccessfolder) = False Then
                Directory.CreateDirectory(strsuccessfolder)
            End If
            If Directory.Exists(strErrorfolder) = False Then
                Directory.CreateDirectory(strErrorfolder)
            End If

            For Each fi In aryFi
                strFilename = fi.FullName
                strSuccessFile = strsuccessfolder & "\" & fi.Name
                strErrorFile = strErrorfolder & "\" & fi.Name
                sr = New StreamReader(fi.FullName, System.Text.Encoding.Default) 'IO.File.OpenText(fil)
                sPath = apath 'System.Windows.Forms.Application.StartupPath & "\ImportLog.txt"

                Dim strLIneStrin As String()
                Try
                    'WriteErrorlog("File Name : " & strFilename, sPath)
                    'WriteErrorlog("Import Process Starting.....", sPath)
                    oApplication.Utilities.WriteErrorlog("Reading Adjustment File Processing...File Name : " & fi.Name, sPath)
                    oApplication.Utilities.WriteErrorlog("Reading Adjustment File Processing...File Name : " & fi.Name, strImportErrorLog)
                    Dim oRec, oRecUpdate, oTemp As SAPbobsCOM.Recordset
                    oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oTemp.DoQuery("SELECT T0.[DfltWhs] FROM OADM T0")
                    strwhs = oTemp.Fields.Item(0).Value

                    oRecUpdate = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRec.DoQuery("Select * from [@Z_XADJ] where U_Z_FileName='" & fi.Name & "' and U_Z_Imported='N'")
                    If oRec.RecordCount > 0 Then
                        oRec.DoQuery("Delete from [@Z_XADJ] where U_Z_FileName='" & fi.Name & "' and U_Z_Imported='N'")
                    End If
                    Do While (sr.Peek <> -1)
                        linje = ""
                        linje = sr.ReadLine()
                        strLIneStrin = linje.Split(vbTab)
                        If strLIneStrin.Length > 7 Then
                            strStorekey = strLIneStrin.GetValue(0)
                            strsokey = strLIneStrin.GetValue(1)
                            strSKU = strLIneStrin.GetValue(2)
                            strbatch = strLIneStrin.GetValue(3)
                            strmfgdate = strLIneStrin.GetValue(4)
                            strexpdate = strLIneStrin.GetValue(5)
                            strQty = strLIneStrin.GetValue(6)
                            If strQty.Contains("-") Then
                                strImpDocType = "Goods Issue"
                            Else
                                strImpDocType = "Goods Recipt"
                            End If
                            strremarks = strLIneStrin.GetValue(7)

                            strdate = strmfgdate.Replace("-", "")
                            If strdate <> "" Then
                                DAY = strdate.Substring(0, 2)
                                MONTH = strdate.Substring(2, 2)
                                YEAR = strdate.Substring(4, 4)
                                DATE1 = DAY & MONTH & YEAR
                                dtMfrDate = GetDateTimeValue(DATE1)
                            End If


                            strdate = strexpdate.Replace("-", "")
                            If strdate <> "" Then
                                DAY = strdate.Substring(0, 2)
                                MONTH = strdate.Substring(2, 2)
                                YEAR = strdate.Substring(4, 4)
                                DATE1 = DAY & MONTH & YEAR
                                dtExpDate = GetDateTimeValue(DATE1)
                            End If
                            Dim oUsertable As SAPbobsCOM.UserTable
                            Dim strsql As String
                            strsql = oApplication.Utilities.getMaxCode("@Z_XADJ", "CODE")
                            oUsertable = oApplication.Company.UserTables.Item("Z_XADJ")
                            oUsertable.Code = strsql
                            oUsertable.Name = strsql & "M"
                            oUsertable.UserFields.Fields.Item("U_Z_StoreKey").Value = strStorekey
                            oUsertable.UserFields.Fields.Item("U_Z_Adjkey").Value = strsokey
                            oUsertable.UserFields.Fields.Item("U_Z_ImpDocType").Value = strImpDocType
                            oUsertable.UserFields.Fields.Item("U_Z_SKU").Value = strSKU
                            oUsertable.UserFields.Fields.Item("U_Z_Remarks").Value = strremarks
                            oUsertable.UserFields.Fields.Item("U_Z_BatchNo").Value = strbatch
                            oUsertable.UserFields.Fields.Item("U_Z_Quantity").Value = CDbl(strQty)
                            oUsertable.UserFields.Fields.Item("U_Z_MfrDate").Value = dtMfrDate
                            oUsertable.UserFields.Fields.Item("U_Z_ExpDate").Value = dtExpDate
                            oUsertable.UserFields.Fields.Item("U_Z_Whs").Value = strwhs
                            oUsertable.UserFields.Fields.Item("U_Z_FileName").Value = fi.Name
                            oUsertable.UserFields.Fields.Item("U_Z_Imported").Value = "N"
                            oUsertable.UserFields.Fields.Item("U_Z_ImpMethod").Value = "M"
                            If oUsertable.Add <> 0 Then
                                MsgBox(oApplication.Company.GetLastErrorDescription)
                                oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oDelrec.DoQuery("Delete from [@Z_XADJ] where Name like '%M' and U_Z_Filename='" & fi.Name & "'")
                            End If
                        End If
                    Loop
                    oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oDelrec.DoQuery("Update [@Z_XADJ] set Name=code where name like '%M' and U_Z_Filename='" & fi.Name & "'")
                    sr.Close()
                    File.Move(fi.FullName, strSuccessFile)
                    oApplication.Utilities.WriteErrorlog("Reading Process Completed: Filename-->" & fi.Name & " Moved to success folder", sPath)
                    oApplication.Utilities.WriteErrorlog("Reading Process Completed: Filename-->" & fi.Name & " Moved to success folder", strImportErrorLog)

                    'Return True
                Catch ex As Exception
                    oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oApplication.Utilities.WriteErrorlog("Reading ADJ File Failed...File Name : " & fi.Name, sPath)
                    oApplication.Utilities.WriteErrorlog("Error -> " & ex.Message, sPath)
                    oApplication.Utilities.WriteErrorlog("Reading ADJ File Failed...File Name : " & fi.Name, strImportErrorLog)
                    oApplication.Utilities.WriteErrorlog("Error -> " & ex.Message, strImportErrorLog)
                    sr.Close()
                    If File.Exists(strErrorFile) Then
                        File.Delete(strErrorFile)
                    End If
                    File.Move(fi.FullName, strErrorFile)
                    oApplication.Utilities.WriteErrorlog("Reading SO file failed: Filename : " & fi.Name & " Moved to Error folder", strImportErrorLog)

                    ' Return False
                End Try
            Next
            oApplication.Utilities.Message("Reading Adjustment file completed...", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            oApplication.Utilities.WriteErrorlog("Reading Adjustment file completed", strImportErrorLog)
            oStaticText = aform.Items.Item("9").Specific
            oStaticText.Caption = ""
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub readHOLImport(ByVal aFolderpath As String, ByVal aform As SAPbouiCOM.Form, ByVal apath As String)
        Dim di As New IO.DirectoryInfo(aFolderpath)
        Dim aryFi As IO.FileInfo() = di.GetFiles("*.csv")
        Dim fi As IO.FileInfo
        Dim strStorekey, strsokey, strfrmwhs, strtowhs, strwhs, strImpDocType, strSuccessFile, strErrorFile, strsuccessfolder, strErrorfolder, strremarks, strType, strdate, strOrderKey, strShipdate, strSKU, strQty, strbatch, strmfgdate, strexpdate, strSusr1, strSur2, strholdcode As String
        Dim dtShipdate, dtMfrDate, dtExpDate As Date
        Dim sr As IO.StreamReader
        Dim YEAR, MONTH, DAY, DATE1, strFilename, linje As String
        Dim oDelrec As SAPbobsCOM.Recordset
        Try
            oApplication.Utilities.Message("Reading HOLD files...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            oStaticText = aform.Items.Item("9").Specific
            oStaticText.Caption = "Processing Reading ADJ file..."
            oApplication.Utilities.WriteErrorlog("Reading HOLD Files processing..", strImportErrorLog)
            strsuccessfolder = aFolderpath
            strsuccessfolder = aFolderpath & "\Success"
            strErrorfolder = aFolderpath & "\Error"
            If Directory.Exists(strsuccessfolder) = False Then
                Directory.CreateDirectory(strsuccessfolder)
            End If
            If Directory.Exists(strErrorfolder) = False Then
                Directory.CreateDirectory(strErrorfolder)
            End If

            For Each fi In aryFi
                strFilename = fi.FullName
                strSuccessFile = strsuccessfolder & "\" & fi.Name
                strErrorFile = strErrorfolder & "\" & fi.Name
                sr = New StreamReader(fi.FullName, System.Text.Encoding.Default) 'IO.File.OpenText(fil)
                sPath = apath 'System.Windows.Forms.Application.StartupPath & "\ImportLog.txt"
                Dim strLIneStrin As String()
                Try
                    oApplication.Utilities.WriteErrorlog("Reading HOLD File Processing...File Name : " & fi.Name, sPath)
                    oApplication.Utilities.WriteErrorlog("Reading HOLD File Processing...File Name : " & fi.Name, strImportErrorLog)
                    Dim oRec, oRecUpdate, oTemp As SAPbobsCOM.Recordset
                    oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oTemp.DoQuery("SELECT T0.[DfltWhs] FROM OADM T0")
                    strwhs = oTemp.Fields.Item(0).Value
                    oRecUpdate = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRec.DoQuery("Select * from [@Z_XHOL] where U_Z_FileName='" & fi.Name & "' and U_Z_Imported='N'")
                    If oRec.RecordCount > 0 Then
                        oRec.DoQuery("Delete from [@Z_XHOL] where U_Z_FileName='" & fi.Name & "' and U_Z_Imported='N'")
                    End If
                    Do While (sr.Peek <> -1)
                        linje = ""
                        linje = sr.ReadLine()
                        strLIneStrin = linje.Split(vbTab)
                        If strLIneStrin.Length > 7 Then
                            strfrmwhs = strLIneStrin.GetValue(0)
                            strtowhs = strLIneStrin.GetValue(1)
                            strremarks = strLIneStrin.GetValue(2)
                            strSKU = strLIneStrin.GetValue(3)
                            strbatch = strLIneStrin.GetValue(4)
                            strmfgdate = strLIneStrin.GetValue(5)
                            strexpdate = strLIneStrin.GetValue(6)
                            strQty = strLIneStrin.GetValue(7)

                            strdate = strmfgdate.Replace("-", "")
                            If strdate <> "" Then
                                DAY = strdate.Substring(0, 2)
                                MONTH = strdate.Substring(2, 2)
                                YEAR = strdate.Substring(4, 4)
                                DATE1 = DAY & MONTH & YEAR
                                dtMfrDate = GetDateTimeValue(DATE1)
                            End If
                            strQty = strQty.Replace(".", CompanyDecimalSeprator)
                            strdate = strexpdate.Replace("-", "")
                            If strdate <> "" Then
                                DAY = strdate.Substring(0, 2)
                                MONTH = strdate.Substring(2, 2)
                                YEAR = strdate.Substring(4, 4)
                                DATE1 = DAY & MONTH & YEAR
                                dtExpDate = GetDateTimeValue(DATE1)
                            End If
                            strImpDocType = "ST"
                            Dim oUsertable As SAPbobsCOM.UserTable
                            Dim strsql As String
                            strsql = oApplication.Utilities.getMaxCode("@Z_XHOL", "CODE")
                            oUsertable = oApplication.Company.UserTables.Item("Z_XHOL")
                            oUsertable.Code = strsql
                            oUsertable.Name = strsql & "M"
                            oUsertable.UserFields.Fields.Item("U_Z_FrmWhs").Value = strfrmwhs
                            oUsertable.UserFields.Fields.Item("U_Z_ToWhs").Value = strtowhs
                            oUsertable.UserFields.Fields.Item("U_Z_ImpDocType").Value = strImpDocType
                            oUsertable.UserFields.Fields.Item("U_Z_SKU").Value = strSKU
                            oUsertable.UserFields.Fields.Item("U_Z_Remarks").Value = strremarks
                            oUsertable.UserFields.Fields.Item("U_Z_BatchNo").Value = strbatch
                            oUsertable.UserFields.Fields.Item("U_Z_Quantity").Value = CDbl(strQty)
                            oUsertable.UserFields.Fields.Item("U_Z_MfrDate").Value = dtMfrDate
                            oUsertable.UserFields.Fields.Item("U_Z_ExpDate").Value = dtExpDate

                            oUsertable.UserFields.Fields.Item("U_Z_FileName").Value = fi.Name
                            oUsertable.UserFields.Fields.Item("U_Z_Imported").Value = "N"
                            oUsertable.UserFields.Fields.Item("U_Z_ImpMethod").Value = "M"
                            If oUsertable.Add <> 0 Then
                                MsgBox(oApplication.Company.GetLastErrorDescription)
                                oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oDelrec.DoQuery("Delete from [@Z_XHOL] where Name like '%M' and U_Z_Filename='" & fi.Name & "'")
                            End If
                        End If
                    Loop
                    oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oDelrec.DoQuery("Update [@Z_XHOL] set Name=code where name like '%M' and U_Z_Filename='" & fi.Name & "'")
                    sr.Close()
                    File.Move(fi.FullName, strSuccessFile)
                    oApplication.Utilities.WriteErrorlog("Reading Process Completed: Filename-->" & fi.Name & " Moved to success folder", sPath)
                    oApplication.Utilities.WriteErrorlog("Reading Process Completed: Filename-->" & fi.Name & " Moved to success folder", strImportErrorLog)

                    'Return True
                Catch ex As Exception
                    oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oApplication.Utilities.WriteErrorlog("Reading HOLD File Failed...File Name : " & fi.Name, sPath)
                    oApplication.Utilities.WriteErrorlog("Error -> " & ex.Message, sPath)
                    oApplication.Utilities.WriteErrorlog("Reading HOLD File Failed...File Name : " & fi.Name, strImportErrorLog)
                    oApplication.Utilities.WriteErrorlog("Error -> " & ex.Message, strImportErrorLog)
                    sr.Close()
                    If File.Exists(strErrorFile) Then
                        File.Delete(strErrorFile)
                    End If
                    File.Move(fi.FullName, strErrorFile)
                    oApplication.Utilities.WriteErrorlog("Reading HOLD file failed: Filename : " & fi.Name & " Moved to Error folder", strImportErrorLog)

                    ' Return False
                End Try
            Next
            oApplication.Utilities.Message("Reading HOLD file completed...", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            oApplication.Utilities.WriteErrorlog("Reading HOLD file completed", strImportErrorLog)
            oStaticText = aform.Items.Item("9").Specific
            oStaticText.Caption = ""
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub


    Private Function ReaddatFile(ByVal strfilename As String, ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim sr As IO.StreamReader
        Dim linje, fil As String
        Dim strAccount, strDes, strAmount, strProject, strprofit, strtype, strRef1, strDay, strMonth, strYear As String
        Dim strCredit, strDebig, strDetails, strRef As String
        Dim dblCredit, dblDebit As Double
        Dim dtDocDate As Date
        Dim blnNegative As Boolean
        Dim oRs As SAPbobsCOM.Recordset
        Dim intReadstatus, intstart, introw As Integer
        Dim YEAR, MONTH, DAY, DATE1 As String
        Dim DTDATE As Date
        Dim amount As Double
        Dim strDate, strTime, strRoaster, strRec, strGreen, strRosted, strRosting, strLoss, strSilo As String
        intReadstatus = 0
        fil = strfilename
        oRs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        sr = New StreamReader(fil, System.Text.Encoding.Default) 'IO.File.OpenText(fil)
        sPath = System.Windows.Forms.Application.StartupPath & "/ImportErrorLog.txt"
        If File.Exists(sPath) Then
            File.Delete(sPath)
        End If
        blnErrorflag = False
        introw = 0
        Dim oStatic As SAPbouiCOM.StaticText
        oStatic = aForm.Items.Item("9").Specific
        oStatic.Caption = "Processing..."
        Try
            Dim oUsertable As SAPbobsCOM.UserTable
            oUsertable = oApplication.Company.UserTables.Item("Z_IRC1")
            Do While (sr.Peek <> -1)
                linje = ""
                linje = sr.ReadLine()
                oStatic = aForm.Items.Item("9").Specific
                oStatic.Caption = "Processing..."
                If linje.Length > 70 Then
                    If linje.StartsWith("|") = False And linje <> "" Then
                        If linje.Length() > 10 Then
                            strDate = linje.Substring(0, 10)
                        End If
                        strDetails = ""
                        If linje.Length > 20 Then
                            strTime = linje.Substring(11, 8)
                            strTime = strTime.Trim()
                        End If
                        If strDate.Length >= 10 Then
                            If strDate <> "" Then
                                DAY = strDate.Substring(0, 2)
                                MONTH = strDate.Substring(3, 2)
                                YEAR = strDate.Substring(6, 4)
                                DATE1 = DAY & MONTH & YEAR
                                dtDocDate = GetDateTimeValue(DATE1)
                            End If
                        End If
                        strRoaster = linje.Substring(21, 8)
                        strRec = linje.Substring(28, 4)
                        strGreen = linje.Substring(32, 10)
                        strRosted = linje.Substring(42, 11)
                        strRosting = linje.Substring(53, 11)
                        strLoss = linje.Substring(65, 7)
                        strSilo = linje.Substring(72, 4)
                        Dim strsql, sCode, strUpdateQuery As String
                        'Dim oSt As String = "Select * from [@Z_IRC1] where U_Z_Date='" & strDate.Trim & "' and U_Z_Time='" & strTime.Trim & "' and U_Z_Roster='" & strRoaster.Trim & "' and U_Z_Rec='" & strRec.Trim & "'"
                        'oSt = oSt & " and U_Z_Green='" & strGreen.Trim & "' and U_Z_Rosted='" & strRosted.Trim & "' and U_Z_Rosting='" & strRosting.Trim & "' and U_Z_Loss='" & strLoss.Trim & "' and U_Z_Silo='" & strSilo.Trim & "'"

                        Dim oSt As String = "Select * from [@Z_IRC1] where U_Z_Date='" & strDate.Trim & "' and U_Z_Time='" & strTime.Trim & "' and U_Z_Roster='" & strRoaster.Trim & "'"
                        oSt = oSt & "  and U_Z_Loss='" & strLoss.Trim & "' and U_Z_Silo='" & strSilo.Trim & "'"
                        Dim otemp As SAPbobsCOM.Recordset
                        otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim strCode As String
                        otemp.DoQuery(oSt)
                        If otemp.RecordCount > 0 Then
                        Else
                            strsql = oApplication.Utilities.getMaxCode("@Z_IRC1", "CODE")
                            oUsertable.Code = strsql
                            oUsertable.Name = strsql & "M"
                            ' oUsertable.UserFields.Fields.Item("U_Z_DocType").Value = "SO"
                            oUsertable.UserFields.Fields.Item("U_Z_Date").Value = strDate.Trim()
                            oUsertable.UserFields.Fields.Item("U_Z_RecDate").Value = dtDocDate
                            oUsertable.UserFields.Fields.Item("U_Z_Time").Value = strTime.Trim()
                            oUsertable.UserFields.Fields.Item("U_Z_Roster").Value = strRoaster.Trim()
                            oUsertable.UserFields.Fields.Item("U_Z_Rec").Value = strRec.Trim()
                            oUsertable.UserFields.Fields.Item("U_Z_Green").Value = strGreen.Trim()
                            oUsertable.UserFields.Fields.Item("U_Z_Rosted").Value = strRosted.Trim
                            oUsertable.UserFields.Fields.Item("U_Z_Rosting").Value = strRosting.Trim
                            oUsertable.UserFields.Fields.Item("U_Z_Loss").Value = strLoss.Trim
                            oUsertable.UserFields.Fields.Item("U_Z_Silo").Value = strSilo.Trim
                            oUsertable.UserFields.Fields.Item("U_Z_Exported").Value = "No"
                            oUsertable.UserFields.Fields.Item("U_Z_PONo").Value = ""
                            oUsertable.UserFields.Fields.Item("U_Z_POEntry").Value = ""
                            oUsertable.UserFields.Fields.Item("U_Z_GreenWasLoss").Value = 0
                            oUsertable.UserFields.Fields.Item("U_Z_GreenWasLossP").Value = 0
                            If oUsertable.Add <> 0 Then
                                MsgBox(oApplication.Company.GetLastErrorDescription)
                                '   oApplication.Utilities.WriteErrorlog("Error --> " & oApplication.Company.GetLastErrorDescription & " File Name : " & fi.Name, sPath)
                            End If
                        End If
                    End If
                End If
                introw = introw + 1
            Loop
            oStatic = aForm.Items.Item("9").Specific
            oStatic.Caption = "Import Process Completed..."
            sr.Close()
            Return True
        Catch ex As Exception
            Return False
        End Try
        oStatic = aForm.Items.Item("9").Specific
        oStatic.Caption = "Import Process Completed..."
    End Function


    Private Function DeleteDataFromUDT(ByVal strfilename As String, ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim sr As IO.StreamReader
        Dim linje, fil As String
        Dim strAccount, strDes, strAmount, strProject, strprofit, strtype, strRef1, strDay, strMonth, strYear As String
        Dim strCredit, strDebig, strDetails, strRef As String
        Dim dblCredit, dblDebit As Double
        Dim dtDocDate As Date
        Dim blnNegative As Boolean
        Dim oRs As SAPbobsCOM.Recordset
        Dim intReadstatus, intstart, introw As Integer
        Dim YEAR, MONTH, DAY, DATE1 As String
        Dim DTDATE As Date
        Dim amount As Double
        Dim strDate, strTime, strRoaster, strRec, strGreen, strRosted, strRosting, strLoss, strSilo As String
        intReadstatus = 0
        fil = strfilename
        oRs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        sr = New StreamReader(fil, System.Text.Encoding.Default) 'IO.File.OpenText(fil)
        sPath = System.Windows.Forms.Application.StartupPath & "/ImportErrorLog.txt"
        If File.Exists(sPath) Then
            File.Delete(sPath)
        End If
        blnErrorflag = False
        introw = 0
        Dim oStatic As SAPbouiCOM.StaticText
        oStatic = aForm.Items.Item("9").Specific
        oStatic.Caption = "Processing..."
        Try
            Dim oUsertable As SAPbobsCOM.UserTable
            oUsertable = oApplication.Company.UserTables.Item("Z_IRC1")
            Do While (sr.Peek <> -1)
                linje = ""
                linje = sr.ReadLine()
                oStatic = aForm.Items.Item("9").Specific
                oStatic.Caption = "Processing..."
                If linje.Length > 70 Then
                    If linje.StartsWith("|") = False And linje <> "" Then
                        If linje.Length() > 10 Then
                            strDate = linje.Substring(0, 10)
                        End If
                        strDetails = ""
                        If linje.Length > 20 Then
                            strTime = linje.Substring(11, 8)
                            strTime = strTime.Trim()
                        End If
                        If strDate.Length >= 10 Then
                            If strDate <> "" Then
                                DAY = strDate.Substring(0, 2)
                                MONTH = strDate.Substring(3, 2)
                                YEAR = strDate.Substring(6, 4)
                                DATE1 = DAY & MONTH & YEAR
                                dtDocDate = GetDateTimeValue(DATE1)
                            End If
                        End If
                        strRoaster = linje.Substring(21, 8)
                        strRec = linje.Substring(28, 4)
                        strGreen = linje.Substring(32, 10)
                        strRosted = linje.Substring(42, 11)
                        strRosting = linje.Substring(53, 11)
                        strLoss = linje.Substring(65, 7)
                        strSilo = linje.Substring(72, 4)
                        Dim strsql, sCode, strUpdateQuery As String
                        Dim oTemp As SAPbobsCOM.Recordset
                        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim oSt1 As String = "Delete from [@Z_IRC1] where U_Z_Exported='No' and U_Z_Date='" & strDate.Trim & "' and U_Z_Time='" & strTime.Trim & "' and U_Z_Roster='" & strRoaster.Trim & "' and U_Z_Rec='" & strRec.Trim & "'"
                        oSt1 = oSt1 & " and U_Z_Green='" & strGreen.Trim & "' and U_Z_Rosted='" & strRosted.Trim & "' and U_Z_Rosting='" & strRosting.Trim & "' and U_Z_Loss='" & strLoss.Trim & "' and U_Z_Silo='" & strSilo.Trim & "'"
                        oTemp.DoQuery(oSt1)
                    End If
                End If
                introw = introw + 1
            Loop
            oStatic = aForm.Items.Item("9").Specific
            oStatic.Caption = "Imported Data removed successfully for selected file.."
            sr.Close()
            Return True
        Catch ex As Exception
            Return False
        End Try
        oStatic = aForm.Items.Item("9").Specific
        oStatic.Caption = "Imported Data removed successfully for selected file"
    End Function

#End Region

#Region "Validation"
    Private Function Validation(ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim oUserTable As SAPbobsCOM.UserTable
        Dim strCode, strECode, strESocial, strEname, strETax, strGLAcc, strLineNo As String
        Dim blnPOMissingflag As Boolean = False
        Dim blnPOEntryFlag As Boolean = False
        Try
            aform.Freeze(True)
            strLineNo = ""
            oGrid = aform.Items.Item("15").Specific
            For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                oApplication.Utilities.Message("Processing....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                strCode = oGrid.DataTable.GetValue("Code", intRow)
                If strCode <> "" Then
                    Dim strRec, strWhs, strPONo, strsql, strins As String
                    oGrid = oForm.Items.Item("15").Specific
                    strRec = oGrid.DataTable.GetValue("U_Z_Rec", intRow)
                    strWhs = oGrid.DataTable.GetValue("U_Z_Silo", intRow)
                    strPONo = oGrid.DataTable.GetValue("U_Z_PONo", intRow)
                    Dim otest As SAPbobsCOM.Recordset
                    otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    If strPONo <> "" Then
                        strsql = "SELECT T0.DocEntry,DocNum  FROM OWOR T0 inner join OITM T1 on T1.itemcode=T0.itemcode INNER JOIN OWHS T2 ON T0.Warehouse = T2.WhsCode and isnull(t1.U_Z_Rec,'')='" & strRec & " '  and T2.U_Z_Silo='" & strWhs & "' and T0.DocNum=" & CDbl(strPONo) & "  and T0.Status='R'"
                        otest.DoQuery(strsql)
                        If otest.RecordCount <= 0 Then
                            blnPOMissingflag = True
                        Else
                            oGrid.DataTable.SetValue("U_Z_PONo", intRow, otest.Fields.Item(1).Value)
                            oGrid.DataTable.SetValue("U_Z_POEntry", intRow, otest.Fields.Item(0).Value)
                        End If
                    End If
                    strECode = oGrid.DataTable.GetValue("U_Z_PONo", intRow)
                    If strECode = "" Then
                        If strLineNo = "" Then
                            strLineNo = intRow
                        Else
                            strLineNo = strLineNo & "," & strLineNo
                        End If
                        blnPOMissingflag = True
                    End If
                    If oGrid.DataTable.GetValue("U_Z_POEntry", intRow) = "" Then
                        blnPOEntryFlag = True
                    End If
                    For intLoop As Integer = intRow + 1 To oGrid.DataTable.Rows.Count - 1
                        Dim st1 As String = oGrid.DataTable.GetValue("U_Z_PONo", intLoop)
                        If st1 <> "" And strPONo = st1 Then
                            oApplication.Utilities.Message("Production order already selected...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            oGrid.Columns.Item("U_Z_PONo").Click(intLoop)
                            aform.Freeze(False)
                            Return False
                        End If
                    Next

                End If
            Next
            If blnPOMissingflag = True Then
                If oApplication.SBO_Application.MessageBox("There are lines with no production orders / Invalid production Order. Do you want to continue the process?", , "Contine", "Cancel") = 2 Then
                    aform.Freeze(False)
                    Return False
                End If
            End If
            aform.Freeze(False)
            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
        End Try
    End Function
#End Region

    Private Function AddtoUDT1(ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim oUserTable As SAPbobsCOM.UserTable
        Dim strCode, strECode, strESocial, strEname, strETax, strGLAcc As String
        oGrid = aform.Items.Item("15").Specific
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            oApplication.Utilities.Message("Processing....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            strCode = oGrid.DataTable.GetValue("Code", intRow)
            strECode = oGrid.DataTable.GetValue("U_Z_PONo", intRow)
            oUserTable = oApplication.Company.UserTables.Item("Z_IRC1")
            If oUserTable.GetByKey(strCode) Then
                'strCode = oApplication.Utilities.getMaxCode("@Z_PAY_OEAR", "Code")
                oUserTable.Code = strCode
                oUserTable.Name = strCode
                oUserTable.UserFields.Fields.Item("U_Z_PONo").Value = strECode
                If strECode = "" Then
                    oUserTable.UserFields.Fields.Item("U_Z_POEntry").Value = "" 'oGrid.DataTable.GetValue("U_Z_POEntry", intRow)
                Else
                    oUserTable.UserFields.Fields.Item("U_Z_POEntry").Value = oGrid.DataTable.GetValue("U_Z_POEntry", intRow)

                End If
                oUserTable.UserFields.Fields.Item("U_Z_Rec").Value = oGrid.DataTable.GetValue("U_Z_Rec", intRow)
                oUserTable.UserFields.Fields.Item("U_Z_Green").Value = oGrid.DataTable.GetValue("U_Z_Green", intRow)
                oUserTable.UserFields.Fields.Item("U_Z_Rosted").Value = oGrid.DataTable.GetValue("U_Z_Rosted", intRow)
                oUserTable.UserFields.Fields.Item("U_Z_Rosting").Value = oGrid.DataTable.GetValue("U_Z_Rosting", intRow)

                If oUserTable.Update <> 0 Then
                    oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                Else
                    
                End If
            End If
        Next
        oApplication.Utilities.Message("Operation completed successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        If CreateIssueforProduction(oGrid) = True Then
            Loaddata(aform)
        Else
            Loaddata(aform)
        End If

    End Function
#Region " Issue for Production"
    Private Function CreateIssueforProduction(ByVal aGrid As SAPbouiCOM.Grid) As Boolean
        Dim strSQl As String
        Dim oTest, otest1 As SAPbobsCOM.Recordset
        strSQl = "Select sum(convert(decimal(9,2),U_Z_Rosted)) 'U_Z_Roasted',Sum(Convert(Decimal(9,2),U_Z_Green)) 'Green',U_Z_POEntry,U_Z_RecDate,Sum(Convert(Decimal(9,2),U_Z_Loss)) 'Loss' from [@Z_IRC1] where isnull(U_Z_Exported,'No')='No' and isnull(U_Z_PONo,'')<>'' group by U_Z_POEntry,U_Z_RecDate"
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otest1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTest.DoQuery(strSQl)
        If oApplication.Company.InTransaction() Then
            oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
        End If
        oApplication.Company.StartTransaction()
        oApplication.Utilities.Message("Processing...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Try
            For intRow As Integer = 0 To oTest.RecordCount - 1
                If oApplication.Utilities.createIssueforProduction(oTest.Fields.Item("U_Z_POEntry").Value, oTest.Fields.Item("U_Z_Roasted").Value, oTest.Fields.Item("U_Z_RecDate").Value, oTest.Fields.Item("Green").Value, oTest.Fields.Item("Loss").Value) = False Then
                    If oApplication.Company.InTransaction() Then
                        oApplication.Utilities.Message("Process completed with error...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    End If
                    Return False
                Else
                    otest1.DoQuery("Update [@Z_IRC1] set U_Z_Exported='Yes' where isnull(U_Z_Exported,'No')='No' and isnull(U_Z_POEntry,'')='" & oTest.Fields.Item("U_Z_POEntry").Value & "'")
                End If
                oTest.MoveNext()
            Next
        Catch ex As Exception
            If oApplication.Company.InTransaction() Then
                oApplication.Utilities.Message("Process completed with error..." & ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If
            Return False
        End Try
        oApplication.Utilities.Message("Operation Completed successfully...", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        If oApplication.Company.InTransaction() Then
            oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
        End If
        Return True
    End Function
#End Region

#End Region

#Region "GetDatetimevalue"
    Public Function GetDateTimeValue(ByVal DateString As String) As DateTime
        Dim objBridge As SAPbobsCOM.SBObob
        objBridge = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
        Return objBridge.Format_StringToDate(DateString).Fields.Item(0).Value
    End Function

#End Region

    Private Sub ReadFile()
        Dim r As New BinaryReader(File.Open("C:\IRC1.dat", FileMode.Open))
        Dim l As New List(Of String)
        Dim Mess(8) As Integer 'array used as buffer
        While r.BaseStream.Position < r.BaseStream.Length  'useful property!
            For i As Integer = 0 To 7 'format is 8x16bit and ....

                Mess(i) = r.ReadInt16
                MsgBox(Mess(i))

            Next
            Mess(8) = r.ReadInt32 '... 1x32bit

            '   ListBox1.Items.Add(New Date(Mess(0), Mess(1), Mess(3), Mess(4), Mess(5), Mess(6)).ToString & " " & Mess(8).ToString)
            'Y, M, D, H, m, s, event code12.    
        End While
        '  TextBox1.Text = l.Item(l.Count - 1)
    End Sub
    Private Sub Loaddata(ByVal aform As SAPbouiCOM.Form)
        aform.Freeze(True)
        oGrid = aform.Items.Item("15").Specific
        oGrid.DataTable.ExecuteQuery("Select * from [@Z_IRC1] where isnull(U_Z_Exported,'No')='No'")
        oGrid.Columns.Item("Code").Visible = False
        oGrid.Columns.Item("Name").Visible = False
        oGrid.Columns.Item("U_Z_Date").TitleObject.Caption = "Date"
        oGrid.Columns.Item("U_Z_Date").Editable = False

        oGrid.Columns.Item("U_Z_Time").TitleObject.Caption = "Time"
        oGrid.Columns.Item("U_Z_Time").Editable = False

        oGrid.Columns.Item("U_Z_Roster").TitleObject.Caption = "Roster"
        oGrid.Columns.Item("U_Z_Roster").Editable = False

        oGrid.Columns.Item("U_Z_Rec").TitleObject.Caption = "Rec"
        oGrid.Columns.Item("U_Z_Rec").Editable = True

        oGrid.Columns.Item("U_Z_Green").TitleObject.Caption = "Green Kgs"
        oGrid.Columns.Item("U_Z_Green").Editable = True

        oGrid.Columns.Item("U_Z_Rosted").TitleObject.Caption = "Rosted Kgs"
        oGrid.Columns.Item("U_Z_Rosted").Editable = True

        oGrid.Columns.Item("U_Z_Rosting").TitleObject.Caption = "Rosting Kgs"
        oGrid.Columns.Item("U_Z_Rosting").Editable = True

        oGrid.Columns.Item("U_Z_Loss").TitleObject.Caption = "Rosting Loss %"
        oGrid.Columns.Item("U_Z_Loss").Editable = False

        oGrid.Columns.Item("U_Z_Silo").TitleObject.Caption = "Silo"
        oGrid.Columns.Item("U_Z_Silo").Editable = False
        oGrid.Columns.Item("U_Z_RecDate").TitleObject.Caption = "Received Date"
        oGrid.Columns.Item("U_Z_RecDate").Visible = False
        oGrid.Columns.Item("U_Z_RecDate").Editable = False
        oGrid.Columns.Item("U_Z_Exported").TitleObject.Caption = "Exported"
        oGrid.Columns.Item("U_Z_Exported").Visible = False
        oGrid.Columns.Item("U_Z_PONo").TitleObject.Caption = "Production Order"
        oEditTextColumn = oGrid.Columns.Item("U_Z_PONo")
        ' oEditTextColumn.ChooseFromListUID = "CFL_2"
        ' oEditTextColumn.ChooseFromListAlias = "DocNum"

        oGrid.Columns.Item("U_Z_GreenWasLoss").Visible = False
        oGrid.Columns.Item("U_Z_GreenWasLossP").Visible = False

        oGrid.Columns.Item("U_Z_POEntry").TitleObject.Caption = "Production Order DocEntry"
        oGrid.Columns.Item("U_Z_POEntry").Editable = False
        oEditTextColumn = oGrid.Columns.Item("U_Z_POEntry")
        oEditTextColumn.LinkedObjectType = "202"
        oGrid.AutoResizeColumns()
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            oGrid.RowHeaders.SetText(intRow, intRow + 1)
        Next
        Dim otest As SAPbobsCOM.Recordset
        otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otest.DoQuery("Select * from [@Z_IRC1] where isnull(U_Z_Exported,'No')='No'")
        If otest.RecordCount > 0 Then
            aform.Items.Item("13").Enabled = True
        Else
            oApplication.Utilities.Message("No data found", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            aform.Items.Item("13").Enabled = False
        End If
        aform.Freeze(False)
    End Sub



#End Region

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_Import Then

                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "13" Then
                                    If Validation(oForm) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)

                                'Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                                '    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                '    If pVal.ItemUID = "15" And pVal.ColUID = "U_Z_PONo" Then
                                '        oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                '        Dim strRec, strWhs, strPONo, strsql, strins As String
                                '        oGrid = oForm.Items.Item("15").Specific
                                '        strRec = oGrid.DataTable.GetValue("U_Z_Rec", pVal.Row)
                                '        strWhs = oGrid.DataTable.GetValue("U_Z_Silo", pVal.Row)
                                '        strPONo = oGrid.DataTable.GetValue("U_Z_PONo", pVal.Row)
                                '        Dim otest As SAPbobsCOM.Recordset
                                '        otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                '        If strPONo <> "" Then
                                '            strsql = "SELECT T0.DocEntry,DocNum  FROM OWOR T0 inner join OITM T1 on T1.itemcode=T0.itemcode INNER JOIN OWHS T2 ON T0.Warehouse = T2.WhsCode and isnull(t1.U_Z_Rec,'')='" & strRec & " '  and T2.U_Z_Silo='" & strWhs & "' and T0.DocNum=" & CDbl(strPONo) & "  and T0.Status='R'"
                                '            otest.DoQuery(strsql)
                                '            If otest.RecordCount <= 0 Then
                                '                oGrid.DataTable.SetValue("U_Z_PONo", pVal.Row, "")
                                '                strins = ""
                                '            Else
                                '                oGrid.DataTable.SetValue("U_Z_PONo", pVal.Row, otest.Fields.Item(1).Value)
                                '                oGrid.DataTable.SetValue("U_Z_POEntry", pVal.Row, otest.Fields.Item(0).Value)
                                '                Exit Sub
                                '            End If
                                '        Else
                                '            oGrid.DataTable.SetValue("U_Z_POEntry", pVal.Row, "")
                                '        End If
                                '    End If


                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "15" And pVal.ColUID = "U_Z_PONo" And pVal.CharPressed = 9 Then
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    Dim strRec, strWhs, strPONo, strsql, strins As String
                                    oGrid = oForm.Items.Item("15").Specific

                                    strRec = oGrid.DataTable.GetValue("U_Z_Rec", pVal.Row)
                                    strWhs = oGrid.DataTable.GetValue("U_Z_Silo", pVal.Row)
                                    strPONo = oGrid.DataTable.GetValue("U_Z_PONo", pVal.Row)
                                    Dim otest As SAPbobsCOM.Recordset
                                    otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    If strPONo <> "" Then
                                        strsql = "SELECT T0.DocEntry,DocNum  FROM OWOR T0 inner join OITM T1 on T1.itemcode=T0.itemcode INNER JOIN OWHS T2 ON T0.Warehouse = T2.WhsCode and isnull(t1.U_Z_Rec,'')='" & strRec & " '  and T2.U_Z_Silo='" & strWhs & "' and T0.DocNum=" & CDbl(strPONo) & "  and T0.Status='R'"
                                        otest.DoQuery(strsql)
                                        If otest.RecordCount <= 0 Then
                                            oGrid.DataTable.SetValue("U_Z_PONo", pVal.Row, "")
                                            strins = ""
                                        Else
                                            oGrid.DataTable.SetValue("U_Z_PONo", pVal.Row, otest.Fields.Item(1).Value)
                                            oGrid.DataTable.SetValue("U_Z_POEntry", pVal.Row, otest.Fields.Item(0).Value)
                                            Exit Sub
                                        End If
                                    Else
                                        oGrid.DataTable.SetValue("U_Z_POEntry", pVal.Row, "")
                                    End If
                                    Dim objChooseForm As SAPbouiCOM.Form
                                    Dim objChoose As New clsChooseFromList
                                    clsChooseFromList.ItemUID = pVal.ItemUID
                                    clsChooseFromList.SourceFormUID = FormUID
                                    clsChooseFromList.SourceLabel = 0
                                    clsChooseFromList.CFLChoice = "Training" 'oCombo.Selected.Value
                                    clsChooseFromList.choice = "Bin"
                                    clsChooseFromList.Documentchoice = "Training"
                                    clsChooseFromList.ItemCode = strWhs
                                    clsChooseFromList.BinDescrUID = strRec
                                    clsChooseFromList.sourceColumID = pVal.ColUID
                                    clsChooseFromList.SourceLabel = pVal.Row
                                    oApplication.Utilities.LoadForm("\CFL.xml", frm_ChoosefromList)
                                    objChooseForm = oApplication.SBO_Application.Forms.ActiveForm()
                                    objChoose.databound(objChooseForm)
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "13" Then
                                    If oApplication.SBO_Application.MessageBox("Do you want to create receipt from production?.", , "Yes", "No") = 2 Then
                                        BubbleEvent = False
                                        Exit Sub
                                    Else
                                        AddtoUDT1(oForm)
                                    End If
                                End If
                                If pVal.ItemUID = "12" Then
                                    fillopen()
                                    oEditText = oForm.Items.Item("6").Specific
                                    oEditText.String = strSelectedFilepath
                                ElseIf pVal.ItemUID = "14" Then
                                    If oApplication.Utilities.getEdittextvalue(oForm, "6") = "" Then
                                        oApplication.Utilities.Message("File missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        Exit Sub
                                    End If
                                    If oApplication.SBO_Application.MessageBox("Do you want to Clear the  Imported Details for selected File?", , "Yes", "No") = 2 Then
                                        Exit Sub
                                    End If
                                    If DeleteDataFromUDT(oApplication.Utilities.getEdittextvalue(oForm, "6"), oForm) = True Then
                                    End If

                                ElseIf pVal.ItemUID = "3" Then
                                    If oApplication.Utilities.getEdittextvalue(oForm, "6") = "" Then
                                        oApplication.Utilities.Message("File missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        Exit Sub
                                    End If
                                    If oApplication.SBO_Application.MessageBox("Do you want to import the Details into UDT?", , "Yes", "No") = 2 Then
                                        Exit Sub
                                    End If
                                    If ReaddatFile(oApplication.Utilities.getEdittextvalue(oForm, "6"), oForm) = True Then
                                    End If
                                ElseIf pVal.ItemUID = "8" Then
                                ElseIf pVal.ItemUID = "11" Then
                                    oForm.PaneLevel = oForm.PaneLevel - 1
                                ElseIf pVal.ItemUID = "10" Then
                                    If oForm.PaneLevel = 1 Then
                                        If oApplication.Utilities.getEdittextvalue(oForm, "6") = "" Then
                                            Loaddata(oForm)
                                            oForm.PaneLevel = oForm.PaneLevel + 1
                                        Else
                                            If oApplication.SBO_Application.MessageBox("Make sure that the Import process is completed", , "Continue", "No") = 2 Then
                                                'Loaddata(oForm)
                                                Exit Sub
                                            Else
                                                '  ReaddatFile(oApplication.Utilities.getEdittextvalue(oForm, "6"), oForm)
                                                Loaddata(oForm)
                                                oForm.PaneLevel = oForm.PaneLevel + 1
                                            End If

                                        End If
                                    End If
                                ElseIf pVal.ItemUID = "13" Then
                                    '   oForm.PaneLevel = oForm.PaneLevel + 1
                                    If oForm.PaneLevel = 2 Then
                                        ' Loaddata(oForm)
                                    End If
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                Dim val1, val2, val3 As String
                                Dim sCHFL_ID, val, val4, val5 As String
                                Dim intChoice As Integer
                                Dim codebar As String
                                Try
                                    oCFLEvento = pVal
                                    sCHFL_ID = oCFLEvento.ChooseFromListUID
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    oCFL = oForm.ChooseFromLists.Item(sCHFL_ID)
                                    If (oCFLEvento.BeforeAction = False) Then
                                        Dim oDataTable As SAPbouiCOM.DataTable
                                        oDataTable = oCFLEvento.SelectedObjects
                                        intChoice = 0
                                        If pVal.ItemUID = "15" And pVal.ColUID = "U_Z_PONo" Then
                                            val = oDataTable.GetValue("DocNum", 0)
                                            val1 = oDataTable.GetValue("DocEntry", 0)
                                            oGrid = oForm.Items.Item("15").Specific
                                            If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                                If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                                End If
                                            End If
                                            Try
                                                oGrid.DataTable.SetValue("U_Z_POEntry", pVal.Row, val1)
                                                oGrid.DataTable.SetValue("U_Z_PONo", pVal.Row, val)
                                                '  oGrid.DataTable.SetValue("U_Z_HRPeoCategory", pVal.Row, val2)

                                            Catch ex As Exception

                                            End Try
                                        End If
                                    End If

                                Catch ex As Exception
                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    End If
                                    oForm.Freeze(False)
                                End Try

                        End Select
                End Select
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_Import
                    If pVal.BeforeAction = False Then
                        'oApplication.Utilities.Message("Import process under development", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        'Exit Sub
                        LoadForm()
                    End If
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD) Then
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class
