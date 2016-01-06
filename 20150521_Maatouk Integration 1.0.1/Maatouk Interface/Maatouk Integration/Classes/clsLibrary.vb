Imports System.Configuration.ConfigurationManager
Imports System.IO

Public Class clsLibrary
    Private oFileDt As DataTable
    Private oCompany As SAPbobsCOM.Company
    Public Function mainFunction() As Boolean
        Dim _retVal As Boolean = True
        Try
            Dim strPath As String = System.Configuration.ConfigurationManager.AppSettings("OpenFolder").ToString
            Dim txtFiles = Directory.EnumerateFiles(strPath, "*.txt")
            If txtFiles.Count() > 0 Then
                connectSAP()
                For Each currentFile As String In txtFiles
                    traceService(currentFile)
                    oFileDt = New DataTable()
                    If currentFile.Length > 0 Then
                        Dim intCol As Integer = 0
                        Dim txtRows() As String
                        Dim fields() As String
                        Dim oDr As DataRow
                        txtRows = System.IO.File.ReadAllLines(currentFile)
                        Dim intRow As Integer = 0
                        For Each txtrow As String In txtRows
                            If intRow = 0 Then
                                fields = txtrow.Split(vbTab)
                                For index As Integer = 0 To fields.Length - 1
                                    oFileDt.Columns.Add(fields(intCol).ToUpper(), GetType(String)).Caption = fields(intCol).ToUpper()
                                    intCol += 1
                                Next
                                Exit For
                            End If
                        Next

                        intRow = 0
                        For Each txtrow As String In txtRows
                            If intRow = 0 Then

                            ElseIf intRow > 0 Then
                                fields = txtrow.Split(vbTab)
                                oDr = oFileDt.NewRow()
                                oDr.ItemArray = fields
                                oFileDt.Rows.Add(oDr)
                            End If
                            intRow = intRow + 1
                        Next
                    End If
                    traceService("Pushing Data")
                    PushToDIAPI(oFileDt)
                    MoveToFolder(currentFile)
                Next
                disconnectSAP()
            End If
            Return _retVal
        Catch ex As Exception
            traceService(ex.ToString())
        Finally
            disconnectSAP()
        End Try
    End Function

    Private Sub PushToDIAPI(ByVal oDT As DataTable)
        Try
            Dim oBP As SAPbobsCOM.BusinessPartners
            If oDT.Rows.Count > 0 Then
                For index = 0 To oDT.Rows.Count - 1
                    Dim strCardCode As String = oDT.Rows(index)("Family Id").ToString()
                    Dim strCustomer As String = String.Empty
                    Dim strFamilyContName As String = ""
                    Dim strPassword As String = ""
                    oBP = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
                    If oBP.GetByKey(strCardCode) Then
                        If oDT.Rows(index)("Paying").ToString = "F" Then
                            oBP.CardName = oDT.Rows(index)("Father")
                            strCustomer = oDT.Rows(index)("Father")
                            strFamilyContName = oDT.Rows(index)("Mother")
                            oBP.Phone1 = oDT.Rows(index)("Fatherdayphone")
                            oBP.Phone2 = oDT.Rows(index)("Father Home Phone")
                            oBP.Cellular = oDT.Rows(index)("Father Cell Phone")
                            oBP.Fax = oDT.Rows(index)("Father Cell Phone 2")
                        Else
                            oBP.CardName = oDT.Rows(index)("Mother")
                            strCustomer = oDT.Rows(index)("Mother")
                            strFamilyContName = oDT.Rows(index)("Father")
                            oBP.Phone1 = oDT.Rows(index)("Motherdayphone")
                            oBP.Phone2 = oDT.Rows(index)("Mother Home Phone")
                            oBP.Cellular = oDT.Rows(index)("Mother Cell Phone")
                            oBP.Fax = oDT.Rows(index)("Mother Cell Phone 2")
                        End If

                        'oBP.ContactPerson = oDt.Rows(index)(3)
                        Dim blnAExist As Boolean = False
                        For intAddress As Integer = 0 To oBP.Addresses.Count - 1
                            oBP.Addresses.SetCurrentLine(intAddress)
                            If oBP.Addresses.AddressName = strCustomer Then
                                blnAExist = True
                                oBP.Addresses.AddressName = strCustomer
                                oBP.Addresses.Street = oDT.Rows(index)("Street")
                                oBP.Addresses.ZipCode = oDT.Rows(index)("Zip")
                                oBP.Addresses.City = oDT.Rows(index)("City")
                                'oBP.Addresses.Country = oDt.Rows(index)("State Country")
                                oBP.Addresses.Country = "BH"
                                oBP.Addresses.GlobalLocationNumber = oDT.Rows(index)("Mailing City")
                                oBP.Addresses.BuildingFloorRoom = oDT.Rows(index)("Mailing Street")
                                oBP.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_BillTo
                            End If
                        Next


                        If Not blnAExist Then
                            If oBP.Addresses.Count > 0 Then
                                oBP.Addresses.Add()
                                oBP.Addresses.SetCurrentLine(oBP.Addresses.Count - 1)
                                oBP.Addresses.AddressName = strCustomer
                                oBP.Addresses.Street = oDT.Rows(index)("Street")
                                oBP.Addresses.ZipCode = oDT.Rows(index)("Zip")
                                oBP.Addresses.City = oDT.Rows(index)("City")
                                'oBP.Addresses.Country = oDt.Rows(index)("State Country")
                                oBP.Addresses.Country = "BH"
                                oBP.Addresses.GlobalLocationNumber = oDT.Rows(index)("Mailing City")
                                oBP.Addresses.BuildingFloorRoom = oDT.Rows(index)("Mailing Street")
                                oBP.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_BillTo
                            End If
                        End If

                        oBP.EmailAddress = oDT.Rows(index)("Guardianemail")
                        oBP.Website = oDT.Rows(index)("Special Conditions Custom")
                        oBP.Password = oDT.Rows(index)("Custody Role")
                        oBP.AdditionalID = oDT.Rows(index)("Pick Up")
                        oBP.CompanyRegistrationNumber = oDT.Rows(index)("Pick Up Contact Mobile")
                        Dim oTes As SAPbobsCOM.Recordset
                        oTes = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oTes.DoQuery("SELECT T0.[IndCode], T0.[IndName], T0.[IndDesc] FROM OOND T0 where IndName='" & oDT.Rows(index)("Section") & "'")
                        If oTes.RecordCount > 0 Then
                            oBP.Industry = oTes.Fields.Item(0).Value
                        End If
                        'oBP.Industry = oDt.Rows(index)("Section")
                        Dim strEmp As String()
                        Dim strName As String = oDT.Rows(index)("Staff Name").ToString.Replace("""", "")
                        strName = strName.Replace("''", "")
                        strEmp = strName.Split(",")
                        Dim strFirstName, strMiddleName, strlatName, strNameCondition As String
                        If strEmp.Length < 1 Then
                            strNameCondition = ""
                        ElseIf strEmp.Length < 2 Then
                            strFirstName = strEmp(0)
                            strNameCondition = " where firstName='" & strFirstName & "'"
                        ElseIf strEmp.Length < 3 Then
                            strFirstName = strEmp(0)
                            strlatName = strEmp(1)
                            strNameCondition = " where firstName='" & strFirstName & "' and lastName='" & strlatName & "'"
                        ElseIf strEmp.Length < 4 Then
                            strFirstName = strEmp(0)
                            strMiddleName = strEmp(1)
                            strlatName = strEmp(2)
                            strNameCondition = " where firstName='" & strFirstName & "' and lastName='" & strlatName & "' and MiddleName='" & strMiddleName & "'"
                        Else
                            strFirstName = strEmp(0)
                            strMiddleName = strEmp(1)
                            strlatName = strEmp(2)
                            strNameCondition = " where firstName='" & strFirstName & "' and lastName='" & strlatName & "' and MiddleName='" & strMiddleName & "'"
                        End If
                        If strNameCondition <> "" Then
                            oTes.DoQuery("SELECT empID FROM OHEM T0  " & strNameCondition)
                            If oTes.RecordCount > 0 Then
                                oBP.DefaultTechnician = oTes.Fields.Item(0).Value
                            End If
                        End If
                        'oBP.DefaultTechnician = oDt.Rows(index)("Staff Name")

                        Dim blnExist As Boolean = False
                        Dim blnFamilyContact As Boolean = False
                        For intContact As Integer = 0 To oBP.ContactEmployees.Count - 1
                            oBP.ContactEmployees.SetCurrentLine(intContact)
                            If oBP.ContactEmployees.Name = oDT.Rows(index)("Student Number") Then
                                blnExist = True

                                oBP.ContactEmployees.Name = oDT.Rows(index)("Student Number") 'Person ID
                                oBP.ContactEmployees.Position = oDT.Rows(index)("Home Room")
                                oBP.ContactEmployees.FirstName = oDT.Rows(index)("First Name") 'Contact FirstName
                                oBP.ContactEmployees.MiddleName = oDT.Rows(index)("Middle Name")
                                oBP.ContactEmployees.LastName = oDT.Rows(index)("Last Name")
                                oBP.ContactEmployees.Remarks1 = oDT.Rows(index)("Grade Level")
                                oBP.ContactEmployees.Title = oDT.Rows(index)("Ethnicity")
                                strPassword = oDT.Rows(index)("Ssn")
                                If strPassword.Length > 7 Then
                                    strPassword = strPassword.Substring(0, 8)
                                Else
                                    strPassword = strPassword
                                End If
                                oBP.ContactEmployees.Password = strPassword ' oDt.Rows(index)("Ssn")
                                If oDT.Rows(index)("Gender") = "M" Then
                                    oBP.ContactEmployees.Gender = SAPbobsCOM.BoGenderTypes.gt_Male
                                Else
                                    oBP.ContactEmployees.Gender = SAPbobsCOM.BoGenderTypes.gt_Female
                                End If

                                oBP.ContactEmployees.DateOfBirth = oDT.Rows(index)("Dob")
                                If oDT.Rows(index)("Paying").ToString = "F" Then
                                    oBP.ContactEmployees.Phone1 = oDT.Rows(index)("Motherdayphone")
                                    oBP.ContactEmployees.Phone2 = oDT.Rows(index)("Mother Home Phone")
                                    oBP.ContactEmployees.MobilePhone = oDT.Rows(index)("Mother Cell Phone")
                                    oBP.ContactEmployees.Fax = oDT.Rows(index)("Mother Cell Phone 2")
                                Else
                                    oBP.ContactEmployees.Phone1 = oDT.Rows(index)("Fatherdayphone")
                                    oBP.ContactEmployees.Phone2 = oDT.Rows(index)("Father Home Phone")
                                    oBP.ContactEmployees.MobilePhone = oDT.Rows(index)("Father Cell Phone")
                                    oBP.ContactEmployees.Fax = oDT.Rows(index)("Father Cell Phone 2")
                                End If
                            End If
                            If oDT.Rows(index)("Paying").ToString = "F" Then
                                If oBP.ContactEmployees.Name = strFamilyContName Then
                                    blnFamilyContact = True
                                    oBP.ContactEmployees.Phone1 = oDT.Rows(index)("Motherdayphone")
                                    oBP.ContactEmployees.Phone2 = oDT.Rows(index)("Mother Home Phone")
                                    oBP.ContactEmployees.MobilePhone = oDT.Rows(index)("Mother Cell Phone")
                                    oBP.ContactEmployees.Fax = oDT.Rows(index)("Mother Cell Phone 2")
                                End If
                            ElseIf oDT.Rows(index)("Paying").ToString = "M" Then
                                If oBP.ContactEmployees.Name = strFamilyContName Then
                                    blnFamilyContact = True
                                    oBP.ContactEmployees.Phone1 = oDT.Rows(index)("Fatherdayphone")
                                    oBP.ContactEmployees.Phone2 = oDT.Rows(index)("Father Home Phone")
                                    oBP.ContactEmployees.MobilePhone = oDT.Rows(index)("Father Cell Phone")
                                    oBP.ContactEmployees.Fax = oDT.Rows(index)("Father Cell Phone 2")
                                End If
                            End If

                        Next

                        If Not blnExist Then
                            If oBP.ContactEmployees.Count > 0 Then
                                oBP.ContactEmployees.Add()
                                oBP.ContactEmployees.SetCurrentLine(oBP.ContactEmployees.Count - 1)

                                oBP.ContactEmployees.Name = oDT.Rows(index)("Student Number") 'Person ID
                                oBP.ContactEmployees.Position = oDT.Rows(index)("Home Room")
                                oBP.ContactEmployees.FirstName = oDT.Rows(index)("First Name") 'Contact FirstName
                                oBP.ContactEmployees.MiddleName = oDT.Rows(index)("Middle Name")
                                oBP.ContactEmployees.LastName = oDT.Rows(index)("Last Name")
                                oBP.ContactEmployees.Remarks1 = oDT.Rows(index)("Grade Level")
                                oBP.ContactEmployees.Title = oDT.Rows(index)("Ethnicity")
                                'oBP.ContactEmployees.Password = oDt.Rows(index)("Ssn")
                                strPassword = oDT.Rows(index)("Ssn")
                                If strPassword.Length > 7 Then
                                    strPassword = strPassword.Substring(0, 8)
                                Else
                                    strPassword = strPassword
                                End If
                                oBP.ContactEmployees.Password = strPassword '
                                If oDT.Rows(index)("Gender") = "M" Then
                                    oBP.ContactEmployees.Gender = SAPbobsCOM.BoGenderTypes.gt_Male
                                Else
                                    oBP.ContactEmployees.Gender = SAPbobsCOM.BoGenderTypes.gt_Female
                                End If

                                oBP.ContactEmployees.DateOfBirth = oDT.Rows(index)("Dob")
                                If oDT.Rows(index)("Paying").ToString = "F" Then
                                    oBP.ContactEmployees.Phone1 = oDT.Rows(index)("Motherdayphone")
                                    oBP.ContactEmployees.Phone2 = oDT.Rows(index)("Mother Home Phone")
                                    oBP.ContactEmployees.MobilePhone = oDT.Rows(index)("Mother Cell Phone")
                                    oBP.ContactEmployees.Fax = oDT.Rows(index)("Mother Cell Phone 2")
                                Else
                                    oBP.ContactEmployees.Phone1 = oDT.Rows(index)("Fatherdayphone")
                                    oBP.ContactEmployees.Phone2 = oDT.Rows(index)("Father Home Phone")
                                    oBP.ContactEmployees.MobilePhone = oDT.Rows(index)("Father Cell Phone")
                                    oBP.ContactEmployees.Fax = oDT.Rows(index)("Father Cell Phone 2")
                                End If


                            Else
                                oBP.ContactEmployees.Name = oDT.Rows(index)("Student Number") 'Person ID
                                oBP.ContactEmployees.Position = oDT.Rows(index)("Home Room")
                                oBP.ContactEmployees.FirstName = oDT.Rows(index)("First Name") 'Contact FirstName
                                oBP.ContactEmployees.MiddleName = oDT.Rows(index)("Middle Name")
                                oBP.ContactEmployees.LastName = oDT.Rows(index)("Last Name")
                                oBP.ContactEmployees.Remarks1 = oDT.Rows(index)("Grade Level")
                                oBP.ContactEmployees.Title = oDT.Rows(index)("Ethnicity")
                                'oBP.ContactEmployees.Password = oDt.Rows(index)("Ssn")
                                strPassword = oDT.Rows(index)("Ssn")
                                If strPassword.Length > 7 Then
                                    strPassword = strPassword.Substring(0, 8)
                                Else
                                    strPassword = strPassword
                                End If
                                oBP.ContactEmployees.Password = strPassword '
                                If oDT.Rows(index)("Gender") = "M" Then
                                    oBP.ContactEmployees.Gender = SAPbobsCOM.BoGenderTypes.gt_Male
                                Else
                                    oBP.ContactEmployees.Gender = SAPbobsCOM.BoGenderTypes.gt_Female
                                End If

                                oBP.ContactEmployees.DateOfBirth = oDT.Rows(index)("Dob")
                                If oDT.Rows(index)("Paying").ToString = "F" Then
                                    oBP.ContactEmployees.Phone1 = oDT.Rows(index)("Motherdayphone")
                                    oBP.ContactEmployees.Phone2 = oDT.Rows(index)("Mother Home Phone")
                                    oBP.ContactEmployees.MobilePhone = oDT.Rows(index)("Mother Cell Phone")
                                    oBP.ContactEmployees.Fax = oDT.Rows(index)("Mother Cell Phone 2")
                                Else
                                    oBP.ContactEmployees.Phone1 = oDT.Rows(index)("Fatherdayphone")
                                    oBP.ContactEmployees.Phone2 = oDT.Rows(index)("Father Home Phone")
                                    oBP.ContactEmployees.MobilePhone = oDT.Rows(index)("Father Cell Phone")
                                    oBP.ContactEmployees.Fax = oDT.Rows(index)("Father Cell Phone 2")
                                End If
                            End If
                        End If

                        If blnFamilyContact = False Then
                            If oDT.Rows(index)("Paying").ToString = "F" Then
                                If 1 = 1 Then
                                    oBP.ContactEmployees.Add()
                                    oBP.ContactEmployees.SetCurrentLine(oBP.ContactEmployees.Count - 1)
                                    oBP.ContactEmployees.Name = strFamilyContName
                                    oBP.ContactEmployees.Phone1 = oDT.Rows(index)("Motherdayphone")
                                    oBP.ContactEmployees.Phone2 = oDT.Rows(index)("Mother Home Phone")
                                    oBP.ContactEmployees.MobilePhone = oDT.Rows(index)("Mother Cell Phone")
                                    oBP.ContactEmployees.Fax = oDT.Rows(index)("Mother Cell Phone 2")
                                End If
                            ElseIf oDT.Rows(index)("Paying").ToString = "M" Then
                                If 1 = 1 Then
                                    oBP.ContactEmployees.Add()
                                    oBP.ContactEmployees.SetCurrentLine(oBP.ContactEmployees.Count - 1)
                                    oBP.ContactEmployees.Name = strFamilyContName
                                    oBP.ContactEmployees.Phone1 = oDT.Rows(index)("Fatherdayphone")
                                    oBP.ContactEmployees.Phone2 = oDT.Rows(index)("Father Home Phone")
                                    oBP.ContactEmployees.MobilePhone = oDT.Rows(index)("Father Cell Phone")
                                    oBP.ContactEmployees.Fax = oDT.Rows(index)("Father Cell Phone 2")
                                End If
                            End If
                        End If

                        If oDT.Rows(index)("Family Rep").ToString() = "1" Then
                            oBP.ContactPerson = oDT.Rows(index)("Student Number")
                        End If

                        Dim intStatus As Integer = oBP.Update()

                        If intStatus <> 0 Then
                            traceService(strCardCode & "-" & strCustomer & "-" & " Error While Update : " & oCompany.GetLastErrorDescription())
                        End If

                    Else

                        oBP.CardCode = strCardCode

                        If oDT.Rows(index)("Paying").ToString = "F" Then
                            oBP.CardName = oDT.Rows(index)("Father")
                            strCustomer = oDT.Rows(index)("Father")
                            strFamilyContName = oDT.Rows(index)("Mother")
                            oBP.Phone1 = oDT.Rows(index)("Fatherdayphone")
                            oBP.Phone2 = oDT.Rows(index)("Father Home Phone")
                            oBP.Cellular = oDT.Rows(index)("Father Cell Phone")
                            oBP.Fax = oDT.Rows(index)("Father Cell Phone 2")
                        Else
                            oBP.CardName = oDT.Rows(index)("Mother")
                            strCustomer = oDT.Rows(index)("Mother")
                            strFamilyContName = oDT.Rows(index)("Father")
                            oBP.Phone1 = oDT.Rows(index)("Motherdayphone")
                            oBP.Phone2 = oDT.Rows(index)("Mother Home Phone")
                            oBP.Cellular = oDT.Rows(index)("Mother Cell Phone")
                            oBP.Fax = oDT.Rows(index)("Mother Cell Phone 2")
                        End If


                        oBP.Addresses.AddressName = strCustomer
                        oBP.Addresses.Street = oDT.Rows(index)("Street")
                        oBP.Addresses.ZipCode = oDT.Rows(index)("Zip")
                        oBP.Addresses.City = oDT.Rows(index)("City")
                        'oBP.Addresses.Country = oDt.Rows(index)("State Country")
                        oBP.Addresses.Country = "BH"
                        oBP.Addresses.GlobalLocationNumber = oDT.Rows(index)("Mailing City")
                        oBP.Addresses.BuildingFloorRoom = oDT.Rows(index)("Mailing Street")
                        oBP.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_BillTo

                        oBP.EmailAddress = oDT.Rows(index)("Guardianemail")
                        oBP.Website = oDT.Rows(index)("Special Conditions Custom")
                        oBP.Password = oDT.Rows(index)("Custody Role")
                        oBP.AdditionalID = oDT.Rows(index)("Pick Up")
                        oBP.CompanyRegistrationNumber = oDT.Rows(index)("Pick Up Contact Mobile")
                        Dim oTes As SAPbobsCOM.Recordset
                        oTes = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oTes.DoQuery("SELECT T0.[IndCode], T0.[IndName], T0.[IndDesc] FROM OOND T0 where IndName='" & oDT.Rows(index)("Section") & "'")
                        If oTes.RecordCount > 0 Then
                            oBP.Industry = oTes.Fields.Item(0).Value
                        End If
                        'oBP.Industry = oDt.Rows(index)("Section")
                        Dim strEmp As String()
                        Dim strName As String = oDT.Rows(index)("Staff Name")
                        strEmp = strName.Split(",")
                        Dim strFirstName, strMiddleName, strlatName, strNameCondition As String
                        If strEmp.Length < 1 Then
                            strNameCondition = ""
                        ElseIf strEmp.Length < 2 Then
                            strFirstName = strEmp(0)
                            strNameCondition = " where firstName='" & strFirstName & "'"
                        ElseIf strEmp.Length < 3 Then
                            strFirstName = strEmp(0)
                            strlatName = strEmp(1)
                            strNameCondition = " where firstName='" & strFirstName & "' and lastName='" & strlatName & "'"
                        ElseIf strEmp.Length < 4 Then
                            strFirstName = strEmp(0)
                            strMiddleName = strEmp(1)
                            strlatName = strEmp(2)
                            strNameCondition = " where firstName='" & strFirstName & "' and lastName='" & strlatName & "' and MiddleName='" & strMiddleName & "'"
                        Else
                            strFirstName = strEmp(0)
                            strMiddleName = strEmp(1)
                            strlatName = strEmp(2)
                            strNameCondition = " where firstName='" & strFirstName & "' and lastName='" & strlatName & "' and MiddleName='" & strMiddleName & "'"
                        End If
                        If strNameCondition <> "" Then
                            oTes.DoQuery("SELECT empID FROM OHEM T0  " & strNameCondition)
                            If oTes.RecordCount > 0 Then
                                oBP.DefaultTechnician = oTes.Fields.Item(0).Value
                            End If
                        End If
                        oBP.ContactEmployees.Name = oDT.Rows(index)("Student Number") 'Person ID
                        oBP.ContactEmployees.Position = oDT.Rows(index)("Home Room")
                        oBP.ContactEmployees.FirstName = oDT.Rows(index)("First Name") 'Contact FirstName
                        oBP.ContactEmployees.MiddleName = oDT.Rows(index)("Middle Name")
                        oBP.ContactEmployees.LastName = oDT.Rows(index)("Last Name")
                        oBP.ContactEmployees.Remarks1 = oDT.Rows(index)("Grade Level")
                        oBP.ContactEmployees.Title = oDT.Rows(index)("Ethnicity")
                        'oBP.ContactEmployees.Password = oDt.Rows(index)("Ssn")
                        strPassword = oDT.Rows(index)("Ssn")
                        If strPassword.Length > 7 Then
                            strPassword = strPassword.Substring(0, 8)
                        Else
                            strPassword = strPassword
                        End If
                        oBP.ContactEmployees.Password = strPassword '
                        If oDT.Rows(index)("Gender") = "M" Then
                            oBP.ContactEmployees.Gender = SAPbobsCOM.BoGenderTypes.gt_Male
                        Else
                            oBP.ContactEmployees.Gender = SAPbobsCOM.BoGenderTypes.gt_Female
                        End If

                        oBP.ContactEmployees.DateOfBirth = oDT.Rows(index)("Dob")
                        If oDT.Rows(index)("Paying").ToString = "F" Then
                            oBP.ContactEmployees.Phone1 = oDT.Rows(index)("Motherdayphone")
                            oBP.ContactEmployees.Phone2 = oDT.Rows(index)("Mother Home Phone")
                            oBP.ContactEmployees.MobilePhone = oDT.Rows(index)("Mother Cell Phone")
                            oBP.ContactEmployees.Fax = oDT.Rows(index)("Mother Cell Phone 2")
                        Else
                            oBP.ContactEmployees.Phone1 = oDT.Rows(index)("Fatherdayphone")
                            oBP.ContactEmployees.Phone2 = oDT.Rows(index)("Father Home Phone")
                            oBP.ContactEmployees.MobilePhone = oDT.Rows(index)("Father Cell Phone")
                            oBP.ContactEmployees.Fax = oDT.Rows(index)("Father Cell Phone 2")
                        End If

                        If oDT.Rows(index)("Family Rep").ToString() = "1" Then
                            oBP.ContactPerson = oDT.Rows(index)("Student Number")
                        End If

                        If 1 = 1 Then
                            If oDT.Rows(index)("Paying").ToString = "F" Then
                                If 1 = 1 Then
                                    oBP.ContactEmployees.Add()
                                    oBP.ContactEmployees.SetCurrentLine(1)
                                    oBP.ContactEmployees.Name = strFamilyContName
                                    oBP.ContactEmployees.Phone1 = oDT.Rows(index)("Motherdayphone")
                                    oBP.ContactEmployees.Phone2 = oDT.Rows(index)("Mother Home Phone")
                                    oBP.ContactEmployees.MobilePhone = oDT.Rows(index)("Mother Cell Phone")
                                    oBP.ContactEmployees.Fax = oDT.Rows(index)("Mother Cell Phone 2")
                                End If
                            ElseIf oDT.Rows(index)("Paying").ToString = "M" Then
                                If 1 = 1 Then
                                    oBP.ContactEmployees.Add()
                                    oBP.ContactEmployees.SetCurrentLine(1)
                                    oBP.ContactEmployees.Name = strFamilyContName
                                    oBP.ContactEmployees.Phone1 = oDT.Rows(index)("Fatherdayphone")
                                    oBP.ContactEmployees.Phone2 = oDT.Rows(index)("Father Home Phone")
                                    oBP.ContactEmployees.MobilePhone = oDT.Rows(index)("Father Cell Phone")
                                    oBP.ContactEmployees.Fax = oDT.Rows(index)("Father Cell Phone 2")
                                End If
                            End If
                        End If

                        Dim intStatus As Integer = oBP.Add()
                        If intStatus <> 0 Then
                            traceService(strCardCode & "-" & strCustomer & "-" & " Error While Add : " & oCompany.GetLastErrorDescription())
                        End If

                    End If
                Next
            End If



        Catch ex As Exception
            traceService(ex.ToString)
        End Try


    End Sub

    Private Sub connectSAP()
        Try
            traceService("Company Connection")
            oCompany = New SAPbobsCOM.Company()
            Dim strCompany As String = System.Configuration.ConfigurationManager.AppSettings("CompanyDB").ToString
            oCompany.CompanyDB = strCompany

            oCompany.UserName = System.Configuration.ConfigurationManager.AppSettings("UserName").ToString
            oCompany.Password = System.Configuration.ConfigurationManager.AppSettings("Password").ToString
            oCompany.Server = System.Configuration.ConfigurationManager.AppSettings("Server").ToString
            oCompany.DbUserName = System.Configuration.ConfigurationManager.AppSettings("DbUserName").ToString
            oCompany.DbPassword = System.Configuration.ConfigurationManager.AppSettings("DbPassword").ToString

            If System.Configuration.ConfigurationManager.AppSettings("DBType").ToString = "2008" Then
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008
            ElseIf System.Configuration.ConfigurationManager.AppSettings("DBType").ToString = "2012" Then
                'oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012
            End If

            oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English
            oCompany.LicenseServer = System.Configuration.ConfigurationManager.AppSettings("LicenseServer").ToString

            If (0 <> oCompany.Connect()) Then
                traceService("Company Connection Failed")
                traceService(oCompany.GetLastErrorDescription())
                Exit Sub
            Else
                traceService("Connected successfully")
            End If
        Catch ex As Exception
            traceService(ex.ToString())
        End Try
    End Sub

    Private Sub disconnectSAP()
        Try
            If oCompany.Connected Then
                oCompany.Disconnect()
            End If
        Catch ex As Exception
            traceService(ex.ToString())
        End Try
    End Sub

    Private Sub MoveToFolder(ByVal strExFile As String)
        Try
            Dim OpenFolder As String
            Dim SuceessFolder As String
            OpenFolder = System.Configuration.ConfigurationManager.AppSettings("OpenFolder").ToString
            SuceessFolder = System.Configuration.ConfigurationManager.AppSettings("SuccessFolder").ToString
            Dim strFileName As String = Path.GetFileName(strExFile)
            If System.IO.File.Exists(strExFile) = True Then
                'System.IO.File.Move(strExFile, SuceessFolder + "\" + strFileName)
                File.Copy(strExFile, SuceessFolder + "\" + strFileName, True)
                File.Delete(strExFile)
                traceService("File : " & strFileName & " File Moved Successfully")
            End If
        Catch ex As Exception
            traceService(ex.ToString)
        End Try
    End Sub

    Private Sub traceService(ByVal content As String)
        Try
            Dim strFile As String = "\BAYAN_SERVICE_" + System.DateTime.Now.ToString("yyyyMMdd") + ".txt"
            Dim strPath As String = System.Windows.Forms.Application.StartupPath.ToString() & strFile

            If Not File.Exists(strPath) Then
                Dim fs As New FileStream(strPath, FileMode.Create, FileAccess.Write)
                Dim sw As New StreamWriter(fs)
                sw.BaseStream.Seek(0, SeekOrigin.[End])
                sw.WriteLine(content)
                sw.Flush()
                sw.Close()
            Else
                Dim fs As New FileStream(strPath, FileMode.Append, FileAccess.Write)
                Dim sw As New StreamWriter(fs)
                sw.BaseStream.Seek(0, SeekOrigin.[End])
                sw.WriteLine(content)
                sw.Flush()
                sw.Close()
            End If
        Catch ex As Exception
            'traceService(ex.ToString)
        End Try
    End Sub

End Class
