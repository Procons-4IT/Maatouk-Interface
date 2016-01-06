Public Class clsSystemForms

    Private oForm As SAPbouiCOM.Form
    Private oExistingItem As SAPbouiCOM.Item
    Private oItem As SAPbouiCOM.Item
    Private oChkBox As SAPbouiCOM.CheckBox
    Private oButton As SAPbouiCOM.Button

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub AddItems(ByVal FormType As Integer, ByVal FormUID As String)
        Select Case FormType

            Case frm_WAREHOUSES

        End Select
    End Sub
End Class
