Module clsModule
    Public objcompany As SAPbobsCOM.Company

    Public Sub Main()
        CompanyConnection()

        Dim objhrposting As New clsHRPosting
        objhrposting.HR_Posting() 'All Postings

        End
    End Sub

    Private Sub CompanyConnection()
        Dim lretcode
        objcompany = New SAPbobsCOM.Company
        objcompany.Server = Getvalue_webconfig("SAPServername")
        objcompany.SLDServer = Getvalue_webconfig("SLDSERVER")
        objcompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2019
        objcompany.DbUserName = Getvalue_webconfig("SQLUserName")
        objcompany.DbPassword = Getvalue_webconfig("SQLPassword")
        objcompany.LicenseServer = Getvalue_webconfig("SAPLicenseName")
        objcompany.CompanyDB = Getvalue_webconfig("database")
        objcompany.UserName = Getvalue_webconfig("SAPUsername")
        objcompany.Password = Getvalue_webconfig("SAPPassword")
        lretcode = objcompany.Connect()

        If lretcode <> 0 Then
            Dim errcode As String = objcompany.GetLastErrorDescription
            MsgBox(objcompany.GetLastErrorDescription)
        Else
            'MsgBox("Company Connected")
        End If
    End Sub

    Public Function Getvalue_webconfig(ByVal key As String) As String
        Try
            Dim strConnectionString As String = Configuration.ConfigurationManager.AppSettings(key)
            Return strConnectionString
        Catch ex As Exception
            MsgBox(ex.ToString)
            Return ""
        End Try
    End Function

End Module
