Public Class clsAttendance
    Dim strsql As String
    Dim objrs As SAPbobsCOM.Recordset
    Dim objrsdetail As SAPbobsCOM.Recordset
    Dim objupdate As SAPbobsCOM.Recordset
    Dim attendanceentry As Integer

    Public Sub Automation_DailyAttendance()
        Create_Attendance_New()
        Update_Attendance_7days()
    End Sub

    Private Sub Create_Attendance_New()
        Try
            Dim location As Integer = 4
            'strsql = "Select Distinct Convert(date,T0.pdate)pdate,T1.Code from [192.168.1.209].timepaq.dbo.ProcessedPunchDetails T0 inner join [@SMPR_OPYP] T1 on Convert(date,T0.pdate) between Convert(date,T1.U_FromDate) and Convert(date,T1.U_ToDate) "
            strsql = "Select Distinct Convert(date,T0.pdate)pdate,T1.Code from (select Distinct convert(Date,Ratedate)[Pdate] from ORTT) T0 inner join [@SMPR_OPYP] T1 on Convert(date,T0.pdate) between Convert(date,T1.U_FromDate) and Convert(date,T1.U_ToDate) "
            strsql += vbCrLf + " where convert(date,T0.Pdate)<convert(date,getdate()) and Convert(date,T0.Pdate)>'20180331' and "
            strsql += vbCrLf + "Convert(date,T0.PDAte) not in (Select Convert(date,U_AttdDate) from [@SMPR_ODAS] where U_AttdDate >'20180331' and U_Location='" & location.ToString & "') "
            objrs = objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objrs.DoQuery(strsql)
            If objrs.RecordCount = 0 Then Exit Sub
            For i As Integer = 0 To objrs.RecordCount - 1
                attendanceentry = Add_DailyAttendance(objrs.Fields.Item("pdate").Value, objrs.Fields.Item("Code").Value, location.ToString)
                If attendanceentry = -1 Then Continue For
                'Attendance Details Update
                strsql = "EXEC [Innova_Attendance_Automation] '" & attendanceentry.ToString & "'"
                objupdate = objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objupdate.DoQuery(strsql)

                'Mail Automation
                strsql = "INNOVA_SAP_SENDMAIL 'ODAS','A','" & attendanceentry.ToString & "'"
                objupdate = objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objupdate.DoQuery(strsql)

                objrs.MoveNext()
            Next
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Update_Attendance_7days()
        Try
            Dim location As Integer = 4
            strsql = "Select Distinct DocEntry,U_AttdDate  from [@SMPR_ODAS] T0 Where Convert(Date,T0.U_AttdDate) between convert(Date,dateadd(dd,-8,getdate())) and convert(Date,dateadd(dd,-1,getdate()))  and U_Location='" & location.ToString & "'"
            strsql += " and U_AttdDate >(select Max(U_Todate) from [@SMPR_OPRC] where U_Process='Y') and T0.U_AttdDate>'20180331'"
            objrs = objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objrs.DoQuery(strsql)
            If objrs.RecordCount = 0 Then Exit Sub
            For i As Integer = 0 To objrs.RecordCount - 1
                attendanceentry = objrs.Fields.Item("DocEntry").Value
                strsql = "EXEC [Innova_Attendance_Automation] '" & attendanceentry.ToString & "'"
                objupdate = objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objupdate.DoQuery(strsql)
                objrs.MoveNext()
            Next
        Catch ex As Exception

        End Try
    End Sub

    Private Function Add_DailyAttendance(ByVal AttnDate As Date, ByVal PayPeriod As String, ByVal Location As String) As Integer
        Try
            Dim oGeneralService As SAPbobsCOM.GeneralService
            Dim oCompanyservice As SAPbobsCOM.CompanyService
            Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
            Dim oGeneralData As SAPbobsCOM.GeneralData
            Dim ochildren As SAPbobsCOM.GeneralDataCollection
            Dim ochild As SAPbobsCOM.GeneralData

            oCompanyservice = objcompany.GetCompanyService
            oGeneralService = oCompanyservice.GetGeneralService("ODAS")
            oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

            oGeneralData.SetProperty("Series", "527")
            'oGeneralData.SetProperty("Status", "O")
            oGeneralData.SetProperty("U_DocDate", AttnDate.Date)
            oGeneralData.SetProperty("U_AttdDate", AttnDate.Date)
            oGeneralData.SetProperty("U_Location", location.ToString)
            oGeneralData.SetProperty("U_PayPerid", PayPeriod.ToString)
            oGeneralData.SetProperty("U_PreByCod", "5")
            oGeneralData.SetProperty("U_PreByNam", "Ashly Susan Benoy")
            oGeneralData.SetProperty("U_Remarks", "Automated From TimePAQ Machine")
            oGeneralData.SetProperty("U_EmpGrpCode", "-1")
            oGeneralData.SetProperty("U_EmpGrpName", "All")
            'oGeneralData.SetProperty("U_EmpType", "")
            'oGeneralData.SetProperty("U_Upload", "")
            'oGeneralData.SetProperty("U_FindType", "")
            'oGeneralData.SetProperty("U_Find", "")

            ochildren = oGeneralData.Child("SMPR_DAS1")
            strsql = "SELECT T0.Empid ,T0.extempno extempid ,T0.firstname+' '+T0.lastname as Empname,T0.dept,ISNULL(T0.Position,'') Position"
            strsql += vbCrLf + " from ohem T0 where T0.Active ='Y' and T0.status=1 and T0.u_location='" & Location.ToString & "' "
            strsql += vbCrLf + " AND CONVERT(DATE,'" & AttnDate.ToString("yyyyMMdd") & "') between isnull(T0.Startdate,'1950-10-09 00:00:00.000') and isnull(T0.termdate,'2080-10-09 00:00:00.000')"
            objrsdetail = objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objrsdetail.DoQuery(strsql)

            For i As Integer = 0 To objrsdetail.RecordCount - 1
                ochild = ochildren.Add

                ochild.SetProperty("U_empID", objrsdetail.Fields.Item("Empid").Value.ToString)
                ochild.SetProperty("U_empName", objrsdetail.Fields.Item("Empname").Value.ToString)
                ochild.SetProperty("U_Dept", objrsdetail.Fields.Item("dept").Value.ToString)
                ochild.SetProperty("U_Designat", objrsdetail.Fields.Item("Position").Value.ToString)
                ochild.SetProperty("U_IDNo", objrsdetail.Fields.Item("extempid").Value.ToString)
                'ochild.SetProperty("U_ShiftCode", objrsdetail.Fields.Item("").Value.ToString)
                'ochild.SetProperty("U_ShiftName", objrsdetail.Fields.Item("").Value.ToString)
                'ochild.SetProperty("U_TimeIn", objrsdetail.Fields.Item("").Value.ToString)
                'ochild.SetProperty("U_TimeOut", objrsdetail.Fields.Item("").Value.ToString)
                'ochild.SetProperty("U_HrsWrk", objrsdetail.Fields.Item("").Value.ToString)
                ochild.SetProperty("U_Friday", "N")
                ochild.SetProperty("U_Holiday", "N")
                ochild.SetProperty("U_Halfday", "N")
                ochild.SetProperty("U_AttStatus", "")
                'ochild.SetProperty("U_HalfStatus", objrsdetail.Fields.Item("").Value.ToString)
                'ochild.SetProperty("U_OTHrs", objrsdetail.Fields.Item("").Value.ToString)
                'ochild.SetProperty("U_NMOTHrs", objrsdetail.Fields.Item("").Value.ToString)
                'ochild.SetProperty("U_WKOffOTHrs", objrsdetail.Fields.Item("").Value.ToString)
                'ochild.SetProperty("U_HLOTHrs", objrsdetail.Fields.Item("").Value.ToString)
                'ochild.SetProperty("U_ANMOTHrs", objrsdetail.Fields.Item("").Value.ToString)
                'ochild.SetProperty("U_AWKOffOTHrs", objrsdetail.Fields.Item("").Value.ToString)
                'ochild.SetProperty("U_HalfDaySt", objrsdetail.Fields.Item("").Value.ToString)
                'ochild.SetProperty("U_HWStatus", objrsdetail.Fields.Item("").Value.ToString)
                'ochild.SetProperty("U_Permission", objrsdetail.Fields.Item("").Value.ToString)
                'ochild.SetProperty("U_OTApproval", objrsdetail.Fields.Item("").Value.ToString)
                'ochild.SetProperty("U_OTapprovedHrs", objrsdetail.Fields.Item("").Value.ToString)
                'ochild.SetProperty("U_AttendDate", objrsdetail.Fields.Item("").Value.ToString)
                'ochild.SetProperty("U_Loccode", objrsdetail.Fields.Item("").Value.ToString)
                'ochild.SetProperty("U_GrpCode", objrsdetail.Fields.Item("").Value.ToString)
                'ochild.SetProperty("U_GrpName", objrsdetail.Fields.Item("").Value.ToString)
                'ochild.SetProperty("U_PPeridCode", objrsdetail.Fields.Item("").Value.ToString)
                'ochild.SetProperty("U_MachHrs", objrsdetail.Fields.Item("").Value.ToString)

                'ochild.SetProperty("", "")
                objrsdetail.MoveNext()
            Next

            oGeneralParams = oGeneralService.Add(oGeneralData)

            Return oGeneralParams.GetProperty("DocEntry")
        Catch ex As Exception
            Return -1
        End Try
    End Function

End Class
