Public Class clsHRPosting
    Dim strsql As String = ""
    Dim objrs As SAPbobsCOM.Recordset
    Dim posted_entryno As String
    Dim lretcode

    Public Sub HR_Posting()

        'EmployeeMaster_Creation()
        LoanApplication()
        LoanRepayment_Manual()
        LeaveSettlement()
        PayrollProcess()
        'MsgBox("1")
        Provision_posting_Table()
        'MsgBox("2")
        Provision_posting_JV()
    End Sub

    Private Sub EmployeeMaster_Creation()
        Try
            strsql = " Select Code,Name,DocEntry,U_empid,U_ExtEmpNo,isnull(U_firstNam,'')[firstNam],isnull(U_lastName,'.')[lastName],left(isnull(U_jobtitle,''),20)[jobtitle],isnull(U_position,'')[position],"
            strsql += vbCrLf + " isnull(U_dept,'')[dept],isnull(U_branch,'')[branch],isnull(U_manager,'')[Manager],isnull(U_userid,'')[Userid],isnull(U_slpcode,'')[slpcode]"
            strsql += vbCrLf + " from [@SMPR_OHEM] where U_ExtEmpNo  not in (Select ExtEmpNo from OHEM) order by convert(int,code)"
            objrs = objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objrs.DoQuery(strsql)
            If objrs.RecordCount = 0 Then GoTo UpdateEMPDetails

            For i As Integer = 0 To objrs.RecordCount - 1
                Try
                    Dim Employeemaster As SAPbobsCOM.EmployeesInfo
                    Employeemaster = objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oEmployeesInfo)

                    Employeemaster.ExternalEmployeeNumber = objrs.Fields.Item("U_ExtEmpNo").Value.ToString
                    Employeemaster.FirstName = objrs.Fields.Item("firstNam").Value.ToString
                    Employeemaster.LastName = objrs.Fields.Item("lastName").Value.ToString
                    If objrs.Fields.Item("jobtitle").Value.ToString <> "" Then Employeemaster.JobTitle = objrs.Fields.Item("jobtitle").Value.ToString
                    If (objrs.Fields.Item("position").Value.ToString <> "" And objrs.Fields.Item("position").Value.ToString <> "0") Then Employeemaster.Position = objrs.Fields.Item("position").Value.ToString
                    If objrs.Fields.Item("dept").Value.ToString <> "" Then Employeemaster.Department = objrs.Fields.Item("dept").Value.ToString
                    If objrs.Fields.Item("branch").Value.ToString <> "" Then Employeemaster.Branch = objrs.Fields.Item("branch").Value.ToString
                    If objrs.Fields.Item("Manager").Value.ToString <> "" Then Employeemaster.Manager = objrs.Fields.Item("Manager").Value.ToString
                    If objrs.Fields.Item("Userid").Value.ToString <> "" Then Employeemaster.ApplicationUserID = objrs.Fields.Item("Userid").Value.ToString
                    If objrs.Fields.Item("slpcode").Value.ToString <> "" Then Employeemaster.SalesPersonCode = objrs.Fields.Item("slpcode").Value.ToString

                    lretcode = Employeemaster.Add()
                    If lretcode <> 0 Then
                        status_Update("OHEM", objrs.Fields.Item("U_ExtEmpNo").Value.ToString, 0, objcompany.GetLastErrorDescription, -1)
                        Exit Sub
                    Else
                        objcompany.GetNewObjectCode(posted_entryno)
                        status_Update("OHEM", objrs.Fields.Item("U_ExtEmpNo").Value.ToString, 1, "Success", posted_entryno)
                    End If
                Catch ex As Exception
                    status_Update("OHEM", objrs.Fields.Item("U_ExtEmpNo").Value.ToString, 0, ex.Message.ToString, -1)
                    Exit Sub
                End Try
                objrs.MoveNext()
            Next

UpdateEMPDetails:
            strsql = "EXEC [Innova_HRMS_EmployeeDetailsUpdate_OHEM] "
            objrs = objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objrs.DoQuery(strsql)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub LoanApplication()
        Try
            strsql = "Exec [Innova_HRMS_Posting_LoanApplication]"
            objrs = objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objrs.DoQuery(strsql)
            If objrs.RecordCount > 0 Then
                For i As Integer = 0 To objrs.RecordCount - 1
                    Try
                        If Not objcompany.InTransaction Then objcompany.StartTransaction()
                        'Dim oloanjv As SAPbobsCOM.JournalVouchers
                        'oloanjv = objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalVouchers)

                        Dim oloanjv As SAPbobsCOM.JournalEntries
                        oloanjv = objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)

                        oloanjv.ReferenceDate = objrs.Fields.Item("Date").Value
                        oloanjv.DueDate = objrs.Fields.Item("Date").Value
                        oloanjv.TaxDate = objrs.Fields.Item("Date").Value
                        If objrs.Fields.Item("Transcode").Value.ToString <> "" Then oloanjv.TransactionCode = objrs.Fields.Item("Transcode").Value.ToString

                        If objrs.Fields.Item("Memo").Value.ToString <> "" Then oloanjv.Memo = objrs.Fields.Item("Memo").Value.ToString
                        If objrs.Fields.Item("Narration").Value.ToString <> "" Then oloanjv.UserFields.Fields.Item("U_Narration").Value = objrs.Fields.Item("Narration").Value.ToString

                        If objrs.Fields.Item("Ref1").Value.ToString <> "" Then oloanjv.Reference = objrs.Fields.Item("Ref1").Value.ToString
                        If objrs.Fields.Item("Ref2").Value.ToString <> "" Then oloanjv.Reference2 = objrs.Fields.Item("Ref2").Value.ToString
                        If objrs.Fields.Item("Ref3").Value.ToString <> "" Then oloanjv.Reference3 = objrs.Fields.Item("Ref3").Value.ToString

                        oloanjv.Lines.AccountCode = objrs.Fields.Item("DebitAccount").Value
                        oloanjv.Lines.Debit = objrs.Fields.Item("Amount").Value
                        'If objrs.Fields.Item("Ocrcode1").Value.ToString <> "" Then oloanjv.Lines.CostingCode = objrs.Fields.Item("Ocrcode1").Value
                        'If objrs.Fields.Item("Ocrcode2").Value.ToString <> "" Then oloanjv.Lines.CostingCode2 = objrs.Fields.Item("Ocrcode2").Value
                        'If objrs.Fields.Item("Ocrcode3").Value.ToString <> "" Then oloanjv.Lines.CostingCode3 = objrs.Fields.Item("Ocrcode3").Value
                        'If objrs.Fields.Item("Ocrcode4").Value.ToString <> "" Then oloanjv.Lines.CostingCode4 = objrs.Fields.Item("Ocrcode4").Value
                        'If objrs.Fields.Item("Ocrcode5").Value.ToString <> "" Then oloanjv.Lines.CostingCode5 = objrs.Fields.Item("Ocrcode5").Value
                        oloanjv.Lines.Add()

                        oloanjv.Lines.AccountCode = objrs.Fields.Item("CreditAccount").Value
                        oloanjv.Lines.Credit = objrs.Fields.Item("Amount").Value
                        'If objrs.Fields.Item("Ocrcode1").Value.ToString <> "" Then oloanjv.Lines.CostingCode = objrs.Fields.Item("Ocrcode1").Value
                        'If objrs.Fields.Item("Ocrcode2").Value.ToString <> "" Then oloanjv.Lines.CostingCode2 = objrs.Fields.Item("Ocrcode2").Value
                        'If objrs.Fields.Item("Ocrcode3").Value.ToString <> "" Then oloanjv.Lines.CostingCode3 = objrs.Fields.Item("Ocrcode3").Value
                        'If objrs.Fields.Item("Ocrcode4").Value.ToString <> "" Then oloanjv.Lines.CostingCode4 = objrs.Fields.Item("Ocrcode4").Value
                        'If objrs.Fields.Item("Ocrcode5").Value.ToString <> "" Then oloanjv.Lines.CostingCode5 = objrs.Fields.Item("Ocrcode5").Value
                        oloanjv.Lines.Add()

                        'oloanjv.JournalEntries.Add()

                        lretcode = oloanjv.Add()
                        If lretcode <> 0 Then
                            If objcompany.InTransaction Then objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                            status_Update("OLOA", objrs.Fields.Item("DocEntry").Value.ToString, 0, objcompany.GetLastErrorDescription, -1)
                        Else
                            objcompany.GetNewObjectCode(posted_entryno)
                            status_Update("OLOA", objrs.Fields.Item("DocEntry").Value.ToString, 1, "Success", posted_entryno.ToString)
                            If Update_query("update [@SMPR_OLOA] set U_jeno='" & posted_entryno & "' where Docentry='" & objrs.Fields.Item("DocEntry").Value.ToString & "'") Then
                                If objcompany.InTransaction Then objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                            Else
                                If objcompany.InTransaction Then objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                            End If
                        End If
                    Catch ex As Exception
                        If objcompany.InTransaction Then objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        status_Update("OLOA", objrs.Fields.Item("DocEntry").Value.ToString, 0, objcompany.GetLastErrorDescription, -1)
                    End Try
                    objrs.MoveNext()
                Next
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub LoanRepayment_Manual()
        Try
            strsql = "Exec [Innova_HRMS_Posting_LoanRepayment_Manual]"
            objrs = objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objrs.DoQuery(strsql)
            If objrs.RecordCount > 0 Then
                For i As Integer = 0 To objrs.RecordCount - 1
                    Try
                        If Not objcompany.InTransaction Then objcompany.StartTransaction()
                        'Dim oloanjv As SAPbobsCOM.JournalVouchers
                        'oloanjv = objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalVouchers)

                        Dim oloanjv As SAPbobsCOM.JournalEntries
                        oloanjv = objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)

                        oloanjv.ReferenceDate = objrs.Fields.Item("Date").Value
                        oloanjv.DueDate = objrs.Fields.Item("Date").Value
                        oloanjv.TaxDate = objrs.Fields.Item("Date").Value
                        If objrs.Fields.Item("Transcode").Value.ToString <> "" Then oloanjv.TransactionCode = objrs.Fields.Item("Transcode").Value.ToString

                        If objrs.Fields.Item("Memo").Value.ToString <> "" Then oloanjv.Memo = objrs.Fields.Item("Memo").Value.ToString
                        If objrs.Fields.Item("Narration").Value.ToString <> "" Then oloanjv.UserFields.Fields.Item("U_Narration").Value = objrs.Fields.Item("Narration").Value.ToString

                        If objrs.Fields.Item("Ref1").Value.ToString <> "" Then oloanjv.Reference = objrs.Fields.Item("Ref1").Value.ToString
                        If objrs.Fields.Item("Ref2").Value.ToString <> "" Then oloanjv.Reference2 = objrs.Fields.Item("Ref2").Value.ToString
                        If objrs.Fields.Item("Ref3").Value.ToString <> "" Then oloanjv.Reference3 = objrs.Fields.Item("Ref3").Value.ToString

                        oloanjv.Lines.AccountCode = objrs.Fields.Item("DebitAccount").Value
                        oloanjv.Lines.Debit = objrs.Fields.Item("Amount").Value
                        'If objrs.Fields.Item("Ocrcode1").Value.ToString <> "" Then oloanjv.Lines.CostingCode = objrs.Fields.Item("Ocrcode1").Value
                        'If objrs.Fields.Item("Ocrcode2").Value.ToString <> "" Then oloanjv.Lines.CostingCode2 = objrs.Fields.Item("Ocrcode2").Value
                        'If objrs.Fields.Item("Ocrcode3").Value.ToString <> "" Then oloanjv.Lines.CostingCode3 = objrs.Fields.Item("Ocrcode3").Value
                        'If objrs.Fields.Item("Ocrcode4").Value.ToString <> "" Then oloanjv.Lines.CostingCode4 = objrs.Fields.Item("Ocrcode4").Value
                        'If objrs.Fields.Item("Ocrcode5").Value.ToString <> "" Then oloanjv.Lines.CostingCode5 = objrs.Fields.Item("Ocrcode5").Value
                        oloanjv.Lines.Add()

                        oloanjv.Lines.AccountCode = objrs.Fields.Item("CreditAccount").Value
                        oloanjv.Lines.Credit = objrs.Fields.Item("Amount").Value
                        'If objrs.Fields.Item("Ocrcode1").Value.ToString <> "" Then oloanjv.Lines.CostingCode = objrs.Fields.Item("Ocrcode1").Value
                        'If objrs.Fields.Item("Ocrcode2").Value.ToString <> "" Then oloanjv.Lines.CostingCode2 = objrs.Fields.Item("Ocrcode2").Value
                        'If objrs.Fields.Item("Ocrcode3").Value.ToString <> "" Then oloanjv.Lines.CostingCode3 = objrs.Fields.Item("Ocrcode3").Value
                        'If objrs.Fields.Item("Ocrcode4").Value.ToString <> "" Then oloanjv.Lines.CostingCode4 = objrs.Fields.Item("Ocrcode4").Value
                        'If objrs.Fields.Item("Ocrcode5").Value.ToString <> "" Then oloanjv.Lines.CostingCode5 = objrs.Fields.Item("Ocrcode5").Value
                        oloanjv.Lines.Add()

                        'oloanjv.JournalEntries.Add()

                        lretcode = oloanjv.Add()
                        If lretcode <> 0 Then
                            If objcompany.InTransaction Then objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                            status_Update("LOA1", objrs.Fields.Item("DocEntry").Value.ToString, 0, objcompany.GetLastErrorDescription, -1)
                        Else
                            objcompany.GetNewObjectCode(posted_entryno)
                            status_Update("LOA1", objrs.Fields.Item("DocEntry").Value.ToString, 1, "Success", posted_entryno.ToString)
                            If Update_query("update [@SMPR_LOA1] set U_jeno='" & posted_entryno & "' where Docentry='" & objrs.Fields.Item("DocEntry").Value.ToString & "' and Lineid='" & objrs.Fields.Item("LineId").Value.ToString & "'") Then
                                If objcompany.InTransaction Then objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                            Else
                                If objcompany.InTransaction Then objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                            End If
                        End If
                    Catch ex As Exception
                        If objcompany.InTransaction Then objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        status_Update("LOA1", objrs.Fields.Item("DocEntry").Value.ToString, 0, objcompany.GetLastErrorDescription, -1)
                    End Try
                    objrs.MoveNext()
                Next
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub LeaveSettlement()
        Try
            Dim objrsheader As SAPbobsCOM.Recordset
            objrsheader = objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strsql = "Select Docentry from [@SMPR_OLSE] where Isnull(U_jeno,'')='' and isnull(U_approved,'')='Y' and isnull(Canceled,'')<>'Y'"
            objrsheader.DoQuery(strsql)
            If objrsheader.RecordCount = 0 Then Exit Sub

            For intheader As Integer = 0 To objrsheader.RecordCount - 1

                strsql = "Exec [Innova_HRMS_Posting_Settlement] '" & objrsheader.Fields.Item("Docentry").Value.ToString & "'"
                objrs = objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objrs.DoQuery(strsql)
                If objrs.RecordCount > 0 Then
                    Try
                        If Not objcompany.InTransaction Then objcompany.StartTransaction()

                        'Dim osettlementjv As SAPbobsCOM.JournalVouchers
                        'osettlementjv = objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalVouchers)
                        Dim osettlementjv As SAPbobsCOM.JournalEntries
                        osettlementjv = objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)

                        osettlementjv.ReferenceDate = objrs.Fields.Item("Date").Value
                        osettlementjv.DueDate = objrs.Fields.Item("Date").Value
                        osettlementjv.TaxDate = objrs.Fields.Item("Date").Value
                        If objrs.Fields.Item("Transcode").Value.ToString <> "" Then osettlementjv.TransactionCode = objrs.Fields.Item("Transcode").Value.ToString

                        If objrs.Fields.Item("Memo").Value.ToString <> "" Then osettlementjv.Memo = objrs.Fields.Item("Memo").Value.ToString
                        If objrs.Fields.Item("Narration").Value.ToString <> "" Then osettlementjv.UserFields.Fields.Item("U_Narration").Value = objrs.Fields.Item("Narration").Value.ToString

                        If objrs.Fields.Item("Ref1").Value.ToString <> "" Then osettlementjv.Reference = objrs.Fields.Item("Ref1").Value.ToString
                        If objrs.Fields.Item("Ref2").Value.ToString <> "" Then osettlementjv.Reference2 = objrs.Fields.Item("Ref2").Value.ToString
                        If objrs.Fields.Item("Ref3").Value.ToString <> "" Then osettlementjv.Reference3 = objrs.Fields.Item("Ref3").Value.ToString

                        For i As Integer = 0 To objrs.RecordCount - 1
                            osettlementjv.Lines.AccountCode = objrs.Fields.Item("AcctCode").Value
                            If objrs.Fields.Item("DebitAmount").Value <> 0 Then osettlementjv.Lines.Debit = objrs.Fields.Item("DebitAmount").Value Else osettlementjv.Lines.Credit = objrs.Fields.Item("CreditAmount").Value

                            osettlementjv.Lines.Reference1 = objrs.Fields.Item("Lref1").Value
                            osettlementjv.Lines.Reference2 = objrs.Fields.Item("Lref2").Value
                            osettlementjv.Lines.AdditionalReference = objrs.Fields.Item("Lref3").Value

                            'If objrs.Fields.Item("Ocrcode1").Value.ToString <> "" Then osettlementjv.Lines.CostingCode = objrs.Fields.Item("Ocrcode1").Value
                            'If objrs.Fields.Item("Ocrcode2").Value.ToString <> "" Then osettlementjv.Lines.CostingCode2 = objrs.Fields.Item("Ocrcode2").Value
                            'If objrs.Fields.Item("Ocrcode3").Value.ToString <> "" Then osettlementjv.Lines.CostingCode3 = objrs.Fields.Item("Ocrcode3").Value
                            'If objrs.Fields.Item("Ocrcode4").Value.ToString <> "" Then osettlementjv.Lines.CostingCode4 = objrs.Fields.Item("Ocrcode4").Value
                            'If objrs.Fields.Item("Ocrcode5").Value.ToString <> "" Then osettlementjv.Lines.CostingCode5 = objrs.Fields.Item("Ocrcode5").Value

                            osettlementjv.Lines.Add()
                            objrs.MoveNext()
                        Next

                        'osettlementjv.JournalEntries.Add()

                        lretcode = osettlementjv.Add()
                        If lretcode <> 0 Then
                            If objcompany.InTransaction Then objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                            status_Update("OLSE", objrsheader.Fields.Item("Docentry").Value.ToString, 0, objcompany.GetLastErrorDescription, -1)
                        Else
                            objcompany.GetNewObjectCode(posted_entryno)
                            status_Update("OLSE", objrsheader.Fields.Item("Docentry").Value.ToString, 1, "Success", posted_entryno.ToString)
                            If Update_query("update [@SMPR_OLSE] set U_jeno='" & posted_entryno & "' where Docentry='" & objrsheader.Fields.Item("Docentry").Value.ToString & "'") Then
                                If objcompany.InTransaction Then objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                            Else
                                If objcompany.InTransaction Then objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                            End If
                        End If

                    Catch ex As Exception
                        If objcompany.InTransaction Then objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        status_Update("OLSE", objrsheader.Fields.Item("Docentry").Value.ToString, 0, objcompany.GetLastErrorDescription, -1)
                    End Try

                End If


                objrsheader.MoveNext()
            Next
        Catch ex As Exception

        End Try
    End Sub

    Private Sub PayrollProcess()
        Try
            Dim objrsheader As SAPbobsCOM.Recordset
            objrsheader = objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strsql = "select DocEntry from [@SMPR_OPRC] Where isnull(U_jeno,'')='' and isnull(U_Process,'')='Y'"
            objrsheader.DoQuery(strsql)
            If objrsheader.RecordCount = 0 Then Exit Sub

            For intheader As Integer = 0 To objrsheader.RecordCount - 1

                strsql = "Exec [Innova_HRMS_Posting_PayrollProcess] '" & objrsheader.Fields.Item("Docentry").Value.ToString & "'"
                objrs = objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objrs.DoQuery(strsql)
                If objrs.RecordCount > 0 Then
                    Try
                        If Not objcompany.InTransaction Then objcompany.StartTransaction()
                        Dim oPayrollJV As SAPbobsCOM.JournalVouchers
                        oPayrollJV = objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalVouchers)

                        oPayrollJV.JournalEntries.ReferenceDate = objrs.Fields.Item("Date").Value
                        oPayrollJV.JournalEntries.DueDate = objrs.Fields.Item("Date").Value
                        oPayrollJV.JournalEntries.TaxDate = objrs.Fields.Item("Date").Value
                        If objrs.Fields.Item("Transcode").Value.ToString <> "" Then oPayrollJV.JournalEntries.TransactionCode = objrs.Fields.Item("Transcode").Value.ToString

                        If objrs.Fields.Item("Memo").Value.ToString <> "" Then oPayrollJV.JournalEntries.Memo = objrs.Fields.Item("Memo").Value.ToString
                        If objrs.Fields.Item("Narration").Value.ToString <> "" Then oPayrollJV.JournalEntries.UserFields.Fields.Item("U_Narration").Value = objrs.Fields.Item("Narration").Value.ToString

                        If objrs.Fields.Item("Ref1").Value.ToString <> "" Then oPayrollJV.JournalEntries.Reference = objrs.Fields.Item("Ref1").Value.ToString
                        If objrs.Fields.Item("Ref2").Value.ToString <> "" Then oPayrollJV.JournalEntries.Reference2 = objrs.Fields.Item("Ref2").Value.ToString
                        If objrs.Fields.Item("Ref3").Value.ToString <> "" Then oPayrollJV.JournalEntries.Reference3 = objrs.Fields.Item("Ref3").Value.ToString

                        For i As Integer = 0 To objrs.RecordCount - 1
                            If objrs.Fields.Item("Type").Value.ToString.ToUpper = "A" Then

                                oPayrollJV.JournalEntries.Lines.AccountCode = objrs.Fields.Item("AccountCode").Value
                                If objrs.Fields.Item("DebitAmount").Value <> 0 Then oPayrollJV.JournalEntries.Lines.Debit = objrs.Fields.Item("DebitAmount").Value Else oPayrollJV.JournalEntries.Lines.Credit = objrs.Fields.Item("CreditAmount").Value

                                oPayrollJV.JournalEntries.Lines.Reference1 = objrs.Fields.Item("Lref1").Value
                                oPayrollJV.JournalEntries.Lines.Reference2 = objrs.Fields.Item("Lref2").Value
                                oPayrollJV.JournalEntries.Lines.AdditionalReference = objrs.Fields.Item("Lref3").Value

                                'If objrs.Fields.Item("Ocrcode1").Value.ToString <> "" Then oPayrollJV.JournalEntries.Lines.CostingCode = objrs.Fields.Item("Ocrcode1").Value
                                'If objrs.Fields.Item("Ocrcode2").Value.ToString <> "" Then oPayrollJV.JournalEntries.Lines.CostingCode2 = objrs.Fields.Item("Ocrcode2").Value
                                'If objrs.Fields.Item("Ocrcode3").Value.ToString <> "" Then oPayrollJV.JournalEntries.Lines.CostingCode3 = objrs.Fields.Item("Ocrcode3").Value
                                'If objrs.Fields.Item("Ocrcode4").Value.ToString <> "" Then oPayrollJV.JournalEntries.Lines.CostingCode4 = objrs.Fields.Item("Ocrcode4").Value
                                'If objrs.Fields.Item("Ocrcode5").Value.ToString <> "" Then oPayrollJV.JournalEntries.Lines.CostingCode5 = objrs.Fields.Item("Ocrcode5").Value
                                oPayrollJV.JournalEntries.Lines.Add()

                            End If
                            objrs.MoveNext()
                        Next

                        oPayrollJV.JournalEntries.Add()

                        objrs.MoveFirst()
                        oPayrollJV.JournalEntries.ReferenceDate = objrs.Fields.Item("Date").Value
                        oPayrollJV.JournalEntries.DueDate = objrs.Fields.Item("Date").Value
                        oPayrollJV.JournalEntries.TaxDate = objrs.Fields.Item("Date").Value
                        If objrs.Fields.Item("Transcode").Value.ToString <> "" Then oPayrollJV.JournalEntries.TransactionCode = objrs.Fields.Item("Transcode").Value.ToString

                        If objrs.Fields.Item("Memo").Value.ToString <> "" Then oPayrollJV.JournalEntries.Memo = objrs.Fields.Item("Memo").Value.ToString
                        If objrs.Fields.Item("Narration").Value.ToString <> "" Then oPayrollJV.JournalEntries.UserFields.Fields.Item("U_Narration").Value = objrs.Fields.Item("Narration").Value.ToString

                        If objrs.Fields.Item("Ref1").Value.ToString <> "" Then oPayrollJV.JournalEntries.Reference = objrs.Fields.Item("Ref1").Value.ToString
                        If objrs.Fields.Item("Ref2").Value.ToString <> "" Then oPayrollJV.JournalEntries.Reference2 = objrs.Fields.Item("Ref2").Value.ToString
                        If objrs.Fields.Item("Ref3").Value.ToString <> "" Then oPayrollJV.JournalEntries.Reference3 = objrs.Fields.Item("Ref3").Value.ToString

                        For i As Integer = 0 To objrs.RecordCount - 1
                            If objrs.Fields.Item("Type").Value.ToString.ToUpper = "D" Then

                                oPayrollJV.JournalEntries.Lines.AccountCode = objrs.Fields.Item("AccountCode").Value
                                If objrs.Fields.Item("DebitAmount").Value <> 0 Then oPayrollJV.JournalEntries.Lines.Debit = objrs.Fields.Item("DebitAmount").Value Else oPayrollJV.JournalEntries.Lines.Credit = objrs.Fields.Item("CreditAmount").Value

                                oPayrollJV.JournalEntries.Lines.Reference1 = objrs.Fields.Item("Lref1").Value
                                oPayrollJV.JournalEntries.Lines.Reference2 = objrs.Fields.Item("Lref2").Value
                                oPayrollJV.JournalEntries.Lines.AdditionalReference = objrs.Fields.Item("Lref3").Value

                                'If objrs.Fields.Item("Ocrcode1").Value.ToString <> "" Then oPayrollJV.JournalEntries.Lines.CostingCode = objrs.Fields.Item("Ocrcode1").Value
                                'If objrs.Fields.Item("Ocrcode2").Value.ToString <> "" Then oPayrollJV.JournalEntries.Lines.CostingCode2 = objrs.Fields.Item("Ocrcode2").Value
                                'If objrs.Fields.Item("Ocrcode3").Value.ToString <> "" Then oPayrollJV.JournalEntries.Lines.CostingCode3 = objrs.Fields.Item("Ocrcode3").Value
                                'If objrs.Fields.Item("Ocrcode4").Value.ToString <> "" Then oPayrollJV.JournalEntries.Lines.CostingCode4 = objrs.Fields.Item("Ocrcode4").Value
                                'If objrs.Fields.Item("Ocrcode5").Value.ToString <> "" Then oPayrollJV.JournalEntries.Lines.CostingCode5 = objrs.Fields.Item("Ocrcode5").Value
                                oPayrollJV.JournalEntries.Lines.Add()

                            End If
                            objrs.MoveNext()
                        Next

                        oPayrollJV.JournalEntries.Add()

                        lretcode = oPayrollJV.Add()
                        If lretcode <> 0 Then
                            If objcompany.InTransaction Then objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                            status_Update("OPRC", objrsheader.Fields.Item("Docentry").Value.ToString, 0, objcompany.GetLastErrorDescription, -1)
                        Else
                            objcompany.GetNewObjectCode(posted_entryno)
                            status_Update("OPRC", objrsheader.Fields.Item("Docentry").Value.ToString, 1, "Success", posted_entryno.ToString)
                            If Update_query("update [@SMPR_OPRC] set U_jeno='" & posted_entryno & "' where Docentry='" & objrsheader.Fields.Item("Docentry").Value.ToString & "'") Then
                                If objcompany.InTransaction Then objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                            Else
                                If objcompany.InTransaction Then objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                            End If
                        End If

                    Catch ex As Exception
                        If objcompany.InTransaction Then objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        status_Update("OPRC", objrsheader.Fields.Item("Docentry").Value.ToString, 0, ex.Message.ToString, -1)
                    End Try
                End If
                objrsheader.MoveNext()
            Next
        Catch ex As Exception
            status_Update("OPRC", objrs.Fields.Item("DocEntry").Value.ToString, 0, ex.Message.ToString, -1)
        End Try
    End Sub

    Private Sub Provision_posting_Table()
        Try
            strsql = "select DocEntry  from [@SMPR_OPRC] where U_ToDate=EOMONTH(getdate(),-1) and isnull(U_process,'')='Y'"
            'strsql = "Select 1 where ((select Max(U_providay) from [@SMPR_ACCT] where isnull(U_providay,0)<>0)<=Datepart(DD,getdate()))"
            objrs = objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objrs.DoQuery(strsql)
            If objrs.RecordCount = 0 Then Exit Sub
        Catch ex As Exception
            status_Update("PROV", "", 0, ex.Message.ToString, -1)
        End Try

        Try
            strsql = "select EOMONTH(getdate(),-1)  [Date] where (select count(1) from HRMS_PROVISION_details Where ProvisionDate =EOMONTH(getdate(),-1))=0"
            objrs = objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objrs.DoQuery(strsql)
            If objrs.RecordCount = 0 Then Exit Sub
        Catch ex As Exception
            status_Update("PROV", "", 0, ex.Message.ToString, -1)
        End Try

        Try

            strsql = " Declare @asondate as date"
            strsql += vbCrLf + " set @asondate=(select EOMONTH(getdate(),-1))"
            strsql += vbCrLf + " Exec [Innova_HRMS_Provision_Creation] @asondate"
            objrs = objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objrs.DoQuery(strsql)

        Catch ex As Exception
            status_Update("PROV", "", 0, ex.Message.ToString, -1)
        End Try

    End Sub

    Private Sub Provision_posting_JV()
        Try

            Dim objrsheader As SAPbobsCOM.Recordset
            objrsheader = objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strsql = "select Distinct Docentry from HRMS_PROVISION_DETAILS WHere Isnull(JENO,'')='' and isnull(finalize,'')='Y'"
            objrsheader.DoQuery(strsql)

            If objrsheader.RecordCount = 0 Then Exit Sub
            For intheader As Integer = 0 To objrsheader.RecordCount - 1

                strsql = "EXEC [Innova_HRMS_Provision_Posting] '" & objrsheader.Fields.Item("Docentry").Value.ToString & "'"
                objrs = objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objrs.DoQuery(strsql)
                If objrs.RecordCount > 0 Then
                    Try
                        If Not objcompany.InTransaction Then objcompany.StartTransaction()

                        Dim OprovisionJE As SAPbobsCOM.JournalVouchers
                        OprovisionJE = objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalVouchers)

                        objrs.MoveFirst()
                        OprovisionJE.JournalEntries.ReferenceDate = objrs.Fields.Item("ProvisionDate").Value
                        OprovisionJE.JournalEntries.DueDate = objrs.Fields.Item("ProvisionDate").Value
                        OprovisionJE.JournalEntries.TaxDate = objrs.Fields.Item("ProvisionDate").Value

                        If objrs.Fields.Item("Transcode").Value.ToString <> "" Then OprovisionJE.JournalEntries.TransactionCode = objrs.Fields.Item("Transcode").Value.ToString

                        OprovisionJE.JournalEntries.Memo = "Leave Salary Provision"
                        OprovisionJE.JournalEntries.UserFields.Fields.Item("U_Narration").Value = "Leave Salary Provision " + objrs.Fields.Item("Period").Value.ToString
                        OprovisionJE.JournalEntries.Reference = objrs.Fields.Item("Period").Value.ToString
                        OprovisionJE.JournalEntries.Reference2 = "Leave Salary Provision"
                        'If objrs.Fields.Item("Ref3").Value.ToString <> "" Then OprovisionJE.JournalEntries.Reference3 = objrs.Fields.Item("Ref3").Value.ToString

                        For i As Integer = 0 To objrs.RecordCount - 1
                            OprovisionJE.JournalEntries.Lines.AccountCode = objrs.Fields.Item("Leave_debitCode").Value
                            OprovisionJE.JournalEntries.Lines.Debit = objrs.Fields.Item("Leave_Amount").Value
                            'OprovisionJE.JournalEntries.Lines.Reference1 = objrs.Fields.Item("Lref1").Value 'OprovisionJE.JournalEntries.Lines.Reference2 = objrs.Fields.Item("Lref2").Value 'OprovisionJE.JournalEntries.Lines.AdditionalReference = objrs.Fields.Item("Lref3").Value

                            'If objrs.Fields.Item("Ocrcode1").Value.ToString <> "" Then OprovisionJE.JournalEntries.Lines.CostingCode = objrs.Fields.Item("Ocrcode1").Value
                            'If objrs.Fields.Item("Ocrcode2").Value.ToString <> "" Then OprovisionJE.JournalEntries.Lines.CostingCode2 = objrs.Fields.Item("Ocrcode2").Value
                            'If objrs.Fields.Item("Ocrcode3").Value.ToString <> "" Then OprovisionJE.JournalEntries.Lines.CostingCode3 = objrs.Fields.Item("Ocrcode3").Value
                            'If objrs.Fields.Item("Ocrcode4").Value.ToString <> "" Then OprovisionJE.JournalEntries.Lines.CostingCode4 = objrs.Fields.Item("Ocrcode4").Value
                            'If objrs.Fields.Item("Ocrcode5").Value.ToString <> "" Then OprovisionJE.JournalEntries.Lines.CostingCode5 = objrs.Fields.Item("Ocrcode5").Value
                            OprovisionJE.JournalEntries.Lines.Add()

                            OprovisionJE.JournalEntries.Lines.AccountCode = objrs.Fields.Item("Leave_CreditCode").Value
                            OprovisionJE.JournalEntries.Lines.Credit = objrs.Fields.Item("Leave_Amount").Value
                            ''OprovisionJE.JournalEntries.Lines.Reference1 = objrs.Fields.Item("Lref1").Value 'OprovisionJE.JournalEntries.Lines.Reference2 = objrs.Fields.Item("Lref2").Value 'OprovisionJE.JournalEntries.Lines.AdditionalReference = objrs.Fields.Item("Lref3").Value

                            'If objrs.Fields.Item("Ocrcode1").Value.ToString <> "" Then OprovisionJE.JournalEntries.Lines.CostingCode = objrs.Fields.Item("Ocrcode1").Value
                            'If objrs.Fields.Item("Ocrcode2").Value.ToString <> "" Then OprovisionJE.JournalEntries.Lines.CostingCode2 = objrs.Fields.Item("Ocrcode2").Value
                            'If objrs.Fields.Item("Ocrcode3").Value.ToString <> "" Then OprovisionJE.JournalEntries.Lines.CostingCode3 = objrs.Fields.Item("Ocrcode3").Value
                            'If objrs.Fields.Item("Ocrcode4").Value.ToString <> "" Then OprovisionJE.JournalEntries.Lines.CostingCode4 = objrs.Fields.Item("Ocrcode4").Value
                            'If objrs.Fields.Item("Ocrcode5").Value.ToString <> "" Then OprovisionJE.JournalEntries.Lines.CostingCode5 = objrs.Fields.Item("Ocrcode5").Value
                            OprovisionJE.JournalEntries.Lines.Add()

                            objrs.MoveNext()
                        Next

                        OprovisionJE.JournalEntries.Add()


                        objrs.MoveFirst()
                        OprovisionJE.JournalEntries.ReferenceDate = objrs.Fields.Item("ProvisionDate").Value
                        OprovisionJE.JournalEntries.DueDate = objrs.Fields.Item("ProvisionDate").Value
                        OprovisionJE.JournalEntries.TaxDate = objrs.Fields.Item("ProvisionDate").Value

                        If objrs.Fields.Item("Transcode").Value.ToString <> "" Then OprovisionJE.JournalEntries.TransactionCode = objrs.Fields.Item("Transcode").Value.ToString

                        OprovisionJE.JournalEntries.Memo = "Air Ticket Provision"
                        OprovisionJE.JournalEntries.UserFields.Fields.Item("U_Narration").Value = "Air Ticket Provision " + objrs.Fields.Item("Period").Value.ToString
                        OprovisionJE.JournalEntries.Reference = objrs.Fields.Item("Period").Value.ToString
                        OprovisionJE.JournalEntries.Reference2 = "Air Ticket Provision"
                        'If objrs.Fields.Item("Ref3").Value.ToString <> "" Then OprovisionJE.JournalEntries.Reference3 = objrs.Fields.Item("Ref3").Value.ToString

                        For i As Integer = 0 To objrs.RecordCount - 1
                            OprovisionJE.JournalEntries.Lines.AccountCode = objrs.Fields.Item("Air_debitCode").Value
                            OprovisionJE.JournalEntries.Lines.Debit = objrs.Fields.Item("AirTicket_Amount").Value
                            'OprovisionJE.JournalEntries.Lines.Reference1 = objrs.Fields.Item("Lref1").Value 'OprovisionJE.JournalEntries.Lines.Reference2 = objrs.Fields.Item("Lref2").Value 'OprovisionJE.JournalEntries.Lines.AdditionalReference = objrs.Fields.Item("Lref3").Value

                            'If objrs.Fields.Item("Ocrcode1").Value.ToString <> "" Then OprovisionJE.JournalEntries.Lines.CostingCode = objrs.Fields.Item("Ocrcode1").Value
                            'If objrs.Fields.Item("Ocrcode2").Value.ToString <> "" Then OprovisionJE.JournalEntries.Lines.CostingCode2 = objrs.Fields.Item("Ocrcode2").Value
                            'If objrs.Fields.Item("Ocrcode3").Value.ToString <> "" Then OprovisionJE.JournalEntries.Lines.CostingCode3 = objrs.Fields.Item("Ocrcode3").Value
                            'If objrs.Fields.Item("Ocrcode4").Value.ToString <> "" Then OprovisionJE.JournalEntries.Lines.CostingCode4 = objrs.Fields.Item("Ocrcode4").Value
                            'If objrs.Fields.Item("Ocrcode5").Value.ToString <> "" Then OprovisionJE.JournalEntries.Lines.CostingCode5 = objrs.Fields.Item("Ocrcode5").Value
                            OprovisionJE.JournalEntries.Lines.Add()

                            OprovisionJE.JournalEntries.Lines.AccountCode = objrs.Fields.Item("Air_CreditCode").Value
                            OprovisionJE.JournalEntries.Lines.Credit = objrs.Fields.Item("AirTicket_Amount").Value
                            ''OprovisionJE.JournalEntries.Lines.Reference1 = objrs.Fields.Item("Lref1").Value 'OprovisionJE.JournalEntries.Lines.Reference2 = objrs.Fields.Item("Lref2").Value 'OprovisionJE.JournalEntries.Lines.AdditionalReference = objrs.Fields.Item("Lref3").Value

                            'If objrs.Fields.Item("Ocrcode1").Value.ToString <> "" Then OprovisionJE.JournalEntries.Lines.CostingCode = objrs.Fields.Item("Ocrcode1").Value
                            'If objrs.Fields.Item("Ocrcode2").Value.ToString <> "" Then OprovisionJE.JournalEntries.Lines.CostingCode2 = objrs.Fields.Item("Ocrcode2").Value
                            'If objrs.Fields.Item("Ocrcode3").Value.ToString <> "" Then OprovisionJE.JournalEntries.Lines.CostingCode3 = objrs.Fields.Item("Ocrcode3").Value
                            'If objrs.Fields.Item("Ocrcode4").Value.ToString <> "" Then OprovisionJE.JournalEntries.Lines.CostingCode4 = objrs.Fields.Item("Ocrcode4").Value
                            'If objrs.Fields.Item("Ocrcode5").Value.ToString <> "" Then OprovisionJE.JournalEntries.Lines.CostingCode5 = objrs.Fields.Item("Ocrcode5").Value
                            OprovisionJE.JournalEntries.Lines.Add()

                            objrs.MoveNext()
                        Next

                        OprovisionJE.JournalEntries.Add()


                        objrs.MoveFirst()
                        OprovisionJE.JournalEntries.ReferenceDate = objrs.Fields.Item("ProvisionDate").Value
                        OprovisionJE.JournalEntries.DueDate = objrs.Fields.Item("ProvisionDate").Value
                        OprovisionJE.JournalEntries.TaxDate = objrs.Fields.Item("ProvisionDate").Value

                        If objrs.Fields.Item("Transcode").Value.ToString <> "" Then OprovisionJE.JournalEntries.TransactionCode = objrs.Fields.Item("Transcode").Value.ToString

                        OprovisionJE.JournalEntries.Memo = "Gratuity Provision"
                        OprovisionJE.JournalEntries.UserFields.Fields.Item("U_Narration").Value = "Gratuity Provision " + objrs.Fields.Item("Period").Value.ToString
                        OprovisionJE.JournalEntries.Reference = objrs.Fields.Item("Period").Value.ToString
                        OprovisionJE.JournalEntries.Reference2 = "Gratuity Provision"
                        'If objrs.Fields.Item("Ref3").Value.ToString <> "" Then OprovisionJE.JournalEntries.Reference3 = objrs.Fields.Item("Ref3").Value.ToString

                        For i As Integer = 0 To objrs.RecordCount - 1
                            OprovisionJE.JournalEntries.Lines.AccountCode = objrs.Fields.Item("Gratuity_debitCode").Value
                            OprovisionJE.JournalEntries.Lines.Debit = objrs.Fields.Item("Gratuity_Amount").Value
                            'OprovisionJE.JournalEntries.Lines.Reference1 = objrs.Fields.Item("Lref1").Value 'OprovisionJE.JournalEntries.Lines.Reference2 = objrs.Fields.Item("Lref2").Value 'OprovisionJE.JournalEntries.Lines.AdditionalReference = objrs.Fields.Item("Lref3").Value

                            'If objrs.Fields.Item("Ocrcode1").Value.ToString <> "" Then OprovisionJE.JournalEntries.Lines.CostingCode = objrs.Fields.Item("Ocrcode1").Value
                            'If objrs.Fields.Item("Ocrcode2").Value.ToString <> "" Then OprovisionJE.JournalEntries.Lines.CostingCode2 = objrs.Fields.Item("Ocrcode2").Value
                            'If objrs.Fields.Item("Ocrcode3").Value.ToString <> "" Then OprovisionJE.JournalEntries.Lines.CostingCode3 = objrs.Fields.Item("Ocrcode3").Value
                            'If objrs.Fields.Item("Ocrcode4").Value.ToString <> "" Then OprovisionJE.JournalEntries.Lines.CostingCode4 = objrs.Fields.Item("Ocrcode4").Value
                            'If objrs.Fields.Item("Ocrcode5").Value.ToString <> "" Then OprovisionJE.JournalEntries.Lines.CostingCode5 = objrs.Fields.Item("Ocrcode5").Value
                            OprovisionJE.JournalEntries.Lines.Add()

                            OprovisionJE.JournalEntries.Lines.AccountCode = objrs.Fields.Item("Gratuity_CreditCode").Value
                            OprovisionJE.JournalEntries.Lines.Credit = objrs.Fields.Item("Gratuity_Amount").Value
                            ''OprovisionJE.JournalEntries.Lines.Reference1 = objrs.Fields.Item("Lref1").Value 'OprovisionJE.JournalEntries.Lines.Reference2 = objrs.Fields.Item("Lref2").Value 'OprovisionJE.JournalEntries.Lines.AdditionalReference = objrs.Fields.Item("Lref3").Value

                            'If objrs.Fields.Item("Ocrcode1").Value.ToString <> "" Then OprovisionJE.JournalEntries.Lines.CostingCode = objrs.Fields.Item("Ocrcode1").Value
                            'If objrs.Fields.Item("Ocrcode2").Value.ToString <> "" Then OprovisionJE.JournalEntries.Lines.CostingCode2 = objrs.Fields.Item("Ocrcode2").Value
                            'If objrs.Fields.Item("Ocrcode3").Value.ToString <> "" Then OprovisionJE.JournalEntries.Lines.CostingCode3 = objrs.Fields.Item("Ocrcode3").Value
                            'If objrs.Fields.Item("Ocrcode4").Value.ToString <> "" Then OprovisionJE.JournalEntries.Lines.CostingCode4 = objrs.Fields.Item("Ocrcode4").Value
                            'If objrs.Fields.Item("Ocrcode5").Value.ToString <> "" Then OprovisionJE.JournalEntries.Lines.CostingCode5 = objrs.Fields.Item("Ocrcode5").Value
                            OprovisionJE.JournalEntries.Lines.Add()

                            objrs.MoveNext()
                        Next

                        OprovisionJE.JournalEntries.Add()

                        lretcode = OprovisionJE.Add()
                        If lretcode <> 0 Then
                            If objcompany.InTransaction Then objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                            status_Update("PROV", objrsheader.Fields.Item("Docentry").Value.ToString, 0, objcompany.GetLastErrorDescription, -1)
                        Else
                            objcompany.GetNewObjectCode(posted_entryno)
                            status_Update("PROV", objrsheader.Fields.Item("Docentry").Value.ToString, 1, "Success", posted_entryno.ToString)
                            If Update_query("update HRMS_PROVISION_DETAILS set jeno='" & posted_entryno & "' where Docentry='" & objrsheader.Fields.Item("Docentry").Value.ToString & "'") Then
                                If objcompany.InTransaction Then objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                            Else
                                If objcompany.InTransaction Then objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                            End If
                        End If

                    Catch ex As Exception
                        If objcompany.InTransaction Then objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        status_Update("PROV", objrsheader.Fields.Item("Docentry").Value.ToString, 0, ex.Message.ToString, -1)
                    End Try
                End If

                objrsheader.MoveNext()
            Next

        Catch ex As Exception

        End Try
    End Sub

    Private Sub status_Update(ByVal objtype As String, ByVal Docentry As String, ByVal status As String, ByVal remarks As String, ByVal JENO As String)
        Try
            Dim objstatus As SAPbobsCOM.Recordset
            strsql = "insert into HRMS_POSTINGLOG (OBJTYPE,DOCENTRY,JENO,STATUS,Remarks) Values("
            strsql += vbCrLf + "'" & objtype.ToString & "','" & Docentry.ToString & "','" & JENO.ToString & "','" & status.ToString & "','" & remarks.ToString.Replace("'", "") & "') "
            objstatus = objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objstatus.DoQuery(strsql)
        Catch ex As Exception
            'MsgBox("Error in Update")
        End Try
    End Sub

    Private Function Update_query(ByVal strsql As String)
        Try
            Dim objstatus As SAPbobsCOM.Recordset
            objstatus = objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objstatus.DoQuery(strsql)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
End Class
