<%@ Page Language="VB" Theme="AIMSDefault" ValidateRequest="false" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<script runat="server">

    ' STILL TO DO

    ' ConsignmentTrackingStage
    ' [ConsignmentKey] [int]
    ' [TrackingID] [nvarchar](20)
    ' [Location] [nvarchar](50)
    ' [Description] [nvarchar](255)
    ' [TrackedOn] [datetime]
    
    Const CUSTOMER_JUPITER As Int32 = 784
    
    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString

    Protected Sub Page_Load(ByVal Source As Object, ByVal E As EventArgs)
        If Not IsNumeric(Session("CustomerKey")) Then
            Server.Transfer("session_expired.aspx")
        End If
        If Not IsPostBack Then
            Call HideAllPanelsAndRows()
            Call BindPDFUploadView()
            Call BindPrintOrdersView()
            If Session("UserType") <> "On Demand Supplier" Then
                gvPDFStatus.Columns(0).Visible = False
                gvPrintOrders.Columns(0).Visible = False
            End If
        End If
    End Sub

    Protected Sub SetTitle()
        Dim sTitle As String = Session("SiteTitle")
        If sTitle <> String.Empty Then
            sTitle += " - "
        End If
        Page.Header.Title = sTitle & "Print Status"
    End Sub
   
    Protected Sub HideAllPanelsAndRows()
        trJobDetails01.Visible = False
        trJobDetails02.Visible = False
        trJobDetails03.Visible = False
        trJobDetails04.Visible = False
        trJobDetails05.Visible = False
        trJobDetails06.Visible = False
        tbCourierJobNumber.Text = String.Empty
        pnlCourierJobNo.Visible = False
    End Sub
    
    Protected Sub BindPDFUploadView()
        Dim sSQL As String
        If cbIncludeHiddenItems.Checked Then
            sSQL = "SELECT jpc.PrintType, jpu.ReadyToPrint, ISNULL(jpu.Hidden, 0) 'Hidden', lp.LogisticProductKey, lp.ProductCode, lp.ProductDate, lp.ProductDescription FROM ClientData_Jupiter_PDFUploads jpu INNER JOIN LogisticProduct lp ON jpu.LogisticProductKey = lp.LogisticProductKey INNER JOIN ClientData_Jupiter_PrintCost jpc ON lp.Misc2 = jpc.[id] ORDER BY ReadyToPrint, ProductCode"
        Else
            sSQL = "SELECT jpc.PrintType, jpu.ReadyToPrint, ISNULL(jpu.Hidden, 0) 'Hidden', lp.LogisticProductKey, lp.ProductCode, lp.ProductDate, lp.ProductDescription FROM ClientData_Jupiter_PDFUploads jpu INNER JOIN LogisticProduct lp ON jpu.LogisticProductKey = lp.LogisticProductKey INNER JOIN ClientData_Jupiter_PrintCost jpc ON lp.Misc2 = jpc.[id] WHERE ISNULL(Hidden, 0) = 0 ORDER BY ReadyToPrint, ProductCode"
        End If
        'sSQL = "SELECT TOP " & ddlPDFUploads.SelectedValue & " jpc.PrintType, jpu.ReadyToPrint, lp.LogisticProductKey, lp.ProductCode, lp.ProductDate, lp.ProductDescription FROM ClientData_Jupiter_PDFUploads jpu INNER JOIN LogisticProduct lp ON jpu.LogisticProductKey = lp.LogisticProductKey INNER JOIN ClientData_Jupiter_PrintCost jpc ON lp.Misc2 = jpc.[id] ORDER BY ReadyToPrint, ProductCode"
        'sSQL = "SELECT TOP " & ddlPDFUploads.SelectedValue & " jpc.PrintType, jpu.ReadyToPrint, lp.LogisticProductKey, lp.ProductCode, lp.ProductDate, lp.ProductDescription FROM ClientData_Jupiter_PDFUploads jpu INNER JOIN LogisticProduct lp ON jpu.LogisticProductKey = lp.LogisticProductKey INNER JOIN ClientData_Jupiter_PrintCost jpc ON lp.Misc2 = jpc.[id] WHERE lp.ArchiveFlag = 'N'"
        Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
        gvPDFStatus.DataSource = dt
        gvPDFStatus.DataBind()
    End Sub
    
    Protected Sub BindPrintOrdersView()
        Dim sSQL As String
        If cbIncludeDespatchedJobs.Checked Then
            'sSQL = "SELECT CAST(REPLACE(CONVERT(VARCHAR(11),  CreatedOn, 106), ' ', '-') AS varchar(20)) + ' ' + SUBSTRING((CONVERT(VARCHAR(8), CreatedOn, 108)),1,5) 'CreatedOn', c.[key] 'ConsignmentKey', AWB, FirstName + ' ' + LastName 'OrderedBy', AgentAWB 'CourierAWB', AgentRef 'PrintStatus', ISNULL(SpecialInstructions, '') 'SpecialInstructions', ISNULL(Misc1, '') 'PrintTurnround' FROM Consignment c INNER JOIN UserProfile up ON c.UserKey = up.[key] WHERE CreatedOn > GETDATE() - " & ddlPrintOrdersLastNDays.SelectedValue & " AND c.CustomerKey = " & CUSTOMER_JUPITER & " AND (AgentRef LIKE '%PRINT%' OR AgentRef LIKE '%JOB%') AND NOT StateId IN ('CANCELLED', 'WITH OPERATIONS') ORDER BY c.[key] DESC"
            sSQL = "SELECT CAST(REPLACE(CONVERT(VARCHAR(11),  CreatedOn, 106), ' ', '-') AS varchar(20)) + ' ' + SUBSTRING((CONVERT(VARCHAR(8), CreatedOn, 108)),1,5) 'CreatedOn', c.[key] 'ConsignmentKey', AWB, FirstName + ' ' + LastName 'OrderedBy', AgentAWB 'CourierAWB', AgentRef 'PrintStatus', ISNULL(SpecialInstructions, '') 'SpecialInstructions', ISNULL(Misc1, '') 'PrintTurnround' FROM Consignment c INNER JOIN UserProfile up ON c.UserKey = up.[key] WHERE CreatedOn > GETDATE() - " & ddlPrintOrdersLastNDays.SelectedValue & " AND c.CustomerKey = " & CUSTOMER_JUPITER & " AND (AgentRef LIKE '%PRINT%' OR AgentRef LIKE '%JOB%') AND NOT StateId IN ('CANCELLED') ORDER BY c.[key] DESC"
        Else
            'sSQL = "SELECT CAST(REPLACE(CONVERT(VARCHAR(11),  CreatedOn, 106), ' ', '-') AS varchar(20)) + ' ' + SUBSTRING((CONVERT(VARCHAR(8), CreatedOn, 108)),1,5) 'CreatedOn', c.[key] 'ConsignmentKey', AWB, FirstName + ' ' + LastName 'OrderedBy', AgentAWB 'CourierAWB', AgentRef 'PrintStatus', ISNULL(SpecialInstructions, '') 'SpecialInstructions', ISNULL(Misc1, '') 'PrintTurnround' FROM Consignment c INNER JOIN UserProfile up ON c.UserKey = up.[key] WHERE CreatedOn > GETDATE() - " & ddlPrintOrdersLastNDays.SelectedValue & " AND c.CustomerKey = " & CUSTOMER_JUPITER & " AND (AgentRef LIKE '%PRINT%' OR AgentRef LIKE '%JOB%') AND NOT AgentRef LIKE '%DESPATCHED%' AND NOT StateId IN ('CANCELLED', 'WITH OPERATIONS') ORDER BY c.[key] DESC"
            sSQL = "SELECT CAST(REPLACE(CONVERT(VARCHAR(11),  CreatedOn, 106), ' ', '-') AS varchar(20)) + ' ' + SUBSTRING((CONVERT(VARCHAR(8), CreatedOn, 108)),1,5) 'CreatedOn', c.[key] 'ConsignmentKey', AWB, FirstName + ' ' + LastName 'OrderedBy', AgentAWB 'CourierAWB', AgentRef 'PrintStatus', ISNULL(SpecialInstructions, '') 'SpecialInstructions', ISNULL(Misc1, '') 'PrintTurnround' FROM Consignment c INNER JOIN UserProfile up ON c.UserKey = up.[key] WHERE CreatedOn > GETDATE() - " & ddlPrintOrdersLastNDays.SelectedValue & " AND c.CustomerKey = " & CUSTOMER_JUPITER & " AND (AgentRef LIKE '%PRINT%' OR AgentRef LIKE '%JOB%') AND NOT (AgentRef LIKE '%DESPATCHED%' OR StateId IN ('CANCELLED')) ORDER BY c.[key] DESC"
        End If
        Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
        gvPrintOrders.DataSource = dt
        gvPrintOrders.DataBind()
    End Sub
    
    Protected Function ExecuteQueryToDataTable(ByVal sQuery As String) As DataTable
        Dim oDataTable As New DataTable
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter(sQuery, oConn)
        Dim oCmd As SqlCommand = New SqlCommand(sQuery, oConn)
        Try
            oAdapter.Fill(oDataTable)
            oConn.Open()
        Catch ex As Exception
            WebMsgBox.Show("Error in ExecuteQueryToDataTable executing: " & sQuery & " : " & ex.Message)
        Finally
            oConn.Close()
        End Try
        ExecuteQueryToDataTable = oDataTable
    End Function

    Protected Sub SetJupiterPODProductArchiveFlag(ByVal nLogisticProductKey As Int32, ByVal sArchiveFlag As String)
        Dim sSQL As String = "UPDATE LogisticProduct SET ArchiveFlag = '" & sArchiveFlag & "' WHERE LogisticProductKey = " & nLogisticProductKey
        Call ExecuteQueryToDataTable(sSQL)
    End Sub
    
    Protected Sub btnReadyToPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim b As Button = sender
        Dim nLogisticProductKey As Int32 = b.CommandArgument
        Dim sSQL As String = "UPDATE ClientData_Jupiter_PDFUploads SET ReadyToPrint = 1 WHERE LogisticProductKey = " & nLogisticProductKey
        Call ExecuteQueryToDataTable(sSQL)
        Call LogJupiterAuditEvent("READY_TO_PRINT", "Printer marked product as ready to print", nLogisticProductKey)
        Call BindPDFUploadView()
        Call SetJupiterPODProductArchiveFlag(nLogisticProductKey, "N")
    End Sub

    Protected Sub btnCourierJobNo_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim b As Button = sender
        lblConsignmentNumberForCourierJobNo.Text = b.CommandArgument
        btnSaveCourierJobNumber.CommandArgument = b.CommandArgument
        pnlCourierJobNo.Visible = True
        tbCourierJobNumber.Text = String.Empty
        tbCourierJobNumber.Focus()
        cbOverrideValidation.Checked = False
    End Sub
    
    Protected Sub btnJobReceived_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim b As Button = sender
        Dim sConsignmentKey = b.CommandArgument
        Dim sSQL As String = "UPDATE Consignment SET AgentRef = '2. JOB RECEIVED BY PRINTER' + ' (' + CAST(REPLACE(CONVERT(VARCHAR(11),  GETDATE(), 106), ' ', '-') AS varchar(20)) + ' ' + SUBSTRING((CONVERT(VARCHAR(8), GETDATE(), 108)),1,5) + ')' WHERE [key] = " & sConsignmentKey
        Call ExecuteQueryToDataTable(sSQL)
        Call LogJupiterAuditEvent("JOB_RECEIVED", "Printer acknowledged receipt of job " & sConsignmentKey, , sConsignmentKey)
        Call BindPrintOrdersView()
    End Sub

    Protected Sub btnJobPrinted_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim b As Button = sender
        Dim sConsignmentKey = b.CommandArgument
        Dim sSQL As String = "UPDATE Consignment SET AgentRef = '3. PRINT COMPLETED' + ' (' + CAST(REPLACE(CONVERT(VARCHAR(11),  GETDATE(), 106), ' ', '-') AS varchar(20)) + ' ' + SUBSTRING((CONVERT(VARCHAR(8), GETDATE(), 108)),1,5) + ')' WHERE [key] = " & sConsignmentKey
        Call ExecuteQueryToDataTable(sSQL)
        Call LogJupiterAuditEvent("JOB_PRINTED", "Printer marked job " & sConsignmentKey & " as print complete", , sConsignmentKey)
        Call BindPrintOrdersView()
    End Sub
    
    Protected Sub btnJobDespatched_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim b As Button = sender
        Dim sConsignmentKey = b.CommandArgument
        Dim sSQL As String = "UPDATE Consignment SET AgentRef = '4. JOB DESPATCHED' + ' (' + CAST(REPLACE(CONVERT(VARCHAR(11),  GETDATE(), 106), ' ', '-') AS varchar(20)) + ' ' + SUBSTRING((CONVERT(VARCHAR(8), GETDATE(), 108)),1,5) + ')' WHERE [key] = " & sConsignmentKey
        Call ExecuteQueryToDataTable(sSQL)
        Call LogJupiterAuditEvent("JOB_DESPATCHED", "Printer marked job " & sConsignmentKey & " as despatched", , sConsignmentKey)
        Call BindPrintOrdersView()
    End Sub
    
    Protected Sub LogJupiterAuditEvent(ByVal sEventCode As String, ByVal sEventDescription As String, Optional ByVal nProductKey As Int32 = 0, Optional ByVal nConsignmentKey As Int32 = 0)
        Dim sSQL As String = "INSERT INTO ClientData_Jupiter_AuditTrail (EventCode, EventDescription, ProductKey, ConsignmentKey, EventDateTime, EventAuthor) VALUES ('" & sEventCode & "', '" & sEventDescription & "', " & nProductKey & ", " & nConsignmentKey & ", GETDATE(), " & Session("UserKey") & ")"
        Call JupiterNotification(sEventCode, sEventDescription, nConsignmentKey)
        Call ExecuteQueryToDataTable(sSQL)
        If nConsignmentKey > 0 Then
            sSQL = "INSERT INTO ConsignmentTrackingStage (ConsignmentKey, TrackingID, Location, Description, TrackedOn) VALUES (" & nConsignmentKey & ", '" & sEventCode & "', '', '" & sEventDescription.Replace("'", "''") & "', GETDATE())"
            Call ExecuteQueryToDataTable(sSQL)
        End If
    End Sub
    
    Protected Sub JupiterNotification(ByVal sEventCode As String, ByVal sEventDescription As String, nConsignmentKey As Int32)
        If sEventCode = "EVENT_NOTIFICATION" Then
            Exit Sub
        End If
        Dim sSQL As String = "SELECT EmailAddr FROM ClientData_Jupiter_EventNotification WHERE EventCode = '" & sEventCode & "'"
        Dim sbMessage As New StringBuilder
        Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
        sbMessage.Append("Jupiter Asset Management Event Notification")
        sbMessage.Append(Environment.NewLine)
        sbMessage.Append(Environment.NewLine)
        sbMessage.Append("Event Code: ")
        sbMessage.Append(sEventCode)
        sbMessage.Append(Environment.NewLine)
        sbMessage.Append("Event Description: ")
        sbMessage.Append(sEventDescription)
        sbMessage.Append(Environment.NewLine)
        sbMessage.Append(Environment.NewLine)
        sbMessage.Append("Visit http://my.transworld.eu.com/jupiter to view further information.")
        sbMessage.Append(Environment.NewLine)
        sbMessage.Append(Environment.NewLine)
        sbMessage.Append("Please do not reply to this email as replies are not monitored.  Thank you.")
        sbMessage.Append(Environment.NewLine)
        sbMessage.Append("Transworld")
        sbMessage.Append(Environment.NewLine)
        Dim sPlainTextBody As String = sbMessage.ToString
        Dim sHTMLBody As String = sbMessage.ToString.Replace(Environment.NewLine, "<br />" & Environment.NewLine)
        For Each dr As DataRow In dt.Rows
            Dim sRecipient As String = dr(0)
            If sRecipient = "#user#" AndAlso nConsignmentKey > 0 Then
                sSQL = "SELECT EmailAddr FROM UserProfile up INNER JOIN Consignment c ON up.[key] = c.UserKey WHERE c.[key] = " & nConsignmentKey
                Dim dtEmailAddr As DataTable = ExecuteQueryToDataTable(sSQL)
                If dtEmailAddr.Rows.Count = 1 Then
                    sRecipient = dtEmailAddr.Rows(0).Item(0)
                End If
            End If
            Call SendMail("JUPITER_EVENT", sRecipient, "Jupiter Event Notification - " & sEventCode, sPlainTextBody, sHTMLBody)
            Call LogJupiterAuditEvent("EVENT_NOTIFICATION", sEventCode & " to: " & dr(0))
        Next
    End Sub

    Protected Sub SendMail(ByVal sType As String, ByVal sRecipient As String, ByVal sSubject As String, ByVal sBodyText As String, ByVal sBodyHTML As String)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Email_AddToQueue", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Try
            oCmd.Parameters.Add(New SqlParameter("@MessageId", SqlDbType.NVarChar, 20))
            oCmd.Parameters("@MessageId").Value = sType
    
            oCmd.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
            oCmd.Parameters("@CustomerKey").Value = Session("CustomerKey")
    
            oCmd.Parameters.Add(New SqlParameter("@StockBookingKey", SqlDbType.Int))
            oCmd.Parameters("@StockBookingKey").Value = 0
    
            oCmd.Parameters.Add(New SqlParameter("@ConsignmentKey", SqlDbType.Int))
            oCmd.Parameters("@ConsignmentKey").Value = 0
    
            oCmd.Parameters.Add(New SqlParameter("@ProductKey", SqlDbType.Int))
            oCmd.Parameters("@ProductKey").Value = 0
    
            oCmd.Parameters.Add(New SqlParameter("@To", SqlDbType.NVarChar, 100))
            oCmd.Parameters("@To").Value = sRecipient
    
            oCmd.Parameters.Add(New SqlParameter("@Subject", SqlDbType.NVarChar, 60))
            oCmd.Parameters("@Subject").Value = sSubject
    
            oCmd.Parameters.Add(New SqlParameter("@BodyText", SqlDbType.NText))
            oCmd.Parameters("@BodyText").Value = sBodyText
    
            oCmd.Parameters.Add(New SqlParameter("@BodyHTML", SqlDbType.NText))
            oCmd.Parameters("@BodyHTML").Value = sBodyHTML
    
            oCmd.Parameters.Add(New SqlParameter("@QueuedBy", SqlDbType.Int))
            oCmd.Parameters("@QueuedBy").Value = Session("UserKey")
    
            oConn.Open()
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            WebMsgBox.Show("Error in SendMail: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub

    Protected Sub btnSaveCourierJobNumber_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        'Dim bValidCourierConsignmentNumber As Boolean
        Dim b As Button = sender
        Dim sConsignmentKey = b.CommandArgument
        tbCourierJobNumber.Text = tbCourierJobNumber.Text.Trim
        If tbCourierJobNumber.Text = String.Empty Then
            WebMsgBox.Show("Please enter a valid courier job number.")
        Else
            'bValidCourierConsignmentNumber = tbCourierJobNumber.Text.StartsWith("2") AndAlso tbCourierJobNumber.Text.Length = 8
            
            If (tbCourierJobNumber.Text.StartsWith("2") AndAlso tbCourierJobNumber.Text.Length = 8) Or cbOverrideValidation.Checked Then
                Dim sSQL As String = "UPDATE Consignment SET AgentAWB = '" & tbCourierJobNumber.Text.Replace("'", "''") & "' WHERE [key] = " & sConsignmentKey
                Call ExecuteQueryToDataTable(sSQL)
                Call LogJupiterAuditEvent("COURIER_JOBNO_LOGGED", "Courier job number entered: " & tbCourierJobNumber.Text.Replace("'", "''") & " for job " & sConsignmentKey, , sConsignmentKey)
                Call BindPrintOrdersView()
                sSQL = "DELETE FROM ClientData_Jupiter_TrackingHook WHERE AWB_POD = '" & sConsignmentKey & "' INSERT INTO ClientData_Jupiter_TrackingHook (AWB_POD, AWB_Courier, CreatedBy, CreatedOn) VALUES ('" & sConsignmentKey & "', '" & tbCourierJobNumber.Text.Replace("'", "''") & "', " & Session("UserKey") & ", GETDATE())"
                Call ExecuteQueryToDataTable(sSQL)
                Call HideAllPanelsAndRows()
            Else
                WebMsgBox.Show("The courier consignment value you entered (" & tbCourierJobNumber.Text & ") is not in the expected format.\n\nEither correct the value, or to proceed with the entered value first click the 'override validation' check box, then click the save button.")
            End If
        End If
    End Sub

    Protected Sub gvPDFStatus_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs)
        Const COL_READY_TO_PRINT As Int32 = 5
        Dim gvr As GridViewRow = e.Row
        Dim btnReadyToPrint As Button
        Dim hidReadyToPrint As HiddenField
        If gvr.RowType = DataControlRowType.DataRow Then
            btnReadyToPrint = gvr.Cells(0).FindControl("btnReadyToPrint")
            hidReadyToPrint = gvr.Cells(0).FindControl("hidReadyToPrint")
            If hidReadyToPrint.Value = True Then
                btnReadyToPrint.Visible = False
            End If
            Dim sReadyToPrint As String = gvr.Cells(COL_READY_TO_PRINT).Text
            If sReadyToPrint.ToLower = "true" Then
                gvr.Cells(COL_READY_TO_PRINT).Text = "YES"
                gvr.Cells(COL_READY_TO_PRINT).BackColor = Drawing.Color.LightGreen
            Else
                gvr.Cells(COL_READY_TO_PRINT).Text = "NO"
                gvr.Cells(COL_READY_TO_PRINT).BackColor = Drawing.Color.Silver
            End If
        End If
    End Sub
    
    Protected Sub ddlPDFUploads_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        gvPDFStatus.PageIndex = 0
        gvPDFStatus.PageSize = ddlPDFUploads.SelectedValue
        Call BindPDFUploadView()
    End Sub

    Protected Sub ddlPrintOrdersLastNDays_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call BindPrintOrdersView()
    End Sub
    
    Protected Sub lnkbtnRefresh_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        gvPDFStatus.PageIndex = 0
        Call HideAllPanelsAndRows()
        Call BindPDFUploadView()
        Call BindPrintOrdersView()
    End Sub
    
    Protected Function GetHyperLinkNavigateURL(ByVal DataItem As Object) As String
        GetHyperLinkNavigateURL = ConfigLib.GetConfigItem_Virtual_PDF_URL & DataBinder.Eval(DataItem, "LogisticProductKey") & ".pdf"
    End Function
    
    Protected Sub lnkbtnAWB_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim lnkbtn As LinkButton = sender
        Dim sConsignmentKey As String = lnkbtn.CommandArgument
        Dim sSQL As String = "SELECT ProductCode 'Product Code', ProductDate 'Product Date', ProductDescription 'Description', jpc.PrintType 'Print Type', ItemsOut 'Qty' FROM LogisticMovement lm INNER JOIN LogisticProduct lp ON lm.LogisticProductKey = lp.LogisticProductKey INNER JOIN ClientData_Jupiter_PrintCost jpc ON lp.Misc2 = jpc.[id] WHERE ConsignmentKey = " & sConsignmentKey
        Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
        gvJobDetails.DataSource = dt
        gvJobDetails.DataBind()
        
        sSQL = "SELECT CneeName, ISNULL(CneeCtcName,'') 'CneeCtcName', CneeAddr1, ISNULL(CneeAddr2,'') 'CneeAddr2', ISNULL(CneeAddr3,'') 'CneeAddr3', CneeTown, ISNULL(CneePostCode,'') 'CneePostCode', CountryName FROM Consignment c INNER JOIN Country ctry ON c.CneeCountryKey = ctry.CountryKey WHERE [key] = " & sConsignmentKey
        Dim dtDeliverTo As DataRow = ExecuteQueryToDataTable(sSQL).Rows(0)
        Dim sbDeliverTo As New StringBuilder
        If dtDeliverTo("CneeCtcName").ToString.Trim <> String.Empty Then
            sbDeliverTo.Append(dtDeliverTo("CneeCtcName"))
            sbDeliverTo.Append(", ")
        End If
        sbDeliverTo.Append(dtDeliverTo("CneeName"))
        sbDeliverTo.Append(", ")
        sbDeliverTo.Append(dtDeliverTo("CneeAddr1"))
        sbDeliverTo.Append(", ")
        If dtDeliverTo("CneeAddr2").ToString.Trim <> String.Empty Then
            sbDeliverTo.Append(dtDeliverTo("CneeAddr2"))
            sbDeliverTo.Append(", ")
        End If
        If dtDeliverTo("CneeAddr2").ToString.Trim <> String.Empty Then
            sbDeliverTo.Append(dtDeliverTo("CneeAddr3"))
            sbDeliverTo.Append(", ")
        End If
        sbDeliverTo.Append(dtDeliverTo("CneeTown"))
        sbDeliverTo.Append(", ")
        If dtDeliverTo("CneePostCode").ToString.Trim <> String.Empty Then
            sbDeliverTo.Append(dtDeliverTo("CneePostCode"))
            sbDeliverTo.Append(", ")
        End If
        sbDeliverTo.Append(dtDeliverTo("CneeTown"))
        sbDeliverTo.Append(", ")
        sbDeliverTo.Append(dtDeliverTo("CountryName"))
        lblDeliverTo.Text = sbDeliverTo.ToString
        
        sSQL = "SELECT TrackingID 'Tracking ID', Description, CAST(REPLACE(CONVERT(VARCHAR(11),  TrackedOn, 106), ' ', '-') AS varchar(20)) + ' ' + SUBSTRING((CONVERT(VARCHAR(8), TrackedOn, 108)),1,5) 'Tracked On' FROM ConsignmentTrackingStage WHERE ConsignmentKey = " & sConsignmentKey & " ORDER BY [key]"
        dt = ExecuteQueryToDataTable(sSQL)
        gvTracking.DataSource = dt
        gvTracking.DataBind()
        [lblJobNumber].Text = sConsignmentKey
        Call HideAllPanelsAndRows()
        trJobDetails01.Visible = True
        trJobDetails02.Visible = True
        trJobDetails03.Visible = True
        trJobDetails04.Visible = True
        trJobDetails05.Visible = True
        trJobDetails06.Visible = True
    End Sub
    
    Protected Sub btnCancelCourierJobNumber_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call HideAllPanelsAndRows()
        'tbCourierJobNumber.Text = String.Empty
        'pnlCourierJobNo.Visible = False
    End Sub
    
    Protected Sub gvPrintOrders_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs)
        Const COLUMN_STATUS As Int16 = 5
        Dim gvr As GridViewRow = e.Row
        'Dim hidReadyToPrint As HiddenField
        If gvr.RowType = DataControlRowType.DataRow Then
            Dim btnJobReceived As Button = gvr.Cells(0).FindControl("btnJobReceived")
            Dim btnJobPrinted As Button = gvr.Cells(0).FindControl("btnJobPrinted")
            Dim sStatusLevel As String = gvr.Cells(COLUMN_STATUS).Text.Substring(0, 1)
            If IsNumeric(sStatusLevel) Then
                Select Case CInt(sStatusLevel)
                    Case 1
                        btnJobReceived.Visible = True
                        btnJobPrinted.Visible = True
                    Case 2
                        btnJobReceived.Visible = False
                        btnJobPrinted.Visible = True
                    Case 3
                        btnJobReceived.Visible = False
                        btnJobPrinted.Visible = False
                End Select
            End If
        End If
    End Sub
    
    Protected Sub cbIncludeDespatchedJobs_CheckedChanged(sender As Object, e As System.EventArgs)
        Call BindPrintOrdersView()
    End Sub

    Protected Function XsGetDestination(ByVal DataItem As Object) As String
        Dim sbDeliveryAddress As New StringBuilder
        Dim nID As Int32 = DataBinder.Eval(DataItem, "ConsignmentKey")
        Dim sSQL As String = "SELECT CneeName, ISNULL(CneeCtcName,'') 'CneeCtcName', CneeAddr1, ISNULL(CneeAddr2,'') 'CneeAddr2', CneeTown, ISNULL(CneePostCode,'') 'CneePostCode', CountryName FROM Consignment c INNER JOIN Country ctry ON c.CneeCountryKey = ctry.CountryKey WHERE [key] = " & nID
        Dim drDeliveryAddress As DataRow = ExecuteQueryToDataTable(sSQL).Rows(0)
        If drDeliveryAddress("CneeCtcName").ToString.Trim <> String.Empty Then
            sbDeliveryAddress.Append(drDeliveryAddress("CneeCtcName"))
            sbDeliveryAddress.Append("<br />" & Environment.NewLine)
        End If
        sbDeliveryAddress.Append(drDeliveryAddress("CneeName"))
        sbDeliveryAddress.Append("<br />" & Environment.NewLine)
        sbDeliveryAddress.Append(drDeliveryAddress("CneeAddr1"))
        sbDeliveryAddress.Append("<br />" & Environment.NewLine)
        If drDeliveryAddress("CneeAddr2").ToString.Trim <> String.Empty Then
            sbDeliveryAddress.Append(drDeliveryAddress("CneeAddr2"))
            sbDeliveryAddress.Append("<br />" & Environment.NewLine)
        End If
        sbDeliveryAddress.Append(drDeliveryAddress("CneeTown"))
        sbDeliveryAddress.Append("<br />" & Environment.NewLine)
        If drDeliveryAddress("CneePostCode").ToString.Trim <> String.Empty Then
            sbDeliveryAddress.Append(drDeliveryAddress("CneePostCode"))
            sbDeliveryAddress.Append("<br />" & Environment.NewLine)
        End If
        sbDeliveryAddress.Append(drDeliveryAddress("CountryName"))
        XsGetDestination = sbDeliveryAddress.ToString
    End Function

    Protected Function sGetDestination(ByVal DataItem As Object) As String
        Dim sbDeliveryAddress As New StringBuilder
        Dim nID As Int32 = DataBinder.Eval(DataItem, "ConsignmentKey")
        Dim sSQL As String = "SELECT CneeName, ISNULL(CneeCtcName,'') 'CneeCtcName', CneeAddr1, ISNULL(CneeAddr2,'') 'CneeAddr2', CneeTown, ISNULL(CneePostCode,'') 'CneePostCode', CountryName FROM Consignment c INNER JOIN Country ctry ON c.CneeCountryKey = ctry.CountryKey WHERE [key] = " & nID
        Dim drDeliveryAddress As DataRow = ExecuteQueryToDataTable(sSQL).Rows(0)
        sbDeliveryAddress.Append(drDeliveryAddress("CneeTown"))
        sbDeliveryAddress.Append("<br />" & Environment.NewLine)
        'If drDeliveryAddress("CneePostCode").ToString.Trim <> String.Empty Then
        '    sbDeliveryAddress.Append(drDeliveryAddress("CneePostCode"))
        '    sbDeliveryAddress.Append("<br />" & Environment.NewLine)
        'End If
        sbDeliveryAddress.Append(drDeliveryAddress("CountryName"))
        sGetDestination = sbDeliveryAddress.ToString
    End Function

    Protected Sub lnkbtnExportToExcel_Click(sender As Object, e As System.EventArgs)
        Dim sSQL As String
        If cbIncludeDespatchedJobs.Checked Then
            sSQL = "SELECT CAST(REPLACE(CONVERT(VARCHAR(11),  CreatedOn, 106), ' ', '-') AS varchar(20)) + ' ' + SUBSTRING((CONVERT(VARCHAR(8), CreatedOn, 108)),1,5) 'CreatedOn', c.[key] 'ConsignmentKey', AWB, FirstName + ' ' + LastName 'OrderedBy', ISNULL(Misc1, '') 'Turnround', AgentAWB 'CourierAWB', AgentRef 'PrintStatus', ISNULL(SpecialInstructions, '') 'SpecialInstructions', CneeName + ' ' + ISNULL(CneeCtcName,'') + ' ' + CneeAddr1 + ' ' + ISNULL(CneeAddr2,'') + ' ' + CneeTown + ' ' + ISNULL(CneePostCode,'') + ' ' +  CountryName 'Destination' FROM Consignment c INNER JOIN UserProfile up ON c.UserKey = up.[key] INNER JOIN Country ctry ON c.CneeCountryKey = ctry.CountryKey WHERE CreatedOn > GETDATE() - " & ddlPrintOrdersLastNDays.SelectedValue & " AND c.CustomerKey = " & CUSTOMER_JUPITER & " AND (AgentRef LIKE '%PRINT%' OR AgentRef LIKE '%JOB%') AND NOT StateId IN ('CANCELLED') ORDER BY c.[key]"   ' include despatched jobs
        Else
            sSQL = "SELECT CAST(REPLACE(CONVERT(VARCHAR(11),  CreatedOn, 106), ' ', '-') AS varchar(20)) + ' ' + SUBSTRING((CONVERT(VARCHAR(8), CreatedOn, 108)),1,5) 'CreatedOn', c.[key] 'ConsignmentKey', AWB, FirstName + ' ' + LastName 'OrderedBy', ISNULL(Misc1, '') 'Turnround', AgentAWB 'CourierAWB', AgentRef 'PrintStatus', ISNULL(SpecialInstructions, '') 'SpecialInstructions', CneeName + ' ' + ISNULL(CneeCtcName,'') + ' ' + CneeAddr1 + ' ' + ISNULL(CneeAddr2,'') + ' ' + CneeTown + ' ' + ISNULL(CneePostCode,'') + ' ' +  CountryName 'Destination' FROM Consignment c INNER JOIN UserProfile up ON c.UserKey = up.[key] INNER JOIN Country ctry ON c.CneeCountryKey = ctry.CountryKey WHERE CreatedOn > GETDATE() - " & ddlPrintOrdersLastNDays.SelectedValue & " AND c.CustomerKey = " & CUSTOMER_JUPITER & " AND (AgentRef LIKE '%PRINT%' OR AgentRef LIKE '%JOB%') AND NOT (AgentRef LIKE '%DESPATCHED%' OR StateId IN ('CANCELLED')) ORDER BY c.[key]"
        End If
        Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
        
        Response.Clear()
        Response.ContentType = "text/csv"
        Dim sbResponseValue As New StringBuilder
        sbResponseValue.Append("attachment; filename=""")
        sbResponseValue.Append("JupiterOrderList")
        sbResponseValue.Append("_")
        sbResponseValue.Append("EXPORTED")
        sbResponseValue.Append("_")
        sbResponseValue.Append(Format(Date.Now, "ddMMMyyyy_hhmmss"))
        sbResponseValue.Append(".csv")
        sbResponseValue.Append("""")
        Response.AddHeader("Content-Disposition", sbResponseValue.ToString)

        Response.Write("Created On")
        Response.Write(",")
        Response.Write("Original AWB")
        Response.Write(",")
        Response.Write("Current AWB")
        Response.Write(",")
        Response.Write("Ordered By")
        Response.Write(",")
        Response.Write("Turnround")
        Response.Write(",")
        Response.Write("Courier Job #")
        Response.Write(",")
        Response.Write("Last Status")
        Response.Write(",")
        Response.Write("Special Instructions")
        Response.Write(",")
        Response.Write("Destination")
        Response.Write(",")
        Response.Write("Item(s)")
        Response.Write(vbCrLf)
        
        Dim sItem As String
        For Each dr As DataRow In dt.Rows
            For i = 0 To dt.Columns.Count - 1
                sItem = dr(i) & String.Empty
                sItem = sItem.Replace(ControlChars.Quote, ControlChars.Quote & ControlChars.Quote)
                sItem = ControlChars.Quote & sItem & ControlChars.Quote
                Response.Write(sItem)
                Response.Write(",")
            Next
            Response.Write(sGetItems(dr(1)))
            Response.Write(vbCrLf)
        Next
        Response.End()
    End Sub
    
    Protected Function sGetItems(sConsignmentKey As String) As String
        Dim sbItems As New StringBuilder
        Dim sSQL As String = "SELECT ProductCode, ProductDate, ProductDescription, ItemsOut 'Qty' FROM LogisticMovement lm INNER JOIN LogisticProduct lp ON lm.LogisticProductKey = lp.LogisticProductKey WHERE ConsignmentKey = " & sConsignmentKey
        Dim dtItems As DataTable = ExecuteQueryToDataTable(sSQL)
        For Each dr As DataRow In dtItems.Rows
            sbItems.Append(dr("ProductCode"))
            sbItems.Append(" ")
            sbItems.Append(dr("ProductDate"))
            sbItems.Append(" ")
            sbItems.Append(dr("ProductDescription"))
            sbItems.Append(" QTY: ")
            sbItems.Append(dr("Qty"))
            sbItems.Append("; ")
        Next
        sGetItems = sbItems.ToString
    End Function
    
    Protected Sub gvPDFStatus_PageIndexChanging(sender As Object, e As System.Web.UI.WebControls.GridViewPageEventArgs)
        gvPDFStatus.PageIndex = e.NewPageIndex
        Call BindPDFUploadView()
    End Sub
    
    Protected Function EntryEnabled(ByVal DataItem As Object) As String
        Return DataBinder.Eval(DataItem, "ReadyToPrint")
        EntryEnabled = False
        If DataBinder.Eval(DataItem, "ReadyToPrint") > 0 Then
            EntryEnabled = True
        End If
    End Function
        
    Protected Sub cbHide_CheckedChanged(sender As Object, e As System.EventArgs)
        Dim cb As CheckBox = sender
        Dim cell As DataControlFieldCell = cb.Parent
        Dim b As Button = cell.FindControl("btnReadyToPrint")
        Dim sLogisticProductKey As String = b.CommandArgument
        Dim sSQL As String = "UPDATE ClientData_Jupiter_PDFUploads SET Hidden = "
        If cb.Checked Then
            sSQL &= "1"
        Else
            sSQL &= "0"
        End If
        sSQL &= " WHERE LogisticProductKey = " & sLogisticProductKey
        Call ExecuteQueryToDataTable(sSQL)
        Call BindPDFUploadView()
    End Sub
    
    Protected Sub cbIncludeHiddenItems_CheckedChanged(sender As Object, e As System.EventArgs)
        gvPDFStatus.PageIndex = 0
        Call BindPDFUploadView()
    End Sub
    
</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Print Status</title>
</head>
<body>
    <form id="form1" runat="server">
    <main:Header ID="ctlHeader" runat="server" />
    <table style="width: 100%">
        <tr>
            <td style="width: 2%">
                &nbsp;
            </td>
            <td style="width: 26%">
                <asp:Label ID="lblLegendTitle" runat="server" Font-Size="Small" Font-Names="Verdana" Font-Bold="True" ForeColor="Gray">Jupiter Print Status </asp:Label>
            </td>
            <td style="width: 40%">
                &nbsp;
            </td>
            <td style="width: 30%">
                &nbsp;
            </td>
            <td style="width: 2%">
                &nbsp;
            </td>
        </tr>
        <tr>
            <td>
                &nbsp;
            </td>
            <td colspan="2">
                <asp:Label ID="lblLegendLastNPDFUploads1" runat="server" Font-Names="Verdana" Font-Size="XX-Small">Last</asp:Label>
                &nbsp;<asp:DropDownList ID="ddlPDFUploads" runat="server" Font-Names="Verdana" Font-Size="XX-Small" AutoPostBack="True" OnSelectedIndexChanged="ddlPDFUploads_SelectedIndexChanged">
                    <asp:ListItem Selected="True" Value="10">10</asp:ListItem>
                    <asp:ListItem Value="25">25</asp:ListItem>
                    <asp:ListItem Value="50">50</asp:ListItem>
                    <asp:ListItem Value="200">200</asp:ListItem>
                </asp:DropDownList>
                <asp:Label ID="lblLegendLastNPDFUploads2" runat="server" Font-Names="Verdana" Font-Size="XX-Small">PDF uploads (does not include archived products):</asp:Label>
                &nbsp;&nbsp;&nbsp;&nbsp;
                <asp:CheckBox ID="cbIncludeHiddenItems" runat="server" AutoPostBack="True" 
                    Font-Names="Verdana" Font-Size="XX-Small" 
                    oncheckedchanged="cbIncludeHiddenItems_CheckedChanged" 
                    Text="include hidden items" />
            </td>
            <td align="right">
                &nbsp;
                <asp:LinkButton ID="lnkbtnRefresh" runat="server" Font-Names="Verdana" Font-Size="XX-Small" OnClick="lnkbtnRefresh_Click">refresh</asp:LinkButton>
            </td>
            <td>
                &nbsp;
            </td>
        </tr>
        <tr>
            <td>
                &nbsp;
            </td>
            <td colspan="3">
                <asp:GridView ID="gvPDFStatus" runat="server" CellPadding="2" 
                    Font-Names="Verdana" Font-Size="XX-Small" Width="100%" 
                    OnRowDataBound="gvPDFStatus_RowDataBound" AutoGenerateColumns="False" 
                    AllowPaging="True" OnPageIndexChanging="gvPDFStatus_PageIndexChanging">
                    <Columns>
                        <asp:TemplateField>
                            <ItemTemplate>
                                <asp:Label ID="lblSpacer" runat="server" Text="&amp;nbsp;&amp;nbsp;&amp;nbsp;"></asp:Label>
                                <asp:HyperLink ID="hlnk_PDF" runat="server" Width="120px" NavigateUrl='<%# GetHyperLinkNavigateURL(Container.DataItem) %>' Font-Names="Verdana" Font-Size="XX-Small" Target="_blank" ToolTip="click to download PDF file">download PDF</asp:HyperLink>
                                <br />
                                <br />
                                <asp:Button ID="btnReadyToPrint" runat="server" Text="ready to print" Width="120px" OnClick="btnReadyToPrint_Click" CommandArgument='<%# Container.DataItem("LogisticProductKey")%>' />
                                <asp:HiddenField ID="hidReadyToPrint" runat="server" Value='<%# Container.DataItem("ReadyToPrint")%>' />
                                <asp:CheckBox ID="cbHide" runat="server" Text="hide" 
                                    Checked='<%# Container.DataItem("Hidden") %>' 
                                    Visible='<%# EntryEnabled(Container.DataItem) %>' 
                                    oncheckedchanged="cbHide_CheckedChanged" AutoPostBack="True" />
                                <br />
                            </ItemTemplate>
                            <ItemStyle HorizontalAlign="Center" Width="130px" />
                        </asp:TemplateField>
                        <asp:BoundField DataField="ProductCode" HeaderText="Product Code" ReadOnly="True" SortExpression="ProductCode" />
                        <asp:BoundField DataField="ProductDate" HeaderText="Product Date" ReadOnly="True" SortExpression="ProductDate" />
                        <asp:BoundField DataField="ProductDescription" HeaderText="Description" ReadOnly="True" SortExpression="ProductDescription" />
                        <asp:BoundField DataField="PrintType" HeaderText="Type" ReadOnly="True" 
                            SortExpression="PrintType" />
                        <asp:BoundField DataField="ReadyToPrint" HeaderText="Ready To Print?" ReadOnly="True" SortExpression="ReadyToPrint">
                            <ItemStyle HorizontalAlign="Center" />
                        </asp:BoundField>
                    </Columns>
                    <EmptyDataTemplate>
                        (no PDF uploads found)
                    </EmptyDataTemplate>
                    <PagerSettings FirstPageText="First" LastPageText="Last" 
                        Mode="NumericFirstLast" NextPageText="Next" PreviousPageText="Prev" />
                    <PagerStyle HorizontalAlign="Center" Font-Bold="True" Font-Size="Small" />
                </asp:GridView>
            </td>
            <td>
                &nbsp;
            </td>
        </tr>
        <tr>
            <td>
                &nbsp;
            </td>
            <td>
                &nbsp;
            </td>
            <td>
                &nbsp;
            </td>
            <td>
                &nbsp;
            </td>
            <td>
                &nbsp;
            </td>
        </tr>
        <tr>
            <td>
                &nbsp;
            </td>
            <td colspan="3">
                <asp:Label ID="lblLegendPrintOrders" runat="server" Font-Names="Verdana" Font-Size="XX-Small">Print Orders  -  last</asp:Label>
                <asp:DropDownList ID="ddlPrintOrdersLastNDays" runat="server" Font-Names="Verdana" Font-Size="XX-Small" AutoPostBack="True" OnSelectedIndexChanged="ddlPrintOrdersLastNDays_SelectedIndexChanged">
                    <asp:ListItem Selected="True" Value="6">5 days</asp:ListItem>
                    <asp:ListItem Value="11">10 days</asp:ListItem>
                    <asp:ListItem Value="31">30 days</asp:ListItem>
                    <asp:ListItem Value="61">60 days</asp:ListItem>
                </asp:DropDownList>
                &nbsp;<asp:CheckBox ID="cbIncludeDespatchedJobs" runat="server" 
                    Font-Names="Verdana" Font-Size="XX-Small" Text="include despatched jobs" 
                    AutoPostBack="True" oncheckedchanged="cbIncludeDespatchedJobs_CheckedChanged" />
                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                <asp:LinkButton ID="lnkbtnExportToExcel" runat="server" Font-Names="Verdana" 
                    Font-Size="XX-Small" onclick="lnkbtnExportToExcel_Click">export to excel</asp:LinkButton>
                <asp:GridView ID="gvPrintOrders" runat="server" CellPadding="2" Font-Names="Verdana" Font-Size="XX-Small" Width="100%" AutoGenerateColumns="False" OnRowDataBound="gvPrintOrders_RowDataBound">
                    <Columns>
                        <asp:TemplateField>
                            <ItemTemplate>
                                <br />
                                <asp:Button ID="btnJobReceived" runat="server" Text="job received" Width="120px" OnClick="btnJobReceived_Click" CommandArgument='<%# DataBinder.Eval(Container.DataItem, "ConsignmentKey") %>' />
                                <br />
                                <asp:Button ID="btnJobPrinted" runat="server" Text="job printed" Width="120px" OnClick="btnJobPrinted_Click" CommandArgument='<%# DataBinder.Eval(Container.DataItem, "ConsignmentKey") %>' />
                                <br />
                                <asp:Button ID="btnCourierJobNo" runat="server" Text="courier job #" Width="120px" OnClick="btnCourierJobNo_Click" CommandArgument='<%# DataBinder.Eval(Container.DataItem, "ConsignmentKey") %>' />
                                <br />
                                <asp:Button ID="btnJobDespatched" runat="server" Text="job despatched" Width="120px" OnClick="btnJobDespatched_Click" CommandArgument='<%# DataBinder.Eval(Container.DataItem, "ConsignmentKey") %>' />
                            </ItemTemplate>
                            <ItemStyle HorizontalAlign="Center" Width="130px" />
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Order">
                            <ItemTemplate>
                                <asp:LinkButton ID="lnkbtnAWB" runat="server" Text='<%# DataBinder.Eval(Container.DataItem, "AWB") %>' CommandArgument='<%# DataBinder.Eval(Container.DataItem, "ConsignmentKey") %>' Font-Names="Verdana" Font-Size="X-Small" OnClick="lnkbtnAWB_Click" />
                            </ItemTemplate>
                            <ItemStyle HorizontalAlign="Center" />
                        </asp:TemplateField>
                        <asp:BoundField DataField="CreatedOn" HeaderText="Created" ReadOnly="True" SortExpression="CreatedOn" >
                        <ItemStyle HorizontalAlign="Center" />
                        </asp:BoundField>
                        <asp:BoundField DataField="PrintTurnround" HeaderText="Print Turnround" 
                            ReadOnly="True" SortExpression="PrintTurnround" />
                        <asp:BoundField DataField="OrderedBy" HeaderText="Ordered By" ReadOnly="True" SortExpression="OrderedBy" >
                        <ItemStyle HorizontalAlign="Center" />
                        </asp:BoundField>
                        <asp:BoundField DataField="SpecialInstructions" 
                            HeaderText="Special Instructions" ReadOnly="True" 
                            SortExpression="SpecialInstructions" />
                        <asp:TemplateField HeaderText="Destination">
                            <ItemTemplate>
                                <asp:Label ID="lblDeliveryAddress" runat="server" Text='<%# sGetDestination(Container.DataItem) %>' />
                            </ItemTemplate>
                            <ItemStyle HorizontalAlign="Center" />
                        </asp:TemplateField>
                        <asp:BoundField DataField="CourierAWB" HeaderText="Courier Job #" ReadOnly="True" SortExpression="CourierAWB" >
                        <ItemStyle HorizontalAlign="Center" />
                        </asp:BoundField>
                        <asp:BoundField DataField="PrintStatus" HeaderText="Status" ReadOnly="True" SortExpression="PrintStatus" >
                        <ItemStyle HorizontalAlign="Center" />
                        </asp:BoundField>
                    </Columns>
                    <EmptyDataTemplate>
                        (no print orders found)
                    </EmptyDataTemplate>
                </asp:GridView>
            </td>
            <td>
                &nbsp;
            </td>
        </tr>
        <tr>
            <td>
                &nbsp;
            </td>
            <td>
                &nbsp;
            </td>
            <td>
                &nbsp;
            </td>
            <td>
                &nbsp;
            </td>
            <td>
                &nbsp;
            </td>
        </tr>
        <tr id="trJobDetails01" runat="server" visible="false">
            <td>
                &nbsp;
            </td>
            <td>
                <asp:Label ID="lblLegendJobDetails" runat="server" Font-Names="Verdana" Font-Size="XX-Small">Job details for job</asp:Label>
                &nbsp;<asp:Label ID="lblJobNumber" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Font-Bold="True"></asp:Label>
            </td>
            <td>
                &nbsp;
            </td>
            <td>
                &nbsp;
            </td>
            <td>
                &nbsp;
            </td>
        </tr>
        <tr id="trJobDetails02" runat="server" visible="false">
            <td>
                &nbsp;
            </td>
            <td colspan="3">
                <asp:GridView ID="gvJobDetails" runat="server" CellPadding="2" Font-Names="Verdana" Font-Size="XX-Small" Width="100%">
                </asp:GridView>
            </td>
            <td>
                &nbsp;
            </td>
        </tr>
        <tr id="trJobDetails03" runat="server" visible="false">
            <td>
                &nbsp;</td>
            <td>
                <asp:Label ID="Label2" runat="server" Font-Names="Verdana" Font-Size="XX-Small">Deliver To:</asp:Label>
            </td>
            <td>
                &nbsp;</td>
            <td>
                &nbsp;</td>
            <td>
                &nbsp;</td>
        </tr>
        <tr id="trJobDetails04" runat="server" visible="false">
            <td>
                &nbsp;</td>
            <td colspan="3">
                <asp:Label ID="lblDeliverTo" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Font-Bold="True"></asp:Label>
            </td>
            <td>
                &nbsp;</td>
        </tr>
        <tr id="trJobDetails05" runat="server" visible="false">
            <td>
                &nbsp;
            </td>
            <td>
                <asp:Label ID="Label1" runat="server" Font-Names="Verdana" Font-Size="XX-Small">Tracking:</asp:Label>
            </td>
            <td>
                &nbsp;
            </td>
            <td>
                &nbsp;
            </td>
            <td>
                &nbsp;
            </td>
        </tr>
        <tr id="trJobDetails06" runat="server" visible="false">
            <td>
                &nbsp;
            </td>
            <td colspan="3">
                <asp:GridView ID="gvTracking" runat="server" CellPadding="2" Font-Names="Verdana" Font-Size="XX-Small" Width="100%">
                    <EmptyDataTemplate>
                        (no tracking found)
                    </EmptyDataTemplate>
                </asp:GridView>
            </td>
            <td>
                &nbsp;
            </td>
        </tr>
        <tr>
            <td>
                &nbsp;
            </td>
            <td>
                &nbsp;
            </td>
            <td>
                &nbsp;
            </td>
            <td>
                &nbsp;
            </td>
            <td>
                &nbsp;
            </td>
        </tr>
    </table>
    <asp:Panel ID="pnlCourierJobNo" runat="server" Visible="false" Width="100%">
        <table style="width: 100%">
            <tr>
                <td style="width: 2%">
                    &nbsp;
                </td>
                <td style="width: 26%">
                    <asp:Label ID="lblLegendJobDetails0" runat="server" Font-Names="Verdana" Font-Size="XX-Small">Job</asp:Label>
                    &nbsp;<asp:Label ID="lblConsignmentNumberForCourierJobNo" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small"></asp:Label>
                </td>
                <td style="width: 40%">
                    &nbsp;
                </td>
                <td style="width: 30%">
                    &nbsp;
                </td>
                <td style="width: 2%">
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td>
                    &nbsp;
                </td>
                <td align="right">
                    &nbsp;
                    <asp:Label ID="lblLegendCourierJobNo" runat="server" Font-Names="Verdana" Font-Size="XX-Small">Courier job number:</asp:Label>
                </td>
                <td colspan="2">
                    <asp:TextBox ID="tbCourierJobNumber" runat="server" Font-Names="Verdana" Font-Size="X-Small" Width="150px" />
                    &nbsp;<asp:Button ID="btnSaveCourierJobNumber" runat="server" OnClick="btnSaveCourierJobNumber_Click" Text="save" Width="80px" />
                    &nbsp;<asp:Button ID="btnCancelCourierJobNumber" runat="server" Text="cancel" OnClick="btnCancelCourierJobNumber_Click" Width="80px" />
                    &nbsp;
                    <asp:CheckBox ID="cbOverrideValidation" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="override validation" />
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
        </table>
    </asp:Panel>
    </form>
</body>
</html>
