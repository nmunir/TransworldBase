<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>
<script runat="server">

    ' TEST VALUES
    ' 9000000 - OK
    ' 9000001 - BAD_VALUE
    ' 9000002 - DUPLICATE
    ' 9000003 - ALREADY_SET
    ' 9000004 - TOO_LATE
    ' 9000005 - JOB_CANCELLED
    ' 9xxxxxx - NO_MATCH
    
    ' EMPTY_COURIER_AWB
    ' INVALID_COURIER_AWB
    ' UNEXPECTED_COURIER_AWB_LENGTH
    
    ' BAD_QUERYSTRING
    
    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString()

    Protected Sub Page_Load(ByVal Source As Object, ByVal E As EventArgs)
        If Not IsPostBack Then
            If IsNothing(Request.QueryString("PODAWB")) Or IsNothing(Request.QueryString("CourierAWB")) Then
                litReturnValue.Text = "BAD_QUERYSTRING_(USE_PODAWB_AND_CourierAWB)"
                Exit Sub
            End If
            Dim sPODAWB As String = Request.QueryString("PODAWB")
            Dim sCourierAWB As String = Request.QueryString("CourierAWB")
            If sPODAWB.StartsWith("9") Then
                litReturnValue.Text = ReturnTestValues(sPODAWB, sCourierAWB)
            Else
                litReturnValue.Text = ValidatePODAWB(sPODAWB, sCourierAWB)
                If litReturnValue.Text = "OK" Or litReturnValue.Text = "TOO_LATE" Then
                    Dim sSQL As String = "UPDATE Consignment SET AgentAWB = '" & sCourierAWB & "' WHERE [key] = " & sPODAWB
                    Call ExecuteQueryToDataTable(sSQL)
                    Call LogJupiterAuditEvent("COURIER_JOBNO_LOGGED", "Courier job number entered: " & sCourierAWB, , sPODAWB)
                End If
            End If
        End If
    End Sub
    
    Protected Function ReturnTestValues(ByVal sAWB As String, ByVal sCourierAWB As String) As String
        Select sAWB
            Case 9000000
                Return "OK"
            Case 9000001
                Return "BAD_VALUE"
            Case 9000002
                Return "DUPLICATE"
            Case 9000003
                Return "ALREADY_SET"
            Case 9000004
                Return "TOO_LATE"
            Case 9000005
                Return "JOB_CANCELLED"
            Case Else
                Return "NO_MATCH"
        End Select
    End Function
    
    Protected Function ValidatePODAWB(ByVal sAWB As String, ByVal sCourierAWB As String) As String
        If sCourierAWB.Trim = String.Empty Then
            Return "EMPTY_COURIER_AWB"
        End If
        If Not IsNumeric(sCourierAWB.Trim) Then
            Return "INVALID_COURIER_AWB"
        End If
        If Not sCourierAWB.Trim.Length = 8 Then
            Return "UNEXPECTED_COURIERAWB_LENGTH"
        End If
        If Not (sAWB.Trim.Length = 7 And IsNumeric(sAWB)) Then
            Return "BAD_VALUE"
        End If
        Dim dt As DataTable = ExecuteQueryToDataTable("SELECT CustomerKey, ISNULL(AgentAWB, ''), ISNULL(AgentRef,''), StateId FROM Consignment WHERE [key] = " & sAWB)
        If dt.Rows.Count = 0 Then
            Return "NO_MATCH"
        End If
        Dim dr = dt.Rows(0)
        If sCourierAWB = dr("AgentAWB") Then
            Return "DUPLICATE"
        End If
        If dr("AgentAWB") <> "" Then
            Return "ALREADY_SET"
        End If
        If dr("StateId") <> "CANCELLED" Then
            Return "JOB_CANCELLED"
        End If
        If dr("StateId") <> "WITH_OPERATIONS" Then
            Return "TOO_LATE"
        End If
        Return "OK"
    End Function
    
    Protected Sub LogJupiterAuditEvent(ByVal sEventCode As String, ByVal sEventDescription As String, Optional ByVal nProductKey As Int32 = 0, Optional ByVal nConsignmentKey As Int32 = 0)
        Dim sSQL As String = "INSERT INTO ClientData_Jupiter_AuditTrail (EventCode, EventDescription, ProductKey, ConsignmentKey, EventDateTime, EventAuthor) VALUES ('" & sEventCode & "', '" & sEventDescription & "', " & nProductKey & ", " & nConsignmentKey & ", GETDATE(), " & Session("UserKey") & ")"
        Call JupiterNotification(sEventCode, sEventDescription)
        Call ExecuteQueryToDataTable(sSQL)
        If nConsignmentKey > 0 Then
            sSQL = "INSERT INTO ConsignmentTrackingStage (ConsignmentKey, TrackingID, Location, Description, TrackedOn) VALUES (" & nConsignmentKey & ", '" & sEventCode & "', '', '" & sEventDescription.Replace("'", "''") & "', GETDATE())"
            Call ExecuteQueryToDataTable(sSQL)
        End If
    End Sub

    Protected Sub JupiterNotification(ByVal sEventCode As String, ByVal sEventDescription As String)
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
        sbMessage.Append("Please do not reply to this email as replies are not monitored.  Thank you.")
        sbMessage.Append(Environment.NewLine)
        sbMessage.Append("Transworld")
        Dim sPlainTextBody As String = sbMessage.ToString
        Dim sHTMLBody As String = sbMessage.ToString.Replace(Environment.NewLine, "<br />" & Environment.NewLine)
        For Each dr As DataRow In dt.Rows
            Call SendMail("JUPITER_EVENT", dr(0), "Jupiter Event Notification - " & sEventCode, sPlainTextBody, sHTMLBody)
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

</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Link AWB</title>
</head>
<body>
    <form id="form1" runat="server">
    <asp:Literal ID="litReturnValue" runat="server"></asp:Literal>
    </form>
    </body>
</html>
