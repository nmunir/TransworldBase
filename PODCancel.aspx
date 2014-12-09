<%@ Page Language="VB" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">
    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString()

    Protected Sub Page_Load(ByVal Source As Object, ByVal E As EventArgs)
        Call ExecuteNonQuery("INSERT INTO AAA_Debug (Result) VALUES ('PODContinue called at " & Now.ToShortDateString & " " & Now.ToShortTimeString & "')")

        If Request.QueryString("TransactionGUID") Is Nothing Then
            WebMsgBox.Show("Parameter TransactionGUID not supplied in query string")
            lblErrorMessage.Text = "Parameter TransactionGUID not supplied in query string"
            Call ExecuteNonQuery("INSERT INTO AAA_Debug (Result) VALUES ('Parameter TransactionGUID not supplied in query string')")
            Exit Sub
        Else
            Call ExecuteNonQuery("INSERT INTO AAA_Debug (Result) VALUES ('TransactionGUID: '" & Request.QueryString("TransactionGUID") & "')")        ' DEBUG
        End If

        If Request.QueryString("Status") Is Nothing Then
            WebMsgBox.Show("Status parameter not supplied in query string")
            lblErrorMessage.Text = "Status parameter not supplied in query string"
            Call ExecuteNonQuery("INSERT INTO AAA_Debug (Result) VALUES ('Status parameter not supplied in query string')")
            Exit Sub
        Else
            Call ExecuteNonQuery("INSERT INTO AAA_Debug (Result) VALUES ('Status: '" & Request.QueryString("Status") & "')")        ' DEBUG
        End If
                
        Dim sTransactionGUID As String = Request.QueryString("TransactionGUID").Trim
        Dim sStatus As String = Request.QueryString("Status").Trim.ToLower

        Dim oDataTable As DataTable = ExecuteQueryToDataTable("SELECT * FROM OnDemandTransactionStatus WHERE TransactionGUID = '" & sTransactionGUID & "'")
        If oDataTable.Rows.Count = 1 Then
            Call ExecuteNonQuery("INSERT INTO AAA_Debug (Result) VALUES ('Retrieving OnDemandTransactionStatus record')")        ' DEBUG
            Dim dr As DataRow = oDataTable.Rows(0)
            If sStatus = "cancel" Then
                Call ExecuteNonQuery("INSERT INTO AAA_Debug (Result) VALUES ('Status is OK')")        ' DEBUG
                pnlSuccess.Visible = True
                pnlError.Visible = False
                If dr("TransactionStatus").ToString.ToLower = "start" Then
                    lblLegendPreviousChangesSaved.Visible = False
                    lblLegendNoChangesYetMade.Visible = True
                ElseIf dr("TransactionStatus").ToString.ToLower = "edited" Then
                    lblLegendPreviousChangesSaved.Visible = True
                    lblLegendNoChangesYetMade.Visible = False
                Else
                    Call ReportError("Unexpected status value returned")
                End If
            Else
                Call ReportError("Error: bad status value returned from call to customisation engine")
                Call ExecuteNonQuery("INSERT INTO AAA_Debug (Result) VALUES ('Status NOT OK')")        ' DEBUG
            End If
        ElseIf oDataTable.Rows.Count = 0 Then
            Call ReportError("Error: No matching entry found for Transaction GUID " & sTransactionGUID & " in OnDemandTransactionStatus")
        Else
            Call ReportError("Error: More than one matching entry found for Transaction GUID " & sTransactionGUID & " in OnDemandTransactionStatus")
        End If
    End Sub

    Protected Sub ReportError(ByVal sMessage As String)
        WebMsgBox.Show(sMessage)
        lblErrorMessage.Text = sMessage
        Call ExecuteNonQuery("INSERT INTO AAA_Debug (Result) VALUES ('" & sMessage & "')")
        pnlSuccess.Visible = False
        pnlError.Visible = True
    End Sub
    
    Protected Function ExecuteNonQuery(ByVal sQuery As String) As Boolean
        ExecuteNonQuery = False
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand
        Try
            oConn.Open()
            oCmd = New SqlCommand(sQuery, oConn)
            oCmd.ExecuteNonQuery()
            ExecuteNonQuery = True
        Catch ex As Exception
            WebMsgBox.Show("Error in ExecuteNonQuery executing: " & sQuery & " : " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Function
   
    Protected Function ExecuteQueryToListItemCollection(ByVal sQuery As String, ByVal sTextFieldName As String, ByVal sValueFieldName As String) As ListItemCollection
        Dim oListItemCollection As New ListItemCollection
        Dim oDataReader As SqlDataReader = Nothing
        Dim oConn As New SqlConnection(gsConn)
        Dim sTextField As String
        Dim sValueField As String
        Dim oCmd As SqlCommand = New SqlCommand(sQuery, oConn)
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            If oDataReader.HasRows Then
                While oDataReader.Read
                    If Not IsDBNull(oDataReader(sTextFieldName)) Then
                        sTextField = oDataReader(sTextFieldName)
                    Else
                        sTextField = String.Empty
                    End If
                    If Not IsDBNull(oDataReader(sValueFieldName)) Then
                        sValueField = oDataReader(sValueFieldName)
                    Else
                        sValueField = String.Empty
                    End If
                    oListItemCollection.Add(New ListItem(sTextField, sValueField))
                End While
            End If
        Catch ex As Exception
            WebMsgBox.Show("Error in ExecuteQueryToListItemCollection executing: " & sQuery & " : " & ex.Message)
        Finally
            oConn.Close()
        End Try
        ExecuteQueryToListItemCollection = oListItemCollection
    End Function

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

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Please close this window</title>
</head>
<body>
    <form id="form1" runat="server">
        <div style="text-align: center">
            <br />
            <br />
            <asp:Panel ID="pnlSuccess" Visible="false" runat="server" Width="100%">
                <asp:Label ID="Label1" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="X-Small" Text="You cancelled this editing session." />
                <br />
                <br />
                <asp:Label ID="lblLegendPreviousChangesSaved" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="X-Small" Visible="false" Text="Any changes you made have been lost, but changes you made previously have been saved." />
                <asp:Label ID="lblLegendNoChangesYetMade" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="X-Small" Visible="false" Text="No changes have yet been made to this document." />
                <br />
                <br />
                <asp:Label ID="Label3" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="X-Small" Text="Please close this window and continue your order in the main window." />
            </asp:Panel>
            <asp:Panel ID="pnlError" Visible="true" runat="server" Width="100%">
                <br />
                <br />
                <br />
                <br />
                <br />
                <asp:Label ID="lblErrorMessage" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="X-Small" ForeColor="Red" /></asp:Panel>
            <br />
            <br />
            <br />
            <asp:Button ID="btnCloseWindow" runat="server" Text="close window" OnClientClick="javascript: self.close ()" />
        </div>
    </form>
</body>
</html>
