<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="Microsoft.VisualBasic" %>
<script runat="server">

    Const CUSTOMER_POSITIVENOISE As Int32 = 821
    'Const CUSTOMER_POSITIVENOISE As Int32 = 16
    
    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsNumeric(Session("CustomerKey")) Then
            Response.Write("You must be logged in to run this report.")
            pnlMain.Visible = False
            Exit Sub
        Else
            pnlMain.Visible = True
        End If

        If Not IsPostBack Then
            If Request.Cookies("Transworld") Is Nothing Then
                Call CreateNextStartConsignmentCookie()
                tbStartConsignment.Text = "1"
            Else
                Dim sNextStartConsignment As String = Request.Cookies("Transworld")("NextStartConsignment") & String.Empty
                If IsNumeric(sNextStartConsignment) Then
                    tbStartConsignment.Text = sNextStartConsignment
                End If
            End If
        End If
        Call SetTitle()
    End Sub

    Protected Sub SetTitle()
        Dim sTitle As String = Session("SiteTitle")
        If sTitle <> String.Empty Then
            sTitle += " - "
        End If
        Page.Header.Title = sTitle & "Positive Noise Consignment Report"
    End Sub
    
    Protected Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim sm As New ScriptManager
        sm.ID = "ScriptMgr"
        Try
            PlaceHolderForScriptManager.Controls.Add(sm)
        Catch ex As Exception
        End Try
    End Sub

    Protected Sub btnGo_Click(sender As Object, e As System.EventArgs)
        If Not IsNumeric(tbStartConsignment.Text) Then
            WebMsgBox.Show("Please enter a valid numeric value for the start consignment.")
            Exit Sub
        End If
        Call CalculateNextStartConsignment()
        Call Export()
    End Sub
    
    Protected Sub CalculateNextStartConsignment()
        Dim sSQL As String = "SELECT TOP 1 [key] FROM Consignment WHERE CustomerKey = " & CUSTOMER_POSITIVENOISE & " AND [key] >= " & tbStartConsignment.Text & " AND StateId = 'WITH_OPERATIONS' ORDER BY [key] DESC"
        Dim dtNextStartConsignment As DataTable = ExecuteQueryToDataTable(sSQL)
        If dtNextStartConsignment.Rows.Count = 1 Then
            Dim sNextStartConsignment As String = dtNextStartConsignment.Rows(0).Item(0)
            Dim c As HttpCookie = New HttpCookie("Transworld")
            c.Values.Add("NextStartConsignment", sNextStartConsignment)
            c.Expires = DateTime.Now.AddDays(365)
            Response.Cookies.Add(c)
        End If
    End Sub
    
    Protected Sub Export()
        Dim sSQL As String = "SELECT AWB 'Consignment #', CreatedOn 'Created On', CneeName 'Consignee', CneeAddr1 'Addr 1', CneeTown 'Town', StateId 'Status', ISNULL(CustomerRef1, '') 'PNOrderRef', ISNULL(CustomerRef2, '') 'MarketOrderRef', ISNULL(Misc1, '') 'Market', ISNULL(Misc2, '') 'ExpeditedPostalService', ISNULL(ExternalSystemId, '') 'CN22 / User' FROM Consignment WHERE CustomerKey = " & CUSTOMER_POSITIVENOISE & " AND [key] >= " & tbStartConsignment.Text & " ORDER BY [key]"
        Dim dtConsignments As DataTable = ExecuteQueryToDataTable(sSQL)
        If dtConsignments.Rows.Count > 0 Then
            Response.Clear()
            Response.ContentType = "text/csv"
            Response.AddHeader("Content-Disposition", "attachment; filename=PositiveNoiseConsignments_ " & Format(Date.Now, "yyMMMddhhmmss") & ".csv")
            Dim sItem As String
            Dim IgnoredItems As New ArrayList
            'IgnoredItems.Add("UserKey")
            For Each dc As DataColumn In dtConsignments.Columns  ' write column header
                If Not IgnoredItems.Contains(dc.ColumnName) Then
                    Response.Write(dc.ColumnName)
                    Response.Write(",")
                End If
            Next
            Response.Write(vbCrLf)
    
            For Each dr As DataRow In dtConsignments.Rows
                For Each dc As DataColumn In dtConsignments.Columns
                    If Not IgnoredItems.Contains(dc.ColumnName) Then
                        sItem = dr(dc.ColumnName).ToString
                        sItem = sItem.Replace(ControlChars.Quote, ControlChars.Quote & ControlChars.Quote)
                        sItem = ControlChars.Quote & sItem & ControlChars.Quote
                        Response.Write(sItem)
                        Response.Write(",")
                    End If
                Next
                Response.Write(vbCrLf)
            Next
            Response.End()
        Else
            lblError.Text = "... no data found"
        End If
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

    Protected Sub CreateNextStartConsignmentCookie()
        Dim c As HttpCookie = New HttpCookie("Transworld")
        c.Values.Add("NextStartConsignment", String.Empty)
        c.Expires = DateTime.Now.AddDays(365)
        Response.Cookies.Add(c)
    End Sub


</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Positive Noise Consignment Report</title>
</head>
<body>
    <form id="frmOrderProcessor" runat="Server">
    <%-- <main:Header id="ctlHeader" runat="server"/>--%>
    <asp:PlaceHolder ID="PlaceHolderForScriptManager" runat="server" />
    <asp:Panel ID="pnlMain" Width="100%" runat="server" Visible="false">
        <asp:Label ID="lblLegendOrderProcessor" runat="server" Font-Names="Verdana" Font-Size="X-Small"
            Text="Positive Noise Consignment Report" Font-Bold="True" />
        <br />
        <br />
        <asp:Label ID="lblLegendSelect" Text="Start consignment #:" runat="server" Font-Bold="False"
            Font-Names="Verdana" Font-Size="XX-Small" />
        <asp:TextBox ID="tbStartConsignment" runat="server" Font-Names="Verdana" Font-Size="XX-Small" />
        &nbsp;<asp:Button ID="btnGo" runat="server" OnClick="btnGo_Click" Text="go" />
        <br />
        <br />
        <asp:Label ID="lblError" runat="server" Font-Names="Verdana" Font-Size="X-Small" />
    </asp:Panel>
    </form>
</body>
</html>
