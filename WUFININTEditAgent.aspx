<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ Register Assembly="Telerik.Web.UI" Namespace="Telerik.Web.UI" TagPrefix="telerik" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

    Const ITEMS_PER_REQUEST As Integer = 30

    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        'If Not IsNumeric(Session("UserKey")) Then
        '    Server.Transfer("session_expired.aspx")
        'End If
        If Not IsPostBack Then
            Call SetTitle()
            rcbAgents.Focus()
        End If
    End Sub

    Protected Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim sm As New ScriptManager
        sm.ID = "ScriptMgr"
        Try
            PlaceHolderForScriptManager.Controls.Add(sm)
        Catch ex As Exception
        End Try
    End Sub
    
    Protected Sub SetTitle()
        Dim sTitle As String = Session("SiteTitle")
        If sTitle <> String.Empty Then
            sTitle += " - "
        End If
        Page.Header.Title = sTitle & "Edit FININT Agent"
    End Sub

    Protected Sub rcbAgents_SelectedIndexChanged(ByVal o As Object, ByVal e As Telerik.Web.UI.RadComboBoxSelectedIndexChangedEventArgs)
        Dim sAgentID As String = e.Value
        Dim sSQL As String = "SELECT * FROM ClientData_WU_LegacyNetwork WHERE AgentID = " & sAgentID
        Dim dtAgent As DataTable = ExecuteQueryToDataTable(sSQL)
        If dtAgent.Rows.Count = 1 Then
            Dim dr As DataRow = dtAgent.Rows(0)
            lblAgentID.Text = dr("AgentID").ToString.Trim
            tbAccountNumber.Text = dr("AccountNumber").ToString.Trim
            tbLegalName.Text = dr("LegalName").ToString.Trim
            tbLocationName.Text = dr("LocationName").ToString.Trim
            tbAddressLine1.Text = dr("AddressLine1").ToString.Trim
            tbAddressLine2.Text = dr("AddressLine2").ToString.Trim
            tbTownCity.Text = dr("CityName").ToString.Trim
            tbCounty.Text = dr("Province/County/State").ToString.Trim
            tbPostCode.Text = dr("PostalCode").ToString.Trim
            tbLocationName.Focus()
            tabForm.Visible = True
            lblMessage.Text = String.Empty
        Else
            WebMsgBox.Show("Could not retrieve agent details - please inform development.")
        End If
    End Sub
    
    Protected Sub rcbAgents_ItemsRequested(ByVal o As Object, ByVal e As Telerik.Web.UI.RadComboBoxItemsRequestedEventArgs)
        Dim s As String = e.Text
        Dim IListOfParams As New List(Of SqlParameter)
        Dim paramSearch As New SqlParameter("@Search", SqlDbType.VarChar)
        paramSearch.Value = e.Text
        IListOfParams.Add(paramSearch)
        Dim data As DataTable = ExecuteStoredProcedureToDataTable("WU_GetAllFrom_FININT_Agents", IListOfParams)
        Dim itemOffset As Integer = e.NumberOfItems
        Dim endOffset As Integer = Math.Min(itemOffset + ITEMS_PER_REQUEST, data.Rows.Count)
        e.EndOfItems = endOffset = data.Rows.Count
        For i As Int32 = itemOffset To endOffset - 1
            Dim rcb As New RadComboBoxItem
            rcb.Text = data.Rows(i)("AgentID").ToString() + " " + data.Rows(i)("LocationName").ToString()
            rcb.Value = data.Rows(i)("AgentID").ToString()
            rcbAgents.Items.Add(rcb)
            Dim lblAgent As Label = rcb.FindControl("lblAgent")
            lblAgent.Text = data.Rows(i)("AgentID").ToString() + " " + data.Rows(i)("LocationName").ToString()
        Next
        e.Message = GetStatusMessage(endOffset, data.Rows.Count)
    End Sub

    Private Shared Function GetStatusMessage(ByVal nOffset As Integer, ByVal nTotal As Integer) As String
        If nTotal <= 0 Then
            Return "No matches"
        End If
        If nOffset <= ITEMS_PER_REQUEST Then
            GetStatusMessage = "Click for more items"
        End If
        If nOffset = nTotal Then
            GetStatusMessage = "No more items"
        Else
            GetStatusMessage = "Click for more items"
        End If
    End Function
    
    Public Shared Function CheckNull(ByVal DataItem As Object) As Object
        CheckNull = Nothing
        If (DataItem IsNot Nothing AndAlso Not IsDBNull(DataItem)) Then
            CheckNull = DataItem
        End If
    End Function

    Protected Function ExecuteStoredProcedureToDataTable(ByVal sp_name As String, Optional ByVal IListPrams As List(Of SqlParameter) = Nothing) As DataTable
        Dim oDataTable As New DataTable
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter(sp_name, oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        If Not IListPrams Is Nothing AndAlso IListPrams.Count > 0 Then
            oAdapter.SelectCommand.Parameters.AddRange(IListPrams.ToArray)
        End If
        Try
            oAdapter.Fill(oDataTable)
        Catch ex As Exception
            WebMsgBox.Show(ex.Message.ToString())
        End Try
        ExecuteStoredProcedureToDataTable = oDataTable
    End Function
    
    Protected Function ExecuteQueryToDataTable(ByVal sQuery As String) As DataTable
        Dim oDataTable As New DataTable
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter(sQuery, oConn)
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

    Protected Sub btnUpdate_Click(sender As Object, e As System.EventArgs)
        If Not IsNumeric(tbAccountNumber.Text) Then
            WebMsgBox.Show("Account Number must be numeric")
            Exit Sub
        End If
        
        Dim sbSQL As New StringBuilder
        sbSQL.Append("UPDATE ClientData_WU_LegacyNetwork ")
        sbSQL.Append("SET ")
        sbSQL.Append("AccountNumber")
        sbSQL.Append(" = ")
        sbSQL.Append(tbAccountNumber.Text.Replace("'", "''"))
        sbSQL.Append(", ")

        sbSQL.Append("LegalName")
        sbSQL.Append(" = '")
        sbSQL.Append(tbLegalName.Text.Replace("'", "''"))
        sbSQL.Append("', ")

        sbSQL.Append("LocationName")
        sbSQL.Append(" = '")
        sbSQL.Append(tbLocationName.Text.Replace("'", "''"))
        sbSQL.Append("', ")

        sbSQL.Append("AddressLine1")
        sbSQL.Append(" = '")
        sbSQL.Append(tbAddressLine1.Text.Replace("'", "''"))
        sbSQL.Append("', ")

        sbSQL.Append("AddressLine2")
        sbSQL.Append(" = '")
        sbSQL.Append(tbAddressLine2.Text.Replace("'", "''"))
        sbSQL.Append("', ")

        sbSQL.Append("CityName")
        sbSQL.Append(" = '")
        sbSQL.Append(tbTownCity.Text.Replace("'", "''"))
        sbSQL.Append("', ")

        sbSQL.Append("[Province/County/State]")
        sbSQL.Append(" = '")
        sbSQL.Append(tbCounty.Text.Replace("'", "''"))
        sbSQL.Append("', ")

        sbSQL.Append("PostalCode")
        sbSQL.Append(" = '")
        sbSQL.Append(tbPostCode.Text.Replace("'", "''"))
        sbSQL.Append("' ")

        sbSQL.Append("WHERE AgentID = ")
        sbSQL.Append(lblAgentID.Text)
        Call ExecuteQueryToDataTable(sbSQL.ToString)
        lblMessage.Text = "Updated agent " & lblAgentID.Text & " " & tbLocationName.Text
        Call ClearForm()
    End Sub

    Protected Sub btnCancel_Click(sender As Object, e As System.EventArgs)
        Call ClearForm()
    End Sub
    
    Protected Sub ClearForm()
        lblAgentID.Text = String.Empty
        tbAccountNumber.Text = String.Empty
        tbLegalName.Text = String.Empty
        tbLocationName.Text = String.Empty
        tbAddressLine1.Text = String.Empty
        tbAddressLine2.Text = String.Empty
        tbTownCity.Text = String.Empty
        tbCounty.Text = String.Empty
        tbPostCode.Text = String.Empty
        rcbAgents.Text = String.Empty
        rcbAgents.Focus()
        tabForm.Visible = False
    End Sub
    
</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <main:Header ID="ctlHeader" runat="server" />
    <asp:PlaceHolder ID="PlaceHolderForScriptManager" runat="server"/>
    <div style="font-size: small; font-family: Verdana">
        <strong>&nbsp;Edit FININT Agent</strong>
        <br />
        <br />
        &nbsp;<asp:Label ID="Label1" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Search Agent Number / Agent Name / Postcode:" />
        &nbsp;
        <telerik:RadComboBox ID="rcbAgents" runat="server" Width="300px" Height="500" Font-Names="Arial"
            EmptyMessage="- enter Agent Number, Agent Name or Postcode" Font-Size="X-Small" Font-Bold="true"
            OnSelectedIndexChanged="rcbAgents_SelectedIndexChanged" AutoPostBack="True" HighlightTemplatedItems="true"
            CausesValidation="False" EnableLoadOnDemand="True" OnItemsRequested="rcbAgents_ItemsRequested"
            EnableVirtualScrolling="True" ShowMoreResultsBox="True" Filter="Contains">
            <ItemTemplate>
                <table>
                    <tr>
                        <asp:Label ID="lblAgent" runat="server" />
                    </tr>
                </table>
            </ItemTemplate>
        </telerik:RadComboBox>
        &nbsp;
                    <asp:Label ID="lblMessage" runat="server" Font-Names="Verdana" Font-Size="Small"
                        Font-Bold="True" />
        <br />
        <br />
        <table id="tabForm" runat="server" visible="false" style="width: 100%">
            <tr>
                <td style="width: 20%" align="right">
                    <asp:Label ID="lblLegendAgentID" runat="server" Font-Names="Verdana" Font-Size="Small"
                        Text="Agent ID:" />
                </td>
                <td style="width: 80%">
                    <asp:Label ID="lblAgentID" runat="server" Font-Names="Verdana" Font-Size="Small"
                        Text="" />
                </td>
            </tr>
            <tr>
                <td align="right">
                    <asp:Label ID="Label3" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Account Number:" />
                </td>
                <td>
                    <asp:TextBox ID="tbAccountNumber" runat="server" Font-Bold="True" Font-Names="Verdana"
                        Font-Size="XX-Small" Width="250px" MaxLength="10" />
                </td>
            </tr>
            <tr>
                <td align="right">
                    <asp:Label ID="Label4" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Legal Name:" />
                </td>
                <td>
                    <asp:TextBox ID="tbLegalName" runat="server" Font-Bold="True" Font-Names="Verdana"
                        Font-Size="XX-Small" Width="250px" MaxLength="50" />
                </td>
            </tr>
            <tr>
                <td align="right">
                    <asp:Label ID="Label5" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Location Name:" />
                </td>
                <td>
                    <asp:TextBox ID="tbLocationName" runat="server" Font-Bold="True" Font-Names="Verdana"
                        Font-Size="XX-Small" Width="250px" MaxLength="50" />
                </td>
            </tr>
            <tr>
                <td align="right">
                    <asp:Label ID="Label6" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Address Line 1:" />
                </td>
                <td>
                    <asp:TextBox ID="tbAddressLine1" runat="server" Font-Bold="True" Font-Names="Verdana"
                        Font-Size="XX-Small" Width="250px" MaxLength="50" />
                </td>
            </tr>
            <tr>
                <td align="right">
                    <asp:Label ID="Label7" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Address Line 2:" />
                </td>
                <td>
                    <asp:TextBox ID="tbAddressLine2" runat="server" Font-Bold="True" Font-Names="Verdana"
                        Font-Size="XX-Small" Width="250px" MaxLength="50" />
                </td>
            </tr>
            <tr>
                <td align="right">
                    <asp:Label ID="Label8" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Town / City:" />
                </td>
                <td>
                    <asp:TextBox ID="tbTownCity" runat="server" Font-Bold="True" Font-Names="Verdana"
                        Font-Size="XX-Small" Width="250px" MaxLength="50" />
                </td>
            </tr>
            <tr>
                <td align="right">
                    <asp:Label ID="Label9" runat="server" Font-Names="Verdana" Font-Size="Small" Text="County:" />
                </td>
                <td>
                    <asp:TextBox ID="tbCounty" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small"
                        Width="250px" MaxLength="50" />
                </td>
            </tr>
            <tr>
                <td align="right">
                    <asp:Label ID="Label10" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Post code:" />
                </td>
                <td>
                    <asp:TextBox ID="tbPostCode" runat="server" Font-Bold="True" Font-Names="Verdana"
                        Font-Size="XX-Small" Width="100px" MaxLength="10" />
                </td>
            </tr>
            <tr>
                <td align="right">
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td align="right">
                    &nbsp;
                </td>
                <td>
                    <asp:Button ID="btnUpdate" runat="server" Text="Update" Width="98px" OnClick="btnUpdate_Click" />
                    &nbsp;<asp:Button ID="btnCancel" runat="server" Text="Cancel" OnClick="btnCancel_Click" />
                    &nbsp;&nbsp;
                    </td>
            </tr>
        </table>
        <br />
    </div>
    </form>
</body>
</html>
