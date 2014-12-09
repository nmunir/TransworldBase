<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Web.UI" %>
<%@ Register Assembly="Telerik.Web.UI" Namespace="Telerik.Web.UI" TagPrefix="telerik" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<style type="text/css">
    label
    {
        color: Red;
        font-weight: normal;
    }
</style>
<script runat="server">    
   
    Private Shared gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    Const CUSTOMER_KEY As Int32 = 579
    Const CUSTOMER_KEY_WUIRE As Int32 = 686
    Const CUSTOMER_KEY_WUFIN As Int32 = 798
    Const CUSTOMER_KEY_WURS_DEMO As Int32 = 788
    
    Const STATUS_DESCRIPTION_ACTIVE = "Active"
    Const STATUS_DESCRIPTION_SUSPENDED = "Suspended"
    
    Const PC_EQUIPPED_YES = "Y"
    Const PC_EQUIPPED_NO = "N"
    
    Const NETWORK_AGENT_YES = "Y"
    Const NETWORK_AGENT_NO = "N"
   
    Const PENDING_DELETED_FLAG_NO = "N"
    
    Const ITEMS_PER_REQUEST As Integer = 30
    
    Protected Sub Page_Load(ByVal Source As Object, ByVal E As EventArgs)
        If Not IsNumeric(Session("CustomerKey")) Then
            Server.Transfer("session_expired.aspx")
        End If
        If Not IsPostBack Then
            Call SetTitle()
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
        Page.Header.Title = sTitle & "WU Agent Addresses"
    End Sub

    Protected Sub rgWURSAgents_NeedDataSource(ByVal source As Object, ByVal e As Telerik.Web.UI.GridNeedDataSourceEventArgs) Handles rgWURSAgents.NeedDataSource
        Dim sSearch As String = rcbAgents.SelectedValue
        Dim IListOfParams As New List(Of SqlParameter)
        Dim paramSearch As New SqlParameter("@Search", SqlDbType.VarChar)
        paramSearch.Value = sSearch
        IListOfParams.Add(paramSearch)
        rgWURSAgents.DataSource = ExecuteStoredProcedureToDataTable("WU_GetAllFrom_WU_Agents", IListOfParams)
    End Sub

    Protected Sub rgWURSAgents_ItemDataBound(ByVal source As Object, ByVal e As Telerik.Web.UI.GridItemEventArgs) Handles rgWURSAgents.ItemDataBound
        If TypeOf e.Item Is GridEditableItem And e.Item.IsInEditMode Then
            If TypeOf e.Item Is GridEditFormInsertItem Then
                Dim tdTermIDInsert As HtmlTableCell = e.Item.FindControl("tdTermIDInsert")
                tdTermIDInsert.Visible = True
            Else
                Dim tdTermIDUpdate As HtmlTableCell = e.Item.FindControl("tdTermIDUpdate")
                tdTermIDUpdate.Visible = True
            End If
        ElseIf TypeOf e.Item Is GridDataItem Then
        End If
    End Sub
    
    Protected Sub rcbAgents_SelectedIndexChanged(ByVal o As Object, ByVal e As Telerik.Web.UI.RadComboBoxSelectedIndexChangedEventArgs)
        rgWURSAgents.Rebind()
    End Sub
    
    Protected Sub rcbAgents_ItemsRequested(ByVal o As Object, ByVal e As Telerik.Web.UI.RadComboBoxItemsRequestedEventArgs)
        Dim s As String = e.Text
        Dim IListOfParams As New List(Of SqlParameter)
        Dim paramSearch As New SqlParameter("@Search", SqlDbType.VarChar)
        paramSearch.Value = e.Text
        IListOfParams.Add(paramSearch)
        Dim data As DataTable = ExecuteStoredProcedureToDataTable("WU_GetAllFrom_WU_Agents", IListOfParams)
        'Dim sThumbnailImage As String = String.Empty
        Dim itemOffset As Integer = e.NumberOfItems
        Dim endOffset As Integer = Math.Min(itemOffset + ITEMS_PER_REQUEST, data.Rows.Count)
        e.EndOfItems = endOffset = data.Rows.Count
        For i As Int32 = itemOffset To endOffset - 1
            Dim rcb As New RadComboBoxItem
            rcb.Text = data.Rows(i)("TermID").ToString() + " " + data.Rows(i)("AgentName").ToString()
            rcb.Value = data.Rows(i)("TermID").ToString()
            rcbAgents.Items.Add(rcb)
            Dim lblAgent As Label = rcb.FindControl("lblAgent")
            lblAgent.Text = data.Rows(i)("TermID").ToString() + " " + data.Rows(i)("AgentName").ToString()
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

    Protected Sub rgWURSAgents_ItemCommand(ByVal source As Object, ByVal e As Telerik.Web.UI.GridCommandEventArgs) Handles rgWURSAgents.ItemCommand
        If e.CommandName = "PerformInsert" Then
            Insert(e)
        ElseIf e.CommandName = "Update" Then
            Update(e)
        End If
    End Sub

    Protected Sub Insert(ByVal e As Telerik.Web.UI.GridCommandEventArgs)
        If (TypeOf e.Item Is Telerik.Web.UI.GridEditFormInsertItem) AndAlso e.Item.IsInEditMode Then
            Dim txtTermID As TextBox = e.Item.FindControl("txtTermID")
            Dim txtAgentName As TextBox = e.Item.FindControl("txtAgentName")
            Dim txtWUAccountNumber As TextBox = e.Item.FindControl("txtWUAccountNumber")
            Dim txtCompanyName As TextBox = e.Item.FindControl("txtCompanyName")
            Dim txtAddress1 As TextBox = e.Item.FindControl("txtAddress1")
            Dim txtAddress2 As TextBox = e.Item.FindControl("txtAddress2")
            Dim txtAddress3 As TextBox = e.Item.FindControl("txtAddress3")
            Dim txtCity As TextBox = e.Item.FindControl("txtCity")
            Dim txtState As TextBox = e.Item.FindControl("txtState")
            Dim txtPostCode As TextBox = e.Item.FindControl("txtPostCode")
            Dim txtContact As TextBox = e.Item.FindControl("txtContact")
            Dim txtPhoneNumber As TextBox = e.Item.FindControl("txtPhoneNumber")
            
            Dim chkStatusDesc As CheckBox = e.Item.FindControl("chkStatusDesc")
            Dim chkPCEquipped As CheckBox = e.Item.FindControl("chkPCEquipped")
            Dim chkNetworkAgent As CheckBox = e.Item.FindControl("chkNetworkAgent")
            Dim chkPendingDeletionFlag As CheckBox = e.Item.FindControl("chkPendingDeletionFlag")
            
            Dim lblMessage As Label = e.Item.FindControl("lblMessage")
            
            If Not IsTermIDExists(txtTermID.Text.Trim.ToUpper) Then
                Dim IListParameters As New List(Of SqlParameter)

                Dim paramTermID As New SqlParameter("@TermID", SqlDbType.VarChar, 4)
                paramTermID.Value = txtTermID.Text.Trim.ToUpper
                IListParameters.Add(paramTermID)

                Dim paramAgentName As New SqlParameter("@AgentName", SqlDbType.VarChar)
                paramAgentName.Value = txtAgentName.Text.Trim
                IListParameters.Add(paramAgentName)
                
                Dim paramWUAccountNumber As New SqlParameter("@WUAccountNumber", SqlDbType.VarChar)
                paramWUAccountNumber.Value = txtWUAccountNumber.Text.Trim
                IListParameters.Add(paramWUAccountNumber)

                Dim paramCompanyName As New SqlParameter("@CompanyName", SqlDbType.VarChar)
                paramCompanyName.Value = txtCompanyName.Text.Trim
                IListParameters.Add(paramCompanyName)

                Dim paramAddress1 As New SqlParameter("@Address1", SqlDbType.VarChar)
                paramAddress1.Value = txtAddress1.Text.Trim
                IListParameters.Add(paramAddress1)
            
                Dim paramAddress2 As New SqlParameter("@Address2", SqlDbType.VarChar)
                paramAddress2.Value = txtAddress2.Text.Trim
                IListParameters.Add(paramAddress2)

                Dim paramAddress3 As New SqlParameter("@Address3", SqlDbType.VarChar)
                paramAddress3.Value = txtAddress3.Text.Trim
                IListParameters.Add(paramAddress3)
            
                Dim paramCity As New SqlParameter("@City", SqlDbType.VarChar)
                paramCity.Value = txtCity.Text.Trim
                IListParameters.Add(paramCity)

                Dim paramState As New SqlParameter("@State", SqlDbType.VarChar)
                paramState.Value = txtState.Text.Trim
                IListParameters.Add(paramState)

                Dim paramPostCode As New SqlParameter("@PostCode", SqlDbType.VarChar)
                paramPostCode.Value = txtPostCode.Text.Trim
                IListParameters.Add(paramPostCode)
            
                Dim paramContact As New SqlParameter("@Contact", SqlDbType.VarChar)
                paramContact.Value = txtContact.Text.Trim
                IListParameters.Add(paramContact)

                Dim paramPhoneNumber As New SqlParameter("@PhoneNumber", SqlDbType.VarChar)
                paramPhoneNumber.Value = txtPhoneNumber.Text.Trim
                IListParameters.Add(paramPhoneNumber)
            
                Dim paramStatusDesc As New SqlParameter("@StatusDesc", SqlDbType.VarChar)
                paramStatusDesc.Value = IIf(chkStatusDesc.Checked, "Active", "Suspended")
                IListParameters.Add(paramStatusDesc)
            
                Dim paramPCEquipped As New SqlParameter("@PCEquipped", SqlDbType.VarChar)
                paramPCEquipped.Value = IIf(chkPCEquipped.Checked, "Y", "N")
                IListParameters.Add(paramPCEquipped)

                Dim paramNetworkAgent As New SqlParameter("@NetworkAgent", SqlDbType.VarChar)
                paramNetworkAgent.Value = IIf(chkNetworkAgent.Checked, "Y", "N")
                IListParameters.Add(paramNetworkAgent)

                Dim paramPendingDeletionFlag As New SqlParameter("@PendingDeletionFlag", SqlDbType.VarChar)
                paramPendingDeletionFlag.Value = "N"
                IListParameters.Add(paramPendingDeletionFlag)
                
                Dim paramIsDeleted As New SqlParameter("@IsDeleted", SqlDbType.Bit)
                paramIsDeleted.Value = 0
                IListParameters.Add(paramIsDeleted)
                
                Dim paramLastChangedOn As New SqlParameter("@LastChangedOn", SqlDbType.SmallDateTime)
                paramLastChangedOn.Value = DateTime.Now
                IListParameters.Add(paramLastChangedOn)
                
                Dim paramLastChangedBy As New SqlParameter("@LastChangedBy", SqlDbType.Int)
                paramLastChangedBy.Value = Convert.ToInt32(Session("UserKey"))
                IListParameters.Add(paramLastChangedBy)

                ExecuteStoredProcedureToDataTable("WU_InsertUpdate_WU_Agents", IListParameters)
            Else
                e.Canceled = True
                lblMessage.Text = "Terminal ID already exists."
            End If
        End If
    End Sub

    Protected Sub Update(ByVal e As Telerik.Web.UI.GridCommandEventArgs)
        If TypeOf e.Item Is GridEditableItem AndAlso e.Item.IsInEditMode Then
            Dim lblTermID As Label = e.Item.FindControl("lblTermID")
            Dim txtAgentName As TextBox = e.Item.FindControl("txtAgentName")
            Dim txtWUAccountNumber As TextBox = e.Item.FindControl("txtWUAccountNumber")
            Dim txtCompanyName As TextBox = e.Item.FindControl("txtCompanyName")
            Dim txtAddress1 As TextBox = e.Item.FindControl("txtAddress1")
            Dim txtAddress2 As TextBox = e.Item.FindControl("txtAddress2")
            Dim txtAddress3 As TextBox = e.Item.FindControl("txtAddress3")
            Dim txtCity As TextBox = e.Item.FindControl("txtCity")
            Dim txtState As TextBox = e.Item.FindControl("txtState")
            Dim txtPostCode As TextBox = e.Item.FindControl("txtPostCode")
            Dim txtContact As TextBox = e.Item.FindControl("txtContact")
            Dim txtPhoneNumber As TextBox = e.Item.FindControl("txtPhoneNumber")
            
            Dim chkStatusDesc As CheckBox = e.Item.FindControl("chkStatusDesc")
            Dim chkPCEquipped As CheckBox = e.Item.FindControl("chkPCEquipped")
            Dim chkNetworkAgent As CheckBox = e.Item.FindControl("chkNetworkAgent")
            Dim chkPendingDeletionFlag As CheckBox = e.Item.FindControl("chkPendingDeletionFlag")

            Dim IListParameters As New List(Of SqlParameter)

            Dim paramTermID As New SqlParameter("@TermID", SqlDbType.VarChar, 4)
            paramTermID.Value = lblTermID.Text
            IListParameters.Add(paramTermID)

            Dim paramAgentName As New SqlParameter("@AgentName", SqlDbType.VarChar)
            paramAgentName.Value = txtAgentName.Text.Trim
            IListParameters.Add(paramAgentName)
            
            Dim paramWUAccountNumber As New SqlParameter("@WUAccountNumber", SqlDbType.VarChar)
            paramWUAccountNumber.Value = txtWUAccountNumber.Text.Trim
            IListParameters.Add(paramWUAccountNumber)

            Dim paramCompanyName As New SqlParameter("@CompanyName", SqlDbType.VarChar)
            paramCompanyName.Value = txtCompanyName.Text.Trim
            IListParameters.Add(paramCompanyName)

            Dim paramAddress1 As New SqlParameter("@Address1", SqlDbType.VarChar)
            paramAddress1.Value = txtAddress1.Text.Trim
            IListParameters.Add(paramAddress1)
            
            Dim paramAddress2 As New SqlParameter("@Address2", SqlDbType.VarChar)
            paramAddress2.Value = txtAddress2.Text.Trim
            IListParameters.Add(paramAddress2)

            Dim paramAddress3 As New SqlParameter("@Address3", SqlDbType.VarChar)
            paramAddress3.Value = txtAddress3.Text.Trim
            IListParameters.Add(paramAddress3)
            
            Dim paramCity As New SqlParameter("@City", SqlDbType.VarChar)
            paramCity.Value = txtCity.Text.Trim
            IListParameters.Add(paramCity)

            Dim paramState As New SqlParameter("@State", SqlDbType.VarChar)
            paramState.Value = txtState.Text.Trim
            IListParameters.Add(paramState)

            Dim paramPostCode As New SqlParameter("@PostCode", SqlDbType.VarChar)
            paramPostCode.Value = txtPostCode.Text.Trim
            IListParameters.Add(paramPostCode)
            
            Dim paramContact As New SqlParameter("@Contact", SqlDbType.VarChar)
            paramContact.Value = txtContact.Text.Trim
            IListParameters.Add(paramContact)

            Dim paramPhoneNumber As New SqlParameter("@PhoneNumber", SqlDbType.VarChar)
            paramPhoneNumber.Value = txtPhoneNumber.Text.Trim
            IListParameters.Add(paramPhoneNumber)
            
            Dim paramStatusDesc As New SqlParameter("@StatusDesc", SqlDbType.VarChar)
            paramStatusDesc.Value = IIf(chkStatusDesc.Checked, "Active", "Suspended")
            IListParameters.Add(paramStatusDesc)
            
            Dim paramPCEquipped As New SqlParameter("@PCEquipped", SqlDbType.VarChar)
            paramPCEquipped.Value = IIf(chkPCEquipped.Checked, "Y", "N")
            IListParameters.Add(paramPCEquipped)

            Dim paramNetworkAgent As New SqlParameter("@NetworkAgent", SqlDbType.VarChar)
            paramNetworkAgent.Value = IIf(chkNetworkAgent.Checked, "Y", "N")
            IListParameters.Add(paramNetworkAgent)

            Dim paramPendingDeletionFlag As New SqlParameter("@PendingDeletionFlag", SqlDbType.VarChar)
            paramPendingDeletionFlag.Value = "N"
            IListParameters.Add(paramPendingDeletionFlag)
            
            Dim paramIsDeleted As New SqlParameter("@IsDeleted", SqlDbType.Bit)
            paramIsDeleted.Value = 0
            IListParameters.Add(paramIsDeleted)
            
            Dim paramLastChangedOn As New SqlParameter("@LastChangedOn", SqlDbType.SmallDateTime)
            paramLastChangedOn.Value = DateTime.Now
            IListParameters.Add(paramLastChangedOn)
                
            Dim paramLastChangedBy As New SqlParameter("@LastChangedBy", SqlDbType.Int)
            paramLastChangedBy.Value = Convert.ToInt32(Session("UserKey"))
            IListParameters.Add(paramLastChangedBy)

            ExecuteStoredProcedureToDataTable("WU_InsertUpdate_WU_Agents", IListParameters)
        End If
    End Sub

    Protected Sub lnkDelete_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim lnkDelete As LinkButton = sender
        Dim sTermID As String = lnkDelete.CommandArgument.ToUpper
        Dim sSQL As String = "Update ClientData_WU_Agents set IsDeleted = 1 where TermID = '" & sTermID & "'"
        ExecuteQueryToDataTable(sSQL)
        rgWURSAgents.Rebind()

    End Sub
    
    Protected Function IsTermIDExists(ByVal sTermID As String) As Boolean
        IsTermIDExists = False
        Dim sSQL As String = "select TermID from ClientData_WU_Agents where TermID = '" & sTermID & "'"
        Dim oDataTable As DataTable = ExecuteQueryToDataTable(sSQL)
        If Not oDataTable Is Nothing AndAlso oDataTable.Rows.Count <> 0 Then
            IsTermIDExists = True
        End If
    End Function
    
    Protected Shared Function ExecuteQueryToDataTable(ByVal sQuery As String) As DataTable
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
    
    Public Shared Function ExecuteStoredProcedureToDataTable(ByVal sp_name As String, Optional ByVal IListPrams As List(Of SqlParameter) = Nothing) As DataTable
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
    
    Protected Sub btnMissingAddressesReport_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim dtAccountsMissingAddress As DataTable = ExecuteQueryToDataTable("SELECT UserId, FirstName + ' ' + LastName 'Name', CASE WHEN CustomerKey = 579 THEN 'WURS' WHEN CustomerKey = 686 THEN 'WUIRE' ELSE 'undefined' END 'Account' FROM UserProfile up WHERE CustomerKey IN (579, 686) AND Status = 'Active' AND Type = 'User' AND LEN(UserID) = 4 AND NOT EmailAddr LIKE '%@westernunion%' AND NOT UserId IN ('GeodisWUIRE') AND NOT UserID IN (SELECT Termid FROM ClientData_WU_Agents) ORDER BY CustomerKey, UserID")
        Dim sbText As New StringBuilder
        Call AddHTMLPreamble(sbText, "Western Union Missing Addresses Report")
        sbText.Append(Bold("WESTERN UNION MISSING ADDRESSES REPORT"))
        Call NewLine(sbText)
        sbText.Append("report generated " & DateTime.Now)
        Call NewLine(sbText)
        Call NewLine(sbText)
        Dim nMissingAddressCount As Int32 = dtAccountsMissingAddress.Rows.Count
        If nMissingAddressCount = 0 Then
            sbText.Append("All Western Union user accounts have address information")
        Else
            sbText.Append("The " & nMissingAddressCount.ToString & " Western Union user accounts below have no address information.")
            Call NewLine(sbText)
            Call NewLine(sbText)
            sbText.Append("<hr />")
            Call NewLine(sbText)
            Call NewLine(sbText)
            sbText.Append("<table style='font-size: small'>")
            For Each dr As DataRow In dtAccountsMissingAddress.Rows
                sbText.Append("<tr>")
                sbText.Append("<td>")
                sbText.Append("<b>")
                sbText.Append(dr("UserID"))
                sbText.Append("</b>")
                sbText.Append("</td>")
                sbText.Append("<td>")
                sbText.Append(dr("Name"))
                sbText.Append("</td>")
                sbText.Append("<td>")
                sbText.Append(dr("Account"))
                sbText.Append("</td>")
                sbText.Append("</tr>")
            Next
            sbText.Append("</table>")
            sbText.Append("<hr />")
        End If
        Call NewLine(sbText)
        Call NewLine(sbText)
        sbText.Append("[end]")
        Call AddHTMLPostamble(sbText)
        Call ExportData(sbText.ToString, "WesternUnionMissingAddressesReport")
    End Sub
    
    Protected Function Bold(ByVal sString As String) As String
        Bold = "<b>" & sString & "</b>"
    End Function

    Protected Sub NewLine(ByRef sbText As StringBuilder)
        sbText.Append("<br />" & Environment.NewLine)
    End Sub

    Protected Sub AddHTMLPreamble(ByRef sbText As StringBuilder, ByVal sTitle As String)
        sbText.Append("<html><head><title>")
        sbText.Append(sTitle)
        sbText.Append("</title><style>")
        sbText.Append("body { font-family: Verdana; font-size : xx-small }")
        sbText.Append("</style></head><body>")
    End Sub

    Protected Sub AddHTMLPostamble(ByRef sbText As StringBuilder)
        sbText.Append("</body></html>")
    End Sub

    Private Sub ExportData(ByVal sData As String, ByVal sFilename As String)
        Response.Clear()
        Response.AddHeader("Content-Disposition", "attachment;filename=" & sFilename & ".htm")
        Response.ContentType = "text/html"

        Dim eEncoding As Encoding = Encoding.GetEncoding("Windows-1252")
        Dim eUnicode As Encoding = Encoding.Unicode
        Dim byUnicodeBytes As Byte() = eUnicode.GetBytes(sData)
        Dim byEncodedBytes As Byte() = Encoding.Convert(eUnicode, eEncoding, byUnicodeBytes)
        Response.BinaryWrite(byEncodedBytes)

        Response.Flush()
        Response.End()
    End Sub

</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <main:Header ID="ctlHeader" runat="server" />
    <asp:PlaceHolder ID="PlaceHolderForScriptManager" runat="server"/>
    <table style="width: 100%" cellpadding="0" cellspacing="0">
        <tr>
            <td style="width: 70%; white-space: nowrap">
                &nbsp;
                <asp:Label ID="Label1" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Search Term ID / Agent Name:"/>
                <telerik:RadComboBox ID="rcbAgents" runat="server" Width="300px" Height="500" Font-Names="Arial" EmptyMessage="- select Term ID / Agent Name" Font-Size="X-Small" Font-Bold="true" OnSelectedIndexChanged="rcbAgents_SelectedIndexChanged" AutoPostBack="True" HighlightTemplatedItems="true" CausesValidation="False" EnableLoadOnDemand="True" OnItemsRequested="rcbAgents_ItemsRequested" EnableVirtualScrolling="True" ShowMoreResultsBox="True" Filter="Contains">
                    <ItemTemplate>
                        <table>
                            <tr>
                                <asp:Label ID="lblAgent" runat="server" />
                            </tr>
                        </table>
                    </ItemTemplate>
                </telerik:RadComboBox>
            </td>
            <td style="width: 30%; white-space: nowrap" align="right">
                <asp:Button ID="btnMissingAddressesReport" runat="server" Text="Missing addresses report" onclick="btnMissingAddressesReport_Click" CausesValidation="False" />
            &nbsp;
            </td>
        </tr>
    </table>
    <div>
        &nbsp;
    </div>
    <div>
        <telerik:RadGrid ID="rgWURSAgents" runat="server" CellPadding="2" Font-Names="Verdana" AllowPaging="true" PageSize="10" AllowSorting="true" Font-Size="XX-Small" Width="100%" AutoGenerateColumns="False">
            <MasterTableView CommandItemDisplay="Top">
                <Columns>
                    <telerik:GridEditCommandColumn ButtonType="ImageButton" UniqueName="EditCommandColumn" ItemStyle-Width="30px"/>
                    <telerik:GridBoundColumn DataField="PendingDeletionFlag" HeaderText="Delete" Visible="false"/>
                    <telerik:GridBoundColumn DataField="TermID" HeaderText="Term ID" Visible="true"/>
                    <telerik:GridBoundColumn DataField="StatusDesc" HeaderText="Status"/>
                    <telerik:GridBoundColumn DataField="PCEquipped" HeaderText="PC Equipped"/>
                    <telerik:GridBoundColumn DataField="NetworkAgent" HeaderText="Network Agent"/>
                    <telerik:GridBoundColumn DataField="AgentName" HeaderText="Agent Name"/>
                    <telerik:GridBoundColumn DataField="Address1" HeaderText="Address1"/>
                    <telerik:GridBoundColumn DataField="Address2" HeaderText="Address2"/>
                    <telerik:GridBoundColumn DataField="Address3" HeaderText="Address3"/>
                    <telerik:GridBoundColumn DataField="City" HeaderText="City"/>
                    <telerik:GridBoundColumn DataField="State" HeaderText="State"/>
                    <telerik:GridBoundColumn DataField="PostCode" HeaderText="PostCode"/>
                    <telerik:GridBoundColumn DataField="Contact" HeaderText="Contact"/>
                    <telerik:GridBoundColumn DataField="PhoneNumber" HeaderText="PhoneNumber"/>
                    <telerik:GridTemplateColumn>
                        <ItemTemplate>
                            <asp:LinkButton ID="lnkDelete" runat="server" OnClick="lnkDelete_Click" CommandArgument='<%# Bind("TermID") %>' Text="Delete" ToolTip="Delete" CommandName="delete" OnClientClick="return confirm('Are you sure you want to delete this record?');">                                
                            </asp:LinkButton>
                        </ItemTemplate>
                    </telerik:GridTemplateColumn>
                </Columns>
                <EditFormSettings EditFormType="Template" InsertCaption="Add New Agent">
                    <FormTemplate>
                        <table width="100%">
                            <tr>
                                <td width="20%">
                                    <label style="color: Red">
                                        Term ID</label>
                                </td>
                                <td id="tdTermIDInsert" visible="false" runat="server">
                                    <asp:TextBox ID="txtTermID" Text='<%# Bind("TermID") %>' MaxLength="4" runat="server"/>
                                    <asp:Label ID="lblMessage" Text="" ForeColor="Red" runat="server"/>
                                    <asp:RequiredFieldValidator ID="rfvTermID" ControlToValidate="txtTermID" Display="Dynamic" ErrorMessage="Please enter terminal ID" ForeColor="Red" runat="server"/>
                                </td>
                                <td id="tdTermIDUpdate" visible="false" runat="server">
                                    <asp:Label ID="lblTermID" Text='<%# Bind("TermID") %>' runat="server"/>
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    <label style="color: Black">
                                        WU Account Number</label>
                                </td>
                                <td>
                                    <asp:TextBox ID="txtWUAccountNumber" Text='<%# Bind("WUAccountNumber") %>' MaxLength="50" runat="server"/>
                                    <asp:RequiredFieldValidator ID="RequiredFieldValidator8" ControlToValidate="txtWUAccountNumber" Display="Dynamic" ErrorMessage="Please enter WU account number" ForeColor="Red" runat="server" Enabled="false"/>
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    <label style="color: Black">
                                        Company Name</label>
                                </td>
                                <td>
                                    <asp:TextBox ID="txtCompanyName" Text='<%# Bind("CompanyName") %>' MaxLength="50" runat="server"/>
                                    <asp:RequiredFieldValidator ID="RequiredFieldValidator9" ControlToValidate="txtCompanyName" Display="Dynamic" ErrorMessage="Please enter company name" ForeColor="Red" runat="server" Enabled="false"/>
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    <label>
                                        Agent Name</label>
                                </td>
                                <td>
                                    <asp:TextBox ID="txtAgentName" Text='<%# Bind("AgentName") %>' MaxLength="50" runat="server"/>
                                    <asp:RequiredFieldValidator ID="rfvAgentName" ControlToValidate="txtAgentName" Display="Dynamic" ErrorMessage="Please enter agent name" ForeColor="Red" runat="server"/>
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    <label>
                                        Address 1</label>
                                </td>
                                <td>
                                    <asp:TextBox ID="txtAddress1" Text='<%# Bind("Address1") %>' MaxLength="50" runat="server"/>
                                    <asp:RequiredFieldValidator ID="rfvAddress" ControlToValidate="txtAddress1" Display="Dynamic" ErrorMessage="Please enter address 1" ForeColor="Red" runat="server"/>
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    <label style="color: Black">
                                        Address 2</label>
                                </td>
                                <td>
                                    <asp:TextBox ID="txtAddress2" Text='<%# Bind("Address2") %>' MaxLength="50" runat="server"/>
                                    <asp:RequiredFieldValidator ID="RequiredFieldValidator1" ControlToValidate="txtAddress2" ErrorMessage="Please enter address 2" Display="Dynamic" ForeColor="Red" runat="server" Enabled="false"/>
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    <label style="color: Black">
                                        Address 3</label>
                                </td>
                                <td>
                                    <asp:TextBox ID="txtAddress3" Text='<%# Bind("Address3") %>' MaxLength="50" runat="server"/>
                                    <asp:RequiredFieldValidator ID="RequiredFieldValidator2" ControlToValidate="txtAddress3" Display="Dynamic" ErrorMessage="Please enter address 3" ForeColor="Red" runat="server" Enabled="false"/>
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    <label>
                                        Town/City</label>
                                </td>
                                <td>
                                    <asp:TextBox ID="txtCity" Text='<%# Bind("City") %>' MaxLength="50" runat="server"/>
                                    <asp:RequiredFieldValidator ID="RequiredFieldValidator3" ControlToValidate="txtCity" Display="Dynamic" ErrorMessage="Please enter city" ForeColor="Red" runat="server"/>
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    <label style="color: Black">
                                        State</label>
                                </td>
                                <td>
                                    <asp:TextBox ID="txtState" Text='<%# Bind("State") %>' MaxLength="50" runat="server"/>
                                    <asp:RequiredFieldValidator ID="RequiredFieldValidator4" ControlToValidate="txtState" Display="Dynamic" ErrorMessage="Please enter state" ForeColor="Red" runat="server" Enabled="false"/>
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    <label>
                                        Post Code</label>
                                </td>
                                <td>
                                    <asp:TextBox ID="txtPostCode" Text='<%# Bind("PostCode") %>' MaxLength="50" runat="server"/>
                                    <asp:RequiredFieldValidator ID="RequiredFieldValidator5" ControlToValidate="txtPostCode" Display="Dynamic" ErrorMessage="Please enter post code" ForeColor="Red" runat="server"/>
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    <label style="color: Black">
                                        Contact</label>
                                </td>
                                <td>
                                    <asp:TextBox ID="txtContact" Text='<%# Bind("Contact") %>' MaxLength="50" runat="server"/>
                                    <asp:RequiredFieldValidator ID="RequiredFieldValidator6" ControlToValidate="txtContact" Display="Dynamic" ErrorMessage="Please enter contact name" ForeColor="Red" runat="server" Enabled="false"/>
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    <label style="color: Black">
                                        Phone Number</label>
                                </td>
                                <td>
                                    <asp:TextBox ID="txtPhoneNumber" Text='<%# Bind("PhoneNumber") %>' MaxLength="50" runat="server"/>
                                    <asp:RequiredFieldValidator ID="RequiredFieldValidator7" ControlToValidate="txtAddress1" Display="Dynamic" ErrorMessage="Please enter telephone number" ForeColor="Red" runat="server" Enabled="false"/>
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    Status Desc.
                                </td>
                                <td>
                                    <asp:CheckBox ID="chkStatusDesc" Checked='<%# IIf( CheckNull(Eval("StatusDesc")) = "Active", "true" , "false") %>' runat="server" />
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    PC Equipped
                                </td>
                                <td>
                                    <asp:CheckBox ID="chkPCEquipped" Checked='<%# IIf( CheckNull(Eval("PCEquipped")) = "Y", "true" , "false") %>' runat="server" />
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    Network Agent
                                </td>
                                <td>
                                    <asp:CheckBox ID="chkNetworkAgent" Checked='<%# IIf( CheckNull(Eval("NetworkAgent")) = "Y", "true" , "false") %>' runat="server" />
                                </td>
                            </tr>
                            <tr>
                                <td>
                                </td>
                                <td>
                                    <asp:LinkButton ID="lnkbtnUpdate" runat="server" CausesValidation="True" Text='<%# IIf( DataBinder.Eval(Container, "OwnerTableView.IsItemInserted"), "Insert", "Update") %>' CommandName='<%# IIf( DataBinder.Eval(Container, "OwnerTableView.IsItemInserted"), "PerformInsert", "Update") %>'></asp:LinkButton>
                                    <asp:LinkButton ID="lnkbtnCancel" runat="server" Text="Cancel" CausesValidation="False" CommandName="Cancel"></asp:LinkButton>
                                    <asp:HiddenField ID="hidTermID" Value='<%# Bind("TermID") %>' runat="server" />
                                </td>
                            </tr>
                        </table>
                    </FormTemplate>
                </EditFormSettings>
            </MasterTableView>
        </telerik:RadGrid>
    </div>
    </form>
</body>
</html>