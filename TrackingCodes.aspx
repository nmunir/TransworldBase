<%@ Page Language="VB" ValidateRequest="false" %>
<%@ Import Namespace="System.Collections" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="Telerik.Web.UI" %>
<%@ Register Assembly="Telerik.Web.UI" Namespace="Telerik.Web.UI" TagPrefix="telerik" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">
    
    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    Private gridMessage As String = Nothing

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsNumeric(Session("UserKey")) Then
            Server.Transfer("session_expired.aspx")
        End If
        If Not Page.IsPostBack Then
            LoadData()
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
    
    Private Function LoadData() As DataTable
        Dim sSQL As String = "SELECT ID, Code, Description, ClientVisibility, Analysis, SurchargeTrigger from TrackingCodesMETACS order by ID desc"
        Try
            LoadData = ExecuteQueryToDataTable(sSQL)
        Catch
            LoadData = Nothing
        End Try
        rgTrackingCodes.DataSource = LoadData
    End Function
    
    Public Function ExecuteStoredProcedure_GetIdentity(sStoredProcedure As String, parameter As List(Of SqlParameter), Optional op As SqlParameter = Nothing) As Int32
        'Dim gsConn As String = "Data Source=KAZIM;Initial Catalog=Logistics;User ID=sa;Password=rugby22"
        Dim oConn As New SqlConnection(gsConn)
        Try
            oConn.Open()
            Dim result As Int32 = -1
            Dim dt As New DataTable()
            Dim oCmd As New SqlCommand()
            oCmd.Connection = oConn
            oCmd.CommandText = sStoredProcedure
            oCmd.CommandType = CommandType.StoredProcedure
            If parameter IsNot Nothing AndAlso parameter.Count <> 0 Then
                oCmd.Parameters.AddRange(parameter.ToArray())
            End If

            Dim reader As SqlDataReader = oCmd.ExecuteReader()
            If op IsNot Nothing Then
                If Not IsDBNull(oCmd.Parameters(op.ParameterName).Value) Then
                    result = CInt(oCmd.Parameters(op.ParameterName).Value)

                End If
            Else
                If Not IsDBNull(oCmd.Parameters("@Id").Value) Then
                    result = CInt(oCmd.Parameters("@Id").Value)
                End If
            End If
            Return result
        Catch ex As Exception
            ex.Message.ToString()
            Return -1
        Finally
            oConn.Close()
        End Try
    End Function

    Protected Function ExecuteQueryToDataTable(ByVal sQuery As String) As DataTable
        'Dim gsConn As String = "Data Source=KAZIM;Initial Catalog=Logistics;User ID=sa;Password=rugby22"
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

    Protected Sub rgTrackingCodes_ItemDataBound(ByVal source As Object, ByVal e As Telerik.Web.UI.GridItemEventArgs) Handles rgTrackingCodes.ItemDataBound
        If TypeOf (e.Item) Is GridDataItem Then
        ElseIf TypeOf (e.Item) Is GridEditFormItem And e.Item.IsInEditMode Then
            Dim item As GridEditFormItem = e.Item
            Dim hidClientVisibility As HiddenField = item.FindControl("hidClientVisibility")
            Dim hidAnalysis As HiddenField = item.FindControl("hidAnalysis")
            Dim chkClientVisibility As CheckBox = item.FindControl("chkClientVisibility")
            Dim chkAnalysis As CheckBox = item.FindControl("chkAnalysis")
            Dim cClientVisibility As Char
            Dim cAnalysis As Char
            cClientVisibility = hidClientVisibility.Value.ToString
            cAnalysis = hidAnalysis.Value.ToString
            If cClientVisibility = "Y" Then
                chkClientVisibility.Checked = True
            Else
                chkClientVisibility.Checked = False
            End If
            If cAnalysis = "Y" Then
                chkAnalysis.Checked = True
            Else
                chkAnalysis.Checked = False
            End If
        End If
    End Sub

    Protected Sub rgTrackingCodes_ItemCommand(ByVal source As Object, ByVal e As Telerik.Web.UI.GridCommandEventArgs) Handles rgTrackingCodes.ItemCommand
        If e.CommandName = RadGrid.InitInsertCommandName Then
        ElseIf e.CommandName = RadGrid.UpdateCommandName Then
            Update(e)
        ElseIf e.CommandName = RadGrid.PerformInsertCommandName Then
            Insert(e)
        ElseIf e.CommandName = RadGrid.DeleteCommandName Then
            Delete(e)
        End If
    End Sub
    
    Private Sub Delete(ByVal e As GridCommandEventArgs)
        Dim item As GridDataItem = e.Item
        Dim nID As Integer = Convert.ToInt32(item("ID").Text)
        Dim sql As String = "Delete from TrackingCodesMetacs where ID = " & nID
        ExecuteQueryToDataTable(sql)
    End Sub
    
    Private Sub Insert(ByVal e As GridCommandEventArgs)
        Dim cClientVisibility As Char = "N"
        Dim cAnalysis As Char = "N"
        Dim txtCode As TextBox = e.Item.FindControl("txtCode")
        Dim txtDescription As TextBox = e.Item.FindControl("txtDescription")
        Dim chkClientVisibility As CheckBox = e.Item.FindControl("chkClientVisibility")
        Dim chkAnalysis As CheckBox = e.Item.FindControl("chkAnalysis")
        Dim txtSurchargeTrigger As TextBox = e.Item.FindControl("txtSurchargeTrigger")
        If chkClientVisibility.Checked Then
            cClientVisibility = "Y"
        Else
            cClientVisibility = "N"
        End If
        If chkAnalysis.Checked Then
            cAnalysis = "Y"
        Else
            cAnalysis = "N"
        End If
        If Not IsCodeExists(txtCode.Text.Trim) Then
            Dim IParam As New List(Of SqlParameter)
            IParam.Add(New SqlParameter("@Id", -1))
            IParam.Add(New SqlParameter("@Code", txtCode.Text.Trim))
            IParam.Add(New SqlParameter("@Description", txtDescription.Text.Trim))
            IParam.Add(New SqlParameter("@ClientVisibility", cClientVisibility))
            IParam.Add(New SqlParameter("@Analysis", cAnalysis))
            IParam.Add(New SqlParameter("@SurchargeTrigger", txtSurchargeTrigger.Text.Trim))
            ExecuteStoredProcedure_GetIdentity("UpdateTrackingCodes", IParam)
        Else
            WebMsgBox.Show("Code already exists.")
        End If
        
        'Dim opParam As New SqlParameter()
        'opParam.ParameterName = "@Id"
        'opParam.Value = nID
        'opParam.SqlDbType = SqlDbType.Int
        'opParam.Direction = ParameterDirection.InputOutput
        'IParam.Add(opParam)
    End Sub

    Private Sub Update(ByVal e As Telerik.Web.UI.GridCommandEventArgs)
        Dim cClientVisibility As Char = "N"
        Dim cAnalysis As Char = "N"
        Dim hidID As HiddenField = e.Item.FindControl("hidID")
        Dim nID As Integer = Convert.ToInt32(hidID.Value)
        Dim txtCode As TextBox = e.Item.FindControl("txtCode")
        Dim txtDescription As TextBox = e.Item.FindControl("txtDescription")
        Dim chkClientVisibility As CheckBox = e.Item.FindControl("chkClientVisibility")
        Dim chkAnalysis As CheckBox = e.Item.FindControl("chkAnalysis")
        Dim txtSurchargeTrigger As TextBox = e.Item.FindControl("txtSurchargeTrigger")
        If chkClientVisibility.Checked Then
            cClientVisibility = "Y"
        Else
            cClientVisibility = "N"
        End If
        If chkAnalysis.Checked Then
            cAnalysis = "Y"
        Else
            cAnalysis = "N"
        End If
        Dim IParam As New List(Of SqlParameter)
        IParam.Add(New SqlParameter("@Id", nID))
        IParam.Add(New SqlParameter("@Code", txtCode.Text.Trim))
        IParam.Add(New SqlParameter("@Description", txtDescription.Text.Trim))
        IParam.Add(New SqlParameter("@ClientVisibility", cClientVisibility))
        IParam.Add(New SqlParameter("@Analysis", cAnalysis))
        IParam.Add(New SqlParameter("@SurchargeTrigger", txtSurchargeTrigger.Text.Trim))
        'Dim opParam As New SqlParameter()
        'opParam.ParameterName = "@Id"
        'opParam.Value = nID
        'opParam.SqlDbType = SqlDbType.Int
        'opParam.Direction = ParameterDirection.InputOutput
        'IParam.Add(opParam)
        ExecuteStoredProcedure_GetIdentity("UpdateTrackingCodes", IParam)
    End Sub
    
    Private Function IsCodeExists(ByVal sCode As String) As Boolean
        Dim bExits As Boolean = False
        Dim sql As String = "select Code from TrackingCodesMetacs where Code = '" & sCode & "'"
        Dim dt As DataTable = ExecuteQueryToDataTable(sql)
        If dt.Rows.Count = 0 Then
            IsCodeExists = False
        Else
            IsCodeExists = True
        End If
    End Function
 
    Protected Sub rgTrackingCodes_NeedDataSource(ByVal sender As Object, ByVal e As EventArgs) Handles rgTrackingCodes.NeedDataSource
        LoadData()
    End Sub
    
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <main:Header ID="ctlHeader" runat="server" />
    <asp:PlaceHolder ID="PlaceHolderForScriptManager" runat="server" />
    <telerik:RadGrid ID="rgTrackingCodes" GridLines="None" runat="server" PageSize="10"
        AllowSorting="true" AllowPaging="True" AutoGenerateColumns="False" Font-Names="Verdana" Font-Size="XX-Small">
        <PagerStyle Mode="NextPrevAndNumeric" />
        <MasterTableView Width="100%" ShowHeadersWhenNoRecords="true" CommandItemDisplay="Top"
            EditMode="EditForms" DataKeyNames="ID" HorizontalAlign="NotSet" AutoGenerateColumns="False">
            <Columns>
                <telerik:GridEditCommandColumn ButtonType="ImageButton" UniqueName="EditCommandColumn">
                </telerik:GridEditCommandColumn>
                <telerik:GridBoundColumn DataField="ID" HeaderText="ID" SortExpression="ID" UniqueName="ID">
                </telerik:GridBoundColumn>
                <telerik:GridBoundColumn DataField="Code" HeaderText="Code" SortExpression="Code"
                    UniqueName="Code">
                </telerik:GridBoundColumn>
                <telerik:GridBoundColumn DataField="Description" HeaderText="Description" SortExpression="Description"
                    UniqueName="Description">
                </telerik:GridBoundColumn>
                <telerik:GridBoundColumn DataField="ClientVisibility" HeaderText="Client Visibility"
                    SortExpression="ClientVisibility" UniqueName="ClientVisibility">
                </telerik:GridBoundColumn>
                <telerik:GridBoundColumn DataField="Analysis" HeaderText="Analysis" SortExpression="ClientVisibility"
                    UniqueName="Analysis">
                </telerik:GridBoundColumn>
                <telerik:GridBoundColumn DataField="SurchargeTrigger" HeaderText="Surcharge Trigger"
                    SortExpression="SurchargeTrigger" UniqueName="SurchargeTrigger">
                </telerik:GridBoundColumn>
                <telerik:GridButtonColumn ConfirmText="Delete this code?" ConfirmDialogType="RadWindow"
                    ConfirmTitle="Delete" ButtonType="ImageButton" CommandName="Delete" Text="Delete"
                    UniqueName="DeleteColumn">
                    <ItemStyle HorizontalAlign="Center" CssClass="MyImageButton" />
                </telerik:GridButtonColumn>
            </Columns>
            <EditFormSettings EditFormType="Template">
                <FormTemplate>
                    <asp:Panel ID="pnlValidation" runat="server">
                        <asp:ValidationSummary ID="vs" ValidationGroup="vg" runat="server" />
                    </asp:Panel>
                    <table width="100%">
                        <tr>
                            <td>
                                Code
                            </td>
                            <td>
                                <asp:TextBox ID="txtCode" MaxLength="50" Text='<%# Bind("Code") %>' ValidationGroup="vg"
                                    runat="server" Font-Names="Verdana" Font-Size="XX-Small"/>
                                <asp:RequiredFieldValidator ID="rfvCode" ControlToValidate="txtCode" ValidationGroup="vg"
                                    ErrorMessage="Please enter code" runat="server" Font-Names="Verdana" Font-Size="XX-Small"/>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                Description
                            </td>
                            <td>
                                <asp:TextBox ID="txtDescription" MaxLength="50" Text='<%# Bind("Description") %>'
                                    runat="server" Font-Names="Verdana" Font-Size="XX-Small"/>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                ClientVisibility
                            </td>
                            <td>
                                <asp:CheckBox ID="chkClientVisibility" runat="server" Font-Names="Verdana" Font-Size="XX-Small" />
                            </td>
                        </tr>
                        <tr>
                            <td>
                                Analysis
                            </td>
                            <td>
                                <asp:CheckBox ID="chkAnalysis" runat="server" Font-Names="Verdana" Font-Size="XX-Small" />
                            </td>
                        </tr>
                        <tr>
                            <td>
                                Surcharge Trigger
                            </td>
                            <td>
                                <asp:TextBox ID="txtSurchargeTrigger" MaxLength="50" Text='<%# Bind("SurchargeTrigger") %>'
                                    runat="server" Font-Names="Verdana" Font-Size="XX-Small"/>
                            </td>
                        </tr>
                        <tr>
                            <td align="right" colspan="2">
                                <asp:Button ID="btnUpdate" Text='<%# IIf((TypeOf(Container) is GridEditFormInsertItem), "Insert", "Update") %>'
                                    runat="server" CommandName='<%# IIf((TypeOf(Container) is GridEditFormInsertItem), "PerformInsert", "Update")%>'>
                                </asp:Button>&nbsp;
                                <asp:Button ID="btnCancel" Text="Cancel" runat="server" CausesValidation="False"
                                    CommandName="Cancel"/>
                                <asp:HiddenField ID="hidClientVisibility" Value='<%# Bind("ClientVisibility") %>'
                                    runat="server" />
                                <asp:HiddenField ID="hidAnalysis" Value='<%# Bind("Analysis") %>' runat="server" />
                                <asp:HiddenField ID="hidID" Value='<%# Bind("ID") %>' runat="server" />
                            </td>
                        </tr>
                    </table>
                </FormTemplate>
            </EditFormSettings>
        </MasterTableView>
    </telerik:RadGrid>
    </form>
</body>
</html>