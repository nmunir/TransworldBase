<%@ Page Language="VB" Theme="AIMSDefault" %>

<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="Telerik.Web.UI" %>
<%@ Register Assembly="Telerik.Web.UI" Namespace="Telerik.Web.UI" TagPrefix="telerik" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<script runat="server">

    ' TO DO
    ' AND CustomerKey IN (579, 686, 788, 798)
    
    Const ITEMS_PER_REQUEST As Integer = 30

    Const USER_PERMISSION_ACCOUNT_HANDLER As Integer = 1
    Const USER_PERMISSION_SITE_ADMINISTRATOR As Integer = 2
    Const USER_PERMISSION_DEPUTY_SITE_ADMINISTRATOR As Integer = 4
    Const USER_PERMISSION_SITE_EDITOR As Integer = 8
    Const USER_PERMISSION_DEPUTY_SITE_EDITOR As Integer = 16

    Const USER_PERMISSION_VIEW_STOCK As Integer = 1024
    Const USER_PERMISSION_CREATE_STOCK_BOOKING As Integer = 2048
    Const USER_PERMISSION_PRINT_ON_DEMAND_TAB As Integer = 4096
    Const USER_PERMISSION_ADVANCED_PERMISSIONS_TAB As Integer = 8192
    Const USER_PERMISSION_FILE_UPLOAD_TAB As Integer = 16384

    Const USER_PERMISSION_WU_INTERNAL_USER As Integer = &H8000
    Const USER_PERMISSION_WU_CREATE_DELETE_ACCOUNTS As Integer = &H10000
    Const USER_PERMISSION_WU_RESET_PASSWORDS As Integer = &H20000
    Const USER_PERMISSION_WU_ACCESS_REPORTS As Integer = &H40000
    Const USER_PERMISSION_WU_VIEW_STOCK As Integer = &H80000
    Const USER_PERMISSION_WU_ORDER_STOCK As Integer = &H100000
    Const USER_PERMISSION_WU_IS_TSE As Integer = &H200000
    Const USER_PERMISSION_WU_IS_PILOT_USER As Integer = &H400000
    Const USER_PERMISSION_WU_IS_SALES As Integer = &H800000

    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    Dim sSQL As String = String.Empty
    Dim gbNewMessages As Boolean
    
    Protected Function IListOfPermissions() As IList(Of Int32)
        
        IListOfPermissions = New List(Of Int32)
        IListOfPermissions.Add(USER_PERMISSION_ACCOUNT_HANDLER)
        IListOfPermissions.Add(USER_PERMISSION_SITE_ADMINISTRATOR)
        IListOfPermissions.Add(USER_PERMISSION_DEPUTY_SITE_ADMINISTRATOR)
        IListOfPermissions.Add(USER_PERMISSION_SITE_EDITOR)
        IListOfPermissions.Add(USER_PERMISSION_DEPUTY_SITE_EDITOR)
        'IListOfPermissions.Add(USER_PERMISSION_VIEW_STOCK)
        'IListOfPermissions.Add(USER_PERMISSION_CREATE_STOCK_BOOKING)
        'IListOfPermissions.Add(USER_PERMISSION_PRINT_ON_DEMAND_TAB)
        'IListOfPermissions.Add(USER_PERMISSION_ADVANCED_PERMISSIONS_TAB)
        'IListOfPermissions.Add(USER_PERMISSION_FILE_UPLOAD_TAB)
        
        IListOfPermissions.Add(USER_PERMISSION_WU_INTERNAL_USER)
        IListOfPermissions.Add(USER_PERMISSION_WU_CREATE_DELETE_ACCOUNTS)
        IListOfPermissions.Add(USER_PERMISSION_WU_RESET_PASSWORDS)
        IListOfPermissions.Add(USER_PERMISSION_WU_ACCESS_REPORTS)
        IListOfPermissions.Add(USER_PERMISSION_WU_VIEW_STOCK)
        IListOfPermissions.Add(USER_PERMISSION_WU_ORDER_STOCK)
        IListOfPermissions.Add(USER_PERMISSION_WU_IS_TSE)
        IListOfPermissions.Add(USER_PERMISSION_WU_IS_PILOT_USER)
        IListOfPermissions.Add(USER_PERMISSION_WU_IS_SALES)
        
    End Function
    
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsNumeric(Session("CustomerKey")) Then
            Server.Transfer("session_expired.aspx")
            Exit Sub
        End If
        If Not IsPostBack Then
            Call HideAllPanels()
            trFindAccount.Visible = True
            Call PopulateUserGridview()
            Call SetGridVisibility(True)
        End If
    End Sub
    
    Protected Sub SetTitle()
        Dim sTitle As String = Session("SiteTitle")
        If sTitle <> String.Empty Then
            sTitle += " - "
        End If
        Page.Header.Title = sTitle & "Western Union Permissions"
    End Sub

    Protected Sub HideAllPanels()
        Call SetEditVisibility(False)
        Call SetGridVisibility(False)
    End Sub

    Protected Sub PopulateUserGridview()
        Dim sSQL As String = "SELECT [key] 'UserKey', UserID + ' (' + FirstName + ' ' + LastName + ')' UserName, ISNULL(UserPermissions, 0) 'UserPermissions' FROM UserProfile WHERE  UserPermissions > 0 AND CustomerKey IN (579, 686, 788, 798) ORDER BY UserID"
        Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
        gvUsers.DataSource = dt
        gvUsers.DataBind()
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
            Err.Raise(ex.Message)
            WebMsgBox.Show("Error in ExecuteQueryToDataTable executing: " & sQuery & " : " & ex.Message)
        Finally
            oConn.Close()
        End Try
        ExecuteQueryToDataTable = oDataTable
    End Function

    Property pnUserID() As Integer
        Get
            Dim o As Object = ViewState("WUP_UserID")
            If o Is Nothing Then
                Return 0
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Integer)
            ViewState("WUP_UserID") = Value
        End Set
    End Property

    Protected Sub rcbUser_ItemsRequested(ByVal o As Object, ByVal e As Telerik.Web.UI.RadComboBoxItemsRequestedEventArgs)
        Dim s As String = e.Text
        Dim data As DataTable = GetUsers(e.Text)
        Dim itemOffset As Integer = e.NumberOfItems
        Dim endOffset As Integer = Math.Min(itemOffset + ITEMS_PER_REQUEST, data.Rows.Count)
        e.EndOfItems = endOffset = data.Rows.Count
        rcbUser.DataTextField = "UserName"
        rcbUser.DataValueField = "UserKey"
        For i As Int32 = itemOffset To endOffset - 1
            Dim rcbi As New RadComboBoxItem
            rcbi.Text = data.Rows(i)("UserName").ToString()
            rcbi.Value = data.Rows(i)("UserKey").ToString()
            rcbUser.Items.Add(rcbi)
        Next
    End Sub

    Protected Function GetUsers(Optional ByVal sFilter As String = "") As DataTable
        GetUsers = Nothing
        Dim sSQL As String = "SELECT [key] 'UserKey',  FirstName + ' ' + LastName + ' (' + UserId + ')' 'UserName' FROM UserProfile WHERE [Status] = 'Active' AND CustomerKey IN (579, 686, 788, 798)"
        If sFilter <> String.Empty Then
            sFilter = sFilter.Replace("'", "''")
            sSQL += " AND UserId LIKE '%" & sFilter & "%'"
        End If
        sSQL += " ORDER BY UserId"
        GetUsers = ExecuteQueryToDataTable(sSQL)
    End Function
    
    Protected Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim btn As Button = sender
        pnUserID = btn.CommandArgument
        Call ExtractPermissionsByUserKey(pnUserID)
    End Sub
    
    Protected Sub SetGridVisibility(ByVal bVisible As Boolean)
        trPageSizeDropDown.Visible = bVisible
        trUserGrid1.Visible = bVisible
        trUserGrid2.Visible = bVisible
    End Sub
    
    Protected Sub SetEditVisibility(ByVal bVisible As Boolean)
        trModifyPermissions1.Visible = bVisible
        trModifyPermissions2.Visible = bVisible
        trModifyPermissions2a.Visible = bVisible
        trModifyPermissions3.Visible = bVisible
    End Sub
    
    Protected Sub btnEditPermissions_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If IsNumeric(rcbUser.SelectedValue) Then   ' rcbUser.SelectedValue > 0
            pnUserID = rcbUser.SelectedValue
            Call ExtractPermissionsByUserKey(pnUserID)
        Else
            WebMsgBox.Show("Please select a user account")
            rcbUser.Focus()
        End If
    End Sub
    
    Protected Sub lnkbtnClearFilter_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        rcbUser.Text = String.Empty
        rcbUser.SelectedIndex = 0
        rcbUser.SelectedValue = String.Empty
    End Sub
    
    Protected Sub ExtractPermissionsByUserKey(ByVal nUserKey As Int32)
        Dim sSQL As String
        sSQL = "SELECT ISNULL(UserPermissions, 0) 'UserPermissions', UserID + ' (' + FirstName + ' ' + LastName + ')' 'UserName' FROM UserProfile WHERE [key] = " & nUserKey
        Dim dr As DataRow = ExecuteQueryToDataTable(sSQL).Rows(0)
        Dim nUserPermissions As Int32 = dr("UserPermissions")
        lblUser.Text = dr("UserName")
        cbIsWUInternalUser.Checked = nUserPermissions And USER_PERMISSION_WU_INTERNAL_USER
        cbCreateDeleteAccounts.Checked = nUserPermissions And USER_PERMISSION_WU_CREATE_DELETE_ACCOUNTS
        cbResetPasswords.Checked = nUserPermissions And USER_PERMISSION_WU_RESET_PASSWORDS
        cbViewProducts.Checked = nUserPermissions And USER_PERMISSION_WU_VIEW_STOCK
        cbOrderProducts.Checked = nUserPermissions And USER_PERMISSION_WU_ORDER_STOCK
        cbAccessReports.Checked = nUserPermissions And USER_PERMISSION_WU_ACCESS_REPORTS
        cbIsTSE.Checked = nUserPermissions And USER_PERMISSION_WU_IS_TSE
        cbIsPilotUser.Checked = nUserPermissions And USER_PERMISSION_WU_IS_PILOT_USER
        cbIsSales.Checked = nUserPermissions And USER_PERMISSION_WU_IS_SALES
        
        Call SetCheckboxState(cbIsWUInternalUser.Checked)

        trFindAccount.Visible = False
        Call SetEditVisibility(True)
        Call SetGridVisibility(False)
    End Sub
    
    Protected Sub SavePermissions(ByVal nUserKey As Int32)
        Dim sSQL As String
        sSQL = "SELECT ISNULL(UserPermissions, 0) 'UserPermissions' FROM UserProfile WHERE [key] = " & nUserKey
        Dim dr As DataRow = ExecuteQueryToDataTable(sSQL).Rows(0)
        Dim nUserPermissions As Int32 = dr("UserPermissions")
        If cbIsWUInternalUser.Checked Then
            nUserPermissions = (nUserPermissions Or USER_PERMISSION_WU_INTERNAL_USER)
        Else
            nUserPermissions = (nUserPermissions And (Not USER_PERMISSION_WU_INTERNAL_USER))
        End If
        If cbCreateDeleteAccounts.Checked Then
            nUserPermissions = (nUserPermissions Or USER_PERMISSION_WU_CREATE_DELETE_ACCOUNTS)
        Else
            nUserPermissions = (nUserPermissions And (Not USER_PERMISSION_WU_CREATE_DELETE_ACCOUNTS))
        End If
        If cbResetPasswords.Checked Then
            nUserPermissions = (nUserPermissions Or USER_PERMISSION_WU_RESET_PASSWORDS)
        Else
            nUserPermissions = (nUserPermissions And (Not USER_PERMISSION_WU_RESET_PASSWORDS))
        End If
        If cbViewProducts.Checked Then
            nUserPermissions = (nUserPermissions Or USER_PERMISSION_WU_VIEW_STOCK)
        Else
            nUserPermissions = (nUserPermissions And (Not USER_PERMISSION_WU_VIEW_STOCK))
        End If
        If cbOrderProducts.Checked Then
            nUserPermissions = (nUserPermissions Or USER_PERMISSION_WU_ORDER_STOCK)
        Else
            nUserPermissions = (nUserPermissions And (Not USER_PERMISSION_WU_ORDER_STOCK))
        End If
        If cbAccessReports.Checked Then
            nUserPermissions = (nUserPermissions Or USER_PERMISSION_WU_ACCESS_REPORTS)
        Else
            nUserPermissions = (nUserPermissions And (Not USER_PERMISSION_WU_ACCESS_REPORTS))
        End If
        If cbIsTSE.Checked Then
            nUserPermissions = (nUserPermissions Or USER_PERMISSION_WU_IS_TSE)
        Else
            nUserPermissions = (nUserPermissions And (Not USER_PERMISSION_WU_IS_TSE))
        End If
        If cbIsPilotUser.Checked Then
            nUserPermissions = (nUserPermissions Or USER_PERMISSION_WU_IS_PILOT_USER)
        Else
            nUserPermissions = (nUserPermissions And (Not USER_PERMISSION_WU_IS_PILOT_USER))
        End If
        If cbIsSales.Checked Then
            nUserPermissions = (nUserPermissions Or USER_PERMISSION_WU_IS_SALES)
        Else
            nUserPermissions = (nUserPermissions And (Not USER_PERMISSION_WU_IS_SALES))
        End If
        sSQL = "UPDATE UserProfile SET UserPermissions = " & nUserPermissions & " WHERE CustomerKey IN (579, 686, 788, 798) AND [key] = " & nUserKey
        Call ExecuteQueryToDataTable(sSQL)
    End Sub
    
    Protected Sub btnSavePermissions_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call SavePermissions(pnUserID)
        Call HideAllPanels()
        trFindAccount.Visible = True
        Call SetGridVisibility(True)
        Call PopulateUserGridview()
    End Sub

    Protected Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call HideAllPanels()
        trFindAccount.Visible = True
        Call SetGridVisibility(True)
    End Sub
    
    Protected Sub cbIsWUInternalUser_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim cb As CheckBox = sender
        Call SetCheckboxState(cb.Checked)
    End Sub
    
    Protected Sub SetCheckboxState(ByVal bEnabled As Boolean)
        If bEnabled Then
            cbCreateDeleteAccounts.Enabled = True
            cbResetPasswords.Enabled = True
            cbViewProducts.Enabled = True
            cbOrderProducts.Enabled = True
            cbAccessReports.Enabled = True
            cbIsTSE.Enabled = True
            cbIsSales.Enabled = True
        Else
            cbCreateDeleteAccounts.Checked = False
            cbCreateDeleteAccounts.Enabled = False
            cbResetPasswords.Checked = False
            cbResetPasswords.Enabled = False
            cbViewProducts.Checked = False
            cbViewProducts.Enabled = False
            cbOrderProducts.Checked = False
            cbOrderProducts.Enabled = False
            cbAccessReports.Checked = False
            cbAccessReports.Enabled = False
            cbIsTSE.Checked = False
            cbIsTSE.Enabled = False
            cbIsSales.Checked = False
            cbIsSales.Enabled = False
        End If
    End Sub
    
    Protected Sub ddlPageSize_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs)
        
        If ddlPageSize.SelectedValue <> "- select page size -" Then
            gvUsers.PageSize = ddlPageSize.SelectedValue
            Call PopulateUserGridview()
        End If
        
    End Sub
    
    Protected Sub gvUsers_PageIndexChanging(ByVal sender As Object, ByVal e As GridViewPageEventArgs) Handles gvUsers.PageIndexChanging
        
        gvUsers.PageIndex = e.NewPageIndex
        Call PopulateUserGridview()
        
    End Sub
    
    Protected Sub gvUsers_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs)
        Dim gvr As GridViewRow = e.Row
        If gvr.RowType = DataControlRowType.DataRow Then
            Dim lbl As Label = gvr.Cells(2).FindControl("lblPermissions")
            Dim nUserPermissions As Int32 = CInt(lbl.Text)
            Dim sbTextPermissions As New StringBuilder
            If (nUserPermissions And USER_PERMISSION_WU_INTERNAL_USER) > 0 Then
                sbTextPermissions.Append("WU Internal User; ")
            End If
            If (nUserPermissions And USER_PERMISSION_WU_CREATE_DELETE_ACCOUNTS) > 0 Then
                sbTextPermissions.Append("Create/delete accounts; ")
            End If
            If (nUserPermissions And USER_PERMISSION_WU_RESET_PASSWORDS) > 0 Then
                sbTextPermissions.Append("Reset passwords; ")
            End If
            If (nUserPermissions And USER_PERMISSION_WU_VIEW_STOCK) > 0 Then
                sbTextPermissions.Append("View stock; ")
            End If
            If (nUserPermissions And USER_PERMISSION_WU_ORDER_STOCK) > 0 Then
                sbTextPermissions.Append("Order stock; ")
            End If
            If (nUserPermissions And USER_PERMISSION_WU_ACCESS_REPORTS) > 0 Then
                sbTextPermissions.Append("Access reports; ")
            End If
            If (nUserPermissions And USER_PERMISSION_WU_IS_TSE) > 0 Then
                sbTextPermissions.Append("Is a TSE; ")
            End If
            If (nUserPermissions And USER_PERMISSION_WU_IS_SALES) > 0 Then
                sbTextPermissions.Append("Is Sales; ")
            End If
            If (nUserPermissions And USER_PERMISSION_WU_IS_PILOT_USER) > 0 Then
                sbTextPermissions.Append("Is Pilot User; ")
            End If
            If (nUserPermissions And USER_PERMISSION_ACCOUNT_HANDLER) > 0 Then
                sbTextPermissions.Append("Account handler; ")
            End If
            If (nUserPermissions And USER_PERMISSION_SITE_ADMINISTRATOR) > 0 Then
                sbTextPermissions.Append("Site administrator; ")
            End If
            If (nUserPermissions And USER_PERMISSION_DEPUTY_SITE_ADMINISTRATOR) > 0 Then
                sbTextPermissions.Append("Deputy site administrator; ")
            End If
            If (nUserPermissions And USER_PERMISSION_SITE_EDITOR) > 0 Then
                sbTextPermissions.Append("Site editor; ")
            End If
            If (nUserPermissions And USER_PERMISSION_DEPUTY_SITE_EDITOR) > 0 Then
                sbTextPermissions.Append("Deputy site editor; ")
            End If
            sbTextPermissions.Append(" (" & nUserPermissions.ToString & ")")
            lbl.Text = sbTextPermissions.ToString
        End If
    End Sub
    
    Protected Sub btnGenerateReport_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim sbText As New StringBuilder
        Call AddHTMLPreamble(sbText, "Western Union Permissions Report")
        sbText.Append(Bold("WESTERN UNION PERMISSIONS REPORT"))
        Call NewLine(sbText)
        sbText.Append("report generated " & DateTime.Now)
        Call NewLine(sbText)
        Call NewLine(sbText)
        sbText.Append("This report is divided into 2 sections. <b>Section 1</b> shows for each Western Union user account with one or more permissions, the permissions assigned to the account. <b>Section 2</b> shows for each permission type, the users who have been assigned that permission.")
        Call NewLine(sbText)
        Call NewLine(sbText)
        sbText.Append("<hr />")
        Call NewLine(sbText)
        Dim dtUsers As DataTable = ExecuteQueryToDataTable("SELECT UserID + ' (' + FirstName + ' ' + LastName + ')' 'User', UserPermissions FROM UserProfile up WHERE up.CustomerKey IN (579, 686, 798) AND ISNULL(UserPermissions,0) > 0 ORDER BY UserID")
        sbText.Append(Bold("PERMISSIONS BY ACCOUNT"))
        Call NewLine(sbText)
        Call NewLine(sbText)
        For Each drUser As DataRow In dtUsers.Rows
            sbText.Append(Bold(drUser("User").ToString))
            Call NewLine(sbText)
            Dim nUserPermissions As Int32 = CInt(drUser("UserPermissions"))
            If (nUserPermissions And USER_PERMISSION_WU_INTERNAL_USER) > 0 Then
                sbText.AppendLine("WU internal user, ")
            End If
            If (nUserPermissions And USER_PERMISSION_WU_CREATE_DELETE_ACCOUNTS) > 0 Then
                sbText.Append("Create/delete accounts, ")
            End If
            If (nUserPermissions And USER_PERMISSION_WU_RESET_PASSWORDS) > 0 Then
                sbText.Append("Reset passwords, ")
            End If
            If (nUserPermissions And USER_PERMISSION_WU_VIEW_STOCK) > 0 Then
                sbText.Append("View stock, ")
            End If
            If (nUserPermissions And USER_PERMISSION_WU_ORDER_STOCK) > 0 Then
                sbText.Append("Order stock, ")
            End If
            If (nUserPermissions And USER_PERMISSION_WU_ACCESS_REPORTS) > 0 Then
                sbText.Append("Access reports, ")
            End If
            If (nUserPermissions And USER_PERMISSION_WU_IS_TSE) > 0 Then
                sbText.Append("Is a TSE, ")
            End If
            If (nUserPermissions And USER_PERMISSION_WU_IS_SALES) > 0 Then
                sbText.Append("Is Sales, ")
            End If
            If (nUserPermissions And USER_PERMISSION_WU_IS_PILOT_USER) > 0 Then
                sbText.Append("Is Pilot User, ")
            End If
            If (nUserPermissions And USER_PERMISSION_ACCOUNT_HANDLER) > 0 Then
                sbText.Append("Account handler, ")
            End If
            If (nUserPermissions And USER_PERMISSION_SITE_ADMINISTRATOR) > 0 Then
                sbText.Append("Site administrator, ")
            End If
            If (nUserPermissions And USER_PERMISSION_DEPUTY_SITE_ADMINISTRATOR) > 0 Then
                sbText.Append("Deputy site administrator, ")
            End If
            If (nUserPermissions And USER_PERMISSION_SITE_EDITOR) > 0 Then
                sbText.Append("Site editor, ")
            End If
            If (nUserPermissions And USER_PERMISSION_DEPUTY_SITE_EDITOR) > 0 Then
                sbText.Append("Deputy site editor, ")
            End If
            sbText.Remove(sbText.Length - 2, 2)
            Call NewLine(sbText)
            Call NewLine(sbText)
        Next
        sbText.Append("<hr />")
        Call NewLine(sbText)
        sbText.Append(Bold("ACCOUNTS BY PERMISSION"))
        Call NewLine(sbText)
        For Each nPermission As Int32 In IListOfPermissions()
            Dim bUserFound As Boolean = False
            Call NewLine(sbText)
            Select Case nPermission
                Case USER_PERMISSION_ACCOUNT_HANDLER
                    sbText.AppendLine(Bold("Account handler"))
                Case USER_PERMISSION_WU_INTERNAL_USER
                    sbText.AppendLine(Bold("WU Internal User"))
                Case USER_PERMISSION_WU_CREATE_DELETE_ACCOUNTS
                    sbText.Append(Bold("Create/delete accounts"))
                Case USER_PERMISSION_WU_RESET_PASSWORDS
                    sbText.Append(Bold("Reset passwords"))
                Case USER_PERMISSION_WU_VIEW_STOCK
                    sbText.Append(Bold("View stock"))
                Case USER_PERMISSION_WU_ORDER_STOCK
                    sbText.Append(Bold("Order stock"))
                Case USER_PERMISSION_WU_ACCESS_REPORTS
                    sbText.Append(Bold("Access reports"))
                Case USER_PERMISSION_WU_IS_TSE
                    sbText.Append(Bold("Is a TSE"))
                Case USER_PERMISSION_WU_IS_SALES
                    sbText.Append(Bold("Is Sales"))
                Case USER_PERMISSION_WU_IS_PILOT_USER
                    sbText.Append(Bold("Is pilot User"))
              
                Case USER_PERMISSION_SITE_ADMINISTRATOR
                    sbText.Append(Bold("Site administrator"))
                Case USER_PERMISSION_DEPUTY_SITE_ADMINISTRATOR
                    sbText.Append(Bold("Deputy site administrator"))
                    'Case USER_PERMISSION_VIEW_STOCK
                    '    sbText.Append(Bold("View Stock"))
                Case USER_PERMISSION_SITE_EDITOR
                    sbText.Append(Bold("Site editor"))
                Case USER_PERMISSION_DEPUTY_SITE_EDITOR
                    sbText.Append(Bold("Deputy site editor"))
                    'Case USER_PERMISSION_CREATE_STOCK_BOOKING
                    '    sbText.Append(Bold("Create Stock Booking"))
                    'Case USER_PERMISSION_PRINT_ON_DEMAND_TAB
                    '    sbText.Append(Bold("Print on demand tab"))
                    'Case USER_PERMISSION_ADVANCED_PERMISSIONS_TAB
                    '    sbText.Append(Bold("Advanced permissions tab"))
                    'Case USER_PERMISSION_FILE_UPLOAD_TAB
                    '    sbText.Append(Bold("File upload tab"))
            End Select
            Call NewLine(sbText)
            For Each drUser As DataRow In dtUsers.Rows
                Dim nUserPermissions As Int32 = CInt(drUser("UserPermissions"))
                If (nPermission And nUserPermissions) > 0 Then
                    Dim sUser As String = drUser("User").ToString()
                    bUserFound = True
                    sbText.Append(sUser)
                    Call NewLine(sbText)
                End If
            Next
            If Not bUserFound Then
                sbText.Append("(No accounts assigned)")
                Call NewLine(sbText)
            End If
        Next
        Call NewLine(sbText)
        sbText.Append("<hr />")
        Call NewLine(sbText)
        sbText.Append("[end]")
        Call AddHTMLPostamble(sbText)
        Call ExportData(sbText.ToString, "WesternUnionPermissionsReport")
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
    <form id="Form1" runat="Server">
    <main:Header ID="ctlHeader" runat="server" />
    <table style="width: 100%" cellpadding="0" cellspacing="0">
        <tr class="bar_reports">
            <td style="width: 50%; white-space: nowrap">
            </td>
            <td style="width: 50%; white-space: nowrap" align="right">
            </td>
        </tr>
    </table>
    <asp:Panel ID="pnlNewMessage" runat="server" Width="100%">
        <table style="width: 100%">
            <tr>
                <td style="width: 1%">
                </td>
                <td style="width: 11%">
                    &nbsp;
                </td>
                <td style="width: 32%">
                    &nbsp;
                </td>
                <td style="width: 23%">
                </td>
                <td style="width: 32%">
                </td>
                <td style="width: 1%">
                </td>
            </tr>
            <tr>
                <td>
                    &nbsp;
                </td>
                <td colspan="4">
                    <asp:Label ID="lblLegendTopicType" runat="server" Font-Bold="True" Font-Names="Verdana"
                        Font-Size="Small" ForeColor="Navy" Text="Western Union Permissions" />
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td>
                    &nbsp;
                </td>
                <td colspan="4" align="right">
                    <asp:Button ID="btnGenerateReport" runat="server" OnClick="btnGenerateReport_Click"
                        Text="Generate report" />
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr id="trCustomerAccount" runat="server" visible="false">
                <td>
                    &nbsp;
                </td>
                <td align="right">
                    <asp:Label ID="Label2" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Customer account:" />
                </td>
                <td colspan="3">
                    <asp:DropDownList ID="DropDownList1" runat="server">
                    </asp:DropDownList>
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr id="trFindAccount" runat="server" visible="false">
                <td>
                    &nbsp;
                </td>
                <td align="right">
                    <asp:Label ID="Label7" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="User account:" />
                </td>
                <td colspan="3">
                    <telerik:RadComboBox ID="rcbUser" runat="server" AutoPostBack="True" CausesValidation="False"
                        EnableLoadOnDemand="True" EnableVirtualScrolling="True" Filter="Contains" Font-Names="Verdana"
                        Font-Size="XX-Small" HighlightTemplatedItems="true" OnItemsRequested="rcbUser_ItemsRequested"
                        Width="300px" ShowMoreResultsBox="True" ToolTip="Shows all users when no search text is specified. Search for users by typing an agent code or name." />
                    &nbsp;<asp:LinkButton ID="lnkbtnClearFilter" runat="server" Font-Names="Verdana"
                        Font-Size="XX-Small" OnClick="lnkbtnClearFilter_Click">clear filter</asp:LinkButton>
                    &nbsp;<asp:Button ID="btnEditPermissions" runat="server" Text="Edit permissions"
                        OnClick="btnEditPermissions_Click" />
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr id="trModifyPermissions1" runat="server" visible="true">
                <td>
                    &nbsp;
                </td>
                <td align="right">
                    &nbsp;
                </td>
                <td colspan="3">
                    <asp:Label ID="lblUser" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="X-Small" />
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr id="trModifyPermissions2" runat="server" visible="true">
                <td>
                    &nbsp;
                </td>
                <td align="right">
                    &nbsp;
                </td>
                <td colspan="3">
                    <asp:CheckBox ID="cbIsWUInternalUser" runat="server" Font-Names="Arial" Font-Size="XX-Small"
                        Text="Is WU Internal User" AutoPostBack="True" OnCheckedChanged="cbIsWUInternalUser_CheckedChanged" />
                    &nbsp;<asp:CheckBox ID="cbCreateDeleteAccounts" runat="server" Font-Names="Arial"
                        Font-Size="XX-Small" Text="Create / delete accounts" />
                    &nbsp;<asp:CheckBox ID="cbResetPasswords" runat="server" Font-Names="Arial" Font-Size="XX-Small"
                        Text="Reset passwords" />
                    &nbsp;<asp:CheckBox ID="cbViewProducts" runat="server" Font-Names="Arial" Font-Size="XX-Small"
                        Text="View products" />
                    &nbsp;<asp:CheckBox ID="cbOrderProducts" runat="server" Font-Names="Arial" Font-Size="XX-Small"
                        Text="Order products" />
                    &nbsp;&nbsp;<asp:CheckBox ID="cbAccessReports" runat="server" Font-Names="Arial"
                        Font-Size="XX-Small" Text="Access reports" />
                    &nbsp;<asp:CheckBox ID="cbIsTSE" runat="server" Font-Names="Arial" Font-Size="XX-Small"
                        Text="Is a TSE" />
                    &nbsp;<asp:CheckBox ID="cbIsSales" runat="server" Font-Names="Arial" 
                        Font-Size="XX-Small" Text="Is Sales" />
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr id="trModifyPermissions2a" runat="server" visible="true">
                <td>
                    &nbsp;
                </td>
                <td align="right">
                    &nbsp;
                </td>
                <td colspan="3">
                    <asp:CheckBox ID="cbIsPilotUser" runat="server" AutoPostBack="True" Font-Names="Arial"
                        Font-Size="XX-Small" OnCheckedChanged="cbIsWUInternalUser_CheckedChanged" Text="Is Pilot User" />
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr id="trModifyPermissions3" runat="server" visible="true">
                <td>
                    &nbsp;
                </td>
                <td align="right">
                    &nbsp;
                </td>
                <td colspan="3">
                    <asp:Button ID="btnSavePermissions" runat="server" Text="Save" Width="100px" Height="26px"
                        OnClick="btnSavePermissions_Click" />
                    &nbsp;<asp:Button ID="btnCancel" runat="server" Text="Cancel" OnClick="btnCancel_Click" />
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr id="trSpace">
                <td>
                    &nbsp;
                </td>
                <td align="right">
                    &nbsp;
                </td>
                <td colspan="3">
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr id="trUserGrid1" runat="server" visible="true">
                <td>
                    &nbsp;
                </td>
                <td align="right">
                    &nbsp;
                </td>
                <td colspan="3">
                    <asp:Label ID="Label8" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Users:" />
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr runat="server" id="trUserGrid2" visible="false">
                <td>
                    &nbsp;
                </td>
                <td align="right">
                    &nbsp;
                </td>
                <td colspan="3">
                    <asp:GridView ID="gvUsers" runat="server" AutoGenerateColumns="False" CellPadding="2"
                        Font-Names="Arial" Font-Size="XX-Small" OnRowDataBound="gvUsers_RowDataBound"
                        AllowPaging="true" PageSize="5" Width="100%">
                        <Columns>
                            <asp:TemplateField>
                                <ItemTemplate>
                                    <asp:Button ID="btnEdit" runat="server" CommandArgument='<%# Container.DataItem("UserKey")%>'
                                        OnClick="btnEdit_Click" Text="edit" />
                                </ItemTemplate>
                                <ItemStyle Width="150px" Wrap="False" />
                            </asp:TemplateField>
                            <asp:BoundField DataField="UserName" HeaderText="User" ReadOnly="True" SortExpression="UserName" />
                            <asp:TemplateField HeaderText="Permissions" SortExpression="UserPermissions">
                                <ItemTemplate>
                                    <asp:Label ID="lblPermissions" runat="server" Text='<%# Bind("UserPermissions") %>' />
                                </ItemTemplate>
                            </asp:TemplateField>
                        </Columns>
                        <EmptyDataTemplate>
                            no users permissioned
                        </EmptyDataTemplate>
                    </asp:GridView>
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr id="trPageSizeDropDown" runat="server" visible="false">
                <td>
                    &nbsp;
                </td>
                <td align="right">
                    &nbsp;
                </td>
                <td colspan="3">
                    <asp:Label ID="Label1" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Page Size:" />
                    <asp:DropDownList ID="ddlPageSize" Font-Names="Verdana" Font-Size="XX-Small" AutoPostBack="true"
                        OnSelectedIndexChanged="ddlPageSize_SelectedIndexChanged" runat="server">
                        <asp:ListItem Text="- select page size -" Value="- select page size -"></asp:ListItem>
                        <asp:ListItem Text="5" Value="5"></asp:ListItem>
                        <asp:ListItem Text="20" Value="20"></asp:ListItem>
                        <asp:ListItem Text="100" Value="100"></asp:ListItem>
                    </asp:DropDownList>
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