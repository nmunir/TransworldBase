<%@ Page Language="VB" Theme="AIMSDefault" ValidateRequest="false" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>
<%@ import Namespace="System.Collections.Generic" %>
<%@ Register TagPrefix="FCKeditorV2" Namespace="FredCK.FCKeditorV2" Assembly="FredCK.FCKeditorV2" %>

<script runat="server">

    ' See http://www.beansoftware.com/ASP.NET-Tutorials/Modify-Web.Config-Run-Time.aspx for writing to web.config
    ' or http://www.dotnetcurry.com/ShowArticle.aspx?ID=102&AspxAutoDetectCookieSupport=1
   
    Const USER_PERMISSION_ACCOUNT_HANDLER As Integer = 1
    Const USER_PERMISSION_SITE_ADMINISTRATOR As Integer = 2
    Const USER_PERMISSION_DEPUTY_ADMINISTRATOR As Integer = 4
    Const USER_PERMISSION_NOTICE_BOARD_EDITOR As Integer = 8
    Const USER_PERMISSION_DEPUTY_NOTICE_BOARD_EDITOR As Integer = 16
   
    Const STYLESHEET_FILENAME_WORKING As String = "sprint.css"
    Const STYLESHEET_FILENAME_DEFAULT As String = "sprint_default.css"
    Const DEFAULT_STYLESHEET_PATH As String = "~\css\sprint.css"
   
    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString
   
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsNumeric(Session("CustomerKey")) Then
            Server.Transfer("session_expired.aspx")
            Exit Sub
        End If
        If Not IsPostBack Then
            Call GetSiteFeatures()
            Call GetRoles()
            Call PopulateDeputySiteAdministratorDropdown()
            Call PopulateNoticeBoardEditorDropdown()
            Call PopulateDeputyNoticeBoardEditorDropdown()
        End If
        Call SetTitle()
    End Sub
   
    Protected Sub SetTitle()
        Dim sTitle As String = Session("SiteTitle")
        If sTitle <> String.Empty Then
            sTitle += " - "
        End If
        Page.Header.Title = sTitle & "Administrator"
    End Sub
   
    Protected Sub GetSiteFeatures()
        Dim oDataTable As New DataTable
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter("spASPNET_SiteContent3", oConn)
       
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
       
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@Action", SqlDbType.NVarChar, 50))
        oAdapter.SelectCommand.Parameters("@Action").Value = "GET"
       
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SiteKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@SiteKey").Value = Session("SiteKey")
       
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ContentType", SqlDbType.NVarChar, 50))
        oAdapter.SelectCommand.Parameters("@ContentType").Value = "SiteSettings"

        Try
            oAdapter.Fill(oDataTable)
        Catch ex As Exception
            WebMsgBox.Show("GetSiteFeatures: " & ex.Message)
        Finally
            oConn.Close()
        End Try

        Dim dr As DataRow = oDataTable.Rows(0)
        tbSiteTitle.Text = dr("SiteTitle")
        cbShowProductsWithZeroStock.Checked = dr("ShowZeroStock")
        cbApplyMaxGrabs.Checked = dr("ApplyMaxGrabs")
        cbMultipleAddressOrders.Checked = dr("MultipleAddressOrders")
        cbOrderAuthorisation.Checked = dr("OrderAuthorisation")
        cbProductAuthorisation.Checked = dr("ProductAuthorisation")
        cbCalendarManagement.Checked = dr("CalendarManagement")
        cbUserPermissions.Checked = dr("UserPermissions")
        cbFileUpload.Checked = dr("FileUpload")
        cbUserCanChangeCostCentre.Checked = dr("UserCanChangeCostCentre")
        cbShowNotes.Checked = dr("ShowNotes")
        cbProductOwners.Checked = dr("ProductOwners")
        cbOnDemandProducts.Checked = dr("OnDemandProducts")
        cbZeroStockNotifications.Checked = dr("ZeroStockNotifications")
        cbCustomLetters.Checked = dr("CustomLetters")
        If cbShowProductsWithZeroStock.Checked Then
            cbZeroStockNotifications.Enabled = True
        Else
            cbZeroStockNotifications.Enabled = False
        End If
        FCKedAuthorisationAdvisory.Value = dr("AuthorisationAdvisory")
        If cbOrderAuthorisation.Checked Then
            trAuthorisationAdvisory.Visible = True
        Else
            trAuthorisationAdvisory.Visible = False
        End If
        cbShowSellingPrice.Checked = dr("SellingPrice")
        pnlAuthorisation.Visible = dr("OrderAuthorisation") Or dr("ProductAuthorisation")
        If pnlAuthorisation.Visible Then
            Call PopulateExemptionListBoxes()
        End If
    End Sub
   
    Protected Sub GetRoles()
        Dim sSQL As String = String.Empty
        sSQL = "SELECT FirstName, LastName, UserId, UserPermissions FROM UserProfile WHERE ISNULL(UserPermissions,0) > 0 AND CustomerKey = " & Session("CustomerKey")
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        Try
            oConn.Open()
            Dim oSqlDataReader As SqlDataReader = oCmd.ExecuteReader
            If oSqlDataReader.HasRows Then
                While oSqlDataReader.Read
                    If CInt(oSqlDataReader("UserPermissions")) And USER_PERMISSION_ACCOUNT_HANDLER Then
                        lblAccountHandler.Text = oSqlDataReader("FirstName") & " " & oSqlDataReader("LastName") & " (" & oSqlDataReader("UserId") & ")"
                    End If
                    If CInt(oSqlDataReader("UserPermissions")) And USER_PERMISSION_SITE_ADMINISTRATOR Then
                        lblSiteAdministrator.Text = oSqlDataReader("FirstName") & " " & oSqlDataReader("LastName") & " (" & oSqlDataReader("UserId") & ")"
                    End If
                End While
            End If
        Catch ex As Exception
            WebMsgBox.Show("GetRoles: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
   
    Protected Sub PopulateDeputySiteAdministratorDropdown()
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataReader As SqlDataReader = Nothing
        Dim sSQL As String = "SELECT FirstName + ' ' + LastName + ' (' + UserId + ')' Name, [key], UserPermissions FROM UserProfile WHERE Type LIKE 'SuperUser' AND Status LIKE 'active' AND DeletedFlag = 0 AND CustomerKey = " & Session("CustomerKey") & " ORDER BY FirstName"
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        Dim li As ListItem
        li = New ListItem("- please select -", 0)
        ddlRoleDeputyAdministrator.Items.Add(li)
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            While oDataReader.Read
                li = New ListItem(oDataReader("Name"), oDataReader("Key"))
                ddlRoleDeputyAdministrator.Items.Add(li)
                If Not IsDBNull(oDataReader("UserPermissions")) Then
                    If oDataReader("UserPermissions") And USER_PERMISSION_DEPUTY_ADMINISTRATOR Then
                        ddlRoleDeputyAdministrator.SelectedIndex = ddlRoleDeputyAdministrator.Items.Count - 1
                        hidDeputySiteAdministratorKey.Value = oDataReader("Key")
                    End If
                End If
            End While
        Catch ex As Exception
            WebMsgBox.Show("PopulateDeputySiteAdministratorDropdown: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
   
    Protected Sub PopulateNoticeBoardEditorDropdown()
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataReader As SqlDataReader = Nothing
        Dim sSQL As String = "SELECT FirstName + ' ' + LastName + ' (' + UserId + ')' Name, [key], UserPermissions FROM UserProfile WHERE Status LIKE 'active' AND DeletedFlag = 0 AND CustomerKey = " & Session("CustomerKey") & " ORDER BY FirstName"
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        Dim li As ListItem
        li = New ListItem("- please select -", 0)
        ddlRoleNoticeBoardEditor.Items.Add(li)
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            While oDataReader.Read
                li = New ListItem(oDataReader("Name"), oDataReader("Key"))
                ddlRoleNoticeBoardEditor.Items.Add(li)
                If Not IsDBNull(oDataReader("UserPermissions")) Then
                    If oDataReader("UserPermissions") And USER_PERMISSION_NOTICE_BOARD_EDITOR Then
                        ddlRoleNoticeBoardEditor.SelectedIndex = ddlRoleNoticeBoardEditor.Items.Count - 1
                        hidNoticeBoardEditorKey.Value = oDataReader("Key")
                    End If
                End If
            End While
        Catch ex As Exception
            WebMsgBox.Show("PopulateNoticeBoardEditorDropdown: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
   
    Protected Sub PopulateDeputyNoticeBoardEditorDropdown()
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataReader As SqlDataReader = Nothing
        Dim sSQL As String = "SELECT FirstName + ' ' + LastName + ' (' + UserId + ')' Name, [key], UserPermissions FROM UserProfile WHERE Status LIKE 'active' AND DeletedFlag = 0 AND CustomerKey = " & Session("CustomerKey") & " ORDER BY FirstName"
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        Dim li As ListItem
        li = New ListItem("- please select -", 0)
        ddlRoleDeputyNoticeBoardEditor.Items.Add(li)
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            While oDataReader.Read
                li = New ListItem(oDataReader("Name"), oDataReader("Key"))
                ddlRoleDeputyNoticeBoardEditor.Items.Add(li)
                If Not IsDBNull(oDataReader("UserPermissions")) Then
                    If oDataReader("UserPermissions") And USER_PERMISSION_DEPUTY_NOTICE_BOARD_EDITOR Then
                        ddlRoleDeputyNoticeBoardEditor.SelectedIndex = ddlRoleDeputyNoticeBoardEditor.Items.Count - 1
                        hidDeputyNoticeBoardEditorKey.Value = oDataReader("Key")
                    End If
                End If
            End While
        Catch ex As Exception
            WebMsgBox.Show("PopulateDeputyNoticeBoardEditorDropdown: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub

    Protected Sub ddlRoleDeputyAdministrator_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        Dim nOldUserKey As Integer
        If IsNumeric(hidDeputySiteAdministratorKey.Value) Then
            nOldUserKey = hidDeputySiteAdministratorKey.Value
        Else
            nOldUserKey = 0
        End If
        Call SetNewDeputySiteAdministrator(nOldUserKey, ddlRoleDeputyAdministrator.SelectedValue)
        hidDeputySiteAdministratorKey.Value = ddl.SelectedValue
    End Sub

    Protected Sub ddlRoleNoticeBoardEditor_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        Dim nOldUserKey As Integer
        If IsNumeric(hidNoticeBoardEditorKey.Value) Then
            nOldUserKey = hidNoticeBoardEditorKey.Value
        Else
            nOldUserKey = 0
        End If
        Call SetNewNoticeBoardEditor(nOldUserKey, ddl.SelectedValue)
        hidNoticeBoardEditorKey.Value = ddl.SelectedValue
    End Sub

    Protected Sub ddlRoleDeputyNoticeBoardEditor_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        Dim nOldUserKey As Integer
        If IsNumeric(hidDeputyNoticeBoardEditorKey.Value) Then
            nOldUserKey = hidDeputyNoticeBoardEditorKey.Value
        Else
            nOldUserKey = 0
        End If
        Call SetNewDeputyNoticeBoardEditor(nOldUserKey, ddl.SelectedValue)
        hidDeputyNoticeBoardEditorKey.Value = ddl.SelectedValue
    End Sub
   
    Protected Sub SetNewDeputySiteAdministrator(ByVal nOldUserKey As Integer, ByVal nNewUserKey As Integer)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_UserProfile_SetNewDeputySiteAdministrator2", oConn)
        oCmd.CommandType = CommandType.StoredProcedure

        Dim paramOldUserKey As SqlParameter = New SqlParameter("@OldUserKey", SqlDbType.Int)
        paramOldUserKey.Value = nOldUserKey
        oCmd.Parameters.Add(paramOldUserKey)

        Dim paramNewUserKey As SqlParameter = New SqlParameter("@NewUserKey", SqlDbType.Int)
        paramNewUserKey.Value = nNewUserKey
        oCmd.Parameters.Add(paramNewUserKey)

        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As SqlException
            WebMsgBox.Show(ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
   
    Protected Sub SetNewNoticeBoardEditor(ByVal nOldUserKey As Integer, ByVal nNewUserKey As Integer)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_UserProfile_SetNewNoticeBoardEditor2", oConn)
        oCmd.CommandType = CommandType.StoredProcedure

        Dim paramOldUserKey As SqlParameter = New SqlParameter("@OldUserKey", SqlDbType.Int)
        paramOldUserKey.Value = nOldUserKey
        oCmd.Parameters.Add(paramOldUserKey)

        Dim paramNewUserKey As SqlParameter = New SqlParameter("@NewUserKey", SqlDbType.Int)
        paramNewUserKey.Value = nNewUserKey
        oCmd.Parameters.Add(paramNewUserKey)

        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As SqlException
            WebMsgBox.Show(ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
   
    Protected Sub SetNewDeputyNoticeBoardEditor(ByVal nOldUserKey As Integer, ByVal nNewUserKey As Integer)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_UserProfile_SetNewDeputyNoticeBoardEditor2", oConn)
        oCmd.CommandType = CommandType.StoredProcedure

        Dim paramOldUserKey As SqlParameter = New SqlParameter("@OldUserKey", SqlDbType.Int)
        paramOldUserKey.Value = nOldUserKey
        oCmd.Parameters.Add(paramOldUserKey)

        Dim paramNewUserKey As SqlParameter = New SqlParameter("@NewUserKey", SqlDbType.Int)
        paramNewUserKey.Value = nNewUserKey
        oCmd.Parameters.Add(paramNewUserKey)

        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As SqlException
            WebMsgBox.Show(ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
   
    Protected Sub btnSaveSiteFeatureChanges_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call SaveSiteFeatureChanges()
    End Sub
   
    Protected Sub SaveSiteFeatureChanges()
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_SiteContent3", oConn)
        oCmd.CommandType = CommandType.StoredProcedure

        Dim paramAction As New SqlParameter("@Action", SqlDbType.NVarChar, 50)
        paramAction.Value = "SET"
        oCmd.Parameters.Add(paramAction)

        Dim paramSiteKey As New SqlParameter("@SiteKey", SqlDbType.Int)
        paramSiteKey.Value = Session("SiteKey")
        oCmd.Parameters.Add(paramSiteKey)

        Dim paramContentType As New SqlParameter("@ContentType", SqlDbType.NVarChar, 50)
        paramContentType.Value = "SiteSettings3"
        oCmd.Parameters.Add(paramContentType)

        Dim paramSiteTitle As New SqlParameter("@SiteTitle", SqlDbType.NVarChar, 50)
        paramSiteTitle.Value = tbSiteTitle.Text
        oCmd.Parameters.Add(paramSiteTitle)

        Dim paramShowZeroStock As New SqlParameter("@ShowZeroStock", SqlDbType.Bit)
        paramShowZeroStock.Value = cbShowProductsWithZeroStock.Checked
        oCmd.Parameters.Add(paramShowZeroStock)

        Dim paramApplyMaxGrabs As New SqlParameter("@ApplyMaxGrabs", SqlDbType.Bit)
        paramApplyMaxGrabs.Value = cbApplyMaxGrabs.Checked
        oCmd.Parameters.Add(paramApplyMaxGrabs)

        Dim paramMultipleAddressOrders As New SqlParameter("@MultipleAddressOrders", SqlDbType.Bit)
        paramMultipleAddressOrders.Value = cbMultipleAddressOrders.Checked
        oCmd.Parameters.Add(paramMultipleAddressOrders)

        Dim paramOrderAuthorisation As New SqlParameter("@OrderAuthorisation", SqlDbType.Bit)
        paramOrderAuthorisation.Value = cbOrderAuthorisation.Checked
        oCmd.Parameters.Add(paramOrderAuthorisation)

        Dim paramProductAuthorisation As New SqlParameter("@ProductAuthorisation", SqlDbType.Bit)
        paramProductAuthorisation.Value = cbProductAuthorisation.Checked
        oCmd.Parameters.Add(paramProductAuthorisation)

        Dim paramShowNotes As New SqlParameter("@ShowNotes", SqlDbType.Bit)
        paramShowNotes.Value = cbShowNotes.Checked
        oCmd.Parameters.Add(paramShowNotes)

        Dim paramUserCanChangeCostCentre As New SqlParameter("@UserCanChangeCostCentre", SqlDbType.Bit)
        paramUserCanChangeCostCentre.Value = cbUserCanChangeCostCentre.Checked
        oCmd.Parameters.Add(paramUserCanChangeCostCentre)

        Dim paramZeroStockNotifications As New SqlParameter("@ZeroStockNotifications", SqlDbType.Bit)
        paramZeroStockNotifications.Value = cbZeroStockNotifications.Checked
        oCmd.Parameters.Add(paramZeroStockNotifications)

        Dim paramSellingPrice As New SqlParameter("@SellingPrice", SqlDbType.Bit)
        paramSellingPrice.Value = cbShowSellingPrice.Checked
        oCmd.Parameters.Add(paramSellingPrice)
       
        Dim paramAuthorisationAdvisory As New SqlParameter("@AuthorisationAdvisory", SqlDbType.NVarChar, 1000)
        paramAuthorisationAdvisory.Value = FCKedAuthorisationAdvisory.Value.Trim
        oCmd.Parameters.Add(paramAuthorisationAdvisory)

        Dim paramMisc1 As New SqlParameter("@Misc1", SqlDbType.Bit)
        paramMisc1.Value = False
        oCmd.Parameters.Add(paramMisc1)

        Dim paramMisc2 As New SqlParameter("@Misc2", SqlDbType.Bit)
        paramMisc2.Value = False
        oCmd.Parameters.Add(paramMisc2)

        Dim paramMisc3 As New SqlParameter("@Misc3", SqlDbType.Bit)
        paramMisc3.Value = False
        oCmd.Parameters.Add(paramMisc3)

        Dim paramMisc4 As New SqlParameter("@Misc4", SqlDbType.Bit)
        paramMisc4.Value = False
        oCmd.Parameters.Add(paramMisc4)

        Dim paramMisc5 As New SqlParameter("@Misc5", SqlDbType.Bit)
        paramMisc5.Value = False
        oCmd.Parameters.Add(paramMisc5)

        Dim paramMisc6 As New SqlParameter("@Misc6", SqlDbType.Bit)
        paramMisc6.Value = False
        oCmd.Parameters.Add(paramMisc6)

        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As SqlException
            WebMsgBox.Show("SaveSiteFeatureChanges: " & ex.Message)
        Finally
            oConn.Close()
        End Try
        Session("SiteTitle") = tbSiteTitle.Text
        Call SetTitle()
    End Sub
   
    Protected Sub cbOrderAuthorisation_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim cb As CheckBox = sender
        If cb.Checked Then
            cbProductAuthorisation.Checked = False
            trAuthorisationAdvisory.Visible = True
        Else
            trAuthorisationAdvisory.Visible = False
        End If
    End Sub

    Protected Sub cbProductAuthorisation_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim cb As CheckBox = sender
        If cb.Checked Then
            cbOrderAuthorisation.Checked = False
            trAuthorisationAdvisory.Visible = False
        End If
    End Sub
   
    Protected Sub cbShowProductsWithZeroStock_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim cb As CheckBox = sender
        If cb.Checked Then
            cbZeroStockNotifications.Enabled = True
        Else
            If cbZeroStockNotifications.Checked Then
                cbZeroStockNotifications.Checked = False
                WebMsgBox.Show("WARNING: Zero Stock Notifications have been disabled because you have disabled Show Products with Zero Stock. Any notifications already requested will still be sent when stock becomes available.")
            End If
            cbZeroStockNotifications.Enabled = False
        End If
    End Sub
   
    Protected Sub PopulateExemptionListBoxes()
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Customer_GetActiveUsersAuthExempt", oConn)
        oCmd.CommandType = CommandType.StoredProcedure

        oCmd.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
        oCmd.Parameters("@CustomerKey").Value = Session("CustomerKey")

        oCmd.Parameters.Add(New SqlParameter("@IsExempt", SqlDbType.Bit))
        oCmd.Parameters("@IsExempt").Value = 0
        lbUsersExempt.Items.Clear()
        lbUsersUnExempt.Items.Clear()
        Try
            oConn.Open()
            Dim oSqlDataReader As SqlDataReader = oCmd.ExecuteReader
            If oSqlDataReader.HasRows Then
                While oSqlDataReader.Read
                    lbUsersUnExempt.Items.Add(New ListItem(oSqlDataReader("UserName"), oSqlDataReader("UserKey")))
                End While
            End If
            oSqlDataReader.Close()
            oCmd.Parameters("@IsExempt").Value = 1
            oSqlDataReader = oCmd.ExecuteReader
            If oSqlDataReader.HasRows Then
                While oSqlDataReader.Read
                    lbUsersExempt.Items.Add(New ListItem(oSqlDataReader("UserName"), oSqlDataReader("UserKey")))
                End While
            End If
        Catch ex As Exception
            WebMsgBox.Show("PopulateExemptionListBoxes: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
   
    Protected Sub lnkbtnSetExemption_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call SetExemption()
        Call PopulateExemptionListBoxes()
    End Sub

    Protected Sub SetExemption()
        For Each item As ListItem In lbUsersUnExempt.Items
            If item.Selected Then
                Call ExemptUser(item.Value)
            End If
        Next
    End Sub
   
    Protected Sub ExemptUser(ByVal UserKey As String)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand
        Dim sSQL As String = "INSERT INTO LogisticProductAuthoriseExemptions (UserKey, LastModifiedDateTime, LastUpdateBy) VALUES ("
        sSQL += UserKey & ", GETDATE(), " & Session("UserKey") & ")"
        Try
            oConn.Open()
            oCmd = New SqlCommand(sSQL, oConn)
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            WebMsgBox.Show("ExemptUser: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
   
    Protected Sub lnkbtnRemoveExemption_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call RemoveExemption()
        Call PopulateExemptionListBoxes()
    End Sub
   
    Protected Sub RemoveExemption()
        For Each item As ListItem In lbUsersExempt.Items
            If item.Selected Then
                Call UnExemptUser(item.Value)
            End If
        Next
    End Sub

    Protected Sub UnExemptUser(ByVal UserKey As String)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand
        Dim sSQL As String = "DELETE FROM LogisticProductAuthoriseExemptions WHERE UserKey = " & UserKey
        Try
            oConn.Open()
            oCmd = New SqlCommand(sSQL, oConn)
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            WebMsgBox.Show("UnExemptUser: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub

</script>
<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Site Administrator</title>
</head>
<body>
    <form id="frmAdministrator" runat="server">
      <main:Header id="ctlHeader" runat="server"></main:Header>
        <table width="100%" cellpadding="0" cellspacing="0">
            <tr class="bar_siteadministrator">
                <td style="width:50%; white-space:nowrap">
                </td>
                <td style="width:50%; white-space:nowrap" align="right">
                </td>
            </tr>
        </table>
            <strong style="color: navy; font-size:x-small; font-family:Verdana">&nbsp;Site Administrator<br />
            </strong>
        <asp:Panel ID="pnlRoles" runat="server" Width="100%" Visible="True">
            <table style="width: 100%">
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td style="width: 20%">
                    </td>
                    <td style="width: 29%">
                    </td>
                    <td style="width: 20%">
                    </td>
                    <td style="width: 29%" align="right">
                        </td>
                    <td style="width: 1%">
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td>
                        <asp:Label ID="Label1" runat="server" Text="Roles" Font-Names="Verdana" Font-Size="XX-Small" Font-Bold="True"></asp:Label></td>
                    <td>
                    </td>
                    <td>
                    </td>
                    <td align="right">
                        <asp:LinkButton ID="lnkbtnHelpSecurity" runat="server" OnClientClick='window.open("help_security.pdf", "CMHelp","top=10,left=10,width=700,height=450,status=no,toolbar=yes,address=no,menubar=yes,resizable=yes,scrollbars=yes");' Font-Names="Verdana" Font-Size="XX-Small">help on security and passwords</asp:LinkButton></td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label17" runat="server" Text="Account handler:" Font-Names="Verdana" Font-Size="XX-Small" ForeColor="Silver"/></td>
                    <td>
                        <asp:Label ID="lblAccountHandler" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Font-Bold="True" ForeColor="Silver"/></td>
                    <td colspan="2">
                        <asp:Label ID="Label16" runat="server" Text="Site administrator:" Font-Names="Verdana" Font-Size="XX-Small" ForeColor="Silver"/>
                        <asp:Label ID="lblSiteAdministrator" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Font-Bold="True" ForeColor="Silver"/></td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td style="height:4px">
                    </td>
                    <td align="right">
                    </td>
                    <td colspan="2">
                    </td>
                    <td>
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label2" runat="server" Text="Deputy site administrator:" Font-Names="Verdana" Font-Size="XX-Small"/>
                    </td>
                    <td colspan="2">
                        <asp:DropDownList ID="ddlRoleDeputyAdministrator" runat="server" Font-Names="Verdana" Font-Size="XX-Small" OnSelectedIndexChanged="ddlRoleDeputyAdministrator_SelectedIndexChanged" AutoPostBack="True" /><asp:HiddenField ID="hidDeputySiteAdministratorKey" runat="server" />
                    </td>
                    <td>
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label3" runat="server" Text="Site editor:" Font-Names="Verdana" Font-Size="XX-Small"/>
                    </td>
                    <td colspan="2">
                        <asp:DropDownList ID="ddlRoleNoticeBoardEditor" runat="server" Font-Names="Verdana" Font-Size="XX-Small" OnSelectedIndexChanged="ddlRoleNoticeBoardEditor_SelectedIndexChanged" AutoPostBack="True" /><asp:HiddenField ID="hidNoticeBoardEditorKey" runat="server" />
                    </td>
                    <td>
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label4" runat="server" Text="Deputy site editor:" Font-Names="Verdana" Font-Size="XX-Small"/>
                    </td>
                    <td colspan="2"><asp:DropDownList ID="ddlRoleDeputyNoticeBoardEditor" runat="server" Font-Names="Verdana" Font-Size="XX-Small" OnSelectedIndexChanged="ddlRoleDeputyNoticeBoardEditor_SelectedIndexChanged" AutoPostBack="True" /><asp:HiddenField ID="hidDeputyNoticeBoardEditorKey" runat="server" />
                    </td>
                    <td>
                    </td>
                    <td>
                    </td>
                </tr>
            </table>
        </asp:Panel>
        <asp:Panel ID="pnlAuthorisation" runat="server" Width="100%" Visible="False">
            <table style="width: 100%">
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td style="width: 20%">
                        <asp:Label ID="Label24" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="Authorisation Exemption"></asp:Label></td>
                    <td style="width: 29%">
                    </td>
                    <td style="width: 20%">
                    </td>
                    <td style="width: 29%" align="right">
                        </td>
                    <td style="width: 1%">
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td>
                    </td>
                    <td><asp:Label ID="Label25" runat="server" Text="Users:" Font-Names="Verdana" Font-Size="XX-Small"/></td>
                    <td></td>
                    <td align="left">
                    <asp:Label ID="Label26" runat="server" Text="Exempt users:" Font-Names="Verdana" Font-Size="XX-Small"/></td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td valign="top"><asp:Label ID="Label27" runat="server" Text="You can set users to be exempt from authorisation. These users can order any product without requiring authorisation. Typically you set this for authorisers themselves.<br /><br />Click a name then click the arrow pointing to the other box." Font-Names="Verdana" Font-Size="XX-Small" ForeColor="#C04000"/></td>
                    <td>
                        <asp:ListBox ID="lbUsersUnExempt" runat="server" Rows="6" SelectionMode="Multiple" Width="100%" Font-Names="Verdana" Font-Size="XX-Small">
                        </asp:ListBox></td>
                    <td align="center" valign="middle" style="white-space:nowrap">
                        &nbsp;<asp:LinkButton ID="lnkbtnRemoveExemption" runat="server" Font-Names="Verdana" Font-Size="XX-Small" OnClick="lnkbtnRemoveExemption_Click"><<<<< </asp:LinkButton>&nbsp;
                        &nbsp;&nbsp;
                        <asp:LinkButton ID="lnkbtnSetExemption" runat="server" Font-Names="Verdana" Font-Size="XX-Small" OnClick="lnkbtnSetExemption_Click"> >>>>></asp:LinkButton></td>
                    <td align="right">
                        <asp:ListBox ID="lbUsersExempt" runat="server" Rows="6" SelectionMode="Multiple" Width="100%" Font-Names="Verdana" Font-Size="XX-Small">
                        </asp:ListBox></td>
                    <td>
                    </td>
                </tr>
            </table>
        </asp:Panel>
        <asp:Panel ID="pnlSiteFeatures" runat="server" Width="100%" Visible="True">
            <table style="width: 100%">
                <tr>
                    <td style="width: 1%; height: 21px;">
                    </td>
                    <td style="width: 20%; height: 21px;">
                        <asp:Label ID="Label5" runat="server" Text="Site Features" Font-Names="Verdana" Font-Size="XX-Small" Font-Bold="True"></asp:Label></td>
                    <td style="width: 29%; height: 21px;">
                    </td>
                    <td style="width: 20%; height: 21px;">
                    </td>
                    <td style="width: 29%; height: 21px;" align="right">
                        <asp:LinkButton ID="lnkbtnHelpSiteFeatures" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            OnClientClick='window.open("help_sitefeatures.pdf", "CMHelp","top=10,left=10,width=700,height=450,status=no,toolbar=yes,address=no,menubar=yes,resizable=yes,scrollbars=yes");'>help on site features</asp:LinkButton></td>
                    <td style="width: 1%; height: 21px;">
                    </td>
                </tr>
                <tr>
                    <td style="height: 20px">
                    </td>
                    <td align="right" style="height: 20px">
                        <asp:Label ID="Label7" runat="server" Text="Site title:" Font-Names="Verdana" Font-Size="XX-Small"/>
                    </td>
                    <td style="height: 20px">
                        <asp:TextBox ID="tbSiteTitle" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="200px"/>
                    </td>
                    <td style="height: 20px" align="right">
                        <asp:Label ID="Label9" runat="server" Text="Multiple address orders:" Font-Names="Verdana" Font-Size="XX-Small"></asp:Label></td>
                    <td style="height: 20px">
                    <asp:CheckBox ID="cbMultipleAddressOrders" runat="server" Font-Names="Verdana" Font-Size="XX-Small" /></td>
                    <td style="height: 20px">
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label6" runat="server" Text="Show products with zero stock:" Font-Names="Verdana" Font-Size="XX-Small"/>
                    </td>
                    <td>
                        <asp:CheckBox ID="cbShowProductsWithZeroStock" runat="server" Font-Names="Verdana" Font-Size="XX-Small" OnCheckedChanged="cbShowProductsWithZeroStock_CheckedChanged" AutoPostBack="True" />
                    </td>
                    <td align="right">
                        <asp:Label ID="Label8" runat="server" Text="Apply max order amount:" Font-Names="Verdana" Font-Size="XX-Small"/></td>
                    <td>
                        <asp:CheckBox ID="cbApplyMaxGrabs" runat="server" Font-Names="Verdana" Font-Size="XX-Small" /></td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label21" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Zero stock notifications:"/>
                    </td>
                    <td>
                        <asp:CheckBox ID="cbZeroStockNotifications" runat="server" Font-Names="Verdana" Font-Size="XX-Small" AutoPostBack="True" OnCheckedChanged="cbOrderAuthorisation_CheckedChanged" />
                    </td>
                    <td align="right"><asp:Label ID="Label23" runat="server" Text="Show selling price:" Font-Names="Verdana" Font-Size="XX-Small"/></td>
                    <td><asp:CheckBox ID="cbShowSellingPrice" runat="server" Font-Names="Verdana" Font-Size="XX-Small" /></td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right"><asp:Label ID="Label13" runat="server" Text="Users can edit cost centre:" Font-Names="Verdana" Font-Size="XX-Small"/>&nbsp;
                    </td>
                    <td>
                        <asp:CheckBox ID="cbUserCanChangeCostCentre" runat="server" Font-Names="Verdana" Font-Size="XX-Small" />
                    </td>
                    <td align="right"><asp:Label ID="Label14" runat="server" Text="Show notes:" Font-Names="Verdana" Font-Size="XX-Small"/>
                    </td>
                    <td><asp:CheckBox ID="cbShowNotes" runat="server" Font-Names="Verdana" Font-Size="XX-Small" />
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label10" runat="server" Text="Order authorisation:" Font-Names="Verdana" Font-Size="XX-Small"/>
                    </td>
                    <td><asp:CheckBox ID="cbOrderAuthorisation" runat="server" Font-Names="Verdana" Font-Size="XX-Small" AutoPostBack="True" OnCheckedChanged="cbOrderAuthorisation_CheckedChanged" />
                    </td>
                    <td align="right">
                        <asp:Label ID="Label11" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Product authorisation:"/>
                    </td>
                    <td><asp:CheckBox ID="cbProductAuthorisation" runat="server" Font-Names="Verdana" Font-Size="XX-Small" AutoPostBack="True" OnCheckedChanged="cbProductAuthorisation_CheckedChanged" />
                    </td>
                    <td>
                    </td>
                </tr>
                <tr id="trAuthorisationAdvisory" runat="server" visible="false">
                    <td>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label19" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Authorisation advisory:"/>
                    </td>
                    <td colspan="3">
                        <FCKeditorV2:FCKeditor ID="FCKedAuthorisationAdvisory" runat="server" ToolbarSet="CourierSoftware" BasePath="./fckeditor/" Height="100px" />
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right">
                    </td>
                    <td colspan="3">
                        <asp:Label ID="Label20" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            ForeColor="Maroon" Text="The feature settings below are shown for your information. They can only be modified by your Account Handler. Contact your Account Handler for more information."></asp:Label></td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label12" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Calendar management:"/>
                    </td>
                    <td><asp:CheckBox ID="cbCalendarManagement" runat="server" Font-Names="Verdana" Font-Size="XX-Small" OnCheckedChanged="cbProductAuthorisation_CheckedChanged" Enabled="False" /></td>
                    <td align="right">
                        <asp:Label ID="Label15" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Product owners:"/>
                    </td>
                    <td><asp:CheckBox ID="cbProductOwners" runat="server" Font-Names="Verdana" Font-Size="XX-Small" AutoPostBack="True" OnCheckedChanged="cbProductAuthorisation_CheckedChanged" Enabled="False" />
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label18" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="On-demand products:"/>
                    </td>
                    <td>
                        <asp:CheckBox ID="cbOnDemandProducts" runat="server" Enabled="False" Font-Names="Verdana" Font-Size="XX-Small" />
                    </td>
                    <td align="right"><asp:Label ID="Label22" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Custom letters:"/></td>
                    <td><asp:CheckBox ID="cbCustomLetters" runat="server" Font-Names="Verdana" Font-Size="XX-Small" AutoPostBack="True" OnCheckedChanged="cbProductAuthorisation_CheckedChanged" Enabled="False" /></td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label28" runat="server" Text="Advanced user permissions:" Font-Names="Verdana" Font-Size="XX-Small"/></td>
                    <td>
                        <asp:CheckBox ID="cbUserPermissions" runat="server" Font-Names="Verdana" Font-Size="XX-Small" AutoPostBack="True" OnCheckedChanged="cbOrderAuthorisation_CheckedChanged" Enabled="False" /></td>
                    <td align="right"><asp:Label ID="Label29" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Secure file upload:"/></td>
                    <td><asp:CheckBox ID="cbFileUpload" runat="server" Font-Names="Verdana" Font-Size="XX-Small" AutoPostBack="True" OnCheckedChanged="cbProductAuthorisation_CheckedChanged" Enabled="False" /></td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right">
                    </td>
                    <td>
                    </td>
                    <td align="right">
                    </td>
                    <td>
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right">
                    </td>
                    <td>
                        <asp:Button ID="btnSaveSiteFeatureChanges" runat="server" Text="save site feature changes" OnClick="btnSaveSiteFeatureChanges_Click" />
                    </td>
                    <td align="right">
                    </td>
                    <td>
                    </td>
                    <td>
                    </td>
                </tr>
            </table>
        </asp:Panel>
    </form>
</body>
</html>
