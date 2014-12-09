<%@ Page Language="VB" Theme="AIMSDefault" ValidateRequest="false" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<script runat="server">

    Const USER_PERMISSION_A_NOBODY As Integer = 0

    Const USER_PERMISSION_ACCOUNT_HANDLER As Integer = 1
    Const USER_PERMISSION_SITE_ADMINISTRATOR As Integer = 2
    Const USER_PERMISSION_DEPUTY_SITE_ADMINISTRATOR As Integer = 4
    Const USER_PERMISSION_SITE_EDITOR As Integer = 8

    Const USER_PERMISSION_DEPUTY_SITE_EDITOR As Integer = 16
    Const USER_PERMISSION_32_NOT_USED As Integer = 32
    Const USER_PERMISSION_64_NOT_USED As Integer = 64
    Const USER_PERMISSION_128_NOT_USED As Integer = 128
    
    Const USER_PERMISSION_256_NOT_USED As Integer = 256
    Const USER_PERMISSION_512_NOT_USED As Integer = 512
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
    Const USER_PERMISSION_PRODUCT_CREDITS_TAB As Integer = &H800000

    Const USER_PERMISSION_PRODUCT_CREDITS_TAB_CLEAR As Integer = &HF7FFFFF

    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString

    '                 If (CInt(Session("UserPermissions")) And USER_PERMISSION_VIEW_STOCK) > 0 Then
    '                TabView.TabItem = ViewProductsTab()
    '            End If

    Protected Sub Page_Load(ByVal Source As Object, ByVal E As EventArgs)
        'If Not IsNumeric(Session("CustomerKey")) Then
        '    Server.Transfer("session_expired.aspx")
        'End If
        If Not IsPostBack Then
            Call HideAllPanelsAndRows()
            tbUserKey.Focus()
        End If
    End Sub

    Protected Sub SetTitle()
        Dim sTitle As String = Session("SiteTitle")
        If sTitle <> String.Empty Then
            sTitle += " - "
        End If
        Page.Header.Title = sTitle & "Set Permissions"
    End Sub
   
    Protected Sub HideAllPanelsAndRows()
        pnlSpare.Visible = False
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

    Protected Sub UpdatePermissions(nPermissions As Int32)
        Dim sSQL As String = "UPDATE UserProfile SET UserPermissions = " & nPermissions & " WHERE [key] = " & pnUserKey
        Call ExecuteQueryToDataTable(sSQL)
    End Sub
    
    Protected Sub btnSet_Click(sender As Object, e As System.EventArgs)
        Dim nPermissions As Int32
        nPermissions = pnPermissions Or USER_PERMISSION_PRODUCT_CREDITS_TAB
        Call UpdatePermissions(nPermissions)
    End Sub

    Protected Sub btnClear_Click(sender As Object, e As System.EventArgs)
        Dim nPermissions As Int32
        nPermissions = pnPermissions And USER_PERMISSION_PRODUCT_CREDITS_TAB_CLEAR
        Call UpdatePermissions(nPermissions)
    End Sub

    Protected Sub btnGo_Click(sender As Object, e As System.EventArgs)
        If Not IsNumeric(tbUserKey.Text) Then
            lblUserInfo.Text = "User Key is not a valid number."
            Exit Sub
        End If
        Dim dtUser As DataTable = ExecuteQueryToDataTable("SELECT FirstName + ' ' + LastName + ' (' + UserID + ')' 'UserDetails', ISNULL(UserPermissions, 0) 'UserPermissions' FROM UserProfile WHERE [key] = " & tbUserKey.Text)
        If dtUser.Rows.Count = 1 Then
            lblUserInfo.Text = dtUser.Rows(0).Item("UserDetails")
            pnPermissions = dtUser.Rows(0).Item("UserPermissions")
            pnUserKey = CInt(tbUserKey.Text)
        Else
            lblUserInfo.Text = "User not found for this user key."
        End If
    End Sub
    
    Property pnPermissions() As Integer
        Get
            Dim o As Object = ViewState("SP_Permissions")
            If o Is Nothing Then
                Return 0
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Int32)
            ViewState("SP_Permissions") = Value
        End Set
    End Property

    Property pnUserKey() As Int32
        Get
            Dim o As Object = ViewState("SP_UserKey")
            If o Is Nothing Then
                Return 0
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Int32)
            ViewState("SP_UserKey") = Value
        End Set
    End Property

</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Print Status</title>
</head>
<body>
    <form id="form1" runat="server">
    <%--<main:Header ID="ctlHeader" runat="server" />--%>
    <table style="width: 100%">
        <tr>
            <td style="width: 2%">
                &nbsp;
            </td>
            <td style="width: 26%">
                <asp:Label ID="lblLegendTitle" runat="server" Font-Size="Small" Font-Names="Verdana" Font-Bold="True" ForeColor="Gray">Set Product Credits Permission</asp:Label>
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
            <td>
                <asp:Label ID="lblLegendUserID" runat="server" Font-Names="Verdana" Font-Size="XX-Small">User Key:</asp:Label>
                &nbsp; <asp:TextBox ID="tbUserKey" runat="server" Width="197px"></asp:TextBox>
            &nbsp;<asp:Button ID="btnGo" runat="server" onclick="btnGo_Click" Text="go" />
            </td>
            <td>
                <asp:Label ID="lblUserInfo" runat="server" Font-Names="Verdana" 
                    Font-Size="XX-Small"></asp:Label>
            </td>
            <td align="right">
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
            <td>
                &nbsp;
                <asp:Button ID="btnSet" runat="server" Text="set" Width="120px" 
                    onclick="btnSet_Click" />
                &nbsp;<asp:Button ID="btnClear" runat="server" Text="clear" Width="120px" 
                    onclick="btnClear_Click" />
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
    <asp:Panel ID="pnlSpare" runat="server" Visible="false" Width="100%">
        <table style="width: 100%">
            <tr>
                <td style="width: 2%">
                    &nbsp;
                </td>
                <td style="width: 26%">
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
    </asp:Panel>
    </form>
</body>
</html>
