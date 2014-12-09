<%@ Page Language="VB" %>
<%@ Register TagPrefix="barcode2" Assembly="Barcode, Version=1.0.5.40001, Culture=neutral, PublicKeyToken=6dc438ab78a525b3" Namespace="Lesnikowski.Web" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>

<script runat="server">

    Const USER_PERMISSION_ACCOUNT_HANDLER As Integer = 1
    Const USER_PERMISSION_SITE_ADMINISTRATOR As Integer = 2
    Const USER_PERMISSION_DEPUTY_SITE_ADMINISTRATOR As Integer = 4
    Const USER_PERMISSION_SITE_EDITOR As Integer = 8
    Const USER_PERMISSION_DEPUTY_SITE_EDITOR As Integer = 16

    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString()

    Protected Sub Page_Load(ByVal Source As Object, ByVal E As EventArgs)
        If Not IsNumeric(Session("CustomerKey")) Then
            Server.Transfer("session_expired.aspx")
            Exit Sub
        End If
        If Not IsPostBack Then
            lblCustomerName.Text = Session("CustomerName")
            lblCustomerKey.Text = Session("CustomerKey")
            Call GetRoles()
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
                        lblAccountHandler.Text = lblAccountHandler.Text.Replace("kamran.ejaz@westernunion.com", "- - - -")
                    End If
                    If CInt(oSqlDataReader("UserPermissions")) And USER_PERMISSION_SITE_ADMINISTRATOR Then
                        lblSiteAdministrator.Text = oSqlDataReader("FirstName") & " " & oSqlDataReader("LastName") & " (" & oSqlDataReader("UserId") & ")"
                        lblSiteAdministrator.Text = lblSiteAdministrator.Text.Replace("kamran.ejaz@westernunion.com", "- - - -")
                    End If
                    If CInt(oSqlDataReader("UserPermissions")) And USER_PERMISSION_DEPUTY_SITE_ADMINISTRATOR Then
                        lblDeputySiteAdministrator.Text = oSqlDataReader("FirstName") & " " & oSqlDataReader("LastName") & " (" & oSqlDataReader("UserId") & ")"
                        lblDeputySiteAdministrator.Text = lblDeputySiteAdministrator.Text.Replace("kamran.ejaz@westernunion.com", "- - - -")
                    End If
                    If CInt(oSqlDataReader("UserPermissions")) And USER_PERMISSION_SITE_EDITOR Then
                        lblSiteEditor.Text = oSqlDataReader("FirstName") & " " & oSqlDataReader("LastName") & " (" & oSqlDataReader("UserId") & ")"
                        lblSiteEditor.Text = lblSiteEditor.Text.Replace("kamran.ejaz@westernunion.com", "- - - -")
                    End If
                    If CInt(oSqlDataReader("UserPermissions")) And USER_PERMISSION_DEPUTY_SITE_EDITOR Then
                        lblDeputySiteEditor.Text = oSqlDataReader("FirstName") & " " & oSqlDataReader("LastName") & " (" & oSqlDataReader("UserId") & ")"
                        lblDeputySiteEditor.Text = lblDeputySiteEditor.Text.Replace("kamran.ejaz@westernunion.com", "- - - -")
                    End If
                End While
            End If
        Catch ex As Exception
            WebMsgBox.Show("GetRoles: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
    
</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>
Site Roles
</title>
    <style type="text/css">
        .style1
        {
            width: 251px;
        }
        .style2
        {
            height: 14px;
            width: 251px;
        }
    </style>
</head>
<body>
    <form id="frmConsignmentNote" runat="server">
        <table style="width:620px; font-family:Verdana; font-size:xx-small">
            <tr>
                <td valign="top" style="white-space:nowrap; " colspan="2">
                    <asp:Label ID="lbl001" runat="server" font-size="Large" font-bold="True">Roles</asp:Label><br />
                    <asp:Label ID="lblCustomerName" runat="server"></asp:Label>&nbsp;(<asp:Label 
                        ID="lblCustomerKey" runat="server"></asp:Label>)&nbsp;
                </td>
            </tr>
            <tr>
                <td align="right" class="style1">
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td align="right" class="style1">
                    Account Handler:</td>
                <td>
                    <asp:Label ID="lblAccountHandler" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Font-Bold="True">(unassigned)</asp:Label></td>
            </tr>
            <tr>
                <td align="right" class="style2">
                    Site Administrator:</td>
                <td>
                    <asp:Label ID="lblSiteAdministrator" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Font-Bold="True">(unassigned)</asp:Label></td>
            </tr>
            <tr>
                <td align="right" class="style1">
                    Deputy Site Administrator:</td>
                <td>
                    <asp:Label ID="lblDeputySiteAdministrator" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Font-Bold="True">(unassigned)</asp:Label></td>
            </tr>
            <tr>
                <td align="right" class="style1">
                    Site Editor:</td>
                <td>
                    <asp:Label ID="lblSiteEditor" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Font-Bold="True">(unassigned)</asp:Label></td>
            </tr>
            <tr>
                <td align="right" class="style2">
                    Deputy Site Editor:</td>
                <td>
                    <asp:Label ID="lblDeputySiteEditor" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Font-Bold="True">(unassigned)</asp:Label></td>
            </tr>
            <tr>
                <td class="style1">
                </td>
                <td>
                    <br />
                </td>
            </tr>
            <tr>
                <td class="style1">
                </td>
                <td>
                    <asp:Button ID="btnCloseWindow" runat="server" Text="close window" OnClientClick="window.close()" /></td>
            </tr>
        </table>
        &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
    </form>
</body>
</html>