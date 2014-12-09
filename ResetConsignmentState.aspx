<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsPostBack Then
            Call DisplayConsignmentsInEmailErrorState()
        End If
    End Sub
    
    Protected Sub DisplayConsignmentsInEmailErrorState()
        Dim sSQL As String = "SELECT [key] FROM Consignment WHERE StateId = 'EMAIL_ERROR' AND CreatedOn >= '" & Date.Today.AddDays(-ddlDays.SelectedValue).ToString("dd-MMM-yyyy") & "'"
        Dim oDataTable As DataTable = ExecuteQueryToDataTable(sSQL)
        gvConsignments.DataSource = oDataTable
        gvConsignments.DataBind()
    End Sub
    
    Protected Sub btnReset_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim b As Button = sender
        Dim sSQL As String = "UPDATE Consignment SET StateId = 'RECEIVED' WHERE StateId = 'EMAIL_ERROR' AND [key] = " & b.CommandArgument
        If Not ExecuteNonQuery(sSQL) Then
            WebMsgBox.Show("Error resetting status of consignment " & b.CommandArgument)
        Else
            Call DisplayConsignmentsInEmailErrorState()
        End If
    End Sub
    
    Protected Sub gvConsignments_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs)
        Exit Sub
        Dim gvr As GridViewRow = e.Row
        If gvr.RowType = DataControlRowType.DataRow Then
            Dim sConsignmentKey As String = gvr.Cells(1).Text
            Dim b As Button = gvr.FindControl("btnReset")
            b.CommandArgument = sConsignmentKey
        End If
    End Sub

    Protected Sub ddlDays_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call DisplayConsignmentsInEmailErrorState()
    End Sub

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
            WebMsgBox.Show("Error in ExecuteQueryToListItemCollection: " & ex.Message)
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

</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Reset Consignment State</title>
    <style type="text/css">
       BODY {
        font-family: Verdana;
        font-size: xx-small
       }     
    </style>
</head>
<body>
    <form id="form1" runat="server">
    <asp:Label ID="Label4" runat="server" Text="This utility lets you reset back to RECEIVED a consignment that has been set to EMAIL_ERROR. Typically this happens when the consignment contains Chinese or similar characters, which some parts of the system cannot currently handle."></asp:Label>
    <br />
    <br />
    <asp:Label ID="Label5" runat="server" 
        Text="Just click 'reset consignment' for the consignment you want to reset."></asp:Label>
    <br />
    <br />
    <asp:GridView ID="gvConsignments" runat="server" Width="60%" 
        OnRowDataBound="gvConsignments_RowDataBound" CellPadding="5" 
        Font-Names="Verdana" Font-Size="XX-Small">
        <Columns>
            <asp:TemplateField>
                <ItemTemplate>
                    <asp:Button ID="btnReset" runat="server" Text="reset consignment" onclick="btnReset_Click" CommandArgument='<%# DataBinder.Eval(Container, "DataItem.Key") %>' />
                </ItemTemplate>
            </asp:TemplateField>
        </Columns>
        <EmptyDataTemplate>
            &nbsp;
            <asp:Label ID="Label1" runat="server" Text="no consignment found in state EMAIL_ERROR within the date range specified" ForeColor="Red" style="font-weight: 700"/>
        </EmptyDataTemplate>
    </asp:GridView>
    <br />
    <asp:Label ID="Label2" runat="server" Text="Days to go back:"></asp:Label>
    <asp:DropDownList ID="ddlDays" runat="server" AutoPostBack="True" 
        onselectedindexchanged="ddlDays_SelectedIndexChanged" style="height: 22px" 
        Font-Names="Verdana" Font-Size="XX-Small">
        <asp:ListItem Selected="True">1</asp:ListItem>
        <asp:ListItem>2</asp:ListItem>
        <asp:ListItem>5</asp:ListItem>
        <asp:ListItem>20</asp:ListItem>
        <asp:ListItem>100</asp:ListItem>
    </asp:DropDownList>
    &nbsp;<asp:Label ID="Label3" runat="server" 
        Text="(choose a higher number to look back further in the list of consignments)"></asp:Label>
    </form>
</body>
</html>
