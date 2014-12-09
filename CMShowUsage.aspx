<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>

<script runat="server">

    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString()
    Private gsProductKey As String

    Protected Sub Page_Load(ByVal Source As Object, ByVal E As EventArgs)
        If Not IsNumeric(Session("CustomerKey")) Then
            Server.Transfer("session_expired.aspx")
            Exit Sub
        End If
        If Not IsPostBack Then
            Try
                gsProductKey = Request.QueryString("ref")
            Catch ex As Exception
                WebMsgBox.Show("No ref supplied!")
                Exit Sub
            End Try
            If Not IsNumeric(gsProductKey) Then
                WebMsgBox.Show("Not a valid ref!")
                Exit Sub
            End If
            Dim sSQL As String = "SELECT * FROM LogisticProduct WHERE LogisticProductKey = " & gsProductKey
            Dim oDataTable As DataTable = ExecuteQueryToDataTable(sSQL)
            If oDataTable.Rows.Count > 0 Then
                Dim dr As DataRow = oDataTable.Rows(0)
                Dim bCalendarManaged As Boolean
                If Not IsDBNull(dr("CalendarManaged")) Then
                    bCalendarManaged = dr("CalendarManaged")
                Else
                    bCalendarManaged = False
                End If
                If bCalendarManaged Then
                    If dr("ArchiveFlag") = "Y" Or dr("DeletedFlag") = "Y" Then
                        WebMsgBox.Show("Product is archived or deleted!")
                        Exit Sub
                    Else
                        lblProductDetails.Text = dr("ProductCode") & " - " & dr("ProductDescription")
                        sSQL = "SELECT EventDay, EventName FROM CalendarManagedItemDays cmid INNER JOIN CalendarManagedItemEvent cmie ON cmid.EventId = cmie.[id] WHERE EventDay >= GETDATE() AND ISNULL(IsDeleted,0) = 0 AND LogisticProductKey = " & gsProductKey & " ORDER BY EventDay"
                        Dim oDataTable2 As DataTable = ExecuteQueryToDataTable(sSQL)
                        gvProductUsage.DataSource = oDataTable2
                        gvProductUsage.DataBind()
                    End If
                Else
                    WebMsgBox.Show("Not a calendar managed product!")
                    Exit Sub
                End If
            Else
                WebMsgBox.Show("Invalid ref!")
                Exit Sub
            End If
        End If
        Call SetTitle()
    End Sub

    Protected Sub SetTitle()
        Dim sTitle As String = Session("SiteTitle")
        If sTitle <> String.Empty Then
            sTitle += " - "
        End If
        Page.Header.Title = sTitle & "Product Calendar"
    End Sub
   
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
            WebMsgBox.Show("Error in ExecuteQueryToListItemCollection executing: " & sQuery & " : " & ex.Message)
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

</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Product Calendar</title>
</head>
<body>
    <form runat="server">
    <table style="width: 100%">
        <tr>
            <td style="width: 80%">
                <asp:Label ID="Label4" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="X-Small" Text="Product Calendar for: " />
                <asp:Label ID="lblProductDetails" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="X-Small" />
            </td>
            <td style="width: 20%" align="right">
                  <asp:LinkButton ID="lnkbtnCloseWindow" runat="server" OnClientClick="window.close()" Font-Size="XX-Small" CausesValidation="False">close window</asp:LinkButton>            
            </td>
        </tr>
    </table>
    <table style="width: 100%">
        <tr>
            <td style="width: 5%">
            </td>
            <td style="width: 95%">
            </td>
        </tr>
        <tr>
            <td>
            </td>
            <td>
                <asp:GridView ID="gvProductUsage" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="100%" CellPadding="2" AutoGenerateColumns="False">
                    <Columns>
                        <asp:TemplateField HeaderText="Date" SortExpression="EventDay">
                            <ItemTemplate>
                                <asp:Label ID="Label1" runat="server" Text='<%# Format(DataBinder.Eval(Container.DataItem,"EventDay"),"ddd d-MMM-yyyy") %>'></asp:Label>
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Event" SortExpression="EventName">
                            <ItemTemplate>
                                <asp:Label ID="Label2" runat="server" Text='<%# Bind("EventName") %>'></asp:Label>
                            </ItemTemplate>
                        </asp:TemplateField>
                    </Columns>
                    <EmptyDataTemplate>
                        no usage data found
                    </EmptyDataTemplate>
                </asp:GridView>
            </td>
        </tr>
    </table>
    </form>
</body>
</html>
