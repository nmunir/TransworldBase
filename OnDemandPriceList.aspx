<%@ Page Language="VB" Theme="AIMSDefault" %>

<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>

<script runat="server">

    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString()
    Dim nPriceListKey As Integer
    Dim nProductKey As Integer
    Dim nTariffMaxQuantity As Integer = 0
    Dim nLastQuantityRead As Integer = 0
    Dim nTariffItemCount As Integer = 0
    
    Protected Sub Page_Load(ByVal Source As Object, ByVal E As EventArgs)
        If Not IsNumeric(Session("CustomerKey")) Then
            Server.Transfer("session_expired.aspx")
        End If
        If Not IsPostBack Then
            nPriceListKey = Request.QueryString("Key")
            nProductKey = Request.QueryString("Product")
            Call GetProductDetails()
            Call GetTariff()
        End If
    End Sub
    
    Protected Sub GetProductDetails()
        If nProductKey > 0 Then
            Dim oDataReader As SqlDataReader
            Dim oConn As New SqlConnection(gsConn)
            Dim sSQL As String = "SELECT ProductCode, ProductDate, ProductDescription FROM LogisticProduct WHERE CustomerKey = " & Session("CustomerKey") & " AND ISNULL(OnDemand,0) > 0 AND LogisticProductKey = " & nProductKey
            Dim oCmd As New SqlCommand(sSQL, oConn)
            Try
                oConn.Open()
                oDataReader = oCmd.ExecuteReader()
                If oDataReader.HasRows Then
                    oDataReader.Read()
                    lblProduct.Text = oDataReader("ProductCode")
                    Dim sProductDate As String = oDataReader("ProductDate").ToString.Trim
                    If sProductDate <> String.Empty Then
                        lblProduct.Text += " - " & sProductDate
                    End If
                    Dim sProductDescription As String = oDataReader("ProductDescription").ToString.Trim
                    If sProductDescription <> String.Empty Then
                        lblProduct.Text += " - " & sProductDescription
                    End If
                Else
                    lblProduct.Text = "Product details not available"
                End If
            Catch ex As SqlException
                WebMsgBox.Show("Error in GetProductDetails: " & ex.Message)
            Finally
                oConn.Close()
            End Try
        Else
            lblProduct.Text = "Product details not available"
        End If
    End Sub
    
    Protected Sub GetTariff()
        If nPriceListKey > 0 Then
            Dim sSQL As String = "SELECT Quantity, Price FROM OnDemandTariff WHERE TariffId = " & nPriceListKey & " ORDER BY Quantity"
            Dim oListItemCollection As ListItemCollection = ExecuteQueryToListItemCollection(sSQL, "Quantity", "Price")
            If oListItemCollection.Count > 0 Then
                nTariffItemCount = oListItemCollection.Count
                nTariffMaxQuantity = CInt(oListItemCollection.Item(nTariffItemCount - 1).Text)
                gvPODTariff.DataSource = oListItemCollection
                gvPODTariff.DataBind()
            Else
                lblPriceList.Text = "Price list empty or not found"
            End If
        Else
            lblPriceList.Text = "No price list specified"
        End If
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

    Protected Sub gvPODTariff_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs)
        Dim gvr As GridViewRow = e.Row
        If gvr.RowType = DataControlRowType.DataRow Then
            Dim nQuantity As Integer = CInt(gvr.Cells(0).Text)
            If nQuantity = nTariffMaxQuantity Then
                If nTariffItemCount = 1 Then
                    gvr.Cells(0).Text = "(any quantity)"
                Else
                    'gvr.Cells(0).Text = nLastQuantityRead + 1 & " to " & gvr.Cells(0).Text
                    'gvr.Cells(0).Text = gvr.Cells(0).Text & " or more"
                    gvr.Cells(0).Text = nLastQuantityRead + 1 & " or more"
                End If
            Else
                gvr.Cells(0).Text = nLastQuantityRead + 1 & " to " & gvr.Cells(0).Text
            End If
            nLastQuantityRead = nQuantity
        End If
    End Sub
    
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>Price List </title>
</head>
<body>
    <form id="frmOnDemandPriceList" runat="server">
        <table style="width: 100%">
            <tr>
                <td style="width: 90%; white-space: nowrap">
                    &nbsp;<asp:Label ID="Label1" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="X-Small" Text="Price list for product "/>
                    <asp:Label ID="lblProduct" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="X-Small" ForeColor="Blue"/>
                </td>
                <td style="width: 10%" align="right">
                    <asp:Button ID="btnCloseWindow" OnClientClick="window.close()" runat="server" Text="close window" />
                </td>
            </tr>
            <tr>
                <td colspan="2" align="center">
                    <asp:Label ID="lblPriceList" runat="server" Font-Names="Verdana" Font-Size="XX-Small"/>
                    <br />
                    <asp:GridView ID="gvPODTariff" runat="server" Font-Names="Verdana" Font-Size="XX-Small" AutoGenerateColumns="False" Width="95%" 
                        OnRowDataBound="gvPODTariff_RowDataBound" CellPadding="2">
                        <Columns>
                            <asp:BoundField DataField="text" HeaderText="Quantity" />
                            <asp:TemplateField HeaderText="Price per unit">
                                <ItemTemplate>
                                    <asp:Label ID="Label1" runat="server" Text='<%# FormatCurrency(DataBinder.Eval(Container.DataItem,"value")) %>'/>
                                </ItemTemplate>
                            </asp:TemplateField>
                        </Columns>
                    </asp:GridView>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
