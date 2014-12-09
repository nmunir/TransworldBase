<%@ Page Language="VB" %>
<%@ Register TagPrefix="barcode2" Assembly="Barcode, Version=1.0.5.40001, Culture=neutral, PublicKeyToken=6dc438ab78a525b3" Namespace="Lesnikowski.Web" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>

<script runat="server">

    Dim lCollectionKey As Long
    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString()
    
    Protected Sub Page_Load(ByVal Source As Object, ByVal E As EventArgs)
        If Not IsPostBack Then
            lCollectionKey = Request.QueryString("Key")
            lblCourierBookingNumber.Text = lCollectionKey
            barcode.Number = lCollectionKey.ToString
            GetDriversName()
            BindConsignments()
        End If
    End Sub
    
    Protected Sub GetDriversName()
        Dim oDataReader As SqlDataReader
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As New SqlCommand("spASPNET_CourierBooking_GetDriversName", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Dim oParam As SqlParameter = oCmd.Parameters.Add("@CourierBookingKey", SqlDbType.Int, 4)
        oParam.Value = lCollectionKey
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            oDataReader.Read()
            If oDataReader.HasRows Then
                
                If Not IsDBNull(oDataReader("Company")) Then
                    lblCompany.Text = oDataReader("Company")
                End If
                If Not IsDBNull(oDataReader("Addr1")) Then
                    lblAddr1.Text = oDataReader("Addr1")
                End If
                If Not IsDBNull(oDataReader("Town")) Then
                    lblTown.Text = oDataReader("Town")
                End If
                If Not IsDBNull(oDataReader("PostCode")) Then
                    lblPostCode.Text = oDataReader("PostCode")
                End If
                If Not IsDBNull(oDataReader("ReadyAt")) Then
                    lblReadyAt.Text = Format(oDataReader("ReadyAt"), "HH:mm  dd MMMM yyyy ")
                End If
                If Not IsDBNull(oDataReader("DriversName")) Then
                    lblDriversName.Text = oDataReader("DriversName")
                End If
            End If
            lblDateTime.Text = Format(Now, "dd MMMM yyyy HH:mm")
        Catch ex As SqlException
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub BindConsignments()
        If lCollectionKey > 0 Then
            Dim oConn As New SqlConnection(gsConn)
            Dim oDataTable As New DataTable()
            Dim oAdapter As New SqlDataAdapter("spASPNET_CourierBooking_GetConsignments",oConn)
            oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CourierBookingKey", SqlDbType.Int))
            oAdapter.SelectCommand.Parameters("@CourierBookingKey").Value = lCollectionKey
            Try
                oAdapter.Fill(oDataTable)
                If oDataTable.Rows.Count > 0 Then
                    lblTotalConsignments.Text = oDataTable.Rows.Count.ToString
                    Dim nTotalItems As Integer = 0
                    For Each dr As DataRow In oDataTable.Rows
                        nTotalItems += CInt(dr("NOP") & String.Empty)
                    Next
                    lblTotalItems.Text = nTotalItems.ToString
                    dgConsignments.DataSource = oDataTable
                    dgConsignments.DataBind()
                    dgConsignments.Visible = True
                Else
                    dgConsignments.Visible = False
                End If
            Catch ex As SqlException
                lblError.Text = ex.ToString
            Finally
                oConn.Close()
            End Try
        End If
    End Sub

</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Driver's Manifest</title>
</head>
<body>
    <form id="frmDriversManifest" runat="server">
        <table style="font-size:xx-small; font-family:Verdana; width:620px">
            <tr>
                <td style="white-space:nowrap; width:420px" valign="middle">
                    <asp:Label ID="lbl001" runat="server" font-size="Small" font-names="Verdana" font-bold="True">Transworld</asp:Label>
                    <br />
                    <asp:Label ID="lbl002" runat="server" font-size="XX-Small" font-names="Verdana">London Heathrow. TEL: 44 (0)208 751 1111<br /> WEB: www.transworld.eu.com</asp:Label>
                    <br />
                    <asp:Label ID="lbl003" runat="server" font-size="Large" font-bold="True">DRIVER'S MANIFEST</asp:Label> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:Label runat="server" id="lblDate" font-size="X-Small" font-bold="True" font-names="Verdana"></asp:Label>
                </td>
                <td style="white-space:nowrap; width:200px" valign="top">
                    <barcode2:BarcodeControl ID="barcode" Symbology="Code128" XDpi="300" YDpi="300" NarrowBarWidth="2" Height="70" IsNumberVisible="false" runat="server"/>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:Label ID="lbl004" runat="server" font-size="XX-Small" font-names="Verdana">Use this manifest as a receipt when our driver collects from you.</asp:Label>
                    <br />
                    <asp:Label ID="lbl005" runat="server" font-size="XX-Small" font-names="Verdana">Please print an extra copy for the driver to take away - thank you.</asp:Label>
                    <br />
                </td>
                <td style="white-space:nowrap">
                    <asp:Label runat="server" id="lblCourierBookingNumber" font-size="Large" font-names="Courier New" font-bold="True"></asp:Label>
                </td>
            </tr>
            <tr>
                <td></td>
                <td>
                    <asp:Label runat="server" id="lblCompany"></asp:Label>
                </td>
            </tr>
            <tr>
                <td></td>
                <td>
                    <asp:Label runat="server" id="lblAddr1"></asp:Label>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:Label ID="lbl006" runat="server">Booked for:</asp:Label>
                    &nbsp;<asp:Label runat="server" id="lblReadyAt"></asp:Label>
                </td>
                <td>
                    <asp:Label runat="server" id="lblTown"></asp:Label>
                    &nbsp;<asp:Label runat="server" id="lblPostCode"></asp:Label>
                </td>
            </tr>
            <tr>
                <td></td>
                <td></td>
            </tr>
        </table>

        <asp:DataGrid id="dgConsignments" runat="server" Font-Size="X-Small" Font-Names="Verdana" Width="620px" Visible="False" AutoGenerateColumns="False" GridLines="None" ShowFooter="True">
            <FooterStyle wrap="False"></FooterStyle>
            <HeaderStyle font-names="Verdana" wrap="False" forecolor="Black" backcolor="Silver"></HeaderStyle>
            <PagerStyle font-size="X-Small" font-names="Verdana" horizontalalign="Center" forecolor="Blue" backcolor="Silver" wrap="False" mode="NumericPages"></PagerStyle>
            <Columns>
                <asp:BoundColumn Visible="False" DataField="Key">
                    <ItemStyle wrap="False"></ItemStyle>
                </asp:BoundColumn>
                <asp:TemplateColumn>
                    <ItemStyle horizontalalign="Left"></ItemStyle>
                    <ItemTemplate>
                        <asp:Table id="tabConsignmentList" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="620px">
                            <asp:TableRow>
                                <asp:TableCell RowSpan="2" width="90px" VerticalAlign="Top" Wrap="false">
                                    <asp:Label runat="server" font-bold="True" font-size="Small"><%# DataBinder.Eval(Container.DataItem,"Key") %></asp:Label>
                                </asp:TableCell>
                                <asp:TableCell width="50px" VerticalAlign="Top" HorizontalAlign="Right">
                                    <asp:Label runat="server">Consignee:</asp:Label>
                                </asp:TableCell>
                                <asp:TableCell width="190px" VerticalAlign="Top">
                                    <asp:Label runat="server" font-bold="True"><%# DataBinder.Eval(Container.DataItem,"CneeName") %></asp:Label>
                                </asp:TableCell>
                                <asp:TableCell width="60px" VerticalAlign="Top" HorizontalAlign="Right">
                                    <asp:Label runat="server">City:</asp:Label>
                                </asp:TableCell>
                                <asp:TableCell width="190px" VerticalAlign="Top">
                                    <asp:Label runat="server" font-bold="True"><%# DataBinder.Eval(Container.DataItem,"CneeTown") %></asp:Label>
                                </asp:TableCell>
                            </asp:TableRow>
                            <asp:TableRow>
                                <asp:TableCell VerticalAlign="Top" HorizontalAlign="Right">
                                    <asp:Label runat="server">Ref:</asp:Label>
                                </asp:TableCell>
                                <asp:TableCell VerticalAlign="Top">
                                    <asp:Label runat="server" font-bold="True"><%# DataBinder.Eval(Container.DataItem,"CustomerRef1") %></asp:Label>
                                </asp:TableCell>
                                <asp:TableCell VerticalAlign="Top" HorizontalAlign="Right">
                                    <asp:Label runat="server">Country:</asp:Label>
                                </asp:TableCell>
                                <asp:TableCell VerticalAlign="Top" ColumnSpan="3">
                                    <asp:Label runat="server" font-bold="True"><%# DataBinder.Eval(Container.DataItem,"CountryName") %></asp:Label>
                                </asp:TableCell>
                            </asp:TableRow>
                            <asp:TableRow>
                                <asp:TableCell VerticalAlign="Top"></asp:TableCell>
                                <asp:TableCell VerticalAlign="Top" HorizontalAlign="Right">
                                    <asp:Label runat="server">NOP:</asp:Label>
                                </asp:TableCell>
                                <asp:TableCell VerticalAlign="Top">
                                    <asp:Label runat="server" font-bold="True"><%# DataBinder.Eval(Container.DataItem,"NOP") %></asp:Label>
                                </asp:TableCell>
                                <asp:TableCell VerticalAlign="Top" HorizontalAlign="Right">
                                    <asp:Label runat="server">Weight:</asp:Label>
                                </asp:TableCell>
                                <asp:TableCell VerticalAlign="Top" ColumnSpan="3">
                                    <asp:Label runat="server" font-bold="True"><%# DataBinder.Eval(Container.DataItem,"Weight") %></asp:Label>
                                </asp:TableCell>
                            </asp:TableRow>
                            <asp:TableRow>
                                <asp:TableCell ColumnSpan="8" VerticalAlign="Top">
                                    <hr />
                                </asp:TableCell>
                            </asp:TableRow>
                        </asp:Table>
                    </ItemTemplate>
                </asp:TemplateColumn>
            </Columns>
        </asp:DataGrid>
        <asp:Label ID="Label1" runat="server" Font-Names="Verdana" Font-Size="X-Small" Text="(Total consignments: "></asp:Label>
        <asp:Label ID="lblTotalConsignments" runat="server" Font-Bold="True" Font-Names="Verdana"
            Font-Size="X-Small"></asp:Label>
        <asp:Label ID="Label3" runat="server" Font-Names="Verdana" Font-Size="X-Small" Text=" Total items: "></asp:Label>
        <asp:Label ID="lblTotalItems" runat="server" Font-Bold="True" Font-Names="Verdana"
            Font-Size="X-Small"></asp:Label>
        <asp:Label ID="Label5" runat="server" Font-Names="Verdana" Font-Size="X-Small" Text=")"></asp:Label><br />
        <br />

        <table id="tabPlainPaperAWB2" style="font-size:xx-small; font-family:Verdana; width:630px">
            <tr>
                <td style="width:150px">
                    <asp:Label ID="lbl007" runat="server" font-size="X-Small" font-names="Verdana">Driver's Name :</asp:Label>
                    <br />
                    <br />
                </td>
                <td>
                    <asp:Label runat="server" ID="lblDriversName" font-size="X-Small" font-names="Verdana"></asp:Label>
                    <br />
                    <br />
                </td>
            </tr>
            <tr>
                <td style="width:150px">
                    <asp:Label ID="lbl009" runat="server" font-size="X-Small" font-names="Verdana">Driver's Signature:</asp:Label>
                    <br />
                    <br />
                </td>
                <td></td>
            </tr>
            <tr>
                <td style="width:150px">
                    <asp:Label ID="lbl010" runat="server" font-size="X-Small" font-names="Verdana">Date / Time :</asp:Label>
                </td>
                <td>
                    <asp:Label runat="server" id="lblDateTime" font-size="X-Small"></asp:Label>
                    <br />
                </td>
            </tr>
        </table>
        <asp:Label id="lblError" runat="server" forecolor="#00C000" font-names="Verdana" font-size="X-Small"></asp:Label>
    </form>
</body>
</html>
