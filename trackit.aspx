<%@ Page Language="VB" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>
<script runat="server">

    ' trackit.aspx
    
    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString

    Sub Page_Load(ByVal Source As Object, ByVal E As EventArgs)
        If Not IsPostBack Then
            psAWB = Request.QueryString("AWB")
            lblAWB.Text = psAWB
            GetConsignmentFromAWB()
        End If
    End Sub
    
    Sub GetConsignmentFromAWB()
        If psAWB <> "" Then
            ResetForm()
            lblError.Text = ""
            Dim sCneeName As String = String.Empty
            Dim sCneeAddr1 As String = String.Empty
            Dim sCneeAddr2 As String = String.Empty
            Dim sCneeAddr3 As String = String.Empty
            Dim sCneeTown As String = String.Empty
            Dim sCneeState As String = String.Empty
            Dim sCneePostCode As String = String.Empty
            Dim sCneeCountry As String = String.Empty
            Dim sCneeContact As String = String.Empty
            Dim sCneeAdress As String = String.Empty
            Dim oDataReader As SqlDataReader
            Dim oConn As New SqlConnection(gsConn)
            Dim oCmd As New SqlCommand("spASPNET_Tracking_GetConsignmentFromAWB", oConn)
            oCmd.CommandType = CommandType.StoredProcedure
            Dim oParam As New SqlParameter("@AWB", SqlDbType.NVarChar, 50)
            oCmd.Parameters.Add(oParam)
            oParam.Value = psAWB
            Try
                oConn.Open()
                oDataReader = oCmd.ExecuteReader()
                oDataReader.Read()
                plConsignmentKey = oDataReader("Key")
                If Not IsDBNull(oDataReader("CreatedOn")) Then
                    lblDate.Text = Format(oDataReader("CreatedOn"), "dd MMM yyyy HH:mm")
                End If
    
                If Not IsDBNull(oDataReader("CneeName")) Then
                    sCneeName = oDataReader("CneeName")
                End If
                If Not IsDBNull(oDataReader("CneeAddr1")) Then
                    sCneeAddr1 = oDataReader("CneeAddr1")
                End If
                If Not IsDBNull(oDataReader("CneeAddr2")) Then
                    sCneeAddr2 &= "  " & oDataReader("CneeAddr2")
                End If
                If Not IsDBNull(oDataReader("CneeAddr3")) Then
                    sCneeAddr3 &= "  " & oDataReader("CneeAddr3")
                End If
                If Not IsDBNull(oDataReader("CneeTown")) Then
                    sCneeTown = oDataReader("CneeTown")
                End If
                If Not IsDBNull(oDataReader("CneeState")) Then
                    sCneeState = oDataReader("CneeState")
                End If
                If Not IsDBNull(oDataReader("CneePostCode")) Then
                    sCneePostCode = oDataReader("CneePostCode")
                End If
                If Not IsDBNull(oDataReader("CneeCountryName")) Then
                    sCneeCountry = oDataReader("CneeCountryName")
                End If
                If Not IsDBNull(oDataReader("CneeCtcName")) Then
                    sCneeContact = oDataReader("CneeCtcName")
                End If
                If Not IsDBNull(oDataReader("CneeTel")) Then
                    sCneeContact &= "  " & oDataReader("CneeTel")
                End If
                If Not IsDBNull(oDataReader("Weight")) Then
                    If oDataReader("Weight") <> "0" Then
                        lblWeight.Text = oDataReader("Weight")
                    End If
                End If
                If Not IsDBNull(oDataReader("NOP")) Then
                    If oDataReader("NOP") <> "0" Then
                        lblNOP.Text = oDataReader("NOP")
                    End If
                End If
                If Not IsDBNull(oDataReader("SpecialInstructions")) Then
                    lblSpclInstructions.Text = oDataReader("SpecialInstructions")
                End If
                If Not IsDBNull(oDataReader("Description")) Then
                    lblContents.Text = oDataReader("Description")
                End If
                If Not IsDBNull(oDataReader("CustomerRef1")) Then
                    lblCustRef1.Text = oDataReader("CustomerRef1")
                End If
                If Not IsDBNull(oDataReader("CustomerRef2")) Then
                    lblCustRef2.Text = oDataReader("CustomerRef2")
                End If
                If Not IsDBNull(oDataReader("PODDate")) Then
                    lblPODDate.Text = oDataReader("PODDate")
                End If
                If Not IsDBNull(oDataReader("PODName")) Then
                    lblPODName.Text = oDataReader("PODName")
                End If
                If Not IsDBNull(oDataReader("PODTime")) Then
                    lblPODTime.Text = oDataReader("PODTime")
                End If
                oDataReader.Close()
            Catch ex As SqlException
                lblError.Text = ex.ToString
            Finally
                oConn.Close()
            End Try
    
            sCneeAdress = sCneeName & "</br>"
            sCneeAdress &= sCneeAddr1 & "</br>"
            If sCneeAddr2 <> "" Then
                sCneeAdress &= sCneeAddr2 & "</br>"
            End If
            If sCneeAddr3 <> "" Then
                sCneeAdress &= sCneeAddr3 & "</br>"
            End If
            sCneeAdress &= sCneeTown & "</br>"
            If sCneeState <> "" Then
                sCneeAdress &= sCneeState & "</br>"
            End If
            If sCneePostCode <> "" Then
                sCneeAdress &= sCneePostCode & "</br>"
            End If
            sCneeAdress &= sCneeCountry & "</br>"
            sCneeAdress &= sCneeContact & "</br>"
    
            lblCneeAddr.Text = sCneeAdress
    
            GetTracking()
        End If
    End Sub
    
    Sub GetTracking()
        If psAWB <> "" Then
            lblError.Text = ""
            Dim oConn As New SqlConnection(gsConn)
            Dim oDataSet As New DataSet()
            Dim oAdapter As New SqlDataAdapter("spASPNET_Consignment_GetTracking", oConn)
            Try
                oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ConsignmentKey", SqlDbType.Int))
                oAdapter.SelectCommand.Parameters("@ConsignmentKey").Value = plConsignmentKey
                oAdapter.Fill(oDataSet, "Tracking")
                Dim Source As DataView = oDataSet.Tables("Tracking").DefaultView
                If Source.Count > 0 Then
                    grid_Tracking.DataSource = Source
                    grid_Tracking.DataBind()
                    grid_Tracking.Visible= True
                Else
                    grid_Tracking.Visible= False
                End If
            Catch ex As SqlException
                lblError.Text = ex.ToString
            Finally
                oConn.Close()
            End Try
        End If
    End Sub
    
    Sub ResetForm()
        lblDate.Text = ""
        lblCneeAddr.Text = ""
        lblWeight.Text = ""
        lblNOP.Text = ""
        lblSpclInstructions.Text = ""
        lblContents.Text = ""
        lblCustRef1.Text = ""
        lblCustRef2.Text = ""
        lblPODDate.Text = ""
        lblPODName.Text = ""
        lblPODTime.Text = ""
    End Sub

    Property psAWB() As String
        Get
            Dim o As Object = ViewState("AWB")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("AWB") = Value
        End Set
    End Property
    
    
    Property plConsignmentKey() As Long
        Get
            Dim o As Object = ViewState("ConsignmentKey")
            If o Is Nothing Then
                Return -1
            End If
            Return CLng(o)
        End Get
        Set(ByVal Value As Long)
            ViewState("ConsignmentKey") = Value
        End Set
    End Property
    
</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Transworld Consignment Tracking</title>
</head>
<body>
    <form id="Form1" method="post" enctype="multipart/form-data" runat="server">
        <asp:Table id="tblFrame" runat="server" Width="100%">
            <asp:TableRow>
                <asp:TableCell HorizontalAlign="Center">
                    <asp:Table id="tblHeading1" runat="server" Width="640px" font-names="Verdana" forecolor="Gray" font-size="XX-Small">
                        <asp:TableRow>
                            <asp:TableCell Width="30%" Wrap="False" VerticalAlign="Top">
                                <asp:Image runat="server" ImageUrl="http://my.transworld.eu.com/common/images/logos/transworld.jpg"></asp:Image>
                            </asp:TableCell>
                            <asp:TableCell Width="70%" Wrap="False" VerticalAlign="Top" ColumnSpan="2">
                                <asp:Label runat="server" font-size="X-Small">Transworld</asp:Label>
                                <br />
                                <asp:Label runat="server">The Mercury Centre, Central Way, Feltham</asp:Label>
                                <br />
                                <asp:Label runat="server">TW14 0RN  United Kingdom.</asp:Label>
                                <br />
                                <asp:Label runat="server">Tel: 44 (0)208 751 7501 Fax: 44 (0)208 890 9090</asp:Label>
                                <br />
                                <br />
                            </asp:TableCell>
                        </asp:TableRow>
                    </asp:Table>
                    <asp:Table id="tblHeading2" runat="server" Width="640px" font-names="Verdana" forecolor="Gray" font-size="XX-Small">
                        <asp:TableRow>
                            <asp:TableCell Wrap="False" VerticalAlign="Bottom" ColumnSpan="2">
                                <br />
                                <asp:Label runat="server" font-size="Small" font-bold="True">Consignment Tracking</asp:Label>
                            </asp:TableCell>
                        </asp:TableRow>
                        <asp:TableRow>
                            <asp:TableCell Width="140px" Wrap="False">
                                <asp:Label runat="server" font-size="X-Small">Consignment No</asp:Label>
                            </asp:TableCell>
                            <asp:TableCell Width="500px" Wrap="False">
                                <asp:Label runat="server" id="lblAWB" font-size="X-Small" forecolor="Red"></asp:Label>
                            </asp:TableCell>
                        </asp:TableRow>
                        <asp:TableRow>
                            <asp:TableCell Width="140px" Wrap="False">
                                <asp:Label runat="server" font-size="X-Small">Booked On</asp:Label>
                            </asp:TableCell>
                            <asp:TableCell Width="500px" Wrap="False">
                                <asp:Label runat="server" id="lblDate" font-size="X-Small" forecolor="Red"></asp:Label>
                            </asp:TableCell>
                        </asp:TableRow>
                    </asp:Table>
                    <br />
                    <asp:Table id="tabConsignmentDetail1" runat="server" Width="640px" Font-Names="Verdana" Font-Size="XX-Small" forecolor="Navy">
                        <asp:TableRow>
                            <asp:TableCell Width="140px"></asp:TableCell>
                            <asp:TableCell Width="500px"></asp:TableCell>
                        </asp:TableRow>
                        <asp:TableRow>
                            <asp:TableCell valign="top">
                                <asp:Label runat="server" forecolor="Gray">Consignee:</asp:Label>
                            </asp:TableCell>
                            <asp:TableCell valign="top">
                                <asp:Label runat="server" id="lblCneeAddr"></asp:Label>
                            </asp:TableCell>
                        </asp:TableRow>
                    </asp:Table>
                    <asp:Table id="tabConsignmentDetail2" runat="server" Width="640px" Font-Names="Verdana" Font-Size="XX-Small" forecolor="Navy">
                        <asp:TableRow>
                            <asp:TableCell Width="140px"></asp:TableCell>
                            <asp:TableCell Width="180px"></asp:TableCell>
                            <asp:TableCell Width="140px"></asp:TableCell>
                            <asp:TableCell Width="180px"></asp:TableCell>
                        </asp:TableRow>
                        <asp:TableRow>
                            <asp:TableCell>
                                <asp:Label runat="server" forecolor="Gray">NOP:</asp:Label>
                            </asp:TableCell>
                            <asp:TableCell>
                                <asp:Label runat="server" id="lblNOP" ></asp:Label>
                            </asp:TableCell>
                            <asp:TableCell>
                                <asp:Label runat="server" forecolor="Gray">Weight:</asp:Label>
                            </asp:TableCell>
                            <asp:TableCell>
                                <asp:Label runat="server" id="lblWeight" ></asp:Label>
                            </asp:TableCell>
                        </asp:TableRow>
                        <asp:TableRow>
                            <asp:TableCell>
                                <asp:Label runat="server" forecolor="Gray">Ref 1:</asp:Label>
                            </asp:TableCell>
                            <asp:TableCell>
                                <asp:Label runat="server" id="lblCustRef1"></asp:Label>
                            </asp:TableCell>
                            <asp:TableCell>
                                <asp:Label runat="server" forecolor="Gray">Ref 2:</asp:Label>
                            </asp:TableCell>
                            <asp:TableCell>
                                <asp:Label runat="server" id="lblCustRef2"></asp:Label>
                            </asp:TableCell>
                        </asp:TableRow>
                        <asp:TableRow>
                            <asp:TableCell>
                                <asp:Label runat="server" forecolor="Gray">Contents:</asp:Label>
                            </asp:TableCell>
                            <asp:TableCell ColumnSpan="3">
                                <asp:Label runat="server" id="lblContents"></asp:Label>
                            </asp:TableCell>
                        </asp:TableRow>
                        <asp:TableRow>
                            <asp:TableCell>
                                <asp:Label runat="server" forecolor="Gray">Spcl Instr:</asp:Label>
                            </asp:TableCell>
                            <asp:TableCell ColumnSpan="3">
                                <asp:Label runat="server" id="lblSpclInstructions"></asp:Label>
                            </asp:TableCell>
                        </asp:TableRow>
                        <asp:TableRow>
                            <asp:TableCell>
                                <asp:Label runat="server" forecolor="Gray">Received By:</asp:Label>
                            </asp:TableCell>
                            <asp:TableCell ColumnSpan="3">
                                <asp:Label runat="server" id="lblPODDate" forecolor="Red"></asp:Label> &nbsp;<asp:Label runat="server" id="lblPODName" forecolor="Red"></asp:Label> &nbsp;<asp:Label runat="server" id="lblPODTime" forecolor="Red"></asp:Label>
                            </asp:TableCell>
                        </asp:TableRow>
                    </asp:Table>
                    <br />
                    <asp:DataGrid id="grid_Tracking" runat="server" Width="640px" Font-Names="Verdana" Font-Size="XX-Small" ShowFooter="False" GridLines="None" AutoGenerateColumns="False" Visible="False">
                        <FooterStyle wrap="False"></FooterStyle>
                        <HeaderStyle font-names="Verdana" wrap="False"></HeaderStyle>
                        <PagerStyle font-size="X-Small" font-names="Verdana" font-bold="True" horizontalalign="Center" forecolor="Blue" backcolor="Silver" wrap="False" mode="NumericPages"></PagerStyle>
                        <Columns>
                            <asp:BoundColumn DataField="Time" HeaderText="Time" DataFormatString="{0:dd.MM.yy HH:mm}">
                                <HeaderStyle wrap="False" forecolor="Gray" width="20%"></HeaderStyle>
                                <ItemStyle wrap="False" forecolor="Navy" verticalalign="Top"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="Location" HeaderText="Location">
                                <HeaderStyle wrap="False" horizontalalign="Left" forecolor="Gray" width="20%"></HeaderStyle>
                                <ItemStyle wrap="False" forecolor="Navy" verticalalign="Top"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="Description" HeaderText="Description">
                                <HeaderStyle wrap="False" forecolor="Gray" width="60%"></HeaderStyle>
                                <ItemStyle forecolor="Navy" verticalalign="Top"></ItemStyle>
                            </asp:BoundColumn>
                        </Columns>
                    </asp:DataGrid>
                    <asp:Table id="tblFooter" runat="server" Width="640px" font-names="Verdana" forecolor="Gray" font-size="XX-Small">
                        <asp:TableRow>
                            <asp:TableCell Wrap="False">
                                <br />
                                <asp:HyperLink id="HyperLink1" runat="server" Font-Names="Verdana" Font-Size="XX-Small" ForeColor="Blue" NavigateUrl="http://www.transworld.eu.com">www.transworld.eu.com</asp:HyperLink>
                            </asp:TableCell>
                        </asp:TableRow>
                    </asp:Table>
                </asp:TableCell>
            </asp:TableRow>
        </asp:Table>
        <asp:Label id="lblError" runat="server" forecolor="#00C000" font-size="X-Small" font-names="Verdana"></asp:Label>
    </form>
</body>
</html>
