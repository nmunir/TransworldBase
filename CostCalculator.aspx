<%@ Page Language="VB" MasterPageFile="~/MasterPage.master" Title="AIMS Store and Despatch" Theme="SkinFile" %>
<script runat="server">
    ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '   Copyright AIMS Store & Despatch
    ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    Property lCustomerKey() As Long
        Get
            Dim o As Object = ViewState("CustomerKey")
            If o Is Nothing Then
                Return 0
            End If
            Return CLng(o)
        End Get
        Set(ByVal Value As Long)
            ViewState("CustomerKey") = Value
        End Set
    End Property
    
    Property lServiceLevelKey() As Long
        Get
            Dim o As Object = ViewState("ServiceLevelKey")
            If o Is Nothing Then
                Return 0
            End If
            Return CLng(o)
        End Get
        Set(ByVal Value As Long)
            ViewState("ServiceLevelKey") = Value
        End Set
    End Property
    
    Property lCountryKey() As Long
        Get
            Dim o As Object = ViewState("CountryKey")
            If o Is Nothing Then
                Return 0
            End If
            Return CLng(o)
        End Get
        Set(ByVal Value As Long)
            ViewState("CountryKey") = Value
        End Set
    End Property 'Page Properties
    
    Public Class CostEstimate
        Public WeightCharge As Double
        Public EstimatedPackagingWeight As Double
        Public NonDoCSurCharge As Double
        Public DiscountRate As Double
        Public LocalTaxRate As Double
    End Class
    
    
    Public Class CostCalculator
    
        Public Function GetCostEstimate(ByVal lCustomerKey As Long, _
                                        ByVal lServiceLevelKey As Long, _
                                        ByVal sDocumentFlag As String, _
                                        ByVal sEstimatePackagingFlag As String, _
                                        ByVal lCountryKey As Long, _
                                        ByVal sTown As String, _
                                        ByVal sPostCode As String, _
                                        ByVal dblWeight As Double) As CostEstimate
    
    
            Dim dblWeightCharge As Double
            Dim dblMatrixBandFee As Double
            Dim bIsBaseRate As Boolean = True
            Dim dblBaseRate As Double
            Dim dblRemainder As Double
            Dim dblProductWeight As Double = dblWeight
            Dim dblPackagingWeight As Double = 0.0
    
            ' Create CustomerDetails Struct
            Dim oCostEstimate As CostEstimate = New CostEstimate()
    
            Dim dr As DataRow
            Dim sConn As String = ConfigurationManager.ConnectionStrings("LogisticsConnectionString").ConnectionString
            Dim oConn As New SqlConnection(sConn)
            Dim oDataSet As New DataSet()
            Dim oAdapter As New SqlDataAdapter("spASPNET_Tariff_GetZoneMatrixFromAddress", oConn)
            oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
            Try
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
                oAdapter.SelectCommand.Parameters("@CustomerKey").Value = lCustomerKey
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ServiceLevelKey", SqlDbType.Int))
                oAdapter.SelectCommand.Parameters("@ServiceLevelKey").Value = lServiceLevelKey
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@DocumentFlag", SqlDbType.NVarChar, 1))
                oAdapter.SelectCommand.Parameters("@DocumentFlag").Value = sDocumentFlag
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CountryKey", SqlDbType.Int))
                oAdapter.SelectCommand.Parameters("@CountryKey").Value = lCountryKey
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@Town", SqlDbType.NVarChar, 50))
                oAdapter.SelectCommand.Parameters("@Town").Value = sTown
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@PostalCode", SqlDbType.NVarChar, 50))
                oAdapter.SelectCommand.Parameters("@PostalCode").Value = sPostCode
    
                oAdapter.Fill(oDataSet, "ZoneMatrix")
    
                If sEstimatePackagingFlag = "Y" Then
                    Do While dblProductWeight > 0
                        'Add 230 grams (for packaging) per 12.5 kilos of product
                        dblPackagingWeight = dblPackagingWeight + 0.23
                        dblProductWeight = dblProductWeight - 12.5
                    Loop
                    'Now add packaging weight onto the consignment Weight
                    dblWeight = dblWeight + dblPackagingWeight
                End If
    
                For Each dr In oDataSet.Tables("ZoneMatrix").Rows
                    'First make a record of the Base Charge
                    If bIsBaseRate Then
                        dblBaseRate = ((dr("WeightTo") - dr("WeightFrom")) / dr("Units")) * dr("Fee")
                        bIsBaseRate = False
                    End If
                    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ Iterating through the zone matrix ~~~~~~~~~~~~~~~~~~~~~~~~
                    'If parcel weight is heavier than this row's from and to delimiters then work out how much
                    'this band charges out at, add it to the running total [dblWeightCharge] and go to next row.
                    If dblWeight >= dr("WeightTo") Then
                        'Normally the weight charge is calculated like your tax and you work out what each
                        'portion of the weight is charged at. This model is broken when the Tariff is a flat rate
                        'tariff. Flat Rate Tariffs just multiply the unit charge by the parcel weight divided by
                        'the units
                        If dr("FlatRate") = False Then  ' not a flat rate
                            dblMatrixBandFee = ((dr("WeightTo") - dr("WeightFrom")) / dr("Units")) * dr("Fee")
                            dblWeightCharge = dblWeightCharge + dblMatrixBandFee
                        Else                       ' this is now (possible already was) a flat rate charge
                            'This matrix row is marked as Flat Rate which means we now apply the this rows's unit
                            'charge to all units. Before disregarding everything we must also see if the 'Hold Base'
                            'flag is set. If it is we must still charge the first rows unit charge and only apply
                            'flat rate for units thereafter.
                            If dr("HoldBase") = True Then
                                dblWeightCharge = ((dblWeight / dr("Units")) * dr("Fee")) + dblBaseRate
                            Else
                                dblWeightCharge = (dblWeight / dr("Units")) * dr("Fee")
                            End If
                        End If
                    ElseIf dblWeight >= dr("WeightFrom") And dblWeight < dr("WeightTo") Then 'Stop here: weight lies between this row's from and to                         
                        'ElseIf parcel weight lies between this row's from and to delimiters then this is the last
                        'row we need look at in this Zone Matrix. Calculate the weight charge and add it to running total.
                        If dr("FlatRate") = False Then  ' not a flat rate
                            dblRemainder = (dblWeight - dr("WeightFrom")) / dr("Units")
                            Do While dblRemainder > 0
                                dblWeightCharge = dblWeightCharge + dr("Fee")
                                dblRemainder = dblRemainder - 1
                            Loop
                        Else
                            If dr("HoldBase") = True Then  'see above for explanation
                                dblWeightCharge = ((dblWeight / dr("Units")) * dr("Fee")) + dblBaseRate
                            Else
                                dblWeightCharge = (dblWeight / dr("Units")) * dr("Fee")
                            End If
                        End If
                    ElseIf dblWeight >= dr("WeightTo") Then  'This weight exceeds the highest weight band
                        'We have calculated all of the tariff bands so now just loop up to the weight using
                        ' the last matrix fee
                        
                    End If
                    'Following variables returned in each row of recordset - all the same
                    'When I find out how to return more than one recordset from a stored procedure
                    'and then iterate through them, I'll change this code.
                    oCostEstimate.NonDoCSurCharge = CDbl(dr("NonDocSurcharge"))
                    oCostEstimate.DiscountRate = CDbl(dr("DiscountRate"))
                    oCostEstimate.LocalTaxRate = CDbl(dr("LocalTaxRate"))
                Next
    
                oCostEstimate.WeightCharge = dblWeightCharge
                oCostEstimate.EstimatedPackagingWeight = dblPackagingWeight
    
            Catch ex As SqlException
            Finally
                oConn.Close()
            End Try
    
            Return oCostEstimate
    
        End Function
    
    End Class 'Classes
    
    Sub Page_Load(ByVal Source As Object, ByVal E As EventArgs)
    
        'catch session expiry
        If Not IsNumeric(Session("UserKey")) Then
            Server.Transfer("session_expired.aspx")
        End If
    
        If Not IsPostBack Then
            chkDocuments.Checked = True
            chkPackaging.Checked = False
            GetCustomersWithTariffs()
            GetCustomerServiceLevels()
            GetCountries()
            txtWeight.Text = "0.500"
        End If
    End Sub
    
    Sub drop_Customer_Changed(ByVal s As Object, ByVal e As EventArgs)
        lCustomerKey = CLng(drop_Customer.SelectedItem.Value)
        GetCustomerServiceLevels()
    End Sub
    
    Sub drop_ServiceLevel_Changed(ByVal s As Object, ByVal e As EventArgs)
        lServiceLevelKey = CLng(drop_ServiceLevel.SelectedItem.Value)
    End Sub
    
    Sub drop_Country_Changed(ByVal s As Object, ByVal e As EventArgs)
        lCountryKey = CLng(drop_Country.SelectedItem.Value)
    End Sub
    
    Sub btn_Calculate_Click(ByVal s As Object, ByVal e As ImageClickEventArgs)
    
        'catch session expiry
        If Not IsNumeric(Session("UserKey")) Then
            Server.Transfer("session_expired.aspx")
        End If
    
        lblError.Text = ""
    
        lblTotalWeight.Text = "0.000"
        lblWeightCharge.Text = "0.00"
        lblDiscountRate.text = "0.00"
        lblDiscountAmount.Text = "0.00"
        lblDiscountedCharge.Text = "0.00"
        lblNDS.text = "0.00"
        lblSubTotal.text = "0.00"
    
        CalculateCost()
    
    End Sub
    
    Sub CalculateCost()
        Dim sDocumentFlag As String
        Dim sEstimatePackagingFlag As String
        Dim oCostCalculator As CostCalculator = New CostCalculator()
        Dim oCostEstimate As CostEstimate = New CostEstimate()
    
        Dim dblProductWeight As Double
        Dim dblTotalWeight As Double
        Dim dblWeightCharge As Double
        Dim dblDiscountRate As Double
        Dim dblDiscountAmount As Double
        Dim dblDiscountedCharge As Double
        Dim dblNDS As Double
        Dim dblSubTotal As Double
        Dim dblLocalTaxRate As Double
        Dim dblLocalTaxAmount As Double
        Dim dblTotal As Double
    
    
        If lCustomerKey > 0 And lServiceLevelKey > 0 And lCountryKey > 0 Then
            If IsNumeric(txtWeight.Text) Then
    
                dblProductWeight = CDbl(txtWeight.Text)
    
                If chkDocuments.Checked Then
                    sDocumentFlag = "Y"
                Else
                    sDocumentFlag = "N"
                End If
    
                If chkPackaging.Checked Then
                    sEstimatePackagingFlag = "Y"
                Else
                    sEstimatePackagingFlag = "N"
                End If
    
                Try
                    oCostEstimate = oCostCalculator.GetCostEstimate(lCustomerKey, _
                                                                        lServiceLevelKey, _
                                                                        sDocumentFlag, _
                                                                        sEstimatePackagingFlag, _
                                                                        lCountryKey, _
                                                                        txtCity.Text, _
                                                                        txtPostCode.Text, _
                                                                        Format(dblProductWeight, "#,##0.000"))
    
                    dblTotalWeight = oCostEstimate.EstimatedPackagingWeight + CDbl(txtWeight.Text)
                    dblWeightCharge = oCostEstimate.WeightCharge
                    dblDiscountRate = oCostEstimate.DiscountRate
                    dblDiscountAmount = (oCostEstimate.WeightCharge * oCostEstimate.DiscountRate) / 100
                    dblDiscountedCharge = dblWeightCharge + dblDiscountAmount
                    dblNDS = oCostEstimate.NonDoCSurCharge
                    dblSubTotal = dblDiscountedCharge + dblNDS
                    dblLocalTaxRate = Format(oCostEstimate.LocalTaxRate, "#,##0.00")
                    dblLocalTaxAmount = (dblSubTotal * dblLocalTaxRate) / 100
                    dblTotal = dblSubTotal + dblLocalTaxAmount
    
                    txtWeight.Text = Format(dblProductWeight, "#,##0.000")
                    lblTotalWeight.Text = Format(dblTotalWeight, "#,##0.000")
                    lblWeightCharge.Text = Format(dblWeightCharge, "#,##0.00")
                    lblDiscountRate.text = Format(dblDiscountRate, "#,##0.00")
                    lblDiscountAmount.Text = Format(dblDiscountAmount, "#,##0.00")
                    lblDiscountedCharge.Text = Format(dblDiscountedCharge, "#,##0.00")
                    lblNDS.Text = Format(dblNDS, "#,##0.00")
                    lblSubTotal.Text = Format(dblSubTotal, "#,##0.00")
                    lblLocalTaxRate.Text = Format(oCostEstimate.LocalTaxRate, "#,##0.00")
                    lblLocalTaxAmount.Text = Format(dblLocalTaxAmount, "#,##0.00")
                    lblTotal.Text = Format(dblTotal, "#,##0.00")
    
                Catch ex As Exception
                    lblError.Text = ex.ToString
                End Try
            Else
                lblError.Text = "Error: invalid weight."
            End If
        Else
            lblError.Text = "lCustomerKey = [" & lCustomerKey & "], lServiceLevelKey = [" & lServiceLevelKey & "], lCountryKey = [" & lCountryKey & "]"
        End If
    End Sub
    
    Sub GetCustomersWithTariffs()
        Dim sConn As String = ConfigurationSettings.AppSettings("ConnectionString")
        Dim oConn As New SqlConnection(sConn)
        Dim oCmd As New SQLCommand("spASPNET_Customer_GetAccountCodesWithTariffs", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Try
            oConn.Open()
            drop_Customer.DataSource = oCmd.ExecuteReader()
            drop_Customer.DataTextField = "CustomerAccountCode"
            drop_Customer.DataValueField = "CustomerKey"
            drop_Customer.DataBind()
        Catch ex As SQLException
            lblError.Text = ""
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Sub GetCustomerServiceLevels()
        Dim sConn As String = ConfigurationSettings.AppSettings("ConnectionString")
        Dim oConn As New SqlConnection(sConn)
        Dim oCmd As New SQLCommand("spStockMngr_TariffAssignment_GetServiceLevelsForCustomer", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Dim oParam As SqlParameter = oCmd.Parameters.Add("@CustomerKey", SqlDbType.Int, 4)
        oParam.Value = lCustomerKey
        Try
            oConn.Open()
            drop_ServiceLevel.DataSource = oCmd.ExecuteReader()
            drop_ServiceLevel.DataTextField = "ServiceLevel"
            drop_ServiceLevel.DataValueField = "ServiceLevelKey"
            drop_ServiceLevel.DataBind()
        Catch ex As SQLException
            lblError.Text = ""
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
    End Sub
    
    
    Sub GetCountries()
        Dim sConn As String = ConfigurationSettings.AppSettings("ConnectionString")
        Dim oConn As New SqlConnection(sConn)
        Dim oCmd As New SQLCommand("spASPNET_Country_GetCountries", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Try
            oConn.Open()
            drop_Country.DataSource = oCmd.ExecuteReader()
            drop_Country.DataTextField = "CountryName"
            drop_Country.DataValueField = "CountryKey"
            drop_Country.DataBind()
        Catch ex As SQLException
            lblError.Text = ex.ToString
        End Try
        oConn.Close()
    End Sub 'Page Events
    
</script>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <asp:Table id="TabHeader" runat="server" Width="750px">
        <asp:TableRow>
            <asp:TableCell Width="750px" HorizontalAlign="Center">
                <asp:Image ID="Image1" runat="server" ImageUrl="images/AIMS_Banner.jpg"></asp:Image>
            </asp:TableCell>
        </asp:TableRow>
    </asp:Table>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder2" Runat="Server">
    <asp:Table id="TabData" runat="server" Width="750px" font-size="XX-Small">
        <asp:TableRow>
            <asp:TableCell BackColor="Silver" ColumnSpan="4" HorizontalAlign="Center">
                <asp:Label ID="Label3" runat="server" font-size="X-Small" forecolor="White" font-bold="True">Cost Calculator</asp:Label>
            </asp:TableCell>
        </asp:TableRow>
        <asp:TableRow>
            <asp:TableCell Width="200px">
                <br />
            </asp:TableCell>
            <asp:TableCell Width="200px"></asp:TableCell>
            <asp:TableCell Width="150px"></asp:TableCell>
            <asp:TableCell Width="200px"></asp:TableCell>
        </asp:TableRow>
        <asp:TableRow>
            <asp:TableCell HorizontalAlign="Left">
                <asp:Label ID="Label4" runat="server">Customer: </asp:Label>
            </asp:TableCell>
            <asp:TableCell HorizontalAlign="Left">
                <asp:DropDownList runat="server" AutoPostBack="True" OnSelectedIndexChanged="drop_Customer_Changed" id="drop_Customer"></asp:DropDownList>
            </asp:TableCell>
            <asp:TableCell HorizontalAlign="Left">
                <asp:Label ID="Label5" runat="server">Service Level: </asp:Label>
            </asp:TableCell>
            <asp:TableCell HorizontalAlign="Left">
                <asp:DropDownList AutoPostBack="True" OnSelectedIndexChanged="drop_ServiceLevel_Changed" id="drop_ServiceLevel" runat="server"></asp:DropDownList>
            </asp:TableCell>
        </asp:TableRow>
        <asp:TableRow>
            <asp:TableCell HorizontalAlign="Left">
                <asp:Label ID="Label6" runat="server">Documents: </asp:Label>
            </asp:TableCell>
            <asp:TableCell HorizontalAlign="Left">
                <asp:CheckBox id="chkDocuments" runat="server"></asp:CheckBox>
            </asp:TableCell>
            <asp:TableCell HorizontalAlign="Left">
                <asp:Label ID="Label7" runat="server">Country: </asp:Label>
            </asp:TableCell>
            <asp:TableCell HorizontalAlign="Left">
                <asp:DropDownList AutoPostBack="True" OnSelectedIndexChanged="drop_Country_Changed" id="drop_Country" runat="server"></asp:DropDownList>
            </asp:TableCell>
        </asp:TableRow>
        <asp:TableRow>
            <asp:TableCell HorizontalAlign="Left">
                <asp:Label ID="Label8" runat="server">Estimate Pakaging: </asp:Label>
            </asp:TableCell>
            <asp:TableCell HorizontalAlign="Left">
                <asp:CheckBox id="chkPackaging" runat="server"></asp:CheckBox>
            </asp:TableCell>
            <asp:TableCell HorizontalAlign="Left">
                <asp:Label ID="Label9" runat="server">City (optional): </asp:Label>
            </asp:TableCell>
            <asp:TableCell HorizontalAlign="Left">
                <asp:TextBox runat="server" ID="txtCity"></asp:TextBox>
            </asp:TableCell>
        </asp:TableRow>
        <asp:TableRow>
            <asp:TableCell HorizontalAlign="Left"></asp:TableCell>
            <asp:TableCell HorizontalAlign="Left"></asp:TableCell>
            <asp:TableCell HorizontalAlign="Left">
                <asp:Label ID="Label10" runat="server">Post Code (optional): </asp:Label>
            </asp:TableCell>
            <asp:TableCell HorizontalAlign="Left">
                <asp:TextBox runat="server" ID="txtPostCode"></asp:TextBox>
            </asp:TableCell>
        </asp:TableRow>
        <asp:TableRow>
            <asp:TableCell HorizontalAlign="Left"></asp:TableCell>
            <asp:TableCell HorizontalAlign="Left"></asp:TableCell>
            <asp:TableCell HorizontalAlign="Left">
                <asp:Label ID="Label11" runat="server">Weight (kilos): </asp:Label>
            </asp:TableCell>
            <asp:TableCell HorizontalAlign="Left">
                <asp:TextBox runat="server" ID="txtWeight"></asp:TextBox>
            </asp:TableCell>
        </asp:TableRow>
        <asp:TableRow>
            <asp:TableCell HorizontalAlign="Left"></asp:TableCell>
            <asp:TableCell HorizontalAlign="Left"></asp:TableCell>
            <asp:TableCell HorizontalAlign="Left"></asp:TableCell>
            <asp:TableCell HorizontalAlign="Left">
                <asp:ImageButton id="btn_Calculate" onclick="btn_Calculate_Click" runat="server" ImageUrl="images/btn_submit.gif"></asp:ImageButton>
            </asp:TableCell>
        </asp:TableRow>
    </asp:Table>
    <asp:Table id="TabResults" runat="server" Font-Names="Verdana" ForeColor="Navy" Width="750px" font-size="XX-Small">
        <asp:TableRow>
            <asp:TableCell BackColor="Silver" ColumnSpan="4" HorizontalAlign="Center">
                <asp:Label ID="Label12" runat="server" font-size="X-Small" forecolor="White" font-bold="True">Tariff Results</asp:Label>
            </asp:TableCell>
        </asp:TableRow>
        <asp:TableRow>
            <asp:TableCell Width="250px">
                <br />
            </asp:TableCell>
            <asp:TableCell Width="150px"></asp:TableCell>
            <asp:TableCell Width="150px"></asp:TableCell>
            <asp:TableCell Width="200px"></asp:TableCell>
        </asp:TableRow>
        <asp:TableRow>
            <asp:TableCell HorizontalAlign="Right">
                <asp:Label ID="Label13" runat="server">Weight including packaging (kgs): </asp:Label>
            </asp:TableCell>
            <asp:TableCell HorizontalAlign="Right">
                <asp:Label runat="server" id="lblTotalWeight">0.000</asp:Label>
            </asp:TableCell>
            <asp:TableCell HorizontalAlign="Right">
                <asp:Label ID="Label14" runat="server">Weight Charge: </asp:Label>
            </asp:TableCell>
            <asp:TableCell HorizontalAlign="Right">
                <asp:Label runat="server" id="lblWeightCharge">0.00</asp:Label>
            </asp:TableCell>
        </asp:TableRow>
        <asp:TableRow>
            <asp:TableCell HorizontalAlign="Right">
                <asp:Label ID="Label15" runat="server">Plus/Minus %: </asp:Label>
            </asp:TableCell>
            <asp:TableCell HorizontalAlign="Right">
                <asp:Label runat="server" id="lblDiscountRate">0.00</asp:Label>
            </asp:TableCell>
            <asp:TableCell HorizontalAlign="Right">
                <asp:Label ID="Label16" runat="server">Plus/Minus Amount: </asp:Label>
            </asp:TableCell>
            <asp:TableCell HorizontalAlign="Right">
                <asp:Label runat="server" id="lblDiscountAmount">0.00</asp:Label>
            </asp:TableCell>
        </asp:TableRow>
        <asp:TableRow>
            <asp:TableCell></asp:TableCell>
            <asp:TableCell></asp:TableCell>
            <asp:TableCell></asp:TableCell>
            <asp:TableCell>
                <hr />
            </asp:TableCell>
        </asp:TableRow>
        <asp:TableRow>
            <asp:TableCell></asp:TableCell>
            <asp:TableCell></asp:TableCell>
            <asp:TableCell></asp:TableCell>
            <asp:TableCell HorizontalAlign="Right">
                <asp:Label runat="server" id="lblDiscountedCharge">0.00</asp:Label>
            </asp:TableCell>
        </asp:TableRow>
        <asp:TableRow>
            <asp:TableCell></asp:TableCell>
            <asp:TableCell></asp:TableCell>
            <asp:TableCell HorizontalAlign="Right">
                <asp:Label ID="Label17" runat="server">Non Doc Surcharge: </asp:Label>
            </asp:TableCell>
            <asp:TableCell HorizontalAlign="Right">
                <asp:Label runat="server" id="lblNDS">0.00</asp:Label>
            </asp:TableCell>
        </asp:TableRow>
        <asp:TableRow>
            <asp:TableCell></asp:TableCell>
            <asp:TableCell></asp:TableCell>
            <asp:TableCell></asp:TableCell>
            <asp:TableCell>
                <hr />
            </asp:TableCell>
        </asp:TableRow>
        <asp:TableRow>
            <asp:TableCell></asp:TableCell>
            <asp:TableCell></asp:TableCell>
            <asp:TableCell HorizontalAlign="Right">
                <asp:Label ID="Label18" runat="server">Sub Total: </asp:Label>
            </asp:TableCell>
            <asp:TableCell HorizontalAlign="Right">
                <asp:Label runat="server" id="lblSubTotal">0.00</asp:Label>
            </asp:TableCell>
        </asp:TableRow>
        <asp:TableRow>
            <asp:TableCell HorizontalAlign="Right">
                <asp:Label ID="Label19" runat="server">Local Tax Rate: </asp:Label>
            </asp:TableCell>
            <asp:TableCell HorizontalAlign="Right">
                <asp:Label runat="server" id="lblLocalTaxRate">0.00</asp:Label>
            </asp:TableCell>
            <asp:TableCell HorizontalAlign="Right">
                <asp:Label ID="Label20" runat="server">Local Tax Amount: </asp:Label>
            </asp:TableCell>
            <asp:TableCell HorizontalAlign="Right">
                <asp:Label runat="server" id="lblLocalTaxAmount">0.00</asp:Label>
            </asp:TableCell>
        </asp:TableRow>
        <asp:TableRow>
            <asp:TableCell></asp:TableCell>
            <asp:TableCell></asp:TableCell>
            <asp:TableCell></asp:TableCell>
            <asp:TableCell>
                <hr />
            </asp:TableCell>
        </asp:TableRow>
        <asp:TableRow>
            <asp:TableCell></asp:TableCell>
            <asp:TableCell></asp:TableCell>
            <asp:TableCell HorizontalAlign="Right">
                <asp:Label ID="Label21" runat="server">Total: </asp:Label>
            </asp:TableCell>
            <asp:TableCell HorizontalAlign="Right" BackColor="#FFFFC0">
                <asp:Label runat="server" id="lblTotal">0.00</asp:Label>
            </asp:TableCell>
        </asp:TableRow>
        <asp:TableRow>
            <asp:TableCell>
                <br />
            </asp:TableCell>
            <asp:TableCell></asp:TableCell>
            <asp:TableCell></asp:TableCell>
            <asp:TableCell></asp:TableCell>
        </asp:TableRow>
    </asp:Table>
</asp:Content>
<asp:Content ID="Content3" ContentPlaceHolderID="ContentPlaceHolder3" Runat="Server">
    <asp:Table id="TabFooter" runat="server" Font-Names="Verdana" ForeColor="Navy" Width="750px" font-size="XX-Small">
        <asp:TableRow>
            <asp:TableCell HorizontalAlign="Center">
                <asp:Image ID="Image2" runat="server" ImageUrl="images/AIMS_Footer.jpg"></asp:Image>
            </asp:TableCell>
        </asp:TableRow>
        <asp:TableRow>
            <asp:TableCell HorizontalAlign="Center">
                <asp:Label ID="Label1" runat="server">distributed by</asp:Label><asp:HyperLink id="HyperLink1" runat="server" NavigateUrl="http://www.couriersoftware.co.uk" style="text-decoration: none;">
                    <span style="color: #00008B; font: 'Courier New', Courier, monospace; font-size: xx-small;">courier</span><span style="color: #DC143C; font: 'Courier New', Courier, monospace; font-size: xx-small;">software</span>
                </asp:HyperLink>
                <asp:Label ID="Label2" runat="server">© 2006</asp:Label>
            </asp:TableCell>
        </asp:TableRow>
    </asp:Table>
    <br />
    <asp:Label id="lblError" runat="server" forecolor="#00C000" font-size="XX-Small"></asp:Label>
</asp:Content>

