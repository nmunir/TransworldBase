<%@ Page Language="VB" Theme="AIMSDefault" %>

<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<script runat="server">
    
    Const NO_VALUE_AVAILABLE_MESSAGE As String = "(no value available)"

    Dim gsConn As String = ConfigurationManager.AppSettings("AIMSRootConnectionString")
    Dim arrMonthNames() As String = {"Zero-Based", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"}
    
    Protected Sub Page_Load(ByVal Source As Object, ByVal E As EventArgs)
        If Not IsPostBack Then
            Call ShowAvailableYears()
        End If
        If Not IsNumeric(Session("CustomerKey")) Then
            'Server.Transfer("../session_expired.aspx")
        End If
    End Sub

    Protected Sub SetTitle()
        Dim sTitle As String = Session("SiteTitle")
        If sTitle <> String.Empty Then
            sTitle += " - "
        End If
        Page.Header.Title = sTitle & "Western Union Management Information Monthly Report"
    End Sub

    Protected Sub ShowAvailableYears()
        Dim sSQL As String = "SELECT DISTINCT Year FROM ClientData_WU_MIMonthlyReport WHERE VisibleToClient = 1"
        Dim dtYears As DataTable = ExecuteQueryToDataTable(sSQL)
        If dtYears.Rows.Count > 0 Then
            rptYear.DataSource = dtYears
            rptYear.DataBind()
            rptYear.Visible = True
        Else
            rptYear.Visible = False
            lblLegendYear.Text = "No data available"
        End If
    End Sub

    Protected Sub rptYear_item_click(ByVal s As Object, ByVal e As RepeaterCommandEventArgs)
        Dim item As RepeaterItem
        For Each item In s.Items
            Dim x As LinkButton = CType(item.Controls(3), LinkButton)
            x.ForeColor = System.Drawing.Color.Blue
        Next
        Dim Link As LinkButton = CType(e.CommandSource, LinkButton)
        Link.ForeColor = System.Drawing.Color.Red
    End Sub
    
    Protected Sub btn_ShowMonths_click(ByVal sender As Object, ByVal e As CommandEventArgs)
        pnYear = CStr(e.CommandArgument)
        'lblYearHeader.Text = psYear
        rptMonth.Visible = True
        lblLegendMonth.Visible = True
        GetAvailableMonths()
        rbUK.Enabled = False
        rbIreland.Enabled = False
    End Sub
    
    Protected Sub btn_ShowReport_click(ByVal sender As Object, ByVal e As CommandEventArgs)
        pnMonth = e.CommandArgument
        Call ShowReport()
        rbUK.Enabled = True
        rbIreland.Enabled = True
    End Sub
    
    Protected Sub ShowReport()
        Dim sSQL As String = "SELECT * FROM ClientData_WU_MIMonthlyReport WHERE Year = " & pnYear & " AND Month = " & pnMonth & " AND Country = '"
        If rbUK.Checked Then
            sSQL &= "UK'"
        Else
            sSQL &= "IRELAND'"
        End If
        Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
        If dt.Rows.Count = 1 Then
            Call PopulateForm(dt.Rows(0))
            tabData.Visible = True
            lblReportPeriod.Text = arrMonthNames(pnMonth) & " " & pnYear.ToString
        End If
    End Sub
    
    Protected Sub GetAvailableMonths()
        Dim sSQL As String = "SELECT DISTINCT Month FROM ClientData_WU_MIMonthlyReport WHERE VisibleToClient = 1 AND Year = " & pnYear
        Dim dtMonths As DataTable = ExecuteQueryToDataTable(sSQL)
        If dtMonths.Rows.Count > 0 Then
            rptMonth.DataSource = dtMonths
            rptMonth.DataBind()
            rptMonth.Visible = True
        Else
            rptMonth.Visible = False
        End If
    End Sub

    Protected Sub PopulateForm(dr As DataRow)
        Const NO_VALUE_AVAILABLE_MESSAGE As String = "(no value available)"
        
        Dim nVisibleToClient As Int32 = dr("VisibleToClient")

        If dr("OrderBreakdownOperations") = -1 Then
            lblOrderBreakdownOperations.Text = NO_VALUE_AVAILABLE_MESSAGE
        Else
            lblOrderBreakdownOperations.Text = dr("OrderBreakdownOperations")
        End If

        If dr("OrderBreakdownMarketing") = -1 Then
            lblOrderBreakdownMarketing.Text = NO_VALUE_AVAILABLE_MESSAGE
        Else
            lblOrderBreakdownMarketing.Text = dr("OrderBreakdownMarketing")
        End If

        If dr("OrderBreakdownFININT") = -1 Then
            lblOrderBreakdownFININT.Text = NO_VALUE_AVAILABLE_MESSAGE
        Else
            lblOrderBreakdownFININT.Text = dr("OrderBreakdownFININT")
        End If

        If dr("OrderBreakdownCosta") = -1 Then
            lblOrderBreakdownCosta.Text = NO_VALUE_AVAILABLE_MESSAGE
        Else
            lblOrderBreakdownCosta.Text = dr("OrderBreakdownCosta")
        End If

        If dr("OrderBreakdownPrePaid") = -1 Then
            lblOrderBreakdownPrePaid.Text = NO_VALUE_AVAILABLE_MESSAGE
        Else
            lblOrderBreakdownPrePaid.Text = dr("OrderBreakdownPrePaid")
        End If

        If dr("StorageCostsOperations") = -1 Then
            lblStorageCostsOperations.Text = NO_VALUE_AVAILABLE_MESSAGE
        Else
            lblStorageCostsOperations.Text = dr("StorageCostsOperations")
        End If

        If dr("StorageCostsMarketing") = -1 Then
            lblStorageCostsMarketing.Text = NO_VALUE_AVAILABLE_MESSAGE
        Else
            lblStorageCostsMarketing.Text = dr("StorageCostsMarketing")
        End If

        If dr("StorageCostsFININT") = -1 Then
            lblStorageCostsFININT.Text = NO_VALUE_AVAILABLE_MESSAGE
        Else
            lblStorageCostsFININT.Text = dr("StorageCostsFININT")
        End If

        If dr("StorageCostsCosta") = -1 Then
            lblStorageCostsCosta.Text = NO_VALUE_AVAILABLE_MESSAGE
        Else
            lblStorageCostsCosta.Text = dr("StorageCostsCosta")
        End If

        If dr("StorageCostsPrePaid") = -1 Then
            lblStorageCostsPrePaid.Text = NO_VALUE_AVAILABLE_MESSAGE
        Else
            lblStorageCostsPrePaid.Text = dr("StorageCostsPrePaid")
        End If

        If dr("LogisticsCostsCourierOperations") = -1 Then
            lblLogisticsCostsCourierOperations.Text = NO_VALUE_AVAILABLE_MESSAGE
        Else
            lblLogisticsCostsCourierOperations.Text = dr("LogisticsCostsCourierOperations")
        End If

        If dr("LogisticsCostsCourierMarketing") = -1 Then
            lblLogisticsCostsCourierMarketing.Text = NO_VALUE_AVAILABLE_MESSAGE
        Else
            lblLogisticsCostsCourierMarketing.Text = dr("LogisticsCostsCourierMarketing")
        End If

        If dr("LogisticsCostsCourierFININT") = -1 Then
            lblLogisticsCostsCourierFININT.Text = NO_VALUE_AVAILABLE_MESSAGE
        Else
            lblLogisticsCostsCourierFININT.Text = dr("LogisticsCostsCourierFININT")
        End If

        If dr("LogisticsCostsCourierCosta") = -1 Then
            lblLogisticsCostsCourierCosta.Text = NO_VALUE_AVAILABLE_MESSAGE
        Else
            lblLogisticsCostsCourierCosta.Text = dr("LogisticsCostsCourierCosta")
        End If

        If dr("LogisticsCostsPrepaid") = -1 Then
            lblLogisticsCostsPrepaid.Text = NO_VALUE_AVAILABLE_MESSAGE
        Else
            lblLogisticsCostsPrepaid.Text = dr("LogisticsCostsPrepaid")
        End If

        If dr("LogisticsCostsMailFulfilment") = -1 Then
            lblLogisticsCostsMailFulfilment.Text = NO_VALUE_AVAILABLE_MESSAGE
        Else
            lblLogisticsCostsMailFulfilment.Text = dr("LogisticsCostsMailFulfilment")
        End If

        If dr("LogisticsCostsAdHocFulfilment") = -1 Then
            lblLogisticsCostsAdHocFulfilment.Text = NO_VALUE_AVAILABLE_MESSAGE
        Else
            lblLogisticsCostsAdHocFulfilment.Text = dr("LogisticsCostsAdHocFulfilment")
        End If

        If dr("ServiceFeesPickFeesOperations") = -1 Then
            lblServiceFeesPickFeesOperations.Text = NO_VALUE_AVAILABLE_MESSAGE
        Else
            lblServiceFeesPickFeesOperations.Text = dr("ServiceFeesPickFeesOperations")
        End If

        If dr("ServiceFeesPickFeesMarketing") = -1 Then
            lblServiceFeesPickFeesMarketing.Text = NO_VALUE_AVAILABLE_MESSAGE
        Else
            lblServiceFeesPickFeesMarketing.Text = dr("ServiceFeesPickFeesMarketing")
        End If

        If dr("ServiceFeesPickFeesFININT") = -1 Then
            lblServiceFeesPickFeesFININT.Text = NO_VALUE_AVAILABLE_MESSAGE
        Else
            lblServiceFeesPickFeesFININT.Text = dr("ServiceFeesPickFeesFININT")
        End If

        If dr("ServiceFeesPickFeesCosta") = -1 Then
            lblServiceFeesPickFeesCosta.Text = NO_VALUE_AVAILABLE_MESSAGE
        Else
            lblServiceFeesPickFeesCosta.Text = dr("ServiceFeesPickFeesCosta")
        End If

        If dr("ServiceFeesPickFeesPrePaid") = -1 Then
            lblServiceFeesPickFeesPrePaid.Text = NO_VALUE_AVAILABLE_MESSAGE
        Else
            lblServiceFeesPickFeesPrePaid.Text = dr("ServiceFeesPickFeesPrePaid")
        End If

        If dr("ServiceFeesGoodsInOperations") = -1 Then
            lblServiceFeesGoodsInOperations.Text = NO_VALUE_AVAILABLE_MESSAGE
        Else
            lblServiceFeesGoodsInOperations.Text = dr("ServiceFeesGoodsInOperations")
        End If

        If dr("ServiceFeesGoodsInMarketing") = -1 Then
            lblServiceFeesGoodsInMarketing.Text = NO_VALUE_AVAILABLE_MESSAGE
        Else
            lblServiceFeesGoodsInMarketing.Text = dr("ServiceFeesGoodsInMarketing")
        End If

        If dr("ServiceFeesGoodsInFININT") = -1 Then
            lblServiceFeesGoodsInFININT.Text = NO_VALUE_AVAILABLE_MESSAGE
        Else
            lblServiceFeesGoodsInFININT.Text = dr("ServiceFeesGoodsInFININT")
        End If

        If dr("ServiceFeesGoodsInCosta") = -1 Then
            lblServiceFeesGoodsInCosta.Text = NO_VALUE_AVAILABLE_MESSAGE
        Else
            lblServiceFeesGoodsInCosta.Text = dr("ServiceFeesGoodsInCosta")
        End If

        If dr("ServiceFeesGoodsInPrePaid") = -1 Then
            lblServiceFeesGoodsInPrePaid.Text = NO_VALUE_AVAILABLE_MESSAGE
        Else
            lblServiceFeesGoodsInPrePaid.Text = dr("ServiceFeesGoodsInPrePaid")
        End If

        If dr("ServiceFeesDestructionFeesOperations") = -1 Then
            lblServiceFeesDestructionFeesOperations.Text = NO_VALUE_AVAILABLE_MESSAGE
        Else
            lblServiceFeesDestructionFeesOperations.Text = dr("ServiceFeesDestructionFeesOperations")
        End If

        If dr("ServiceFeesDestructionFeesMarketing") = -1 Then
            lblServiceFeesDestructionFeesMarketing.Text = NO_VALUE_AVAILABLE_MESSAGE
        Else
            lblServiceFeesDestructionFeesMarketing.Text = dr("ServiceFeesDestructionFeesMarketing")
        End If

        If dr("ServiceFeesDestructionFeesFININT") = -1 Then
            lblServiceFeesDestructionFeesFININT.Text = NO_VALUE_AVAILABLE_MESSAGE
        Else
            lblServiceFeesDestructionFeesFININT.Text = dr("ServiceFeesDestructionFeesFININT")
        End If

        If dr("ServiceFeesDestructionFeesCosta") = -1 Then
            lblServiceFeesDestructionFeesCosta.Text = NO_VALUE_AVAILABLE_MESSAGE
        Else
            lblServiceFeesDestructionFeesCosta.Text = dr("ServiceFeesDestructionFeesCosta")
        End If

        If dr("ServiceFeesDestructionFeesPrePaid") = -1 Then
            lblServiceFeesDestructionFeesPrePaid.Text = NO_VALUE_AVAILABLE_MESSAGE
        Else
            lblServiceFeesDestructionFeesPrePaid.Text = dr("ServiceFeesDestructionFeesPrePaid")
        End If

        If dr("ClientNotes") <> String.Empty Then
            lblNotes.Text = dr("ClientNotes")
            trNotes.Visible = True
        Else
            trNotes.Visible = False
        End If
        
        Call FormatLabel(lblOrderBreakdownOperations)
        Call FormatLabel(lblOrderBreakdownMarketing)
        Call FormatLabel(lblOrderBreakdownFININT)
        Call FormatLabel(lblOrderBreakdownCosta)
        Call FormatLabel(lblOrderBreakdownPrePaid)
        Call FormatLabel(lblStorageCostsOperations)
        Call FormatLabel(lblStorageCostsMarketing)
        Call FormatLabel(lblStorageCostsFININT)
        Call FormatLabel(lblStorageCostsCosta)
        Call FormatLabel(lblStorageCostsPrePaid)
        Call FormatLabel(lblLogisticsCostsCourierOperations)
        Call FormatLabel(lblLogisticsCostsCourierMarketing)
        Call FormatLabel(lblLogisticsCostsCourierFININT)
        Call FormatLabel(lblLogisticsCostsCourierCosta)
        Call FormatLabel(lblLogisticsCostsPrepaid)
        Call FormatLabel(lblLogisticsCostsMailFulfilment)
        Call FormatLabel(lblLogisticsCostsAdHocFulfilment)
        Call FormatLabel(lblServiceFeesPickFeesOperations)
        Call FormatLabel(lblServiceFeesPickFeesMarketing)
        Call FormatLabel(lblServiceFeesPickFeesFININT)
        Call FormatLabel(lblServiceFeesPickFeesCosta)
        Call FormatLabel(lblServiceFeesPickFeesPrePaid)
        Call FormatLabel(lblServiceFeesGoodsInOperations)
        Call FormatLabel(lblServiceFeesGoodsInMarketing)
        Call FormatLabel(lblServiceFeesGoodsInFININT)
        Call FormatLabel(lblServiceFeesGoodsInCosta)
        Call FormatLabel(lblServiceFeesGoodsInPrePaid)
        Call FormatLabel(lblServiceFeesDestructionFeesOperations)
        Call FormatLabel(lblServiceFeesDestructionFeesMarketing)
        Call FormatLabel(lblServiceFeesDestructionFeesFININT)
        Call FormatLabel(lblServiceFeesDestructionFeesCosta)
        Call FormatLabel(lblServiceFeesDestructionFeesPrePaid)

    End Sub
    
    Protected Sub FormatLabel(lbl As Label)
        If lbl.Text = NO_VALUE_AVAILABLE_MESSAGE Then
            lbl.ForeColor = Drawing.Color.Silver
            lbl.Font.Bold = False
            'lbl.Font.Size = FontSize.Small
        Else
            lbl.ForeColor = Drawing.Color.Black
            'lbl.Font.Size = FontSize.Medium
            lbl.Font.Bold = True
            If lbl.Text.Contains(".") Then
                Dim nDotPos As Int32 = lbl.Text.IndexOf(".")
                lbl.Text = lbl.Text.Substring(0, nDotPos + 3)
                lbl.Text = "£" & lbl.Text
            End If
        End If
    End Sub
    
    ' lblOrderBreakdownOperations.Text 
    ' lblOrderBreakdownMarketing.Text
    ' lblOrderBreakdownFININT.Text
    ' lblOrderBreakdownCosta.Text
    ' lblOrderBreakdownPrePaid.Text
    ' lblStorageCostsOperations.Text
    ' lblStorageCostsMarketing.Text
    ' lblStorageCostsFININT.Text
    ' lblStorageCostsCosta.Text
    ' lblStorageCostsPrePaid.Text
    ' lblLogisticsCostsCourierOperations.Text
    ' lblLogisticsCostsCourierMarketing.Text
    ' lblLogisticsCostsCourierFININT.Text
    ' lblLogisticsCostsCourierCosta.Text
    ' lblLogisticsCostsPrepaid.Text
    ' lblLogisticsCostsMailFulfilment.Text
    ' lblLogisticsCostsAdHocFulfilment.Text
    ' lblServiceFeesPickFeesOperations.Text
    ' lblServiceFeesPickFeesMarketing.Text
    ' lblServiceFeesPickFeesFININT.Text
    ' lblServiceFeesPickFeesCosta.Text
    ' lblServiceFeesPickFeesPrePaid.Text
    ' lblServiceFeesGoodsInOperations.Text
    ' lblServiceFeesGoodsInMarketing.Text
    ' lblServiceFeesGoodsInFININT.Text
    ' lblServiceFeesGoodsInCosta.Text
    ' lblServiceFeesGoodsInPrePaid.Text
    ' lblServiceFeesDestructionFeesOperations.Text
    ' lblServiceFeesDestructionFeesMarketing.Text
    ' lblServiceFeesDestructionFeesFININT.Text
    ' lblServiceFeesDestructionFeesCosta.Text
    ' lblServiceFeesDestructionFeesPrePaid.Text

    Protected Function ExecuteQueryToDataTable(ByVal sQuery As String) As DataTable
        Dim oDataTable As New DataTable
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter(sQuery, oConn)
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

    Property pnYear() As Int32
        Get
            Dim o As Object = ViewState("WIMR_Year")
            If o Is Nothing Then
                Return -1
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Int32)
            ViewState("WIMR_Year") = Value
        End Set
    End Property

    'Property psYear() As String
    '    Get
    '        Dim o As Object = ViewState("WIMR_Year")
    '        If o Is Nothing Then
    '            Return ""
    '        End If
    '        Return CStr(o)
    '    End Get
    '    Set(ByVal Value As String)
    '        ViewState("WIMR_Year") = Value
    '    End Set
    'End Property
    
    Property pnMonth() As Int32
        Get
            Dim o As Object = ViewState("WIMR_Month")
            If o Is Nothing Then
                Return -1
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Int32)
            ViewState("WIMR_Month") = Value
        End Set
    End Property

    Property psCountry() As String
        Get
            Dim o As Object = ViewState("WIMR_Country")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("WIMR_Country") = Value
        End Set
    End Property
    
    Protected Sub rbUK_CheckedChanged(sender As Object, e As System.EventArgs)
        Dim rb As RadioButton = sender
        If rb.Checked Then
            psCountry = "UK"
            Call ShowReport()
        End If
    End Sub

    Protected Sub rbIreland_CheckedChanged(sender As Object, e As System.EventArgs)
        Dim rb As RadioButton = sender
        If rb.Checked Then
            psCountry = "IRELAND"
            Call ShowReport()
        End If
    End Sub
    
</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <style type="text/css">
        .style1
        {
            height: 25px;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
    <%--<main:Header ID="ctlHeader" runat="server" />--%>
    <asp:ScriptManager ID="ScriptManager1" runat="server" />
    <div>
        <table id="tabHeader" style="width: 100%">
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
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td align="right">
                    <asp:Label ID="lblLegendYear" runat="server" Font-Names="Verdana" Font-Size="Small"
                        Text="Year" Font-Bold="True" />
                </td>
                <td>
                    <asp:Repeater ID="rptYear" runat="server" OnItemCommand="rptYear_item_click">
                        <ItemTemplate>
                            <asp:Image runat="server" ImageUrl="../images/icon_arrow.gif" />
                            <asp:LinkButton runat="server" CommandArgument='<%# Container.DataItem("Year")%>'
                                ForeColor="Blue" OnCommand="btn_ShowMonths_click" Text='<%# Container.DataItem("Year")%>' />
                            <br />
                        </ItemTemplate>
                    </asp:Repeater>
                </td>
                <td align="right">
                    <asp:Label ID="lblLegendMonth" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Month"
                        Font-Bold="True" Visible="false" />
                </td>
                <td>
                    <asp:Repeater ID="rptMonth" runat="server" Visible="False">
                        <ItemTemplate>
                            <asp:Image runat="server" ImageUrl="../images/icon_arrow.gif" />
                            <asp:LinkButton runat="server" CommandArgument='<%# Container.DataItem("Month")%>'
                                ForeColor="Blue" OnCommand="btn_ShowReport_click" Text='<%# arrMonthNames(Container.DataItem("Month"))%>' />
                            <br />
                        </ItemTemplate>
                    </asp:Repeater>
                </td>
                <td align="right">
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
        </table>
        <table id="tabData" runat="server" visible="false" style="width: 100%">
            <tr>
                <td colspan="6">
                    <hr></td>
            </tr>
            <tr>
                <td colspan="6" class="style1">
                    &nbsp;
                    &nbsp;
                    &nbsp;
                    &nbsp;
                    &nbsp;
                    &nbsp;
                    <asp:Label ID="Label81" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Western Union Management Information Summary for "
                        Font-Bold="True" />
                    <asp:Label ID="lblReportPeriod" runat="server" Font-Names="Verdana" Font-Size="Small" 
                        Font-Bold="True" />
                &nbsp;
                    <asp:RadioButton ID="rbUK" runat="server" GroupName="Country" 
            Text="UK" AutoPostBack="True" oncheckedchanged="rbUK_CheckedChanged" 
            Enabled="False" Checked="True" Font-Names="Verdana" Font-Size="Small" />
        &nbsp;<asp:RadioButton ID="rbIreland" runat="server" GroupName="Country" 
            Text="IRELAND" AutoPostBack="True" 
            oncheckedchanged="rbIreland_CheckedChanged" Enabled="False" Font-Names="Verdana" 
                        Font-Size="Small" />
                </td>
            </tr>
            <tr>
                <td align="right">
                    &nbsp;</td>
                <td>
                    &nbsp;
                </td>
                <td align="right">
                    &nbsp;</td>
                <td>
                    &nbsp;
                </td>
                <td align="right">
                    &nbsp;</td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td align="right">
                    <asp:Label ID="Label1" runat="server" Font-Names="Verdana" Font-Size="Small" Text="ORDERS"
                        Font-Bold="True" />
                </td>
                <td>
                    &nbsp;
                </td>
                <td align="right">
                    <asp:Label ID="Label2" runat="server" Font-Names="Verdana" Font-Size="Small" Text="STORAGE COSTS"
                        Font-Bold="True" />
                </td>
                <td>
                    &nbsp;
                </td>
                <td align="right">
                    <asp:Label ID="Label42" runat="server" Font-Names="Verdana" Font-Size="Small" Text="LOGISTICS COSTS"
                        Font-Bold="True" />
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td align="right">
                    <asp:Label ID="Label43" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Operations:" />
                </td>
                <td>
                    <asp:Label ID="lblOrderBreakdownOperations" runat="server" Font-Names="Verdana" Font-Size="Small"/>
                </td>
                <td align="right">
                    <asp:Label ID="Label47" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Operations:" />
                </td>
                <td>
                    <asp:Label ID="lblStorageCostsOperations" runat="server" Font-Names="Verdana" Font-Size="Small"/>
                </td>
                <td align="right">
                    <asp:Label ID="Label54" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Courier Operations:" />
                </td>
                <td>
                    <asp:Label ID="lblLogisticsCostsCourierOperations" runat="server" Font-Names="Verdana"
                        Font-Size="Small"/>
                </td>
            </tr>
            <tr>
                <td align="right">
                    <asp:Label ID="Label44" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Marketing:" />
                </td>
                <td>
                    <asp:Label ID="lblOrderBreakdownMarketing" runat="server" Font-Names="Verdana" Font-Size="Small"/>
                </td>
                <td align="right">
                    <asp:Label ID="Label48" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Marketing:" />
                </td>
                <td>
                    <asp:Label ID="lblStorageCostsMarketing" runat="server" Font-Names="Verdana" Font-Size="Small"/>
                </td>
                <td align="right">
                    <asp:Label ID="Label55" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Courier Marketing:" />
                </td>
                <td>
                    <asp:Label ID="lblLogisticsCostsCourierMarketing" runat="server" Font-Names="Verdana"
                        Font-Size="Small"/>
                </td>
            </tr>
            <tr>
                <td align="right">
                    <asp:Label ID="Label45" runat="server" Font-Names="Verdana" Font-Size="Small" Text="FININT:" />
                </td>
                <td>
                    <asp:Label ID="lblOrderBreakdownFININT" runat="server" Font-Names="Verdana" Font-Size="Small"/>
                </td>
                <td align="right">
                    <asp:Label ID="Label49" runat="server" Font-Names="Verdana" Font-Size="Small" Text="FININT:" />
                </td>
                <td>
                    <asp:Label ID="lblStorageCostsFININT" runat="server" Font-Names="Verdana" Font-Size="Small"/>
                </td>
                <td align="right">
                    <asp:Label ID="Label56" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Courier FININT:" />
                </td>
                <td>
                    <asp:Label ID="lblLogisticsCostsCourierFININT" runat="server" Font-Names="Verdana"
                        Font-Size="Small"/>
                </td>
            </tr>
            <tr>
                <td align="right">
                    <asp:Label ID="Label50" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Costa:" />
                </td>
                <td>
                    <asp:Label ID="lblOrderBreakdownCosta" runat="server" Font-Names="Verdana" Font-Size="Small"/>
                </td>
                <td align="right">
                    <asp:Label ID="Label51" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Costa:" />
                </td>
                <td>
                    <asp:Label ID="lblStorageCostsCosta" runat="server" Font-Names="Verdana" Font-Size="Small"/>
                </td>
                <td align="right">
                    <asp:Label ID="Label57" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Courier Costa:" />
                </td>
                <td>
                    <asp:Label ID="lblLogisticsCostsCourierCosta" runat="server" Font-Names="Verdana" Font-Size="Small"/>
                </td>
            </tr>
            <tr>
                <td align="right">
                    <asp:Label ID="Label52" runat="server" Font-Names="Verdana" Font-Size="Small" Text="PrePaid:" />
                </td>
                <td>
                    <asp:Label ID="lblOrderBreakdownPrePaid" runat="server" Font-Names="Verdana" Font-Size="Small" />
                </td>
                <td align="right">
                    <asp:Label ID="Label53" runat="server" Font-Names="Verdana" Font-Size="Small" Text="PrePaid:" />
                </td>
                <td>
                    <asp:Label ID="lblStorageCostsPrePaid" runat="server" Font-Names="Verdana" Font-Size="Small"/>
                </td>
                <td align="right">
                    <asp:Label ID="Label58" runat="server" Font-Names="Verdana" Font-Size="Small" Text="PrePaid:" />
                </td>
                <td>
                    <asp:Label ID="lblLogisticsCostsPrepaid" runat="server" Font-Names="Verdana" Font-Size="Small"/>
                </td>
            </tr>
            <tr>
                <td align="right">
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
                <td align="right">
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
                <td align="right">
                    <asp:Label ID="Label59" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Mail Fulfilment:" />
                </td>
                <td>
                    <asp:Label ID="lblLogisticsCostsMailFulfilment" runat="server" Font-Names="Verdana"
                        Font-Size="Small"></asp:Label>
                </td>
            </tr>
            <tr>
                <td align="right">
                </td>
                <td>
                </td>
                <td align="right">
                </td>
                <td>
                </td>
                <td align="right">
                    <asp:Label ID="Label60" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Ad Hoc Fulfilment:" />
                </td>
                <td>
                    <asp:Label ID="lblLogisticsCostsAdHocFulfilment" runat="server" Font-Names="Verdana"
                        Font-Size="Small"></asp:Label>
                </td>
            </tr>
            <tr>
                <td align="right">
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
                <td align="right">
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
                <td align="right">
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td align="right">
                    <asp:Label ID="Label61" runat="server" Font-Names="Verdana" Font-Size="Small" Text="SERVICE FEES - PICK FEES"
                        Font-Bold="True" />
                </td>
                <td>
                    &nbsp;
                </td>
                <td align="right">
                    <asp:Label ID="Label63" runat="server" Font-Names="Verdana" Font-Size="Small" Text="SERVICE FEES - GOODS IN"
                        Font-Bold="True" />
                </td>
                <td>
                    &nbsp;
                </td>
                <td align="right">
                    <asp:Label ID="Label64" runat="server" Font-Names="Verdana" Font-Size="Small" Text="SERVICE FEES - DESTRUCTION FEES"
                        Font-Bold="True" />
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td align="right">
                    <asp:Label ID="Label65" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Operations:" />
                </td>
                <td style="margin-left: 40px">
                    <asp:Label ID="lblServiceFeesPickFeesOperations" runat="server" Font-Names="Verdana"
                        Font-Size="Small"></asp:Label>
                </td>
                <td align="right">
                    <asp:Label ID="Label70" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Operations:" />
                </td>
                <td>
                    <asp:Label ID="lblServiceFeesGoodsInOperations" runat="server" Font-Names="Verdana"
                        Font-Size="Small"></asp:Label>
                </td>
                <td align="right">
                    <asp:Label ID="Label71" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Operations:" />
                </td>
                <td>
                    <asp:Label ID="lblServiceFeesDestructionFeesOperations" runat="server" Font-Names="Verdana"
                        Font-Size="Small"></asp:Label>
                </td>
            </tr>
            <tr>
                <td align="right">
                    <asp:Label ID="Label66" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Marketing:" />
                </td>
                <td>
                    <asp:Label ID="lblServiceFeesPickFeesMarketing" runat="server" Font-Names="Verdana"
                        Font-Size="Small"></asp:Label>
                </td>
                <td align="right">
                    <asp:Label ID="Label72" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Marketing:" />
                </td>
                <td>
                    <asp:Label ID="lblServiceFeesGoodsInMarketing" runat="server" Font-Names="Verdana"
                        Font-Size="Small"></asp:Label>
                </td>
                <td align="right">
                    <asp:Label ID="Label73" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Marketing:" />
                </td>
                <td>
                    <asp:Label ID="lblServiceFeesDestructionFeesMarketing" runat="server" Font-Names="Verdana"
                        Font-Size="Small"></asp:Label>
                </td>
            </tr>
            <tr>
                <td align="right">
                    <asp:Label ID="Label67" runat="server" Font-Names="Verdana" Font-Size="Small" Text="FININT:" />
                </td>
                <td>
                    <asp:Label ID="lblServiceFeesPickFeesFININT" runat="server" Font-Names="Verdana"
                        Font-Size="Small"></asp:Label>
                </td>
                <td align="right">
                    <asp:Label ID="Label74" runat="server" Font-Names="Verdana" Font-Size="Small" Text="FININT:" />
                </td>
                <td>
                    <asp:Label ID="lblServiceFeesGoodsInFININT" runat="server" Font-Names="Verdana" Font-Size="Small"></asp:Label>
                </td>
                <td align="right">
                    <asp:Label ID="Label75" runat="server" Font-Names="Verdana" Font-Size="Small" Text="FININT:" />
                </td>
                <td>
                    <asp:Label ID="lblServiceFeesDestructionFeesFININT" runat="server" Font-Names="Verdana"
                        Font-Size="Small"></asp:Label>
                </td>
            </tr>
            <tr>
                <td align="right">
                    <asp:Label ID="Label68" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Costa:" />
                </td>
                <td>
                    <asp:Label ID="lblServiceFeesPickFeesCosta" runat="server" Font-Names="Verdana" Font-Size="Small"></asp:Label>
                </td>
                <td align="right">
                    <asp:Label ID="Label76" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Costa:" />
                </td>
                <td>
                    <asp:Label ID="lblServiceFeesGoodsInCosta" runat="server" Font-Names="Verdana" Font-Size="Small"></asp:Label>
                </td>
                <td align="right">
                    <asp:Label ID="Label77" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Costa:" />
                </td>
                <td>
                    <asp:Label ID="lblServiceFeesDestructionFeesCosta" runat="server" Font-Names="Verdana"
                        Font-Size="Small"></asp:Label>
                </td>
            </tr>
            <tr>
                <td align="right">
                    <asp:Label ID="Label69" runat="server" Font-Names="Verdana" Font-Size="Small" Text="PrePaid:" />
                </td>
                <td>
                    <asp:Label ID="lblServiceFeesPickFeesPrePaid" runat="server" Font-Names="Verdana"
                        Font-Size="Small"></asp:Label>
                </td>
                <td align="right">
                    <asp:Label ID="Label78" runat="server" Font-Names="Verdana" Font-Size="Small" Text="PrePaid:" />
                </td>
                <td>
                    <asp:Label ID="lblServiceFeesGoodsInPrePaid" runat="server" Font-Names="Verdana"
                        Font-Size="Small"></asp:Label>
                </td>
                <td align="right">
                    <asp:Label ID="Label79" runat="server" Font-Names="Verdana" Font-Size="Small" Text="PrePaid:" />
                </td>
                <td>
                    <asp:Label ID="lblServiceFeesDestructionFeesPrePaid" runat="server" Font-Names="Verdana"
                        Font-Size="Small"></asp:Label>
                </td>
            </tr>
            <tr>
                <td align="right">
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
                <td align="right">
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
                <td align="right">
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr id="trNotes" runat="server" visible="false">
                <td align="right" colspan="6">
                    <table style="width: 100%">
                        <tr>
                            <td style="width: 2%">
                                &nbsp;
                            </td>
                            <td style="width: 47%" align="left">
                                <asp:Label ID="Label80" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Notes:" />
                            </td>
                            <td style="width: 2%">
                                &nbsp;
                            </td>
                            <td style="width: 47%" align="left">
                                &nbsp;
                            </td>
                            <td style="width: 2%">
                                &nbsp;
                            </td>
                        </tr>
                        <tr>
                            <td style="width: 2%">
                                &nbsp;</td>
                            <td style="width: 47%" align="left">
                                &nbsp;</td>
                            <td style="width: 2%">
                                &nbsp;</td>
                            <td style="width: 47%" align="left">
                                &nbsp;</td>
                            <td style="width: 2%">
                                &nbsp;</td>
                        </tr>
                        <tr>
                            <td>
                                &nbsp;
                            </td>
                            <td colspan="3" align="left">
                                <asp:Label ID="lblNotes" runat="server" Font-Names="Verdana" Font-Size="Small" />
                                <br />
                            </td>
                            <td>
                                &nbsp;
                                <br />
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </div>
    </form>
</body>
</html>
