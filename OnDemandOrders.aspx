<%@ Page Language="VB" Theme="AIMSDefault" %>

<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>

<script runat="server">

    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString()

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsPostBack Then
            Call SetTitle()
            Dim dteFromDefault As Date = Date.Today.AddDays(-2)
            Dim dteToDefault As Date = Date.Today.AddDays(1)
            tbFromDate.Text = dteFromDefault.ToString("dd-MMM-yyyy")
            tbToDate.Text = dteToDefault.ToString("dd-MMM-yyyy")
            Call HideAllPanels()
            Call ShowOrders(tbFromDate.Text, tbToDate.Text, bSearchByDate:=True)
            pnlOrders.Visible = True
        End If
    End Sub

    Protected Sub SetTitle()
        Dim sTitle As String = Session("SiteTitle")
        If sTitle <> String.Empty Then
            sTitle += " - "
        End If
        Page.Header.Title = sTitle & "Reports"
    End Sub

    Protected Sub HideAllPanels()
        pnlOrders.Visible = False
        pnlAddAnnotation.Visible = False
        pnlViewAnnotations.Visible = False
    End Sub
    
    Protected Sub btnShowOrdersByOrderNo_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Page.Validate("ByOrderNo")
        If Page.IsValid Then
            Call ShowOrders(tbFromOrderNo.Text, tbToOrderNo.Text, bSearchByDate:=False)
        End If
    End Sub
    
    Protected Sub btnShowOrdersByDate_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Page.Validate("ByDate")
        If Page.IsValid Then
            Call ShowOrders(tbFromDate.Text, tbToDate.Text, bSearchByDate:=True)
        End If
    End Sub
    
    Protected Sub ShowOrders(ByVal sFrom As String, ByVal sTo As String, ByVal bSearchByDate As Boolean)
        Dim oDataReader As SqlDataReader
        Dim oConn As New SqlConnection(gsConn)
        Dim sSQL As String
        Dim sFieldList As String = "id 'Order No', CustomerAccountCode 'Customer', BookingKey 'Transworld Booking Ref', ConsignmentKey 'Transworld Consignment Ref', OrderDateTime 'Order Placed', FirstName + ' ' + LastName 'By', OrderEmail 'Email Text' "
        Dim sQualifiers As String
        If bSearchByDate Then
            sQualifiers = " OrderDateTime >= '" & sFrom & "' AND OrderDateTime <= '" & sTo & "'"
            pbSearchTypeDate = True
        Else
            sQualifiers = " [id] >= " & sFrom & " AND [id] <= " & sTo
            pbSearchTypeDate = False
        End If
        sSQL = "SELECT " & sFieldList & " FROM OnDemandOrder odo INNER JOIN UserProfile up ON odo.OrderPlacedBy = up.[key] INNER JOIN Customer c ON odo.CustomerKey = c.CustomerKey WHERE " & sQualifiers
        Dim oCmd As New SqlCommand(sSQL, oConn)
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            gvOrders.DataSource = oDataReader
            gvOrders.DataBind()
        Catch ex As SqlException
            WebMsgBox.Show("Error in ShowOrdersByOrderNo: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
        
    Protected Sub lnkbtnAddAnnotation_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim sOrderNo = sender.NamingContainer.cells(1).text
        lblAddOrderNo.Text = sOrderNo.ToString
        Call HideAllPanels()
        pnlAddAnnotation.Visible = True
        tbAnnotation.Text = String.Empty
        tbAnnotation.Focus()
    End Sub

    Protected Sub lnkbtnViewAnnotations_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim sOrderNo = sender.NamingContainer.cells(1).text
        lblViewOrderNo.Text = sOrderNo
        Call GetAnnotations(sOrderNo)
        Call HideAllPanels()
        pnlViewAnnotations.Visible = True
    End Sub

    Protected Sub GetAnnotations(ByVal sOrderNo As Integer)
        Dim oDataReader As SqlDataReader
        Dim oConn As New SqlConnection(gsConn)
        Dim sSQL As String
        sSQL = "SELECT FirstName + ' ' + LastName 'Added By', AnnotationDateTime 'Added On', AnnotationText 'Annotation' FROM OnDemandOrderAnnotation odoa INNER JOIN UserProfile up ON AnnotationAddedBy = up.[key] WHERE OrderNo = " & sOrderNo & " ORDER BY AnnotationDateTime"
        Dim oCmd As New SqlCommand(sSQL, oConn)
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            gvAnnotations.DataSource = oDataReader
            gvAnnotations.DataBind()
        Catch ex As SqlException
            WebMsgBox.Show("Error in GetAnnotations: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub btnSaveAnnotation_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call SaveAnnotation()
        If pbSearchTypeDate Then
            Call ShowOrders(tbFromDate.Text, tbToDate.Text, bSearchByDate:=pbSearchTypeDate)
        Else
            Call ShowOrders(tbFromOrderNo.Text, tbToOrderNo.Text, bSearchByDate:=pbSearchTypeDate)
        End If
        Call HideAllPanels()
        pnlOrders.Visible = True
    End Sub
    
    Protected Sub SaveAnnotation()
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand
        oCmd = New SqlCommand("spASPNET_OnDemand_AddAnnotation", oConn)
        oCmd.CommandType = CommandType.StoredProcedure

        Dim paramOrderNo As SqlParameter = New SqlParameter("@OrderNo", SqlDbType.Int)
        paramOrderNo.Value = CInt(lblAddOrderNo.Text)
        oCmd.Parameters.Add(paramOrderNo)
        
        Dim paramAnnotationText As SqlParameter = New SqlParameter("@AnnotationText", SqlDbType.NVarChar, 1000)
        paramAnnotationText.Value = tbAnnotation.Text
        oCmd.Parameters.Add(paramAnnotationText)
        
        Dim paramUserKey As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int)
        paramUserKey.Value = Session("UserKey")
        oCmd.Parameters.Add(paramUserKey)
        
        Try
            oConn.Open()
            oCmd.ExecuteNonQuery()

        Catch ex As Exception
            WebMsgBox.Show("Error in ShowOrdersByOrderNo: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub

    Protected Sub btnBackFromAddAnnotation_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call HideAllPanels()
        pnlOrders.Visible = True
    End Sub

    Protected Sub btnHideAnnotations_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call HideAllPanels()
        pnlOrders.Visible = True
    End Sub
    
    Protected Function sSetViewVisibility(ByVal DataItem As Object) As String
        sSetViewVisibility = String.Empty
        Dim nOrderNo As Integer = DataBinder.Eval(DataItem, "Order No")
        Dim oDataReader As SqlDataReader
        Dim oConn As New SqlConnection(gsConn)
        Dim sSQL As String
        sSQL = "SELECT COUNT(*) FROM OnDemandOrderAnnotation WHERE OrderNo = " & nOrderNo
        Dim oCmd As New SqlCommand(sSQL, oConn)
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            oDataReader.Read()
            sSetViewVisibility = oDataReader(0)
        Catch ex As SqlException
            WebMsgBox.Show("Error in sSetViewVisibility: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Function
    
    Property pbSearchTypeDate() As Boolean
        Get
            Dim o As Object = ViewState("OD_SearchTypeDate")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("OD_SearchTypeDate") = Value
        End Set
    End Property
    
    Protected Sub lnkbtnRefreshOrders_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If pbSearchTypeDate Then
            Call ShowOrders(tbFromDate.Text, tbToDate.Text, bSearchByDate:=pbSearchTypeDate)
        Else
            Call ShowOrders(tbFromOrderNo.Text, tbToOrderNo.Text, bSearchByDate:=pbSearchTypeDate)
        End If
    End Sub
    
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <link href="sprint.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="Form1" runat="Server">
        <main:Header ID="ctlHeader" runat="server"></main:Header>
        <table style="width: 100%" cellpadding="0" cellspacing="0">
            <tr class="bar_reports">
                <td style="width: 50%; white-space: nowrap">
                </td>
                <td style="width: 50%; white-space: nowrap" align="right">
                </td>
            </tr>
        </table>
        <asp:Panel ID="pnlOrders" runat="server" Width="100%">
            &nbsp; <strong><span style="font-size: 10pt; color: #000080">&nbsp;On Demand Orders</span></strong>
            <table style="width: 100%">
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td style="width: 20%">
                    </td>
                    <td style="width: 29%">
                    </td>
                    <td style="width: 20%">
                    </td>
                    <td style="width: 29%" align="right">
                        <asp:LinkButton ID="lnkbtnRefreshOrders" runat="server" OnClick="lnkbtnRefreshOrders_Click">refresh orders</asp:LinkButton></td>
                    <td style="width: 1%">
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td>
                        <asp:Button ID="btnShowOrdersByOrderNo" runat="server" Text="show orders by order #"
                            Width="170px" OnClick="btnShowOrdersByOrderNo_Click" /></td>
                    <td align="left" colspan="3" style="white-space: nowrap">
                        <asp:Label ID="Label2" runat="server" Text="from order #" Font-Names="Verdana" Font-Size="XX-Small"
                            Font-Bold="False" />
                        <asp:TextBox ID="tbFromOrderNo" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Width="65px" />
                        <a runat="server" id="aHelpFromOrderNo" visible="false" onmouseover="return escape('')"
                            style="color: gray; cursor: help">&nbsp;?&nbsp;</a>
                        <asp:Label ID="Label3" runat="server" Text="to order #" Font-Names="Verdana" Font-Size="XX-Small"
                            Font-Bold="False" />
                        <asp:TextBox ID="tbToOrderNo" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Width="65px" />
                        <a runat="server" id="aHelpToOrderNo" visible="false" onmouseover="return escape('')"
                            style="color: gray; cursor: help">&nbsp;?&nbsp;</a>
                        <asp:RangeValidator ID="rvFromOrderNo" runat="server" ControlToValidate="tbFromOrderNo"
                            ErrorMessage="invalid from order #" MaximumValue="99999" MinimumValue="1" Type="Integer"
                            ValidationGroup="ByOrderNo"></asp:RangeValidator>
                        <asp:RangeValidator ID="RangeValidator1" runat="server" ControlToValidate="tbToOrderNo"
                            ErrorMessage="invalid to order #" MaximumValue="99999" MinimumValue="1" Type="Integer"
                            ValidationGroup="ByOrderNo"></asp:RangeValidator>
                        <asp:RequiredFieldValidator ID="rfvFromOrderNo" runat="server" ControlToValidate="tbFromOrderNo"
                            ErrorMessage="from # required" ValidationGroup="ByOrderNo"></asp:RequiredFieldValidator>
                        <asp:RequiredFieldValidator ID="rfvToOrderNo" runat="server" ControlToValidate="tbToOrderNo"
                            ErrorMessage="to # required" ValidationGroup="ByOrderNo"></asp:RequiredFieldValidator></td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td>
                        <asp:Button ID="btnShowOrdersByDate" runat="server" Text="show orders by date" Width="170px"
                            OnClick="btnShowOrdersByDate_Click" /></td>
                    <td colspan="3" style="white-space: nowrap">
                        <asp:Label ID="Label4" runat="server" Text="from (eg 29-Jan-2008)" Font-Names="Verdana"
                            Font-Size="XX-Small" Font-Bold="False" />
                        <asp:TextBox ID="tbFromDate" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Width="85px" />
                        <a href="javascript:;" onclick="window.open('PopupCalendar4.aspx?textbox=tbFromDate','cal','width=300,height=305,left=270,top=180')">
                            <img id="Img1" alt="" src="~/images/SmallCalendar.gif" runat="server" border="0"
                                ie:visible="true" visible="false" /></a> <a runat="server" id="aHelpFromDate" visible="false"
                                    onmouseover="return escape('The date from which you want to see On Demand orders. Click the calendar icon (available in Internet Explorer only) and follow the instructions to select a date , or type the date directly in the format dd-mmm-yyyy, eg 29-Jan-2008')"
                                    style="color: gray; cursor: help">&nbsp;?&nbsp;</a>
                        <asp:Label ID="Label5" runat="server" Text="to (eg 1-Mar-2008)" Font-Names="Verdana"
                            Font-Size="XX-Small" Font-Bold="False" />
                        <asp:TextBox ID="tbToDate" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Width="85px" />
                        <a href="javascript:;" onclick="window.open('PopupCalendar4.aspx?textbox=tbToDate','cal','width=300,height=305,left=270,top=180')">
                            <img id="Img2" alt="" src="~/images/SmallCalendar.gif" runat="server" border="0"
                                ie:visible="true" visible="false" /></a> <a runat="server" id="aHelpToDate" visible="false"
                                    onmouseover="return escape('The date from which you want to see On Demand orders. Click the calendar icon (available in Internet Explorer only) and follow the instructions to select a date , or type the date directly in the format dd-mmm-yyyy, eg 29-Jan-2009')"
                                    style="color: gray; cursor: help">&nbsp;?&nbsp;</a>
                        <asp:RegularExpressionValidator ID="revFromDate" runat="server" ErrorMessage="Invalid from date - use dd-mmm-yyyy"
                            ControlToValidate="tbFromDate" ValidationExpression="^\d\d-(jan|Jan|feb|Feb|mar|Mar|apr|Apr|may|May|jun|Jun|jul|Jul|aug|Aug|sep|Sep|oct|Oct|nov|Nov|dec|Dec)-\d\d\d\d"
                            Font-Names="Verdana" Font-Size="XX-Small" ValidationGroup="ByDate" />
                        <asp:RegularExpressionValidator ID="revToDate" runat="server" ErrorMessage="Invalid to date - use dd-mmm-yyyy"
                            ControlToValidate="tbToDate" ValidationExpression="^\d\d-(jan|Jan|feb|Feb|mar|Mar|apr|Apr|may|May|jun|Jun|jul|Jul|aug|Aug|sep|Sep|oct|Oct|nov|Nov|dec|Dec)-\d\d\d\d"
                            Font-Names="Verdana" Font-Size="XX-Small" ValidationGroup="ByDate" />
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td>
                        <asp:Label ID="Label1" runat="server" Text="Orders:" Font-Names="Verdana" Font-Size="XX-Small"
                            Font-Bold="False" /></td>
                    <td>
                    </td>
                    <td>
                    </td>
                    <td align="right">
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td colspan="4">
                        <asp:GridView ID="gvOrders" runat="server" Width="100%" CellPadding="2" AutoGenerateColumns="False">
                            <Columns>
                                <asp:TemplateField HeaderText="Annotations">
                                    <ItemTemplate>
                                        <asp:LinkButton ID="lnkbtnAddAnnotation" runat="server" OnClick="lnkbtnAddAnnotation_Click">add</asp:LinkButton>
                                        <asp:LinkButton ID="lnkbtnViewAnnotations" runat="server" OnClick="lnkbtnViewAnnotations_Click">view</asp:LinkButton>
                                        <asp:Label ID="Label9" runat="server" Text="("/><asp:Label ID="Label10" runat="server" Text='<%# sSetViewVisibility(Container.DataItem) %>' /><asp:Label ID="Label11" runat="server" Text=")"/>
                                        <asp:HiddenField ID="hidOrderNo" Value='<%# DataBinder.Eval(Container.DataItem,"Order No") %>' runat="server" />
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:BoundField DataField="Order No" HeaderText="Order No" ReadOnly="True" SortExpression="Order No" DataFormatString="{0:d5}" >
                                    <ControlStyle Font-Bold="True" />
                                    <ItemStyle Font-Bold="True" />
                                </asp:BoundField>
                                <asp:BoundField DataField="Customer" HeaderText="Customer" ReadOnly="True" SortExpression="Customer" />
                                <asp:BoundField DataField="Transworld Booking Ref" HeaderText="Transworld Booking Ref" ReadOnly="True"
                                    SortExpression="Transworld Booking Ref" />
                                <asp:BoundField DataField="Transworld Consignment Ref" HeaderText="Transworld Consignment Ref"
                                    ReadOnly="True" SortExpression="Transworld Consignment Ref" />
                                <asp:BoundField DataField="Order Placed" HeaderText="Order Placed" ReadOnly="True"
                                    SortExpression="Order Placed" />
                                <asp:BoundField DataField="By" HeaderText="By" ReadOnly="True" SortExpression="By" />
                                <asp:TemplateField HeaderText="Email Text">
                                    <ItemTemplate>
                                        <asp:Label ID="lblEmailText" runat="server" Text='<%# DataBinder.Eval(Container.DataItem,"Email Text") %>' />
                                    </ItemTemplate>
                                </asp:TemplateField>
                            </Columns>
                            <EmptyDataTemplate>
                                no orders found for this range
                            </EmptyDataTemplate>
                            <AlternatingRowStyle BackColor="#C0FFC0" />
                        </asp:GridView>
                    </td>
                    <td>
                    </td>
                </tr>
            </table>
        </asp:Panel>
        <asp:Panel ID="pnlAddAnnotation" runat="server" Width="100%">
            <asp:Label ID="Label6" runat="server" Text=" Add Annotation to Order" Font-Names="Verdana"
                Font-Size="X-Small" Font-Bold="True" />
            <asp:Label ID="lblAddOrderNo" runat="server" Font-Names="Verdana" Font-Size="X-Small"
                Font-Bold="True" /><br />
            <table style="width: 100%">
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td style="width: 20%">
                    </td>
                    <td style="width: 29%">
                    </td>
                    <td style="width: 20%">
                    </td>
                    <td style="width: 29%" align="right">
                        <asp:Button ID="btnBackFromAddAnnotation" runat="server" Text="go back" OnClick="btnBackFromAddAnnotation_Click" />
                    </td>
                    <td style="width: 1%">
                    </td>
                </tr>
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td style="width: 20%" align="right" valign="top">
                        <asp:Label ID="Label8" runat="server" Text="Annotation:" Font-Names="Verdana" Font-Size="XX-Small"
                            Font-Bold="False" /></td>
                    <td colspan="2">
                        <asp:TextBox ID="tbAnnotation" runat="server" Rows="6" TextMode="MultiLine" Width="100%"
                            MaxLength="950" Font-Names="Verdana" Font-Size="XX-Small" />
                        <br />
                        <asp:Button ID="btnSaveAnnotation" runat="server" OnClick="btnSaveAnnotation_Click"
                            Text="save annotation" />
                    </td>
                    <td style="width: 29%" align="right">
                    </td>
                    <td style="width: 1%">
                    </td>
                </tr>
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td style="width: 20%">
                    </td>
                    <td style="width: 29%">
                    </td>
                    <td style="width: 20%">
                    </td>
                    <td align="right" style="width: 29%">
                    </td>
                    <td style="width: 1%">
                    </td>
                </tr>
            </table>
        </asp:Panel>
        <asp:Panel ID="pnlViewAnnotations" runat="server" Width="100%">
            <asp:Label ID="Label7" runat="server" Text=" Annotations for Order" Font-Names="Verdana"
                Font-Size="X-Small" Font-Bold="True" />
            <asp:Label ID="lblViewOrderNo" runat="server" Font-Names="Verdana" Font-Size="X-Small"
                Font-Bold="True" />
            <table style="width: 100%">
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td style="width: 20%">
                    </td>
                    <td style="width: 29%">
                    </td>
                    <td style="width: 20%">
                    </td>
                    <td style="width: 29%" align="right">
                        <asp:Button ID="btnHideAnnotations1" runat="server" Text="go back" OnClick="btnHideAnnotations_Click" /><td
                            style="width: 1%">
                        </td>
                </tr>
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td colspan="4">
                        <asp:GridView ID="gvAnnotations" runat="server" CellPadding="2" Width="100%">
                        </asp:GridView>
                    </td>
                    <td style="width: 1%;">
                    </td>
                </tr>
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td style="width: 20%">
                        <asp:Button ID="btnHideAnnotations2" runat="server" OnClick="btnHideAnnotations_Click"
                            Text="go back" />
                    </td>
                    <td style="width: 29%">
                    </td>
                    <td style="width: 20%">
                    </td>
                    <td align="right" style="width: 29%">
                    </td>
                    <td style="width: 1%">
                    </td>
                </tr>
            </table>
        </asp:Panel>
    </form>
</body>
</html>
