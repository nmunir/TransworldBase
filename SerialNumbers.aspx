<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>

<script runat="server">

    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString()

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsPostBack Then
            Call SetTitle()
            Dim dteFromDefault As Date = Date.Today.AddDays(-30)
            Dim dteToDefault As Date = Date.Today.AddDays(1)
            tbFromDate.Text = dteFromDefault.ToString("dd-MMM-yyyy")
            tbToDate.Text = dteToDefault.ToString("dd-MMM-yyyy")
            'Call HideAllPanels()
            Call GetProducts()
            'Call ShowSerialNos(tbFromDate.Text, tbToDate.Text, bSearchByDate:=True)
            gvSerialNos.Visible = False
            lblSerialNumbers.Visible = False
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
        pnlSerialNumbers.Visible = False
    End Sub
    
    Protected Sub btnShowSerialNosByConsignmentNo_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Page.Validate("ByOrderNo")
        If Page.IsValid Then
            Call ShowSerialNos(tbFromConsignmentNo.Text, tbToConsignmentNo.Text, bSearchByDate:=False)
        End If
    End Sub
    
    Protected Sub btnShowSerialNosByDate_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Page.Validate("ByDate")
        If Page.IsValid Then
            Call ShowSerialNos(tbFromDate.Text, tbToDate.Text, bSearchByDate:=True)
        End If
    End Sub
    
    Protected Sub GetProducts()
        Dim oDataReader As SqlDataReader
        Dim oConn As New SqlConnection(gsConn)
        Dim sSQL As String
        sSQL = "SELECT ProductCode + '  ' + ISNULL(ProductDate,''), LogisticProductKey FROM LogisticProduct WHERE CustomerKey = 6 AND SerialNumbersFlag = 'Y' ORDER BY ProductCode"
        Dim oCmd As New SqlCommand(sSQL, oConn)
        Dim li As ListItem
        ddlProductCode.Items.Clear()
        li = New ListItem("- please select -", 0)
        ddlProductCode.Items.Add(li)
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            While oDataReader.Read
                li = New ListItem(oDataReader(0), oDataReader(1))
                ddlProductCode.Items.Add(li)
            End While
        Catch ex As SqlException
            WebMsgBox.Show("Error in GetProducts: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub ShowSerialNos(ByVal sFrom As String, ByVal sTo As String, ByVal bSearchByDate As Boolean)
        Dim oDataReader As SqlDataReader
        Dim oConn As New SqlConnection(gsConn)
        Dim sSQL As String
        If bSearchByDate Then
            sSQL = "SELECT snr.ConsignmentKey 'Consignment', LogisticMovementStartDateTime 'Date', snr.RangeFrom 'From', snr.RangeTo 'To', snr.SubTotal 'Items' FROM SerialNumberRange snr INNER JOIN LogisticMovement lm ON snr.LogisticMovementKey = lm.LogisticMovementKey WHERE LogisticMovementStartDateTime >= '" & tbFromDate.Text & "' AND LogisticMovementStartDateTime <= '" & tbToDate.Text & "' AND snr.ProductKey = " & ddlProductCode.SelectedValue & " ORDER BY RangeFrom"
        Else
            sSQL = "SELECT snr.ConsignmentKey 'Consignment', LogisticMovementStartDateTime 'Date', snr.RangeFrom 'From', snr.RangeTo 'To', snr.SubTotal 'Items' FROM SerialNumberRange snr INNER JOIN LogisticMovement lm ON snr.LogisticMovementKey = lm.LogisticMovementKey WHERE snr.ConsignmentKey >= " & tbFromConsignmentNo.Text & " AND snr.ConsignmentKey <= " & tbToConsignmentNo.Text & " AND snr.ProductKey = " & ddlProductCode.SelectedValue & " ORDER BY RangeFrom"
        End If
        Dim oCmd As New SqlCommand(sSQL, oConn)
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            gvSerialNos.DataSource = oDataReader
            gvSerialNos.DataBind()
        Catch ex As SqlException
            WebMsgBox.Show("Error in ShowSerialNos: " & ex.Message)
        Finally
            oConn.Close()
        End Try
        gvSerialNos.Visible = True
        lblSerialNumbers.Visible = True
    End Sub
        
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
            Call ShowSerialNos(tbFromDate.Text, tbToDate.Text, bSearchByDate:=pbSearchTypeDate)
        Else
            Call ShowSerialNos(tbFromConsignmentNo.Text, tbToConsignmentNo.Text, bSearchByDate:=pbSearchTypeDate)
        End If
    End Sub
    
    Protected Sub ddlProductCode_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        If ddl.SelectedIndex > 0 Then
            btnShowSerialNosByConsignmentNo.Enabled = True
            btnShowSerialNosByDate.Enabled = True
        Else
            btnShowSerialNosByConsignmentNo.Enabled = False
            btnShowSerialNosByDate.Enabled = False
        End If
        gvSerialNos.Visible = False
        lblSerialNumbers.Visible = False
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
        <asp:Panel ID="pnlSerialNumbers" runat="server" Width="100%">
            &nbsp; <strong><span style="font-size: 10pt; color: #000080">&nbsp;MAN Serial Numbers</span></strong><table style="width: 100%">
                <tr>
                    <td style="width: 1%; height: 22px;">
                    </td>
                    <td style="width: 20%; height: 22px;" align="right"><asp:Label ID="Label6" runat="server" Text="Product code:" Font-Names="Verdana" Font-Size="XX-Small"
                            Font-Bold="False" /></td>
                    <td style="height: 22px;" colspan="2">
                        <asp:DropDownList ID="ddlProductCode" runat="server" Font-Names="Verdana" Font-Size="XX-Small" AutoPostBack="True" OnSelectedIndexChanged="ddlProductCode_SelectedIndexChanged">
                        </asp:DropDownList></td>
                    <td style="width: 29%; height: 22px;" align="right">
                        </td>
                    <td style="width: 1%; height: 22px;">
                    </td>
                </tr>
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td align="right" style="width: 20%">
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
                <tr>
                    <td>
                    </td>
                    <td>
                        <asp:Button ID="btnShowSerialNosByConsignmentNo" runat="server" Text="show by consignment #"
                            Width="170px" OnClick="btnShowSerialNosByConsignmentNo_Click" Enabled="False" /></td>
                    <td align="left" colspan="3" style="white-space: nowrap">
                        <asp:Label ID="Label2" runat="server" Text="from consignment #" Font-Names="Verdana" Font-Size="XX-Small"
                            Font-Bold="False" />
                        <asp:TextBox ID="tbFromConsignmentNo" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Width="65px" >1</asp:TextBox>
                        <a runat="server" id="aHelpFromOrderNo" visible="false" onmouseover="return escape('')"
                            style="color: gray; cursor: help">&nbsp;?&nbsp;</a>
                        <asp:Label ID="Label3" runat="server" Text="to consignment #" Font-Names="Verdana" Font-Size="XX-Small"
                            Font-Bold="False" />
                        <asp:TextBox ID="tbToConsignmentNo" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Width="65px" >9999999</asp:TextBox>
                        <a runat="server" id="aHelpToOrderNo" visible="false" onmouseover="return escape('')"
                            style="color: gray; cursor: help">&nbsp;?&nbsp;</a>
                        <asp:RangeValidator ID="rvFromOrderNo" runat="server" ControlToValidate="tbFromConsignmentNo"
                            ErrorMessage="invalid from #" MaximumValue="9999999" MinimumValue="1" Type="Integer"
                            ValidationGroup="ByOrderNo"></asp:RangeValidator>
                        <asp:RangeValidator ID="RangeValidator1" runat="server" ControlToValidate="tbToConsignmentNo"
                            ErrorMessage="invalid to #" MaximumValue="9999999" MinimumValue="1" Type="Integer"
                            ValidationGroup="ByOrderNo"></asp:RangeValidator>
                        <asp:RequiredFieldValidator ID="rfvFromOrderNo" runat="server" ControlToValidate="tbFromConsignmentNo"
                            ErrorMessage="from # required" ValidationGroup="ByOrderNo"></asp:RequiredFieldValidator>
                        <asp:RequiredFieldValidator ID="rfvToOrderNo" runat="server" ControlToValidate="tbToConsignmentNo"
                            ErrorMessage="to # required" ValidationGroup="ByOrderNo"></asp:RequiredFieldValidator></td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td>
                        <asp:Button ID="btnShowSerialNosByDate" runat="server" Text="show by date" Width="170px"
                            OnClick="btnShowSerialNosByDate_Click" Enabled="False" /></td>
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
                    <td style="height: 18px">
                    </td>
                    <td style="height: 18px">
                        <asp:Label ID="lblSerialNumbers" runat="server" Text="Serial numbers:" Font-Names="Verdana" Font-Size="XX-Small"
                            Font-Bold="False" /></td>
                    <td style="height: 18px">
                    </td>
                    <td style="height: 18px">
                    </td>
                    <td align="right" style="height: 18px">
                    </td>
                    <td style="height: 18px">
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td colspan="4">
                        <asp:GridView ID="gvSerialNos" runat="server" Width="100%" CellPadding="2" AutoGenerateColumns="True">
                            <EmptyDataTemplate>
                                no serial numbers found for this range
                            </EmptyDataTemplate>
                            <AlternatingRowStyle BackColor="#C0FFC0" />
                        </asp:GridView>
                    </td>
                    <td>
                    </td>
                </tr>
            </table>
        </asp:Panel>
    </form>
</body>
</html>
