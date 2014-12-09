<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<script runat="server">

    Dim gsConn As String = ConfigurationSettings.AppSettings("AIMSRootConnectionString")

    Const CUSTOMER_HYSTER As Integer = 77
    Const CUSTOMER_YALE As Int32 = 680

    Const INTERVAL_SHORT_MESSAGE As Int32 = 3000
    Const INTERVAL_SHORT_CONFIRMATION_MESSAGE As Int32 = 3000
    Const INTERVAL_SHORT_ERROR_MESSAGE As Int32 = 3001
    Const INTERVAL_LONG_MESSAGE As Int32 = 6000
    Const INTERVAL_PERMANENT_MESSAGE As Int32 = -1

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsNumeric(Session("CustomerKey")) Then
            Server.Transfer("session_expired.aspx")
            Exit Sub
        End If
        If Not IsPostBack Then
            Call SetTitle()
            tbConsignment.Attributes.Add("onkeypress", "return clickButton(event,'" + btnShow.ClientID + "')")
            tbConsignment.Focus()
        End If
    End Sub
    
    Protected Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim sm As New ScriptManager
        sm.ID = "ScriptMgr"
        Try
            PlaceHolderForScriptManager.Controls.Add(sm)
        Catch ex As Exception
        End Try
    End Sub
    
    Protected Sub SetTitle()
        Dim sTitle As String = Session("SiteTitle")
        If sTitle <> String.Empty Then
            sTitle += " - "
        End If
        Page.Header.Title = sTitle & "Hyster Yale Invoice Editor"
    End Sub
   
    Protected Sub btnGo_Click(sender As Object, e As System.EventArgs)
        tbConsignment.Text = tbConsignment.Text.Trim
        If tbConsignment.Text.Length > 0 Then
            Call RetrieveConsignment()
        End If
    End Sub
    
    Protected Sub RetrieveConsignment()
        Dim nCustomerKey As Int32
        trRow1.Visible = False
        trRow2.Visible = False
        Dim sSQL As String = "SELECT [key] ConsignmentKey, CustomerKey, CreatedOn FROM Consignment WHERE AWB = '" & tbConsignment.Text.Replace("'", "''") & "'"
        Dim dtConsignment As DataTable
        dtConsignment = ExecuteQueryToDataTable(sSQL)
        If dtConsignment.Rows.Count > 0 Then
            If dtConsignment.Rows.Count > 1 Then
                WebMsgBox.Show("Error: there appear to be multiple consignments with AWB " & tbConsignment.Text & ". Please inform development.")
                Exit Sub
            Else
                nCustomerKey = dtConsignment.Rows(0).Item("CustomerKey")
                If nCustomerKey = CUSTOMER_HYSTER Then
                    sSQL = "SELECT [key], DealerOrderDate, ConsignmentNo, BookedBy, DealerCompanyName, ProductCode, ProductDescription, Quantity, UnitPrice, CarriageCost FROM HysterInvoicingData WHERE ConsignmentNo = '" & tbConsignment.Text.Replace("'", "''") & "' ORDER BY [key]"
                    psTableName = "HysterInvoicingData"
                ElseIf nCustomerKey = CUSTOMER_YALE Then
                    sSQL = "SELECT [key], DealerOrderDate, ConsignmentNo, BookedBy, DealerCompanyName, ProductCode, ProductDescription, Quantity, UnitPrice, CarriageCost FROM YaleInvoicingData WHERE ConsignmentNo = '" & tbConsignment.Text.Replace("'", "''") & "' ORDER BY [key]"
                    psTableName = "YaleInvoicingData"
                Else
                    WebMsgBox.Show("The specified AWB is neither a Hyster nor a Yale consignment.")
                    Exit Sub
                End If
                'pnConsignmentKey = dtConsignment.Rows(0).Item("ConsignmentKey")
                psConsignmentNo = tbConsignment.Text
                Dim dt As DateTime = dtConsignment.Rows(0).Item("CreatedOn")
                psConsignmentDate = dt.ToString("dd-MMM-yyyy hh:mm")
                dtConsignment = ExecuteQueryToDataTable(sSQL)
                If dtConsignment.Rows.Count > 0 Then
                    lblConsignmentData.Text = "Date: " & psConsignmentDate & ", Booked By: " & dtConsignment.Rows(0).Item("BookedBy") & ", Dealer: " & dtConsignment.Rows(0).Item("DealerCompanyName")
                    tbCarriageCost.Text = Format(dtConsignment.Rows(0).Item("CarriageCost"), "0.00")
                    psInitialCarriageCost = tbCarriageCost.Text
                    gvConsignment.DataSource = dtConsignment
                    gvConsignment.DataBind()
                    trRow1.Visible = True
                    trRow2.Visible = True
                Else
                    WebMsgBox.Show("Could not find an invoice record for this consignment. It may not yet have been added to the invoice file (eg because the carriage cost is not yet available from METACS).")
                    Exit Sub
                End If
            End If
        Else
            WebMsgBox.Show("Cannot find a consignment with AWB " & tbConsignment.Text & ".")
            Exit Sub
        End If
    End Sub
    
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
    
    Property psConsignmentDate() As String
        Get
            Dim o As Object = ViewState("HYIE_ConsignmentDate")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("HYIE_ConsignmentDate") = Value
        End Set
    End Property
    
    Property psTableName() As String
        Get
            Dim o As Object = ViewState("HYIE_TableName")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("HYIE_TableName") = Value
        End Set
    End Property
    
    Property psConsignmentNo() As String
        Get
            Dim o As Object = ViewState("HYIE_ConsignmentNo")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("HYIE_ConsignmentNo") = Value
        End Set
    End Property
    
    Property psInitialCarriageCost() As String
        Get
            Dim o As Object = ViewState("HYIE_InitialCarriageCost")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("HYIE_InitialCarriageCost") = Value
        End Set
    End Property
    
    'Property pnConsignmentKey() As Int32
    '    Get
    '        Dim o As Object = ViewState("HYIE_ConsignmentKey")
    '        If o Is Nothing Then
    '            Return 0
    '        End If
    '        Return CInt(o)
    '    End Get
    '    Set(ByVal Value As Int32)
    '        ViewState("HYIE_ConsignmentKey") = Value
    '    End Set
    'End Property
    
    Protected Sub btnUpdateCarriageCost_Click(sender As Object, e As System.EventArgs)
        tbCarriageCost.Text = tbCarriageCost.Text.Trim
        If IsNumeric(tbCarriageCost.Text) Then
            Dim sSQL As String = "UPDATE " & psTableName & " SET CarriageCost = " & tbCarriageCost.Text & " WHERE ConsignmentNo = '" & psConsignmentNo & "'"
            Call ExecuteQueryToDataTable(sSQL)
            sSQL = "INSERT INTO ClientData_HysterYale_InvoiceEditAuditTrail (Action, CreatedOn, CreatedBy) VALUES ('"
            sSQL &= "CARRCOST: "
            sSQL &= psConsignmentNo
            sSQL &= ": was "
            sSQL &= psInitialCarriageCost
            sSQL &= ", now "
            sSQL &= tbCarriageCost.Text
            sSQL &= "', GETDATE(), "
            sSQL &= Session("UserKey")
            sSQL &= ")"
            Call ExecuteQueryToDataTable(sSQL)
            Call ShowMessage("Saved new carriage cost value.")
        Else
            WebMsgBox.Show("Carriage cost must be numeric.")
        End If
    End Sub
    
    Protected Sub btnUpdateQty_Click(sender As Object, e As System.EventArgs)
        Dim btn As Button = sender
        Dim nID As Int32 = btn.CommandArgument
        For Each gvr As GridViewRow In gvConsignment.Rows
            If gvr.RowType = DataControlRowType.DataRow Then
                Dim btnUpdateQty As Button = gvr.Cells(3).FindControl("btnUpdateQty")
                If btnUpdateQty.CommandArgument = nID Then
                    Dim tbQty As TextBox = gvr.Cells(3).FindControl("tbQty")
                    Dim sQty As String = tbQty.Text
                    If IsNumeric(sQty) Then
                        Dim nQty = CInt(sQty)
                        If nQty >= 0 Then
                            Dim hidInitialValue As HiddenField = gvr.Cells(3).FindControl("hidInitialValue")
                            Dim sSQL As String = "UPDATE " & psTableName & " SET Quantity = " & sQty & " WHERE [key] = " & nID
                            Call ExecuteQueryToDataTable(sSQL)
                            sSQL = "INSERT INTO ClientData_HysterYale_InvoiceEditAuditTrail (Action, CreatedOn, CreatedBy) VALUES ('"
                            sSQL &= "QTY: "
                            sSQL &= nID.ToString
                            sSQL &= " = "
                            sSQL &= sQty
                            sSQL &= ", "
                            sSQL &= "was "
                            sSQL &= hidInitialValue.Value
                            sSQL &= "', "
                            sSQL &= "GETDATE(), "
                            sSQL &= Session("UserKey")
                            sSQL &= ")"
                            Call ExecuteQueryToDataTable(sSQL)
                            Call ShowMessage("Saved new item quantity value.")
                        Else
                            WebMsgBox.Show("Invalid quantity - quantity cannot be negative.")
                        End If
                    Else
                        WebMsgBox.Show("Invalid quantity.")
                    End If
                End If
            End If
        Next
    End Sub
    
    Protected Sub tmrNotificationTimer_Tick(ByVal sender As Object, ByVal e As System.EventArgs)
        lblMessage.Visible = False
        'lblMessage.Text = ""
        tmrNotificationTimer.Enabled = False
    End Sub

    Protected Sub SetNotificationTimer(ByVal c As Control)
        c.Visible = True
        tmrNotificationTimer.Enabled = True
    End Sub

    Protected Sub ShowMessage(ByVal sMessage As String, Optional ByVal nInterval As Int32 = INTERVAL_SHORT_CONFIRMATION_MESSAGE)
        If lblMessage.Text.Contains("SYSTEM ERROR") Then
            Exit Sub
        End If
        lblMessage.Text = "&nbsp;" & sMessage & "&nbsp;"
        lblMessage.Visible = True
        lblMessage.BackColor = Drawing.Color.Red
        If nInterval = INTERVAL_SHORT_CONFIRMATION_MESSAGE Then
            lblMessage.BackColor = Drawing.Color.LawnGreen
        End If
        If nInterval <> INTERVAL_PERMANENT_MESSAGE Then
            tmrNotificationTimer.Interval = nInterval
            tmrNotificationTimer.Enabled = True
        End If
    End Sub

</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="Form1" runat="Server">
    <main:Header ID="ctlHeader" runat="server" />
    <asp:PlaceHolder ID="PlaceHolderForScriptManager" runat="server" />
        <asp:Timer ID="tmrNotificationTimer" runat="server" OnTick="tmrNotificationTimer_Tick"
            Interval="2000" Enabled="False" />
    <table width="100%">
        <tr>
            <td style="width: 1%;">
                &nbsp;
            </td>
            <td style="width: 98%;">
                <asp:Label ID="lblLegendTitle" runat="server" Font-Size="X-Small" Font-Names="Verdana"
                    Font-Bold="True" ForeColor="Gray">Hyster / Yale Invoice Editor</asp:Label>
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <asp:Label ID="lblMessage" runat="server" BackColor="Red" Font-Bold="True" Font-Names="Verdana"
                Font-Size="Small" ForeColor="White" Text="MESSAGE" Visible="False" />
            </td>
            <td style="width: 1%;">
                &nbsp;
            </td>
        </tr>
        <tr>
            <td>
            </td>
            <td>
                <asp:Label ID="lblLegendTitle1" runat="server" Font-Size="X-Small" Font-Names="Verdana"
                    Font-Bold="True" ForeColor="Gray">Consignment:</asp:Label>
                &nbsp; &nbsp;<asp:TextBox ID="tbConsignment" runat="server"></asp:TextBox>
                &nbsp;&nbsp;&nbsp;&nbsp;
                <asp:Button ID="btnShow" runat="server" OnClick="btnGo_Click" Text="find" Width="130px" />
                &nbsp;
            </td>
            <td>
            </td>
        </tr>
        <tr id="trRow1" runat="server" visible="false">
            <td>
            </td>
            <td>
                <asp:Label ID="lblConsignmentData" Font-Names="Verdana" Font-Size="XX-Small" runat="server"
                    Font-Bold="True" />
                &nbsp;&nbsp;&nbsp;
                <asp:Label ID="lblLegendCarriageCost" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                    Text="Carriage Cost (£):" />
                &nbsp;<asp:TextBox ID="tbCarriageCost" Font-Names="Verdana" Font-Size="XX-Small"
                    runat="server" />
                &nbsp;<asp:Button ID="btnUpdateCarriageCost" runat="server" Text="update" OnClick="btnUpdateCarriageCost_Click" />
            </td>
            <td>
            </td>
        </tr>
        <tr id="trRow2" runat="server" visible="false">
            <td>
            </td>
            <td>
                <asp:GridView ID="gvConsignment" runat="server" CellPadding="2" Font-Names="Verdana"
                    Font-Size="XX-Small" Width="100%" AutoGenerateColumns="False">
                    <Columns>
                        <asp:BoundField DataField="ProductCode" HeaderText="Product Code" ReadOnly="True"
                            SortExpression="ProductCode" />
                        <asp:BoundField DataField="ProductDescription" HeaderText="Description" ReadOnly="True"
                            SortExpression="ProductDescription" />
                        <asp:TemplateField HeaderText="Qty">
                            <ItemTemplate>
                                <asp:TextBox ID="tbQty" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                    Text='<%# Container.DataItem("Quantity")%>' Width="50px" />
                                &nbsp;<asp:Button ID="btnUpdateQty" runat="server" Text="update"  
                                    CommandArgument='<%# Container.DataItem("key")%>' 
                                    onclick="btnUpdateQty_Click"/>
                                <asp:HiddenField ID="hidInitialValue" runat="server" Value='<%# Container.DataItem("Quantity")%>' />
                            </ItemTemplate>
                            <ItemStyle HorizontalAlign="Center" />
                        </asp:TemplateField>
                        <asp:BoundField DataField="UnitPrice" HeaderText="Unit Price (£/€)" ReadOnly="True" SortExpression="UnitPrice" DataFormatString="{0:#,##0.00}">
                            <ItemStyle HorizontalAlign="Center" />
                        </asp:BoundField>
                    </Columns>
                </asp:GridView>
            </td>
            <td>
            </td>
        </tr>
        <tr>
            <td>
            </td>
            <td>
            </td>
            <td>
            </td>
        </tr>
        <tr>
            <td>
            </td>
            <td>
            </td>
            <td>
            </td>
        </tr>
    </table>
    </form>
</body>
</html>
