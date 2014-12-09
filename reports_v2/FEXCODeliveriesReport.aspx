<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>
<script runat="server">

    '   Product Usage Report - HEADERLESS VERSION FOR COMMON REPORTING FACILITY - CN
    '   Shows total Goods Out for all products over selected period

    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    Dim iDefaultHistoryPeriod As Integer = -1      'last month

    Protected Sub Page_Load(ByVal Source As Object, ByVal E As EventArgs)
        If Not IsNumeric(Session("CustomerKey")) Then
            Server.Transfer("../session_expired.aspx")
        End If
    
        lblReportGeneratedDateTime.Text = "Report generated: " & Now().ToString("dd-MMM-yy HH:mm")
        If Not IsPostBack Then
            Dim dteLastYear As Date = Date.Today.AddMonths(iDefaultHistoryPeriod)
            tbFromDate.Text = dteLastYear.ToString("dd-MMM-yy")
            tbToDate.Text = Now.ToString("dd-MMM-yy")
            
            Dim sYear As String = Year(Now)
            Dim i As Integer
            For i = CInt(sYear) To CInt(sYear) - 6 Step -1
                ddlToYear.Items.Add(i.ToString)
                ddlFromYear.Items.Add(i.ToString)
            Next
            ddlToYear.Items(0).Selected = True
            ddlFromYear.Items(1).Selected = True
            
            ddlToMonth.Items(Month(Now) - 1).Selected = True
            ddlFromMonth.Items(Month(Now) - 1).Selected = True

            ddlToDay.Items(Day(Now) - 1).Selected = True
            ddlFromDay.Items(Day(Now) - 1).Selected = True

        End If
    End Sub
    
    Protected Sub btnReselectDateRange_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ReselectDateRange()
    End Sub
    
    Protected Sub ReselectDateRange()
        pnlData.Visible = False
        btnExport1.Visible = True
        btnExport2.Visible = True
        btnReselectDateRange1.Visible = False
        btnReselectDateRange2.Visible = False
        tbFromDate.Enabled = True
        tbToDate.Enabled = True
        ddlFromDay.Enabled = True
        ddlFromMonth.Enabled = True
        ddlFromYear.Enabled = True
        ddlToDay.Enabled = True
        ddlToMonth.Enabled = True
        ddlToYear.Enabled = True
        spnDateExample1.Visible = True
        spnDateExample2.Visible = True
        imgCalendarButton1.Visible = True
        imgCalendarButton2.Visible = True
    End Sub

    Protected Sub ShowDataPanel()
        pnlData.Visible = True
    End Sub
    
    Protected Sub GetDateRange()
        btnExport1.Visible = False
        btnExport2.Visible = False
        btnReselectDateRange1.Visible = True
        btnReselectDateRange2.Visible = True
        tbFromDate.Enabled = False
        tbToDate.Enabled = False
        ddlFromDay.Enabled = False
        ddlFromMonth.Enabled = False
        ddlFromYear.Enabled = False
        ddlToDay.Enabled = False
        ddlToMonth.Enabled = False
        ddlToYear.Enabled = False
        
        If CalendarInterface.Visible Then
            psToDate = tbToDate.Text
            psFromDate = tbFromDate.Text
        Else
            psFromDate = ddlFromDay.SelectedItem.Text & "-" & ddlFromMonth.SelectedItem.Text & "-" & ddlFromYear.SelectedItem.Text
            psToDate = ddlToDay.SelectedItem.Text & "-" & ddlToMonth.SelectedItem.Text & "-" & ddlToYear.SelectedItem.Text
            tbFromDate.Text = psFromDate
            tbToDate.Text = psToDate
        End If
        
    End Sub
    
    Protected Sub btnExport_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        lblFromErrorMessage.Text = ""
        lblToErrorMessage.Text = ""
        
        If CalendarInterface.Visible Then
            Page.Validate("CalendarInterface")
        Else
            Dim sDate = ddlToDay.SelectedItem.Text & ddlToMonth.SelectedItem.Text & ddlToYear.SelectedItem.Text
            If IsDate(sDate) Then
                
            End If
        End If

        If (CalendarInterface.Visible _
          And Page.IsValid) _
         Or _
          (DropdownInterface.Visible _
          And IsDate(ddlFromDay.SelectedItem.Text & ddlFromMonth.SelectedItem.Text & ddlFromYear.SelectedItem.Text) _
          And IsDate(ddlToDay.SelectedItem.Text & ddlToMonth.SelectedItem.Text & ddlToYear.SelectedItem.Text)) Then

            Call GetDateRange()
            
            spnDateExample1.Visible = False
            spnDateExample2.Visible = False
            imgCalendarButton1.Visible = False
            imgCalendarButton2.Visible = False
            
            'lblReportGeneratedDateTime.Visible = True
            Call ExportCSVData(ConvertDataTableToCSVString(GetDeliveries))
        Else
            If DropdownInterface.Visible Then
                If Not IsDate(ddlFromDay.SelectedItem.Text & ddlFromMonth.SelectedItem.Text & ddlFromYear.SelectedItem.Text) Then
                    lblFromErrorMessage.Text = "Invalid date"
                End If
                If Not IsDate(ddlToDay.SelectedItem.Text & ddlToMonth.SelectedItem.Text & ddlToYear.SelectedItem.Text) Then
                    lblToErrorMessage.Text = "Invalid date"
                End If
            End If
        End If
    End Sub
    
    Protected Function GetDeliveries() As DataTable
        GetDeliveries = Nothing
        Dim sbSQL As New StringBuilder
        ' sbSQL.Append("SELECT  ProductCode + ' - ' + ISNULL(ProductDate,'') + ' - ' + ProductDescription 'Product', CONVERT(VARCHAR(11), ArrivalDate, 106) 'Arrival Date', cfd.Quantity, cfd.Value, cfd.Supplier, cfd.Notes FROM ClientData_FEXCO_Deliveries cfd ")
        sbSQL.Append("SELECT  ProductCode 'Product Code', ISNULL(ProductDate,'') 'Value / Date', ProductDescription 'Description', ProductCategory 'Product Category', CONVERT(VARCHAR(11), ArrivalDate, 106) 'Arrival Date', cfd.Quantity, cfd.Value, cfd.Supplier, cfd.Notes FROM ClientData_FEXCO_Deliveries cfd ")
        sbSQL.Append("INNER JOIN LogisticProduct lp ")
        sbSQL.Append("ON cfd.LogisticProductKey = lp.LogisticProductKey ")
        sbSQL.Append("WHERE ArrivalDate >= '" & psFromDate & "' AND ArrivalDate <= '" & psToDate & "' ")
        sbSQL.Append("ORDER BY ArrivalDate ASC")
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter(sbSQL.ToString, oConn)
        Dim oDataTable As New DataTable
        Try
            oAdapter.Fill(oDataTable)
            GetDeliveries = oDataTable
        Catch ex As Exception
            WebMsgBox.Show("Error in GetDeliveries: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Function
    
    Public Function ConvertDataTableToCSVString(ByVal oDataTable As DataTable) As String
        Dim sbResult As New StringBuilder
        Dim oDataColumn As DataColumn
        Dim oDataRow As DataRow

        For Each oDataColumn In oDataTable.Columns         ' column headings in line 1
            sbResult.Append(oDataColumn.ColumnName)
            sbResult.Append(",")
        Next
        If sbResult.Length > 1 Then
            sbResult.Length = sbResult.Length - 1
        End If
        sbResult.Append(Environment.NewLine)
        Dim s2 As String
        For Each oDataRow In oDataTable.Rows
            For Each s As Object In oDataRow.ItemArray
                Try
                    s2 = s
                Catch
                    s2 = String.Empty
                End Try
                s2 = s2.Replace(Environment.NewLine, " ")
                sbResult.Append(s2.Replace(",", " "))
                sbResult.Append(",")
            Next
            sbResult.Length = sbResult.Length - 1
            sbResult.Append(Environment.NewLine)
        Next oDataRow

        If Not sbResult Is Nothing Then
            Return sbResult.ToString()
        Else
            Return String.Empty
        End If
    End Function
   
    Private Sub ExportCSVData(ByVal sCSVString As String)
        Response.Clear()
        Response.AddHeader("Content-Disposition", "attachment;filename=" & "WesternUnionDeliveriesReport.csv")
        'Response.ContentType = "application/vnd.ms-excel"
        Response.ContentType = "text/csv"
   
        Dim eEncoding As Encoding = Encoding.GetEncoding("Windows-1252")
        Dim eUnicode As Encoding = Encoding.Unicode
        Dim byUnicodeBytes As Byte() = eUnicode.GetBytes(sCSVString)
        Dim byEncodedBytes As Byte() = Encoding.Convert(eUnicode, eEncoding, byUnicodeBytes)
        Response.BinaryWrite(byEncodedBytes)
        Response.End()
        ' Response.Flush()
    End Sub

    Protected Sub lnkbtnToggleSelectionStyle_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If CalendarInterface.Visible = True Then
            CalendarInterface.Visible = False
            DropdownInterface.Visible = True
        Else
            CalendarInterface.Visible = True
            DropdownInterface.Visible = False
        End If
        lblFromErrorMessage.Text = ""
        lblToErrorMessage.Text = ""
    End Sub

    Property psToDate() As String
        Get
            Dim o As Object = ViewState("ToDate")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("ToDate") = Value
        End Set
    End Property
    
    Property psFromDate() As String
        Get
            Dim o As Object = ViewState("FromDate")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("FromDate") = Value
        End Set
    End Property
    
</script>
<html xmlns="http://www.w3.org/1999/xhtml" >
<head>
    <title>Product Usage Report</title>
    <link rel="stylesheet" type="text/css" href="../Reports.css" />
</head>
<body>
    <form id="form1" runat="server">
        <table>
            <tr runat="server" visible="true">
                <td colspan="4" style="white-space:nowrap">
                  <asp:Label ID="lblPageHeading"
                             runat="server"
                             forecolor="Silver"
                             font-size="Small"
                             font-bold="True"
                             font-names="Arial">Western Union Retail Services GB Deliveries</asp:Label><br /><br />
                </td>
            </tr>
            <tr runat="server" visible="true" id="CalendarInterface">
                <td style="width: 265px; white-space:nowrap" >
                    <span class="informational dark">From:</span>
                        <asp:TextBox ID="tbFromDate"
                                     font-names="Verdana"
                                     font-size="XX-Small"
                                     Width="90"
                                     runat="server">
                          </asp:TextBox>
                         <a id="imgCalendarButton1" runat="server" visible="true" href="javascript:;"
                            onclick="window.open('../PopupCalendar4.aspx?textbox=tbFromDate','cal','width=300,height=305,left=270,top=180')">
                            <img id="Img1" alt=""
                                 src="../images/SmallCalendar.gif"
                                 runat="server"
                                 border="0"
                              IE:visible="true"
                                 visible="false"
                               /></a>
                    <span id="spnDateExample1" runat="server" visible="true" class="informational light">(eg&nbsp;12-Jan-2005)</span>
                    </td>
                <td style="width: 265px; white-space:nowrap" >
                    <span class="informational dark">To:</span>
                        <asp:TextBox ID="tbToDate"
                                     font-names="Verdana"
                                     font-size="XX-Small"
                                     Width="90"
                                     runat="server">
                          </asp:TextBox>
                           <a id="imgCalendarButton2" runat="server" visible="true" href="javascript:;"
                            onclick="window.open('../PopupCalendar4.aspx?textbox=tbToDate','cal','width=300,height=305,left=270,top=180')"><img id="Img2"
                                 src="../images/SmallCalendar.gif" alt=""
                                 runat="server"
                                 border="0"
                              IE:visible="true"
                                 visible="false"
                               ></a>
                    <span id="spnDateExample2" runat="server" visible="true" class="informational light">(eg&nbsp;12-Jan-2006)</span>
                   </td>
                <td style="width: 253px"><asp:Button ID="btnExport1" runat="server" OnClick="btnExport_Click" Text="export deliveries data to excel"
                Width="300px" Font-Names="Verdana" Font-Size="XX-Small" />
                    <asp:Button ID="btnReselectDateRange1"
                     runat="server"
                     Text="select another period"
                     Visible="false"
                      OnClick="btnReselectDateRange_Click" Font-Names="Verdana" Font-Size="XX-Small" />
                      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                </td>
                <td style="width: 169px">
                    <asp:LinkButton ID="lnkbtnToggleSelectionStyle1"
                                      runat="server"
                                      OnClick="lnkbtnToggleSelectionStyle_Click"
                                      ToolTip="toggles between calendar interface and dropdown interface">change&nbsp;selection&nbsp;style</asp:LinkButton></td>
            </tr>
            <tr runat="server" visible="false" id="DropdownInterface">
                <td style="width: 265px; height: 26px;">
                    <span class="informational dark">From:</span>
                    &nbsp;<asp:DropDownList ID="ddlFromDay" runat="server" Font-Names="Verdana" Font-Size="XX-Small">
                        <asp:ListItem>01</asp:ListItem><asp:ListItem>02</asp:ListItem>
                        <asp:ListItem>03</asp:ListItem>
                        <asp:ListItem>04</asp:ListItem>
                        <asp:ListItem>05</asp:ListItem>
                        <asp:ListItem>06</asp:ListItem>
                        <asp:ListItem>07</asp:ListItem>
                        <asp:ListItem>08</asp:ListItem>
                        <asp:ListItem>09</asp:ListItem>
                        <asp:ListItem>10</asp:ListItem>
                        <asp:ListItem>11</asp:ListItem>
                        <asp:ListItem>12</asp:ListItem>
                        <asp:ListItem>13</asp:ListItem>
                        <asp:ListItem>14</asp:ListItem>
                        <asp:ListItem>15</asp:ListItem>
                        <asp:ListItem>16</asp:ListItem>
                        <asp:ListItem>17</asp:ListItem>
                        <asp:ListItem>18</asp:ListItem>
                        <asp:ListItem>19</asp:ListItem>
                        <asp:ListItem>20</asp:ListItem>
                        <asp:ListItem>21</asp:ListItem>
                        <asp:ListItem>22</asp:ListItem>
                        <asp:ListItem>23</asp:ListItem>
                        <asp:ListItem>24</asp:ListItem>
                        <asp:ListItem>25</asp:ListItem>
                        <asp:ListItem>26</asp:ListItem>
                        <asp:ListItem>27</asp:ListItem>
                        <asp:ListItem>28</asp:ListItem>
                        <asp:ListItem>29</asp:ListItem>
                        <asp:ListItem>30</asp:ListItem>
                        <asp:ListItem>31</asp:ListItem>
                        </asp:DropDownList>&nbsp;<asp:DropDownList ID="ddlFromMonth" runat="server" Font-Names="Verdana" Font-Size="XX-Small">
                        <asp:ListItem>Jan</asp:ListItem>
                        <asp:ListItem>Feb</asp:ListItem>
                        <asp:ListItem>Mar</asp:ListItem>
                        <asp:ListItem>Apr</asp:ListItem>
                        <asp:ListItem>May</asp:ListItem>
                        <asp:ListItem>Jun</asp:ListItem>
                        <asp:ListItem>Jul</asp:ListItem>
                        <asp:ListItem>Aug</asp:ListItem>
                        <asp:ListItem>Sep</asp:ListItem>
                        <asp:ListItem>Oct</asp:ListItem>
                        <asp:ListItem>Nov</asp:ListItem>
                        <asp:ListItem>Dec</asp:ListItem>
                    </asp:DropDownList>&nbsp;<asp:DropDownList ID="ddlFromYear" runat="server" Font-Names="Verdana" Font-Size="XX-Small">
                    </asp:DropDownList>&nbsp;</td>
                <td style="width: 265px; height: 26px;">
                    <span class="informational dark">To:</span>
                    &nbsp;<asp:DropDownList ID="ddlToDay" runat="server" Font-Names="Verdana" Font-Size="XX-Small">
                        <asp:ListItem>01</asp:ListItem>
                        <asp:ListItem>02</asp:ListItem>
                        <asp:ListItem>03</asp:ListItem>
                        <asp:ListItem>04</asp:ListItem>
                        <asp:ListItem>05</asp:ListItem>
                        <asp:ListItem>06</asp:ListItem>
                        <asp:ListItem>07</asp:ListItem>
                        <asp:ListItem>08</asp:ListItem>
                        <asp:ListItem>09</asp:ListItem>
                        <asp:ListItem>10</asp:ListItem>
                        <asp:ListItem>11</asp:ListItem>
                        <asp:ListItem>12</asp:ListItem>
                        <asp:ListItem>13</asp:ListItem>
                        <asp:ListItem>14</asp:ListItem>
                        <asp:ListItem>15</asp:ListItem>
                        <asp:ListItem>16</asp:ListItem>
                        <asp:ListItem>17</asp:ListItem>
                        <asp:ListItem>18</asp:ListItem>
                        <asp:ListItem>19</asp:ListItem>
                        <asp:ListItem>20</asp:ListItem>
                        <asp:ListItem>21</asp:ListItem>
                        <asp:ListItem>22</asp:ListItem>
                        <asp:ListItem>23</asp:ListItem>
                        <asp:ListItem>24</asp:ListItem>
                        <asp:ListItem>25</asp:ListItem>
                        <asp:ListItem>26</asp:ListItem>
                        <asp:ListItem>27</asp:ListItem>
                        <asp:ListItem>28</asp:ListItem>
                        <asp:ListItem>29</asp:ListItem>
                        <asp:ListItem>30</asp:ListItem>
                        <asp:ListItem>31</asp:ListItem>
                    </asp:DropDownList>&nbsp;<asp:DropDownList ID="ddlToMonth" runat="server" Font-Names="Verdana" Font-Size="XX-Small">
                        <asp:ListItem>Jan</asp:ListItem>
                        <asp:ListItem>Feb</asp:ListItem>
                        <asp:ListItem>Mar</asp:ListItem>
                        <asp:ListItem>Apr</asp:ListItem>
                        <asp:ListItem>May</asp:ListItem>
                        <asp:ListItem>Jun</asp:ListItem>
                        <asp:ListItem>Jul</asp:ListItem>
                        <asp:ListItem>Aug</asp:ListItem>
                        <asp:ListItem>Sep</asp:ListItem>
                        <asp:ListItem>Oct</asp:ListItem>
                        <asp:ListItem>Nov</asp:ListItem>
                        <asp:ListItem>Dec</asp:ListItem>
                    </asp:DropDownList>&nbsp;<asp:DropDownList ID="ddlToYear" runat="server" Font-Names="Verdana" Font-Size="XX-Small">
                    </asp:DropDownList>&nbsp;</td>
                <td style="width: 253px; height: 26px;"><asp:Button ID="btnExport2" runat="server" OnClick="btnExport_Click" Text="export deliveries data to excel"
                Width="300px" Font-Names="Verdana" Font-Size="XX-Small" />
                    <asp:Button ID="btnReselectDateRange2"
                     runat="server"
                     Text="select another period"
                     Visible="false"
                      OnClick="btnReselectDateRange_Click" Font-Names="Verdana" Font-Size="XX-Small" />
                      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                      </td>
                <td style="width: 169px; height: 26px;">
                    <asp:LinkButton ID="lnkbtnToggleSelectionStyle2"
                                      runat="server"
                                      OnClick="lnkbtnToggleSelectionStyle_Click"
                                      ToolTip="toggles between calendar interface and dropdown interface">change&nbsp;selection&nbsp;style</asp:LinkButton></td>
            </tr>
            <tr id="Tr1" runat="server" visible="true">
                <td colspan="4" style="white-space:nowrap">
                  <br />
                    &nbsp;&nbsp;
                </td>
            </tr>
            <tr runat="server" visible="true" id="DateValidationMessages">
                <td style="width: 265px">
                    <asp:RegularExpressionValidator ID="revFromDate" runat="server" ControlToValidate="tbFromDate" ErrorMessage=" - invalid format for expiry date - use dd-mmm-yy"
                        Font-Names="Verdana" Font-Size="XX-Small" ValidationExpression="^\d\d-(jan|Jan|JAN|feb|Feb|FEB|mar|Mar|MAR|apr|Apr|APR|may|May|MAY|jun|Jun|JUN|jul|Jul|JUL|aug|Aug|AUG|sep|Sep|SEP|oct|Oct|OCT|nov|Nov|NOV|dec|Dec|DEC)-\d(\d+)" SetFocusOnError="True" ValidationGroup="CalendarInterface"></asp:RegularExpressionValidator>
                    <asp:RangeValidator ID="rvFromDate" runat="server" ControlToValidate="tbFromDate"
                        CultureInvariantValues="True" ErrorMessage=" - expiry year before 2000, after 2020, or not a valid date!"
                        Font-Names="Verdana" Font-Size="XX-Small" MaximumValue="2019/1/1" MinimumValue="2000/1/1" ValidationGroup="CalendarInterface" EnableClientScript="False" Type="Date"></asp:RangeValidator>
                    <asp:Label ID="lblFromErrorMessage" runat="server" Font-Names="Verdana"
                        Font-Size="XX-Small" ForeColor="Red"></asp:Label></td>
                <td style="width: 265px">
                    <asp:RegularExpressionValidator ID="RegularevToDate" runat="server" ControlToValidate="tbToDate" ErrorMessage=" - invalid format for expiry date - use dd-mmm-yy"
                        Font-Names="Verdana" Font-Size="XX-Small" ValidationExpression="^\d\d-(jan|Jan|JAN|feb|Feb|FEB|mar|Mar|MAR|apr|Apr|APR|may|May|MAY|jun|Jun|JUN|jul|Jul|JUL|aug|Aug|AUG|sep|Sep|SEP|oct|Oct|OCT|nov|Nov|NOV|dec|Dec|DEC)-\d(\d+)" ValidationGroup="CalendarInterface"></asp:RegularExpressionValidator><asp:RangeValidator
                            ID="rvToDate" runat="server" ControlToValidate="tbToDate" CultureInvariantValues="True" ErrorMessage=" - expiry year before 2000, after 2020, or not a valid date!"
                            Font-Names="Verdana" Font-Size="XX-Small" MaximumValue="2019/1/1" MinimumValue="2000/1/1" ValidationGroup="CalendarInterface" EnableClientScript="False" Type="Date"></asp:RangeValidator>
                    <asp:Label ID="lblToErrorMessage" runat="server" Font-Names="Verdana"
                        Font-Size="XX-Small" ForeColor="Red"></asp:Label></td>
                <td style="width: 253px">
                </td>
                <td style="width: 169px">
                </td>
            </tr>
        </table>
        <!-- Panel: Data -->
        <asp:Panel id="pnlData" runat="server">
            &nbsp;<br />
            <asp:Label ID="lblReportGeneratedDateTime" runat="server" Text="" font-size="XX-Small" font-names="Verdana, Sans-Serif" forecolor="Green" Visible="false"></asp:Label>
        </asp:Panel>
        <br />
        <asp:Label id="lblError" runat="server" Font-Names="Arial" Font-Size="XX-Small" ForeColor="red"></asp:Label>
    </form>
</body>
</html>
