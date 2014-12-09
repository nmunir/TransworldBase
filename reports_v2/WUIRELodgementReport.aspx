<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>
<script runat="server">

    Dim iDefaultHistoryPeriodDays As Integer = -14    'last 2 weeks
    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString()
    
    Sub Page_Load(ByVal Source As Object, ByVal E As EventArgs)
        Session("CustomerKey") = 0
        If Not IsNumeric(Session("CustomerKey")) Then
            Server.Transfer("../session_expired.aspx")
        End If
    
        lblReportGeneratedDateTime.Text = "Report generated: " & Now().ToString("dd-MMM-yy HH:mm")
        If Not IsPostBack Then
            Dim dteLastYear As Date = Date.Today.AddDays(iDefaultHistoryPeriodDays)
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
        btnRunReport1.Visible = True
        btnRunReport2.Visible = True
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

    Protected Sub GetDateRange()
        btnRunReport1.Visible = False
        btnRunReport2.Visible = False
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
    
    Protected Sub btnRunReport_Click(ByVal s As Object, ByVal e As EventArgs)
        lblFromErrorMessage.Text = ""
        lblToErrorMessage.Text = ""
        
        If CalendarInterface.Visible Then
            Page.Validate("CalendarInterface")
        Else
            Dim sDate = ddlToDay.SelectedItem.Text & ddlToMonth.SelectedItem.Text & ddlToYear.SelectedItem.Text
            If IsDate(sDate) Then
                
            End If
        End If

        If (CalendarInterface.Visible And Page.IsValid) _
         Or _
          (DropdownInterface.Visible _
          And IsDate(ddlFromDay.SelectedItem.Text & ddlFromMonth.SelectedItem.Text & ddlFromYear.SelectedItem.Text) _
          And IsDate(ddlToDay.SelectedItem.Text & ddlToMonth.SelectedItem.Text & ddlToYear.SelectedItem.Text)) Then

            Call GetDateRange()
            spnDateExample1.Visible = False
            spnDateExample2.Visible = False
            imgCalendarButton1.Visible = False
            imgCalendarButton2.Visible = False
            
            lblReportGeneratedDateTime.Visible = True
            
            Call GenerateReport()
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
    
    Protected Sub GenerateReport()
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataTable As New DataTable()
        Dim sbSQL As New StringBuilder
        'sbSQL.Append("SELECT c.CreatedOn 'Order Rcvd', lp.ProductCode 'Product Code', lp.ProductDescription 'Description' , BookNumber 'Book Number', FirstPageNumber 'First Page Number', c.AWB 'Consignment' ")
        'sbSQL.Append("c.CneeName + ', ' + CneeAddr1 + ' ' + CneeTown + ' ' + CneePostcode 'Consignee', up.UserId 'Terminal ID', CONVERT(VARCHAR(8), WarehouseCutOffTime, 3) 'Despatch Date'  ")
        'sbSQL.Append("FROM LogisticMovement lm ")
        'sbSQL.Append("LEFT OUTER JOIN ClientData_WUIRE_SerialNumbers fsn ")
        'sbSQL.Append("ON lm.ConsignmentKey = fsn.ConsignmentKey ")
        'sbSQL.Append("INNER JOIN LogisticProduct lp ")
        'sbSQL.Append("ON lm.LogisticProductKey = lp.LogisticProductKey ")
        'sbSQL.Append("INNER JOIN Consignment c ")
        'sbSQL.Append("ON lm.ConsignmentKey = c.[key] ")
        'sbSQL.Append("INNER JOIN UserProfile up ")
        'sbSQL.Append("ON c.CreatedBy = up.[key] ")
        'sbSQL.Append("WHERE lm.ConsignmentKey IN ")
        'sbSQL.Append("( ")
        'sbSQL.Append("SELECT c.[key] FROM Consignment c INNER JOIN ClientData_WUIRE_SerialNumbers fsn ON c.[key] = fsn.ConsignmentKey ")
        'sbSQL.Append(") ")
        'sbSQL.Append("AND lm.LogisticProductKey IN ")
        'sbSQL.Append("( ")
        'sbSQL.Append("SELECT LogisticProductKey FROM LogisticProduct WHERE SerialNumbersFlag = 'Y' ")
        'sbSQL.Append(") ")
        'sbSQL.Append("AND c.CreatedOn BETWEEN '")
        'sbSQL.Append(psFromDate)
        'sbSQL.Append("' AND '")
        'sbSQL.Append(psToDate)
        'sbSQL.Append("'")

        sbSQL.Append("SELECT c.CreatedOn 'Order Rcvd', lp.ProductCode 'Product Code', lp.ProductDescription 'Description' , BookNumber 'Book Number', FirstPageNumber 'First Page Number', c.AWB 'Consignment', ")
        sbSQL.Append("ISNULL(c.CneeName,'') + ', ' + ISNULL(CneeAddr1,'') + ' ' + ISNULL(CneeTown,'') + ' ' + ISNULL(CneePostcode,'') 'Consignee', up.UserId 'Terminal ID', CONVERT(VARCHAR(8), WarehouseCutOffTime, 3) 'Despatch Date'  ")
        sbSQL.Append("FROM LogisticMovement lm ")
        sbSQL.Append("LEFT OUTER JOIN ClientData_WUIRE_SerialNumbers fsn ")
        sbSQL.Append("ON lm.ConsignmentKey = fsn.ConsignmentKey ")
        sbSQL.Append("LEFT OUTER JOIN LogisticProduct lp ")
        sbSQL.Append("ON fsn.LogisticProductKey = lp.LogisticProductKey ")
        sbSQL.Append("INNER JOIN Consignment c ")
        sbSQL.Append("ON lm.ConsignmentKey = c.[key] ")
        sbSQL.Append("INNER JOIN UserProfile up ")
        sbSQL.Append("ON c.CreatedBy = up.[key] ")
        sbSQL.Append("WHERE lm.ConsignmentKey IN ")
        sbSQL.Append("( ")
        sbSQL.Append("SELECT c.[key] FROM Consignment c INNER JOIN ClientData_WUIRE_SerialNumbers fsn ON c.[key] = fsn.ConsignmentKey ")
        sbSQL.Append(") ")
        sbSQL.Append("AND lm.LogisticProductKey IN ")
        sbSQL.Append("( ")
        sbSQL.Append("SELECT LogisticProductKey FROM LogisticProduct WHERE SerialNumbersFlag = 'Y' ")
        sbSQL.Append(") ")
        sbSQL.Append("AND c.CreatedOn >= CAST('")
        sbSQL.Append(psFromDate)
        sbSQL.Append("' AS smalldatetime) AND c.CreatedOn < DATEADD(DAY, 1, '")
        sbSQL.Append(psToDate)
        sbSQL.Append("') ")
        sbSQL.Append("ORDER BY c.CreatedOn ")
        Dim oAdapter As New SqlDataAdapter(sbSQL.ToString, oConn)
        Try
            oAdapter.Fill(oDataTable)
            If oDataTable.Rows.Count > 0 Then
                Dim sCSVString As String = ConvertDataTableToCSVString(oDataTable)
                Call ExportCSVData(sCSVString)
            Else
                WebMsgBox.Show("No data found for this date range")
                lblReportGeneratedDateTime.Visible = False
            End If
        Catch ex As SqlException
            WebMsgBox.Show("Error in GenerateReport: " & ex.Message)
        Finally
            oConn.Close()
        End Try
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
        Response.AddHeader("Content-Disposition", "attachment;filename=" & "WesternUnionLodgements.csv")
        'Response.ContentType = "application/vnd.ms-excel"
        Response.ContentType = "text/csv"
    
        Dim eEncoding As Encoding = Encoding.GetEncoding("Windows-1252")
        Dim eUnicode As Encoding = Encoding.Unicode
        Dim byUnicodeBytes As Byte() = eUnicode.GetBytes(sCSVString)
        Dim byEncodedBytes As Byte() = Encoding.Convert(eUnicode, eEncoding, byUnicodeBytes)
        Response.BinaryWrite(byEncodedBytes)
        Response.End()
    End Sub

    Property psToDate() As String
        Get
            Dim o As Object = ViewState("BHR_ToDate")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("BHR_ToDate") = Value
        End Set
    End Property
    
    Property psFromDate() As String
        Get
            Dim o As Object = ViewState("BHR_FromDate")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("BHR_FromDate") = Value
        End Set
    End Property
    
</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>Western Union IRELAND Lodgement Report</title>
    <link rel="stylesheet" type="text/css" href="../css/sprint.css" />
</head>
<body>
    <form id="frmBookingHistoryReport" runat="server">
        <table id="tblDateRangeSelector" runat="server" visible="true">
            <tr id="Tr1" runat="server">
                <td colspan="4" style="white-space:nowrap; height: 21px;">
                    &nbsp;
                    <asp:Label ID="Label1" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="X-Small"
                        Text="Western Union IRELAND Lodgement Report"></asp:Label></td>
            </tr>
            <tr runat="server" visible="true">
                <td runat="server" colspan="2" style="white-space: nowrap">
                </td>
                <td style="width: 253px; height: 14px">
                </td>
                <td align="right" style="width: 169px; height: 14px">
                </td>
            </tr>
            <tr runat="server" visible="true">
                <td style="width: 265px; white-space: nowrap; height: 14px">
                </td>
                <td style="white-space: nowrap; height: 14px">
                </td>
                <td style="width: 253px; height: 14px">
                </td>
                <td style="width: 169px; height: 14px">
                </td>
            </tr>
            <tr runat="server" visible="true" id="CalendarInterface">
                <td style="width: 265px; white-space:nowrap">
                    <span class="informational dark">&nbsp;From:</span>
                        <asp:TextBox ID="tbFromDate" font-names="Verdana" font-size="XX-Small" Width="90" runat="server"/>
                         <a id="imgCalendarButton1" runat="server" visible="true" href="javascript:;"
                            onclick="window.open('../PopupCalendar4.aspx?textbox=tbFromDate','cal','width=300,height=305,left=270,top=180')">
                            <img id="Img1" src="../images/SmallCalendar.gif" runat="server" border="0" alt="" IE:visible="true" visible="false"
                               /></a>
                    <span id="spnDateExample1" runat="server" visible="true" class="informational light">(eg&nbsp;12-Jan-2009)</span>
                </td>
                <td style="white-space:nowrap" >
                    <span class="informational dark">To:</span>
                        <asp:TextBox ID="tbToDate" font-names="Verdana" font-size="XX-Small" Width="90" runat="server"/>
                           <a id="imgCalendarButton2" runat="server" visible="true" href="javascript:;"
                            onclick="window.open('../PopupCalendar4.aspx?textbox=tbToDate','cal','width=300,height=305,left=270,top=180')"><img id="Img2"
                                 src="../images/SmallCalendar.gif" runat="server" border="0" alt="" IE:visible="true" visible="false"
                               /></a>
                    <span id="spnDateExample2" runat="server" visible="true" class="informational light">(eg&nbsp;12-Jan-2009)</span>
                </td>
                <td style="width: 253px">
                    <asp:Button ID="btnRunReport1" runat="server" Text="run" Visible="true" OnClick="btnRunReport_Click" />
                    <asp:Button ID="btnReselectDateRange1" runat="server" Text="reselect report parameters" Visible="false" OnClick="btnReselectDateRange_Click" />
                      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                </td>
                <td style="width: 169px">
                    <asp:LinkButton ID="lnkbtnToggleSelectionStyle1" runat="server" OnClick="lnkbtnToggleSelectionStyle_Click" ToolTip="toggles between calendar interface and dropdown interface">change&nbsp;selection&nbsp;style</asp:LinkButton>
                </td>
            </tr>
            <tr runat="server" visible="false" id="DropdownInterface">
                <td style="width: 265px">
                    <span class="informational dark">&nbsp;From:</span>
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
                <td style="width: 265px">
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
                <td style="width: 253px">
                <asp:Button ID="btnRunReport2"
                     runat="server"
                     Text="run"
                      OnClick="btnRunReport_Click" />
                    <asp:Button ID="btnReselectDateRange2"
                     runat="server"
                     Text="reselect report parameters"
                     Visible="false"
                      OnClick="btnReselectDateRange_Click" />
                      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                      </td>
                <td style="width: 169px">
                    <asp:LinkButton ID="lnkbtnToggleSelectionStyle2"
                                      runat="server"
                                      OnClick="lnkbtnToggleSelectionStyle_Click"
                                      ToolTip="toggles between easy-to-use calendar interface and clunky dropdown interface">change&nbsp;selection&nbsp;style</asp:LinkButton></td>
            </tr>
            <tr runat="server" visible="true" id="DateValidationMessages">
                <td style="width: 265px">
                    <asp:RegularExpressionValidator ID="revFromDate" runat="server" ControlToValidate="tbFromDate" ErrorMessage=" - invalid format for expiry date - use dd-mmm-yy"
                        Font-Names="Verdana" Font-Size="XX-Small" ValidationExpression="^\d\d-(jan|Jan|JAN|feb|Feb|FEB|mar|Mar|MAR|apr|Apr|APR|may|May|MAY|jun|Jun|JUN|jul|Jul|JUL|aug|Aug|AUG|sep|Sep|SEP|oct|Oct|OCT|nov|Nov|NOV|dec|Dec|DEC)-\d(\d+)" SetFocusOnError="True" ValidationGroup="CalendarInterface"></asp:RegularExpressionValidator>
                    <asp:RangeValidator ID="rvFromDate" runat="server" ControlToValidate="tbFromDate"
                        CultureInvariantValues="True" ErrorMessage=" - expiry year before 2000, after 2020, or not a valid date!"
                        Font-Names="Verdana" Font-Size="XX-Small" MaximumValue="2019/1/1" MinimumValue="2000/1/1" ValidationGroup="CalendarInterface" EnableClientScript="False" Type="Date"></asp:RangeValidator>
                    <asp:Label ID="lblFromErrorMessage" runat="server" Font-Names="Verdana,Sans-Serif"
                        Font-Size="XX-Small" ForeColor="Red"></asp:Label></td>
                <td style="width: 265px">
                    <asp:RegularExpressionValidator ID="RegularevToDate" runat="server" ControlToValidate="tbToDate" ErrorMessage=" - invalid format for expiry date - use dd-mmm-yy"
                        Font-Names="Verdana" Font-Size="XX-Small" ValidationExpression="^\d\d-(jan|Jan|JAN|feb|Feb|FEB|mar|Mar|MAR|apr|Apr|APR|may|May|MAY|jun|Jun|JUN|jul|Jul|JUL|aug|Aug|AUG|sep|Sep|SEP|oct|Oct|OCT|nov|Nov|NOV|dec|Dec|DEC)-\d(\d+)" ValidationGroup="CalendarInterface"></asp:RegularExpressionValidator><asp:RangeValidator
                            ID="rvToDate" runat="server" ControlToValidate="tbToDate" CultureInvariantValues="True" ErrorMessage=" - expiry year before 2000, after 2020, or not a valid date!"
                            Font-Names="Verdana" Font-Size="XX-Small" MaximumValue="2019/1/1" MinimumValue="2000/1/1" ValidationGroup="CalendarInterface" EnableClientScript="False" Type="Date"></asp:RangeValidator>
                    <asp:Label ID="lblToErrorMessage" runat="server" Font-Names="Verdana,Sans-Serif"
                        Font-Size="XX-Small" ForeColor="Red"></asp:Label></td>
                <td style="width: 253px">
                </td>
                <td style="width: 169px">
                </td>
            </tr>
        </table>
        <asp:Label ID="lblReportGeneratedDateTime" runat="server" Text="" font-size="XX-Small" font-names="Verdana, Sans-Serif" forecolor="Green" Visible="false"></asp:Label>
    </form>
</body>
</html>
