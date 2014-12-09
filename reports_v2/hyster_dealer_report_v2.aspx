<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ import Namespace="System" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>
<script runat="server">

    ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '   Copyright Jonathan Hare May 2004
    '   Part of the web interface to Stock Manager
    ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    
    Property FromDate() As String
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
    
    Property ToDate() As String
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
    
    Property DealershipCode() As String
        Get
            Dim o As Object = ViewState("FilterName")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("FilterName") = Value
        End Set
    End Property
    
    Public Class ConsignmentList
        Inherits CollectionBase
        Default Public Property Item(ByVal Index As Long) As Consignment
            Get
                Return CType(List(Index), Consignment)
            End Get
            Set(ByVal Value As Consignment)
                List(Index) = Value
            End Set
        End Property
    
        Public Function Add(ByVal Value As Consignment) As Long
            Return List.Add(Value)
        End Function
    
        Public Shared Function GetConsignmentList(ByVal sFromDate As String, ByVal sToDate As String, iUserKey As Integer) As ConsignmentList
    
            Dim obj As ConsignmentList = New ConsignmentList
            Dim dr as DataRow
            Dim sConn As String = ConfigurationSettings.AppSettings("AIMSRootConnectionString")
            Dim oConn As New SqlConnection(sConn)
            Dim oDataSet As New DataSet()
            Dim oAdapter As New SqlDataAdapter("spASPNET_Hyster_DealersStockBookings_Rpt",oConn)
            Try
                oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
                oAdapter.SelectCommand.Parameters("@CustomerKey").Value = 77
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@FromDate", SqlDbType.Smalldatetime))
                oAdapter.SelectCommand.Parameters("@FromDate").Value = sFromDate
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ToDate", SqlDbType.Smalldatetime))
                oAdapter.SelectCommand.Parameters("@ToDate").Value = sToDate
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@UserKey", SqlDbType.Int))
                oAdapter.SelectCommand.Parameters("@UserKey").Value = iUserKey
    
                oAdapter.Fill(oDataSet, "Consignments")
    
                For Each dr in oDataSet.Tables("Consignments").Rows
                    obj.Add(New Consignment(CLng(dr("ConsignmentKey")), _
                                            CLng(dr("LogisticBookingKey")), _
                                            dr("AWB").ToString(), _
                                            CDate(dr("ShipDate")), _
                                            dr("CneeName").ToString(), _
                                            dr("CneeAddr1").ToString(), _
                                            dr("CneeAddr2").ToString(), _
                                            dr("CneeAddr3").ToString(), _
                                            dr("CneeTown").ToString(), _
                                            dr("CneeState").ToString(), _
                                            dr("CneePostCode").ToString(), _
                                            dr("CneeCountry").ToString(), _
                                            dr("CneeCtcName").ToString(), _
                                            dr("CneeTel").ToString(), _
                                            dr("CustomerRef1").ToString(), _
                                            dr("CustomerRef2").ToString(), _
                                            CInt(dr("NOP")), _
                                            CDbl(dr("Weight")), _
                                            "", _
                                            dr("Misc1").ToString(), _
                                            dr("Misc2").ToString(), _
                                            CDbl(dr("ShippingCost")), _
                                            dr("PODName").ToString(), _
                                            dr("PODDate").ToString(), _
                                            dr("PODTime").ToString(), _
                                            dr("BookedBy").ToString(), _
                                            dr("Title").ToString(), _
                                            dr("Department").ToString()))
    
                Next
    
                Return obj
    
            Catch ex As SqlException
            Finally
                oConn.Close()
            End Try
        End Function
    End Class
    
    Public Class Consignment
        Private _ConsignmentKey As Long
        Private _LogisticBookingKey As Long
        Private _AWB As String
        Private _ShipDate As Date
        Private _CneeName As String
        Private _CneeAddr1 As String
        Private _CneeAddr2 As String
        Private _CneeAddr3 As String
        Private _CneeTown As String
        Private _CneeState As String
        Private _CneePostCode As String
        Private _CneeCountry As String
        Private _CneeCtcName As String
        Private _CneeTel As String
        Private _CustomerRef1 As String
        Private _CustomerRef2 As String
        Private _NOP As Integer
        Private _Weight As Double
        Private _UnitOfMeasurementId As String
        Private _Misc1 As String
        Private _Misc2 As String
        Private _ShippingCost As Double
        Private _PODName As String
        Private _PODDate As String
        Private _PODTime As String
        Private _BookedBy As String
        Private _Title As String
        Private _Department As String
    
        Private _StockItemList As StockItemList
        Private _StockItemListValue As Double
    
        Public Property ConsignmentKey() As Long
            Get
                Return _ConsignmentKey
            End Get
            Set(ByVal Value As Long)
                _ConsignmentKey = Value
            End Set
        End Property
    
        Public Property LogisticBookingKey() As Long
            Get
                Return _LogisticBookingKey
            End Get
            Set(ByVal Value As Long)
                _LogisticBookingKey = Value
            End Set
        End Property
    
        Public Property AWB() As String
            Get
                Return _AWB
            End Get
            Set(ByVal Value As String)
                _AWB = Value
            End Set
        End Property
    
        Public Property ShipDate() As Date
            Get
                Return _ShipDate
            End Get
            Set(ByVal Value As Date)
                _ShipDate = Value
            End Set
        End Property
    
        Public Property CneeName() As String
            Get
                Return _CneeName
            End Get
            Set(ByVal Value As String)
                _CneeName = Value
            End Set
        End Property
    
        Public Property CneeAddr1() As String
            Get
                Return _CneeAddr1
            End Get
            Set(ByVal Value As String)
                _CneeAddr1 = Value
            End Set
        End Property
    
        Public Property CneeAddr2() As String
            Get
                Return _CneeAddr2
            End Get
            Set(ByVal Value As String)
                _CneeAddr2 = Value
            End Set
        End Property
    
        Public Property CneeAddr3() As String
            Get
                Return _CneeAddr3
            End Get
            Set(ByVal Value As String)
                _CneeAddr3 = Value
            End Set
        End Property
    
        Public Property CneeTown() As String
            Get
                Return _CneeTown
            End Get
            Set(ByVal Value As String)
                _CneeTown = Value
            End Set
        End Property
    
        Public Property CneeState() As String
            Get
                Return _CneeState
            End Get
            Set(ByVal Value As String)
                _CneeState = Value
            End Set
        End Property
    
        Public Property CneePostCode() As String
            Get
                Return _CneePostCode
            End Get
            Set(ByVal Value As String)
                _CneePostCode = Value
            End Set
        End Property
    
        Public Property CneeCountry() As String
            Get
                Return _CneeCountry
            End Get
            Set(ByVal Value As String)
                _CneeCountry = Value
            End Set
        End Property
    
        Public Property CneeCtcName() As String
            Get
                Return _CneeCtcName
            End Get
            Set(ByVal Value As String)
                _CneeCtcName = Value
            End Set
        End Property
    
        Public Property CneeTel() As String
            Get
                Return _CneeTel
            End Get
            Set(ByVal Value As String)
                _CneeTel = Value
            End Set
        End Property
    
        Public Property CustomerRef1() As String
            Get
                Return _CustomerRef1
            End Get
            Set(ByVal Value As String)
                _CustomerRef1 = Value
            End Set
        End Property
    
        Public Property CustomerRef2() As String
            Get
                Return _CustomerRef2
            End Get
            Set(ByVal Value As String)
                _CustomerRef2 = Value
            End Set
        End Property
    
        Public Property NOP() As Integer
            Get
                Return _NOP
            End Get
            Set(ByVal Value As Integer)
                _NOP = Value
            End Set
        End Property
    
        Public Property Weight() As Double
            Get
                Return _Weight
            End Get
            Set(ByVal Value As Double)
                _Weight = Value
            End Set
        End Property
    
        Public Property UnitOfMeasurementId() As String
            Get
                Return _UnitOfMeasurementId
            End Get
            Set(ByVal Value As String)
                _UnitOfMeasurementId = Value
            End Set
        End Property
    
        Public Property Misc1() As String
            Get
                Return _Misc1
            End Get
            Set(ByVal Value As String)
                _Misc1 = Value
            End Set
        End Property
    
        Public Property Misc2() As String
            Get
                Return _Misc2
            End Get
            Set(ByVal Value As String)
                _Misc2 = Value
            End Set
        End Property
    
        Public Property ShippingCost() As Double
            Get
                Return _ShippingCost
            End Get
            Set(ByVal Value As Double)
                _ShippingCost = Value
            End Set
        End Property
    
        Public Property PODName() As String
            Get
                Return _PODName
            End Get
            Set(ByVal Value As String)
                _PODName = Value
            End Set
        End Property
    
        Public Property PODDate() As String
            Get
                Return _PODDate
            End Get
            Set(ByVal Value As String)
                _PODDate = Value
            End Set
        End Property
    
        Public Property PODTime() As String
            Get
                Return _PODTime
            End Get
            Set(ByVal Value As String)
                _PODTime = Value
            End Set
        End Property
    
        Public Property BookedBy() As String
            Get
                Return _BookedBy
            End Get
            Set(ByVal Value As String)
                _BookedBy = Value
            End Set
        End Property
    
        Public Property Title() As String
            Get
                Return _Title
            End Get
            Set(ByVal Value As String)
                _Title = Value
            End Set
        End Property
    
        Public Property Department() As String
            Get
                Return _Department
            End Get
            Set(ByVal Value As String)
                _Department = Value
            End Set
        End Property
    
        Public ReadOnly Property StockItemList() As StockItemList
            Get
                Return _StockItemList
            End Get
        End Property
    
        Public ReadOnly Property StockItemListValue() As Double
            Get
                Return _StockItemListValue
            End Get
        End Property
    
        Public Sub New()
        End Sub
    
        Public Sub New(ByVal ConsignmentKey As Long, _
                        ByVal LogisticBookingKey As Long, _
                        ByVal AWB As String, _
                        ByVal ShipDate As Date, _
                        ByVal CneeName As String, _
                        ByVal CneeAddr1 As String, _
                        ByVal CneeAddr2 As String, _
                        ByVal CneeAddr3 As String, _
                        ByVal CneeTown As String, _
                        ByVal CneeState As String, _
                        ByVal CneePostCode As String, _
                        ByVal CneeCountry As String, _
                        ByVal CneeCtcName As String, _
                        ByVal CneeTel As String, _
                        ByVal CustomerRef1 As String, _
                        ByVal CustomerRef2 As String, _
                        ByVal NOP As Integer, _
                        ByVal Weight As Double, _
                        ByVal UnitOfMeasurementId As String, _
                        ByVal Misc1 As String, _
                        ByVal Misc2 As String, _
                        ByVal ShippingCost As Double, _
                        ByVal PODName As String, _
                        ByVal PODDate As String, _
                        ByVal PODTime As String, _
                        ByVal BookedBy As String, _
                        ByVal Title As String, _
                        ByVal Department As String)
    
            _ConsignmentKey = ConsignmentKey
            _LogisticBookingKey = LogisticBookingKey
            _AWB = AWB
            _ShipDate = ShipDate
            _CneeName = CneeName
            _CneeAddr1 = CneeAddr1
            _CneeAddr2 = CneeAddr2
            _CneeAddr3 = CneeAddr3
            _CneeTown = CneeTown
            _CneeState = CneeState
            _CneePostCode = CneePostCode
            _CneeCountry = CneeCountry
            _CneeCtcName = CneeCtcName
            _CneeTel = CneeTel
            _CustomerRef1 = CustomerRef1
            _CustomerRef2 = CustomerRef2
            _NOP = NOP
            _Weight = Weight
            _UnitOfMeasurementId = UnitOfMeasurementId
            _Misc1 = Misc1
            _Misc2 = Misc2
            _ShippingCost = ShippingCost
            _BookedBy = BookedBy
            _Title = Title
            _Department = Department
            _PODName = PODName
            _PODDate = PODDate
            _PODTime = PODTime
    
            _StockItemList = StockItemList.GetStockItemList(LogisticBookingKey)
            _StockItemListValue = StockItemList.GetStockItemListValue(LogisticBookingKey)
    
        End Sub
    
    End Class
    
    Public Class StockItemList
        Inherits CollectionBase
    
        Public Shared Function GetStockItemList(ByVal LogisticBookingKey As Long) As StockItemList
            Dim obj As StockItemList = New StockItemList
            Dim dr as DataRow
            Dim sConn As String = ConfigurationSettings.AppSettings("AIMSRootConnectionString")
            Dim oConn As New SqlConnection(sConn)
            Dim oDataSet As New DataSet()
            Dim oAdapter As New SqlDataAdapter("spASPNET_StockBooking_GetProducts",oConn)
            Try
                oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@StockBookingKey", SqlDbType.Int))
                oAdapter.SelectCommand.Parameters("@StockBookingKey").Value = LogisticBookingKey
    
                oAdapter.Fill(oDataSet, "StockItems")
    
                For Each dr in oDataSet.Tables("StockItems").Rows
                    obj.Add(New StockItem(  dr("ProductCode"), _
                                            dr("ProductDate"), _
                                            dr("ProductDescription"), _
                                            dr("ItemsOut"), _
                                            CDbl(dr("UnitValue"))))
    
                Next
    
                Return obj
    
            Catch ex As SqlException
            Finally
                oConn.Close()
            End Try
    
        End Function
    
        Public Shared Function GetStockItemListValue(ByVal LogisticBookingKey As Long) As Double
            Dim dr as DataRow
            Dim sConn As String = ConfigurationSettings.AppSettings("AIMSRootConnectionString")
            Dim oConn As New SqlConnection(sConn)
            Dim oDataSet As New DataSet()
            Dim oAdapter As New SqlDataAdapter("spASPNET_StockBooking_GetProducts",oConn)
            Dim Value As Double
            Try
                oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@StockBookingKey", SqlDbType.Int))
                oAdapter.SelectCommand.Parameters("@StockBookingKey").Value = LogisticBookingKey
    
                oAdapter.Fill(oDataSet, "StockItems")
    
                For Each dr in oDataSet.Tables("StockItems").Rows
                    If dr("UnitValue") > 0 Then
                        Value = Value + (CLng(dr("ItemsOut")) * CDbl(dr("UnitValue")))
                    End If
                Next
    
                Return Value
    
            Catch ex As SqlException
            Finally
                oConn.Close()
            End Try
    
        End Function
    
        Default Public Property Item(ByVal Index As Long) As StockItem
            Get
                Return CType(List(Index), StockItem)
            End Get
            Set(ByVal Value As StockItem)
                List(Index) = Value
            End Set
        End Property
    
        Public Function Add(ByVal Value As StockItem) As Long
            Return List.Add(Value)
        End Function
    
    End Class
    
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '~~~~~~~~~~~~~ StockItem Class ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    
    Public Class StockItem
        Private _ProdCode As String
        Private _ProdDate As String
        Private _ProdDescription As String
        Private _Quantity As Integer
        Private _UnitValue As Double
    
        Public Sub New( ByVal ProdCode As String, _
                        ByVal ProdDate As String, _
                        ByVal ProdDescription As String, _
                        ByVal Quantity As Integer, _
                        ByVal UnitValue As Double)
    
            _ProdCode = ProdCode
            _ProdDate = ProdDate
            _ProdDescription = ProdDescription
            _Quantity = Quantity
            _UnitValue = UnitValue
    
        End Sub
    
        Public Property ProdCode() As String
            Get
                Return _ProdCode
            End Get
            Set(ByVal Value As String)
                _ProdCode = Value
            End Set
        End Property
    
        Public Property ProdDate() As String
            Get
                Return _ProdDate
            End Get
            Set(ByVal Value As String)
                _ProdDate = Value
            End Set
        End Property
    
        Public Property ProdDescription() As String
            Get
                Return _ProdDescription
            End Get
            Set(ByVal Value As String)
                _ProdDescription = Value
            End Set
        End Property
    
        Public Property Quantity() As Integer
            Get
                Return _Quantity
            End Get
            Set(ByVal Value As Integer)
                _Quantity = Value
            End Set
        End Property
    
        Public Property UnitValue() As Double
            Get
                Return _UnitValue
            End Get
            Set(ByVal Value As Double)
                _UnitValue = Value
            End Set
        End Property
    End Class
    
    '~~~~~~~~~~~~~ Page Load ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    
    Sub Page_Load(Source As Object, E As EventArgs)
        If Not IsPostBack Then
    
            Dim dteFromDate As Date = Date.Today.AddMonths(-1)
    
            Dim iFromDay As integer = Day(Now)
            Dim iFromMonth As Integer = Datepart(DateInterval.Month,dteFromDate)
            Dim iFromYear As Integer = Year(dteFromDate)
    
            Dim iToDay As Integer = Day(Now)
            Dim iToMonth As Integer = Datepart(DateInterval.Month,Now)
            Dim iToYear As Integer = Year(Now)
    
            drop_FromDay.SelectedIndex = iFromDay
            drop_FromMonth.SelectedIndex = iFromMonth
            drop_FromYear.SelectedValue = iFromYear
            drop_ToDay.SelectedIndex = iToDay
            drop_ToMonth.SelectedIndex = iToMonth
            drop_ToYear.SelectedValue = iToYear
    
            ShowReportCriteria()
    
        End IF
    
    End Sub
    
    Sub btn_RunReport_Click(ByVal s As Object, ByVal e As ImageClickEventArgs)
        If ValidDate() Then
            DataList1.DataSource = ConsignmentList.GetConsignmentList(FromDate, ToDate, Session("UserKey"))
            DataList1.DataBind()
            ShowReportData()
        End If
    End Sub
    
    Sub ShowReportCriteria()
        pnlReportCriteria.Visible = True
        pnlReportData.Visible = False
    End Sub
    
    Sub ShowReportData()
        pnlReportCriteria.Visible = False
        pnlReportData.Visible = True
    End Sub
    
    Function ValidDate()As Boolean
        Dim sFromDayPart As String
        Dim sFromMonthPart As String
        Dim sFromYearPart As String
        Dim sToDayPart As String
        Dim sToMonthPart As String
        Dim sToYearPart As String
        Dim sMessage As String
        If drop_FromDay.SelectedItem.Text = "DAY" Then
            ValidDate = False
            sMessage = "[FROM DAY]"
        End If
        If drop_FromMonth.SelectedItem.Text = "MONTH" Then
            ValidDate = False
            sMessage &= "[FROM MONTH]"
        End If
        If drop_FromYear.SelectedItem.Text = "YEAR" Then
            ValidDate = False
            sMessage &= "[FROM YEAR]"
        End If
        If drop_ToDay.SelectedItem.Text = "DAY" Then
            ValidDate = False
            sMessage &= "[TO DAY]"
        End If
        If drop_ToMonth.SelectedItem.Text = "MONTH" Then
            ValidDate = False
            sMessage &= "[TO MONTH]"
        End If
        If drop_ToYear.SelectedItem.Text = "YEAR" Then
            ValidDate = False
            sMessage &= "[TO YEAR]"
        End If
    
        If sMessage <> "" Then
            lblDateError.Text = "Invalid date: " & sMessage
        Else
            ValidDate = True
            lblDateError.Text = ""
            sFromDayPart = drop_FromDay.SelectedItem.Text
            sFromMonthPart = drop_FromMonth.SelectedItem.Text
            sFromYearPart = drop_FromYear.SelectedItem.Text
            sToDayPart = drop_ToDay.SelectedItem.Text
            sToMonthPart = drop_ToMonth.SelectedItem.Text
            sToYearPart = drop_ToYear.SelectedItem.Text
            Try
                FromDate = DateTime.Parse(sFromDayPart & " " & sFromMonthPart & " " & sFromYearPart)
                'FromDate = sFromDayPart & " " & sFromMonthPart & " " & sFromYearPart
            Catch ex As Exception
                ValidDate = False
                sMessage &= "Incorrect 'From' date"
                lblDateError.Text = "Invalid date: " & sMessage & " "
            End Try
            Try
                ToDate = DateTime.Parse(sToDayPart & " " & sToMonthPart & " " & sToYearPart)
                'ToDate = sToDayPart & " " & sToMonthPart & " " & sToYearPart
            Catch ex As Exception
                ValidDate = False
                sMessage &= "Incorrect 'To' date"
                lblDateError.Text = "Invalid date: " & sMessage
            End Try
        End If
    End Function

</script>
<html>
<head>
    <title>Hyster Dealer Report</title>
    <LINK rel="stylesheet" type="text/css" href="../Reports.css" />
</head>
<body>
    <form runat="server">
          <asp:Table id="TableHeader" runat="server" width="100%">
              <asp:TableRow>
                  <asp:TableCell VerticalAlign="Bottom" width="0%"></asp:TableCell>
                  <asp:TableCell Wrap="False" width="50%">
                      <asp:Label ID="Label1" runat="server" forecolor="silver" font-size="Small" font-bold="True" font-names="Arial">Hyster
                      Dealer Report</asp:Label><br /><br />
                  </asp:TableCell>
                  <asp:TableCell Wrap="False" HorizontalAlign="Right" width="50%"></asp:TableCell>
              </asp:TableRow>
          </asp:Table>
        <asp:Panel id="pnlReportCriteria" runat="server" visible="False">
            <asp:Table id="tabDates" runat="server" Width="650px" font-names="Verdana">
                <asp:TableRow>
                    <asp:TableCell VerticalAlign="Top" Width="20%" Wrap="False"><asp:Label ID="Label2" runat="server" font-size="X-Small">Select period:</asp:Label></asp:TableCell>
                    <asp:TableCell VerticalAlign="Top" Width="20%" Wrap="False"></asp:TableCell>
                    <asp:TableCell VerticalAlign="Top" Width="30%" HorizontalAlign="Right" Wrap="False">
                        <asp:Label runat="server" font-size="X-Small">From:</asp:Label> &nbsp;
                        <asp:DropDownList runat="server" Font-Size="XX-Small" Font-Names="Verdana" ID="drop_FromDay">
                            <asp:ListItem Value="0">DAY</asp:ListItem>
                            <asp:ListItem Value="1">1</asp:ListItem>
                            <asp:ListItem Value="2">2</asp:ListItem>
                            <asp:ListItem Value="3">3</asp:ListItem>
                            <asp:ListItem Value="4">4</asp:ListItem>
                            <asp:ListItem Value="5">5</asp:ListItem>
                            <asp:ListItem Value="6">6</asp:ListItem>
                            <asp:ListItem Value="7">7</asp:ListItem>
                            <asp:ListItem Value="8">8</asp:ListItem>
                            <asp:ListItem Value="9">9</asp:ListItem>
                            <asp:ListItem Value="10">10</asp:ListItem>
                            <asp:ListItem Value="11">11</asp:ListItem>
                            <asp:ListItem Value="12">12</asp:ListItem>
                            <asp:ListItem Value="13">13</asp:ListItem>
                            <asp:ListItem Value="14">14</asp:ListItem>
                            <asp:ListItem Value="15">15</asp:ListItem>
                            <asp:ListItem Value="16">16</asp:ListItem>
                            <asp:ListItem Value="17">17</asp:ListItem>
                            <asp:ListItem Value="18">18</asp:ListItem>
                            <asp:ListItem Value="19">19</asp:ListItem>
                            <asp:ListItem Value="20">20</asp:ListItem>
                            <asp:ListItem Value="21">21</asp:ListItem>
                            <asp:ListItem Value="22">22</asp:ListItem>
                            <asp:ListItem Value="23">23</asp:ListItem>
                            <asp:ListItem Value="24">24</asp:ListItem>
                            <asp:ListItem Value="25">25</asp:ListItem>
                            <asp:ListItem Value="26">26</asp:ListItem>
                            <asp:ListItem Value="27">27</asp:ListItem>
                            <asp:ListItem Value="28">28</asp:ListItem>
                            <asp:ListItem Value="29">29</asp:ListItem>
                            <asp:ListItem Value="30">30</asp:ListItem>
                            <asp:ListItem Value="31">31</asp:ListItem>
                        </asp:DropDownList>
                        <asp:DropDownList runat="server" Font-Size="XX-Small" Font-Names="Verdana" ID="drop_FromMonth">
                            <asp:ListItem Value="0" Selected="True">MONTH</asp:ListItem>
                            <asp:ListItem Value="1">JAN</asp:ListItem>
                            <asp:ListItem Value="2">FEB</asp:ListItem>
                            <asp:ListItem Value="3">MAR</asp:ListItem>
                            <asp:ListItem Value="4">APR</asp:ListItem>
                            <asp:ListItem Value="5">MAY</asp:ListItem>
                            <asp:ListItem Value="6">JUN</asp:ListItem>
                            <asp:ListItem Value="7">JUL</asp:ListItem>
                            <asp:ListItem Value="8">AUG</asp:ListItem>
                            <asp:ListItem Value="9">SEP</asp:ListItem>
                            <asp:ListItem Value="10">OCT</asp:ListItem>
                            <asp:ListItem Value="11">NOV</asp:ListItem>
                            <asp:ListItem Value="12">DEC</asp:ListItem>
                        </asp:DropDownList>
                        <asp:DropDownList runat="server" Font-Size="XX-Small" Font-Names="Verdana" ID="drop_FromYear">
                            <asp:ListItem Value="0">YEAR</asp:ListItem>
                            <asp:ListItem Value="1998">1998</asp:ListItem>
                            <asp:ListItem Value="1999">1999</asp:ListItem>
                            <asp:ListItem Value="2000">2000</asp:ListItem>
                            <asp:ListItem Value="2001">2001</asp:ListItem>
                            <asp:ListItem Value="2002">2002</asp:ListItem>
                            <asp:ListItem Value="2003">2003</asp:ListItem>
                            <asp:ListItem Value="2004">2004</asp:ListItem>
                            <asp:ListItem Value="2005">2005</asp:ListItem>
                            <asp:ListItem Value="2006">2006</asp:ListItem>
                            <asp:ListItem Value="2007">2007</asp:ListItem>
                            <asp:ListItem Value="2008">2008</asp:ListItem>
                            <asp:ListItem Value="2009">2009</asp:ListItem>
                            <asp:ListItem Value="2010">2010</asp:ListItem>
                            <asp:ListItem Value="2011">2011</asp:ListItem>
                            <asp:ListItem Value="2012">2012</asp:ListItem>
                            <asp:ListItem Value="2013">2013</asp:ListItem>
                            <asp:ListItem Value="2014">2014</asp:ListItem>
                        </asp:DropDownList>
                    </asp:TableCell>
                    <asp:TableCell VerticalAlign="Top" Width="30%" HorizontalAlign="Right" Wrap="False">
                        <asp:Label runat="server" font-size="X-Small">To:</asp:Label> &nbsp;
                        <asp:DropDownList runat="server" Font-Size="XX-Small" Font-Names="Verdana" ID="drop_ToDay">
                            <asp:ListItem Value="0">DAY</asp:ListItem>
                            <asp:ListItem Value="1">1</asp:ListItem>
                            <asp:ListItem Value="2">2</asp:ListItem>
                            <asp:ListItem Value="3">3</asp:ListItem>
                            <asp:ListItem Value="4">4</asp:ListItem>
                            <asp:ListItem Value="5">5</asp:ListItem>
                            <asp:ListItem Value="6">6</asp:ListItem>
                            <asp:ListItem Value="7">7</asp:ListItem>
                            <asp:ListItem Value="8">8</asp:ListItem>
                            <asp:ListItem Value="9">9</asp:ListItem>
                            <asp:ListItem Value="10">10</asp:ListItem>
                            <asp:ListItem Value="11">11</asp:ListItem>
                            <asp:ListItem Value="12">12</asp:ListItem>
                            <asp:ListItem Value="13">13</asp:ListItem>
                            <asp:ListItem Value="14">14</asp:ListItem>
                            <asp:ListItem Value="15">15</asp:ListItem>
                            <asp:ListItem Value="16">16</asp:ListItem>
                            <asp:ListItem Value="17">17</asp:ListItem>
                            <asp:ListItem Value="18">18</asp:ListItem>
                            <asp:ListItem Value="19">19</asp:ListItem>
                            <asp:ListItem Value="20">20</asp:ListItem>
                            <asp:ListItem Value="21">21</asp:ListItem>
                            <asp:ListItem Value="22">22</asp:ListItem>
                            <asp:ListItem Value="23">23</asp:ListItem>
                            <asp:ListItem Value="24">24</asp:ListItem>
                            <asp:ListItem Value="25">25</asp:ListItem>
                            <asp:ListItem Value="26">26</asp:ListItem>
                            <asp:ListItem Value="27">27</asp:ListItem>
                            <asp:ListItem Value="28">28</asp:ListItem>
                            <asp:ListItem Value="29">29</asp:ListItem>
                            <asp:ListItem Value="30">30</asp:ListItem>
                            <asp:ListItem Value="31">31</asp:ListItem>
                        </asp:DropDownList>
                        <asp:DropDownList runat="server" Font-Size="XX-Small" Font-Names="Verdana" ID="drop_ToMonth">
                            <asp:ListItem Value="0" Selected="True">MONTH</asp:ListItem>
                            <asp:ListItem Value="1">JAN</asp:ListItem>
                            <asp:ListItem Value="2">FEB</asp:ListItem>
                            <asp:ListItem Value="3">MAR</asp:ListItem>
                            <asp:ListItem Value="4">APR</asp:ListItem>
                            <asp:ListItem Value="5">MAY</asp:ListItem>
                            <asp:ListItem Value="6">JUN</asp:ListItem>
                            <asp:ListItem Value="7">JUL</asp:ListItem>
                            <asp:ListItem Value="8">AUG</asp:ListItem>
                            <asp:ListItem Value="9">SEP</asp:ListItem>
                            <asp:ListItem Value="10">OCT</asp:ListItem>
                            <asp:ListItem Value="11">NOV</asp:ListItem>
                            <asp:ListItem Value="12">DEC</asp:ListItem>
                        </asp:DropDownList>
                        <asp:DropDownList runat="server" Font-Size="XX-Small" Font-Names="Verdana" ID="drop_ToYear">
                            <asp:ListItem Value="0">YEAR</asp:ListItem>
                            <asp:ListItem Value="1998">1998</asp:ListItem>
                            <asp:ListItem Value="1999">1999</asp:ListItem>
                            <asp:ListItem Value="2000">2000</asp:ListItem>
                            <asp:ListItem Value="2001">2001</asp:ListItem>
                            <asp:ListItem Value="2002">2002</asp:ListItem>
                            <asp:ListItem Value="2003">2003</asp:ListItem>
                            <asp:ListItem Value="2004">2004</asp:ListItem>
                            <asp:ListItem Value="2005">2005</asp:ListItem>
                            <asp:ListItem Value="2006">2006</asp:ListItem>
                            <asp:ListItem Value="2007">2007</asp:ListItem>
                            <asp:ListItem Value="2008">2008</asp:ListItem>
                            <asp:ListItem Value="2009">2009</asp:ListItem>
                            <asp:ListItem Value="2010">2010</asp:ListItem>
                            <asp:ListItem Value="2011">2011</asp:ListItem>
                            <asp:ListItem Value="2012">2012</asp:ListItem>
                            <asp:ListItem Value="2013">2013</asp:ListItem>
                            <asp:ListItem Value="2014">2014</asp:ListItem>
                        </asp:DropDownList>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow>
                    <asp:TableCell ColumnSpan="4" HorizontalAlign="Right">
                        <asp:Label id="lblDateError" runat="server" forecolor="Red" font-size="XX-Small"></asp:Label>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow>
                    <asp:TableCell ColumnSpan="4" HorizontalAlign="Right">
                        <br/>
                        <asp:Button ID="btnRunReport" runat="server" Text="run report" onclick="btn_RunReport_Click" />
                    </asp:TableCell>
                </asp:TableRow>
            </asp:Table>
        </asp:Panel>
        <asp:Panel id="pnlReportData" runat="server" visible="False">
            <asp:DataList id="DataList1" runat="server" EnableViewState="False">
                <ItemTemplate>
                    <asp:Table id="Table1" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="650px">
                        <asp:TableRow>
                            <asp:TableCell ColumnSpan="8">
                                <hr />
                            </asp:TableCell>
                        </asp:TableRow>
                        <asp:TableRow>
                            <asp:TableCell width="50px"></asp:TableCell>
                            <asp:TableCell width="100px"></asp:TableCell>
                            <asp:TableCell width="100px"></asp:TableCell>
                            <asp:TableCell width="75px"></asp:TableCell>
                            <asp:TableCell width="50px"></asp:TableCell>
                            <asp:TableCell width="100px"></asp:TableCell>
                            <asp:TableCell width="100px"></asp:TableCell>
                            <asp:TableCell width="75px"></asp:TableCell>
                        </asp:TableRow>
                        <asp:TableRow>
                            <asp:TableCell ColumnSpan="4">
                                <asp:Label runat="server" font-size="X-Small">Stock Booking No</asp:Label> &nbsp;<asp:Label runat="server" font-size="X-Small" font-bold="True"><%# Format(DataBinder.Eval(Container.DataItem,"LogisticBookingKey"),"0000000") %></asp:Label>
                            </asp:TableCell>
                            <asp:TableCell ColumnSpan="4" HorizontalAlign="Right">
                                <asp:Label runat="server" font-size="X-Small">Booked On</asp:Label> &nbsp;<asp:Label runat="server" font-size="X-Small" font-bold="True"><%# Format(DataBinder.Eval(Container.DataItem,"ShipDate"),"dd MMM yyyy HH:mm") %></asp:Label>
                            </asp:TableCell>
                        </asp:TableRow>
                        <asp:TableRow>
                            <asp:TableCell ColumnSpan="4">
                                <asp:Label runat="server" font-size="X-Small">Dealership</asp:Label> &nbsp;<asp:Label runat="server" font-size="X-Small" font-bold="True"><%# DataBinder.Eval(Container.DataItem,"Department") %></asp:Label> &nbsp;<asp:Label runat="server" font-size="X-Small">[</asp:Label> <asp:Label runat="server" font-size="X-Small" font-bold="True"><%# DataBinder.Eval(Container.DataItem,"Title") %></asp:Label> <asp:Label runat="server" font-size="X-Small">]</asp:Label>
                            </asp:TableCell>
                            <asp:TableCell ColumnSpan="4" HorizontalAlign="Right">
                                <asp:Label runat="server" font-size="X-Small">Booked By</asp:Label> &nbsp;<asp:Label runat="server" font-size="X-Small" font-bold="True"><%# DataBinder.Eval(Container.DataItem,"BookedBy") %></asp:Label>
                            </asp:TableCell>
                        </asp:TableRow>
                        <asp:TableRow>
                            <asp:TableCell ColumnSpan="4">
                                <asp:Label runat="server" font-size="X-Small">Air Waybill</asp:Label> &nbsp;<asp:Label runat="server" font-size="X-Small" font-bold="True"><%# (DataBinder.Eval(Container.DataItem, "AWB"))%></asp:Label> &nbsp;<asp:Label runat="server" font-size="X-Small">/</asp:Label> &nbsp;<asp:Label runat="server" font-size="X-Small"><%# (DataBinder.Eval(Container.DataItem, "NOP")) & " @ " & Format((DataBinder.Eval(Container.DataItem, "Weight")),"#,##0.0") & " " %></asp:Label>
                            </asp:TableCell>
                            <asp:TableCell ColumnSpan="4" HorizontalAlign="Right">
                                <asp:Label runat="server" font-size="X-Small">Delivered To</asp:Label> &nbsp;<asp:Label runat="server" font-size="X-Small" font-bold="True"><%# DataBinder.Eval(Container.DataItem,"PODName")%></asp:Label> &nbsp;<asp:Label runat="server" font-size="X-Small"><%# DataBinder.Eval(Container.DataItem,"PODDate") & " " & DataBinder.Eval(Container.DataItem,"PODTime") %></asp:Label>
                            </asp:TableCell>
                        </asp:TableRow>
                        <asp:TableRow>
                            <asp:TableCell ColumnSpan="4">
                                <br />
                                <asp:Label runat="server" font-size="X-Small"><%# DataBinder.Eval(Container.DataItem,"CneeName") %></asp:Label>
                            </asp:TableCell>
                            <asp:TableCell></asp:TableCell>
                            <asp:TableCell></asp:TableCell>
                            <asp:TableCell></asp:TableCell>
                            <asp:TableCell></asp:TableCell>
                        </asp:TableRow>
                        <asp:TableRow>
                            <asp:TableCell ColumnSpan="4">
                                <asp:Label runat="server" font-size="XX-Small"><%# DataBinder.Eval(Container.DataItem,"CneeAddr1") %></asp:Label>
                            </asp:TableCell>
                            <asp:TableCell></asp:TableCell>
                            <asp:TableCell></asp:TableCell>
                            <asp:TableCell></asp:TableCell>
                            <asp:TableCell></asp:TableCell>
                        </asp:TableRow>
                        <asp:TableRow>
                            <asp:TableCell ColumnSpan="4">
                                <asp:Label runat="server" font-size="XX-Small"><%# DataBinder.Eval(Container.DataItem,"CneeTown") %></asp:Label>
                            </asp:TableCell>
                            <asp:TableCell></asp:TableCell>
                            <asp:TableCell></asp:TableCell>
                            <asp:TableCell></asp:TableCell>
                            <asp:TableCell></asp:TableCell>
                        </asp:TableRow>
                        <asp:TableRow>
                            <asp:TableCell ColumnSpan="4">
                                <asp:Label runat="server" font-size="XX-Small"><%# DataBinder.Eval(Container.DataItem,"CneeCountry") %></asp:Label>
                            </asp:TableCell>
                            <asp:TableCell></asp:TableCell>
                            <asp:TableCell></asp:TableCell>
                            <asp:TableCell></asp:TableCell>
                            <asp:TableCell></asp:TableCell>
                        </asp:TableRow>
                        <asp:TableRow>
                            <asp:TableCell ColumnSpan="4">
                                <asp:Label runat="server" font-size="XX-Small"><%# DataBinder.Eval(Container.DataItem,"CneeCtcName") %></asp:Label>
                                <br />
                                <br />
                            </asp:TableCell>
                            <asp:TableCell></asp:TableCell>
                            <asp:TableCell></asp:TableCell>
                            <asp:TableCell></asp:TableCell>
                            <asp:TableCell></asp:TableCell>
                        </asp:TableRow>
                        <asp:TableRow>
                            <asp:TableCell ColumnSpan="8">
                                <hr />
                            </asp:TableCell>
                        </asp:TableRow>
                        <asp:TableRow>
                            <asp:TableCell></asp:TableCell>
                            <asp:TableCell></asp:TableCell>
                            <asp:TableCell></asp:TableCell>
                            <asp:TableCell></asp:TableCell>
                            <asp:TableCell></asp:TableCell>
                            <asp:TableCell></asp:TableCell>
                            <asp:TableCell></asp:TableCell>
                            <asp:TableCell></asp:TableCell>
                        </asp:TableRow>
                    </asp:Table>
                    <asp:DataGrid id="DataGrid1" runat="server" DataSource='<%# DataBinder.Eval(Container.DataItem,"StockItemList") %>' AutoGenerateColumns="False" Font-Names="Verdana" Font-Size="XX-Small" Width="650px" GridLines="None">
                        <Columns>
                            <asp:BoundColumn DataField="ProdCode" HeaderText="Product Code">
                                <HeaderStyle font-bold="True" width="90px"></HeaderStyle>
                                <ItemStyle verticalalign="Top"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="ProdDate" HeaderText="Product Date">
                                <HeaderStyle font-bold="True" width="100px"></HeaderStyle>
                                <ItemStyle verticalalign="Top"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="ProdDescription" HeaderText="Product Description">
                                <HeaderStyle font-bold="True" width="250px"></HeaderStyle>
                                <ItemStyle verticalalign="Top"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="Quantity" HeaderText="Quantity">
                                <HeaderStyle font-bold="True" horizontalalign="Right" width="70px"></HeaderStyle>
                                <ItemStyle horizontalalign="Right" verticalalign="Top"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="UnitValue" HeaderText="Unit Cost" DataFormatString="{0:#,##0.00}">
                                <HeaderStyle font-bold="True" horizontalalign="Right" width="70px"></HeaderStyle>
                                <ItemStyle horizontalalign="Right" verticalalign="Top"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:TemplateColumn HeaderText="Value (€)">
                                <HeaderStyle font-bold="True" horizontalalign="Right" width="70px"></HeaderStyle>
                                <ItemStyle horizontalalign="Right" verticalalign="Top"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Label runat="server"><%# Format((DataBinder.Eval(Container.DataItem, "Quantity")) * (DataBinder.Eval(Container.DataItem, "UnitValue")),"#,##0.00") %></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                        </Columns>
                    </asp:DataGrid>
                    <asp:Table runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="650px">
                        <asp:TableRow>
                            <asp:TableCell width="50px"></asp:TableCell>
                            <asp:TableCell width="100px"></asp:TableCell>
                            <asp:TableCell width="100px"></asp:TableCell>
                            <asp:TableCell width="100px"></asp:TableCell>
                            <asp:TableCell width="100px"></asp:TableCell>
                            <asp:TableCell width="100px"></asp:TableCell>
                            <asp:TableCell width="100px"></asp:TableCell>
                        </asp:TableRow>
                        <asp:TableRow>
                            <asp:TableCell ColumnSpan="6"></asp:TableCell>
                            <asp:TableCell ColumnSpan="1">
                                <hr />
                            </asp:TableCell>
                        </asp:TableRow>
                        <asp:TableRow>
                            <asp:TableCell horizontalalign="Right" ColumnSpan="6">
                                <asp:Label runat="server" font-size="X-Small">Total value this order (€)</asp:Label>
                            </asp:TableCell>
                            <asp:TableCell horizontalalign="Right" ColumnSpan="1">
                                <asp:Label runat="server" font-size="X-Small" font-bold="True"><%# Format(DataBinder.Eval(Container.DataItem,"StockItemListValue"),"#,##0.00") %></asp:Label>
                            </asp:TableCell>
                        </asp:TableRow>
                        <asp:TableRow>
                            <asp:TableCell ColumnSpan="3"></asp:TableCell>
                            <asp:TableCell horizontalalign="Right" ColumnSpan="3">
                                <asp:Label runat="server" font-size="X-Small">Shipping costs this order (£)</asp:Label>
                            </asp:TableCell>
                            <asp:TableCell horizontalalign="Right" ColumnSpan="1">
                                <asp:Label runat="server" font-size="X-Small" font-bold="True"><%# Format(DataBinder.Eval(Container.DataItem,"ShippingCost"),"#,##0.00") %></asp:Label>
                            </asp:TableCell>
                        </asp:TableRow>
                        <asp:TableRow>
                            <asp:TableCell ColumnSpan="4"></asp:TableCell>
                            <asp:TableCell ColumnSpan="3">
                                <hr />
                            </asp:TableCell>
                        </asp:TableRow>
                        <asp:TableRow>
                            <asp:TableCell ColumnSpan="7">
                                <br />
                                <br />
                                <br />
                                <br />
                            </asp:TableCell>
                        </asp:TableRow>
                    </asp:Table>
                </ItemTemplate>
            </asp:DataList>
        </asp:Panel>
    </form>
</body>
</html>
