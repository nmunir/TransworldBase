<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ import Namespace="System" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>
<script runat="server">

    '   Consignments By Consignee Report

    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString()
    Private gsSiteType As String = ConfigLib.GetConfigItem_SiteType
    Private gbSiteTypeDefined = gsSiteType.Length > 0
    
    Sub Page_Load(Source As Object, E As EventArgs)
        If Not IsNumeric(Session("CustomerKey")) Then
            Server.Transfer("../session_expired.aspx")
        End If

        lblReportGeneratedDateTime.Text = "Report generated: " & Now().ToString("dd-MMM-yy HH:mm")
        If Not IsPostBack Then
            pbIsProductOwner = CBool(Session("UserType").ToString.ToLower.Contains("owner"))
            Call GetSiteFeatures()
            trProductGroups.Visible = pbProductOwners
            ' pbProductOwners = site-wide Product Owners attribute; pbIsProductOwner = this user
            If pbIsProductOwner Then
                If pbProductOwners Then
                    ddlProductGroup.Visible = True
                    PopulateProductGroups(Session("UserKey"))
                    btnShowProductGroups.Visible = False
                Else
                    WebMsgBox.Show("Cannot show report as Product Owners attribute is not enabled for this web site")
                    Exit Sub
                End If
            Else
                If pbProductOwners Then
                    btnShowProductGroups.Visible = True
                Else
                    btnShowProductGroups.Visible = False
                End If
                pnSelectedProductGroup = 0
            End If
            Call SetDateDropdowns()
            Call ShowReportCriteria()
        End If
    End Sub
    
    Protected Sub SetDateDropdowns()
        Dim dteFromDate As Date = Date.Today.AddMonths(-1)
    
        Dim iFromDay As Integer = Day(Now)
        Dim iFromMonth As Integer = DatePart(DateInterval.Month, dteFromDate)
        Dim iFromYear As Integer = Year(dteFromDate)
    
        Dim iToDay As Integer = Day(Now)
        Dim iToMonth As Integer = DatePart(DateInterval.Month, Now)
        Dim iToYear As Integer = Year(Now)
    
        Call SetCalendarYears()
        ddlFromDay.SelectedIndex = iFromDay
        ddlFromMonth.SelectedIndex = iFromMonth
        If iFromYear <> iToYear Then
            ddlFromYear.SelectedIndex = 2
        Else
            ddlFromYear.SelectedIndex = 3
        End If
        ddlToDay.SelectedIndex = iToDay
        ddlToMonth.SelectedIndex = iToMonth
        ddlToYear.SelectedIndex = 3
    End Sub
    
    Protected Sub SetCalendarYears()
        Dim iThisYear As Integer = Year(Now)
        ddlFromYear.Items.Add(New ListItem(iThisYear - 3, iThisYear - 3))
        ddlFromYear.Items.Add(New ListItem(iThisYear - 2, iThisYear - 2))
        ddlFromYear.Items.Add(New ListItem(iThisYear - 1, iThisYear - 1))
        ddlFromYear.Items.Add(New ListItem(iThisYear, iThisYear))
        ddlFromYear.Items.Add(New ListItem(iThisYear + 1, iThisYear + 1))
    
        ddlToYear.Items.Add(New ListItem(iThisYear - 3, iThisYear - 3))
        ddlToYear.Items.Add(New ListItem(iThisYear - 2, iThisYear - 2))
        ddlToYear.Items.Add(New ListItem(iThisYear - 1, iThisYear - 1))
        ddlToYear.Items.Add(New ListItem(iThisYear, iThisYear))
        ddlToYear.Items.Add(New ListItem(iThisYear + 1, iThisYear + 1))
    End Sub
    
    Protected Sub GetSiteFeatures()
        Dim oDataTable As New DataTable
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter("spASPNET_SiteContent", oConn)
        
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@Action", SqlDbType.NVarChar, 50))
        oAdapter.SelectCommand.Parameters("@Action").Value = "GET"
        
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SiteKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@SiteKey").Value = Session("SiteKey")
        
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ContentType", SqlDbType.NVarChar, 50))
        oAdapter.SelectCommand.Parameters("@ContentType").Value = "SiteSettings"

        Try
            oAdapter.Fill(oDataTable)
        Catch ex As Exception
            WebMsgBox.Show("GetSiteFeatures: " & ex.Message)
        Finally
            oConn.Close()
        End Try

        Dim dr As DataRow = oDataTable.Rows(0)
        pbProductOwners = dr("ProductOwners")
    End Sub
    
    Protected Sub PopulateProductGroups(ByVal nProductOwner As Integer)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Product_GetGroupsForOwner", oConn)
        oCmd.CommandType = CommandType.StoredProcedure

        Dim paramProductOwner As SqlParameter = New SqlParameter("@ProductOwner", SqlDbType.Int)
        paramProductOwner.Value = nProductOwner
        oCmd.Parameters.Add(paramProductOwner)
       
        Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int)
        paramCustomerKey.Value = Session("CustomerKey")
        oCmd.Parameters.Add(paramCustomerKey)
       
        Try
            oConn.Open()
            Dim oSqlDataReader As SqlDataReader = oCmd.ExecuteReader
            If oSqlDataReader.HasRows Then
                ddlProductGroup.Items.Add(New ListItem("- select product group -", -1))
                If Not pbIsProductOwner Then
                    ddlProductGroup.Items.Add(New ListItem("- all products -", 0))
                End If
                While oSqlDataReader.Read()
                    ddlProductGroup.Items.Add(New ListItem(oSqlDataReader("ProductGroupName"), oSqlDataReader("ProductGroupKey")))
                End While
            End If
        Catch ex As Exception
            WebMsgBox.Show("PopulateProductgGroupsDropdown: " & ex.Message)
        Finally
            oConn.Close()
        End Try
        
        If ddlProductGroup.Items.Count <= 2 Then
            lblProductGroup.Text = "Product group: " & ddlProductGroup.Items(1).Text
            pnSelectedProductGroup = ddlProductGroup.Items(1).Value
            ddlProductGroup.Visible = False
        Else
            btnRunReport.Enabled = False
            pnSelectedProductGroup = -1
        End If
    End Sub
    
    Protected Sub dgProducts_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs)
        If e.Item.ItemType = ListItemType.Header Then
            If IsHyster() Then
                e.Item.Cells(5).Text = "Value (€)"
            Else
                e.Item.Cells(5).Text = "Value (£)"
            End If
        End If
    End Sub

    Sub btnRunReport_Click(ByVal s As Object, ByVal e As EventArgs)
        If bValidDate() Then
            dlConsignments.DataSource = ConsignmentList.GetConsignmentList(sFromDate, sToDate, Session("UserKey"), Session("CustomerKey"), pnSelectedProductGroup)
            dlConsignments.DataBind()
            If dlConsignments.Items.Count = 0 Then
                lblResult.Visible = True
                lblReportGeneratedDateTime.Visible = False
            Else
                lblResult.Visible = False
                lblReportGeneratedDateTime.Visible = True
            End If
            ShowReportData()
        End If
    End Sub
    
    Protected Function IsHyster() As Boolean
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsHyster = IIf(gbSiteTypeDefined, gsSiteType = "hyster", nCustomerKey = 77)
    End Function
   
    Protected Function IsNotHyster() As Boolean
        IsNotHyster = Not IsHyster()
    End Function

    Protected Function sCurrency(ByVal sLegend As String) As String
        If IsHyster() Then
            sCurrency = sLegend & " (€)"
        Else
            sCurrency = sLegend & " (£)"
        End If
    End Function

    Sub ShowReportCriteria()
        pnlReportCriteria.Visible = True
        pnlReportData.Visible = False
    End Sub
    
    Sub ShowReportData()
        'pnlReportCriteria.Visible = False
        pnlReportData.Visible = True
    End Sub
    
    Function bValidDate() As Boolean
        Dim bIsValid As Boolean = True
        Dim sFromDayPart As String
        Dim sFromMonthPart As String
        Dim sFromYearPart As String
        Dim sToDayPart As String
        Dim sToMonthPart As String
        Dim sToYearPart As String
        Dim sTestDate As String
        Dim sMessage As String = String.Empty
        If ddlFromDay.SelectedItem.Text = "DAY" Then
            bIsValid = False
            sMessage = "[FROM DAY]"
        End If
        If ddlFromMonth.SelectedItem.Text = "MONTH" Then
            bIsValid = False
            sMessage &= "[FROM MONTH]"
        End If
        If ddlFromYear.SelectedItem.Text = "YEAR" Then
            bIsValid = False
            sMessage &= "[FROM YEAR]"
        End If
        If ddlToDay.SelectedItem.Text = "DAY" Then
            bIsValid = False
            sMessage &= "[TO DAY]"
        End If
        If ddlToMonth.SelectedItem.Text = "MONTH" Then
            bIsValid = False
            sMessage &= "[TO MONTH]"
        End If
        If ddlToYear.SelectedItem.Text = "YEAR" Then
            bIsValid = False
            sMessage &= "[TO YEAR]"
        End If
    
        If sMessage <> "" Then
            lblDateError.Text = "Invalid date: " & sMessage
        Else
            bIsValid = True
            lblDateError.Text = ""
            sFromDayPart = ddlFromDay.SelectedItem.Text
            sFromMonthPart = ddlFromMonth.SelectedItem.Text
            sFromYearPart = ddlFromYear.SelectedItem.Text
            sToDayPart = ddlToDay.SelectedItem.Text
            sToMonthPart = ddlToMonth.SelectedItem.Text
            sToYearPart = ddlToYear.SelectedItem.Text
            sFromDate = sFromDayPart & " " & sFromMonthPart & " " & sFromYearPart
            Try
                sTestDate = DateTime.Parse(sFromDate)
            Catch ex As Exception
                bIsValid = False
                sMessage &= "Incorrect 'From' date"
                lblDateError.Text = "Invalid date: " & sMessage & " "
            End Try

            sToDate = sToDayPart & " " & sToMonthPart & " " & sToYearPart
            Try
                sTestDate = DateTime.Parse(sToDate)
            Catch ex As Exception
                bIsValid = False
                sMessage &= "Incorrect 'To' date"
                lblDateError.Text = "Invalid date: " & sMessage
            End Try
        End If
        If bIsValid Then
            If DateTime.Parse(sToDate) < DateTime.Parse(sFromDate) Then
                lblDateError.Text = "From date is more recent than To date"
                bIsValid = False
            End If
        End If
        bValidDate = bIsValid
    End Function

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
    
        Public Shared Function GetConsignmentList(ByVal sFromDate As String, ByVal sToDate As String, ByVal iUserKey As Integer, ByVal lCustomerKey As Long, ByVal nSelectedProductGroup As Integer) As ConsignmentList
            Dim obj As ConsignmentList = New ConsignmentList
            Dim sConn As String = ConfigLib.GetConfigItem_ConnectionString
            Dim oConn As New SqlConnection(sConn)
            Dim oDataTable As New DataTable
            Try
                Dim sbSQL As New StringBuilder
                sbSQL.Append("SELECT awb.[Key] ConsignmentKey, lb.LogisticBookingKey, awb.AWB, awb.CreatedOn ShipDate, ")
                sbSQL.Append("awb.CneeName, awb.CneeAddr1, awb.CneeAddr2, awb.CneeAddr3, awb.CneeTown, awb.CneeState, ")
                sbSQL.Append("awb.CneePostCode, c1.CountryName CneeCountry, awb.CneeCtcName, awb.CneeTel, awb.CustomerRef1, ")
                sbSQL.Append("awb.CustomerRef2, NOP = ISNULL(awb.NOP,0), Weight = ISNULL(awb.Weight,0.0), ")
                sbSQL.Append("awb.Misc1, awb.Misc2, ShippingCost = ISNULL(awb.CashOnDelAmount,0.0), awb.PODName, awb.PODDate, ")
                sbSQL.Append("awb.PODTime, UserProfile.FirstName + ' ' + UserProfile.LastName BookedBy, UserProfile.Title, UserProfile.Department ")
                sbSQL.Append("FROM Consignment As awb ")
                sbSQL.Append("LEFT OUTER JOIN Country As c1 ")
                sbSQL.Append("ON awb.CneeCountryKey = c1.CountryKey ")
                sbSQL.Append("INNER JOIN LogisticBooking As lb ")
                sbSQL.Append("ON awb.StockBookingKey = lb.LogisticBookingKey ")
                sbSQL.Append("INNER JOIN UserProfile ")
                sbSQL.Append("ON lb.BookedByKey = UserProfile.[Key] ")
                sbSQL.Append("INNER JOIN LogisticMovement lm ")
                sbSQL.Append("ON awb.[key] = lm.ConsignmentKey ")
                sbSQL.Append("INNER JOIN LogisticProduct lp ")
                sbSQL.Append("ON lm.[LogisticProductkey] = lp.LogisticProductkey ")
                sbSQL.Append("WHERE awb.CustomerKey = " & lCustomerKey)
                sbSQL.Append(" AND awb.CreatedOn BETWEEN '" & sFromDate & "' AND '" & sToDate)
                sbSQL.Append("' AND NOT awb.StateId = 'CANCELLED' ")
                If nSelectedProductGroup > 0 Then
                    sbSQL.Append("AND lp.StockOwnedByKey = " & nSelectedProductGroup)
                End If
                sbSQL.Append(" ORDER BY UserProfile.Department, awb.CreatedOn")
                Dim oAdapter As New SqlDataAdapter(sbSQL.ToString, oConn)

                oAdapter.Fill(oDataTable)
                For Each dr As DataRow In oDataTable.Rows
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
    
        Public Property lConsignmentKey() As Long
            Get
                Return _ConsignmentKey
            End Get
            Set(ByVal Value As Long)
                _ConsignmentKey = Value
            End Set
        End Property
    
        Public Property lLogisticBookingKey() As Long
            Get
                Return _LogisticBookingKey
            End Get
            Set(ByVal Value As Long)
                _LogisticBookingKey = Value
            End Set
        End Property
    
        Public Property sAWB() As String
            Get
                Return _AWB
            End Get
            Set(ByVal Value As String)
                _AWB = Value
            End Set
        End Property
    
        Public Property dtShipDate() As Date
            Get
                Return _ShipDate
            End Get
            Set(ByVal Value As Date)
                _ShipDate = Value
            End Set
        End Property
    
        Public Property sCneeName() As String
            Get
                Return _CneeName
            End Get
            Set(ByVal Value As String)
                _CneeName = Value
            End Set
        End Property
    
        Public Property sCneeAddr1() As String
            Get
                Return _CneeAddr1
            End Get
            Set(ByVal Value As String)
                _CneeAddr1 = Value
            End Set
        End Property
    
        Public Property sCneeAddr2() As String
            Get
                Return _CneeAddr2
            End Get
            Set(ByVal Value As String)
                _CneeAddr2 = Value
            End Set
        End Property
    
        Public Property sCneeAddr3() As String
            Get
                Return _CneeAddr3
            End Get
            Set(ByVal Value As String)
                _CneeAddr3 = Value
            End Set
        End Property
    
        Public Property sCneeTown() As String
            Get
                Return _CneeTown
            End Get
            Set(ByVal Value As String)
                _CneeTown = Value
            End Set
        End Property
    
        Public Property sCneeState() As String
            Get
                Return _CneeState
            End Get
            Set(ByVal Value As String)
                _CneeState = Value
            End Set
        End Property
    
        Public Property sCneePostCode() As String
            Get
                Return _CneePostCode
            End Get
            Set(ByVal Value As String)
                _CneePostCode = Value
            End Set
        End Property
    
        Public Property sCneeCountry() As String
            Get
                Return _CneeCountry
            End Get
            Set(ByVal Value As String)
                _CneeCountry = Value
            End Set
        End Property
    
        Public Property sCneeCtcName() As String
            Get
                Return _CneeCtcName
            End Get
            Set(ByVal Value As String)
                _CneeCtcName = Value
            End Set
        End Property
    
        Public Property sCneeTel() As String
            Get
                Return _CneeTel
            End Get
            Set(ByVal Value As String)
                _CneeTel = Value
            End Set
        End Property
    
        Public Property sCustomerRef1() As String
            Get
                Return _CustomerRef1
            End Get
            Set(ByVal Value As String)
                _CustomerRef1 = Value
            End Set
        End Property
    
        Public Property sCustomerRef2() As String
            Get
                Return _CustomerRef2
            End Get
            Set(ByVal Value As String)
                _CustomerRef2 = Value
            End Set
        End Property
    
        Public Property nNOP() As Integer
            Get
                Return _NOP
            End Get
            Set(ByVal Value As Integer)
                _NOP = Value
            End Set
        End Property
    
        Public Property dblWeight() As Double
            Get
                Return _Weight
            End Get
            Set(ByVal Value As Double)
                _Weight = Value
            End Set
        End Property
    
        Public Property sMisc1() As String
            Get
                Return _Misc1
            End Get
            Set(ByVal Value As String)
                _Misc1 = Value
            End Set
        End Property
    
        Public Property sMisc2() As String
            Get
                Return _Misc2
            End Get
            Set(ByVal Value As String)
                _Misc2 = Value
            End Set
        End Property
    
        Public Property dblShippingCost() As Double
            Get
                Return _ShippingCost
            End Get
            Set(ByVal Value As Double)
                _ShippingCost = Value
            End Set
        End Property
    
        Public Property sPODName() As String
            Get
                Return _PODName
            End Get
            Set(ByVal Value As String)
                _PODName = Value
            End Set
        End Property
    
        Public Property sPODDate() As String
            Get
                Return _PODDate
            End Get
            Set(ByVal Value As String)
                _PODDate = Value
            End Set
        End Property
    
        Public Property sPODTime() As String
            Get
                Return _PODTime
            End Get
            Set(ByVal Value As String)
                _PODTime = Value
            End Set
        End Property
    
        Public Property sBookedBy() As String
            Get
                Return _BookedBy
            End Get
            Set(ByVal Value As String)
                _BookedBy = Value
            End Set
        End Property
    
        Public Property sTitle() As String
            Get
                Return _Title
            End Get
            Set(ByVal Value As String)
                _Title = Value
            End Set
        End Property
    
        Public Property sDepartment() As String
            Get
                Return _Department
            End Get
            Set(ByVal Value As String)
                _Department = Value
            End Set
        End Property
    
        Public ReadOnly Property dblStockItemList() As StockItemList
            Get
                Return _StockItemList
            End Get
        End Property
    
        Public ReadOnly Property dblStockItemListValue() As Double
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
            _Misc1 = Misc1
            _Misc2 = Misc2
            _ShippingCost = ShippingCost
            _BookedBy = BookedBy
            _Title = Title
            _Department = Department
            _PODName = PODName
            _PODDate = PODDate
            _PODTime = PODTime
    
            _StockItemList = dblStockItemList.GetStockItemList(LogisticBookingKey)
            _StockItemListValue = dblStockItemList.GetStockItemListValue(LogisticBookingKey)
    
        End Sub
    
    End Class
    
    Public Class StockItemList
        Inherits CollectionBase
    
        Public Shared Function GetStockItemList(ByVal LogisticBookingKey As Long) As StockItemList
            Dim obj As StockItemList = New StockItemList
            Dim dr As DataRow
            'Dim sConn As String = ConfigurationManager.AppSettings("AIMSRootConnectionString")
            Dim sConn As String = ConfigLib.GetConfigItem_ConnectionString
            Dim oConn As New SqlConnection(sConn)
            Dim oDataSet As New DataSet()
            Dim oAdapter As New SqlDataAdapter("spASPNET_StockBooking_GetProducts", oConn)
            Try
                oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@StockBookingKey", SqlDbType.Int))
                oAdapter.SelectCommand.Parameters("@StockBookingKey").Value = LogisticBookingKey
    
                oAdapter.Fill(oDataSet, "StockItems")
    
                For Each dr In oDataSet.Tables("StockItems").Rows
                    obj.Add(New StockItem(dr("ProductCode"), _
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
            Dim dr As DataRow
            ' Dim sConn As String = ConfigurationManager.AppSettings("AIMSRootConnectionString")
            Dim sConn As String = ConfigLib.GetConfigItem_ConnectionString
            Dim oConn As New SqlConnection(sConn)
            Dim oDataSet As New DataSet()
            Dim oAdapter As New SqlDataAdapter("spASPNET_StockBooking_GetProducts", oConn)
            Dim Value As Double
            Try
                oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@StockBookingKey", SqlDbType.Int))
                oAdapter.SelectCommand.Parameters("@StockBookingKey").Value = LogisticBookingKey
    
                oAdapter.Fill(oDataSet, "StockItems")
    
                For Each dr In oDataSet.Tables("StockItems").Rows
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
    
        Default Public Property lItem(ByVal Index As Long) As StockItem
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
    
    Public Class StockItem
        Private _ProdCode As String
        Private _ProdDate As String
        Private _ProdDescription As String
        Private _Quantity As Integer
        Private _UnitValue As Double
    
        Public Sub New(ByVal ProdCode As String, _
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
    
        Public Property sProdCode() As String
            Get
                Return _ProdCode
            End Get
            Set(ByVal Value As String)
                _ProdCode = Value
            End Set
        End Property
    
        Public Property sProdDate() As String
            Get
                Return _ProdDate
            End Get
            Set(ByVal Value As String)
                _ProdDate = Value
            End Set
        End Property
    
        Public Property sProdDescription() As String
            Get
                Return _ProdDescription
            End Get
            Set(ByVal Value As String)
                _ProdDescription = Value
            End Set
        End Property
    
        Public Property nQuantity() As Integer
            Get
                Return _Quantity
            End Get
            Set(ByVal Value As Integer)
                _Quantity = Value
            End Set
        End Property
    
        Public Property dblUnitValue() As Double
            Get
                Return _UnitValue
            End Get
            Set(ByVal Value As Double)
                _UnitValue = Value
            End Set
        End Property
    End Class
    
    Property sFromDate() As String
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
    
    Property sToDate() As String
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
    
    Property sDealershipCode() As String
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
    
    Property pnSelectedProductGroup() As Integer
        Get
            Dim o As Object = ViewState("BHR_SelectedProductGroup")
            If o Is Nothing Then
                Return 2
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Integer)
            ViewState("BHR_SelectedProductGroup") = Value
        End Set
    End Property
   
    Property pbProductOwners() As Boolean
        Get
            Dim o As Object = ViewState("BHR_ProductOwners")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("BHR_ProductOwners") = Value
        End Set
    End Property
   
    Property pbIsProductOwner() As Boolean
        Get
            Dim o As Object = ViewState("BHR_IsProductOwner")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("BHR_IsProductOwner") = Value
        End Set
    End Property
   
    Protected Sub btnShowProductGroups_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowProductGroups()
    End Sub

    Protected Sub ShowProductGroups()
        ddlProductGroup.Visible = True
        Call PopulateProductGroups(0)
        btnShowProductGroups.Visible = False
    End Sub
    
    Protected Sub ddlProductGroup_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        If ddl.Items(0).Value = -1 Then
            ddlProductGroup.Items.RemoveAt(0)
        End If
        pnSelectedProductGroup = ddl.SelectedValue
        btnRunReport.Enabled = True
    End Sub
    
</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>Consignments By Consignee Report</title>
    <link rel="stylesheet" type="text/css" href="../css/sprint.css" />
</head>
<body>
    <form id="frmCbyC" runat="server">
          <table width="100%">
              <tr>
                  <td style="width:50%; white-space:nowrap">
                      <asp:Label ID="Label1" runat="server" forecolor="silver" font-size="Small" font-bold="True" font-names="Arial">Consignments
                      By Consignee Report</asp:Label><br /><br />
                  </td>
                  <td style="width:50%; white-space:nowrap" align="right"></td>
              </tr>
          </table>
        <asp:Panel id="pnlReportCriteria" runat="server" visible="False">
            <table style="width:650px; font-family:Verdana">
                <tr>
                    <td align="right" style="width: 40%; white-space: nowrap" valign="top">
                    </td>
                    <td align="right" style="width: 40%; white-space: nowrap" valign="top">
                    </td>
                    <td style="width: 10%; white-space: nowrap" valign="top">
                    </td>
                </tr>
                <tr runat="server" id="trProductGroups">
                    <td align="left" colspan="2" style="white-space: nowrap" valign="top">
                        &nbsp;<asp:DropDownList ID="ddlProductGroup" runat="server" AutoPostBack="True" Font-Names="Verdana"
                            Font-Size="XX-Small" OnSelectedIndexChanged="ddlProductGroup_SelectedIndexChanged"
                            Visible="False">
                        </asp:DropDownList><asp:Label ID="lblProductGroup" runat="server" Font-Names="Verdana"
                            Font-Size="X-Small" Font-Bold="True"></asp:Label></td>
                    <td align="right" style="width: 10%; white-space: nowrap" valign="top">
                        <asp:Button ID="btnShowProductGroups" runat="server" OnClick="btnShowProductGroups_Click"
                            Text="show product groups" Visible="False" /></td>
                </tr>
                <tr>
                    <td align="right" style="width: 40%; white-space: nowrap" valign="top">
                    </td>
                    <td align="right" style="width: 40%; white-space: nowrap" valign="top">
                    </td>
                    <td style="width: 10%; white-space: nowrap" valign="top">
                    </td>
                </tr>
                <tr>
                    <td valign="top" style="width:40%; white-space:nowrap"align="right">
                        &nbsp;<asp:Label ID="l002" runat="server" font-size="X-Small">From:</asp:Label> &nbsp;
                        <asp:DropDownList runat="server" Font-Size="XX-Small" Font-Names="Verdana" ID="ddlFromDay">
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
                        <asp:DropDownList runat="server" Font-Size="XX-Small" Font-Names="Verdana" ID="ddlFromMonth">
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
                        <asp:DropDownList runat="server" Font-Size="XX-Small" Font-Names="Verdana" ID="ddlFromYear"/>
                    </td>
                    <td  style="width:40%; white-space:nowrap" valign="Top" align="right">
                        <asp:Label ID="l001" runat="server" font-size="X-Small">To:</asp:Label> &nbsp;
                        <asp:DropDownList runat="server" Font-Size="XX-Small" Font-Names="Verdana" ID="ddlToDay">
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
                        <asp:DropDownList runat="server" Font-Size="XX-Small" Font-Names="Verdana" ID="ddlToMonth">
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
                        <asp:DropDownList runat="server" Font-Size="XX-Small" Font-Names="Verdana" ID="ddlToYear"/>
                    </td>
                    <td valign="top" style="width:10%; white-space:nowrap">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                <asp:Button ID="btnRunReport"
                     runat="server"
                     Text="generate report"
                     Visible="true"
                     OnClick="btnRunReport_Click" Width="160px" />
                    </td>
                </tr>
                <tr>
                    <td colspan="3" align="right">
                        <asp:Label id="lblDateError" runat="server" forecolor="Red" font-size="XX-Small"></asp:Label>
                    </td>
                </tr>
            </table>
        </asp:Panel>
        <asp:Panel id="pnlReportData" runat="server" visible="False" Width="100%">
            <asp:DataList id="dlConsignments" runat="server" EnableViewState="False">
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
                                <asp:Label runat="server" font-size="X-Small">Stock Booking No</asp:Label> &nbsp;<asp:Label runat="server" font-size="X-Small" font-bold="True"><%#Format(DataBinder.Eval(Container.DataItem, "lLogisticBookingKey"), "0000000")%></asp:Label>
                            </asp:TableCell>
                            <asp:TableCell ColumnSpan="4" HorizontalAlign="Right">
                                <asp:Label runat="server" font-size="X-Small">Booked On</asp:Label> &nbsp;<asp:Label runat="server" font-size="X-Small" font-bold="True"><%#Format(DataBinder.Eval(Container.DataItem, "dtShipDate"), "dd MMM yyyy HH:mm")%></asp:Label>
                            </asp:TableCell>
                        </asp:TableRow>
                        <asp:TableRow>
                            <asp:TableCell ColumnSpan="4">
                                <asp:Label runat="server" font-size="X-Small">Consignee</asp:Label> &nbsp;<asp:Label runat="server" font-size="X-Small" font-bold="True"><%# DataBinder.Eval(Container.DataItem,"sCneeName") %></asp:Label> &nbsp;
                            </asp:TableCell>
                            <asp:TableCell ColumnSpan="4" HorizontalAlign="Right">
                                <asp:Label runat="server" font-size="X-Small">Booked By</asp:Label> &nbsp;<asp:Label runat="server" font-size="X-Small" font-bold="True"><%# DataBinder.Eval(Container.DataItem,"sBookedBy") %></asp:Label>
                            </asp:TableCell>
                        </asp:TableRow>
                        <asp:TableRow>
                            <asp:TableCell ColumnSpan="4">
                                <asp:Label runat="server" font-size="X-Small">Air Waybill</asp:Label> &nbsp;<asp:Label runat="server" font-size="X-Small" font-bold="True"><%#(DataBinder.Eval(Container.DataItem, "sAWB"))%></asp:Label> &nbsp;<asp:Label runat="server" font-size="X-Small">/</asp:Label> &nbsp;<asp:Label runat="server" font-size="X-Small"><%#(DataBinder.Eval(Container.DataItem, "nNOP")) & " @ " & Format((DataBinder.Eval(Container.DataItem, "dblWeight")), "#,##0.0")%></asp:Label>
                            </asp:TableCell>
                            <asp:TableCell ColumnSpan="4" HorizontalAlign="Right">
                                <asp:Label runat="server" font-size="X-Small">Delivered To</asp:Label> &nbsp;<asp:Label runat="server" font-size="X-Small" font-bold="True"><%# DataBinder.Eval(Container.DataItem,"sPODName")%></asp:Label> &nbsp;<asp:Label runat="server" font-size="X-Small"><%#DataBinder.Eval(Container.DataItem, "sPODDate") & " " & DataBinder.Eval(Container.DataItem, "sPODTime")%></asp:Label>
                            </asp:TableCell>
                        </asp:TableRow>
                        <asp:TableRow>
                            <asp:TableCell ColumnSpan="4">
                                <br />
                                <asp:Label runat="server" font-size="X-Small"><%#DataBinder.Eval(Container.DataItem, "sCneeName")%></asp:Label>
                            </asp:TableCell>
                            <asp:TableCell></asp:TableCell>
                            <asp:TableCell></asp:TableCell>
                            <asp:TableCell></asp:TableCell>
                            <asp:TableCell></asp:TableCell>
                        </asp:TableRow>
                        <asp:TableRow>
                            <asp:TableCell ColumnSpan="4">
                                <asp:Label runat="server" font-size="XX-Small"><%#DataBinder.Eval(Container.DataItem, "sCneeAddr1")%></asp:Label>
                            </asp:TableCell>
                            <asp:TableCell></asp:TableCell>
                            <asp:TableCell></asp:TableCell>
                            <asp:TableCell></asp:TableCell>
                            <asp:TableCell></asp:TableCell>
                        </asp:TableRow>
                        <asp:TableRow>
                            <asp:TableCell ColumnSpan="4">
                                <asp:Label runat="server" font-size="XX-Small"><%#DataBinder.Eval(Container.DataItem, "sCneeTown")%></asp:Label>
                            </asp:TableCell>
                            <asp:TableCell></asp:TableCell>
                            <asp:TableCell></asp:TableCell>
                            <asp:TableCell></asp:TableCell>
                            <asp:TableCell></asp:TableCell>
                        </asp:TableRow>
                        <asp:TableRow>
                            <asp:TableCell ColumnSpan="4">
                                <asp:Label runat="server" font-size="XX-Small"><%# DataBinder.Eval(Container.DataItem,"sCneeCountry") %></asp:Label>
                            </asp:TableCell>
                            <asp:TableCell></asp:TableCell>
                            <asp:TableCell></asp:TableCell>
                            <asp:TableCell></asp:TableCell>
                            <asp:TableCell></asp:TableCell>
                        </asp:TableRow>
                        <asp:TableRow>
                            <asp:TableCell ColumnSpan="4">
                                <asp:Label runat="server" font-size="XX-Small"><%# DataBinder.Eval(Container.DataItem,"sCneeCtcName") %></asp:Label>
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
                    <asp:DataGrid id="dgProducts" OnItemDataBound="dgProducts_ItemDataBound" runat="server" DataSource='<%# DataBinder.Eval(Container.DataItem,"dblStockItemList") %>' AutoGenerateColumns="False" Font-Names="Verdana" Font-Size="XX-Small" Width="650px" GridLines="None">
                        <Columns>
                            <asp:BoundColumn DataField="sProdCode" HeaderText="Product Code">
                                <HeaderStyle font-bold="True" width="90px"></HeaderStyle>
                                <ItemStyle verticalalign="Top"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="sProdDate" HeaderText="Product Date">
                                <HeaderStyle font-bold="True" width="100px"></HeaderStyle>
                                <ItemStyle verticalalign="Top"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="sProdDescription" HeaderText="Product Description">
                                <HeaderStyle font-bold="True" width="250px"></HeaderStyle>
                                <ItemStyle verticalalign="Top"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="nQuantity" HeaderText="Quantity">
                                <HeaderStyle font-bold="True" horizontalalign="Right" width="70px"></HeaderStyle>
                                <ItemStyle horizontalalign="Right" verticalalign="Top"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="dblUnitValue" HeaderText="Unit Cost" DataFormatString="{0:#,##0.00}">
                                <HeaderStyle font-bold="True" horizontalalign="Right" width="70px"></HeaderStyle>
                                <ItemStyle horizontalalign="Right" verticalalign="Top"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:TemplateColumn HeaderText="Value (€)">
                                <HeaderStyle font-bold="True" horizontalalign="Right" width="70px"></HeaderStyle>
                                <ItemStyle horizontalalign="Right" verticalalign="Top"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Label runat="server"><%# Format((DataBinder.Eval(Container.DataItem, "nQuantity")) * (DataBinder.Eval(Container.DataItem, "dblUnitValue")),"#,##0.00") %></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                        </Columns>
                    </asp:DataGrid>
                    <asp:Table ID="tbl001" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="650px">
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
                                <asp:Label runat="server" font-size="X-Small" Visible="<%# IsNotHyster() %>">Total value this order (£)</asp:Label><asp:Label runat="server" font-size="X-Small" Visible="<%# IsHyster() %>">Total value this order (€)</asp:Label>
                            </asp:TableCell>
                            <asp:TableCell horizontalalign="Right" ColumnSpan="1">
                                <asp:Label runat="server" font-size="X-Small" font-bold="True"><%# Format(DataBinder.Eval(Container.DataItem,"dblStockItemListValue"),"#,##0.00") %></asp:Label>
                            </asp:TableCell>
                        </asp:TableRow>
                        <asp:TableRow>
                            <asp:TableCell ColumnSpan="3"></asp:TableCell>
                            <asp:TableCell horizontalalign="Right" ColumnSpan="3">
                                <asp:Label runat="server" font-size="X-Small">Shipping costs this order (£)</asp:Label>
                            </asp:TableCell>
                            <asp:TableCell horizontalalign="Right" ColumnSpan="1">
                                <asp:Label runat="server" font-size="X-Small" font-bold="True"><%#Format(DataBinder.Eval(Container.DataItem, "dblShippingCost"), "#,##0.00")%></asp:Label>
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
            </asp:DataList>&nbsp;
            <asp:Label ID="lblReportGeneratedDateTime" runat="server" Font-Names="Verdana, Sans-Serif"
                Font-Size="XX-Small" ForeColor="Green" Text="" Visible="false"></asp:Label>
            <asp:Label ID="lblResult" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                ForeColor="Red" Text="no consignments found"></asp:Label></asp:Panel>
    </form>
</body>
</html>
