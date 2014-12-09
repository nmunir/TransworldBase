<%@ Page Language="VB" Theme="AIMSDefault" %>

<%@ Import Namespace="System.Data " %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Globalization" %>
<%@ Import Namespace="System.Threading" %>
<%@ Import Namespace="System.Collections.Generic" %>
<script type="text/VB" runat="server">

    '   Consignments By Consignee Report
    
    Const STYLENAME_CALENDAR As String = "calendar style dates"
    Const STYLENAME_DROPDOWN As String = "dropdown style dates"

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
            lnkbtnToggleSelectionStyle1.Text = STYLENAME_DROPDOWN
            lnkbtnToggleSelectionStyle2.Text = STYLENAME_DROPDOWN

        End If
    End Sub
    
    Protected Sub lnkbtnToggleSelectionStyle_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If CalendarInterface.Visible = True Then
            CalendarInterface.Visible = False
            DropdownInterface.Visible = True
            If Page.IsValid Then
                Dim dDate As Date
                Dim nVal As Integer
                If IsDate(tbFromDate.Text) Then
                    dDate = Date.Parse(tbFromDate.Text)
                    nVal = dDate.Day
                    ddlFromDay.SelectedIndex = nVal - 1
                    nVal = CStr(dDate.Month)
                    ddlFromMonth.SelectedIndex = nVal - 1
                    nVal = CStr(dDate.Year)
                    For i As Integer = 0 To ddlFromYear.Items.Count - 1
                        If ddlFromYear.Items(i).Text = CStr(nVal) Then
                            ddlFromYear.SelectedIndex = i
                            Exit For
                        End If
                    Next
                End If

                If IsDate(tbToDate.Text) Then
                    dDate = Date.Parse(tbToDate.Text)
                    nVal = dDate.Day
                    ddlToDay.SelectedIndex = nVal - 1
                    nVal = CStr(dDate.Month)
                    ddlToMonth.SelectedIndex = nVal - 1
                    nVal = CStr(dDate.Year)
                    For i As Integer = 0 To ddlToYear.Items.Count - 1
                        If ddlToYear.Items(i).Text = CStr(nVal) Then
                            ddlToYear.SelectedIndex = i
                            Exit For
                        End If
                    Next
                End If
            End If
        Else
            CalendarInterface.Visible = True
            DropdownInterface.Visible = False
            Dim arrMonths() As String = {"Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"}
            If IsDate(ddlFromDay.SelectedItem.Text & ddlFromMonth.SelectedItem.Text & ddlFromYear.SelectedItem.Text) Then
                tbFromDate.Text = ddlFromDay.SelectedValue & "-" & arrMonths(ddlFromMonth.SelectedIndex) & "-" & ddlFromYear.SelectedValue
            End If
            If IsDate(ddlToDay.SelectedItem.Text & ddlToMonth.SelectedItem.Text & ddlToYear.SelectedItem.Text) Then
                tbToDate.Text = ddlToDay.SelectedValue & "-" & arrMonths(ddlToMonth.SelectedIndex) & "-" & ddlToYear.SelectedValue
            End If
        End If
        lblFromErrorMessage.Text = ""
        lblToErrorMessage.Text = ""
        If lnkbtnToggleSelectionStyle1.Text = STYLENAME_CALENDAR Then
            lnkbtnToggleSelectionStyle1.Text = STYLENAME_DROPDOWN
            lnkbtnToggleSelectionStyle2.Text = STYLENAME_DROPDOWN
        Else
            lnkbtnToggleSelectionStyle1.Text = STYLENAME_CALENDAR
            lnkbtnToggleSelectionStyle2.Text = STYLENAME_CALENDAR
        End If
        
        btnExportByConsignment.Visible = True
        btnExportByProduct.Visible = True
        
    End Sub
    
    
    Public Sub ExportToExcelByConsignment()
        
        Dim fileName As String = "Consignment_List_ByConsignment_" & DateTime.Now.ToString("yyyymmddhhmmss") & ".csv"
        
        Dim Source As ConsignmentList = ConsignmentList.GetConsignmentList(sFromDate, sToDate, Session("UserKey"), Session("CustomerKey"), pnSelectedProductGroup)
        Response.Clear()
        Response.ContentType = "text/csv"
        Response.AddHeader("Content-Disposition", "attachment; filename=" & fileName)
        
        Response.Write(ControlChars.Quote & "Consignment" & ControlChars.Quote)
        Response.Write(",")
        Response.Write(ControlChars.Quote & "LogisticBookingKey" & ControlChars.Quote)
        Response.Write(",")
        Response.Write(ControlChars.Quote & "CneeName" & ControlChars.Quote)
        Response.Write(",")
        Response.Write(ControlChars.Quote & "Address1" & ControlChars.Quote)
        Response.Write(",")
        Response.Write(ControlChars.Quote & "Address2" & ControlChars.Quote)
        Response.Write(",")
        Response.Write(ControlChars.Quote & "Address3" & ControlChars.Quote)
        Response.Write(",")
        Response.Write(ControlChars.Quote & "Town" & ControlChars.Quote)
        Response.Write(",")
        Response.Write(ControlChars.Quote & "State" & ControlChars.Quote)
        Response.Write(",")
        Response.Write(ControlChars.Quote & "Country" & ControlChars.Quote)
        Response.Write(",")
        Response.Write(ControlChars.Quote & "PostCode" & ControlChars.Quote)
        Response.Write(",")
        Response.Write(ControlChars.Quote & "Telephone" & ControlChars.Quote)
        Response.Write(",")
        Response.Write(ControlChars.Quote & "Customer Ref1" & ControlChars.Quote)
        Response.Write(",")
        Response.Write(ControlChars.Quote & "Customer Ref2" & ControlChars.Quote)
        Response.Write(",")
        Response.Write(ControlChars.Quote & "Customer Ref3" & ControlChars.Quote)
        Response.Write(",")
        Response.Write(ControlChars.Quote & "CustomerRef4" & ControlChars.Quote)
        Response.Write(",")
        Response.Write(ControlChars.Quote & "POD" & ControlChars.Quote)
        Response.Write(",")
        Response.Write(ControlChars.Quote & "POD Date" & ControlChars.Quote)
        Response.Write(",")
        Response.Write(ControlChars.Quote & "POD Time" & ControlChars.Quote)
        Response.Write(",")
        Response.Write(ControlChars.Quote & "AWB" & ControlChars.Quote)
        Response.Write(",")
        Response.Write(ControlChars.Quote & "Weight" & ControlChars.Quote)
        Response.Write(",")
        Response.Write(ControlChars.Quote & "Shipping Cost" & ControlChars.Quote)
        Response.Write(",")
        Response.Write(ControlChars.Quote & "Product Code" & ControlChars.Quote)
        Response.Write(",")
        Response.Write(ControlChars.Quote & "Product Description" & ControlChars.Quote)
        Response.Write(",")
        Response.Write(ControlChars.Quote & "Quantity" & ControlChars.Quote)
        Response.Write(",")
        Response.Write(ControlChars.Quote & "Unit Value" & ControlChars.Quote)
        Response.Write(",")
        Response.Write(ControlChars.Quote & "Total Value" & ControlChars.Quote)
        Response.Write(vbCrLf)
        
        For Each Consignment As Consignment In Source
            
            For Each _stockItem As StockItem In Consignment.dblStockItemList
                
                Response.Write(ControlChars.Quote & Consignment.lConsignmentKey & ControlChars.Quote)
                Response.Write(",")
                Response.Write(ControlChars.Quote & Consignment.lLogisticBookingKey & ControlChars.Quote)
                Response.Write(",")
                Response.Write(ControlChars.Quote & Consignment.sCneeName & ControlChars.Quote)
                Response.Write(",")
                Response.Write(ControlChars.Quote & Consignment.sCneeAddr1 & ControlChars.Quote)
                Response.Write(",")
                Response.Write(ControlChars.Quote & Consignment.sCneeAddr2 & ControlChars.Quote)
                Response.Write(",")
                Response.Write(ControlChars.Quote & Consignment.sCneeAddr3 & ControlChars.Quote)
                Response.Write(",")
                Response.Write(ControlChars.Quote & Consignment.sCneeTown & ControlChars.Quote)
                Response.Write(",")
                Response.Write(ControlChars.Quote & Consignment.sCneeState & ControlChars.Quote)
                Response.Write(",")
                Response.Write(ControlChars.Quote & Consignment.sCneeCountry & ControlChars.Quote)
                Response.Write(",")
                Response.Write(ControlChars.Quote & Consignment.sCneePostCode & ControlChars.Quote)
                Response.Write(",")
                Response.Write(ControlChars.Quote & Consignment.sCneeTel & ControlChars.Quote)
                Response.Write(",")
                Response.Write(ControlChars.Quote & Consignment.sCustomerRef1 & ControlChars.Quote)
                Response.Write(",")
                Response.Write(ControlChars.Quote & Consignment.sCustomerRef2 & ControlChars.Quote)
                Response.Write(",")
                Response.Write(ControlChars.Quote & Consignment.sMisc1 & ControlChars.Quote)
                Response.Write(",")
                Response.Write(ControlChars.Quote & Consignment.sMisc2 & ControlChars.Quote)
                Response.Write(",")
                Response.Write(ControlChars.Quote & Consignment.sPODName & ControlChars.Quote)
                Response.Write(",")
                Response.Write(ControlChars.Quote & Consignment.sPODDate & ControlChars.Quote)
                Response.Write(",")
                Response.Write(ControlChars.Quote & Consignment.sPODTime & ControlChars.Quote)
                Response.Write(",")
                Response.Write(ControlChars.Quote & Consignment.sAWB & ControlChars.Quote)
                Response.Write(",")
                Response.Write(ControlChars.Quote & Consignment.dblWeight & ControlChars.Quote)
                Response.Write(",")
                Response.Write(ControlChars.Quote & Consignment.dblShippingCost & ControlChars.Quote)
                Response.Write(",")
                Response.Write(ControlChars.Quote & _stockItem.sProdCode & ControlChars.Quote)
                Response.Write(",")
                Response.Write(ControlChars.Quote & _stockItem.sProdDescription & ControlChars.Quote)
                Response.Write(",")
                Response.Write(ControlChars.Quote & _stockItem.nQuantity.ToString & ControlChars.Quote)
                Response.Write(",")
                Response.Write(ControlChars.Quote & _stockItem.dblUnitValue.ToString & ControlChars.Quote)
                Response.Write(",")
                Dim nValue As Decimal = _stockItem.nQuantity * _stockItem.dblUnitValue
                Response.Write(ControlChars.Quote & nValue.ToString & ControlChars.Quote)
                Response.Write(",")
                Response.Write(vbCrLf)
                
            Next
            
            
            
        Next
        
        Response.End()
        
    End Sub
    
    Public Sub ExportToExcelByProduct()
        
        Dim fileName As String = "Consignment_List_ByProduct_" & DateTime.Now.ToString("yyyymmddhhmmss") & ".csv"
        
        Dim Source As ConsignmentList = ConsignmentList.GetConsignmentList(sFromDate, sToDate, Session("UserKey"), Session("CustomerKey"), pnSelectedProductGroup)
        Response.Clear()
        Response.ContentType = "text/csv"
        Response.AddHeader("Content-Disposition", "attachment; filename=" & fileName)
        
        Response.Write(ControlChars.Quote & "Consignment" & ControlChars.Quote)
        Response.Write(",")
        Response.Write(ControlChars.Quote & "LogisticBookingKey" & ControlChars.Quote)
        Response.Write(",")
        Response.Write(ControlChars.Quote & "CneeName" & ControlChars.Quote)
        Response.Write(",")
        Response.Write(ControlChars.Quote & "Address1" & ControlChars.Quote)
        Response.Write(",")
        Response.Write(ControlChars.Quote & "Address2" & ControlChars.Quote)
        Response.Write(",")
        Response.Write(ControlChars.Quote & "Address3" & ControlChars.Quote)
        Response.Write(",")
        Response.Write(ControlChars.Quote & "Town" & ControlChars.Quote)
        Response.Write(",")
        Response.Write(ControlChars.Quote & "State" & ControlChars.Quote)
        Response.Write(",")
        Response.Write(ControlChars.Quote & "Country" & ControlChars.Quote)
        Response.Write(",")
        Response.Write(ControlChars.Quote & "PostCode" & ControlChars.Quote)
        Response.Write(",")
        Response.Write(ControlChars.Quote & "Telephone" & ControlChars.Quote)
        Response.Write(",")
        Response.Write(ControlChars.Quote & "Customer Ref1" & ControlChars.Quote)
        Response.Write(",")
        Response.Write(ControlChars.Quote & "Customer Ref2" & ControlChars.Quote)
        Response.Write(",")
        Response.Write(ControlChars.Quote & "Customer Ref3" & ControlChars.Quote)
        Response.Write(",")
        Response.Write(ControlChars.Quote & "CustomerRef4" & ControlChars.Quote)
        Response.Write(",")
        Response.Write(ControlChars.Quote & "POD" & ControlChars.Quote)
        Response.Write(",")
        Response.Write(ControlChars.Quote & "POD Date" & ControlChars.Quote)
        Response.Write(",")
        Response.Write(ControlChars.Quote & "POD Time" & ControlChars.Quote)
        Response.Write(",")
        Response.Write(ControlChars.Quote & "AWB" & ControlChars.Quote)
        Response.Write(",")
        Response.Write(ControlChars.Quote & "Weight" & ControlChars.Quote)
        Response.Write(",")
        Response.Write(ControlChars.Quote & "Shipping Cost" & ControlChars.Quote)
        Response.Write(",")
        Response.Write("Product Code")
        Response.Write(";")
        Response.Write("Product Description")
        Response.Write(";")
        Response.Write("Quantity")
        Response.Write(";")
        Response.Write("Unit Value")
        Response.Write(";")
        Response.Write("Total Value")
        Response.Write(vbCrLf)
        
        For Each Consignment As Consignment In Source
            
            Response.Write(ControlChars.Quote & Consignment.lConsignmentKey & ControlChars.Quote)
            Response.Write(",")
            Response.Write(ControlChars.Quote & Consignment.lLogisticBookingKey & ControlChars.Quote)
            Response.Write(",")
            Response.Write(ControlChars.Quote & Consignment.sCneeName & ControlChars.Quote)
            Response.Write(",")
            Response.Write(ControlChars.Quote & Consignment.sCneeAddr1 & ControlChars.Quote)
            Response.Write(",")
            Response.Write(ControlChars.Quote & Consignment.sCneeAddr2 & ControlChars.Quote)
            Response.Write(",")
            Response.Write(ControlChars.Quote & Consignment.sCneeAddr3 & ControlChars.Quote)
            Response.Write(",")
            Response.Write(ControlChars.Quote & Consignment.sCneeTown & ControlChars.Quote)
            Response.Write(",")
            Response.Write(ControlChars.Quote & Consignment.sCneeState & ControlChars.Quote)
            Response.Write(",")
            Response.Write(ControlChars.Quote & Consignment.sCneeCountry & ControlChars.Quote)
            Response.Write(",")
            Response.Write(ControlChars.Quote & Consignment.sCneePostCode & ControlChars.Quote)
            Response.Write(",")
            Response.Write(ControlChars.Quote & Consignment.sCneeTel & ControlChars.Quote)
            Response.Write(",")
            Response.Write(ControlChars.Quote & Consignment.sCustomerRef1 & ControlChars.Quote)
            Response.Write(",")
            Response.Write(ControlChars.Quote & Consignment.sCustomerRef2 & ControlChars.Quote)
            Response.Write(",")
            Response.Write(ControlChars.Quote & Consignment.sMisc1 & ControlChars.Quote)
            Response.Write(",")
            Response.Write(ControlChars.Quote & Consignment.sMisc2 & ControlChars.Quote)
            Response.Write(",")
            Response.Write(ControlChars.Quote & Consignment.sPODName & ControlChars.Quote)
            Response.Write(",")
            Response.Write(ControlChars.Quote & Consignment.sPODDate & ControlChars.Quote)
            Response.Write(",")
            Response.Write(ControlChars.Quote & Consignment.sPODTime & ControlChars.Quote)
            Response.Write(",")
            Response.Write(ControlChars.Quote & Consignment.sAWB & ControlChars.Quote)
            Response.Write(",")
            Response.Write(ControlChars.Quote & Consignment.dblWeight & ControlChars.Quote)
            Response.Write(",")
            Response.Write(ControlChars.Quote & Consignment.dblShippingCost & ControlChars.Quote)
            Response.Write(",")
            
            Dim sb As New StringBuilder
            
            For Each _stockItem As StockItem In Consignment.dblStockItemList
                
                sb.Append(_stockItem.sProdCode)
                sb.Append(";")
                sb.Append(_stockItem.sProdDescription)
                sb.Append(";")
                sb.Append(_stockItem.nQuantity.ToString)
                sb.Append(";")
                sb.Append(_stockItem.dblUnitValue.ToString)
                sb.Append(";")
                Dim nValue As Decimal = _stockItem.nQuantity * _stockItem.dblUnitValue
                sb.Append(nValue.ToString)
                sb.Append("|")
                
            Next
            
            Dim finalstr As String = sb.ToString
            
            finalstr = finalstr.Replace(ControlChars.Quote, ControlChars.Quote & ControlChars.Quote)
            Response.Write(ControlChars.Quote & finalstr & ControlChars.Quote)
            Response.Write(vbCrLf)
            
        Next
        
        Response.End()
           
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
        
        If DropdownInterface.Visible Then
            
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
            
        ElseIf CalendarInterface.Visible Then
            Page.Validate("CalendarInterface")
            sFromDate = tbFromDate.Text.Trim
            sToDate = tbToDate.Text.Trim
            bIsValid = True
            
        Else
            bIsValid = False
        End If
        
       
        Return bIsValid
      
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
                ' I (Muhammad) have added distinct keyword to prevent duplication of records
                sbSQL.Append("SELECT distinct awb.[Key] ConsignmentKey, lb.LogisticBookingKey, awb.AWB, awb.CreatedOn ShipDate, ")
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
                sbSQL.Append(" AND awb.CreatedOn BETWEEN '" & sFromDate & "' AND '" & sToDate & " 23:59:00")
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
                Return Nothing
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
    
    Protected Sub dlConsignments_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.DataListItemEventArgs)
        
        If e.Item.ItemType = ListItemType.Item Or e.Item.ItemType = ListItemType.AlternatingItem Then
                
            Dim trCneeAddr2 As TableRow = e.Item.FindControl("trCneeAddr2")
            Dim trCneeAddr3 As TableRow = e.Item.FindControl("trCneeAddr3")
            Dim trCustomerRef1 As TableRow = e.Item.FindControl("trCustomerRef1")
            Dim trCustomerRef2 As TableRow = e.Item.FindControl("trCustomerRef2")
            Dim trMisc1 As TableRow = e.Item.FindControl("trMisc1")
            Dim trMisc2 As TableRow = e.Item.FindControl("trMisc2")

            
                
            Dim lblCneeAddr2 As Label = e.Item.FindControl("lblCneeAddr2")
            Dim lblCneeAddr3 As Label = e.Item.FindControl("lblCneeAddr3")
            Dim lblCustomerRef1 As Label = e.Item.FindControl("lblCustomerRef1")
            Dim lblCustomerRef2 As Label = e.Item.FindControl("lblCustomerRef2")
            Dim lblMisc1 As Label = e.Item.FindControl("lblMisc1")
            Dim lblMisc2 As Label = e.Item.FindControl("lblMisc2")
                
            If lblCneeAddr2.Text.Trim <> String.Empty Then
                trCneeAddr2.Visible = True
            End If
                
            If lblCneeAddr3.Text.Trim <> String.Empty Then
                trCneeAddr3.Visible = True
            End If
            
            If lblCustomerRef1.Text.Trim <> String.Empty Then
                trCustomerRef1.Visible = True
            Else
                trCustomerRef1.Visible = False
            End If
            
            If lblCustomerRef2.Text.Trim <> String.Empty Then
                trCustomerRef2.Visible = True
            Else
                trCustomerRef2.Visible = False
            End If
            
            If lblMisc1.Text.Trim <> String.Empty Then
                trMisc1.Visible = True
            Else
                trMisc1.Visible = False
            End If
            
            If lblMisc2.Text.Trim <> String.Empty Then
                trMisc2.Visible = True
            Else
                trMisc2.Visible = False
            End If
            
                
        End If

    End Sub

    Protected Sub btnExportByConsignment_Click(sender As Object, e As System.EventArgs)

        Call ExportToExcelByProduct()
        
    End Sub

    Protected Sub btnExportByProduct_Click(sender As Object, e As System.EventArgs)
        
        Call ExportToExcelByConsignment()
        
    End Sub
    
    Protected Sub btnReselectReportFilter_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ReselectReportFilter()
    End Sub
    
    Protected Sub ReselectReportFilter()
        btnRunReport1.Visible = True
        'btnRunReport2.Visible = True
        btnReselectReportFilter1.Visible = False
        'btnReselectReportFilter2.Visible = False
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
        
        ddlProductGroup.Enabled = True
        
        'pnlData.Visible = False
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
            <td style="width: 50%; white-space: nowrap">
                <asp:Label ID="Label1" runat="server" ForeColor="silver" Font-Size="Small" Font-Bold="True"
                    Font-Names="Arial">Consignments
                      By Consignee Report</asp:Label><br />
                <br />
            </td>
            <td style="width: 50%; white-space: nowrap" align="right">
            </td>
        </tr>
    </table>
    <asp:Panel ID="pnlReportCriteria" runat="server" Visible="False">
        <table style="width: 650px; font-family: Verdana">
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
                    </asp:DropDownList>
                    <asp:Label ID="lblProductGroup" runat="server" Font-Names="Verdana" Font-Size="X-Small"
                        Font-Bold="True"></asp:Label>
                </td>
                <td align="right" style="width: 10%; white-space: nowrap" valign="top">
                    <asp:Button ID="btnShowProductGroups" runat="server" OnClick="btnShowProductGroups_Click"
                        Text="show product groups" Visible="False" />
                </td>
            </tr>
            <tr>
                <td align="right" style="width: 40%; white-space: nowrap" valign="top">
                </td>
                <td align="right" style="width: 40%; white-space: nowrap" valign="top">
                </td>
                <td style="width: 10%; white-space: nowrap" valign="top">
                </td>
            </tr>
            <tr runat="server" visible="true" id="CalendarInterface">
                <td style="width: 190px; white-space: nowrap">
                    <span class="informational dark">&nbsp;From:</span>
                    <asp:TextBox ID="tbFromDate" Font-Names="Verdana" Font-Size="XX-Small" Width="90"
                        runat="server" />
                    <a id="imgCalendarButton1" runat="server" visible="true" href="javascript:;" onclick="window.open('../PopupCalendar4.aspx?textbox=tbFromDate','cal','width=300,height=305,left=270,top=180')">
                        <img id="Img1" src="../images/SmallCalendar.gif" runat="server" border="0" alt=""
                            ie:visible="true" visible="false" /></a><span id="spnDateExample1" runat="server"
                                visible="true" class="informational light" style="white-space: nowrap">(eg&nbsp;12-Jan-2011)</span>
                </td>
                <td style="white-space: nowrap; width: 190px">
                    <span class="informational dark">To:</span>
                    <asp:TextBox ID="tbToDate" Font-Names="Verdana" Font-Size="XX-Small" Width="90" runat="server" />
                    <a id="imgCalendarButton2" runat="server" visible="true" href="javascript:;" onclick="window.open('../PopupCalendar4.aspx?textbox=tbToDate','cal','width=300,height=305,left=270,top=180')">
                        <img id="Img2" src="../images/SmallCalendar.gif" runat="server" border="0" alt=""
                            ie:visible="true" visible="false" /></a> <span id="spnDateExample2" runat="server"
                                visible="true" class="informational light" style="white-space: nowrap">(eg&nbsp;12-Jan-2012)</span>
                </td>
                <td>
                    <asp:Button ID="btnRunReport1" runat="server" Text="generate report" Visible="true"
                        OnClick="btnRunReport_Click" Width="180px" />
                    <asp:LinkButton ID="lnkbtnToggleSelectionStyle1" runat="server" OnClick="lnkbtnToggleSelectionStyle_Click"
                        ToolTip="toggles between calendar interface and dropdown interface"></asp:LinkButton>
                    <asp:Button ID="btnReselectReportFilter1" runat="server" Text="re-select report filter"
                        Visible="false" OnClick="btnReselectReportFilter_Click" />
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                </td>
            </tr>
            <tr runat="server" visible="false" id="DropdownInterface">
                <td style="width: 265px">
                    <span class="informational dark">&nbsp;From:</span> &nbsp;<asp:DropDownList ID="ddlFromDay"
                        runat="server" Font-Names="Verdana" Font-Size="XX-Small">
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
                    </asp:DropDownList>
                    &nbsp;<asp:DropDownList ID="ddlFromMonth" runat="server" Font-Names="Verdana" Font-Size="XX-Small">
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
                    </asp:DropDownList>
                    &nbsp;<asp:DropDownList ID="ddlFromYear" runat="server" Font-Names="Verdana" Font-Size="XX-Small">
                    </asp:DropDownList>
                    &nbsp;
                </td>
                <td style="width: 265px">
                    <span class="informational dark">To:</span> &nbsp;<asp:DropDownList ID="ddlToDay"
                        runat="server" Font-Names="Verdana" Font-Size="XX-Small">
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
                    </asp:DropDownList>
                    &nbsp;<asp:DropDownList ID="ddlToMonth" runat="server" Font-Names="Verdana" Font-Size="XX-Small">
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
                    </asp:DropDownList>
                    &nbsp;<asp:DropDownList ID="ddlToYear" runat="server" Font-Names="Verdana" Font-Size="XX-Small">
                    </asp:DropDownList>
                    &nbsp;
                </td>
                <td style="width: 350px">
                    <asp:Button ID="btnRunReport" runat="server" Text="generate report" OnClick="btnRunReport_Click"
                        Width="170px" />
                    <asp:LinkButton ID="lnkbtnToggleSelectionStyle2" runat="server" OnClick="lnkbtnToggleSelectionStyle_Click"
                        ToolTip="toggles between easy-to-use calendar interface and clunky dropdown interface"></asp:LinkButton>
                    <asp:Button ID="btnReselectReportFilter2" runat="server" Text="re-select report filter"
                        Visible="false" OnClick="btnReselectReportFilter_Click" />
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                </td>
                <td style="width: 169px">
                </td>
            </tr>
            <tr>
                <td>
                    <asp:Button ID="btnExportByConsignment" Text="export to excel row / consignment"
                        Visible="true" runat="server" OnClick="btnExportByConsignment_Click" Width="230px" />
                </td>
                <td>
                    <asp:Button ID="btnExportByProduct" Text="export to excel row / product" runat="server"
                        Visible="true" OnClick="btnExportByProduct_Click" Width="230px" />
                </td>
            </tr>
            <tr runat="server" visible="true" id="DateValidationMessages">
                <td style="width: 265px">
                    <asp:RegularExpressionValidator ID="revFromDate" runat="server" ControlToValidate="tbFromDate"
                        ErrorMessage=" - invalid format for expiry date - use dd-mmm-yy" Font-Names="Verdana"
                        Font-Size="XX-Small" ValidationExpression="^\d\d-(jan|Jan|JAN|feb|Feb|FEB|mar|Mar|MAR|apr|Apr|APR|may|May|MAY|jun|Jun|JUN|jul|Jul|JUL|aug|Aug|AUG|sep|Sep|SEP|oct|Oct|OCT|nov|Nov|NOV|dec|Dec|DEC)-\d(\d+)"
                        SetFocusOnError="True" ValidationGroup="CalendarInterface"></asp:RegularExpressionValidator>
                    <asp:RangeValidator ID="rvFromDate" runat="server" ControlToValidate="tbFromDate"
                        CultureInvariantValues="True" ErrorMessage=" - expiry year before 2000, after 2020, or not a valid date!"
                        Font-Names="Verdana" Font-Size="XX-Small" MaximumValue="2019/1/1" MinimumValue="2000/1/1"
                        ValidationGroup="CalendarInterface" EnableClientScript="False" Type="Date"></asp:RangeValidator>
                    <asp:Label ID="lblFromErrorMessage" runat="server" Font-Names="Verdana,Sans-Serif"
                        Font-Size="XX-Small" ForeColor="Red"></asp:Label>
                </td>
                <td style="width: 265px">
                    <asp:RegularExpressionValidator ID="RegularevToDate" runat="server" ControlToValidate="tbToDate"
                        ErrorMessage=" - invalid format for expiry date - use dd-mmm-yy" Font-Names="Verdana"
                        Font-Size="XX-Small" ValidationExpression="^\d\d-(jan|Jan|JAN|feb|Feb|FEB|mar|Mar|MAR|apr|Apr|APR|may|May|MAY|jun|Jun|JUN|jul|Jul|JUL|aug|Aug|AUG|sep|Sep|SEP|oct|Oct|OCT|nov|Nov|NOV|dec|Dec|DEC)-\d(\d+)"
                        ValidationGroup="CalendarInterface"></asp:RegularExpressionValidator><asp:RangeValidator
                            ID="rvToDate" runat="server" ControlToValidate="tbToDate" CultureInvariantValues="True"
                            ErrorMessage=" - expiry year before 2000, after 2020, or not a valid date!" Font-Names="Verdana"
                            Font-Size="XX-Small" MaximumValue="2019/1/1" MinimumValue="2000/1/1" ValidationGroup="CalendarInterface"
                            EnableClientScript="False" Type="Date"></asp:RangeValidator>
                    <asp:Label ID="lblToErrorMessage" runat="server" Font-Names="Verdana,Sans-Serif"
                        Font-Size="XX-Small" ForeColor="Red"></asp:Label>
                </td>
                <td style="width: 253px">
                </td>
                <td style="width: 169px">
                </td>
            </tr>
            <tr>
                <td colspan="3" align="right">
                    <asp:Label ID="lblDateError" runat="server" ForeColor="Red" Font-Size="XX-Small"></asp:Label>
                </td>
            </tr>
        </table>
    </asp:Panel>
    <asp:Panel ID="pnlReportData" runat="server" Visible="False" Width="100%">
        <asp:DataList ID="dlConsignments" runat="server" EnableViewState="False">
            <ItemTemplate>
                <asp:Table ID="Table1" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="650px">
                    <asp:TableRow>
                        <asp:TableCell ColumnSpan="8">
                                <hr />
                        </asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell Width="50px"></asp:TableCell>
                        <asp:TableCell Width="100px"></asp:TableCell>
                        <asp:TableCell Width="100px"></asp:TableCell>
                        <asp:TableCell Width="75px"></asp:TableCell>
                        <asp:TableCell Width="50px"></asp:TableCell>
                        <asp:TableCell Width="100px"></asp:TableCell>
                        <asp:TableCell Width="100px"></asp:TableCell>
                        <asp:TableCell Width="75px"></asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell ColumnSpan="4">
                            <asp:Label ID="Label2" runat="server" Font-Size="X-Small">Stock Booking No</asp:Label>
                            &nbsp;<asp:Label ID="Label3" runat="server" Font-Size="X-Small" Font-Bold="True"><%#Format(DataBinder.Eval(Container.DataItem, "lLogisticBookingKey"), "0000000")%></asp:Label>
                        </asp:TableCell>
                        <asp:TableCell ColumnSpan="4" HorizontalAlign="Right">
                            <asp:Label ID="Label4" runat="server" Font-Size="X-Small">Booked On</asp:Label>
                            &nbsp;<asp:Label ID="Label5" runat="server" Font-Size="X-Small" Font-Bold="True"><%#Format(DataBinder.Eval(Container.DataItem, "dtShipDate"), "dd MMM yyyy HH:mm")%></asp:Label>
                        </asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell ColumnSpan="4">
                            <asp:Label ID="Label6" runat="server" Font-Size="X-Small">Consignee</asp:Label>
                            &nbsp;<asp:Label ID="Label7" runat="server" Font-Size="X-Small" Font-Bold="True"><%# DataBinder.Eval(Container.DataItem,"sCneeName") %></asp:Label>
                            &nbsp;
                        </asp:TableCell>
                        <asp:TableCell ColumnSpan="4" HorizontalAlign="Right">
                            <asp:Label ID="Label8" runat="server" Font-Size="X-Small">Booked By</asp:Label>
                            &nbsp;<asp:Label ID="Label9" runat="server" Font-Size="X-Small" Font-Bold="True"><%# DataBinder.Eval(Container.DataItem,"sBookedBy") %></asp:Label>
                        </asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell ColumnSpan="4">
                            <asp:Label ID="Label10" runat="server" Font-Size="X-Small">Air Waybill</asp:Label>
                            &nbsp;<asp:Label ID="Label11" runat="server" Font-Size="X-Small" Font-Bold="True"><%#(DataBinder.Eval(Container.DataItem, "sAWB"))%></asp:Label>
                            &nbsp;<asp:Label ID="Label12" runat="server" Font-Size="X-Small">/</asp:Label>
                            &nbsp;<asp:Label ID="Label13" runat="server" Font-Size="X-Small"><%#(DataBinder.Eval(Container.DataItem, "nNOP")) & " @ " & Format((DataBinder.Eval(Container.DataItem, "dblWeight")), "#,##0.0")%></asp:Label>
                        </asp:TableCell>
                        <asp:TableCell ColumnSpan="4" HorizontalAlign="Right">
                            <asp:Label ID="Label14" runat="server" Font-Size="X-Small">Delivered To</asp:Label>
                            &nbsp;<asp:Label ID="Label15" runat="server" Font-Size="X-Small" Font-Bold="True"><%# DataBinder.Eval(Container.DataItem,"sPODName")%></asp:Label>
                            &nbsp;<asp:Label ID="Label16" runat="server" Font-Size="X-Small"><%#DataBinder.Eval(Container.DataItem, "sPODDate") & " " & DataBinder.Eval(Container.DataItem, "sPODTime")%></asp:Label>
                        </asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell ColumnSpan="4">
                            <br />
                            <asp:Label ID="Label17" runat="server" Font-Size="X-Small"><%#DataBinder.Eval(Container.DataItem, "sCneeName")%></asp:Label>
                        </asp:TableCell>
                        <asp:TableCell></asp:TableCell>
                        <asp:TableCell></asp:TableCell>
                        <asp:TableCell></asp:TableCell>
                        <asp:TableCell></asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell ColumnSpan="4">
                            <asp:Label ID="Label18" runat="server" Font-Size="XX-Small"><%#DataBinder.Eval(Container.DataItem, "sCneeAddr1")%></asp:Label>
                        </asp:TableCell>
                        <asp:TableCell></asp:TableCell>
                        <asp:TableCell></asp:TableCell>
                        <asp:TableCell></asp:TableCell>
                        <asp:TableCell></asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell ColumnSpan="4">
                            <asp:Label ID="Label19" runat="server" Font-Size="XX-Small"><%#DataBinder.Eval(Container.DataItem, "sCneeTown")%></asp:Label>
                        </asp:TableCell>
                        <asp:TableCell></asp:TableCell>
                        <asp:TableCell></asp:TableCell>
                        <asp:TableCell></asp:TableCell>
                        <asp:TableCell></asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell ColumnSpan="4">
                            <asp:Label ID="Label28" runat="server" Font-Size="XX-Small"><%# DataBinder.Eval(Container.DataItem,"sCneePostCode") %></asp:Label>
                        </asp:TableCell>
                        <asp:TableCell></asp:TableCell>
                        <asp:TableCell></asp:TableCell>
                        <asp:TableCell></asp:TableCell>
                        <asp:TableCell></asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell ColumnSpan="4">
                            <asp:Label ID="Label20" runat="server" Font-Size="XX-Small"><%# DataBinder.Eval(Container.DataItem,"sCneeCountry") %></asp:Label>
                        </asp:TableCell>
                        <asp:TableCell></asp:TableCell>
                        <asp:TableCell></asp:TableCell>
                        <asp:TableCell></asp:TableCell>
                        <asp:TableCell></asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell ColumnSpan="4">
                            <asp:Label ID="Label21" runat="server" Font-Size="XX-Small"><%# DataBinder.Eval(Container.DataItem,"sCneeCtcName") %></asp:Label>
                            <br />
                            <br />
                        </asp:TableCell>
                        <asp:TableCell></asp:TableCell>
                        <asp:TableCell></asp:TableCell>
                        <asp:TableCell></asp:TableCell>
                        <asp:TableCell></asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow ID="trCustomerRef1" runat="server">
                        <asp:TableCell ColumnSpan="4">
                            <asp:Label ID="lblCustomerRef1" runat="server" Font-Size="XX-Small" Text='<%# Eval("sCustomerRef1")%>'></asp:Label>
                            <br />                            
                        </asp:TableCell>
                        <asp:TableCell></asp:TableCell>
                        <asp:TableCell></asp:TableCell>
                        <asp:TableCell></asp:TableCell>
                        <asp:TableCell></asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow ID="trCustomerRef2" runat="server">
                        <asp:TableCell ColumnSpan="4">
                            <asp:Label ID="lblCustomerRef2" runat="server" Font-Size="XX-Small" Text='<%# Eval("sCustomerRef2")%>'></asp:Label>
                            <br />                            
                        </asp:TableCell>
                        <asp:TableCell></asp:TableCell>
                        <asp:TableCell></asp:TableCell>
                        <asp:TableCell></asp:TableCell>
                        <asp:TableCell></asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow ID="trMisc1" runat="server">
                        <asp:TableCell ColumnSpan="4">
                            <asp:Label ID="lblMisc1" runat="server" Font-Size="XX-Small" Text='<%# Eval("sMisc1")%>'></asp:Label>
                            <br />                            
                        </asp:TableCell>
                        <asp:TableCell></asp:TableCell>
                        <asp:TableCell></asp:TableCell>
                        <asp:TableCell></asp:TableCell>
                        <asp:TableCell></asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow ID="trMisc2" runat="server">
                        <asp:TableCell ColumnSpan="4">
                            <asp:Label ID="lblMisc2" runat="server" Font-Size="XX-Small" Text='<%# Eval("sMisc2")%>'></asp:Label>
                            <br />                            
                        </asp:TableCell>
                        <asp:TableCell></asp:TableCell>
                        <asp:TableCell></asp:TableCell>
                        <asp:TableCell></asp:TableCell>
                        <asp:TableCell></asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell ColumnSpan="8">
                                <br />
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
                <asp:DataGrid ID="dgProducts" OnItemDataBound="dgProducts_ItemDataBound" runat="server"
                    DataSource='<%# DataBinder.Eval(Container.DataItem,"dblStockItemList") %>' AutoGenerateColumns="False"
                    Font-Names="Verdana" Font-Size="XX-Small" Width="650px" GridLines="None">
                    <Columns>
                        <asp:BoundColumn DataField="sProdCode" HeaderText="Product Code">
                            <HeaderStyle Font-Bold="True" Width="90px"></HeaderStyle>
                            <ItemStyle VerticalAlign="Top"></ItemStyle>
                        </asp:BoundColumn>
                        <asp:BoundColumn DataField="sProdDate" HeaderText="Product Date">
                            <HeaderStyle Font-Bold="True" Width="100px"></HeaderStyle>
                            <ItemStyle VerticalAlign="Top"></ItemStyle>
                        </asp:BoundColumn>
                        <asp:BoundColumn DataField="sProdDescription" HeaderText="Product Description">
                            <HeaderStyle Font-Bold="True" Width="250px"></HeaderStyle>
                            <ItemStyle VerticalAlign="Top"></ItemStyle>
                        </asp:BoundColumn>
                        <asp:BoundColumn DataField="nQuantity" HeaderText="Quantity">
                            <HeaderStyle Font-Bold="True" HorizontalAlign="Right" Width="70px"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Right" VerticalAlign="Top"></ItemStyle>
                        </asp:BoundColumn>
                        <asp:BoundColumn DataField="dblUnitValue" HeaderText="Unit Cost" DataFormatString="{0:#,##0.00}">
                            <HeaderStyle Font-Bold="True" HorizontalAlign="Right" Width="70px"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Right" VerticalAlign="Top"></ItemStyle>
                        </asp:BoundColumn>
                        <asp:TemplateColumn HeaderText="Value (€)">
                            <HeaderStyle Font-Bold="True" HorizontalAlign="Right" Width="70px"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Right" VerticalAlign="Top"></ItemStyle>
                            <ItemTemplate>
                                <asp:Label ID="Label22" runat="server"><%# Format((DataBinder.Eval(Container.DataItem, "nQuantity")) * (DataBinder.Eval(Container.DataItem, "dblUnitValue")),"#,##0.00") %></asp:Label>
                            </ItemTemplate>
                        </asp:TemplateColumn>
                    </Columns>
                </asp:DataGrid>
                <asp:Table ID="tbl001" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="650px">
                    <asp:TableRow>
                        <asp:TableCell Width="50px"></asp:TableCell>
                        <asp:TableCell Width="100px"></asp:TableCell>
                        <asp:TableCell Width="100px"></asp:TableCell>
                        <asp:TableCell Width="100px"></asp:TableCell>
                        <asp:TableCell Width="100px"></asp:TableCell>
                        <asp:TableCell Width="100px"></asp:TableCell>
                        <asp:TableCell Width="100px"></asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell ColumnSpan="6"></asp:TableCell>
                        <asp:TableCell ColumnSpan="1">
                                <hr />
                        </asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell HorizontalAlign="Right" ColumnSpan="6">
                            <asp:Label ID="Label23" runat="server" Font-Size="X-Small" Visible="<%# IsNotHyster() %>">Total value this order (£)</asp:Label><asp:Label
                                ID="Label24" runat="server" Font-Size="X-Small" Visible="<%# IsHyster() %>">Total value this order (€)</asp:Label>
                        </asp:TableCell>
                        <asp:TableCell HorizontalAlign="Right" ColumnSpan="1">
                            <asp:Label ID="Label25" runat="server" Font-Size="X-Small" Font-Bold="True"><%# Format(DataBinder.Eval(Container.DataItem,"dblStockItemListValue"),"#,##0.00") %></asp:Label>
                        </asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell ColumnSpan="3"></asp:TableCell>
                        <asp:TableCell HorizontalAlign="Right" ColumnSpan="3">
                            <asp:Label ID="Label26" runat="server" Font-Size="X-Small">Shipping costs this order (£)</asp:Label>
                        </asp:TableCell>
                        <asp:TableCell HorizontalAlign="Right" ColumnSpan="1">
                            <asp:Label ID="Label27" runat="server" Font-Size="X-Small" Font-Bold="True"><%#Format(DataBinder.Eval(Container.DataItem, "dblShippingCost"), "#,##0.00")%></asp:Label>
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
