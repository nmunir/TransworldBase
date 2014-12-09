<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ import Namespace="System.Collections.Generic" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="Telerik.Web.UI" %>
<%@ Register Assembly="Telerik.Web.UI" Namespace="Telerik.Web.UI" TagPrefix="telerik" %>
<%@ Import Namespace="ClosedXML.Excel" %>
<%@ Import Namespace="DocumentFormat.OpenXml" %>
<%@ Import Namespace="System.Data.Common" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<script runat="server">
    
    Const CUSTOMER_QUANTUM As Integer = 774

    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    Dim gnFirstConsignmentNumber As Int32
    Dim glstConsignments As List(Of Int32)
    
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsNumeric(Session("UserKey")) Then
            Server.Transfer("session_expired.aspx")
        End If

        If Not IsPostBack Then
            Call ReadCookies()
            CreateAmazonFolder()
            Call SetTitle()
            txtReference.Focus()
            tblPaging.Visible = False
        End If
    End Sub
    
    Protected Sub SetTitle()
        Dim sTitle As String = Session("SiteTitle")
        If sTitle <> String.Empty Then
            sTitle += " - "
        End If
        Page.Header.Title = sTitle & "Create Amazon Order"
    End Sub
    
    Private Sub CreateAmazonFolder()
        Dim sPath As String = Server.MapPath("~/")
        If Not Directory.Exists(sPath & "\Amazon") Then
            Directory.CreateDirectory(sPath & "\Amazon")
        End If
    End Sub
    
    Private Sub ReadCookies()
        Dim addressCookie As HttpCookie = Request.Cookies("AddressInfo")
        If addressCookie IsNot Nothing Then
            txtCustomerName.Text = addressCookie("CustomerName")
            txtAddr1.Text = addressCookie("Addr1")
            txtAddr2.Text = addressCookie("Addr2")
            txtAddr3.Text = addressCookie("Addr3")
            txtTown.Text = addressCookie("Town")
            txtPostCode.Text = addressCookie("PostCode")
            txtCtcName.Text = addressCookie("ContactName")
        Else
            txtCustomerName.Text = "Amazon EU SARI"
            txtAddr1.Text = "Plot 8, Marston Gate"
            txtAddr2.Text = "The Old Brickworks Site"
            txtAddr3.Text = "Ridgmont"
            txtTown.Text = "Bedford"
            txtPostCode.Text = "MK43 0ZA"
            txtCtcName.Text = "Inventory  display"
            txtReference.Text = ""
        End If
    End Sub
    
    Private Sub SaveCookie()
        Dim addressCookie As New HttpCookie("AddressInfo")
        addressCookie("CustomerName") = txtCustomerName.Text.Trim
        addressCookie("Addr1") = txtAddr1.Text.Trim
        addressCookie("Addr2") = txtAddr2.Text.Trim
        addressCookie("Addr3") = txtAddr3.Text.Trim
        addressCookie("Town") = txtTown.Text.Trim
        addressCookie("PostCode") = txtPostCode.Text.Trim
        addressCookie("ContactName") = txtCtcName.Text.Trim
        addressCookie.Expires = DateTime.Now.AddDays(30)
        Response.Cookies.Add(addressCookie)
    End Sub
    
    Protected Sub btnUpload_Click(ByVal sender As Object, ByVal e As EventArgs)
        rpProducts.Visible = False
        rpConsignments.Visible = False
        Try
            If ruAmazonFileUpload.UploadedFiles.Count > 0 Then
                lblNoResults.Visible = False
                divFileInfo.Visible = True
                psFileName = ruAmazonFileUpload.UploadedFiles(0).GetName
                lblFileName.Text = ruAmazonFileUpload.UploadedFiles(0).GetName
                'lblFileSize.Text = ruAmazonFileUpload.UploadedFiles(0).ContentLength.ToString()
                btnProcess.Enabled = True
                btnCheckData.Enabled = True
                tbLog.Text = String.Empty
                Call WriteToLog("Uploaded " & ruAmazonFileUpload.UploadedFiles(0).GetName & " @ " & Format(Date.Now, "d-MMM-yyyy hh:mm:ss"))
            Else
                lblNoResults.Visible = True
                divFileInfo.Visible = False
                Call WriteToLog("Nothing uploaded")
            End If
        Catch ex As Exception
            lblNoResults.Text = ex.Message.ToString()
            Call WriteToLog(ex.Message.ToString())
        End Try
    End Sub
    
    Protected Sub WriteToLog(sMessage As String)
        tbLog.Text += sMessage & Environment.NewLine
    End Sub
    
    Protected Sub btnProcess_Click(ByVal sender As Object, ByVal e As EventArgs)
        Call SaveCookie()
        Call ProcessExcelFile()
    End Sub
    
    Protected Function BuildProductsTable(ByVal dtSpreadsheet As DataTable) As DataTable
        Const COL_CATALOGUE_NUMBER As Int32 = 3
        Const COL_QUANTITYORDERED As Int32 = 9
        Dim nLineNo As Int32 = 1
        pdtProductsRequested = New DataTable("ProductsRequested")
        Dim dc As New DataColumn("ProductCode")
        dc.DataType = Type.GetType("System.String")
        dc.AllowDBNull = False
        pdtProductsRequested.Columns.Add(dc)
        pdtProductsRequested.PrimaryKey = New DataColumn() {dc}
        pdtProductsRequested.Columns.Add(New DataColumn("AvailableQty", Type.GetType("System.Int32")))
        pdtProductsRequested.Columns.Add(New DataColumn("RequestedQty", Type.GetType("System.Int32")))
        pdtProductsRequested.Columns.Add(New DataColumn("LogisticProductKey", Type.GetType("System.Int32")))
        
        If dtSpreadsheet IsNot Nothing AndAlso dtSpreadsheet.Rows.Count <> 0 Then
            For Each drSpreadsheet As DataRow In dtSpreadsheet.Rows
                Dim sQtyOrdered As String = drSpreadsheet(COL_QUANTITYORDERED).ToString()
                If IsNumeric(sQtyOrdered) Then
                    Dim drDuplicateRows() As DataRow
                    Dim sProductCode As String = drSpreadsheet(COL_CATALOGUE_NUMBER)
                    Dim sReqQuantity As String = drSpreadsheet(COL_QUANTITYORDERED)
                    If IsNumeric(sReqQuantity) Then
                        'Dim nReqQuantity As Integer = Convert.ToInt32(drSpreadsheet(COL_QUANTITYORDERED))
                        Dim nLogisticProductKey As Integer
                        Dim expression As String = "ProductCode = '" & sProductCode & "'"
                        'pdtProductsInExcelSheet.Select(expression)
                        drDuplicateRows = pdtProductsRequested.Select(expression)
                        If drDuplicateRows.Length = 0 Then
                            Call WriteToLog("Processing line " & nLineNo.ToString)
                            Dim row As DataRow
                            pdtProductsRequested.ImportRow(drSpreadsheet)
                            nLogisticProductKey = GetProductKey(sProductCode)
                            If nLogisticProductKey > 0 Then
                                row = pdtProductsRequested.Rows.Find(sProductCode)
                                row("LogisticProductKey") = nLogisticProductKey
                                Dim nAvailableQty As Int32 = GetAvailableQty(nLogisticProductKey)
                                row("AvailableQty") = nAvailableQty
                                If nAvailableQty > 0 Then
                                    If row("RequestedQty") > nAvailableQty Then
                                        Call WriteToLog("Only " & nAvailableQty & " item(s)  of this product (" & sProductCode & ") will be ordered as not all requested quantity is available")
                                    Else
                                        Call WriteToLog("OK (" & sProductCode & ")")
                                    End If
                                Else
                                    Call WriteToLog("This product (" & sProductCode & ") will not be ordered as no quantity is available")
                                End If
                            Else
                                row = pdtProductsRequested.Rows.Find(sProductCode)
                                row("AvailableQty") = 0
                                row("LogisticProductKey") = -1
                                Call WriteToLog("Could not match product!")
                            End If
                        Else
                            Call WriteToLog("On line " & nLineNo.ToString & " processing second or subsequent reference to product!!")
                            Dim row As DataRow = pdtProductsRequested.Rows.Find(sProductCode)
                            If row IsNot Nothing Then
                                row("RequestedQty") = drDuplicateRows(0)("RequestedQty") + Convert.ToInt32(drSpreadsheet("RequestedQty"))
                                row("AvailableQty") = GetAvailableQty(nLogisticProductKey)
                            Else
                                Call WriteToLog("On line " & nLineNo.ToString & " subsequent reference was NULL")
                            End If
                        End If
                    Else
                        Call WriteToLog("On line " & nLineNo.ToString & " non-numeric quantity ordered value")
                    End If
                    
                End If
                        
                '    Else
                '        Call WriteToLog("On line " & nLineNo.ToString & " bar code non-numeric and/or wrong length")
                '    End If
                'Else
                '    'Console.WriteLine("Code didn't match")
                '    Call WriteToLog("On line " & nLineNo.ToString & " no bar code found")
                'End If
                
                nLineNo += 1
            Next
        End If
        Return pdtProductsRequested
    End Function
    
    'Private Sub rpProductsInExcelSheet_ItemDataBound(ByVal sender As Object, ByVal e As RepeaterItemEventArgs) Handles rpProducts.ItemDataBound
    '    If e.Item.ItemType = ListItemType.Header Then
    '        Dim lblNoOfProducts As Label = e.Item.FindControl("lblNoOfProducts")
    '        lblNoOfProducts.Text = "Number of products in Excel sheet: <b>" & pdtProductsToOrder.Rows.Count.ToString() & "</b>"
    '    End If
    'End Sub
    
    Public Sub ProcessExcelFile()
        Dim dt As DataTable = ExcelSheetToDataTable()
        pdtProductsRequested = BuildProductsTable(dt)
        Call BindProducts()
        Call SubmitOrdersInChunks()
    End Sub
    
    Public Sub BindProducts()
        Dim objPds As New PagedDataSource()
        objPds.DataSource = pdtProductsRequested.DefaultView
        objPds.AllowPaging = True
        objPds.PageSize = 50
        objPds.CurrentPageIndex = pnCurrentPage
        
        lblCurrentPage.Text = "Page: " + (pnCurrentPage + 1).ToString() + " of " + objPds.PageCount.ToString()
        
        cmdPrev.Enabled = Not objPds.IsFirstPage
        cmdNext.Enabled = Not objPds.IsLastPage
        
        tblPaging.Visible = True
        
        rpProducts.DataSource = objPds
        rpProducts.DataBind()
    End Sub
    
    Private Function GetProductKey(ByVal sProductCode As String) As Integer
        Dim sSQL As String = "SELECT LogisticProductKey from LogisticProduct where ProductCode = '" & sProductCode & "' AND CustomerKey = " & CUSTOMER_QUANTUM
        Dim oDataTable As New DataTable
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter(sSQL, oConn)
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        Try
            oAdapter.Fill(oDataTable)
            oConn.Open()
        Catch ex As Exception
            WebMsgBox.Show("Error in ExecuteQueryToDataTable executing: " & sSQL & " : " & ex.Message)
        Finally
            oConn.Close()
        End Try
        If oDataTable IsNot Nothing AndAlso oDataTable.Rows.Count <> 0 Then
            GetProductKey = Convert.ToInt32(oDataTable.Rows(0)(0))
        Else
            GetProductKey = -1
        End If
    End Function
    
    Private Function GetAvailableQty(ByVal nLogisticProductKey As Integer) As Integer
        Dim sSQL As String = "SELECT Quantity = CASE ISNUMERIC((SELECT SUM(LogisticProductQuantity) FROM LogisticProductLocation WHERE LogisticProductKey = " & nLogisticProductKey & ")) WHEN 0 THEN 0 ELSE (SELECT SUM(LogisticProductQuantity) FROM LogisticProductLocation WHERE LogisticProductKey = " & nLogisticProductKey & ") END"
        Dim oDataTable As New DataTable
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter(sSQL, oConn)
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        Try
            oAdapter.Fill(oDataTable)
            oConn.Open()
        Catch ex As Exception
            WebMsgBox.Show("Error in ExecuteQueryToDataTable executing: " & sSQL & " : " & ex.Message)
        Finally
            oConn.Close()
        End Try
        If oDataTable IsNot Nothing AndAlso oDataTable.Rows.Count <> 0 Then
            GetAvailableQty = oDataTable.Rows(0)(0)
        Else
            GetAvailableQty = 0
        End If
    End Function
    
    Private Sub SubmitOrdersInChunks()
        glstConsignments = New List(Of Int32)
        pdtConsignments = New DataTable()
        Dim dtProductsToOrder As New DataTable
        
        Dim dc As New DataColumn("ProductCode")
        dc.DataType = Type.GetType("System.String")
        dc.AllowDBNull = False
        dtProductsToOrder.Columns.Add(dc)
        dtProductsToOrder.PrimaryKey = New DataColumn() {dc}
        dtProductsToOrder.Columns.Add(New DataColumn("AvailableQty", Type.GetType("System.Int32")))
        dtProductsToOrder.Columns.Add(New DataColumn("RequestedQty", Type.GetType("System.Int32")))
        dtProductsToOrder.Columns.Add(New DataColumn("LogisticProductKey", Type.GetType("System.Int32")))

        For Each drProductRequested As DataRow In pdtProductsRequested.Select("(LogisticProductKey <> -1) AND (AvailableQty > 0)")
            Dim drProductToOrder As DataRow = dtProductsToOrder.NewRow()
            drProductToOrder("ProductCode") = drProductRequested("ProductCode")
            drProductToOrder("AvailableQty") = drProductRequested("AvailableQty")
            drProductToOrder("RequestedQty") = drProductRequested("RequestedQty")
            drProductToOrder("LogisticProductKey") = drProductRequested("LogisticProductKey")
            Call WriteToLog("Preparing to order " & drProductToOrder("RequestedQty") & " of " & drProductToOrder("ProductCode") & " (" & drProductToOrder("LogisticProductKey") & ")")
            dtProductsToOrder.Rows.Add(drProductToOrder)
        Next
        Dim nTotProductsToProcess As Int32 = dtProductsToOrder.Rows.Count
        Dim nOrderSize As Int32 = Convert.ToInt32(tbOrderSize.Text.Trim)
        Call WriteToLog("Total products: " & nTotProductsToProcess)
        Call WriteToLog("Chunk size: " & nOrderSize)
        Dim nStart As Int32 = 0
        Dim nEnd As Int32 = nOrderSize - 1
        If nTotProductsToProcess < nOrderSize Then
            nEnd = nTotProductsToProcess - 1
        End If

        Do
            Call nSubmitConsignment(dtProductsToOrder, nStart, nEnd)
            nStart += nOrderSize
            If nStart > (nTotProductsToProcess - 1) Then
                Exit Do
            End If
            
            nEnd = nStart + (nOrderSize - 1)
            If nEnd > (nTotProductsToProcess - 1) Then
                nEnd = nTotProductsToProcess - 1
            End If
        Loop
        
        rpProducts.Visible = False
        tblPaging.Visible = False
        rpConsignments.Visible = True
        For Each nItem As Int32 In glstConsignments
            Call WriteToLog("Created AWB " & nItem.ToString)
        Next
        rpConsignments.DataSource = glstConsignments
        rpConsignments.DataBind()
        Call SendMail("AMAZON_ORDER", "chris.newport@transworld.eu.com", "AmazonOrder", tbLog.Text, tbLog.Text.Replace(Environment.NewLine, "<br />" & Environment.NewLine))
    End Sub
    
    Protected Function nSubmitConsignment(ByRef dt As DataTable, ByVal nStart As Int32, ByVal nEnd As Int32) As Integer
        Dim sConn As String = ConfigLib.GetConfigItem_ConnectionString
        Dim lBookingKey As Long
        Dim lConsignmentKey As Long
        Dim BookingFailed As Boolean
        Dim oConn As New SqlConnection(sConn)
        Dim oTrans As SqlTransaction
        Dim oCmdAddBooking As SqlCommand = New SqlCommand("spASPNET_StockBooking_Add3", oConn)
        nSubmitConsignment = 0
        oCmdAddBooking.CommandType = CommandType.StoredProcedure
        Dim param1 As SqlParameter = New SqlParameter("@UserProfileKey", SqlDbType.Int, 4)
        param1.Value = CInt(Session("UserKey"))
        oCmdAddBooking.Parameters.Add(param1)
        Dim param2 As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
        param2.Value = CUSTOMER_QUANTUM
        oCmdAddBooking.Parameters.Add(param2)
        Dim param2a As SqlParameter = New SqlParameter("@BookingOrigin", SqlDbType.NVarChar, 20)
        param2a.Value = "WEB_BOOKING"
        oCmdAddBooking.Parameters.Add(param2a)
        Dim param3 As SqlParameter = New SqlParameter("@BookingReference1", SqlDbType.NVarChar, 25)
        param3.Value = ""
        oCmdAddBooking.Parameters.Add(param3)
        Dim param4 As SqlParameter = New SqlParameter("@BookingReference2", SqlDbType.NVarChar, 25)
        param4.Value = ""
        oCmdAddBooking.Parameters.Add(param4)
        Dim param5 As SqlParameter = New SqlParameter("@BookingReference3", SqlDbType.NVarChar, 50)
        param5.Value = ""
        oCmdAddBooking.Parameters.Add(param5)
        Dim param6 As SqlParameter = New SqlParameter("@BookingReference4", SqlDbType.NVarChar, 50)
        param6.Value = ""
        oCmdAddBooking.Parameters.Add(param6)
        Dim param6a As SqlParameter = New SqlParameter("@ExternalReference", SqlDbType.NVarChar, 50)
        param6a.Value = Nothing
        oCmdAddBooking.Parameters.Add(param6a)
        Dim param7 As SqlParameter = New SqlParameter("@SpecialInstructions", SqlDbType.NVarChar, 1000)
        param7.Value = ""
        If gnFirstConsignmentNumber > 0 Then
            param7.Value = "PUT WITH " & gnFirstConsignmentNumber.ToString
        End If
        oCmdAddBooking.Parameters.Add(param7)
        Dim param8 As SqlParameter = New SqlParameter("@PackingNoteInfo", SqlDbType.NVarChar, 1000)
        param8.Value = txtReference.Text.Trim
        oCmdAddBooking.Parameters.Add(param8)
        Dim param9 As SqlParameter = New SqlParameter("@ConsignmentType", SqlDbType.NVarChar, 20)
        param9.Value = "STOCK ITEM"
        oCmdAddBooking.Parameters.Add(param9)
        Dim param10 As SqlParameter = New SqlParameter("@ServiceLevelKey", SqlDbType.Int, 4)
        param10.Value = -1
        oCmdAddBooking.Parameters.Add(param10)
        Dim param11 As SqlParameter = New SqlParameter("@Description", SqlDbType.NVarChar, 250)
        param11.Value = "PRINTED MATTER - FREE DOMICILE"
        oCmdAddBooking.Parameters.Add(param11)
        
        Dim sSQL As String = "SELECT * FROM Customer WHERE CustomerKey = " & CUSTOMER_QUANTUM
        Dim dtCnor As New DataTable
        oConn = New SqlConnection(sConn)
        Dim oAdapter As New SqlDataAdapter(sSQL, oConn)
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        Try
            oAdapter.Fill(dtCnor)
            oConn.Open()
        Catch ex As Exception
            WebMsgBox.Show("Error in ExecuteQueryToDataTable executing: " & sSQL & " : " & ex.Message)
        Finally
            oConn.Close()
        End Try
        
        Dim drCnor As DataRow
        If dtCnor.Rows.Count = 1 Then
            drCnor = dtCnor.Rows(0)
        Else
            WebMsgBox.Show("Couldn't find Consignor details.")
            Exit Function
        End If
       
        Dim param13 As SqlParameter = New SqlParameter("@CnorName", SqlDbType.NVarChar, 50)
        'param13.Value = psCnorCompany
        param13.Value = drCnor("CustomerName")
       
        oCmdAddBooking.Parameters.Add(param13)
        Dim param14 As SqlParameter = New SqlParameter("@CnorAddr1", SqlDbType.NVarChar, 50)
        param14.Value = drCnor("CustomerAddr1")
        oCmdAddBooking.Parameters.Add(param14)
        Dim param15 As SqlParameter = New SqlParameter("@CnorAddr2", SqlDbType.NVarChar, 50)
        param15.Value = drCnor("CustomerAddr2")
        oCmdAddBooking.Parameters.Add(param15)
        Dim param16 As SqlParameter = New SqlParameter("@CnorAddr3", SqlDbType.NVarChar, 50)
        param16.Value = drCnor("CustomerAddr3")
        oCmdAddBooking.Parameters.Add(param16)
        Dim param17 As SqlParameter = New SqlParameter("@CnorTown", SqlDbType.NVarChar, 50)
        param17.Value = drCnor("CustomerTown")
        oCmdAddBooking.Parameters.Add(param17)
        Dim param18 As SqlParameter = New SqlParameter("@CnorState", SqlDbType.NVarChar, 50)
        param18.Value = drCnor("CustomerCounty")
        oCmdAddBooking.Parameters.Add(param18)
        Dim param19 As SqlParameter = New SqlParameter("@CnorPostCode", SqlDbType.NVarChar, 50)
        param19.Value = drCnor("CustomerPostCode")
        oCmdAddBooking.Parameters.Add(param19)
        Dim param20 As SqlParameter = New SqlParameter("@CnorCountryKey", SqlDbType.Int, 4)
        param20.Value = drCnor("CustomerCountryKey")
        oCmdAddBooking.Parameters.Add(param20)
        Dim param21 As SqlParameter = New SqlParameter("@CnorCtcName", SqlDbType.NVarChar, 50)
        param21.Value = ""
        oCmdAddBooking.Parameters.Add(param21)
        Dim param22 As SqlParameter = New SqlParameter("@CnorTel", SqlDbType.NVarChar, 50)
        param22.Value = ""
        oCmdAddBooking.Parameters.Add(param22)
        Dim param23 As SqlParameter = New SqlParameter("@CnorEmail", SqlDbType.NVarChar, 50)
        param23.Value = ""
        oCmdAddBooking.Parameters.Add(param23)
        Dim param24 As SqlParameter = New SqlParameter("@CnorPreAlertFlag", SqlDbType.Bit)
        param24.Value = 0
        oCmdAddBooking.Parameters.Add(param24)
        Dim param25 As SqlParameter = New SqlParameter("@CneeName", SqlDbType.NVarChar, 50)
        param25.Value = txtCustomerName.Text.Trim
        oCmdAddBooking.Parameters.Add(param25)
        Dim param26 As SqlParameter = New SqlParameter("@CneeAddr1", SqlDbType.NVarChar, 50)
        param26.Value = txtAddr1.Text.Trim
        oCmdAddBooking.Parameters.Add(param26)
        Dim param27 As SqlParameter = New SqlParameter("@CneeAddr2", SqlDbType.NVarChar, 50)
        param27.Value = txtAddr2.Text.Trim
        oCmdAddBooking.Parameters.Add(param27)
        Dim param28 As SqlParameter = New SqlParameter("@CneeAddr3", SqlDbType.NVarChar, 50)
        param28.Value = txtAddr3.Text.Trim
        oCmdAddBooking.Parameters.Add(param28)
        Dim param29 As SqlParameter = New SqlParameter("@CneeTown", SqlDbType.NVarChar, 50)
        param29.Value = txtTown.Text.Trim
        oCmdAddBooking.Parameters.Add(param29)
        Dim param30 As SqlParameter = New SqlParameter("@CneeState", SqlDbType.NVarChar, 50)
        param30.Value = ""
        oCmdAddBooking.Parameters.Add(param30)
        Dim param31 As SqlParameter = New SqlParameter("@CneePostCode", SqlDbType.NVarChar, 50)
        param31.Value = txtPostCode.Text.Trim
        oCmdAddBooking.Parameters.Add(param31)
        Dim param32 As SqlParameter = New SqlParameter("@CneeCountryKey", SqlDbType.Int, 4)
        param32.Value = 222
        oCmdAddBooking.Parameters.Add(param32)
        Dim param33 As SqlParameter = New SqlParameter("@CneeCtcName", SqlDbType.NVarChar, 50)
        param33.Value = txtCtcName.Text.Trim
        oCmdAddBooking.Parameters.Add(param33)
        Dim param34 As SqlParameter = New SqlParameter("@CneeTel", SqlDbType.NVarChar, 50)
        param34.Value = ""
        oCmdAddBooking.Parameters.Add(param34)
        Dim param35 As SqlParameter = New SqlParameter("@CneeEmail", SqlDbType.NVarChar, 50)
        param35.Value = ""
        oCmdAddBooking.Parameters.Add(param35)
        Dim param36 As SqlParameter = New SqlParameter("@CneePreAlertFlag", SqlDbType.Bit)
        param36.Value = 0
        oCmdAddBooking.Parameters.Add(param36)
        Dim param37 As SqlParameter = New SqlParameter("@LogisticBookingKey", SqlDbType.Int, 4)
        param37.Direction = ParameterDirection.Output
        oCmdAddBooking.Parameters.Add(param37)
        Dim param38 As SqlParameter = New SqlParameter("@ConsignmentKey", SqlDbType.Int, 4)
        param38.Direction = ParameterDirection.Output
        oCmdAddBooking.Parameters.Add(param38)
        
        For i As Int32 = 0 To oCmdAddBooking.Parameters.Count - 1
            Trace.Write(oCmdAddBooking.Parameters(i).ParameterName.ToString)
            Trace.Write(oCmdAddBooking.Parameters(i).DbType.ToString)
            If Not IsNothing(oCmdAddBooking.Parameters(i).Value) Then
                Trace.Write(oCmdAddBooking.Parameters(i).Value.ToString)
            Else
                Trace.Write("NOTHING")
            End If
        Next
        
        Try
            BookingFailed = False
            oConn.Open()
            oTrans = oConn.BeginTransaction(IsolationLevel.ReadCommitted, "AddBooking")
            oCmdAddBooking.Connection = oConn
            oCmdAddBooking.Transaction = oTrans
            oCmdAddBooking.ExecuteNonQuery()
            lBookingKey = CLng(oCmdAddBooking.Parameters("@LogisticBookingKey").Value.ToString)
            lConsignmentKey = CLng(oCmdAddBooking.Parameters("@ConsignmentKey").Value.ToString)
            If lBookingKey > 0 Then
                For i As Int32 = nStart To nEnd
                    Dim dr As DataRow = dt.Rows(i)
                    Dim oCmdAddStockItem As SqlCommand = New SqlCommand("spASPNET_LogisticMovement_Add", oConn)
                    oCmdAddStockItem.CommandType = CommandType.StoredProcedure
                    Dim param51 As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int, 4)
                    param51.Value = CInt(Session("UserKey"))
                    oCmdAddStockItem.Parameters.Add(param51)
                    Dim param52 As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
                    param52.Value = CUSTOMER_QUANTUM
                    oCmdAddStockItem.Parameters.Add(param52)
                    Dim param53 As SqlParameter = New SqlParameter("@LogisticBookingKey", SqlDbType.Int, 4)
                    param53.Value = lBookingKey
                    oCmdAddStockItem.Parameters.Add(param53)
                    Dim param54 As SqlParameter = New SqlParameter("@LogisticProductKey", SqlDbType.Int, 4)
                    param54.Value = Convert.ToInt32(dr("LogisticProductKey"))
                    oCmdAddStockItem.Parameters.Add(param54)
                    Dim param55 As SqlParameter = New SqlParameter("@LogisticMovementStateId", SqlDbType.NVarChar, 20)
                    param55.Value = "PENDING"
                    oCmdAddStockItem.Parameters.Add(param55)
                    Dim param56 As SqlParameter = New SqlParameter("@ItemsOut", SqlDbType.Int, 4)
                    
                    Dim nRequestedQty As Integer = Convert.ToInt32(dr("RequestedQty"))
                    Dim nAvailableQty As Integer = Convert.ToInt32(dr("AvailableQty"))
                    
                    If nRequestedQty > nAvailableQty Then
                        param56.Value = nAvailableQty
                        Call WriteToLog(dr("ProductCode") & " (" & dr("LogisticProductKey") & "): requested order of " & nRequestedQty.ToString & " reduced to " & nAvailableQty.ToString)
                    Else
                        param56.Value = nRequestedQty
                        Call WriteToLog(dr("ProductCode") & " (" & dr("LogisticProductKey") & "): " & nRequestedQty.ToString & " ordered")
                    End If
                    oCmdAddStockItem.Parameters.Add(param56)
                    Dim param57 As SqlParameter = New SqlParameter("@ConsignmentKey", SqlDbType.Int, 8)
                    param57.Value = lConsignmentKey
                    oCmdAddStockItem.Parameters.Add(param57)
                    oCmdAddStockItem.Connection = oConn
                    oCmdAddStockItem.Transaction = oTrans
                    oCmdAddStockItem.ExecuteNonQuery()
                Next
                Dim oCmdCompleteBooking As SqlCommand = New SqlCommand("spASPNET_LogisticBooking_Complete", oConn)
                oCmdCompleteBooking.CommandType = CommandType.StoredProcedure
                Dim param71 As SqlParameter = New SqlParameter("@LogisticBookingKey", SqlDbType.Int, 4)
                param71.Value = lBookingKey
                oCmdCompleteBooking.Parameters.Add(param71)
                oCmdCompleteBooking.Connection = oConn
                oCmdCompleteBooking.Transaction = oTrans
                oCmdCompleteBooking.ExecuteNonQuery()
            Else
                BookingFailed = True
            End If
            If Not BookingFailed Then
                oTrans.Commit()
                nSubmitConsignment = lConsignmentKey
                glstConsignments.Add(lConsignmentKey)
                If gnFirstConsignmentNumber = 0 Then
                    gnFirstConsignmentNumber = lConsignmentKey
                End If
            Else
                oTrans.Rollback("AddBooking")
            End If
        Catch ex As SqlException
            oTrans.Rollback("AddBooking")
        Finally
            oConn.Close()
        End Try
    End Function
    
    Protected Sub SendMail(ByVal sType As String, ByVal sRecipient As String, ByVal sSubject As String, ByVal sBodyText As String, ByVal sBodyHTML As String)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Email_AddToQueue", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
    
        Try
            oCmd.Parameters.Add(New SqlParameter("@MessageId", SqlDbType.NVarChar, 20))
            oCmd.Parameters("@MessageId").Value = sType
    
            oCmd.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
            oCmd.Parameters("@CustomerKey").Value = Session("CustomerKey")
    
            oCmd.Parameters.Add(New SqlParameter("@StockBookingKey", SqlDbType.Int))
            oCmd.Parameters("@StockBookingKey").Value = 0
    
            oCmd.Parameters.Add(New SqlParameter("@ConsignmentKey", SqlDbType.Int))
            oCmd.Parameters("@ConsignmentKey").Value = 0
    
            oCmd.Parameters.Add(New SqlParameter("@ProductKey", SqlDbType.Int))
            oCmd.Parameters("@ProductKey").Value = 0
    
            oCmd.Parameters.Add(New SqlParameter("@To", SqlDbType.NVarChar, 100))
            oCmd.Parameters("@To").Value = sRecipient
    
            oCmd.Parameters.Add(New SqlParameter("@Subject", SqlDbType.NVarChar, 60))
            oCmd.Parameters("@Subject").Value = sSubject
    
            oCmd.Parameters.Add(New SqlParameter("@BodyText", SqlDbType.NText))
            oCmd.Parameters("@BodyText").Value = sBodyText
    
            oCmd.Parameters.Add(New SqlParameter("@BodyHTML", SqlDbType.NText))
            oCmd.Parameters("@BodyHTML").Value = sBodyHTML
    
            oCmd.Parameters.Add(New SqlParameter("@QueuedBy", SqlDbType.Int))
            oCmd.Parameters("@QueuedBy").Value = Session("UserKey")
    
            oConn.Open()
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            WebMsgBox.Show("Error in SendMail: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub

    Public Function ExcelSheetToDataTable() As System.Data.DataTable
        'Dim sFilePath As String = Server.MapPath("~/" & psFileName.Substring(psFileName.LastIndexOf("\")))
        Dim sFilePath As String = Server.MapPath("~/Amazon/" & psFileName)
        If Path.GetExtension(sFilePath) = ".xls" Then
            WebMsgBox.Show("Upload an Excel spreadsheet with file extension .xlsx (Excel 2007 format)")
        End If
        If File.Exists(sFilePath) Then
            Dim workbook = New XLWorkbook(sFilePath, XLEventTracking.Disabled)
            Dim sheetName As String = workbook.Worksheet(1).Name
            Dim dtSpreadsheet = New DataTable()
            Dim xlWorksheet = workbook.Worksheet(sheetName)
            Dim range = xlWorksheet.Range(xlWorksheet.FirstCellUsed(), xlWorksheet.LastCellUsed())   'IXLRange xlRangeRow = xlWorksheet.AsRange();                           ' IXLCell rowCell = xlWorksheet.LastCellUsed();
            Dim col As Integer = range.ColumnCount()
            'Dim row As Integer = range.RowCount()
            dtSpreadsheet.Clear()

            For i As Integer = 1 To col
                Dim column As IXLCell = xlWorksheet.Cell(1, i)
                If column.Value.ToString() = "SUBMITTED_QTY" Then
                    dtSpreadsheet.Columns.Add("RequestedQty")
                ElseIf column.Value.ToString() = "CATALOG_NUMBER" Then
                    dtSpreadsheet.Columns.Add("ProductCode")
                ElseIf column.Value.ToString() = "XXX" Then
                    dtSpreadsheet.Columns.Add(column.Value.ToString())
                Else
                    dtSpreadsheet.Columns.Add(column.Value.ToString())
                End If
            Next

            Dim firstHeadRow As Integer = 0          ' add rows data
            For Each item As IXLRangeRow In range.Rows()
                If firstHeadRow <> 0 Then
                    Dim array = New Object(col - 1) {}
                    For y As Integer = 1 To col
                        array(y - 1) = item.Cell(y).Value
                    Next
                    dtSpreadsheet.Rows.Add(array)
                End If
                firstHeadRow += 1
            Next
            Return dtSpreadsheet
        Else
            Response.Write("File doesn't exist")
            ExcelSheetToDataTable = Nothing
        End If
    End Function
    
    Private Sub cmdPrev_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdPrev.Click
        pnCurrentPage -= 1
        Call BindProducts()
    End Sub

    Private Sub cmdNext_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdNext.Click
        pnCurrentPage += 1
        Call BindProducts()
    End Sub

    Property pdtProductsRequested() As DataTable
        Get
            Dim o As Object = ViewState("AMAZON_PRODUCTS_REQUESTED")
            If o Is Nothing Then
                Return Nothing
            End If
            Return o
        End Get
        Set(ByVal Value As DataTable)
            ViewState("AMAZON_PRODUCTS_REQUESTED") = Value
        End Set
    End Property
    
    Property pdtConsignments() As DataTable
        Get
            Dim o As Object = ViewState("AMAZON_REPORT_CONSIGNMENTS")
            If o Is Nothing Then
                Return Nothing
            End If
            Return o
        End Get
        Set(ByVal Value As DataTable)
            ViewState("AMAZON_REPORT_CONSIGNMENTS") = Value
        End Set
    End Property
    
    Property psFileName() As String
        Get
            Dim o As Object = ViewState("AMAZON_REPORT_FileName")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("AMAZON_REPORT_FileName") = Value
        End Set
    End Property
    
    Property pnCurrentPage() As Int32
        Get
            Dim o As Object = ViewState("AMAZON_REPORT_CurrentPage")
            If o Is Nothing Then
                Return 0
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Int32)
            ViewState("AMAZON_REPORT_CurrentPage") = Value
        End Set
    End Property

    Protected Sub lnkbtnHelp_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        pnlHelp.Visible = True
        lnkbtnHelp.Visible = False
    End Sub
    
    Protected Sub btnCheckData_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        rpProducts.Visible = True
        rpConsignments.Visible = False
        Call SaveCookie()
        Dim dt As DataTable = ExcelSheetToDataTable()
        pdtProductsRequested = BuildProductsTable(dt)
        Call BindProducts()
    End Sub
    
    Protected Sub rpProducts_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.RepeaterItemEventArgs)
        Dim riea As RepeaterItemEventArgs = e
        Dim ri As RepeaterItem
        ri = riea.Item
        If ri.ItemType = ListItemType.Item Or ri.ItemType = ListItemType.AlternatingItem Then
            Dim lblProductCode As Label = ri.Controls(1)
            Dim lblRequestedQty As Label = ri.Controls(3)
            Dim lblAvailableQty As Label = ri.Controls(5)
            Dim nRequestedQty As Int32 = CInt(lblRequestedQty.Text)
            Dim nAvailableQty As Int32 = CInt(lblAvailableQty.Text)
            If nAvailableQty = 0 Then
                lblProductCode.ForeColor = System.Drawing.Color.Red
                lblProductCode.Font.Bold = True
                lblAvailableQty.ForeColor = System.Drawing.Color.Red
                lblAvailableQty.Font.Bold = True
            End If
            If nRequestedQty > nAvailableQty Then
                lblProductCode.ForeColor = System.Drawing.Color.Red
                lblProductCode.Font.Bold = True
                lblRequestedQty.ForeColor = System.Drawing.Color.Red
                lblRequestedQty.Font.Bold = True
            End If
        End If
        
    End Sub
    
</script>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Create Amazon Order</title>
    <style type="text/css">
        .label
        {
            margin-left: 1px;
            width: 15%;
            float: left;
            text-align: left;
        }
    </style>
    <script type="text/javascript" src="https://ajax.googleapis.com/ajax/libs/jquery/1.7.2/jquery.min.js"></script>
    <script type="text/javascript">
        $(document).ready(function () {
            $('label').addClass('label');
        });
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <main:Header ID="ctlHeader" runat="server" />
    <table style="width: 100%" cellpadding="0" cellspacing="0">
        <tr class="bar_addressbook">
            <td style="width: 50%; white-space: nowrap">
                &nbsp;
            </td>
            <td style="width: 50%; white-space: nowrap" align="right">
                &nbsp;
            </td>
        </tr>
    </table>
    <asp:ScriptManager ID="ScriptManager1" runat="server" />
    <asp:Panel ID="pnlHelp" runat="server" Visible="False" Font-Size="XX-Small" BackColor="#FFFFCC">
        This utility creates one or more consignments from an Amazon SKU spreadsheet.<br />
        <br />
        <b>NOTE: When downloaded from Amazon the spreadsheet is in Excel 2003 format. You must
            export the spreadsheet to Excel 2007 format (.xlsx) before uploading it to this
            utility.</b>
            <br />
            <br />
            INSTRUCTIONS
            <ul>
                <li>Ensure the spreadsheet containing the orders you want to place has an XLSX file 
                    extension. If the file has an .XLS extension, you must save it as an Excel 2007 
                    format file.</li>
                <li>Click the Select button, navigate to the filder containing your spreadsheet, and 
                    select the spreadsheet</li>
                <li>Click the Upload Spreadsheet button. The filename of the spreadsheet should 
                    appear.</li>
                <li>Click the Check Data button. You will see a list of the products in the 
                    spreadsheet showing the status of each on the Transworld system.</li>
                <li>The destination address for your order is saved from the last time you placed an 
                    order. Check it is correct.</li>
                <li>Enter the reference number for the order. This appears on the Packing Note.</li>
                <li>Click the Create Consignments button. Your order is created. If there are more 
                    than the maximum permitterd orders for a single consignment, two or more 
                    consignments will be created. The second and subsquent consignments will refer 
                    to the first consignment so they are all sent as a single shipment. You will see 
                    a list of the consignments created.<br />
                    <br />
                </li>
        </ul>
    </asp:Panel>
    <div>
        Max products per consignment:
        <asp:TextBox ID="tbOrderSize" Text="50" MaxLength="5" ValidationGroup="vg" Width="40px"
            runat="server" Font-Size="XX-Small" />
        <asp:RequiredFieldValidator ID="rfvOrderSize" runat="server" ControlToValidate="tbOrderSize"
            ValidationGroup="vg" ErrorMessage="required!" />
        <asp:RangeValidator ID="rvOrderSize" runat="server" MinimumValue="1" MaximumValue="50"
            Type="Integer" ErrorMessage="Max products per consignment should be between 1 and 50"
            Display="None" ControlToValidate="tbOrderSize" ValidationGroup="vg" />
        &nbsp;<asp:LinkButton ID="lnkbtnHelp" runat="server" Font-Size="XX-Small" OnClick="lnkbtnHelp_Click">help</asp:LinkButton>
        <br />
        <br />
    </div>
    <asp:Panel ID="pnlAddress" BackColor="#FDE8E9" runat="server">
        <div>
            <asp:ValidationSummary ID="vs" ValidationGroup="vg" runat="server" Font-Bold="True"
                Font-Size="Medium" />
        </div>
        <div>
            <label>
                Addressee :</label>
            <asp:TextBox ID="txtCustomerName" runat="server" MaxLength="100" Width="200px" ValidationGroup="vg"
                Font-Size="XX-Small" />
            <asp:RequiredFieldValidator ID="rfvCustomerName" runat="server" ControlToValidate="txtCustomerName"
                Display="None" ValidationGroup="vg" ErrorMessage="Addressee required!" />
        </div>
        <div>
            <label>
                Contact Name :</label>
            <asp:TextBox ID="txtCtcName" runat="server" MaxLength="100" Width="200px" ValidationGroup="vg"
                Font-Size="XX-Small" />
            <asp:RequiredFieldValidator ID="rfvCtcName" runat="server" ControlToValidate="txtCtcName"
                Display="None" ValidationGroup="vg" ErrorMessage="Contact required!" />
            <br />
        </div>
        <div>
            <label>
                Addr 1 :</label>
            <asp:TextBox ID="txtAddr1" runat="server" MaxLength="100" Width="200px" ValidationGroup="vg"
                Font-Size="XX-Small" />
            <asp:RequiredFieldValidator ID="rfvAddr1" runat="server" ControlToValidate="txtAddr1"
                Display="None" ValidationGroup="vg" ErrorMessage="Addr 1 required!" />
        </div>
        <div>
            <label>
                Addr 2 :</label>
            <asp:TextBox ID="txtAddr2" runat="server" MaxLength="100" Width="200px" ValidationGroup="vg"
                Font-Size="XX-Small" />
        </div>
        <div>
            <label>
                Addr 3 :</label>
            <asp:TextBox ID="txtAddr3" runat="server" MaxLength="100" Width="200px" ValidationGroup="vg"
                Font-Size="XX-Small" />
        </div>
        <div>
            <label>
                Town :</label>
            <asp:TextBox ID="txtTown" runat="server" MaxLength="100" Width="200px" ValidationGroup="vg"
                Font-Size="XX-Small" />
            <asp:RequiredFieldValidator ID="rfvTown" runat="server" ControlToValidate="txtTown"
                Display="None" ValidationGroup="vg" ErrorMessage="Town/City required!" />
        </div>
        <div>
            <label>
                Post Code :</label>
            <asp:TextBox ID="txtPostCode" runat="server" MaxLength="100" Width="200px" ValidationGroup="vg"
                Font-Size="XX-Small" />
            <asp:RequiredFieldValidator ID="rfvPostcode" runat="server" ControlToValidate="txtPostCode"
                Display="None" ValidationGroup="vg" ErrorMessage="Post Code required!" />
            <br />
        </div>
    </asp:Panel>
    <br />
    <asp:Panel ID="pnlReference" runat="server" BackColor="#ECFFEC">
        <div>
            <label>
                Reference :</label>
            <asp:TextBox ID="txtReference" runat="server" MaxLength="100" Width="200px" ValidationGroup="vg" />
            <asp:RequiredFieldValidator ID="rfvReference" runat="server" ControlToValidate="txtReference"
                Display="None" ValidationGroup="vg" ErrorMessage="Reference required!" />
            <br />
            <br />
        </div>
    </asp:Panel>
    <br />
    <div>
        <label>
            Select an Excel (.xlsx) file :</label>
        <telerik:RadUpload ID="ruAmazonFileUpload" Width="100%" TargetFolder="~/Amazon" AllowedFileExtensions=".xlsx"
            ToolTip="Select an Excel 2007 file (.xlsx)" MaxFileInputsCount="1" OverwriteExistingFiles="true"
            runat="server" BackColor="#FFE7CE" ControlObjectsVisibility="None" />
    </div>
    <asp:Label ID="lblNoResults" runat="server" Visible="false" Text="No files uploaded (note: only Excel 2007 file (.xlsx) accepted)"
        Font-Bold="true" />
    <div id="divFileInfo" runat="server" visible="false">
        File:
        <asp:Label ID="lblFileName" Text="" runat="server" Font-Bold="true" />
        <%--<label> &nbsp; Size :</label><asp:Label ID="lblFileSize" Text="" runat="server" />--%>
        <br />
        <br />
    </div>
    <div>
        <asp:Button runat="server" ID="btnUpload" Text="1. Upload Spreadsheet" OnClick="btnUpload_Click" />
        &nbsp;<asp:Button ID="btnCheckData" runat="server" Text="2. Check Data" OnClick="btnCheckData_Click"
            Enabled="False" />
        &nbsp;<asp:Button runat="server" ID="btnProcess" Text="3. Create Consignments" Enabled="false"
            ValidationGroup="vg" OnClick="btnProcess_Click" Style="height: 26px" />
    </div>
    <div>
        <div style="float: left; margin-left: 5px; width: 700px">
            <asp:UpdatePanel ID="upProducts" runat="server" UpdateMode="Always">
                <ContentTemplate>
                    <asp:Repeater ID="rpProducts" EnableViewState="true" runat="server" OnItemDataBound="rpProducts_ItemDataBound">
                        <HeaderTemplate>
                            <asp:Label ID="lblNoOfProducts" Text="List Of Products" runat="server" />
                            <table border="1" cellpadding="1" cellspacing="1" width="100%">
                                <tr>
                                    <th>
                                        Product Code
                                    </th>
                                    <th>
                                        Requested Quantity
                                    </th>
                                    <th>
                                        Available Qty
                                    </th>
                                    <th>
                                        Match Status
                                    </th>
                                </tr>
                        </HeaderTemplate>
                        <ItemTemplate>
                            <tr>
                                <td>
                                    <asp:Label ID="lblProductCode" Text='<%# Bind("ProductCode") %>' runat="server" />
                                </td>
                                <td>
                                    <asp:Label ID="Label1" Text='<%# Bind("RequestedQty") %>' runat="server" />
                                </td>
                                <td>
                                    <asp:Label ID="Label2" Text='<%# Bind("AvailableQty") %>' runat="server" />
                                </td>
                                <td>
                                    <%# IIf(DataBinder.Eval(Container.DataItem, "LogisticProductKey") = "-1", "<font color='red'>Product not found</font>", "<font color='green'>Product found</font>")%>
                                </td>
                            </tr>
                        </ItemTemplate>
                        <FooterTemplate>
                            </table>
                        </FooterTemplate>
                    </asp:Repeater>
                    <table id="tblPaging" width="100%" border="0" runat="server">
                        <tr>
                            <td>
                                <asp:Label ID="lblCurrentPage" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <asp:Button ID="cmdPrev" runat="server" Text=" << "></asp:Button>
                                <asp:Button ID="cmdNext" runat="server" Text=" >> "></asp:Button>
                            </td>
                        </tr>
                    </table>
                </ContentTemplate>
            </asp:UpdatePanel>
        </div>
        <div style="float: left; margin-right: 5px; width: 300px">
            <asp:Repeater ID="rpConsignments" runat="server">
                <HeaderTemplate>
                    <asp:Label ID="lblConsignments" Text="Consignment(s)" runat="server"/>
                    <table border="1" cellpadding="1" cellspacing="1" width="100%">
                        <tr>
                            <th>
                                Consignment Number
                            </th>
                        </tr>
                </HeaderTemplate>
                <ItemTemplate>
                    <tr>
                        <td>
                            <asp:Label ID="lblConsignmentNumber" Text='<%# Container.DataItem %>' runat="server"/>
                            <%--<asp:Label ID="Label3" Text='<%# Bind("Consignment") %>' runat="server"/>--%>
                        </td>
                    </tr>
                </ItemTemplate>
                <FooterTemplate>
                    </table>
                </FooterTemplate>
            </asp:Repeater>
        </div>
    </div>
    <p>
        &nbsp;</p>
    <p>
        &nbsp;</p>
    <p>
        <asp:TextBox ID="tbLog" runat="server" ReadOnly="True" TextMode="MultiLine" 
            Width="100%" Rows="10" Font-Names="Arial" Font-Size="XX-Small" />
    </p>
    </form>
</body>
</html>
