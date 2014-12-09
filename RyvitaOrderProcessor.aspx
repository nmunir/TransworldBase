<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Globalization" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Text.RegularExpressions" %>
<%@ Import Namespace="Telerik.Web.UI" %>
<%@ Register Assembly="Telerik.Web.UI" Namespace="Telerik.Web.UI" TagPrefix="telerik" %>
<%@ Import Namespace="System.Data.Common" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ Import Namespace="FileHelpers" %>
<script runat="server">
    
    ' Ryvita Order Exporter: https://www.ryvita.co.uk/export/csv_export.php   operator / VRnNeTe3RguZg
    
    'CREATE TABLE [dbo].[ClientData_Ryvita_OrdersPlaced](
    '[id] [int] IDENTITY(1,1) NOT NULL,
    '[RyvitaOrderRef] [varchar](50) NOT NULL,
    '[OrderDateTime] [smalldatetime] NOT NULL,
    '[AWB] [int] NOT NULL
    ') ON [PRIMARY]
 
    'CREATE TABLE [dbo].[ClientData_Ryvita_ProductMapping](
    '[id] [int] IDENTITY(1,1) NOT NULL,
    '[RyvitaProductCode] [varchar](50) NOT NULL,
    '[TransworldProductCode] [varchar](50) NOT NULL,
    '[TransworldProductKey] [int] NOT NULL
    ') ON [PRIMARY]
    
    ' TO DO
    ' dummy user to place orders
    
    Const CUSTOMER_RYVITA As Int32 = 813
    'Const CUSTOMER_RYVITA As Int32 = 16
    
    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsNumeric(Session("UserKey")) Then
            'Server.Transfer("session_expired.aspx")
        End If

        If Not IsPostBack Then
            Call CreateRyvitaOrdersFolder()
            Call SetTitle()
        End If
    End Sub
    
    Protected Sub SetTitle()
        Dim sTitle As String = Session("SiteTitle")
        If sTitle <> String.Empty Then
            sTitle += " - "
        End If
        Page.Header.Title = sTitle & "Create Ryvita Orders"
    End Sub
    
    Protected Sub ReadFile()
        Dim sFilePath As String = Server.MapPath("~/RyvitaOrders/" & psFileName)
        Dim RyvitaOrders As RyvitaOrder()
        Dim engine As New FileHelperEngine(GetType(RyvitaOrder))
        Try
            RyvitaOrders = DirectCast(engine.ReadFile(sFilePath), RyvitaOrder())  ' To Write Use: engine.WriteFile("FileOut.txt", res) 
        Catch ex As Exception
            Call WriteToLog("Could not read file: " & ex.Message)
            Exit Sub
        End Try
        
        Dim dtOrders As DataTable = BuildConsignmentTable()
        Dim dictTotalProducts As New Dictionary(Of Int32, Int32)
        
        Dim nOrderStatusValidated As Int32
        Dim nOrderStatusBad As Int32
        Dim nOrderStatusProcessed As Int32
        
        For Each o As RyvitaOrder In RyvitaOrders   'Console.WriteLine(cust.Name + " - " + cust.AddedDate.ToString("dd/MM/yy"))
            
            Call WriteToLog("")
            Call WriteToLog("Checking order " & o.sOrderRef)

            Dim drOrder As DataRow = dtOrders.NewRow()
            
            Dim sCneeFirstName As String = o.sUserFirstName.Trim
            If IsAllSameCase(sCneeFirstName) Then
                sCneeFirstName = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(sCneeFirstName.ToLower())
            End If

            Dim sCneeLastName As String = o.sUserLastName.Trim
            If IsAllSameCase(sCneeLastName) Then
                sCneeLastName = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(sCneeLastName.ToLower())
            End If

            drOrder("CneeName") = sCneeFirstName & " " & sCneeLastName
                
            drOrder("CneeAddr1") = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(o.sAddress1)

            drOrder("CneeTown") = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(o.sAddress2)

            drOrder("CneeCounty") = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(o.sAddress3)

            drOrder("CneePostcode") = o.sPostcode.ToUpper

            drOrder("OrderRef") = o.sOrderRef

            Dim bError As Boolean = False
            Dim lstProductsHumanReadableList As New List(Of String)
            Dim dictProducts As New Dictionary(Of Int32, Int32)
            Dim nProductKey As Int32
            
            nProductKey = IsValidProductAndQuantity(o.Product, o.Quantity)
            If nProductKey > 0 Then
                Call AddProduct(nProductKey, CInt(o.Quantity), lstProductsHumanReadableList, dictProducts, dictTotalProducts)
            ElseIf nProductKey < 0 Then
                bError = True
            End If

            nProductKey = IsValidProductAndQuantity(o.Product1, o.Quantity1)
            If nProductKey > 0 Then
                Call AddProduct(nProductKey, CInt(o.Quantity1), lstProductsHumanReadableList, dictProducts, dictTotalProducts)
            ElseIf nProductKey < 0 Then
                bError = True
            End If

            nProductKey = IsValidProductAndQuantity(o.Product2, o.Quantity2)
            If nProductKey > 0 Then
                Call AddProduct(nProductKey, CInt(o.Quantity2), lstProductsHumanReadableList, dictProducts, dictTotalProducts)
            ElseIf nProductKey < 0 Then
                bError = True
            End If

            nProductKey = IsValidProductAndQuantity(o.Product3, o.Quantity3)
            If nProductKey > 0 Then
                Call AddProduct(nProductKey, CInt(o.Quantity3), lstProductsHumanReadableList, dictProducts, dictTotalProducts)
            ElseIf nProductKey < 0 Then
                bError = True
            End If

            nProductKey = IsValidProductAndQuantity(o.Product4, o.Quantity4)
            If nProductKey > 0 Then
                Call AddProduct(nProductKey, CInt(o.Quantity4), lstProductsHumanReadableList, dictProducts, dictTotalProducts)
            ElseIf nProductKey < 0 Then
                bError = True
            End If

            nProductKey = IsValidProductAndQuantity(o.Product5, o.Quantity5)
            If nProductKey > 0 Then
                Call AddProduct(nProductKey, CInt(o.Quantity5), lstProductsHumanReadableList, dictProducts, dictTotalProducts)
            ElseIf nProductKey < 0 Then
                bError = True
            End If

            If bError Then
                WriteToLog("One or more incoming product fields had errors.")
                drOrder("StatusExcluded") = -1
                nOrderStatusBad += 1
            ElseIf OrderAlreadyPlaced(drOrder("OrderRef")) Then
                drOrder("StatusExcluded") = 1
                nOrderStatusProcessed += 1
            Else
                drOrder("StatusExcluded") = 0
                nOrderStatusValidated += 1
            End If
            
            drOrder("ProductsHumanReadableList") = String.Empty
            For Each sProduct As String In lstProductsHumanReadableList
                drOrder("ProductsHumanReadableList") &= sProduct & " "
            Next
            If drOrder("ProductsHumanReadableList") = String.Empty Then
                drOrder("ProductsHumanReadableList") = "One or more products invalid or not recognised! (" & o.Product & " " & o.Quantity & " " & o.Product1 & " " & o.Quantity1 & " " & o.Product2 & " " & o.Quantity2 & " " & o.Product3 & " " & o.Quantity3 & " " & o.Product4 & " " & o.Quantity4 & " " & o.Product5 & " " & o.Quantity5 & ")"
            End If

            drOrder("ProductsToOrder") = String.Empty
            For Each kv As KeyValuePair(Of Int32, Int32) In dictProducts
                drOrder("ProductsToOrder") &= kv.Key & "," & kv.Value & ","
            Next
            If drOrder("ProductsToOrder") <> String.Empty Then
                drOrder("ProductsToOrder") = drOrder("ProductsToOrder").ToString.Substring(0, drOrder("ProductsToOrder").ToString.Length - 1)
            End If

            dtOrders.Rows.Add(drOrder)
        Next
        
        Dim bAvailableQuantityError As Boolean = False
        
        For Each kv As KeyValuePair(Of Int32, Int32) In dictTotalProducts
            Dim nAvailableQuantity As Int32 = GetAvailableQty(kv.Key)
            If kv.Value > nAvailableQuantity Then
                WriteToLog("WARNING: Available quantity of product " & GetProductCodeFromProductKey(kv.Key) & " (" & nAvailableQuantity & ") is insufficient to fulfil all orders (required: " & kv.Value & ")")
                bAvailableQuantityError = True
            Else
                WriteToLog("Confirmed required quantity of product " & GetProductCodeFromProductKey(kv.Key) & " (" & kv.Value & ") is less than or equal to total amount available (" & nAvailableQuantity & ")")
            End If
        Next

        If bAvailableQuantityError Then
            lblInsufficientQuantity.Visible = True
        Else
            lblInsufficientQuantity.Visible = False
        End If
        pdtOrders = dtOrders
        
        gvOrders.DataSource = dtOrders
        gvOrders.DataBind()
        lblLegendOrders.Visible = True
        btnRecheckQuantities.Enabled = True
        btnCreateConsignments.Enabled = True

        Dim sMessage As String
        For Each kv As KeyValuePair(Of Int32, Int32) In dictTotalProducts
            Dim nAvailableQuantity As Int32 = GetAvailableQty(kv.Key)
            If kv.Value < nAvailableQuantity Then
                sMessage &= "Available quantity of product " & GetProductCodeFromProductKey(kv.Key) & " (" & nAvailableQuantity & ") is sufficient to fulfil all orders (required: " & kv.Value & ")" & "\r\r"
            Else
                sMessage &= "WARNING: available quantity of product " & GetProductCodeFromProductKey(kv.Key) & " (" & nAvailableQuantity & ") is insufficient to fulfil all orders (required: " & kv.Value & ")" & "\r\r"
            End If
        Next

        WebMsgBox.Show("Total orders: " & nOrderStatusValidated + nOrderStatusBad + nOrderStatusProcessed & "\r\rValidated Unprocessed Orders: " & nOrderStatusValidated & "\r\rFailed Validation: " & nOrderStatusBad & "\r\rAlready Processed: " & nOrderStatusProcessed & "\r\r" & sMessage)
    End Sub
    
    Protected Function OrderAlreadyPlaced(sOrderRef As String) As Boolean
        OrderAlreadyPlaced = False
        Dim sSQL As String = "SELECT TOP 1 * FROM ClientData_Ryvita_OrdersPlaced WHERE RyvitaOrderRef = '" & sOrderRef & "'"
        Dim dtOrderRecord As DataTable = ExecuteQueryToDataTable(sSQL)
        If dtOrderRecord.Rows.Count > 0 Then
            OrderAlreadyPlaced = True
        End If
    End Function
    
    Protected Sub AddProduct(ByVal nProductKey As Int32, ByVal nQuantity As Int32, ByRef lstProductHumanReadableList As List(Of String), ByRef dictProducts As Dictionary(Of Int32, Int32), ByRef dictTotalProducts As Dictionary(Of Int32, Int32))
        lstProductHumanReadableList.Add(GetProductCodeFromProductKey(nProductKey) & " (" & nQuantity.ToString & ")")
        Try
            dictProducts.Add(nProductKey, nQuantity)
        Catch ex As Exception
            dictProducts(nProductKey) = dictProducts(nProductKey) + nQuantity
        End Try
        Try
            dictTotalProducts.Add(nProductKey, nQuantity)
        Catch ex As Exception
            dictTotalProducts(nProductKey) = dictTotalProducts(nProductKey) + nQuantity
        End Try
    End Sub
    
    Protected Function GetProductCodeFromProductKey(ByVal nProductKey As Int32) As String
        GetProductCodeFromProductKey = ExecuteQueryToDataTable("SELECT TransworldProductCode FROM ClientData_Ryvita_ProductMapping WHERE TransworldProductKey = " & nProductKey).Rows(0).Item(0)
    End Function
    
    Protected Function IsValidProductAndQuantity(ByVal sProduct As String, ByVal sQuantity As String) As Int32  ' returns LogisticProductKey if valid, 0 if both fields empty, -1 if invalid
        IsValidProductAndQuantity = 0
        If sProduct <> String.Empty Then
            If IsNumeric(sQuantity) Then
                Dim nQuantity = CInt(sQuantity)
                If nQuantity > 0 Then
                    Dim nLogisticProductCode As Int32 = GetLogisticProductKeyFromRyvitaProductCode(sProduct)
                    If nLogisticProductCode > 0 Then
                        IsValidProductAndQuantity = nLogisticProductCode
                    Else
                        Call WriteToLog("Could not match product (" & sProduct & ").")
                        IsValidProductAndQuantity = -1
                    End If
                Else
                    Call WriteToLog("Quantity must be greater than zero.")
                    IsValidProductAndQuantity = -1
                End If
            Else
                Call WriteToLog("Non-blank product code with non-numeric quantity (" & sQuantity & ").")
                IsValidProductAndQuantity = -1
            End If
        Else
            If sQuantity <> String.Empty Then
                IsValidProductAndQuantity = -1
            End If
        End If
    End Function
    
    Protected Function GetLogisticProductKeyFromRyvitaProductCode(sRyvitaProductCode As String) As Int32
        GetLogisticProductKeyFromRyvitaProductCode = 0
        Dim sSQL As String = "SELECT TransworldProductKey FROM ClientData_Ryvita_ProductMapping WHERE RyvitaProductCode = '" & sRyvitaProductCode & "'"
        Dim dtProduct As DataTable = ExecuteQueryToDataTable(sSQL)
        If dtProduct.Rows.Count = 1 Then
            GetLogisticProductKeyFromRyvitaProductCode = dtProduct.Rows(0).Item(0)
        End If
    End Function
    
    Protected Function IsAllSameCase(sWord As String) As Boolean
        sWord = sWord.Trim
        Dim rCheckLower As Regex = New Regex("^[a-z]*$", RegexOptions.None)
        Dim mLower As Match = rCheckLower.Match(sWord)
        If mLower.Success Then
            Return True
        End If
        Dim rCheckUpper As Regex = New Regex("^[A-Z]*$", RegexOptions.None)
        Dim mUpper As Match = rCheckUpper.Match(sWord)
        If mUpper.Success Then
            Return True
        End If
        Return False
    End Function
    
    Protected Function GetProductsToOrder(sProductsToOrder As String) As Dictionary(Of Int32, Int32)
        Dim dictProducts As New Dictionary(Of Int32, Int32)
        Dim arrProductsToOrder() = sProductsToOrder.Split(",")
        For i As Int32 = 0 To arrProductsToOrder.Count - 1 Step 2
            dictProducts.Add(arrProductsToOrder(i), arrProductsToOrder(i + 1))
        Next
        GetProductsToOrder = dictProducts
    End Function
    
    Protected Function BuildConsignmentTable() As DataTable
        Dim dtOrders As New DataTable
        dtOrders.Columns.Add(New DataColumn("CneeName", Type.GetType("System.String")))
        dtOrders.Columns.Add(New DataColumn("CneeAddr1", Type.GetType("System.String")))
        dtOrders.Columns.Add(New DataColumn("CneeTown", Type.GetType("System.String")))
        dtOrders.Columns.Add(New DataColumn("CneePostcode", Type.GetType("System.String")))
        dtOrders.Columns.Add(New DataColumn("CneeCounty", Type.GetType("System.String")))
        dtOrders.Columns.Add(New DataColumn("OrderRef", Type.GetType("System.String")))
        dtOrders.Columns.Add(New DataColumn("ProductsHumanReadableList", Type.GetType("System.String")))
        dtOrders.Columns.Add(New DataColumn("ProductsToOrder", Type.GetType("System.String")))

        Dim dc As New DataColumn("StatusExcluded")
        dc.DataType = GetType(Integer)
        dc.AllowDBNull = True
        dtOrders.Columns.Add(dc)

        Return dtOrders
    End Function

    '         <FieldConverter(ConverterKind.Date, "dd/MM/yyyy mm:ss")> _        
    '         <FieldConverter(ConverterKind.Date, "yyyy-MM-dd hh:mm:ss")> _

    '<FieldQuoted("""", QuoteMode.OptionalForBoth)> _
    '<FieldConverter(ConverterKind.Date, "yyyy-MM-dd")> _
    'Public dtOrderDate As DateTime

    <DelimitedRecord(",")> _
    <IgnoreFirst(1)> _
    Public Class RyvitaOrder
        <FieldQuoted("""", QuoteMode.OptionalForBoth)> _
        Public sOrderDate As String
        
        <FieldQuoted("""", QuoteMode.OptionalForBoth)> _
        Public sUserFirstName As String
        
        <FieldQuoted("""", QuoteMode.OptionalForBoth)> _
        Public sUserLastName As String
        
        <FieldQuoted("""", QuoteMode.OptionalForBoth)> _
        Public sDeliveryFirstName As String
        
        <FieldQuoted("""", QuoteMode.OptionalForBoth)> _
        Public sDeliveryLastName As String
        
        <FieldQuoted("""", QuoteMode.OptionalForBoth)> _
        Public sAddress1 As String
        
        <FieldQuoted("""", QuoteMode.OptionalForBoth)> _
        Public sAddress2 As String
        
        <FieldQuoted("""", QuoteMode.OptionalForBoth)> _
        Public sAddress3 As String
        
        <FieldQuoted("""", QuoteMode.OptionalForBoth)> _
        Public sPostcode As String
        
        Public sOrderRef As String

        <FieldOptional()> _
        Public Product As String

        <FieldOptional()> _
        Public Quantity As String
        
        <FieldOptional()> _
        Public Product1 As String
        <FieldOptional()> _
        Public Quantity1 As String
        <FieldOptional()> _
        Public Product2 As String
        <FieldOptional()> _
        Public Quantity2 As String
        <FieldOptional()> _
        Public Product3 As String
        <FieldOptional()> _
        Public Quantity3 As String
        <FieldOptional()> _
        Public Product4 As String
        <FieldOptional()> _
        Public Quantity4 As String
        <FieldOptional()> _
        Public Product5 As String
        <FieldOptional()> _
        Public Quantity5 As String
    End Class

    Protected Sub btnCheckData_Click(sender As Object, e As System.EventArgs)
        Call ReadFile()
    End Sub

    Private Sub CreateRyvitaOrdersFolder()
        Dim sPath As String = Server.MapPath("~/")
        If Not Directory.Exists(sPath & "\RyvitaOrders") Then
            Directory.CreateDirectory(sPath & "\RyvitaOrders")
        End If
    End Sub

    Protected Sub btnUpload_Click(ByVal sender As Object, ByVal e As EventArgs)
        Try
            If ruRyvitaFileUpload.UploadedFiles.Count > 0 Then
                lblNoResults.Visible = False
                psFileName = ruRyvitaFileUpload.UploadedFiles(0).GetName
                lblFileName.Text = ruRyvitaFileUpload.UploadedFiles(0).GetName
                lblLegendFile.Visible = True
                tbLog.Text = String.Empty
                Call WriteToLog("Uploaded " & ruRyvitaFileUpload.UploadedFiles(0).GetName & " @ " & Format(Date.Now, "d-MMM-yyyy hh:mm:ss"))
                Call ReadFile()
            Else
                btnRecheckQuantities.Enabled = False
                btnCreateConsignments.Enabled = False
                
                lblNoResults.Visible = True
                divFileInfo.Visible = False
                Call WriteToLog("Nothing uploaded")
            End If
        Catch ex As Exception
            lblNoResults.Text = ex.Message.ToString()
            Call WriteToLog(ex.Message.ToString())
        End Try
    End Sub
    
    Private Function GetAvailableQty(ByVal nLogisticProductKey As Integer) As Integer
        Dim sSQL As String = "SELECT Quantity = CASE ISNUMERIC((SELECT SUM(LogisticProductQuantity) FROM LogisticProductLocation WHERE LogisticProductKey = " & nLogisticProductKey & ")) WHEN 0 THEN 0 ELSE (SELECT SUM(LogisticProductQuantity) FROM LogisticProductLocation WHERE LogisticProductKey = " & nLogisticProductKey & ") END"
        Dim oDataTable As DataTable = ExecuteQueryToDataTable(sSQL)
        If oDataTable IsNot Nothing AndAlso oDataTable.Rows.Count <> 0 Then
            GetAvailableQty = oDataTable.Rows(0)(0)
        Else
            GetAvailableQty = 0
        End If
    End Function

    Protected Sub WriteToLog(sMessage As String)
        tbLog.Text += sMessage & Environment.NewLine
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

    Property psFileName() As String
        Get
            Dim o As Object = ViewState("Ryvita_Order_FileName")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("Ryvita_Order_FileName") = Value
        End Set
    End Property

    Property pdtOrders() As DataTable
        Get
            Dim o As Object = ViewState("RyvitaOrders")
            If o Is Nothing Then
                Return Nothing
            End If
            Return o
        End Get
        Set(ByVal Value As DataTable)
            ViewState("RyvitaOrders") = Value
        End Set
    End Property

    Protected Sub btnRecheckQuantities_Click(sender As Object, e As System.EventArgs)
        Dim dictTotalProducts As New Dictionary(Of Int32, Int32)
        Dim dtOrders As DataTable = pdtOrders
        For Each drOrder As DataRow In dtOrders.Rows
            If drOrder("StatusExcluded") = 0 Then
                Dim dictProductsToOrder As New Dictionary(Of Int32, Int32)
                dictProductsToOrder = GetProductsToOrder(drOrder("ProductsToOrder"))
                For Each kv As KeyValuePair(Of Int32, Int32) In dictProductsToOrder
                    Try
                        dictTotalProducts.Add(kv.Key, kv.Value)
                    Catch ex As Exception
                        dictTotalProducts(kv.Key) = dictTotalProducts(kv.Key) + kv.Value
                    End Try
                Next
            End If
        Next

        Dim sMessage As String = String.Empty
        For Each kv As KeyValuePair(Of Int32, Int32) In dictTotalProducts
            Dim nAvailableQuantity As Int32 = GetAvailableQty(kv.Key)
            If kv.Value < nAvailableQuantity Then
                sMessage &= "Available quantity of product " & GetProductCodeFromProductKey(kv.Key) & " (" & nAvailableQuantity & ") is sufficient to fulfil all orders (required: " & kv.Value & ")" & "\r\r"
            Else
                sMessage &= "WARNING: available quantity of product " & GetProductCodeFromProductKey(kv.Key) & " (" & nAvailableQuantity & ") is insufficient to fulfil all orders (required: " & kv.Value & ")" & "\r\r"
            End If
        Next
        WebMsgBox.Show(sMessage)
    End Sub
    
    Protected Function nSubmitConsignment(ByRef drOrder As DataRow) As Int32
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
        param1.Value = GetGenericUser()
        oCmdAddBooking.Parameters.Add(param1)
        Dim param2 As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
        param2.Value = CUSTOMER_RYVITA
        oCmdAddBooking.Parameters.Add(param2)
        Dim param2a As SqlParameter = New SqlParameter("@BookingOrigin", SqlDbType.NVarChar, 20)
        param2a.Value = "WEB_BOOKING"
        oCmdAddBooking.Parameters.Add(param2a)
        Dim param3 As SqlParameter = New SqlParameter("@BookingReference1", SqlDbType.NVarChar, 25)
        param3.Value = drOrder("OrderRef")
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
        oCmdAddBooking.Parameters.Add(param7)
        Dim param8 As SqlParameter = New SqlParameter("@PackingNoteInfo", SqlDbType.NVarChar, 1000)
        param8.Value = ""
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
        
        Dim sSQL As String = "SELECT * FROM Customer WHERE CustomerKey = " & CUSTOMER_RYVITA
        Dim dtCnor As DataTable = ExecuteQueryToDataTable(sSQL)
        
        Dim drCnor As DataRow
        If dtCnor.Rows.Count = 1 Then
            drCnor = dtCnor.Rows(0)
        Else
            WebMsgBox.Show("Couldn't find Consignor details.")
            Exit Function
        End If
       
        Dim param13 As SqlParameter = New SqlParameter("@CnorName", SqlDbType.NVarChar, 50)
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
        param25.Value = drOrder("CneeName")
        oCmdAddBooking.Parameters.Add(param25)
        Dim param26 As SqlParameter = New SqlParameter("@CneeAddr1", SqlDbType.NVarChar, 50)
        param26.Value = drOrder("CneeAddr1")
        oCmdAddBooking.Parameters.Add(param26)
        Dim param27 As SqlParameter = New SqlParameter("@CneeAddr2", SqlDbType.NVarChar, 50)
        param27.Value = ""
        oCmdAddBooking.Parameters.Add(param27)
        Dim param28 As SqlParameter = New SqlParameter("@CneeAddr3", SqlDbType.NVarChar, 50)
        param28.Value = ""
        oCmdAddBooking.Parameters.Add(param28)
        Dim param29 As SqlParameter = New SqlParameter("@CneeTown", SqlDbType.NVarChar, 50)
        param29.Value = drOrder("CneeTown")
        oCmdAddBooking.Parameters.Add(param29)
        Dim param30 As SqlParameter = New SqlParameter("@CneeState", SqlDbType.NVarChar, 50)
        param30.Value = drOrder("CneeCounty")
        oCmdAddBooking.Parameters.Add(param30)
        Dim param31 As SqlParameter = New SqlParameter("@CneePostCode", SqlDbType.NVarChar, 50)
        param31.Value = drOrder("CneePostcode")
        oCmdAddBooking.Parameters.Add(param31)
        Dim param32 As SqlParameter = New SqlParameter("@CneeCountryKey", SqlDbType.Int, 4)
        param32.Value = 222
        oCmdAddBooking.Parameters.Add(param32)
        Dim param33 As SqlParameter = New SqlParameter("@CneeCtcName", SqlDbType.NVarChar, 50)
        param33.Value = ""
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
                Dim dictProductsToOrder As New Dictionary(Of Int32, Int32)
                dictProductsToOrder = GetProductsToOrder(drOrder("ProductsToOrder"))
                For Each kv As KeyValuePair(Of Int32, Int32) In dictProductsToOrder
                    Dim oCmdAddStockItem As SqlCommand = New SqlCommand("spASPNET_LogisticMovement_Add", oConn)
                    oCmdAddStockItem.CommandType = CommandType.StoredProcedure
                    Dim param51 As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int, 4)
                    param51.Value = GetGenericUser()
                    oCmdAddStockItem.Parameters.Add(param51)
                    Dim param52 As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
                    param52.Value = CUSTOMER_RYVITA
                    oCmdAddStockItem.Parameters.Add(param52)
                    Dim param53 As SqlParameter = New SqlParameter("@LogisticBookingKey", SqlDbType.Int, 4)
                    param53.Value = lBookingKey
                    oCmdAddStockItem.Parameters.Add(param53)
                    Dim param54 As SqlParameter = New SqlParameter("@LogisticProductKey", SqlDbType.Int, 4)
                    param54.Value = kv.Key
                    oCmdAddStockItem.Parameters.Add(param54)
                    Dim param55 As SqlParameter = New SqlParameter("@LogisticMovementStateId", SqlDbType.NVarChar, 20)
                    param55.Value = "PENDING"
                    oCmdAddStockItem.Parameters.Add(param55)
                    Dim param56 As SqlParameter = New SqlParameter("@ItemsOut", SqlDbType.Int, 4)
                    param56.Value = kv.Value
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
            Else
                oTrans.Rollback("AddBooking")
            End If
        Catch ex As SqlException
            oTrans.Rollback("AddBooking")
        Finally
            oConn.Close()
        End Try
    End Function
    
    Protected Function GetGenericUser() As Int32
        GetGenericUser = 0
        Dim sSQL As String = "SELECT [key] FROM UserProfile WHERE UserID = 'ryvitaGU' AND CustomerKey = " & CUSTOMER_RYVITA
        Dim dtGenericUser As DataTable = ExecuteQueryToDataTable(sSQL)
        If dtGenericUser.Rows.Count = 1 Then
            GetGenericUser = dtGenericUser.Rows(0).Item(0)
        End If
    End Function
    
    Protected Sub btnCreateConsignments_Click(sender As Object, e As System.EventArgs)
        Dim dtOrders As DataTable = pdtOrders
        Dim nOrderCount As Int32
        Dim bConsignmentFailed As Boolean = False
        For Each drOrder As DataRow In dtOrders.Rows
            If drOrder("StatusExcluded") = 0 Then
                Dim nConsignmentKey As Int32 = nSubmitConsignment(drOrder)
                If nConsignmentKey > 0 Then
                    WriteToLog("Order " & drOrder("OrderRef") & " successfully created as consignment " & nConsignmentKey.ToString)
                    nOrderCount += 1
                    Dim sSQL As String = "INSERT INTO ClientData_Ryvita_OrdersPlaced (RyvitaOrderRef, OrderDateTime, AWB) VALUES ('" & drOrder("OrderRef") & "', GETDATE(), " & nConsignmentKey & ")"
                    Call ExecuteQueryToDataTable(sSQL)
                Else
                    WriteToLog("WARNING: Could not create consignment for order " & drOrder("OrderRef"))
                    bConsignmentFailed = True
                End If
            End If
        Next
        If bConsignmentFailed Then
            lblConsignmentFailed.Visible = True
        End If
        WebMsgBox.Show("Created " & nOrderCount & " consignment(s)")
        btnCreateConsignments.Enabled = False
    End Sub
    
    Protected Sub lnkbtnExclude_Click(sender As Object, e As System.EventArgs)
        Dim lnkbtn As LinkButton = sender
        Dim sOrderRef As String = lnkbtn.CommandArgument
        Dim dtOrders As DataTable = pdtOrders
        If Not lnkbtn.Text.Contains("RESTORE") Then
            For Each drOrder As DataRow In dtOrders.Rows
                If sOrderRef = drOrder("OrderRef") Then
                    drOrder("StatusExcluded") = 1
                    Exit For
                End If
            Next
        Else
            For Each drOrder As DataRow In dtOrders.Rows
                If sOrderRef = drOrder("OrderRef") Then
                    drOrder("StatusExcluded") = 0
                    Exit For
                End If
            Next
        End If
        pdtOrders = dtOrders
        gvOrders.DataSource = dtOrders
        gvOrders.DataBind()
    End Sub
    
    Protected Sub gvOrders_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs)
        Dim gvr As GridViewRow = e.Row
        If gvr.RowType = DataControlRowType.DataRow Then
            Dim hidExcluded As HiddenField
            Dim lnkbtnExclude As LinkButton
            hidExcluded = gvr.Cells(0).FindControl("hidExcluded")
            lnkbtnExclude = gvr.Cells(0).FindControl("lnkbtnExclude")
            If hidExcluded.Value < 0 Then
                gvr.BackColor = Drawing.Color.Red
                lnkbtnExclude.Text = "excluded"
                lnkbtnExclude.Enabled = False
            ElseIf hidExcluded.Value > 0 Then
                gvr.BackColor = Drawing.Color.LightGray
                lnkbtnExclude.Text = "RESTORE"
            End If
        End If
    End Sub
    
    Protected Sub lnkbtnCheckProductTranslation_Click(sender As Object, e As System.EventArgs)
        Dim sSQL As String = "SELECT RyvitaProductCode 'Ryvita Product Code', TransworldProductCode 'Displayed Product Code', TransworldProductKey 'Transworld Product Key', lp.ProductCode 'Transworld Product Code', lp.ProductDescription 'Description' FROM ClientData_Ryvita_ProductMapping rpm INNER JOIN LogisticProduct lp ON rpm.TransworldProductKey = lp.LogisticProductKey ORDER BY RyvitaProductCode"
        Dim dtProductTranslation As DataTable = ExecuteQueryToDataTable(sSQL)
        gvProductTranslation.DataSource = dtProductTranslation
        gvProductTranslation.DataBind()
        pnlProductTranslation.Visible = True
    End Sub
    
    Protected Sub lnkbtnHideProductTranslationTable_Click(sender As Object, e As System.EventArgs)
        pnlProductTranslation.Visible = False
    End Sub
    
</script>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Create Ryvita Orders</title>
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
                <label>
                    <asp:Label ID="lblTitle" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="Small">Ryvita Order Processor</asp:Label>
                </label>
            </td>
            <td style="width: 50%; white-space: nowrap" align="right">
                &nbsp;
            </td>
        </tr>
    </table>
    <asp:ScriptManager ID="ScriptManager1" runat="server" />
    <div>
            <asp:Label ID="lblLegendSelect" Text="1. Select the Ryvita CSV order file:" runat="server"
                Font-Bold="False" Font-Names="Verdana" Font-Size="X-Small" />
        &nbsp;<telerik:RadUpload ID="ruRyvitaFileUpload" Width="100%" TargetFolder="~/RyvitaOrders"
            AllowedFileExtensions=".csv" ToolTip="Select an Excel 2007 file (.xlsx)" MaxFileInputsCount="1"
            OverwriteExistingFiles="true" runat="server" BackColor="#FFE7CE" ControlObjectsVisibility="None" />
        <label>
            <asp:Label ID="lblLegendFile" Text="File:" runat="server" Font-Bold="False" Font-Names="Verdana"
                Font-Size="X-Small" Visible="False" />
            <asp:Label ID="lblFileName" runat="server" Font-Bold="True" Font-Names="Verdana"
                Font-Size="Small" />
            <br />
        </label>
        <br />
        <asp:Button runat="server" ID="btnUpload" Text="2. Upload &amp; Check Spreadsheet"
            OnClick="btnUpload_Click" Width="200px" />
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        <asp:Button ID="btnRecheckQuantities" runat="server" OnClick="btnRecheckQuantities_Click"
            Text="2a. Re-check Quantities" Width="150px" Enabled="False" />
        &nbsp;<asp:Label ID="lblAdvice" Text="(eg after including or excluding one or more orders)"
            runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="X-Small" />
        <br />
        <br />
        <asp:Button runat="server" ID="btnCreateConsignments" Text="3. Create Consigments"
            Enabled="false" ValidationGroup="vg" Width="200px" OnClick="btnCreateConsignments_Click" />
    </div>
    <br />
    <asp:Label ID="lblNoResults" runat="server" Visible="false" Text="No file uploaded."
        Font-Names="Verdana" Font-Size="X-Small" Font-Bold="true" />
    &nbsp;<asp:Label ID="lblInsufficientQuantity" runat="server" Visible="False" Text="One or more products has insufficient quantity to fulfil all orders."
        Font-Bold="True" ForeColor="Red" Font-Names="Verdana" Font-Size="X-Small" />
    <asp:Label ID="lblConsignmentFailed" runat="server" Visible="False" Text="One or more consignments failed."
        Font-Bold="True" ForeColor="Red" Font-Names="Verdana" Font-Size="X-Small" />
    <br />
    <asp:Panel ID="pnlProductTranslation" runat="server" Width="100%" Visible="false">
        <asp:Label ID="lblLegendJournal0" runat="server" Font-Bold="False" Font-Names="Verdana"
            Font-Size="X-Small" Text="Product Translation Table:" />
        &nbsp;<asp:LinkButton ID="lnkbtnHideProductTranslationTable" runat="server" 
            Font-Names="Verdana" Font-Size="XX-Small" 
            onclick="lnkbtnHideProductTranslationTable_Click">hide</asp:LinkButton>
        <asp:GridView ID="gvProductTranslation" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
            Width="100%">
        </asp:GridView>
        <br />
    </asp:Panel>
    <div id="divFileInfo" runat="server" visible="false">
        &nbsp;
        <br />
        <br />
    </div>
    <asp:Label ID="lblLegendJournal" Text="Journal:" runat="server" Font-Bold="False"
        Font-Names="Verdana" Font-Size="X-Small" />
    <asp:TextBox ID="tbLog" runat="server" ReadOnly="True" TextMode="MultiLine" Width="99%"
        Rows="10" Font-Names="Arial" Font-Size="XX-Small" />
    <br />
    <br />
    <label>
        <asp:Label ID="lblLegendOrders" Text="Orders:" runat="server" Font-Bold="False" Font-Names="Verdana"
            Font-Size="X-Small" Visible="False" />
    </label>
    <asp:GridView ID="gvOrders" runat="server" CellPadding="1" Font-Names="Verdana" Font-Size="XX-Small"
        Width="100%" AutoGenerateColumns="False" OnRowDataBound="gvOrders_RowDataBound">
        <Columns>
            <asp:TemplateField>
                <ItemTemplate>
                    <asp:LinkButton ID="lnkbtnExclude" runat="server" CommandArgument='<%# Container.DataItem("OrderRef")%>'
                        OnClick="lnkbtnExclude_Click">exclude</asp:LinkButton>
                    <asp:HiddenField ID="hidExcluded" runat="server" Value='<%# Container.DataItem("StatusExcluded")%>' />
                </ItemTemplate>
                <ItemStyle Width="70px" />
            </asp:TemplateField>
            <asp:BoundField DataField="OrderRef" HeaderText="Order Ref" ReadOnly="True" SortExpression="OrderRef" />
            <asp:BoundField DataField="CneeName" HeaderText="Consignee" ReadOnly="True" SortExpression="CneeName" />
            <asp:BoundField DataField="CneeAddr1" HeaderText="Address" ReadOnly="True" SortExpression="CneeAddr1" />
            <asp:BoundField DataField="CneeTown" HeaderText="Town / City" ReadOnly="True" SortExpression="CneeTown" />
            <asp:BoundField DataField="CneeCounty" HeaderText="Country" ReadOnly="True" SortExpression="CneeCounty" />
            <asp:BoundField DataField="CneePostcode" HeaderText="Post code" ReadOnly="True" SortExpression="CneePostcode" />
            <asp:BoundField DataField="ProductsHumanReadableList" HeaderText="Products" ReadOnly="True"
                SortExpression="ProductsHumanReadableList" />
        </Columns>
        <EmptyDataTemplate>
            no orders found
        </EmptyDataTemplate>
    </asp:GridView>
    <asp:Panel ID="pnlHelp" runat="server" Width="100%" Font-Names="Verdana" Font-Size="XX-Small">
        <br />
        INSTRUCTIONS &amp; NOTES:<br />
        <br />
        1.&nbsp; Download the order file from the Ryvita web site ( <b><a href="https://www.ryvita.co.uk/export/csv_export.php">
            https://www.ryvita.co.uk/export/csv_export.php</a></b> Username/password: <b>operator</b>
        / <b>VRnNeTe3RguZg</b> ).<br />
        <br />
        2. Click the <b>Select</b> button. Navigate to the downloaded file.<br />
        <br />
        3.&nbsp; Click the button labelled <b>2. Upload &amp; Check Spreadsheet</b> .
        <br />
        <br />
        The system reads the CSV order file and displays the list of orders for you to 
        view before creating consignments,<br />
        <br />
        It validates each order and checks that enough quantity of each requested 
        product exists. Any order that cannot be read is marked is marked &#39;excluded&#39; and 
        highlighted in red. These excluded orders must be investigated and corrected.<br />
        <br />
        It checks the order reference number against the list of orders already placed. 
        Any order that appears already to have been placed is marked &#39;excluded&#39; and 
        highlighted in grey.&nbsp; Orders so excluded can be re-included by clicking the 
        RESTORE link. You can therefore download overlapping blocks of orders to ensure 
        none are missed, and leave the system to exclude any orders that have already 
        been placed.<br />
        <br />
        Click <b>exclude</b> to remove any other order you do not want generated (eg 
        because of address problems).<br />
        <br />
        4.&nbsp; After excluding or restoring orders you can re-check the button <b>2a 
        Re-check Quantities</b> to verify that enough of each requested product is in stock.<br />
        <br />
        5.&nbsp; Click the <b>Create Consignments</b> button to generate the orders in the
        list. The system reports the number of consignments generated.<br />
        <br />
        The incoming order file uses product codes such as RCB, RTRed to refer to Ryvita
        products&nbsp; The system translates these codes to the products configured in the
        Transworld stock system. To view the product translation table click
        <asp:LinkButton ID="lnkbtnCheckProductTranslation" runat="server" Font-Names="Verdana"
            Font-Size="XX-Small" OnClick="lnkbtnCheckProductTranslation_Click">check product translation</asp:LinkButton>
        &nbsp;. The translation table can only be modified by Development at this time.<br />
        <br />
    </asp:Panel>
    </form>
</body>
</html>
