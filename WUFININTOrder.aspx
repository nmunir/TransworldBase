<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<script runat="server">

    ' TO DO
    ' remove workaround for FININTgu and comment in timeout check
    
    ' check if need to load data in newagents_20120802.csv
    
    ' Format of FININT Order File
    ' ID NUMBER   DATETIME  SUBAGENTID  STOCKREF  QTY
    ' 122180,20120729144920,8687,33,3
    ' 122219,20120802163105,274,33,3
    ' 122280,20120813132753,4417,61,1
    ' 122280,20120813132753,4417,81,1
    ' 122215,20120802151037,1009,33,6
    ' 122241,20120806165644,3934,33,2
    ' 122281,20120813135456,5045,33,3
    ' 122243,20120807103049,3864,33,6
    ' 122168,20120726184040,1009,33,5
    ' 122146,20120723194033,3048,33,2
    ' 122273,20120813094900,6765,33,1

    ' Use ProductID instead of ProductCode (they should be synonymous)

    ' Format of COSTA Order File
    'order ID	Time/date	identifier	product code	quantity
    '1	230520131000	VM7O	Cos-001	2
    '2	230520131000	IGV2	Cos-001	8
    '3	230520131000	T6A7	Cos-001	20
    '4	230520131000	T6A7	Cos-005	10
    '5	230520131000	ITMG	Cos-003	10
    '6	230520131000	ITMG	Cos-005	10
    '7	230520131000	Q25T	Cos-001	4
    '8	230520131000	TRL8	Cos-003	5
    '9	230520131000	VM8S	Cos-005	10
    '10	230520131000	VM8S	Cos-007	1
    '11	230520131000	16TH	Cos-003	6
    '12	230520131000	KC8U	Cos-001	2
    '13	230520131000	ZEOJ	Cos-001	8
    '14	230520131000	A330	Cos-001	4
    '15	230520131000	CTQR	Cos-001	4
    '16	230520131000	ZUZX	Cos-001	10


    Const CUSTOMER_WUFIN As Int32 = 798
    Const CUSTOMER_WUCOSTA As Int32 = 826
    
    Dim sOriginalFileName As String, sUniqueFileName As String
    Dim gdtOrderData As DataTable, gdtConflatedOrderData As DataTable
    Dim sFilePrefix As String, sFileSuffix As String = ".csv"
    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    Dim gnCustomerKey As Int32

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsNumeric(Session("UserKey")) Then
            Server.Transfer("session_expired.aspx")
        End If
        If Not IsPostBack Then
            Call SetTitle()
        End If
        If rbCosta.Checked Then
            gnCustomerKey = CUSTOMER_WUCOSTA
        ElseIf rbFinint.Checked Then
            gnCustomerKey = CUSTOMER_WUFIN
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
        Page.Header.Title = sTitle & "Process COSTA or FININT Order Spreadsheet"
    End Sub

    Protected Sub AddProductToTotal(ByVal nProductKey As Int32, ByVal nQuantity As Int32, ByRef dictTotalProducts As Dictionary(Of Int32, Int32))
        Try
            dictTotalProducts.Add(nProductKey, nQuantity)
        Catch ex As Exception
            dictTotalProducts(nProductKey) = dictTotalProducts(nProductKey) + nQuantity
        End Try
    End Sub

    Protected Sub btnReadOrders_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not (rbCosta.Checked Or rbFinint.Checked) Then
            WebMsgBox.Show("Please select COSTA or FININT.")
            Exit Sub
        End If
        Call ReadOrders()
    End Sub
    
    Protected Sub ReadOrders()
        Dim sFileName As String = String.Empty
        Dim sUserMessage As String = String.Empty
        Dim nLineCount As Integer
        
        lblError.Text = ""
        lblHeaderWarning.Text = String.Empty
        lblMessage.Text = String.Empty
        
        sFileName = FileUpload1.FileName
        If Path.GetExtension(sFileName).ToLower <> ".csv" Then
            WebMsgBox.Show("This is not a .CSV file.\n\nDid you forget to convert it perhaps?")
            Exit Sub
        End If
        If My.Computer.FileSystem.FileExists(sFileName) Then
            WebMsgBox.Show("Could not find file " & sFileName)
            Exit Sub
        End If
        If FileUpload1.HasFile Then
            sFilePrefix = Format(Now(), "yyyymmddhhmmssff")
            sUniqueFileName = Server.MapPath("") & "\" & sFilePrefix & sFileName
            FileUpload1.SaveAs(sUniqueFileName)
            psUniqueFilename = sUniqueFileName
            Dim bPossibleHeaderLine As Boolean = False
            bPossibleHeaderLine = RemoveHeaderLineIfNecessaryPlusBlankLines(sUniqueFileName)
            If bPossibleHeaderLine And Not cbColumnHeadingsInRow1.Checked Then
                lblHeaderWarning.Text = "ERROR: Detected what appears to be a header line, but 'Column headings in row 1' check box is not set; "
            End If
            If Not bPossibleHeaderLine And cbColumnHeadingsInRow1.Checked Then
                lblHeaderWarning.Text = "ERROR: Did not detect anything that appears to be a header line, but 'Column headings in row 1' check box is set; "
            End If
            nLineCount = nCSVLineCount(sUniqueFileName) ' return number of lines in CSV order file
            lblMessage.Text = "Found " & nLineCount.ToString & " order lines."
            lblError.Text = FileIsValid()
            If lblHeaderWarning.Text = String.Empty And (lblError.Text = String.Empty Or lblError.Text.StartsWith("Ignored")) Then
                btnGenerateOrderSpreadsheet.Enabled = True
                btnGenerateOrders.Enabled = True
                
                If Not GetConflatedOrderTableOneProductPerRow() Then
                    sUserMessage &= "One or more Product IDs cannot be mapped to a Product Code."
                Else
                    gvOrders.DataSource = gdtConflatedOrderData
                    gvOrders.DataBind()
                End If
            End If
        Else
            sUserMessage &= "Specified file could not be found or file could not be processed."
            FileUpload1.Focus()
            Exit Sub
        End If
        If sUserMessage <> String.Empty Then
            WebMsgBox.Show(sUserMessage)
        End If
        If My.Computer.FileSystem.FileExists(sFileName) Then
            My.Computer.FileSystem.DeleteFile(sFileName)
        End If
    End Sub
    
    Protected Function FileIsValid() As String
        Dim sr As New StreamReader(sUniqueFileName)
        Dim sLine As String = String.Empty
        Dim sbMessage As New StringBuilder
        Dim sLineElements() As String
        Dim nLineCount As Int32 = 0
        Dim nBlankLineCount As Int32 = 0
        Dim bDetectedNonNumericDateField As Boolean = False
        Do While sr.Peek >= 0
            sLine = sr.ReadLine()
            nLineCount += 1
            sLineElements = sLine.Split(",")
            If nLineCount = 1 AndAlso sLineElements(0).ToLower.Contains("order") Then   ' skip first line which appears to be headers
                sLine = sr.ReadLine()
                nLineCount += 1
                sLineElements = sLine.Split(",")
            End If
            
            If sLineElements.Count <> 5 Then
                sbMessage.Append("Line " & nLineCount & ": " & sLineElements.Count & " elements found when 5 were expected.<br />")
            Else
                If sLineElements(0) = String.Empty And sLineElements(1) = String.Empty And sLineElements(2) = String.Empty And sLineElements(3) = String.Empty And sLineElements(4) = String.Empty Then
                    nBlankLineCount += 1
                    nLineCount -= 1
                Else
                    If Not IsNumeric(sLineElements(0)) Then
                        sbMessage.Append("Line " & nLineCount & ": expected numeric reference field.<br />")
                    End If
                    If Not IsNumeric(sLineElements(1)) Then
                        'sbMessage.Append("Line " & nLineCount & ": expected valid, all numeric, date field (did you forget to format the date column as 'Custom 0' before converting to CSV?).<br />")
                    End If
                    If rbFinint.Checked Then
                        If Not AgentIDExists_FININT(sLineElements(2)) Then
                            sbMessage.Append("Line " & nLineCount & ": unidentified Agent ID (" & sLineElements(2) & ").<br />")
                        End If
                    Else
                        If Not AgentIDExists_COSTA(sLineElements(2)) Then
                            sbMessage.Append("Line " & nLineCount & ": unidentified Agent ID (" & sLineElements(2) & ").<br />")
                        End If
                    End If
                    If Not ProductIDExists(sLineElements(3)) Then
                        sbMessage.Append("Line " & nLineCount & ": unidentified Product ID (" & sLineElements(3) & ").<br />")
                    End If
                    If Not IsNumeric(sLineElements(4)) Then
                        sbMessage.Append("Line " & nLineCount & ": expected numeric quantity field.<br />")
                    End If
                End If
            End If
        Loop
        sr.Close()
        If nBlankLineCount > 0 Then
            sbMessage.Append("Ignored " & nBlankLineCount & " blank line(s).<br />")
        End If
        FileIsValid = sbMessage.ToString
    End Function
    
    Protected Function AgentIDExists_FININT(ByVal sAgentID As String) As Boolean
        AgentIDExists_FININT = False
        Dim sSQL As String = "SELECT AgentID FROM ClientData_WU_LegacyNetwork WHERE AgentID = " & sAgentID
        Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
        If dt.Rows.Count > 1 Then
            WebMsgBox.Show("Error in AgentIDExists_FININT: non-unique Agent ID detected.")
        Else
            AgentIDExists_FININT = (dt.Rows.Count = 1)
        End If
    End Function
    
    Protected Function AgentIDExists_COSTA(ByVal sAgentID As String) As Boolean
        AgentIDExists_COSTA = False
        Dim sSQL As String = "SELECT TermID FROM ClientData_WUCOSTA_Agents WHERE TermID = '" & sAgentID & "'"
        Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
        If dt.Rows.Count > 1 Then
            WebMsgBox.Show("Error in AgentIDExists_COSTA: non-unique Agent ID detected.")
        Else
            AgentIDExists_COSTA = (dt.Rows.Count = 1)
        End If
    End Function
    
    Protected Function ProductIDExists(ByVal sProductID As String) As Boolean
        ProductIDExists = False
        Dim sSQL As String = "SELECT LogisticProductKey FROM LogisticProduct WHERE CustomerKey = " & gnCustomerKey & " AND ProductCode  = '" & sProductID & "'"
        Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
        If dt.Rows.Count > 1 Then
            WebMsgBox.Show("Error: non-unique Product ID detected.")
        Else
            ProductIDExists = (dt.Rows.Count = 1)
        End If
    End Function

    Protected Function nCSVLineCount(ByVal sFileName As String) As Integer ' return number of lines in CSV order file
        Dim sr As New StreamReader(sFileName)
        Dim nLineCount As Integer = 0
        Do While sr.Peek >= 0
            sr.ReadLine()
            nLineCount += 1
        Loop
        sr.Close()
        nCSVLineCount = nLineCount
    End Function

    Protected Function RemoveHeaderLineIfNecessaryPlusBlankLines(ByVal sFileName As String) As Boolean   ' return true if possible header line detected, irrespective of 'Column headings in row 1' setting
        RemoveHeaderLineIfNecessaryPlusBlankLines = False
        Dim sr As New StreamReader(sFileName)
        Dim sTempFilename As String = sFileName.Substring(0, sFileName.Length - 4) & "_TEMP.csv"
        Dim sw As New StreamWriter(sTempFilename)
        Dim bProcessedFirstLine As Boolean = False
        Dim sLine As String
        Do While sr.Peek >= 0
            sLine = sr.ReadLine()
            If Not bProcessedFirstLine Then
                If sLine.ToLower.Contains("date") Then RemoveHeaderLineIfNecessaryPlusBlankLines = True
                If cbColumnHeadingsInRow1.Checked Then
                    sLine = sr.ReadLine()
                End If
                bProcessedFirstLine = True
            End If
            sLine = sLine.Replace(" ", "")
            sLine = sLine.Replace(",,", ",")
            If sLine.Substring(sLine.Length - 1, 1) = "," Then
                sLine = sLine.Substring(0, sLine.Length - 1)
            End If
            If sLine.Replace(",", "").Trim <> String.Empty Then
                sw.WriteLine(sLine)
            End If
        Loop
        sr.Close()
        sw.Close()
        
        If My.Computer.FileSystem.FileExists(sTempFilename) Then
            Dim sMsg As String = "it exists!"
        End If

        My.Computer.FileSystem.DeleteFile(sFileName)
        My.Computer.FileSystem.RenameFile(sTempFilename, Path.GetFileName(sFileName))
    End Function
    
    Protected Sub OutputErrorMessage(ByVal sMessage As String) ' ensures any existing message is not overwritten
        If lblError.Text = "" Then
            lblError.Text = sMessage
        End If
    End Sub

    Protected Function GetProductKeyFromProductID(ByVal sProductID As String) As Int32
        GetProductKeyFromProductID = 0
        Dim sSQL As String = "SELECT LogisticProductKey FROM LogisticProduct WHERE CustomerKey = " & gnCustomerKey & " AND ProductCode = '" & sProductID & "'"
        GetProductKeyFromProductID = ExecuteQueryToDataTable(sSQL).Rows(0).Item(0)
    End Function
    
    Protected Function GetProductCodeFromProductKey(ByVal nProductKey As Int32) As String
        GetProductCodeFromProductKey = ExecuteQueryToDataTable("SELECT ProductCode FROM LogisticProduct WHERE LogisticProductKey = " & nProductKey).Rows(0).Item(0)
    End Function
    
    Protected Function GetProductNameFromProductCode(sProductCode As String) As String
        GetProductNameFromProductCode = String.Empty
        Dim sSQL As String = "SELECT ISNULL(ProductDepartmentId, '(no product name shown for COSTA orders)') FROM LogisticProduct WHERE CustomerKey = " & gnCustomerKey & " AND ProductCode = '" & sProductCode & "'"
        Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
        If dt.Rows.Count = 1 Then
            GetProductNameFromProductCode = dt.Rows(0).Item(0)
        End If
    End Function
    
    Protected Function GetProductDescriptionFromProductCode(sProductCode As String) As String
        GetProductDescriptionFromProductCode = String.Empty
        Dim sSQL As String = "SELECT TOP 1 ProductDescription FROM LogisticProduct WHERE CustomerKey = " & gnCustomerKey & " AND ProductCode = '" & sProductCode & "'"
        Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
        If dt.Rows.Count = 1 Then
            GetProductDescriptionFromProductCode = dt.Rows(0).Item(0)
        End If
    End Function
    
    Protected Function GetAgentDetailsFromAgentID_FININT(ByVal sAgentID As String) As DataRow
        GetAgentDetailsFromAgentID_FININT = Nothing
        Dim sSQL As String = "SELECT * FROM ClientData_WU_LegacyNetwork WHERE AgentID = " & sAgentID
        Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
        If dt.Rows.Count <> 1 Then
            WebMsgBox.Show("GetAgentDetailsFromAgentID_FININT: error retrieving agent details for " & sAgentID)
        Else
            GetAgentDetailsFromAgentID_FININT = dt.Rows(0)
        End If
    End Function
    
    Protected Function GetAgentDetailsFromAgentID_COSTA(ByVal sTermID As String) As DataRow
        GetAgentDetailsFromAgentID_COSTA = Nothing
        Dim sSQL As String = "SELECT * FROM ClientData_WUCOSTA_Agents WHERE TermID = '" & sTermID & "'"
        Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
        If dt.Rows.Count <> 1 Then
            WebMsgBox.Show("GetAgentDetailsFromAgentID_COSTA: error retrieving agent details for " & sTermID)
        Else
            GetAgentDetailsFromAgentID_COSTA = dt.Rows(0)
        End If
    End Function
    
    Protected Sub btnGenerateOrderSpreadsheet_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If rbFinint.Checked Then
            GenerateOrderSpreadsheetFININT()
        Else
            GenerateOrderSpreadsheetCOSTA()
        End If
    End Sub

    Protected Sub GenerateOrderSpreadsheetFININT()
        If Not GetConflatedOrderTableOneProductPerRow() Then
            WebMsgBox.Show("One or more Product IDs cannot be mapped to a Product Code.")
        Else
            Response.Clear()
            Response.ContentType = "text/csv"
            Dim sResponseValue As New StringBuilder
            sResponseValue.Append("attachment; filename=""")
            sResponseValue.Append("FININTOrder")
            sResponseValue.Append(Format(Now(), "yyyyMMMdd_hhmmssff"))
            sResponseValue.Append(".csv")
            sResponseValue.Append("""")
            Response.AddHeader("Content-Disposition", sResponseValue.ToString)
            Response.Write("Legacy Agent ID, Account Number, First Name / Legal Name, Last Name / Location, Addr 1, Addr 2, Town/City, County, Post Code, Product Code, Product Name, Product Description, Qty" & vbCrLf)
            For Each dr As DataRow In gdtConflatedOrderData.Rows
                Response.Write(dr("AgentID"))
                Response.Write(",")
                Dim drAgentDetails As DataRow = GetAgentDetailsFromAgentID_FININT(dr("AgentID"))
                Response.Write(drAgentDetails("AccountNumber"))
                Response.Write(",")
                Response.Write(drAgentDetails("LegalName"))
                Response.Write(",")
                Response.Write(drAgentDetails("LocationName"))
                Response.Write(",")
                Response.Write(drAgentDetails("AddressLine1"))
                Response.Write(",")
                Response.Write(drAgentDetails("AddressLine2"))
                Response.Write(",")
                Response.Write(drAgentDetails("CityName"))
                Response.Write(",")
                Response.Write(drAgentDetails("Province/County/State"))
                Response.Write(",")
                Response.Write(drAgentDetails("PostalCode"))
                Response.Write(",")
                Response.Write(dr("ProductID"))
                Response.Write(",")
                Response.Write(GetProductNameFromProductCode(dr("ProductID")))
                Response.Write(",")
                Response.Write(GetProductDescriptionFromProductCode(dr("ProductID")))
                Response.Write(",")
                Response.Write(dr("Qty"))
                Response.Write(vbCrLf)
            Next
            Response.End()
        End If
    End Sub

    Protected Sub GenerateOrderSpreadsheetCOSTA()
        If Not GetConflatedOrderTableOneProductPerRow() Then
            WebMsgBox.Show("One or more Product IDs cannot be mapped to a Product Code.")
        Else
            Response.Clear()
            Response.ContentType = "text/csv"
            Dim sResponseValue As New StringBuilder
            sResponseValue.Append("attachment; filename=""")
            sResponseValue.Append("COSTAOrder")
            sResponseValue.Append(Format(Now(), "yyyyMMMdd_hhmmssff"))
            sResponseValue.Append(".csv")
            sResponseValue.Append("""")
            Response.AddHeader("Content-Disposition", sResponseValue.ToString)
            Response.Write("Terminal ID, Class, Location Name, Address, Post Code, Area Code, Phone, Product Code, Product Description, Qty" & vbCrLf)
            For Each dr As DataRow In gdtConflatedOrderData.Rows
                Response.Write(dr("AgentID"))
                Response.Write(",")
                Dim drAgentDetails As DataRow = GetAgentDetailsFromAgentID_COSTA(dr("AgentID"))
                Response.Write(drAgentDetails("Class"))
                Response.Write(",")
                Response.Write(drAgentDetails("LocationName"))
                Response.Write(",")
                Response.Write(drAgentDetails("Address"))
                Response.Write(",")
                Response.Write(drAgentDetails("PostCode"))
                Response.Write(",")
                Response.Write(drAgentDetails("AreaCode"))
                Response.Write(",")
                Response.Write(drAgentDetails("Phone"))
                Response.Write(",")
                Response.Write(dr("ProductID"))
                Response.Write(",")
                Response.Write(GetProductDescriptionFromProductCode(dr("ProductID")))
                Response.Write(",")
                Response.Write(dr("Qty"))
                Response.Write(vbCrLf)
            Next
            Response.End()
        End If
    End Sub

    Protected Sub btnGenerateOrders_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim bSuccess As Boolean = GetConflatedOrderTableOneProductPerRow()
        If Not bSuccess Then
            WebMsgBox.Show("One or more Product IDs cannot be mapped to a Product Code.")
        Else
            Call MarshallOrders()
        End If
    End Sub

    Protected Function GetConflatedOrderTableOneProductPerRow() As Boolean
        Dim bError As Boolean = False
        Dim dictTotalProducts As New Dictionary(Of Int32, Int32)

        Call BuildOrderTable() ' builds 5-column Datatable gdtOrderData: OrderID, DateTime, AgentID, ProductID, Qty 
        Call FillOrderTable() ' fills 5-column Datatable gdtOrderData: OrderID, DateTime, AgentID, ProductID, Qty from .CSV

        Call BuildConflatedOrderTableOneProductPerRow()  ' builds 8-column Datatable gdtConflatedOrderData: OrderID, DateTime, AgentID, AgentName, ProductID, ProductName, ProductDescription, Qty
        
        Dim oDataView As New DataView(gdtOrderData)
        oDataView.Sort = "AgentID,ProductID"
        Dim sOrderID As String = String.Empty
        Dim sDateTime As String = String.Empty
        Dim sAgentID As String = String.Empty
        Dim sAgentName As String = String.Empty
        Dim sProductID As String = String.Empty
        Dim sProductName As String = String.Empty
        Dim sProductDescription As String = String.Empty
        Dim nProductKey As Int32 = 0
        
        Dim sQty As String = String.Empty
        Dim drConflatedOrderData As DataRow
        For Each drv As DataRowView In oDataView
            Dim dr As DataRow = drv.Row
            If sAgentID <> dr("AgentID").ToString Or sProductID <> dr("ProductID").ToString Then
                If sAgentID <> String.Empty Then
                    drConflatedOrderData = gdtConflatedOrderData.NewRow()
                    drConflatedOrderData("OrderID") = sOrderID
                    drConflatedOrderData("DateTime") = sDateTime
                    drConflatedOrderData("AgentID") = sAgentID
                    drConflatedOrderData("AgentName") = sAgentName
                    drConflatedOrderData("ProductID") = sProductID
                    drConflatedOrderData("ProductName") = sProductName
                    drConflatedOrderData("ProductDescription") = sProductDescription
                    drConflatedOrderData("Qty") = sQty
                    gdtConflatedOrderData.Rows.Add(drConflatedOrderData)
                    nProductKey = GetProductKeyFromProductID(sProductID)
                    If nProductKey = 0 Then
                        bError = True
                    Else
                        Call AddProductToTotal(nProductKey, sQty, dictTotalProducts)
                    End If
                End If
                sOrderID = dr("OrderID")
                sDateTime = dr("DateTime")
                sAgentID = dr("AgentID")
                sAgentName = dr("AgentName")
                sProductID = dr("ProductID")
                sProductName = dr("ProductName")
                sProductDescription = dr("ProductDescription")
                sQty = dr("Qty")
            Else
                If sAgentID <> String.Empty Then
                    sQty += dr("Qty")
                End If
            End If
        Next
        drConflatedOrderData = gdtConflatedOrderData.NewRow()
        drConflatedOrderData("OrderID") = sOrderID
        drConflatedOrderData("DateTime") = sDateTime
        drConflatedOrderData("AgentID") = sAgentID
        drConflatedOrderData("AgentName") = sAgentName
        drConflatedOrderData("ProductID") = sProductID
        drConflatedOrderData("ProductName") = sProductName
        drConflatedOrderData("ProductDescription") = sProductDescription
        drConflatedOrderData("Qty") = sQty
        gdtConflatedOrderData.Rows.Add(drConflatedOrderData)
        nProductKey = GetProductKeyFromProductID(sProductID)
        If nProductKey = 0 Then
            bError = True
        Else
            Call AddProductToTotal(nProductKey, sQty, dictTotalProducts)
        End If
        Dim sMessage As String = String.Empty
        If Not bError Then
            Dim bUnavailable As Boolean = False
            For Each kv As KeyValuePair(Of Int32, Int32) In dictTotalProducts
                Dim nAvailableQuantity As Int32 = GetAvailableQty(kv.Key)
                If kv.Value < nAvailableQuantity Then
                    sMessage &= "Available quantity of product " & GetProductCodeFromProductKey(kv.Key) & " - " & GetProductDescriptionFromProductCode(GetProductCodeFromProductKey(kv.Key)) & " (" & nAvailableQuantity & ") is sufficient to fulfil all orders (required: " & kv.Value & ")" & "\r\r"
                Else
                    sMessage &= "WARNING: available quantity of product " & GetProductCodeFromProductKey(kv.Key) & " - " & GetProductDescriptionFromProductCode(GetProductCodeFromProductKey(kv.Key)) & " (" & nAvailableQuantity & ") is insufficient to fulfil all orders (required: " & kv.Value & ")" & "\r\r"
                    bUnavailable = True
                End If
            Next
            If bUnavailable Then
                sMessage &= "SUMMARY - ONE OR MORE PRODUCTS HAS INSUFFICIENT QUANTITY TO FULFIL ALL ORDERS. Please adjust the order file and resubmit!!!"
            Else
                sMessage &= "There is sufficient product quantity to fulfil all orders."
            End If
        Else
            sMessage = "Product lookup error - cannot check quantities."
        End If
        WebMsgBox.Show(sMessage)
        GetConflatedOrderTableOneProductPerRow = Not bError
    End Function

    Private Function GetAvailableQty(ByVal nLogisticProductKey As Integer) As Integer
        Dim sSQL As String = "SELECT Quantity = CASE ISNUMERIC((SELECT SUM(LogisticProductQuantity) FROM LogisticProductLocation WHERE LogisticProductKey = " & nLogisticProductKey & ")) WHEN 0 THEN 0 ELSE (SELECT SUM(LogisticProductQuantity) FROM LogisticProductLocation WHERE LogisticProductKey = " & nLogisticProductKey & ") END"
        Dim oDataTable As DataTable = ExecuteQueryToDataTable(sSQL)
        If oDataTable IsNot Nothing AndAlso oDataTable.Rows.Count <> 0 Then
            GetAvailableQty = oDataTable.Rows(0)(0)
        Else
            GetAvailableQty = 0
        End If
    End Function

    Protected Sub BuildOrderTable() ' builds a 5 column Datatable gdtOrderData: OrderID, DateTime, AgentID, ProductID, Qty & fills from .CSV
        gdtOrderData = New DataTable
        gdtOrderData.Columns.Add(New DataColumn("OrderID", GetType(Int32)))
        gdtOrderData.Columns.Add(New DataColumn("DateTime", GetType(String)))
        gdtOrderData.Columns.Add(New DataColumn("AgentID", GetType(String)))
        gdtOrderData.Columns.Add(New DataColumn("AgentName", GetType(String)))  ' NEW
        gdtOrderData.Columns.Add(New DataColumn("ProductID", GetType(String)))
        gdtOrderData.Columns.Add(New DataColumn("ProductName", GetType(String)))  ' NEW
        gdtOrderData.Columns.Add(New DataColumn("ProductDescription", GetType(String)))  ' NEW
        gdtOrderData.Columns.Add(New DataColumn("Qty", GetType(Int32)))
    End Sub
    
    Protected Sub FillOrderTable() ' fills 5 column Datatable gdtOrderData: OrderID, DateTime, AgentID, ProductID, Qty from .CSV
        Dim dr As DataRow
        Dim sr As New StreamReader(psUniqueFilename)
        Dim sLine As String = String.Empty
        Dim sLineElements() As String
        Dim bReadPastHeader As Boolean = False

        Do While sr.Peek >= 0
            'If Not bReadPastHeader Then
            '    sLine = sr.ReadLine()
            '    bReadPastHeader = True
            'End If
            sLine = sr.ReadLine()
            sLineElements = sLine.Split(",")
            If Not (sLineElements(0) = String.Empty And sLineElements(1) = String.Empty And sLineElements(2) = String.Empty And sLineElements(3) = String.Empty And sLineElements(4) = String.Empty) Then
                dr = gdtOrderData.NewRow()
                dr("OrderID") = sLineElements(0)
                dr("DateTime") = sLineElements(1)
                dr("AgentID") = sLineElements(2)
                dr("ProductID") = sLineElements(3)
                dr("Qty") = sLineElements(4)
                Dim drAgentDetails As DataRow
                If rbFinint.Checked Then
                    drAgentDetails = GetAgentDetailsFromAgentID_FININT(dr("AgentID"))
                    dr("AgentName") = drAgentDetails("LocationName")
                Else
                    drAgentDetails = GetAgentDetailsFromAgentID_COSTA(dr("AgentID"))
                    dr("AgentName") = drAgentDetails("LocationName")
                End If
                dr("ProductName") = GetProductNameFromProductCode(dr("ProductID"))
                dr("ProductDescription") = GetProductDescriptionFromProductCode(dr("ProductID"))
                gdtOrderData.Rows.Add(dr)
            End If
        Loop
        sr.Close()
    End Sub

    Protected Sub BuildConflatedOrderTableOneProductPerRow()
        gdtConflatedOrderData = New DataTable
        gdtConflatedOrderData.Columns.Add(New DataColumn("OrderID", GetType(Int32)))
        gdtConflatedOrderData.Columns.Add(New DataColumn("DateTime", GetType(String)))
        gdtConflatedOrderData.Columns.Add(New DataColumn("AgentID", GetType(String)))
        gdtConflatedOrderData.Columns.Add(New DataColumn("AgentName", GetType(String)))
        gdtConflatedOrderData.Columns.Add(New DataColumn("ProductID", GetType(String)))
        gdtConflatedOrderData.Columns.Add(New DataColumn("ProductName", GetType(String)))
        gdtConflatedOrderData.Columns.Add(New DataColumn("ProductDescription", GetType(String)))
        gdtConflatedOrderData.Columns.Add(New DataColumn("Qty", GetType(Int32)))
    End Sub

    Protected Sub MarshallOrders()
        If gdtConflatedOrderData.Rows.Count > 0 Then
            Dim lstConsignmentNumbers As New List(Of Int32)
            Dim dictOrderItems As New Dictionary(Of Int32, Int32)
            Dim lstOrderNumbers As New List(Of String)
            Dim sAgentID As String = gdtConflatedOrderData.Rows(0).Item("AgentID")
            For i As Int32 = 0 To gdtConflatedOrderData.Rows.Count - 1
                If gdtConflatedOrderData.Rows(i).Item("AgentID") <> sAgentID Then
                    If rbFinint.Checked Then
                        lstConsignmentNumbers.Add(nSubmitConsignment(GetAgentDetailsFromAgentID_FININT(sAgentID), dictOrderItems, lstOrderNumbers))
                    Else
                        lstConsignmentNumbers.Add(nSubmitConsignment(GetAgentDetailsFromAgentID_COSTA(sAgentID), dictOrderItems, lstOrderNumbers))
                    End If
                    sAgentID = gdtConflatedOrderData.Rows(i).Item("AgentID")
                    dictOrderItems.Clear()
                    dictOrderItems.Add(GetProductKeyFromProductID(gdtConflatedOrderData.Rows(i).Item("ProductID")), gdtConflatedOrderData.Rows(i).Item("Qty"))
                    lstOrderNumbers.Clear()
                    lstOrderNumbers.Add(gdtConflatedOrderData.Rows(i).Item("OrderID"))
                Else
                    dictOrderItems.Add(GetProductKeyFromProductID(gdtConflatedOrderData.Rows(i).Item("ProductID")), gdtConflatedOrderData.Rows(i).Item("Qty"))
                    lstOrderNumbers.Add(gdtConflatedOrderData.Rows(i).Item("OrderID"))
                End If
            Next
            If rbFinint.Checked Then
                lstConsignmentNumbers.Add(nSubmitConsignment(GetAgentDetailsFromAgentID_FININT(sAgentID), dictOrderItems, lstOrderNumbers))
            Else
                lstConsignmentNumbers.Add(nSubmitConsignment(GetAgentDetailsFromAgentID_COSTA(sAgentID), dictOrderItems, lstOrderNumbers))
            End If
            lblResult.Text = "Consignments generated: "
            Dim bFailedConsignment As Boolean = False
            For Each sConsignmentNumber As String In lstConsignmentNumbers
                If CInt(sConsignmentNumber) <= 0 Then
                    bFailedConsignment = True
                End If
                lblResult.Text &= sConsignmentNumber & " - "
            Next
            lblResult.Text = lblResult.Text.Substring(0, lblResult.Text.Length - 2)
            If bFailedConsignment Then
                lblResult.Text &= " WARNING: ONE OR MORE ORDERS FAILED"
            End If
        Else
            lblResult.Text = "No consignments to generate."
        End If
    End Sub
    
    Protected Function nSubmitConsignment(ByRef drAgentDetails As DataRow, ByVal dictOrderItems As Dictionary(Of Int32, Int32), ByVal lstOrderNumbers As List(Of String)) As Integer
        Dim sConn As String = ConfigLib.GetConfigItem_ConnectionString
        Dim lBookingKey As Long
        Dim lConsignmentKey As Long
        Dim BookingFailed As Boolean
        Dim oConn As New SqlConnection(sConn)
        Dim oTrans As SqlTransaction
        Dim oCmdAddBooking As SqlCommand = New SqlCommand("spASPNET_StockBooking_Add3", oConn)
        nSubmitConsignment = 0
        oCmdAddBooking.CommandType = CommandType.StoredProcedure
        Dim dtGenericUser As DataTable = ExecuteQueryToDataTable("SELECT [key] 'UserKey' FROM UserProfile WHERE UserID = 'WUFINgu'")
        
        Dim nGenericUserKey As Int32 = 0
        
        Dim param1 As SqlParameter = New SqlParameter("@UserProfileKey", SqlDbType.Int, 4)
        param1.Value = nGenericUserKey
        oCmdAddBooking.Parameters.Add(param1)
        Dim param2 As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
        param2.Value = gnCustomerKey
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
        For Each s As String In lstOrderNumbers
            param6.Value &= s & " "
        Next
        param6.Value = param6.Value.ToString.Trim
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
        
        Dim sSQL As String = "SELECT * FROM Customer WHERE CustomerKey = " & gnCustomerKey
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
            WebMsgBox.Show("Couldn't retrieve Consignor details.")
            nSubmitConsignment = -1
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
        param25.Value = drAgentDetails("LocationName").ToString.Trim
        oCmdAddBooking.Parameters.Add(param25)
        Dim param26 As SqlParameter = New SqlParameter("@CneeAddr1", SqlDbType.NVarChar, 50)
        If rbFinint.Checked Then
            param26.Value = drAgentDetails("AddressLine1").ToString.Trim
        Else
            param26.Value = drAgentDetails("Address").ToString.Trim
        End If
        oCmdAddBooking.Parameters.Add(param26)
        Dim param27 As SqlParameter = New SqlParameter("@CneeAddr2", SqlDbType.NVarChar, 50)
        If rbFinint.Checked Then
            param27.Value = drAgentDetails("AddressLine2").ToString.Trim
        Else
            param27.Value = String.Empty
        End If
        oCmdAddBooking.Parameters.Add(param27)
        Dim param28 As SqlParameter = New SqlParameter("@CneeAddr3", SqlDbType.NVarChar, 50)
        param28.Value = ""
        oCmdAddBooking.Parameters.Add(param28)
        Dim param29 As SqlParameter = New SqlParameter("@CneeTown", SqlDbType.NVarChar, 50)
        If rbFinint.Checked Then
            param29.Value = drAgentDetails("CityName").ToString.Trim
        Else
            param29.Value = String.Empty
        End If
        oCmdAddBooking.Parameters.Add(param29)
        Dim param30 As SqlParameter = New SqlParameter("@CneeState", SqlDbType.NVarChar, 50)
        If rbFinint.Checked Then
            param30.Value = drAgentDetails("Province/County/State").ToString.Trim
        Else
            param30.Value = String.Empty
        End If
        If rbFinint.Checked Then
            param30.Value = drAgentDetails("Province/County/State").ToString.Trim
        Else
            param30.Value = String.Empty
        End If
        oCmdAddBooking.Parameters.Add(param30)
        Dim param31 As SqlParameter = New SqlParameter("@CneePostCode", SqlDbType.NVarChar, 50)
        If rbFinint.Checked Then
            param31.Value = drAgentDetails("PostalCode").ToString.Trim
        Else
            param31.Value = drAgentDetails("PostCode").ToString.Trim
        End If
        oCmdAddBooking.Parameters.Add(param31)
        Dim param32 As SqlParameter = New SqlParameter("@CneeCountryKey", SqlDbType.Int, 4)
        param32.Value = 222
        oCmdAddBooking.Parameters.Add(param32)
        Dim param33 As SqlParameter = New SqlParameter("@CneeCtcName", SqlDbType.NVarChar, 50)
        If rbFinint.Checked Then
            param33.Value = drAgentDetails("LegalName").ToString.Trim
        Else
            param33.Value = String.Empty
        End If
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
                For Each kvp As KeyValuePair(Of Integer, Integer) In dictOrderItems
                    Dim lProductKey As Long = kvp.Key
                    Dim lPickQuantity As Long = kvp.Value
                    Dim oCmdAddStockItem As SqlCommand = New SqlCommand("spASPNET_LogisticMovement_Add", oConn)
                    oCmdAddStockItem.CommandType = CommandType.StoredProcedure
                    Dim param51 As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int, 4)
                    param51.Value = nGenericUserKey
                    oCmdAddStockItem.Parameters.Add(param51)
                    Dim param52 As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
                    param52.Value = gnCustomerKey
                    oCmdAddStockItem.Parameters.Add(param52)
                    Dim param53 As SqlParameter = New SqlParameter("@LogisticBookingKey", SqlDbType.Int, 4)
                    param53.Value = lBookingKey
                    oCmdAddStockItem.Parameters.Add(param53)
                    Dim param54 As SqlParameter = New SqlParameter("@LogisticProductKey", SqlDbType.Int, 4)
                    param54.Value = lProductKey
                    oCmdAddStockItem.Parameters.Add(param54)
                    Dim param55 As SqlParameter = New SqlParameter("@LogisticMovementStateId", SqlDbType.NVarChar, 20)
                    param55.Value = "PENDING"
                    oCmdAddStockItem.Parameters.Add(param55)
                    Dim param56 As SqlParameter = New SqlParameter("@ItemsOut", SqlDbType.Int, 4)
                    param56.Value = lPickQuantity
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

    Protected Sub lnkbtnViewOrderFile_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If psUniqueFilename = String.Empty Then
            WebMsgBox.Show("Read the order file first.")
            Exit Sub
        End If

        If lnkbtnViewOrderFile.Text.Contains("view") Then
            tbOrderFile.Visible = True
            lnkbtnViewOrderFile.Text = "hide order file"
            tbOrderFile.Text = ""
            
            Dim sr As New StreamReader(psUniqueFilename)
            Dim nLineCount As Integer = 0
            Do While sr.Peek >= 0
                tbOrderFile.Text &= sr.ReadLine() & Environment.NewLine
                nLineCount += 1
            Loop
            sr.Close()
        Else
            tbOrderFile.Visible = False
            lnkbtnViewOrderFile.Text = "view order file"
        End If
    End Sub

    Protected Function ExecuteQueryToDataTable(ByVal sQuery As String) As DataTable
        Dim oDataTable As New DataTable
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter(sQuery, oConn)
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

    Property psUniqueFilename() As String
        Get
            Dim o As Object = ViewState("FINIT_UniqueFilename")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("FINIT_UniqueFilename") = Value
        End Set
    End Property

</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>COSTA / FININT Order Processor</title>
</head>
<body>
    <form id="form1" runat="server" defaultfocus="FormUpload1" enctype="multipart/form-data">
    <main:Header ID="ctlHeader" runat="server" />
    <asp:PlaceHolder ID="PlaceHolderForScriptManager" runat="server" />
    <div style="font-size: small; font-family: Verdana">
        <strong>COSTA / FININT ORDER PROCESSOR</strong> - version 23DEC13<br />
        <br />
        &nbsp;<b>1.</b>&nbsp;&nbsp;&nbsp; Select COSTA<asp:RadioButton ID="rbCosta" runat="server"
            GroupName="customer" />
        &nbsp;or FININT
        <asp:RadioButton ID="rbFinint" runat="server" GroupName="customer" />
        <br />
        <br />
        &nbsp;<strong>2.</strong>&nbsp; &nbsp; Select order&nbsp;CSV&nbsp;file:
        <asp:FileUpload ID="FileUpload1" runat="server" Font-Names="Verdana" Font-Size="X-Small"
            Font-Bold="true" Width="300px" />
        <br />
        <br />
        &nbsp;<b>3</b><strong>.</strong>&nbsp;&nbsp;&nbsp;&nbsp;
        <asp:Button ID="btnReadOrders" runat="server" OnClick="btnReadOrders_Click" Text="Read Orders"
            Width="140px" />
        &nbsp;&nbsp;<asp:CheckBox ID="cbColumnHeadingsInRow1" runat="server" Text="Column&nbsp;headings&nbsp;in&nbsp;row&nbsp;1"
            Font-Size="XX-Small" />
        &nbsp;
        <asp:LinkButton ID="lnkbtnViewOrderFile" runat="server" OnClick="lnkbtnViewOrderFile_Click">view order file</asp:LinkButton>
        <br />
        <br />
        &nbsp;<asp:TextBox ID="tbOrderFile" runat="server" ReadOnly="True" TextMode="MultiLine"
            Visible="False" Width="500px" Rows="5" />
        <br />
&nbsp;<asp:Label ID="lblHeaderWarning" runat="server" Font-Bold="True" 
            ForeColor="Red" Visible="true" />
        &nbsp;<asp:Label ID="lblMessage" runat="server" Font-Bold="True" />
        <br />
        &nbsp;
        <br />
        <br />
        <asp:GridView ID="gvOrders" runat="server" CellPadding="2" Font-Names="Verdana" Font-Size="XX-Small"
            Width="98%" AutoGenerateColumns="False">
            <Columns>
                <asp:BoundField DataField="AgentName" HeaderText="Agent Name" ReadOnly="True" SortExpression="AgentName" />
                <asp:BoundField DataField="ProductID" HeaderText="Product Code" ReadOnly="True" SortExpression="ProductID" />
                <asp:BoundField DataField="ProductName" HeaderText="Product Name" ReadOnly="True"
                    SortExpression="ProductName" />
                <asp:BoundField DataField="ProductDescription" HeaderText="Description" ReadOnly="True"
                    SortExpression="ProductDescription" />
                <asp:BoundField DataField="Qty" HeaderText="Qty" ReadOnly="True" SortExpression="Qty" />
            </Columns>
        </asp:GridView>
        <br />
        &nbsp;<strong>4.</strong>&nbsp;&nbsp;&nbsp;&nbsp;
        <asp:Button ID="btnGenerateOrderSpreadsheet" runat="server" Text="Generate Order Spreadsheet"
            Enabled="False" OnClick="btnGenerateOrderSpreadsheet_Click" />
        &nbsp;&nbsp;<asp:Button ID="btnGenerateOrders" runat="server" Text="Generate Orders"
            Enabled="False" OnClick="btnGenerateOrders_Click" Width="200px" />
        <br />
        <br />
        <asp:Label ID="lblResult" runat="server" Font-Bold="True" />
        <br />
        <asp:Label ID="lblError" runat="server" ForeColor="Red" Font-Bold="True" />
        <br />
        <strong>INSTRUCTIONS:</strong>
        <br />
        <br />
        &nbsp;1. Select the option button to indicate the type of file you are processing
        - COSTA or FININT.<br />
        <br />
        &nbsp;2. Click the <strong>Browse</strong> button, navigate to and select the comma-separated
        variable (CSV) file containing the COSTA or FININT orders.<br />
        <br />
        &nbsp;3. Click the <strong>Read Orders</strong> button. The data is read and validated.&nbsp;
        Each order line in the uploaded file is shown (a single order may comprise several
        order lines).<br />
        <br />
        - Check no error messages are displayed.<br />
        - Check the summary of quantities required and available that is displayed to ensure
        enough quantity of all products is available.<br />
        <br />
        &nbsp;4. To view the contents of the order file, click the <strong>Generate Order Spreadsheet</strong>
        button.<br />
        <br />
        &nbsp;5. To generate and place the orders, click the <strong>Generate Orders</strong>
        button.<br />
        <br />
        <strong>&nbsp;NOTES:</strong>
        <br />
        <br />
        &nbsp;1. When the order is placed, two or more orders placed for the same product
        by the same agent will be combined into a single order for that product.<br />
        &nbsp;2. When the order is placed,&nbsp; multiple orders placed by the same agent
        are combined into a single order.<br />
        &nbsp;3. The order spreadsheet shows one data as it appears in the COSTA or FININT
        file, ie one row per product. That is, if an agent orders 3 products, that agent
        will have 3 consecutive rows in the spreadsheet.<br />
        &nbsp;4.&nbsp;The original order number(s) in the incoming COSTA or FININT order
        file are saved in field Cust Ref 4 of the generated order.<br />
        <br />
    </div>
    </form>
</body>
</html>
