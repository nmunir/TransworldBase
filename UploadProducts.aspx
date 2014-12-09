<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<script runat="server">
  
    ' check if all selected fields are different (no dupes); show number of addresses read; show number of addresses written
    ' later on, split large files into chunks and process sequentially; offer to show 1st line as prompt for column headers on line 1
    ' show if required field is not select in message box instead of erroring every missing instance
    
    ' when checking data, check authoriser specified if any authorisable product specified
  
    Const MARSHALL_PRODUCT_CODE As Integer = 0
    Const MARSHALL_PRODUCT_DATE As Integer = 1
    Const MARSHALL_DESCRIPTION As Integer = 2
    Const MARSHALL_COST_CENTRE_DEPT_ID As Integer = 3
    Const MARSHALL_CATEGORY As Integer = 4
    Const MARSHALL_SUB_CATEGORY As Integer = 5
    Const MARSHALL_SUB_CATEGORY_2 As Integer = 6
    Const MARSHALL_MIN_STOCK_LEVEL As Integer = 7
    Const MARSHALL_UNIT_VALUE As Integer = 8
    Const MARSHALL_SELLING_PRICE As Integer = 9
    Const MARSHALL_UNIT_WEIGHT_GRAMS As Integer = 10
    Const MARSHALL_LANGUAGE As Integer = 11
    Const MARSHALL_ITEMS_PER_BOX As Integer = 12
    Const MARSHALL_EXPIRY_DATE As Integer = 13
    Const MARSHALL_REPLENISHMENT_DATE As Integer = 14
    Const MARSHALL_INACTIVITY_ALERT_DAYS As Integer = 15
    Const MARSHALL_NOTES As Integer = 16
    Const MARSHALL_MISC_1 As Integer = 17
    Const MARSHALL_MISC_2 As Integer = 18
    Const MARSHALL_AUTHORISABLE As Integer = 19
    Const MARSHALL_CALENDAR_MANAGED As Integer = 20
    Const MARSHALL_MAX_GRAB As Integer = 21
    Const MARSHALL_ARCHIVED As Integer = 22
    Const MARSHALL_PRODUCTQUANTITY As Integer = 23

    Dim sFileName As String, sIntermediateFileName As String
    Dim sFilePrefix As String, sFileSuffix As String = ".csv"
    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString
  
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsPostBack Then
            Call Initialise()
        End If
        'tbBayKey.Text = tbBayKey.Text.Trim
        'If tbBayKey.Text <> String.Empty Then
        '    If Not IsNumeric(tbBayKey.Text) Then
        '        WebMsgBox.Show("Non numeric bay key")
        '    Else
        '        lblBayName.Text = BayName()
        '        If lblBayName.Text = String.Empty Then
        '            lblBayName.Text = "NO BAY MATCHED!!"
        '            lblBayName.ForeColor = Drawing.Color.Red
        '        Else
        '            lblBayName.ForeColor = Drawing.Color.Black
        '        End If
                
        '    End If
        'End If
    End Sub
  
    Protected Sub Initialise()
        pnlStart.Visible = True
        pnlGridView.Visible = False
        pnlMessage.Visible = False
        pnlHelp.Visible = False
        Call GetCustomerAccountCodes()
    End Sub
    
    Protected Function bSetExplicitProductPermissionsFlag() As Boolean
        Dim oDataTable As New DataTable()
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter("spASPNET_Customer_ExplicitProductPermissions_GetFlag", oConn)

        Try
            oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
            oAdapter.SelectCommand.Parameters("@CustomerKey").Value = pnSelectedCustomerKey
            oAdapter.Fill(oDataTable)
        Catch ex As SqlException
            WebMsgBox.Show("bSetExplicitProductPermissionsFlag: " & ex.Message)
            bSetExplicitProductPermissionsFlag = False
        Finally
            oConn.Close()
            If IsDBNull(oDataTable.Rows(0).Item(0)) Then
                bSetExplicitProductPermissionsFlag = False
            Else
                bSetExplicitProductPermissionsFlag = oDataTable.Rows(0).Item(0)
            End If
        End Try
    End Function
 
    Protected Sub btnReadProducts_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ReadProducts()
    End Sub
  
    Protected Sub ReadProducts()
        Dim sFileName As String
        Dim nLineCount As Integer
        lblError.Text = ""
        sFileName = FileUpload1.FileName
        If My.Computer.FileSystem.FileExists(sFileName) Then
            Call OutputErrorMessage("Could not find file " & sFileName)
        End If
        If FileUpload1.HasFile Then
            sFilePrefix = Format(Now(), "yyyymmddhhmmssff")
            sFileName = Server.MapPath("") & "\" & sFilePrefix & sFileSuffix
            FileUpload1.SaveAs(sFileName)
            nLineCount = nCSVLineCount(sFileName)
            If nLineCount > 10000 Then
                WebMsgBox.Show("WARNING: The file you are uploading contains " & nLineCount.ToString & " products.  To ensure successful uploading we recommend that you split the file into multiple files of no more than 10,000 products per file.")
            Else
                sIntermediateFileName = Server.MapPath("") & "\" & sFilePrefix & "-2" & sFileSuffix
                Call RemoveEmbeddedLineBreaksAndCommas(sFileName, sIntermediateFileName)
                Dim dt As DataTable
                dt = DelimitFile(sIntermediateFileName, ",", cbColumnHeadingsInRow1.Checked)
                If dt.Rows.Count > 0 Then
                    Dim r As DataRow = dt.NewRow()
                    dt.Rows.InsertAt(r, 0)
                    gvUploadData.DataSource = dt
                    gvUploadData.DataBind()
                End If
                If gvUploadData.Rows.Count > 0 Then
                    pnlStart.Visible = False
                    pnlGridView.Visible = True
                Else
                    If lblError.Text = "" Then
                        Call OutputErrorMessage("No data found in file")
                    End If
                End If
            End If
        Else
            Call OutputErrorMessage("Specified file could not be found or file could not be processed")
            Exit Sub
        End If

        If My.Computer.FileSystem.FileExists(sFileName) Then
            My.Computer.FileSystem.DeleteFile(sFileName)
        End If
        If My.Computer.FileSystem.FileExists(sIntermediateFileName) Then
            My.Computer.FileSystem.DeleteFile(sIntermediateFileName)
        End If
    End Sub
  
    Protected Function nCSVLineCount(ByVal sFileName As String) As Integer
        Dim sr As New StreamReader(sFileName)
        Dim nLineCount As Integer = 0
        Do While sr.Peek >= 0
            sr.ReadLine()
            nLineCount += 1
        Loop
        sr.Close()
        nCSVLineCount = nLineCount
    End Function

    Private Sub RemoveEmbeddedLineBreaksAndCommas(ByVal sFileIn As String, ByVal sFileOut As String)
        Dim bWithinQuotes As Boolean = False
        Dim sr As New StreamReader(sFileIn)
        Dim sw As New StreamWriter(sFileOut, False)
        Dim sFileContents As New StringBuilder
        Dim nChar As Integer, nPrevChar As Integer
        Dim sChar As Char
        While sr.Peek <> -1
            nPrevChar = nChar
            nChar = sr.Read
            Try
                sChar = Chr(nChar)   ' failed on Novartis 'Guide Dogs – Sheffield' (65533) whereas 'Ivy Lodge Vet - Large Animal' worked okay
            Catch
                sChar = "?"
            End Try
            If sChar = """" Then
                If bWithinQuotes Then
                    If sr.Peek <> -1 AndAlso Chr(sr.Peek) = """" Then  ' if next char is a quote this is a quoted quote; map 1st char and consume 2nd; cope with this being final character in file, where peek would fail
                        nChar = sr.Read
                        sChar = "‰"
                    Else
                        bWithinQuotes = False
                    End If
                Else
                    bWithinQuotes = True
                End If
            End If
            If (sChar = vbCr Or sChar = vbLf) And bWithinQuotes Then
                sChar = " "
            End If
            If sChar = "," And bWithinQuotes Then
                sChar = "¡"
            End If
            sw.Write(sChar)
        End While
        sr.Close()
        sw.Close()
    End Sub
  
    Private Function DelimitFile(ByVal sFile As String, ByVal sDelimChar As Char, ByVal bHasRowHeader As Boolean) As DataTable
        Dim bHeaderFlag As Boolean = bHasRowHeader ' rowheaders present?
        Dim sr As New StreamReader(sFile) ' opens file for reading
        Dim dt As New DataTable ' to hold resulting data
        Dim bGetRowCount As Boolean = True ' flag to get first line column count
        Dim nElementCount As Integer, nRecordCount As Integer
        While sr.Peek <> -1
            Dim sLine As String = sr.ReadLine   ' get next line
            sLine = sLine.Replace("""", "")     ' remove all quotes, since quoted quotes have already been transformed
            Dim sElements = Split(sLine, ",")   ' split on delimiter
            If bGetRowCount = True Then         ' unknown # cols until first line is parsed so get the # columns in datatable and add them
                For i As Integer = 0 To sElements.GetLength(0) - 1  ' go through array
                    Dim dc As New DataColumn                        ' create a datacolumn
                    If bHeaderFlag = True Then                      ' use this row as column header text if RowHeader is True
                        dc.ColumnName = sElements(i)
                    End If
                    Try
                        dt.Columns.Add(dc)                          ' need this in a Try Catch in case of duplicate row header text
                    Catch
                        dc.ColumnName = ""
                        dt.Columns.Add(dc)
                    End Try
                Next
                nElementCount = sElements.GetLength(0)              ' save element count
                bGetRowCount = False
                'Dim PrevSplits(30) As String
                'For z As Integer = 0 To sElements.GetLength(0) - 1
                ' PrevSplits(z) = sElements(z)
                ' Next
            Else                                                    ' already have # cols; just check for bad rowcount
                Dim x As Integer = sElements.GetLength(0)
                If nElementCount <> sElements.GetLength(0) Then
                    If (sElements.GetLength(0) > 1) AndAlso (dt.Rows.Count > 0) Then
                        Call OutputErrorMessage("At row " & (dt.Rows.Count + 1) & " the number of items differed from previous row - cannot process file further")
                    End If
                    sr.Close()
                    lblMessage.Text = dt.Rows.Count.ToString & "processed"
                    Return dt
                End If
                'Dim PrevSplits(30) As String
                'For z As Integer = 0 To sElements.GetLength(0) - 1
                ' PrevSplits(z) = sElements(z)
                'Next
              
            End If
            If bHeaderFlag = False Then
                For x As Integer = 0 To sElements.GetLength(0) - 1      ' reconstitute chars & add row data to table
                    sElements(x) = sElements(x).Replace("¡", ",")
                    sElements(x) = sElements(x).Replace("‰", """")
                    'sElements(x) = sElements(x).Replace(sTempDelimiter, sDelimiter)
                Next
                dt.Rows.Add.ItemArray = sElements
                nRecordCount = nRecordCount + 1
            Else                                        ' this is the header row don't add it, set flag to false
                bHeaderFlag = False
            End If
        End While
        sr.Close()
        lblMessage.Text = dt.Rows.Count.ToString & "processed"
        Return dt
    End Function
  
    Protected Sub OutputErrorMessage(ByVal sMessage As String) ' ensures any existing message is not overwritten
        If lblError.Text = "" Then
            lblError.Text = sMessage
        End If
    End Sub
  
    Protected Sub btnCheckData_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call CheckData()
    End Sub
  
    Protected Sub CheckData()
        If SelectedFieldsUnique() Then
            lblMessage.Text = ""
            lblSuccess.Text = ""
            Call ProcessProducts(bAddToDatabase:=False)
            If lblError.Text = "" Then
                lblSuccess.Text = "No errors found"
            End If
        Else
            WebMsgBox.Show("Two or more columns are mapped to the same product field.")
        End If
    End Sub

    Protected Sub btnUploadProducts_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        tbResults.Text = String.Empty
        If SelectedFieldsUnique() Then
            lblMessage.Text = ""
            lblSuccess.Text = ""
            Call ProcessProducts(bAddToDatabase:=False)
            If lblError.Text = "" Then
                Call ProcessProducts(bAddToDatabase:=True)
            End If
            If lblError.Text = "" Then
                lblMessage.Text = gvUploadData.Rows.Count.ToString & " products added"
                pnlGridView.Visible = False
            End If
            If lblError.Text = "" Then
                lblSuccess.Text = "Products successfully loaded"
            End If
        Else
            WebMsgBox.Show("Two or more columns are mapped to the same field.")
        End If
    End Sub
  
    Protected Function SelectedFieldsUnique() As Boolean
        ' DO MSG BOX IN HERE; ALSO CHECK COUNTRY SELECTED; ALSO CHECK SOMETHING SELECTED
        Dim nCol As Integer
        Dim nMaxCols As Integer = gvUploadData.Rows(0).Cells.Count
        Dim nTargetIndex As Integer
        Dim nSelectedField(25) As Int16  ' dropdown fields: "", "", "ShortCode", "Addressee", "Addr1", "Addr2", "Addr3", "City", "State", "PostCode", "Country", "Attn", "Tel"
        SelectedFieldsUnique = True
        For nCol = 0 To nMaxCols - 1
            nTargetIndex = ExtractHiddenFieldValue(nCol)
            If nTargetIndex >= 2 Then     ' allow multiple 'nothing selected' or 'don't use this column'
                nSelectedField(nTargetIndex) += 1
                If nSelectedField(nTargetIndex) > 1 Then
                    SelectedFieldsUnique = False
                    Exit For
                End If
            End If
        Next
    End Function
  
    Protected Sub AddDropDowns()
        Dim gvr As GridViewRow
        gvr = gvUploadData.Rows(0)
        Dim j As Integer
        For j = 0 To gvr.Cells.Count - 1
            Dim dd As New DropDownList
            dd.ID = "Select" & j.ToString.Trim
            'dd.AutoPostBack = True
            ' following list must match the array used to match selection during processing          
            dd.Items.Add("- nothing selected -")
            dd.Items.Add("- don't use this column -")
            dd.Items.Add("Product Code")
            dd.Items.Add("Value / Date")
            dd.Items.Add("Description")
            dd.Items.Add("Cost Centre / Dept Id")
            dd.Items.Add("Category")
            dd.Items.Add("Sub Category")
            dd.Items.Add("Sub Category 2")
            dd.Items.Add("Min Stock Level")
            dd.Items.Add("Unit Value")
            dd.Items.Add("Selling Price")
            dd.Items.Add("Unit Weight (gm)")
            dd.Items.Add("Language")
            dd.Items.Add("Items / Box")
            dd.Items.Add("Expiry Date")
            dd.Items.Add("Replenishment Date")
            dd.Items.Add("Inactivity Alert Days")
            dd.Items.Add("Notes")
            dd.Items.Add("Misc 1")
            dd.Items.Add("Misc 2")
            dd.Items.Add("Authorisable")
            dd.Items.Add("Calendar Managed")
            dd.Items.Add("Max Grab")
            dd.Items.Add("Archived")
            dd.Items.Add("Quantity")
            dd.SelectedIndex = ExtractHiddenFieldValue(j)
            dd.Font.Name = "Verdana"
            dd.Font.Size = FontSize.Small
            
            gvr.Cells(j).Controls.Add(dd)
            Dim cid As String = dd.ClientID
            dd.Attributes.Add("onClick", "HiddenField" & j.ToString.Trim & ".value=" & cid & ".selectedIndex; HiddenFieldChanged.value='TRUE'")
        Next
    End Sub

    Protected Sub ProcessProducts(ByVal bAddToDatabase As Boolean)
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataSet As New DataSet
      
        Dim nRow As Integer, nMaxRows As Integer = gvUploadData.Rows.Count
        Dim nCol As Integer, nMaxCols As Integer = gvUploadData.Columns.Count
        Dim i As Integer, sVal As String
        Dim nTargetIndex As Integer
        ' following 2 lists must match the contents of the dropdown list box used to select target (sMarshaller drops initial 2 items)
        Dim sTarget() As String = {"", "", "ProductCode", "ProductDate", "Description", "CostCentre/DeptId", "Category", "SubCategory", "SubCategory2", "MinimumStockLevel", "UnitValue", "SellingPrice", "UnitWeightGrams", "Language", "ItemsPerBox", "ExpiryDate", "ReplenishmentDate", "InactivityAlertDays", "Notes", "Misc1", "Misc2", "Authorisable", "CalendarManaged", "MaxGrab", "Archived", "Quantity"} ' this is the order of significant items presented in the dropdowns
        Dim sMarshaller() As String = {"ProductCode", "ProductDate", "Description", "CostCentre/DeptId", "Category", "SubCategory", "SubCategory2", "MinimumStockLevel", "UnitValue", "SellingPrice", "UnitWeightGrams", "Language", "ItemsPerBox", "ExpiryDate", "ReplenishmentDate", "InactivityAlertDays", "Notes", "Misc1", "Misc2", "Authorisable", "CalendarManaged", "MaxGrab", "Archived", "Quantity"} ' note initialisation is for documentation purposes
        
        Dim sbErrors As New StringBuilder
        Dim bAuthorisable As Boolean
        lblError.Text = ""
      
        nMaxRows = gvUploadData.Rows.Count
        nMaxCols = gvUploadData.Rows(0).Cells.Count
      
        For nRow = 1 To nMaxRows - 1
            For i = sMarshaller.GetLowerBound(0) To sMarshaller.GetUpperBound(0)
                sMarshaller(i) = ""
            Next
            For nCol = 0 To nMaxCols - 1
                nTargetIndex = ExtractHiddenFieldValue(nCol)
                sVal = HttpUtility.HtmlDecode(gvUploadData.Rows(nRow).Cells(nCol).Text)
                If nTargetIndex > 1 Then  ' need to change this in address prog too
                    sMarshaller(nTargetIndex - 2) = sVal.Trim
                End If
            Next
            
            Dim bEmptyRow As Boolean = True
            For i = sMarshaller.GetLowerBound(0) To sMarshaller.GetUpperBound(0)
                If Not sMarshaller(i) = String.Empty Then
                    bEmptyRow = False
                    Exit For
                End If
            Next

            If Not bEmptyRow Then
                
                If sMarshaller(MARSHALL_PRODUCT_CODE) = String.Empty Then
                    sbErrors.Append("No product code specified in row " & nRow.ToString & "<br />")
                ElseIf sMarshaller(MARSHALL_PRODUCT_CODE).Length > 25 Then
                    sbErrors.Append("Product code too long (max 25 chars) in row " & nRow.ToString & "<br />")
                End If

                If sMarshaller(MARSHALL_PRODUCT_DATE) <> String.Empty AndAlso sMarshaller(MARSHALL_PRODUCT_DATE).Length > 10 Then
                    sbErrors.Append("Product value/date too long (max 10 chars) in row " & nRow.ToString & "<br />")
                End If

                If sMarshaller(MARSHALL_DESCRIPTION) = String.Empty Then
                    sbErrors.Append("No product description specified in row " & nRow.ToString & "<br />")
                ElseIf sMarshaller(MARSHALL_DESCRIPTION).Length > 300 Then
                    sbErrors.Append("Product description too long (max 300 chars) in row " & nRow.ToString & "<br />")
                End If

                If sMarshaller(MARSHALL_LANGUAGE) <> String.Empty AndAlso sMarshaller(MARSHALL_LANGUAGE).Length > 20 Then
                    sbErrors.Append("Product language ID too long (max 20 chars) in row " & nRow.ToString & "<br />")
                End If

                If sMarshaller(MARSHALL_CATEGORY) <> String.Empty AndAlso sMarshaller(MARSHALL_CATEGORY).Length > 50 Then
                    sbErrors.Append("Product category too long (max 50 chars) in row " & nRow.ToString & "<br />")
                End If

                If sMarshaller(MARSHALL_SUB_CATEGORY) <> String.Empty AndAlso sMarshaller(MARSHALL_SUB_CATEGORY).Length > 50 Then
                    sbErrors.Append("Product sub-category too long (max 50 chars) in row " & nRow.ToString & "<br />")
                End If

                If sMarshaller(MARSHALL_SUB_CATEGORY_2) <> String.Empty AndAlso sMarshaller(MARSHALL_SUB_CATEGORY_2).Length > 50 Then
                    sbErrors.Append("Product sub-category 2 too long (max 50 chars) in row " & nRow.ToString & "<br />")
                End If

                If sMarshaller(MARSHALL_MISC_1) <> String.Empty AndAlso sMarshaller(MARSHALL_MISC_1).Length > 50 Then
                    sbErrors.Append("Product Misc1 value too long (max 50 chars) in row " & nRow.ToString & "<br />")
                End If

                If sMarshaller(MARSHALL_MISC_2) <> String.Empty AndAlso sMarshaller(MARSHALL_MISC_2).Length > 50 Then
                    sbErrors.Append("Product Misc2 value too long (max 50 chars) in row " & nRow.ToString & "<br />")
                End If

                If sMarshaller(MARSHALL_NOTES) <> String.Empty AndAlso sMarshaller(MARSHALL_NOTES).Length > 1000 Then
                    sbErrors.Append("Product Notes value too long (max 1000 chars) in row " & nRow.ToString & "<br />")
                End If

                If sMarshaller(MARSHALL_MIN_STOCK_LEVEL) <> String.Empty Then
                    If Not IsNumeric(sMarshaller(MARSHALL_MIN_STOCK_LEVEL)) Then
                        sbErrors.Append("Non-numeric min stock level specified in row " & nRow.ToString & "<br />")
                    End If
                End If

                If sMarshaller(MARSHALL_UNIT_VALUE) <> String.Empty Then
                    If sMarshaller(MARSHALL_UNIT_VALUE).StartsWith("£") Then
                        sMarshaller(MARSHALL_UNIT_VALUE) = sMarshaller(MARSHALL_UNIT_VALUE).Substring(1, sMarshaller(MARSHALL_UNIT_VALUE).Length - 1)
                    End If
                    If Not IsNumeric(sMarshaller(MARSHALL_UNIT_VALUE)) Then
                        sbErrors.Append("Non-numeric unit value specified in row " & nRow.ToString & "<br />")
                    End If
                End If

                If sMarshaller(MARSHALL_SELLING_PRICE) <> String.Empty Then
                    If sMarshaller(MARSHALL_SELLING_PRICE).StartsWith("£") Then
                        sMarshaller(MARSHALL_SELLING_PRICE) = sMarshaller(MARSHALL_UNIT_VALUE).Substring(1, sMarshaller(MARSHALL_SELLING_PRICE).Length - 1)
                    End If
                    If Not IsNumeric(sMarshaller(MARSHALL_SELLING_PRICE)) Then
                        sbErrors.Append("Non-numeric selling price specified in row " & nRow.ToString & "<br />")
                    End If
                End If

                If sMarshaller(MARSHALL_UNIT_WEIGHT_GRAMS) <> String.Empty Then
                    If Not IsNumeric(sMarshaller(MARSHALL_UNIT_WEIGHT_GRAMS)) Then
                        sbErrors.Append("Non-numeric unit weight specified in row " & nRow.ToString & "<br />")
                    End If
                End If

                If sMarshaller(MARSHALL_ITEMS_PER_BOX) <> String.Empty Then
                    If Not IsNumeric(sMarshaller(MARSHALL_ITEMS_PER_BOX)) Then
                        sbErrors.Append("Non-numeric items / box specified in row " & nRow.ToString & "<br />")
                    End If
                End If

                If sMarshaller(MARSHALL_EXPIRY_DATE) <> String.Empty Then
                    If Not IsDate(sMarshaller(MARSHALL_EXPIRY_DATE)) Then
                        sbErrors.Append("Illegal date format for expiry date specified in row " & nRow.ToString & "<br />")
                    End If
                End If

                If sMarshaller(MARSHALL_REPLENISHMENT_DATE) <> String.Empty Then
                    If Not IsDate(sMarshaller(MARSHALL_REPLENISHMENT_DATE)) Then
                        sbErrors.Append("Illegal date format for replenishment date specified in row " & nRow.ToString & "<br />")
                    End If
                End If

                If sMarshaller(MARSHALL_INACTIVITY_ALERT_DAYS) <> String.Empty Then
                    If Not IsNumeric(sMarshaller(MARSHALL_INACTIVITY_ALERT_DAYS)) Then
                        sbErrors.Append("Non-numeric inactivity alert days specified in row " & nRow.ToString & "<br />")
                    End If
                End If

                bAuthorisable = False
                If sMarshaller(MARSHALL_AUTHORISABLE) <> String.Empty Then
                    If Not (sMarshaller(MARSHALL_AUTHORISABLE).ToUpper = "Y" Or sMarshaller(MARSHALL_AUTHORISABLE).ToUpper = "YES" Or sMarshaller(MARSHALL_AUTHORISABLE).ToUpper = "1" _
                      Or sMarshaller(MARSHALL_AUTHORISABLE).ToUpper = "N" Or sMarshaller(MARSHALL_AUTHORISABLE).ToUpper = "NO" Or sMarshaller(MARSHALL_AUTHORISABLE).ToUpper = "0" _
                      Or sMarshaller(MARSHALL_AUTHORISABLE).ToUpper = "TRUE" Or sMarshaller(MARSHALL_AUTHORISABLE).ToUpper = "FALSE") _
                    Then
                        sbErrors.Append("Cannot interpret authorisation value specified in row " & nRow.ToString & "<br />")
                    Else
                        If (sMarshaller(MARSHALL_AUTHORISABLE).ToUpper = "Y" Or sMarshaller(MARSHALL_AUTHORISABLE).ToUpper = "YES" Or sMarshaller(MARSHALL_AUTHORISABLE).ToUpper = "1" Or sMarshaller(MARSHALL_AUTHORISABLE).ToUpper = "TRUE") Then
                            bAuthorisable = True
                        End If
                        If pnAuthoriser = 0 Then
                            sbErrors.Append("Authorisation requested but no authoriser specified in row " & nRow.ToString & "<br />")
                        End If
                    End If
                End If

                If sMarshaller(MARSHALL_CALENDAR_MANAGED) <> String.Empty Then
                    If Not (sMarshaller(MARSHALL_CALENDAR_MANAGED).ToUpper = "Y" Or sMarshaller(MARSHALL_CALENDAR_MANAGED).ToUpper = "YES" Or sMarshaller(MARSHALL_CALENDAR_MANAGED).ToUpper = "1" _
                      Or sMarshaller(MARSHALL_CALENDAR_MANAGED).ToUpper = "N" Or sMarshaller(MARSHALL_CALENDAR_MANAGED).ToUpper = "NO" Or sMarshaller(MARSHALL_CALENDAR_MANAGED).ToUpper = "0" _
                      Or sMarshaller(MARSHALL_CALENDAR_MANAGED).ToUpper = "TRUE" Or sMarshaller(MARSHALL_CALENDAR_MANAGED).ToUpper = "FALSE") _
                    Then
                        sbErrors.Append("Cannot interpret calendar managed value specified in row " & nRow.ToString & "<br />")
                    Else
                        If bAuthorisable = True Then
                            If (sMarshaller(MARSHALL_CALENDAR_MANAGED).ToUpper = "Y" Or sMarshaller(MARSHALL_CALENDAR_MANAGED).ToUpper = "YES" Or sMarshaller(MARSHALL_CALENDAR_MANAGED).ToUpper = "1" Or sMarshaller(MARSHALL_CALENDAR_MANAGED).ToUpper = "TRUE") Then
                                sbErrors.Append("Calendar managed and authorisable cannot be combined in row " & nRow.ToString & "<br />")
                            End If
                        End If
                    End If
                End If

                If sMarshaller(MARSHALL_PRODUCTQUANTITY) <> String.Empty Then
                    If Not IsNumeric(sMarshaller(MARSHALL_PRODUCTQUANTITY)) Then
                        sbErrors.Append("Non-numeric product quantity specified in row " & nRow.ToString & "<br />")
                    End If
                    If IsNumeric(sMarshaller(MARSHALL_PRODUCTQUANTITY)) Then
                        If Not cbIgnoreNegativeQuantities.Checked Then
                            If CInt(sMarshaller(MARSHALL_PRODUCTQUANTITY)) < 0 Then
                                sbErrors.Append("Negative product quantity specified in row " & nRow.ToString & "<br />")
                            End If
                        End If
                    End If
                    'If tbBayKey.Text <> String.Empty Then
                    '    If Not IsNumeric(tbBayKey.Text) Then
                    '        sbErrors.Append("Non numeric bay key specified, processing row " & nRow.ToString & "<br />")
                    '    Else
                    '        If lblBayName.Text = String.Empty Then
                    '            sbErrors.Append("Bay with key " & tbBayKey.Text & " does not exist, processing row " & nRow.ToString & "<br />")
                    '        End If
                    '    End If
                    'Else
                    '    sbErrors.Append("No bay key specified for quantity in row " & nRow.ToString & "<br />")
                    'End If
                    If cbAddQuantity.Checked Then
                        If ddlBay.SelectedIndex <= 0 Then
                            sbErrors.Append("No warehouse location for quantity specified in row " & nRow.ToString & "<br />")
                        End If
                    Else
                        sbErrors.Append("A warehouse location must be selected in row " & nRow.ToString & "<br />")
                    End If
                End If
                
                If sMarshaller(MARSHALL_PRODUCTQUANTITY) = String.Empty Then
                    If cbAddQuantity.Checked Then
                        sbErrors.Append("Add quantity check box selected but no quantity column found, in row " & nRow.ToString & "<br />")
                    End If
                End If

                If ProductExists(sMarshaller(MARSHALL_PRODUCT_CODE) & String.Empty, sMarshaller(MARSHALL_PRODUCT_DATE) & String.Empty) Then
                    sbErrors.Append("A product with this code & value/date combination already exists, in row " & nRow.ToString & "<br />")
                End If

                If bAddToDatabase And sbErrors.Length = 0 Then
                    Call AddNewProduct(sMarshaller)
                End If
            Else
                'sbErrors.Append("...skipping empty row " & nRow.ToString & "<br />")
            End If
        Next
      
        If sbErrors.Length > 0 Then
            lblError.Text = sbErrors.ToString
        End If
    End Sub

    Protected Function ProductExists(sProductCode As String, sProductDate As String) As Boolean
        ProductExists = False
        Dim sSQL As String = "SELECT ProductCode FROM LogisticProduct WHERE CustomerKey = " & pnSelectedCustomerKey & " AND ProductCode = '" & String.Empty & sProductCode.Replace("'", "''") & "' AND ISNULL(ProductDate, '') = '" & sProductDate.Replace("'", "''") & "'"
        Dim dtProduct As DataTable = ExecuteQueryToDataTable(sSQL)
        If dtProduct.Rows.Count > 0 Then
            ProductExists = True
        End If
    End Function
    
    'Protected Function BayName() As String
    '    BayName = String.Empty
    '    Try
    '        BayName = ExecuteQueryToDataTable("SELECT WarehouseBayId FROM WarehouseBay WHERE WarehouseBayKey = " & tbBayKey.Text).Rows(0).Item(0)
    '    Catch ex As Exception
    '    End Try
    'End Function
    
    Protected Sub AddNewProduct(ByVal sMarshaller() As String)
        lblError.Text = ""
        Dim nIndex As Integer
        Dim oConn As New SqlConnection(gsConn)
        Dim oTrans As SqlTransaction
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Product_AddWithAccessControl8", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
 
        Dim paramUserKey As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int)
        paramUserKey.Value = 0
        oCmd.Parameters.Add(paramUserKey)

        Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int)
        paramCustomerKey.Value = pnSelectedCustomerKey
        oCmd.Parameters.Add(paramCustomerKey)
 
        Dim paramProductCode As SqlParameter = New SqlParameter("@ProductCode", SqlDbType.NVarChar, 25)
        paramProductCode.Value = sMarshaller(MARSHALL_PRODUCT_CODE)
        oCmd.Parameters.Add(paramProductCode)
      
        Dim paramProductDate As SqlParameter = New SqlParameter("@ProductDate", SqlDbType.NVarChar, 10)
        paramProductDate.Value = sMarshaller(MARSHALL_PRODUCT_DATE) & String.Empty
        oCmd.Parameters.Add(paramProductDate)
 
        Dim paramMinimumStockLevel As SqlParameter = New SqlParameter("@MinimumStockLevel", SqlDbType.Int, 4)
        If IsNumeric(sMarshaller(MARSHALL_MIN_STOCK_LEVEL)) Then
            paramMinimumStockLevel.Value = CLng(sMarshaller(MARSHALL_MIN_STOCK_LEVEL))
        Else
            paramMinimumStockLevel.Value = 0
        End If
        oCmd.Parameters.Add(paramMinimumStockLevel)
      
        Dim paramDescription As SqlParameter = New SqlParameter("@ProductDescription", SqlDbType.NVarChar, 300)
        paramDescription.Value = sMarshaller(MARSHALL_DESCRIPTION)
        oCmd.Parameters.Add(paramDescription)
      
        Dim paramItemsPerBox As SqlParameter = New SqlParameter("@ItemsPerBox", SqlDbType.Int, 4)
        If IsNumeric(sMarshaller(MARSHALL_ITEMS_PER_BOX)) Then
            paramItemsPerBox.Value = CLng(sMarshaller(MARSHALL_ITEMS_PER_BOX))
        Else
            paramItemsPerBox.Value = 0
        End If
        oCmd.Parameters.Add(paramItemsPerBox)
      
        Dim paramCategory As SqlParameter = New SqlParameter("@ProductCategory", SqlDbType.NVarChar, 50)
        paramCategory.Value = sMarshaller(MARSHALL_CATEGORY)
        oCmd.Parameters.Add(paramCategory)
      
        Dim paramSubCategory As SqlParameter = New SqlParameter("@SubCategory", SqlDbType.NVarChar, 50)
        paramSubCategory.Value = sMarshaller(MARSHALL_SUB_CATEGORY)
        oCmd.Parameters.Add(paramSubCategory)
      
        Dim paramSubCategory2 As SqlParameter = New SqlParameter("@SubCategory2", SqlDbType.NVarChar, 50)
        paramSubCategory2.Value = sMarshaller(MARSHALL_SUB_CATEGORY_2)
        oCmd.Parameters.Add(paramSubCategory2)
      
        Dim paramUnitValue As SqlParameter = New SqlParameter("@UnitValue", SqlDbType.Money, 8)
        If IsNumeric(sMarshaller(MARSHALL_UNIT_VALUE)) Then
            If CDec(sMarshaller(MARSHALL_UNIT_VALUE)) > 0 Then
                paramUnitValue.Value = CDec(sMarshaller(MARSHALL_UNIT_VALUE))
            Else
                paramUnitValue.Value = 0
            End If
        Else
            paramUnitValue.Value = 0
        End If
        oCmd.Parameters.Add(paramUnitValue)
      
        Dim paramUnitValue2 As SqlParameter = New SqlParameter("@UnitValue2", SqlDbType.Money, 8)
        If IsNumeric(sMarshaller(MARSHALL_SELLING_PRICE)) Then
            If CDec(sMarshaller(MARSHALL_SELLING_PRICE)) > 0 Then
                paramUnitValue2.Value = CDec(sMarshaller(MARSHALL_SELLING_PRICE))
            Else
                paramUnitValue2.Value = 0
            End If
        Else
            paramUnitValue2.Value = 0
        End If
        oCmd.Parameters.Add(paramUnitValue2)

        Dim paramLanguage As SqlParameter = New SqlParameter("@LanguageId", SqlDbType.NVarChar, 20)
        paramLanguage.Value = sMarshaller(MARSHALL_LANGUAGE)
        oCmd.Parameters.Add(paramLanguage)

        Dim paramDepartment As SqlParameter = New SqlParameter("@ProductDepartmentId", SqlDbType.NVarChar, 20)
        paramDepartment.Value = sMarshaller(MARSHALL_COST_CENTRE_DEPT_ID)
        oCmd.Parameters.Add(paramDepartment)
      
        Dim paramWeight As SqlParameter = New SqlParameter("@UnitWeightGrams", SqlDbType.Int, 4)
        If IsNumeric(sMarshaller(MARSHALL_UNIT_WEIGHT_GRAMS)) Then
            paramWeight.Value = CLng(sMarshaller(MARSHALL_UNIT_WEIGHT_GRAMS))
        Else
            paramWeight.Value = 0
        End If
        oCmd.Parameters.Add(paramWeight)
      
        Dim paramStockOwnedByKey As SqlParameter = New SqlParameter("@StockOwnedByKey", SqlDbType.Int, 4)
        paramStockOwnedByKey.Value = 0
        oCmd.Parameters.Add(paramStockOwnedByKey)
      
        Dim paramMisc1 As SqlParameter = New SqlParameter("@Misc1", SqlDbType.NVarChar, 50)
        paramMisc1.Value = sMarshaller(MARSHALL_MISC_1)
        oCmd.Parameters.Add(paramMisc1)
      
        Dim paramMisc2 As SqlParameter = New SqlParameter("@Misc2", SqlDbType.NVarChar, 50)
        paramMisc2.Value = sMarshaller(MARSHALL_MISC_2)
        oCmd.Parameters.Add(paramMisc2)
      
        Dim paramArchive As SqlParameter = New SqlParameter("@ArchiveFlag", SqlDbType.NVarChar, 1)
        nIndex = MARSHALL_ARCHIVED
        If sMarshaller(nIndex) <> String.Empty AndAlso (sMarshaller(nIndex) = "1" Or sMarshaller(nIndex).ToUpper = "Y" Or sMarshaller(nIndex).ToUpper = "YES" Or sMarshaller(nIndex).ToUpper = "TRUE") Then
            paramArchive.Value = "Y"
        Else
            paramArchive.Value = "N"
        End If
        oCmd.Parameters.Add(paramArchive)
    
        Dim paramStatus As SqlParameter = New SqlParameter("@Status", SqlDbType.TinyInt)
        paramStatus.Value = 0
        oCmd.Parameters.Add(paramStatus)

        Dim paramExpiryDate As SqlParameter = New SqlParameter("@ExpiryDate", SqlDbType.SmallDateTime)
        Dim sExpiryDate As String
        sExpiryDate = sMarshaller(MARSHALL_EXPIRY_DATE)
        If sExpiryDate <> "" Then
            Try
                sExpiryDate = DateTime.Parse(sExpiryDate)
            Catch ex As Exception
                lblError.Text = "ERROR: Invalid Expiry Date"
                'Exit Sub
                sExpiryDate = String.Empty
            End Try
        End If
        If sExpiryDate = "" Then
            paramExpiryDate.Value = Nothing
        Else
            'paramExpiryDate.Value = sExpiryDate
            paramExpiryDate.Value = DateTime.Parse(sExpiryDate)
            
        End If
        paramExpiryDate.Value = Nothing
        oCmd.Parameters.Add(paramExpiryDate)

        Dim paramReplenishmentDate As SqlParameter = New SqlParameter("@ReplenishmentDate", SqlDbType.SmallDateTime)
        Dim sReplenishmentDate As String
        sReplenishmentDate = sMarshaller(MARSHALL_REPLENISHMENT_DATE)
        If sReplenishmentDate <> "" Then
            Try
                sReplenishmentDate = DateTime.Parse(sReplenishmentDate)
            Catch ex As Exception
                lblError.Text = "ERROR: Invalid Renewal / Review Date"
                'Exit Sub
                sReplenishmentDate = String.Empty
            End Try
        End If
        If sReplenishmentDate = "" Then
            paramReplenishmentDate.Value = Nothing
        Else
            paramReplenishmentDate.Value = sReplenishmentDate
        End If
        paramReplenishmentDate.Value = Nothing
        oCmd.Parameters.Add(paramReplenishmentDate)
    
        Dim paramSerialNumbers As SqlParameter = New SqlParameter("@SerialNumbersFlag", SqlDbType.NVarChar, 1)
        paramSerialNumbers.Value = "N"
        oCmd.Parameters.Add(paramSerialNumbers)
      
        Dim paramAdRotatorText As SqlParameter = New SqlParameter("@AdRotatorText", SqlDbType.NVarChar, 120)
        paramAdRotatorText.Value = String.Empty
        oCmd.Parameters.Add(paramAdRotatorText)
      
        Dim paramWebsiteAdRotatorFlag As SqlParameter = New SqlParameter("@WebsiteAdRotatorFlag", SqlDbType.Bit)
        paramWebsiteAdRotatorFlag.Value = 0
        oCmd.Parameters.Add(paramWebsiteAdRotatorFlag)
      
        Dim paramNotes As SqlParameter = New SqlParameter("@Notes", SqlDbType.NVarChar, 1000)
        paramNotes.Value = sMarshaller(MARSHALL_NOTES)
        oCmd.Parameters.Add(paramNotes)

        Dim paramViewOnWebForm As SqlParameter = New SqlParameter("@ViewOnWebForm", SqlDbType.Bit)
        paramViewOnWebForm.Value = 0
        oCmd.Parameters.Add(paramViewOnWebForm)
 
        Dim paramDefaultAccessFlag As SqlParameter = New SqlParameter("@DefaultAccessFlag", SqlDbType.Bit)
        paramDefaultAccessFlag.Value = Not pbExplicitProductPermissions
        oCmd.Parameters.Add(paramDefaultAccessFlag)

        Dim paramRotationProductKey As SqlParameter = New SqlParameter("@RotationProductKey", SqlDbType.Int, 4)
        paramRotationProductKey.Value = System.Data.SqlTypes.SqlInt32.Null
        oCmd.Parameters.Add(paramRotationProductKey)

        Dim paramInactivityAlertDays As SqlParameter = New SqlParameter("@InactivityAlertDays", SqlDbType.Int, 4)
        If IsNumeric(sMarshaller(MARSHALL_INACTIVITY_ALERT_DAYS)) Then
            paramInactivityAlertDays.Value = CLng(sMarshaller(MARSHALL_INACTIVITY_ALERT_DAYS))
        Else
            paramInactivityAlertDays.Value = 0
        End If
        oCmd.Parameters.Add(paramInactivityAlertDays)
    
        Dim paramCalendarManaged As SqlParameter = New SqlParameter("@CalendarManaged", SqlDbType.Bit)
        nIndex = MARSHALL_CALENDAR_MANAGED
        If sMarshaller(nIndex) <> String.Empty AndAlso (sMarshaller(nIndex) = "1" Or sMarshaller(nIndex).ToUpper = "Y" Or sMarshaller(nIndex).ToUpper = "YES" Or sMarshaller(nIndex).ToUpper = "TRUE") Then
            paramCalendarManaged.Value = 1
        Else
            paramCalendarManaged.Value = 0
        End If
        oCmd.Parameters.Add(paramCalendarManaged)

        Dim paramOnDemand As SqlParameter = New SqlParameter("@OnDemand", SqlDbType.Int)
        paramOnDemand.Value = 0
        oCmd.Parameters.Add(paramOnDemand)
      
        Dim paramOnDemandPriceList As SqlParameter = New SqlParameter("@OnDemandPriceList", SqlDbType.Int)
        paramOnDemandPriceList.Value = 0
        oCmd.Parameters.Add(paramOnDemandPriceList)
      
        Dim paramCustomLetter As SqlParameter = New SqlParameter("@CustomLetter", SqlDbType.Bit)
        paramCustomLetter.Value = 0
        oCmd.Parameters.Add(paramCustomLetter)
      
        Dim paramProductKey As SqlParameter = New SqlParameter("@ProductKey", SqlDbType.Int, 4)
        paramProductKey.Direction = ParameterDirection.Output
        oCmd.Parameters.Add(paramProductKey)
 
        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
            Dim lProductKey As Long = CLng(oCmd.Parameters("@ProductKey").Value)

            tbResults.Text += "Added product " & sMarshaller(MARSHALL_PRODUCT_CODE) & " as product " & lProductKey.ToString & Environment.NewLine

            If sMarshaller(MARSHALL_AUTHORISABLE) <> String.Empty Then
                If (sMarshaller(MARSHALL_AUTHORISABLE) = "1" Or sMarshaller(MARSHALL_AUTHORISABLE).ToUpper = "Y" Or sMarshaller(MARSHALL_AUTHORISABLE).ToUpper = "YES" Or sMarshaller(MARSHALL_AUTHORISABLE).ToUpper = "TRUE") Then
                    If ddlAuthoriser.SelectedIndex > 0 Then
                        Call SetAuthorisable(lProductKey)
                        tbResults.Text += "Product code: " & sMarshaller(MARSHALL_PRODUCT_CODE) & "; value date: " & sMarshaller(MARSHALL_PRODUCT_DATE) & " set AUTHORISABLE" & Environment.NewLine
                    Else
                        tbResults.Text += "ERROR: product code: " & sMarshaller(MARSHALL_PRODUCT_CODE) & "; value date: " & sMarshaller(MARSHALL_PRODUCT_DATE) & "; Could not set this authorisable as no authoriser was specified" & Environment.NewLine
                    End If
                End If
            End If
            
            If sMarshaller(MARSHALL_MAX_GRAB) <> String.Empty Then
                If IsNumeric(sMarshaller(MARSHALL_MAX_GRAB)) Then
                    If CInt(sMarshaller(MARSHALL_MAX_GRAB) > 0) Then
                        Call SetMaxGrab(CInt(sMarshaller(MARSHALL_MAX_GRAB)), lProductKey)
                        tbResults.Text += "Set MAX GRAB on product code: " & sMarshaller(MARSHALL_PRODUCT_CODE) & "; value date: " & sMarshaller(MARSHALL_PRODUCT_DATE) & " to " & sMarshaller(MARSHALL_MAX_GRAB) & Environment.NewLine
                    End If
                End If
            End If
            
            If sMarshaller(MARSHALL_PRODUCTQUANTITY) <> String.Empty Then
                If IsNumeric(sMarshaller(MARSHALL_PRODUCTQUANTITY)) Then
                    If CInt(sMarshaller(MARSHALL_PRODUCTQUANTITY) > 0) Then
                        Call SetQuantity(CInt(sMarshaller(MARSHALL_PRODUCTQUANTITY)), lProductKey)
                        tbResults.Text += "Set Quantity on product code: " & sMarshaller(MARSHALL_PRODUCT_CODE) & "; value date: " & sMarshaller(MARSHALL_PRODUCT_DATE) & " to " & sMarshaller(MARSHALL_PRODUCTQUANTITY) & " in location " & ddlWarehouse.SelectedItem.Text & " / " & ddlRack.SelectedItem.Text & " / " & ddlSection.SelectedItem.Text & " / " & ddlBay.SelectedItem.Text & Environment.NewLine
                    End If
                End If
            End If

        Catch ex As SqlException
            If ex.Number = 2627 Then
                lblError.Text = "ERROR: A record already exists with the same product CODE and DATE combination"
                tbResults.Text += "ERROR: product code: " & sMarshaller(MARSHALL_PRODUCT_CODE) & "; value date: " & sMarshaller(MARSHALL_PRODUCT_DATE) & "; A record already exists with the same product CODE and DATE combination" & Environment.NewLine
            Else
                lblError.Text = ex.ToString
            End If
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub SetMaxGrab(ByVal nMaxGrab As Integer, ByVal lProductKey As Long)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand
        Dim sSQL As String = "UPDATE UserProductProfile SET ApplyMaxGrab = 1, MaxGrabQty = " & nMaxGrab.ToString & " WHERE ProductKey = " & lProductKey.ToString
        Try
            oConn.Open()
            oCmd = New SqlCommand(sSQL, oConn)
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            WebMsgBox.Show("Error in SetMaxGrab: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub SetQuantity(ByVal nQuantity As Integer, ByVal lProductKey As Long)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand
        'Dim sSQL As String = "INSERT INTO LogisticProductLocation (LogisticProductKey, WarehouseBayKey, LogisticProductQuantity, DateStored) VALUES (" & lProductKey & ", " & tbBayKey.Text & ", " & nQuantity & ", GETDATE())"
        Dim sSQL As String = "INSERT INTO LogisticProductLocation (LogisticProductKey, WarehouseBayKey, LogisticProductQuantity, DateStored) VALUES (" & lProductKey & ", " & ddlBay.SelectedValue & ", " & nQuantity & ", GETDATE())"
        Try
            oConn.Open()
            oCmd = New SqlCommand(sSQL, oConn)
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            WebMsgBox.Show("Error in SetQuantity: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub

    Protected Sub SetAuthorisable(ByVal lProductKey As Long)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Product_SetAuthorisable", oConn)
        oCmd.CommandType = CommandType.StoredProcedure

        Dim paramLogisticProductKey As SqlParameter = New SqlParameter("@LogisticProductKey", SqlDbType.Int)
        paramLogisticProductKey.Value = lProductKey
        oCmd.Parameters.Add(paramLogisticProductKey)
      
        Dim paramDefaultAuthorisationGrantTimeoutHours As SqlParameter = New SqlParameter("@DefaultAuthorisationGrantTimeoutHours", SqlDbType.Int)
        paramDefaultAuthorisationGrantTimeoutHours.Value = 0
        oCmd.Parameters.Add(paramDefaultAuthorisationGrantTimeoutHours)
      
        Dim paramDefaultAuthorisationLifetimeHours As SqlParameter = New SqlParameter("@DefaultAuthorisationLifetimeHours", SqlDbType.Int)
        paramDefaultAuthorisationLifetimeHours.Value = 0
        oCmd.Parameters.Add(paramDefaultAuthorisationLifetimeHours)
      
        Dim paramAuthoriser As SqlParameter = New SqlParameter("@AuthoriserKey", SqlDbType.Int)
        paramAuthoriser.Value = pnAuthoriser
        oCmd.Parameters.Add(paramAuthoriser)
      
        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As SqlException
            WebMsgBox.Show("Error in SetAuthorisable: " & ex.Message)
        End Try
        tbResults.Text += "Set product " & lProductKey & " authorisable" & Environment.NewLine
    End Sub
    
    Protected Sub Page_PreRender(ByVal sender As Object, ByVal e As System.EventArgs)
        If gvUploadData.Rows.Count > 0 Then
            Call AddDropDowns()
        End If
    End Sub
  
    Protected Function ExtractHiddenFieldValue(ByVal nIndex As Integer) As Integer
        Select Case nIndex
            Case 0
                If (Not HiddenField0.Value Is Nothing) OrElse (HiddenField0.Value > 0) Then
                    Return CInt(HiddenField0.Value)
                Else
                    Return 0
                End If
            Case 1
                If (Not HiddenField1.Value Is Nothing) OrElse (HiddenField1.Value > 0) Then
                    Return CInt(HiddenField1.Value)
                Else
                    Return 0
                End If
            Case 2
                If (Not HiddenField2.Value Is Nothing) OrElse (HiddenField2.Value > 0) Then
                    Return CInt(HiddenField2.Value)
                Else
                    Return 0
                End If
            Case 3
                If (Not HiddenField3.Value Is Nothing) OrElse (HiddenField3.Value > 0) Then
                    Return CInt(HiddenField3.Value)
                Else
                    Return 0
                End If
            Case 4
                If (Not HiddenField4.Value Is Nothing) OrElse (HiddenField4.Value > 0) Then
                    Return CInt(HiddenField4.Value)
                Else
                    Return 0
                End If
            Case 5
                If (Not HiddenField5.Value Is Nothing) OrElse (HiddenField5.Value > 0) Then
                    Return CInt(HiddenField5.Value)
                Else
                    Return 0
                End If
            Case 6
                If (Not HiddenField6.Value Is Nothing) OrElse (HiddenField6.Value > 0) Then
                    Return CInt(HiddenField6.Value)
                Else
                    Return 0
                End If
            Case 7
                If (Not HiddenField7.Value Is Nothing) OrElse (HiddenField7.Value > 0) Then
                    Return CInt(HiddenField7.Value)
                Else
                    Return 0
                End If
            Case 8
                If (Not HiddenField8.Value Is Nothing) OrElse (HiddenField8.Value > 0) Then
                    Return CInt(HiddenField8.Value)
                Else
                    Return 0
                End If
            Case 9
                If (Not HiddenField9.Value Is Nothing) OrElse (HiddenField9.Value > 0) Then
                    Return CInt(HiddenField9.Value)
                Else
                    Return 0
                End If
            Case 10
                If (Not HiddenField10.Value Is Nothing) OrElse (HiddenField10.Value > 0) Then
                    Return CInt(HiddenField10.Value)
                Else
                    Return 0
                End If
            Case 11
                If (Not HiddenField11.Value Is Nothing) OrElse (HiddenField11.Value > 0) Then
                    Return CInt(HiddenField11.Value)
                Else
                    Return 0
                End If
            Case 12
                If (Not HiddenField12.Value Is Nothing) OrElse (HiddenField12.Value > 0) Then
                    Return CInt(HiddenField12.Value)
                Else
                    Return 0
                End If
            Case 13
                If (Not HiddenField13.Value Is Nothing) OrElse (HiddenField13.Value > 0) Then
                    Return CInt(HiddenField13.Value)
                Else
                    Return 0
                End If
            Case 14
                If (Not HiddenField14.Value Is Nothing) OrElse (HiddenField14.Value > 0) Then
                    Return CInt(HiddenField14.Value)
                Else
                    Return 0
                End If
            Case 15
                If (Not HiddenField15.Value Is Nothing) OrElse (HiddenField15.Value > 0) Then
                    Return CInt(HiddenField15.Value)
                Else
                    Return 0
                End If
            Case 16
                If (Not HiddenField16.Value Is Nothing) OrElse (HiddenField16.Value > 0) Then
                    Return CInt(HiddenField16.Value)
                Else
                    Return 0
                End If
            Case 17
                If (Not HiddenField17.Value Is Nothing) OrElse (HiddenField17.Value > 0) Then
                    Return CInt(HiddenField17.Value)
                Else
                    Return 0
                End If
            Case 18
                If (Not HiddenField18.Value Is Nothing) OrElse (HiddenField18.Value > 0) Then
                    Return CInt(HiddenField18.Value)
                Else
                    Return 0
                End If
            Case 19
                If (Not HiddenField19.Value Is Nothing) OrElse (HiddenField19.Value > 0) Then
                    Return CInt(HiddenField19.Value)
                Else
                    Return 0
                End If
            Case 20
                If (Not HiddenField20.Value Is Nothing) OrElse (HiddenField20.Value > 0) Then
                    Return CInt(HiddenField20.Value)
                Else
                    Return 0
                End If
            Case 21
                If (Not HiddenField21.Value Is Nothing) OrElse (HiddenField21.Value > 0) Then
                    Return CInt(HiddenField21.Value)
                Else
                    Return 0
                End If
            Case 22
                If (Not HiddenField22.Value Is Nothing) OrElse (HiddenField22.Value > 0) Then
                    Return CInt(HiddenField22.Value)
                Else
                    Return 0
                End If
            Case 23
                If (Not HiddenField23.Value Is Nothing) OrElse (HiddenField23.Value > 0) Then
                    Return CInt(HiddenField23.Value)
                Else
                    Return 0
                End If
            Case 24
                If (Not HiddenField24.Value Is Nothing) OrElse (HiddenField24.Value > 0) Then
                    Return CInt(HiddenField24.Value)
                Else
                    Return 0
                End If
            Case 25
                If (Not HiddenField25.Value Is Nothing) OrElse (HiddenField25.Value > 0) Then
                    Return CInt(HiddenField25.Value)
                Else
                    Return 0
                End If
            Case 26
                If (Not HiddenField26.Value Is Nothing) OrElse (HiddenField26.Value > 0) Then
                    Return CInt(HiddenField26.Value)
                Else
                    Return 0
                End If
            Case 27
                If (Not HiddenField27.Value Is Nothing) OrElse (HiddenField27.Value > 0) Then
                    Return CInt(HiddenField27.Value)
                Else
                    Return 0
                End If
            Case 28
                If (Not HiddenField28.Value Is Nothing) OrElse (HiddenField28.Value > 0) Then
                    Return CInt(HiddenField28.Value)
                Else
                    Return 0
                End If
            Case 29
                If (Not HiddenField29.Value Is Nothing) OrElse (HiddenField29.Value > 0) Then
                    Return CInt(HiddenField29.Value)
                Else
                    Return 0
                End If
            Case 30
                If (Not HiddenField30.Value Is Nothing) OrElse (HiddenField30.Value > 0) Then
                    Return CInt(HiddenField30.Value)
                Else
                    Return 0
                End If
        End Select
    End Function

    Protected Sub btnHelp_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If btnHelp.Text.ToLower.Contains("show") Then
            pnlHelp.Visible = True
            btnHelp.Text = "hide help"
        Else
            pnlHelp.Visible = False
            btnHelp.Text = "show help"
        End If
    End Sub

    Protected Sub GetCustomerAccountCodes()
        Dim oConn As New SqlConnection(gsConn)
        ddlCustomers.Items.Clear()
        Dim oCmd As New SqlCommand("spASPNET_Customer_GetActiveCustomerCodes", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Try
            oConn.Open()
            ddlCustomers.DataSource = oCmd.ExecuteReader()
            ddlCustomers.DataTextField = "CustomerAccountCode"
            ddlCustomers.DataValueField = "CustomerKey"
            ddlCustomers.DataBind()
        Catch ex As SqlException
            lblError.Text = ""
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
    End Sub

    Protected Sub ddlCustomers_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        pnSelectedCustomerKey = ddl.SelectedValue
        pbExplicitProductPermissions = bSetExplicitProductPermissionsFlag()
        btnUploadProducts.Enabled = True
        btnCheckData.Enabled = True
        lbLegendlDefaultAuthoriser.Enabled = True
        ddlAuthoriser.Enabled = True
       
        FileUpload1.Enabled = True
        cbColumnHeadingsInRow1.Enabled = True
        btnReadProducts.Enabled = True
       
        lblCustomer.Text = "to " & ddl.SelectedItem.Text
       
        If ddl.Items(0).Text = String.Empty Then
            ddl.Items.RemoveAt(0)
        End If
        Call PopulateSuperUserDropdown(ddlAuthoriser)
    End Sub
    
    Protected Sub PopulateSuperUserDropdown(ByVal ddl As DropDownList)
        Dim oConn As New SqlConnection(gsConn)
        Dim oDatatable As New DataTable
        Dim oAdapter As New SqlDataAdapter("spASPNET_UserProfile_GetAllSuperUsersForCustomer2", oConn)
      
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@CustomerKey").Value = pnSelectedCustomerKey

        oAdapter.Fill(oDatatable)
        ddl.Items.Clear()
        Dim li As New ListItem
        li.Text = "-- select authoriser --"
        li.Value = 0
        ddl.Items.Add(li)

        For Each dr As DataRow In oDatatable.Rows
            Dim li2 As New ListItem
            Dim sFirstName As String = dr("FirstName")
            Dim sLastName As String = dr("LastName")
            li2.Text = Char.ToUpper(sFirstName(0)) & sFirstName.Substring(1) & " " & Char.ToUpper(sLastName(0)) & sLastName.Substring(1) & "  (" & dr("UserId") & ")"
            li2.Value = dr("key")
            ddl.Items.Add(li2)
        Next
    End Sub

    Protected Sub lnkbtnRestart_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call Initialise()
    End Sub
    
    Protected Sub ddlAuthoriser_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        pnAuthoriser = ddl.SelectedValue
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

    Property pnSelectedCustomerKey() As Long
        Get
            Dim o As Object = ViewState("UP_SelectedCustomerKey")
            If o Is Nothing Then
                Return -1
            End If
            Return CLng(o)
        End Get
        Set(ByVal Value As Long)
            ViewState("UP_SelectedCustomerKey") = Value
        End Set
    End Property
   
    Property pnAuthoriser() As Long
        Get
            Dim o As Object = ViewState("UP_Authoriser")
            If o Is Nothing Then
                Return -1
            End If
            Return CLng(o)
        End Get
        Set(ByVal Value As Long)
            ViewState("UP_Authoriser") = Value
        End Set
    End Property
   
    Property pnBayKey() As Long
        Get
            Dim o As Object = ViewState("UP_BayKey")
            If o Is Nothing Then
                Return -1
            End If
            Return CLng(o)
        End Get
        Set(ByVal Value As Long)
            ViewState("UP_BayKey") = Value
        End Set
    End Property

    Property pbExplicitProductPermissions() As Boolean
        Get
            Dim o As Object = ViewState("UP_ExplicitProductPermissions")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("UP_ExplicitProductPermissions") = Value
        End Set
    End Property

    Protected Sub cbAddQuantity_CheckedChanged(sender As Object, e As System.EventArgs)
        Dim cb As CheckBox = sender
        If cb.Checked Then
            divQtyControls.Visible = True
            Call PopulateWarehouseDropdown()
            ddlWarehouse.Enabled = True
            ddlWarehouse.Focus()
        Else
            divQtyControls.Visible = False
        End If
    End Sub

    Protected Sub PopulateWarehouseDropdown()
        Dim sSQL As String = "SELECT WarehouseId, WarehouseKey FROM Warehouse WHERE DeletedFlag = 'N' ORDER BY WarehouseId"
        Dim oListItemCollection As ListItemCollection = ExecuteQueryToListItemCollection(sSQL, "WarehouseId", "WarehouseKey")
        ddlWarehouse.Items.Clear()
        ddlWarehouse.Items.Add(New ListItem("- please select -", 0))
        For Each li As ListItem In oListItemCollection
            ddlWarehouse.Items.Add(li)
        Next
        'For i As Int32 = 0 To ddlWarehouse.Items.Count - 1
        '    If ddlWarehouse.Items(i).Text = "DEMO" Then
        '        ddlWarehouse.SelectedIndex = i
        '        Call InitRackDropdown()
        '        Exit For
        '    End If
        'Next
        ddlRack.Items.Clear()
        ddlRack.Items.Add(New ListItem("- please select -", 0))
        ddlRack.SelectedIndex = 0
        ddlRack.Enabled = False
        ddlSection.Items.Clear()
        ddlSection.Items.Add(New ListItem("- please select -", 0))
        ddlSection.SelectedIndex = 0
        ddlSection.Enabled = False
        ddlBay.Items.Clear()
        ddlBay.Items.Add(New ListItem("- please select -", 0))
        ddlBay.SelectedIndex = 0
        ddlBay.Enabled = False
    End Sub

    Protected Function ExecuteQueryToListItemCollection(ByVal sQuery As String, ByVal sTextFieldName As String, ByVal sValueFieldName As String) As ListItemCollection
        Dim oListItemCollection As New ListItemCollection
        Dim oDataReader As SqlDataReader = Nothing
        Dim oConn As New SqlConnection(gsConn)
        Dim sTextField As String
        Dim sValueField As String
        Dim oCmd As SqlCommand = New SqlCommand(sQuery, oConn)
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            If oDataReader.HasRows Then
                While oDataReader.Read
                    If Not IsDBNull(oDataReader(sTextFieldName)) Then
                        sTextField = oDataReader(sTextFieldName)
                    Else
                        sTextField = String.Empty
                    End If
                    If Not IsDBNull(oDataReader(sValueFieldName)) Then
                        sValueField = oDataReader(sValueFieldName)
                    Else
                        sValueField = String.Empty
                    End If
                    oListItemCollection.Add(New ListItem(sTextField, sValueField))
                End While
            End If
        Catch ex As Exception
            WebMsgBox.Show("Error in ExecuteQueryToListItemCollection: " & ex.Message)
        Finally
            oConn.Close()
        End Try
        ExecuteQueryToListItemCollection = oListItemCollection
    End Function

    Protected Sub ddlWarehouse_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ClearRackDropdown()
        ddlRack.Enabled = False
        Call ClearSectionDropdown()
        ddlSection.Enabled = False
        Call ClearBayDropdown()
        ddlBay.Enabled = False

        If ddlWarehouse.SelectedIndex > 0 Then
            Call InitRackDropdown()
            ddlRack.Enabled = True
            ddlRack.Focus()
        End If
    End Sub

    Protected Sub InitRackDropdown()
        ddlRack.Items.Clear()
        Dim sSQL As String = "SELECT WarehouseRackId, WarehouseRackKey FROM WarehouseRack WHERE DeletedFlag = 'N' AND WarehouseKey = " & ddlWarehouse.SelectedValue & " ORDER BY WarehouseRackId"
        Dim oListItemCollection As ListItemCollection = ExecuteQueryToListItemCollection(sSQL, "WarehouseRackId", "WarehouseRackKey")
        ddlRack.Items.Add(New ListItem("- please select -", 0))
        For Each li As ListItem In oListItemCollection
            ddlRack.Items.Add(li)
        Next
    End Sub
    
    Protected Sub ddlRack_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ClearSectionDropdown()
        ddlSection.Enabled = False
        Call ClearBayDropdown()
        ddlBay.Enabled = False

        If ddlRack.SelectedIndex > 0 Then
            Call InitSectionDropdown()
            ddlSection.Enabled = True
            ddlSection.Focus()
        End If
    End Sub

    Protected Sub InitSectionDropdown()
        ddlSection.Items.Clear()
        Dim sSQL As String = "SELECT WarehouseSectionId, WarehouseSectionKey FROM WarehouseSection WHERE DeletedFlag = 'N' AND WarehouseRackKey = " & ddlRack.SelectedValue & " ORDER BY WarehouseSectionId"
        Dim oListItemCollection As ListItemCollection = ExecuteQueryToListItemCollection(sSQL, "WarehouseSectionId", "WarehouseSectionKey")
        ddlSection.Items.Add(New ListItem("- please select -", 0))
        For Each li As ListItem In oListItemCollection
            ddlSection.Items.Add(li)
        Next
    End Sub
    
    Protected Sub ddlSection_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ClearBayDropdown()
        ddlBay.Enabled = False

        If ddlSection.SelectedIndex > 0 Then
            Call InitBayDropdown()
            ddlBay.Enabled = True
            ddlSection.Focus()
        End If
    End Sub
    
    Protected Sub InitBayDropdown()
        ddlBay.Items.Clear()
        Dim sSQL As String = "SELECT WarehouseBayId, WarehouseBayKey FROM WarehouseBay WHERE DeletedFlag = 'N' AND WarehouseSectionKey = " & ddlSection.SelectedValue & " ORDER BY WarehouseBayId"
        Dim oListItemCollection As ListItemCollection = ExecuteQueryToListItemCollection(sSQL, "WarehouseBayId", "WarehouseBayKey")
        ddlBay.Items.Add(New ListItem("- please select -", 0))
        For Each li As ListItem In oListItemCollection
            ddlBay.Items.Add(li)
        Next
    End Sub

    Protected Sub ClearRackDropdown()
        If ddlRack.Items.Count > 0 Then
            ddlRack.SelectedIndex = 0
        End If
    End Sub
    
    Protected Sub ClearSectionDropdown()
        If ddlSection.Items.Count > 0 Then
            ddlSection.SelectedIndex = 0
        End If
    End Sub
    
    Protected Sub ClearBayDropdown()
        If ddlBay.Items.Count > 0 Then
            ddlBay.SelectedIndex = 0
        End If
        'lblBayKey.Text = String.Empty
    End Sub
    
    Protected Sub ddlBay_SelectedIndexChanged(sender As Object, e As System.EventArgs)
        If ddlBay.SelectedIndex = 0 Then
            'lblBayKey.Text = String.Empty
        Else
            'lblBayKey.Text = "(Bay key: " & ddlBay.SelectedValue & ")"
        End If
    End Sub

</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Upload Products</title>
    <style type="text/css">
        .style1
        {
            color: #FF0000;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server" defaultfocus="FormUpload1" enctype="multipart/form-data">
    <main:header ID="ctlHeader" runat="server" />
    <strong>
    <asp:Label ID="lblLegendUploadProducts" runat="server" Text="UPLOAD PRODUCTS "></asp:Label>
    <asp:Label ID="lblCustomer" runat="server"
        Text="(no customer selected)" Font-Names="Verdana" />
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
        <asp:Button ID="btnHelp" runat="server" Text="show help" Width="112px" OnClick="btnHelp_Click" />
        &nbsp; &nbsp; &nbsp;&nbsp;
        <asp:LinkButton ID="lnkbtnRestart" runat="server" OnClick="lnkbtnRestart_Click">restart</asp:LinkButton><br />
    </strong>
    <br />
    <asp:Panel ID="pnlHelp" runat="server" Width="100%" Font-Names="Verdana">
        The <b>Upload Products</b> facility loads products from a CSV (comma-separated 
        variable) format file you supply. Follow the steps below to specify (a) the 
        location of the file and (b) how to use each field of the product.<br />
        <br />
        1. To import products from an Excel spreadsheet you must first convert the 
        spreadsheet into CSV format. Open the spreadsheet, choose <b>Save As</b>, select
        <b>CSV</b> as the file type, specify a location and filename (eg 
        C:\myproducts.csv), then click <b>Save</b>.<br />
        <br />
        2. Click the <strong>Browse</strong> button to locate the file of products on your
        local machine.<br />
        <br />
        3. Select the customer for whom the products are to be uploaded. Check this 
        carefully - <span class="style1"><b>YOU MUST GET IT RIGHT!!</b></span><br />
        <br />
        4. If your data contains columns headings (ie the first row of data is the name 
        or description of the column) click the <strong>Column Headings in Row 1</strong> 
        check box.<br />
        <br />
        5. If one of your data columns is an initial stock quantity to be set up, click 
        the <b>Add quantity</b> check box. A new line will appear to allow you to select 
        the warehouse location of the products. All products must go into the same 
        location. If any of your quantities are negative, and you want to ignore these 
        values, click the <b>ignore -ve qtys</b> check box.<br />
        <br />
        6. If you are using authorisation, you can select a the Default Authoriser for 
        the products from the drop down box. All authorisable products must have the 
        same default authoriser.<br />
        <br />
        7. Click the <strong>Read Products </strong>button. The system reads and 
        interprets your file, then displays the contents for you to check and confirm 
        before loading it into the system. A message is displayed if there are problems 
        reading the data.<br />
        <br />
        8. For each column of data you want to include, choose the field to associate 
        with this column, using the dropdown list box at the top of each column. If you 
        have a column containing an initial quantity to be added, select the value <b>
        Quantity</b> in the drop down list.<br />
        <br />
        9. Click <strong>Check Data</strong> to check that the system can correctly 
        process your data. The system displays a message if the product already exists, 
        validation failes, or required data is missing. Correct any errors and re-submit 
        the data.<br />
        <br />
        10. Click <strong>Upload Products </strong>to load your data into the system.<br />
        <br />
    </asp:Panel>
    <asp:Panel ID="pnlCommon" runat="server" Width="100%">
        <asp:Label ID="Label1" runat="server" Text="Customer:" Font-Names="Verdana" Font-Size="XX-Small"></asp:Label>
        <asp:DropDownList ID="ddlCustomers" runat="server" OnSelectedIndexChanged="ddlCustomers_SelectedIndexChanged"
            AutoPostBack="True" Font-Names="Verdana" Font-Size="XX-Small" />
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
        <asp:Label ID="lbLegendlDefaultAuthoriser" runat="server" Enabled="False" Text="Default Authoriser:"
            Font-Names="Verdana" Font-Size="XX-Small"></asp:Label>
        <asp:DropDownList ID="ddlAuthoriser" runat="server" Enabled="False" OnSelectedIndexChanged="ddlAuthoriser_SelectedIndexChanged"
            Font-Names="Verdana" Font-Size="XX-Small">
        </asp:DropDownList>
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        <asp:CheckBox ID="cbAddQuantity" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
            Text="Add quantity" oncheckedchanged="cbAddQuantity_CheckedChanged" 
            AutoPostBack="True" />
        <div id="divQtyControls" runat="server" visible="false">
            <asp:Label ID="Label4" runat="server" Font-Names="Verdana" Font-Size="XX-Small" 
                Text="Warehouse:" />
            &nbsp;<asp:DropDownList ID="ddlWarehouse" runat="server" AutoPostBack="True" 
                Font-Names="Verdana" Font-Size="XX-Small" 
                onselectedindexchanged="ddlWarehouse_SelectedIndexChanged" />
            &nbsp;
            <asp:Label ID="Label5" runat="server" Font-Names="Verdana" 
                Font-Size="XX-Small" Text="Rack:" />
&nbsp;<asp:DropDownList ID="ddlRack" runat="server" AutoPostBack="True" 
                Font-Names="Verdana" Font-Size="XX-Small" 
                onselectedindexchanged="ddlRack_SelectedIndexChanged" />
            &nbsp;
            <asp:Label ID="Label6" runat="server" Font-Names="Verdana" Font-Size="XX-Small" 
                Text="Section:" />
&nbsp;<asp:DropDownList ID="ddlSection" runat="server" AutoPostBack="True" 
                Font-Names="Verdana" Font-Size="XX-Small" 
                onselectedindexchanged="ddlSection_SelectedIndexChanged" />
            &nbsp;&nbsp;<asp:Label ID="Label7" runat="server" Font-Names="Verdana" 
                Font-Size="XX-Small" Text="Bay:" />
&nbsp;<asp:DropDownList ID="ddlBay" runat="server" AutoPostBack="True" Font-Names="Verdana" 
                Font-Size="XX-Small" onselectedindexchanged="ddlBay_SelectedIndexChanged" />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:CheckBox ID="cbIgnoreNegativeQuantities" runat="server" Font-Names="Verdana"
                Font-Size="XX-Small" Text="ignore -ve qtys" />
        </div>
    </asp:Panel>
    <asp:Panel ID="pnlStart" runat="server" Width="100%">
        <asp:Label ID="lblLegendProductCSVFile" runat="server" Text="Product CSV file:" Font-Names="Verdana"
            Font-Size="XX-Small" />
        <asp:FileUpload ID="FileUpload1" runat="server" Width="350px" Enabled="False" Font-Names="Verdana"
            Font-Size="XX-Small" />
        &nbsp;&nbsp;<asp:CheckBox ID="cbColumnHeadingsInRow1" runat="server" Text="column&nbsp;headings&nbsp;in&nbsp;row&nbsp;1"
            Enabled="False" Font-Names="Verdana" Font-Size="XX-Small" />&nbsp;
        <asp:Button ID="btnReadProducts" runat="server" OnClick="btnReadProducts_Click" Text="read products"
            Width="140px" Enabled="False" Font-Names="Verdana" Font-Size="XX-Small" />&nbsp;<br />
    </asp:Panel>
    <asp:Panel ID="pnlGridView" runat="server" Width="100%">
        <br />
        <asp:Button ID="btnCheckData" runat="server" Text="check data" OnClick="btnCheckData_Click"
            Enabled="False" />
        <asp:Button ID="btnUploadProducts" runat="server" Text="upload products" OnClick="btnUploadProducts_Click"
            Enabled="False" /><br />
        <br />
        <asp:GridView ID="gvUploadData" runat="server" Width="100%" Font-Names="Verdana"
            Font-Size="XX-Small">
        </asp:GridView>
    </asp:Panel>
    <asp:Panel ID="pnlMessage" runat="server" Width="100%">
        <asp:Label ID="lblMessage" runat="server" Font-Names="Verdana" Font-Size="XX-Small" /><br />
        <br />
        <br />
        &nbsp;</asp:Panel>
    <asp:Panel ID="pnlFinished" runat="server" Width="100%">
        <asp:Label ID="Label3" runat="server" Text="Results:" Font-Names="Verdana" Font-Size="XX-Small" />
        <asp:TextBox ID="tbResults" runat="server" Rows="10" TextMode="MultiLine" Width="100%"
            Font-Names="Verdana" Font-Size="XX-Small" />
        <br />
        <asp:Label ID="lblSuccess" runat="server" ForeColor="Green" Font-Names="Verdana"
            Font-Size="XX-Small"></asp:Label>
        <br />
        <asp:Button ID="btnClose" runat="server" Text="Close" OnClientClick="window.close()" />
        &nbsp;&nbsp;
    </asp:Panel>
    <br />
    <asp:Label ID="lblError" runat="server" ForeColor="Red" Font-Names="Verdana" Font-Size="XX-Small"></asp:Label>
    <asp:HiddenField ID="HiddenField0" runat="server" Value="0" />
    <asp:HiddenField ID="HiddenField1" runat="server" Value="0" />
    <asp:HiddenField ID="HiddenField2" runat="server" Value="0" />
    <asp:HiddenField ID="HiddenField3" runat="server" Value="0" />
    <asp:HiddenField ID="HiddenField4" runat="server" Value="0" />
    <asp:HiddenField ID="HiddenField5" runat="server" Value="0" />
    <asp:HiddenField ID="HiddenField6" runat="server" Value="0" />
    <asp:HiddenField ID="HiddenField7" runat="server" Value="0" />
    <asp:HiddenField ID="HiddenField8" runat="server" Value="0" />
    <asp:HiddenField ID="HiddenField9" runat="server" Value="0" />
    <asp:HiddenField ID="HiddenField10" runat="server" Value="0" />
    <asp:HiddenField ID="HiddenField11" runat="server" Value="0" />
    <asp:HiddenField ID="HiddenField12" runat="server" Value="0" />
    <asp:HiddenField ID="HiddenField13" runat="server" Value="0" />
    <asp:HiddenField ID="HiddenField14" runat="server" Value="0" />
    <asp:HiddenField ID="HiddenField15" runat="server" Value="0" />
    <asp:HiddenField ID="HiddenField16" runat="server" Value="0" />
    <asp:HiddenField ID="HiddenField17" runat="server" Value="0" />
    <asp:HiddenField ID="HiddenField18" runat="server" Value="0" />
    <asp:HiddenField ID="HiddenField19" runat="server" Value="0" />
    <asp:HiddenField ID="HiddenField20" runat="server" Value="0" />
    <asp:HiddenField ID="HiddenField21" runat="server" Value="0" />
    <asp:HiddenField ID="HiddenField22" runat="server" Value="0" />
    <asp:HiddenField ID="HiddenField23" runat="server" Value="0" />
    <asp:HiddenField ID="HiddenField24" runat="server" Value="0" />
    <asp:HiddenField ID="HiddenField25" runat="server" Value="0" />
    <asp:HiddenField ID="HiddenField26" runat="server" Value="0" />
    <asp:HiddenField ID="HiddenField27" runat="server" Value="0" />
    <asp:HiddenField ID="HiddenField28" runat="server" Value="0" />
    <asp:HiddenField ID="HiddenField29" runat="server" Value="0" />
    <asp:HiddenField ID="HiddenField30" runat="server" Value="0" />
    <asp:HiddenField ID="HiddenFieldChanged" runat="server" />
    </form>
</body>
</html>