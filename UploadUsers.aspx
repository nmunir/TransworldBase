<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

    ' STILL TO DO
    
    ' sort out autogeneration of UserID
    ' validate email address
    ' add missing fields (Title, Department, Telephone, etc.)
    ' add count of users supplied, users loaded
    ' option to view input file, to see headers
    
    Const MARSHALL_USER_ID As Integer = 0
    Const MARSHALL_PASSWORD As Integer = 1
    Const MARSHALL_FIRST_NAME As Integer = 2
    Const MARSHALL_LAST_NAME As Integer = 3
    Const MARSHALL_TITLE As Integer = 4
    Const MARSHALL_DEPARTMENT As Integer = 5
    Const MARSHALL_USER_TYPE As Integer = 6
    Const MARSHALL_EMAIL_ADDRESS As Integer = 7
    Const MARSHALL_TELEPHONE As Integer = 8
    Const MARSHALL_COLLECTION_POINT As Integer = 9
    Const MARSHALL_ABLE_TO_VIEW_STOCK As Integer = 10
    Const MARSHALL_ABLE_TO_CREATE_STOCK_BOOKING As Integer = 11
    Const MARSHALL_ABLE_TO_CREATE_COLLECTION_REQUEST As Integer = 12
    Const MARSHALL_ABLE_TO_VIEW_GLOBAL_ADDRESS_BOOK As Integer = 13
    Const MARSHALL_ABLE_TO_EDIT_GLOBAL_ADDRESS_BOOK As Integer = 14
    Const MARSHALL_RECEIVE_OWN_STOCK_BOOKING_ALERT As Integer = 15
    Const MARSHALL_RECEIVE_ALL_STOCK_BOOKING_ALERTS As Integer = 16
    Const MARSHALL_RECEIVE_GOODS_IN_ALERTS As Integer = 17
    Const MARSHALL_RECEIVE_LOW_STOCK_ALERTS As Integer = 18
    Const MARSHALL_RECEIVE_OWN_CONSIGNMENT_BOOKING_ALERT As Integer = 19
    Const MARSHALL_RECEIVE_ALL_CONSIGNMENT_BOOKING_ALERTS As Integer = 20
    Const MARSHALL_RECEIVE_CONSIGNMENT_DESPATCH_ALERTS As Integer = 21
    Const MARSHALL_RECEIVE_CONSIGNMENT_DELIVERY_ALERTS As Integer = 22

    Const DROPDOWN_ITEM_00_NOTHING_SELECTED As Int32 = 0
    Const DROPDOWN_ITEM_01_DONT_USE_THIS_COLUMN As Int32 = 1
    Const DROPDOWN_ITEM_02_USERID As Int32 = 2
    Const DROPDOWN_ITEM_03_PASSWORD As Int32 = 3
    Const DROPDOWN_ITEM_04_FIRSTNAME As Int32 = 4
    Const DROPDOWN_ITEM_05_LASTNAME As Int32 = 5
    Const DROPDOWN_ITEM_06_TITLE As Int32 = 6
    Const DROPDOWN_ITEM_07_DEPARTMENT As Int32 = 7
    Const DROPDOWN_ITEM_08_USERTYPE As Int32 = 8
    Const DROPDOWN_ITEM_09_EMAILADDR As Int32 = 9
    Const DROPDOWN_ITEM_10_TELEPHONE As Int32 = 10
    Const DROPDOWN_ITEM_11_COLLECTIONPOINT As Int32 = 11
    Const DROPDOWN_ITEM_12_ABLETOVIEWSTOCK As Int32 = 12

    Dim sFileName As String, sIntermediateFileName As String
    Dim sFilePrefix As String, sFileSuffix As String = ".csv"
    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString
  
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsPostBack Then
            Call Initialise()
        End If
    End Sub
  
    Protected Sub Initialise()
        pnlStart.Visible = True
        pnlGridView.Visible = False
        pnlMessage.Visible = False
        pnlHelp.Visible = False
        Call GetCustomerAccountCodes()
    End Sub
    
    Protected Sub btnReadUsers_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ReadUsers()
    End Sub
  
    Protected Sub ReadUsers()
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
                WebMsgBox.Show("WARNING: The file you are uploading contains " & nLineCount.ToString & " users.  To ensure successful uploading we recommend that you split the file into multiple files of no more than 10,000 users per file.")
            Else
                sIntermediateFileName = Server.MapPath("") & "\" & sFilePrefix & "-2" & sFileSuffix
                Call RemoveEmbeddedLineBreaksAndCommas(sFileName, sIntermediateFileName)
                Dim dt As DataTable
                dt = DelimitFile(sIntermediateFileName, ",", cbColumnHeadingsInRow1.Checked)
                If dt.Rows.Count > 0 Then
                    Dim r As DataRow = dt.NewRow()
                    dt.Rows.InsertAt(r, 0)
                    gvUserDetails.DataSource = dt
                    gvUserDetails.DataBind()
                End If
                If gvUserDetails.Rows.Count > 0 Then
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
            If RequiredFieldsPresent() Then
                lblMessage.Text = ""
                lblSuccess.Text = ""
                Call ProcessUsers(bAddToDatabase:=False)
                If lblError.Text = "" Then
                    lblSuccess.Text = "No errors found"
                End If
            End If
        Else
            WebMsgBox.Show("Two or more columns are mapped to the same field.")
        End If
    End Sub

    Protected Sub btnUploadUsers_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        tbResults.Text = String.Empty
        If SelectedFieldsUnique() Then
            If RequiredFieldsPresent() Then
                lblMessage.Text = ""
                lblSuccess.Text = ""
                Call ProcessUsers(bAddToDatabase:=False)
                If lblError.Text = "" Then
                    Call ProcessUsers(bAddToDatabase:=True)
                End If
                If lblError.Text = "" Then
                    lblMessage.Text = gvUserDetails.Rows.Count.ToString & " users added"
                    pnlGridView.Visible = False
                End If
                If lblError.Text = "" Then
                    lblSuccess.Text = "Users successfully loaded"
                End If
            End If
        Else
            WebMsgBox.Show("Two or more columns are mapped to the same field.")
        End If
    End Sub
  
    Protected Function SelectedFieldsUnique() As Boolean
        ' DO MSG BOX IN HERE; ALSO CHECK COUNTRY SELECTED; ALSO CHECK SOMETHING SELECTED
        Dim nCol As Integer
        Dim nMaxCols As Integer = gvUserDetails.Rows(0).Cells.Count
        Dim nTargetIndex As Integer
        Dim nSelectedField(22) As Int16
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
  
    Protected Function RequiredFieldsPresent() As Boolean
        Dim nCol As Integer
        Dim nMaxCols As Integer = gvUserDetails.Rows(0).Cells.Count
        Dim nTargetIndex As Integer
        Dim nRequiredField(22) As Int16
        RequiredFieldsPresent = True
        For nCol = 0 To nMaxCols - 1
            nTargetIndex = ExtractHiddenFieldValue(nCol)
            If nTargetIndex >= 2 Then     ' allow multiple 'nothing selected' or 'don't use this column'
                nRequiredField(nTargetIndex) += 1
            End If
        Next
        If nRequiredField(DROPDOWN_ITEM_02_USERID) = 0 And Not cbAutoGenerateUserId.Checked Then
            RequiredFieldsPresent = False
            WebMsgBox.Show("No User ID column is selected, and autogenerate User ID is not checked.")
            Exit Function
        End If
        If nRequiredField(DROPDOWN_ITEM_04_FIRSTNAME) = 0 Then
            RequiredFieldsPresent = False
            WebMsgBox.Show("No First Name column is selected.")
            Exit Function
        End If
        If nRequiredField(DROPDOWN_ITEM_05_LASTNAME) = 0 Then
            RequiredFieldsPresent = False
            WebMsgBox.Show("No Last Name column is selected.")
            Exit Function
        End If
        If nRequiredField(DROPDOWN_ITEM_09_EMAILADDR) = 0 Then
            RequiredFieldsPresent = False
            WebMsgBox.Show("No Email Address column is selected.")
            Exit Function
        End If
        If nRequiredField(DROPDOWN_ITEM_08_USERTYPE) = 0 Then
            WebMsgBox.Show("No User Type column is selected. Created accounts will default to type User.")
        End If
    End Function
  
    Protected Sub AddDropDowns()
        Dim gvr As GridViewRow
        gvr = gvUserDetails.Rows(0)
        Dim j As Integer
        For j = 0 To gvr.Cells.Count - 1
            Dim dd As New DropDownList
            dd.ID = "Select" & j.ToString.Trim
            'dd.AutoPostBack = True
            ' following list must match the array used to match selection during processing          
            dd.Items.Add("- nothing selected -")
            dd.Items.Add("- don't use this column -")
            dd.Items.Add("User ID")
            dd.Items.Add("Password")
            dd.Items.Add("First Name")
            dd.Items.Add("Last Name")
            dd.Items.Add("Title")
            dd.Items.Add("Department / CC")
            dd.Items.Add("User Type")
            dd.Items.Add("Email Address")
            dd.Items.Add("Telephone")
            dd.Items.Add("Collection Point")
            dd.Items.Add("Able To View Stock")
            dd.Items.Add("Able To Create Stock Booking")
            dd.Items.Add("Able To Create Collection Request")
            dd.Items.Add("Able To View Global Address Book")
            dd.Items.Add("Able To Edit Global Address Book")
            dd.Items.Add("Receive Own Stock Booking Alerts")
            dd.Items.Add("Receive All Stock Booking Alerts")
            dd.Items.Add("Receive Goods In Alerts")
            dd.Items.Add("Receive Low Stock Alerts")
            dd.Items.Add("Receive Own Consignment Booking Alerts")
            dd.Items.Add("Receive All Consignment Booking Alerts")
            dd.Items.Add("Receive Consignment Despatch Alerts")
            dd.Items.Add("Receive Consignment Delivery Alerts")
            dd.SelectedIndex = ExtractHiddenFieldValue(j)

            gvr.Cells(j).Controls.Add(dd)
            Dim cid As String = dd.ClientID
            dd.Attributes.Add("onClick", "HiddenField" & j.ToString.Trim & ".value=" & cid & ".selectedIndex; HiddenFieldChanged.value='TRUE'")
        Next
    End Sub

    Protected Sub ProcessUsers(ByVal bAddToDatabase As Boolean)
        Dim nRow As Integer, nMaxRows As Integer = gvUserDetails.Rows.Count
        Dim nCol As Integer, nMaxCols As Integer = gvUserDetails.Columns.Count
        Dim i As Integer, sVal As String
        Dim nTargetIndex As Integer
        ' following 2 lists must match the contents of the dropdown list box used to select target (sMarshaller drops initial 2 items)
        Dim sTarget() As String = {"", "", "UserID", "Password", "FirstName", "LastName", "Title", "Department", "UserType", "EmailAddress", "Telephone", "CollectionPoint", "AbleToViewStock", "AbleToCreateStockBooking", "AbleToCreateCollectionRequest", "AbleToViewGlobalAddressBook", "AbleToEditGlobalAddressBook", "ReceiveOwnStockBookingAlert", "ReceiveAllStockBookingAlerts", "ReceiveGoodsInAlerts", "ReceiveLowStockAlerts", "ReceiveOwnConsignmentBookingAlert", "ReceiveAllConsignmentBookingAlerts", "ReceiveConsignmentDespatchAlerts", "ReceiveConsignmentDeliveryAlerts"}
        Dim sMarshaller() As String = {"UserID", "Password", "FirstName", "LastName", "Title", "Department", "UserType", "EmailAddress", "Telephone", "CollectionPoint", "AbleToViewStock", "AbleToCreateStockBooking", "AbleToCreateCollectionRequest", "AbleToViewGlobalAddressBook", "AbleToEditGlobalAddressBook", "ReceiveOwnStockBookingAlert", "ReceiveAllStockBookingAlerts", "ReceiveGoodsInAlerts", "ReceiveLowStockAlerts", "ReceiveOwnConsignmentBookingAlert", "ReceiveAllConsignmentBookingAlerts", "ReceiveConsignmentDespatchAlerts", "ReceiveConsignmentDeliveryAlerts"}
        Dim nValidationIndex As Integer
        
        Dim sbErrors As New StringBuilder
        lblError.Text = ""
      
        nMaxRows = gvUserDetails.Rows.Count
        nMaxCols = gvUserDetails.Rows(0).Cells.Count
      
        If cbAutoGenerateUserId.Checked Then
            tbUserIdSeparatorChars.Text = tbUserIdSeparatorChars.Text.Trim
            tbUserIdSeparatorChars.Text = tbUserIdSeparatorChars.Text.Replace(" ", "_")
            For nRow = 1 To nMaxRows - 1
                Dim sFirstName As String = String.Empty
                Dim sLastName As String = String.Empty
                Dim sUserId As String = String.Empty
                Dim nUserIdCol As Integer = 0
                For nCol = 0 To nMaxCols - 1
                    nTargetIndex = ExtractHiddenFieldValue(nCol)
                    Select Case nTargetIndex
                        Case DROPDOWN_ITEM_02_USERID
                            sUserId = HttpUtility.HtmlDecode(gvUserDetails.Rows(nRow).Cells(nCol).Text)
                            nUserIdCol = nCol
                        Case DROPDOWN_ITEM_04_FIRSTNAME
                            sFirstName = HttpUtility.HtmlDecode(gvUserDetails.Rows(nRow).Cells(nCol).Text).Trim.ToLower
                        Case DROPDOWN_ITEM_05_LASTNAME
                            sLastName = HttpUtility.HtmlDecode(gvUserDetails.Rows(nRow).Cells(nCol).Text).Trim.ToLower.Replace(" ", "-")
                    End Select
                Next
                
                ' 27MAY13 doesn't appear to handle case where no UserID column is specified at all; next line should perhaps be: If sUserId.Trim.ToLower = "autogenerate" or sUserId = String.Empty
                
                If sUserId.Trim.ToLower = "autogenerate" Then
                    If sFirstName <> String.Empty And sLastName <> String.Empty Then
                        sUserId = sFirstName.Substring(0, 1) & tbUserIdSeparatorChars.Text & sLastName
                        gvUserDetails.Rows(nRow).Cells(nUserIdCol).Text = HttpUtility.HtmlEncode(sUserId)
                    End If
                End If
            Next
        End If

        For nRow = 1 To nMaxRows - 1
            For i = sMarshaller.GetLowerBound(0) To sMarshaller.GetUpperBound(0)
                sMarshaller(i) = ""
            Next
            For nCol = 0 To nMaxCols - 1
                nTargetIndex = ExtractHiddenFieldValue(nCol)
                sVal = HttpUtility.HtmlDecode(gvUserDetails.Rows(nRow).Cells(nCol).Text)
                If nTargetIndex > 1 Then                                      ' need to change this in address & product prog too
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
                If sMarshaller(MARSHALL_USER_ID) = String.Empty Then
                    sbErrors.Append("No User ID specified in row " & nRow.ToString & "<br />")
                Else
                    If sMarshaller(MARSHALL_USER_ID).Length > 100 Then
                        sbErrors.Append("User ID " & sMarshaller(MARSHALL_USER_ID) & " exceeds maximum User ID length of 100 chars in row " & nRow.ToString & ".<br />")
                    End If
                    If ExecuteQueryToDataTable("SELECT UserId from UserProfile WHERE UserId = '" & sMarshaller(MARSHALL_USER_ID).Replace("'", "''") & "'").Rows.Count > 0 Then
                        sbErrors.Append("User ID " & sMarshaller(MARSHALL_USER_ID) & " in row " & nRow.ToString & " already exists on the system!!<br />")
                    End If
                End If
          
                If sMarshaller(MARSHALL_FIRST_NAME) = "" Then
                    sbErrors.Append("No first name specified in row " & nRow.ToString & "<br />")
                Else
                    If sMarshaller(MARSHALL_FIRST_NAME).Length > 50 Then
                        sbErrors.Append("User ID " & sMarshaller(MARSHALL_USER_ID) & " exceeds maximum First Name length of 50 chars in row " & nRow.ToString & ".<br />")
                    End If
                End If

                If sMarshaller(MARSHALL_LAST_NAME) = "" Then
                    sbErrors.Append("No last name specified in row " & nRow.ToString & "<br />")
                Else
                    If sMarshaller(MARSHALL_LAST_NAME).Length > 50 Then
                        sbErrors.Append("User ID " & sMarshaller(MARSHALL_LAST_NAME) & " exceeds maximum Last Name length of 50 chars in row " & nRow.ToString & ".<br />")
                    End If
                End If

                If sMarshaller(MARSHALL_EMAIL_ADDRESS) = "" Then
                    sbErrors.Append("No email address specified in row " & nRow.ToString & "<br />")
                Else
                    If sMarshaller(MARSHALL_EMAIL_ADDRESS).Length > 100 Then
                        sbErrors.Append("User ID " & sMarshaller(MARSHALL_EMAIL_ADDRESS) & " exceeds maximum Email Address length of 100 chars in row " & nRow.ToString & ".<br />")
                    End If
                End If

                nValidationIndex = MARSHALL_USER_TYPE
                If sMarshaller(nValidationIndex) = "" Then
                    sbErrors.Append("No user type specified in row " & nRow.ToString & ". User type must start with U for User, P for Product Owner or S for SuperUser (case insensitive).<br />")
                Else
                    If Not (sMarshaller(nValidationIndex).ToLower.StartsWith("u") Or sMarshaller(nValidationIndex).ToLower.StartsWith("p") Or sMarshaller(nValidationIndex).ToLower.StartsWith("s")) Then
                        sbErrors.Append("Unrecognised user type " & sMarshaller(nValidationIndex) & " specified in row " & nRow.ToString & ". User type must start with U for User, P for Product Owner or S for SuperUser (case insensitive).<br />")
                    End If
                End If
                
                nValidationIndex = MARSHALL_ABLE_TO_VIEW_STOCK
                If sMarshaller(nValidationIndex) <> String.Empty Then
                    If Not (sMarshaller(nValidationIndex).ToUpper = "Y" Or sMarshaller(nValidationIndex).ToUpper = "YES" Or sMarshaller(nValidationIndex).ToUpper = "1" _
                      Or sMarshaller(nValidationIndex).ToUpper = "N" Or sMarshaller(nValidationIndex).ToUpper = "NO" Or sMarshaller(nValidationIndex).ToUpper = "0" _
                      Or sMarshaller(nValidationIndex).ToUpper = "TRUE" Or sMarshaller(nValidationIndex).ToUpper = "FALSE") _
                    Then
                        sbErrors.Append("Cannot interpret Able To View Stock Booking value specified in row " & nRow.ToString & ". Expecting Y, YES, N, NO, 1, 0, TRUE or FALSE, case insensitive.<br />")
                    End If
                End If

                nValidationIndex = MARSHALL_ABLE_TO_CREATE_STOCK_BOOKING
                If sMarshaller(nValidationIndex) <> String.Empty Then
                    If Not (sMarshaller(nValidationIndex).ToUpper = "Y" Or sMarshaller(nValidationIndex).ToUpper = "YES" Or sMarshaller(nValidationIndex).ToUpper = "1" _
                      Or sMarshaller(nValidationIndex).ToUpper = "N" Or sMarshaller(nValidationIndex).ToUpper = "NO" Or sMarshaller(nValidationIndex).ToUpper = "0" _
                      Or sMarshaller(nValidationIndex).ToUpper = "TRUE" Or sMarshaller(nValidationIndex).ToUpper = "FALSE") _
                    Then
                        sbErrors.Append("Cannot interpret Able To Create Stock Booking value specified in row " & nRow.ToString & ". Expecting Y, YES, N, NO, 1, 0, TRUE or FALSE, case insensitive.<br />")
                    End If
                End If

                nValidationIndex = MARSHALL_ABLE_TO_CREATE_COLLECTION_REQUEST
                If sMarshaller(nValidationIndex) <> String.Empty Then
                    If Not (sMarshaller(nValidationIndex).ToUpper = "Y" Or sMarshaller(nValidationIndex).ToUpper = "YES" Or sMarshaller(nValidationIndex).ToUpper = "1" _
                      Or sMarshaller(nValidationIndex).ToUpper = "N" Or sMarshaller(nValidationIndex).ToUpper = "NO" Or sMarshaller(nValidationIndex).ToUpper = "0" _
                      Or sMarshaller(nValidationIndex).ToUpper = "TRUE" Or sMarshaller(nValidationIndex).ToUpper = "FALSE") _
                    Then
                        sbErrors.Append("Cannot interpret Able To Create Collection Request value specified in row " & nRow.ToString & ". Expecting Y, YES, N, NO, 1, 0, TRUE or FALSE, case insensitive.<br />")
                    End If
                End If

                nValidationIndex = MARSHALL_ABLE_TO_VIEW_GLOBAL_ADDRESS_BOOK
                If sMarshaller(nValidationIndex) <> String.Empty Then
                    If Not (sMarshaller(nValidationIndex).ToUpper = "Y" Or sMarshaller(nValidationIndex).ToUpper = "YES" Or sMarshaller(nValidationIndex).ToUpper = "1" _
                      Or sMarshaller(nValidationIndex).ToUpper = "N" Or sMarshaller(nValidationIndex).ToUpper = "NO" Or sMarshaller(nValidationIndex).ToUpper = "0" _
                      Or sMarshaller(nValidationIndex).ToUpper = "TRUE" Or sMarshaller(nValidationIndex).ToUpper = "FALSE") _
                    Then
                        sbErrors.Append("Cannot interpret Able To View Global Address Book value specified in row " & nRow.ToString & ". Expecting Y, YES, N, NO, 1, 0, TRUE or FALSE, case insensitive.<br />")
                    End If
                End If

                nValidationIndex = MARSHALL_ABLE_TO_EDIT_GLOBAL_ADDRESS_BOOK
                If sMarshaller(nValidationIndex) <> String.Empty Then
                    If Not (sMarshaller(nValidationIndex).ToUpper = "Y" Or sMarshaller(nValidationIndex).ToUpper = "YES" Or sMarshaller(nValidationIndex).ToUpper = "1" _
                      Or sMarshaller(nValidationIndex).ToUpper = "N" Or sMarshaller(nValidationIndex).ToUpper = "NO" Or sMarshaller(nValidationIndex).ToUpper = "0" _
                      Or sMarshaller(nValidationIndex).ToUpper = "TRUE" Or sMarshaller(nValidationIndex).ToUpper = "FALSE") _
                    Then
                        sbErrors.Append("Cannot interpret Able To Edit Global Address Book value specified in row " & nRow.ToString & ". Expecting Y, YES, N, NO, 1, 0, TRUE or FALSE, case insensitive.<br />")
                    End If
                End If

                nValidationIndex = MARSHALL_RECEIVE_OWN_STOCK_BOOKING_ALERT
                If sMarshaller(nValidationIndex) <> String.Empty Then
                    If Not (sMarshaller(nValidationIndex).ToUpper = "Y" Or sMarshaller(nValidationIndex).ToUpper = "YES" Or sMarshaller(nValidationIndex).ToUpper = "1" _
                      Or sMarshaller(nValidationIndex).ToUpper = "N" Or sMarshaller(nValidationIndex).ToUpper = "NO" Or sMarshaller(nValidationIndex).ToUpper = "0" _
                      Or sMarshaller(nValidationIndex).ToUpper = "TRUE" Or sMarshaller(nValidationIndex).ToUpper = "FALSE") _
                    Then
                        sbErrors.Append("Cannot interpret Receive Own Stock Booking Alert value specified in row " & nRow.ToString & ". Expecting Y, YES, N, NO, 1, 0, TRUE or FALSE, case insensitive.<br />")
                    End If
                End If

                nValidationIndex = MARSHALL_RECEIVE_ALL_STOCK_BOOKING_ALERTS
                If sMarshaller(nValidationIndex) <> String.Empty Then
                    If Not (sMarshaller(nValidationIndex).ToUpper = "Y" Or sMarshaller(nValidationIndex).ToUpper = "YES" Or sMarshaller(nValidationIndex).ToUpper = "1" _
                      Or sMarshaller(nValidationIndex).ToUpper = "N" Or sMarshaller(nValidationIndex).ToUpper = "NO" Or sMarshaller(nValidationIndex).ToUpper = "0" _
                      Or sMarshaller(nValidationIndex).ToUpper = "TRUE" Or sMarshaller(nValidationIndex).ToUpper = "FALSE") _
                    Then
                        sbErrors.Append("Cannot interpret Receive All Stock Booking Alerts value specified in row " & nRow.ToString & ". Expecting Y, YES, N, NO, 1, 0, TRUE or FALSE, case insensitive.<br />")
                    End If
                End If

                nValidationIndex = MARSHALL_RECEIVE_GOODS_IN_ALERTS
                If sMarshaller(nValidationIndex) <> String.Empty Then
                    If Not (sMarshaller(nValidationIndex).ToUpper = "Y" Or sMarshaller(nValidationIndex).ToUpper = "YES" Or sMarshaller(nValidationIndex).ToUpper = "1" _
                      Or sMarshaller(nValidationIndex).ToUpper = "N" Or sMarshaller(nValidationIndex).ToUpper = "NO" Or sMarshaller(nValidationIndex).ToUpper = "0" _
                      Or sMarshaller(nValidationIndex).ToUpper = "TRUE" Or sMarshaller(nValidationIndex).ToUpper = "FALSE") _
                    Then
                        sbErrors.Append("Cannot interpret Receive Goods In Alerts value specified in row " & nRow.ToString & ". Expecting Y, YES, N, NO, 1, 0, TRUE or FALSE, case insensitive.<br />")
                    End If
                End If

                nValidationIndex = MARSHALL_RECEIVE_LOW_STOCK_ALERTS
                If sMarshaller(nValidationIndex) <> String.Empty Then
                    If Not (sMarshaller(nValidationIndex).ToUpper = "Y" Or sMarshaller(nValidationIndex).ToUpper = "YES" Or sMarshaller(nValidationIndex).ToUpper = "1" _
                      Or sMarshaller(nValidationIndex).ToUpper = "N" Or sMarshaller(nValidationIndex).ToUpper = "NO" Or sMarshaller(nValidationIndex).ToUpper = "0" _
                      Or sMarshaller(nValidationIndex).ToUpper = "TRUE" Or sMarshaller(nValidationIndex).ToUpper = "FALSE") _
                    Then
                        sbErrors.Append("Cannot interpret Receive Low Stock Alerts value specified in row " & nRow.ToString & ". Expecting Y, YES, N, NO, 1, 0, TRUE or FALSE, case insensitive.<br />")
                    End If
                End If

                nValidationIndex = MARSHALL_RECEIVE_OWN_CONSIGNMENT_BOOKING_ALERT
                If sMarshaller(nValidationIndex) <> String.Empty Then
                    If Not (sMarshaller(nValidationIndex).ToUpper = "Y" Or sMarshaller(nValidationIndex).ToUpper = "YES" Or sMarshaller(nValidationIndex).ToUpper = "1" _
                      Or sMarshaller(nValidationIndex).ToUpper = "N" Or sMarshaller(nValidationIndex).ToUpper = "NO" Or sMarshaller(nValidationIndex).ToUpper = "0" _
                      Or sMarshaller(nValidationIndex).ToUpper = "TRUE" Or sMarshaller(nValidationIndex).ToUpper = "FALSE") _
                    Then
                        sbErrors.Append("Cannot interpret Receive Own Consignment Booking Alert value specified in row " & nRow.ToString & ". Expecting Y, YES, N, NO, 1, 0, TRUE or FALSE, case insensitive.<br />")
                    End If
                End If

                nValidationIndex = MARSHALL_RECEIVE_ALL_CONSIGNMENT_BOOKING_ALERTS
                If sMarshaller(nValidationIndex) <> String.Empty Then
                    If Not (sMarshaller(nValidationIndex).ToUpper = "Y" Or sMarshaller(nValidationIndex).ToUpper = "YES" Or sMarshaller(nValidationIndex).ToUpper = "1" _
                      Or sMarshaller(nValidationIndex).ToUpper = "N" Or sMarshaller(nValidationIndex).ToUpper = "NO" Or sMarshaller(nValidationIndex).ToUpper = "0" _
                      Or sMarshaller(nValidationIndex).ToUpper = "TRUE" Or sMarshaller(nValidationIndex).ToUpper = "FALSE") _
                    Then
                        sbErrors.Append("Cannot interpret Receive All Consignment Booking Alerts value specified in row " & nRow.ToString & ". Expecting Y, YES, N, NO, 1, 0, TRUE or FALSE, case insensitive.<br />")
                    End If
                End If

                nValidationIndex = MARSHALL_RECEIVE_CONSIGNMENT_DESPATCH_ALERTS
                If sMarshaller(nValidationIndex) <> String.Empty Then
                    If Not (sMarshaller(nValidationIndex).ToUpper = "Y" Or sMarshaller(nValidationIndex).ToUpper = "YES" Or sMarshaller(nValidationIndex).ToUpper = "1" _
                      Or sMarshaller(nValidationIndex).ToUpper = "N" Or sMarshaller(nValidationIndex).ToUpper = "NO" Or sMarshaller(nValidationIndex).ToUpper = "0" _
                      Or sMarshaller(nValidationIndex).ToUpper = "TRUE" Or sMarshaller(nValidationIndex).ToUpper = "FALSE") _
                    Then
                        sbErrors.Append("Cannot interpret Receive Consignment Despatch Alerts value specified in row " & nRow.ToString & ". Expecting Y, YES, N, NO, 1, 0, TRUE or FALSE, case insensitive.<br />")
                    End If
                End If

                nValidationIndex = MARSHALL_RECEIVE_CONSIGNMENT_DELIVERY_ALERTS
                If sMarshaller(nValidationIndex) <> String.Empty Then
                    If Not (sMarshaller(nValidationIndex).ToUpper = "Y" Or sMarshaller(nValidationIndex).ToUpper = "YES" Or sMarshaller(nValidationIndex).ToUpper = "1" _
                      Or sMarshaller(nValidationIndex).ToUpper = "N" Or sMarshaller(nValidationIndex).ToUpper = "NO" Or sMarshaller(nValidationIndex).ToUpper = "0" _
                      Or sMarshaller(nValidationIndex).ToUpper = "TRUE" Or sMarshaller(nValidationIndex).ToUpper = "FALSE") _
                    Then
                        sbErrors.Append("Cannot interpret Receive Consignment Delivery Alerts value specified in row " & nRow.ToString & ". Expecting Y, YES, N, NO, 1, 0, TRUE or FALSE, case insensitive.<br />")
                    End If
                End If

                If bAddToDatabase And sbErrors.Length = 0 Then
                    Call AddNewUser(sMarshaller)
                End If
            Else
                'sbErrors.Append("...skipping empty row " & nRow.ToString & "<br />")
            End If
        Next
      
        If sbErrors.Length > 0 Then
            lblError.Text = sbErrors.ToString
        End If
    End Sub

    Protected Sub AddNewUser(ByVal sMarshaller() As String)
        lblError.Text = ""
        Dim nIndex As Integer
        Dim oPassword As SprintInternational.Password = New SprintInternational.Password()
        Dim oConn As New SqlConnection(gsConn)
        Dim oTrans As SqlTransaction
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_UserProfile_Add5", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
 
        Dim paramOriginatorKey As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int)
        paramOriginatorKey.Value = 0
        oCmd.Parameters.Add(paramOriginatorKey)

        Dim paramUserId As SqlParameter = New SqlParameter("@UserId", SqlDbType.NVarChar, 100)
        paramUserId.Value = sMarshaller(MARSHALL_USER_ID)
        oCmd.Parameters.Add(paramUserId)
      
        Dim paramPassword As SqlParameter = New SqlParameter("@Password", SqlDbType.NVarChar, 24)
        paramPassword.Value = oPassword.Encrypt(sMarshaller(MARSHALL_PASSWORD))
        oCmd.Parameters.Add(paramPassword)
      
        Dim paramFirstName As SqlParameter = New SqlParameter("@FirstName", SqlDbType.NVarChar, 50)
        paramFirstName.Value = sMarshaller(MARSHALL_FIRST_NAME)
        oCmd.Parameters.Add(paramFirstName)
      
        Dim paramLastName As SqlParameter = New SqlParameter("@LastName", SqlDbType.NVarChar, 50)
        paramLastName.Value = sMarshaller(MARSHALL_LAST_NAME)
        oCmd.Parameters.Add(paramLastName)
      
        Dim paramTitle As SqlParameter = New SqlParameter("@Title", SqlDbType.NVarChar, 50)
        paramTitle.Value = sMarshaller(MARSHALL_TITLE)
        oCmd.Parameters.Add(paramTitle)
      
        Dim paramDepartment As SqlParameter = New SqlParameter("@Department", SqlDbType.NVarChar, 20)
        paramDepartment.Value = sMarshaller(MARSHALL_DEPARTMENT)
        oCmd.Parameters.Add(paramDepartment)
      
        Dim paramUserGroup As SqlParameter = New SqlParameter("@UserGroup", SqlDbType.Int)
        paramUserGroup.Value = 0
        oCmd.Parameters.Add(paramUserGroup)
      
        Dim paramType As SqlParameter = New SqlParameter("@Type", SqlDbType.NVarChar, 20)
        If sMarshaller(MARSHALL_USER_TYPE).ToLower.StartsWith("u") Then
            paramType.Value = "User"
        End If
        If sMarshaller(MARSHALL_USER_TYPE).ToLower.StartsWith("p")
            paramType.Value = "Product Owner"
        End If
        If sMarshaller(MARSHALL_USER_TYPE).ToLower.StartsWith("s") Then
            paramType.Value = "SuperUser"
        End If
        oCmd.Parameters.Add(paramType)
      
        Dim paramStatus As SqlParameter = New SqlParameter("@Status", SqlDbType.NVarChar, 20)
        paramStatus.Value = "Active"
        oCmd.Parameters.Add(paramStatus)
      
        Dim paramCustomer As SqlParameter = New SqlParameter("@Customer", SqlDbType.Bit)
        paramCustomer.Value = 1
        oCmd.Parameters.Add(paramCustomer)

        Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int)
        paramCustomerKey.Value = pnSelectedCustomerKey
        oCmd.Parameters.Add(paramCustomerKey)
 
        Dim paramEmailAddr As SqlParameter = New SqlParameter("@EmailAddr", SqlDbType.NVarChar, 100)
        paramEmailAddr.Value = sMarshaller(MARSHALL_EMAIL_ADDRESS)
        oCmd.Parameters.Add(paramEmailAddr)
 
        Dim paramTelephone As SqlParameter = New SqlParameter("@Telephone", SqlDbType.NVarChar, 20)
        paramTelephone.Value = sMarshaller(MARSHALL_TELEPHONE)
        oCmd.Parameters.Add(paramTelephone)
 
        Dim paramCollectionPoint As SqlParameter = New SqlParameter("@CollectionPoint", SqlDbType.NVarChar, 50)
        paramCollectionPoint.Value = sMarshaller(MARSHALL_COLLECTION_POINT)
        oCmd.Parameters.Add(paramCollectionPoint)
 
        Dim paramURL As SqlParameter = New SqlParameter("@URL", SqlDbType.NVarChar, 100)
        paramURL.Value = sMarshaller(MARSHALL_COLLECTION_POINT)
        oCmd.Parameters.Add(paramURL)
 
        nIndex = MARSHALL_ABLE_TO_VIEW_STOCK
        Dim paramAbleToViewStock As SqlParameter = New SqlParameter("@AbleToViewStock", SqlDbType.Bit)
        If sMarshaller(nIndex) <> String.Empty AndAlso (sMarshaller(nIndex) = "1" Or sMarshaller(nIndex).ToUpper = "Y" Or sMarshaller(nIndex).ToUpper = "YES" Or sMarshaller(nIndex).ToUpper = "TRUE") Then
            paramAbleToViewStock.Value = 1
        Else
            paramAbleToViewStock.Value = 0
        End If
        oCmd.Parameters.Add(paramAbleToViewStock)

        nIndex = MARSHALL_ABLE_TO_CREATE_STOCK_BOOKING
        Dim paramAbleToCreateStockBooking As SqlParameter = New SqlParameter("@AbleToCreateStockBooking", SqlDbType.Bit)
        If sMarshaller(nIndex) <> String.Empty AndAlso (sMarshaller(nIndex) = "0" Or sMarshaller(nIndex).ToUpper = "N" Or sMarshaller(nIndex).ToUpper = "NO" Or sMarshaller(nIndex).ToUpper = "FALSE") Then
            paramAbleToCreateStockBooking.Value = 0
        Else
            paramAbleToCreateStockBooking.Value = 1
        End If
        oCmd.Parameters.Add(paramAbleToCreateStockBooking)

        nIndex = MARSHALL_ABLE_TO_CREATE_COLLECTION_REQUEST
        Dim paramAbleToCreateCollectionRequest As SqlParameter = New SqlParameter("@AbleToCreateCollectionRequest", SqlDbType.Bit)
        If sMarshaller(nIndex) <> String.Empty AndAlso (sMarshaller(nIndex) = "1" Or sMarshaller(nIndex).ToUpper = "Y" Or sMarshaller(nIndex).ToUpper = "YES" Or sMarshaller(nIndex).ToUpper = "TRUE") Then
            paramAbleToCreateCollectionRequest.Value = 1
        Else
            paramAbleToCreateCollectionRequest.Value = 0
        End If
        oCmd.Parameters.Add(paramAbleToCreateCollectionRequest)

        nIndex = MARSHALL_ABLE_TO_VIEW_GLOBAL_ADDRESS_BOOK
        Dim paramAbleToViewGlobalAddressBook As SqlParameter = New SqlParameter("@AbleToViewGlobalAddressBook", SqlDbType.Bit)
        If sMarshaller(nIndex) <> String.Empty AndAlso (sMarshaller(nIndex) = "1" Or sMarshaller(nIndex).ToUpper = "Y" Or sMarshaller(nIndex).ToUpper = "YES" Or sMarshaller(nIndex).ToUpper = "TRUE") Then
            paramAbleToViewGlobalAddressBook.Value = 1
        Else
            paramAbleToViewGlobalAddressBook.Value = 0
        End If
        oCmd.Parameters.Add(paramAbleToViewGlobalAddressBook)

        nIndex = MARSHALL_ABLE_TO_EDIT_GLOBAL_ADDRESS_BOOK
        Dim paramAbleToEditGlobalAddressBook As SqlParameter = New SqlParameter("@AbleToEditGlobalAddressBook", SqlDbType.Bit)
        If sMarshaller(nIndex) <> String.Empty AndAlso (sMarshaller(nIndex) = "1" Or sMarshaller(nIndex).ToUpper = "Y" Or sMarshaller(nIndex).ToUpper = "YES" Or sMarshaller(nIndex).ToUpper = "TRUE") Then
            paramAbleToEditGlobalAddressBook.Value = 1
        Else
            paramAbleToEditGlobalAddressBook.Value = 0
        End If
        oCmd.Parameters.Add(paramAbleToEditGlobalAddressBook)

        Dim paramRunningHeaderImage As SqlParameter = New SqlParameter("@RunningHeaderImage", SqlDbType.NVarChar, 100)
        paramRunningHeaderImage.Value = "default"
        oCmd.Parameters.Add(paramRunningHeaderImage)
 
        nIndex = MARSHALL_RECEIVE_OWN_STOCK_BOOKING_ALERT
        Dim paramStockBookingAlert As SqlParameter = New SqlParameter("@StockBookingAlert", SqlDbType.Bit)
        If sMarshaller(nIndex) <> String.Empty AndAlso (sMarshaller(nIndex) = "0" Or sMarshaller(nIndex).ToUpper = "N" Or sMarshaller(nIndex).ToUpper = "NO" Or sMarshaller(nIndex).ToUpper = "FALSE") Then
            paramStockBookingAlert.Value = 0
        Else
            paramStockBookingAlert.Value = 1
        End If
        oCmd.Parameters.Add(paramStockBookingAlert)

        nIndex = MARSHALL_RECEIVE_ALL_STOCK_BOOKING_ALERTS
        Dim paramStockBookingAlertAll As SqlParameter = New SqlParameter("@StockBookingAlertAll", SqlDbType.Bit)
        If sMarshaller(nIndex) <> String.Empty AndAlso (sMarshaller(nIndex) = "1" Or sMarshaller(nIndex).ToUpper = "Y" Or sMarshaller(nIndex).ToUpper = "YES" Or sMarshaller(nIndex).ToUpper = "TRUE") Then
            paramStockBookingAlertAll.Value = 1
        Else
            paramStockBookingAlertAll.Value = 0
        End If
        oCmd.Parameters.Add(paramStockBookingAlertAll)

        nIndex = MARSHALL_RECEIVE_GOODS_IN_ALERTS
        Dim paramStockArrivalAlert As SqlParameter = New SqlParameter("@StockArrivalAlert", SqlDbType.Bit)
        If sMarshaller(nIndex) <> String.Empty AndAlso (sMarshaller(nIndex) = "1" Or sMarshaller(nIndex).ToUpper = "Y" Or sMarshaller(nIndex).ToUpper = "YES" Or sMarshaller(nIndex).ToUpper = "TRUE") Then
            paramStockArrivalAlert.Value = 1
        Else
            paramStockArrivalAlert.Value = 0
        End If
        oCmd.Parameters.Add(paramStockArrivalAlert)

        nIndex = MARSHALL_RECEIVE_LOW_STOCK_ALERTS
        Dim paramLowStockAlert As SqlParameter = New SqlParameter("@LowStockAlert", SqlDbType.Bit)
        If sMarshaller(nIndex) <> String.Empty AndAlso (sMarshaller(nIndex) = "1" Or sMarshaller(nIndex).ToUpper = "Y" Or sMarshaller(nIndex).ToUpper = "YES" Or sMarshaller(nIndex).ToUpper = "TRUE") Then
            paramLowStockAlert.Value = 1
        Else
            paramLowStockAlert.Value = 0
        End If
        oCmd.Parameters.Add(paramLowStockAlert)

        nIndex = MARSHALL_RECEIVE_OWN_CONSIGNMENT_BOOKING_ALERT
        Dim paramConsignmentBookingAlert As SqlParameter = New SqlParameter("@ConsignmentBookingAlert", SqlDbType.Bit)
        If sMarshaller(nIndex) <> String.Empty AndAlso (sMarshaller(nIndex) = "1" Or sMarshaller(nIndex).ToUpper = "Y" Or sMarshaller(nIndex).ToUpper = "YES" Or sMarshaller(nIndex).ToUpper = "TRUE") Then
            paramConsignmentBookingAlert.Value = 1
        Else
            paramConsignmentBookingAlert.Value = 0
        End If
        oCmd.Parameters.Add(paramConsignmentBookingAlert)

        nIndex = MARSHALL_RECEIVE_ALL_CONSIGNMENT_BOOKING_ALERTS
        Dim paramConsignmentBookingAlertAll As SqlParameter = New SqlParameter("@ConsignmentBookingAlertAll", SqlDbType.Bit)
        If sMarshaller(nIndex) <> String.Empty AndAlso (sMarshaller(nIndex) = "1" Or sMarshaller(nIndex).ToUpper = "Y" Or sMarshaller(nIndex).ToUpper = "YES" Or sMarshaller(nIndex).ToUpper = "TRUE") Then
            paramConsignmentBookingAlertAll.Value = 1
        Else
            paramConsignmentBookingAlertAll.Value = 0
        End If
        oCmd.Parameters.Add(paramConsignmentBookingAlertAll)

        nIndex = MARSHALL_RECEIVE_CONSIGNMENT_DESPATCH_ALERTS
        Dim paramConsignmentDespatchAlert As SqlParameter = New SqlParameter("@ConsignmentDespatchAlert", SqlDbType.Bit)
        If sMarshaller(nIndex) <> String.Empty AndAlso (sMarshaller(nIndex) = "1" Or sMarshaller(nIndex).ToUpper = "Y" Or sMarshaller(nIndex).ToUpper = "YES" Or sMarshaller(nIndex).ToUpper = "TRUE") Then
            paramConsignmentDespatchAlert.Value = 1
        Else
            paramConsignmentDespatchAlert.Value = 0
        End If
        oCmd.Parameters.Add(paramConsignmentDespatchAlert)

        nIndex = MARSHALL_RECEIVE_CONSIGNMENT_DELIVERY_ALERTS
        Dim paramConsignmentDeliveryAlert As SqlParameter = New SqlParameter("@ConsignmentDeliveryAlert", SqlDbType.Bit)
        If sMarshaller(nIndex) <> String.Empty AndAlso (sMarshaller(nIndex) = "1" Or sMarshaller(nIndex).ToUpper = "Y" Or sMarshaller(nIndex).ToUpper = "YES" Or sMarshaller(nIndex).ToUpper = "TRUE") Then
            paramConsignmentDeliveryAlert.Value = 1
        Else
            paramConsignmentDeliveryAlert.Value = 0
        End If
        oCmd.Parameters.Add(paramConsignmentDeliveryAlert)

        'Dim paramUserPermissions As SqlParameter = New SqlParameter("@UserPermissions", SqlDbType.Int)
        'paramUserPermissions.Value = 0
        'oCmd.Parameters.Add(paramUserPermissions)

        Dim paramUserKey As SqlParameter = New SqlParameter("@UserProfileKey", SqlDbType.Int, 4)
        paramUserKey.Direction = ParameterDirection.Output
        oCmd.Parameters.Add(paramUserKey)
 
        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
            Dim lUserKey As Long = CLng(oCmd.Parameters("@UserProfileKey").Value)

            tbResults.Text += "Added user " & sMarshaller(MARSHALL_FIRST_NAME) & " " & sMarshaller(MARSHALL_LAST_NAME) & " (" & sMarshaller(MARSHALL_USER_ID) & ") as user " & lUserKey.ToString & Environment.NewLine

        Catch ex As SqlException
            If ex.Number = 2627 Then
                lblError.Text = "ERROR: A record already exists with the user ID " & sMarshaller(MARSHALL_USER_ID)
                tbResults.Text += "ERROR: A record already exists with the user ID " & sMarshaller(MARSHALL_USER_ID) & Environment.NewLine
            Else
                lblError.Text = ex.ToString
            End If
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub Page_PreRender(ByVal sender As Object, ByVal e As System.EventArgs)
        If gvUserDetails.Rows.Count > 0 Then
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
        If btnHelp.Text = "Show Help" Then
            pnlHelp.Visible = True
            btnHelp.Text = "Hide Help"
        Else
            pnlHelp.Visible = False
            btnHelp.Text = "Show Help"
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
   
    Protected Sub ddlCustomers_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        pnSelectedCustomerKey = ddl.SelectedValue
        btnUploadUsers.Enabled = True
        btnCheckData.Enabled = True
       
        FileUpload1.Enabled = True
        cbColumnHeadingsInRow1.Enabled = True
        btnReadUsers.Enabled = True
       
        lblCustomer.Text = "to " & ddl.SelectedItem.Text
       
        If ddl.Items(0).Text = String.Empty Then
            ddl.Items.RemoveAt(0)
        End If
    End Sub
    
    Protected Sub lnkbtnRestart_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call Initialise()
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

    Protected Function ExecuteNonQuery(ByVal sQuery As String) As Boolean
        ExecuteNonQuery = False
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand
        Try
            oConn.Open()
            oCmd = New SqlCommand(sQuery, oConn)
            oCmd.ExecuteNonQuery()
            ExecuteNonQuery = True
        Catch ex As Exception
            WebMsgBox.Show("Error in ExecuteNonQuery executing: " & sQuery & " : " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Function

</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Upload Products</title>
</head>
<body>
    <form id="form1" runat="server" defaultfocus="FormUpload1" enctype="multipart/form-data">
        <strong>UPLOAD USERS &nbsp;&nbsp;&nbsp; &nbsp;<asp:Label ID="lblCustomer" runat="server"
            Text="(no customer selected)"></asp:Label>
            &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
            <asp:Button ID="btnHelp" runat="server" Text="show help" Width="112px" OnClick="btnHelp_Click" />
            &nbsp; &nbsp; &nbsp;&nbsp;
            <asp:LinkButton ID="lnkbtnRestart" runat="server" OnClick="lnkbtnRestart_Click">restart</asp:LinkButton><br />
        </strong>
        <br />
        <asp:Panel ID="pnlHelp" runat="server" Width="100%">
            The Upload Users facility loads user information from a CSV (comma-separated variable)
            format file you supply. Follow the steps below to specify (a) the location of the
            file and (b) how to use each field of user information.<br />
            <br />
            1. To import user details from Excel you must first convert the data into CSV format.
            Open the spreadsheet, choose Save As, select CSV as the file type, specify a location
            and filename (eg C:\myusers.csv), then click Save.<br />
            <br />
            2. Click the <strong>Browse</strong> button to locate the file of user information
            on your
            local machine.<br />
            <br />
            3. If your data contains columns headings (ie the first row of data is the name
            or description of the column) click the <strong>Column Headings in Row 1</strong>
            check box.<br />
            <br />
            4. Click the <strong>Read User Details </strong>button. The system reads and interprets
            your file, then displays the contents for you to check and confirm before loading
            it into the system. A message is displayed if there are problems reading the data.<br />
            <br />
            5. For each column of data you want to include, choose the field to associate with this column, using the dropdown list box at the top of each column. The columns you <strong>must</strong> include are:<br /> - User ID (unless the automatically generated User IDs option is checked)<br /> - First Name<br /> - Last Name<br /> - Email Address<br />
            <br />
            6. Click <strong>Check Data</strong> if you want to check that the system can correctly
            process your data. The system checks your data and displays a message if a problem
            is found or required data is missing. Correct any errors and re-submit the data.<br />
            <br />
            7. Click <strong>Upload User Details </strong>to load your data into the system.<br />

    <strong style="font-weight:normal;"></strong>

    <p dir="ltr" style="line-height:1;margin-top:0pt;margin-bottom:0pt;">
    <strong style="font-weight:normal;"><span style=
    "font-size:13px;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;white-space:pre-wrap;">
    <strong>DEFAULTS</strong></span></strong></p>

    <p dir="ltr" style="line-height:1;margin-top:0pt;margin-bottom:0pt;">
    <strong style="font-weight:normal;">&nbsp;</strong></p>

    <p dir="ltr" style="line-height:1;margin-top:0pt;margin-bottom:0pt;">
    <strong style="font-weight:normal;">User Type = <strong>User</strong></strong></p>

    <p dir="ltr" style="line-height:1;margin-top:0pt;margin-bottom:0pt;">
    <strong style="font-weight:normal;">&nbsp;</strong></p>

    <p dir="ltr" style="line-height:1;margin-top:0pt;margin-bottom:0pt;">
    <strong style="font-weight:normal;"><span style=
    "font-size:13px;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;white-space:pre-wrap;">
    ABLE_TO_CREATE_STOCK_BOOKING = <strong>true</strong></span></strong></p>

    <p dir="ltr" style="line-height:1;margin-top:0pt;margin-bottom:0pt;">
    <strong style="font-weight:normal;"><span style=
    "font-size:13px;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;white-space:pre-wrap;">
    RECEIVE_OWN_STOCK_BOOKING_ALERT = <strong>true</strong></span></strong></p>

    <p dir="ltr" style="line-height:1;margin-top:0pt;margin-bottom:0pt;">
    <strong style="font-weight:normal;">&nbsp;</strong></p>

    <p dir="ltr" style="line-height:1;margin-top:0pt;margin-bottom:0pt;">
    <strong style="font-weight:normal;"><span style=
    "font-size:13px;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;white-space:pre-wrap;">
    ABLE_TO_CREATE_COLLECTION_REQUEST = false</span></strong></p>

    <p dir="ltr" style="line-height:1;margin-top:0pt;margin-bottom:0pt;">
    <strong style="font-weight:normal;"><span style=
    "font-size:13px;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;white-space:pre-wrap;">
    ABLE_TO_VIEW_GLOBAL_ADDRESS_BOOK = false</span></strong></p>

    <p dir="ltr" style="line-height:1;margin-top:0pt;margin-bottom:0pt;">
    <strong style="font-weight:normal;"><span style=
    "font-size:13px;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;white-space:pre-wrap;">
    ABLE_TO_EDIT_GLOBAL_ADDRESS_BOOK = false</span></strong></p>

    <p dir="ltr" style="line-height:1;margin-top:0pt;margin-bottom:0pt;">
    <strong style="font-weight:normal;"><span style=
    "font-size:13px;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;white-space:pre-wrap;">
    RECEIVE_ALL_STOCK_BOOKING_ALERTS = false</span></strong></p>

    <p dir="ltr" style="line-height:1;margin-top:0pt;margin-bottom:0pt;">
    <strong style="font-weight:normal;"><span style=
    "font-size:13px;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;white-space:pre-wrap;">
    RECEIVE_GOODS_IN_ALERTS = false</span></strong></p>

    <p dir="ltr" style="line-height:1;margin-top:0pt;margin-bottom:0pt;">
    <strong style="font-weight:normal;"><span style=
    "font-size:13px;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;white-space:pre-wrap;">
    RECEIVE_LOW_STOCK_ALERTS = false</span></strong></p>

    <p dir="ltr" style="line-height:1;margin-top:0pt;margin-bottom:0pt;">
    <strong style="font-weight:normal;">&nbsp;</strong></p>

    <p dir="ltr" style="line-height:1;margin-top:0pt;margin-bottom:0pt;">
    <strong style="font-weight:normal;"><span style=
    "font-size:13px;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;white-space:pre-wrap;">
    ABLE_TO_VIEW_STOCK = false (NOTE: defined but not used
    internally)</span></strong></p>

    <p dir="ltr" style="line-height:1;margin-top:0pt;margin-bottom:0pt;">
    <strong style="font-weight:normal;"><span style=
    "font-size:13px;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;white-space:pre-wrap;">
    RECEIVE_OWN_CONSIGNMENT_BOOKING_ALERT = false (NOTE: no longer
    used)</span></strong></p>

    <p dir="ltr" style="line-height:1;margin-top:0pt;margin-bottom:0pt;">
    <strong style="font-weight:normal;"><span style=
    "font-size:13px;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;white-space:pre-wrap;">
    RECEIVE_ALL_CONSIGNMENT_BOOKING_ALERTS = false (NOTE: no longer
    used)</span></strong></p>

    <p dir="ltr" style="line-height:1;margin-top:0pt;margin-bottom:0pt;">
    <strong style="font-weight:normal;"><span style=
    "font-size:13px;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;white-space:pre-wrap;">
    RECEIVE_CONSIGNMENT_DESPATCH_ALERTS = false (NOTE: no longer
    used)</span></strong></p>

    <p dir="ltr" style="line-height:1;margin-top:0pt;margin-bottom:0pt;">
    <strong style="font-weight:normal;"><span style=
    "font-size:13px;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;white-space:pre-wrap;">
    RECEIVE_CONSIGNMENT_DELIVERY_ALERTS = false (NOTE: no longer
    used)</span></strong></p><br />
            [end]<br />
        </asp:Panel>
        <asp:Panel ID="pnlCommon" runat="server" Width="100%">
        <asp:Label ID="Label1" runat="server" Text="Customer:" Font-Names="Verdana" Font-Size="XX-Small"></asp:Label>
            <asp:DropDownList ID="ddlCustomers" runat="server" OnSelectedIndexChanged="ddlCustomers_SelectedIndexChanged"
                AutoPostBack="True" Font-Names="Verdana" Font-Size="XX-Small" />
            &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
            </asp:Panel>
        <asp:Panel ID="pnlStart" runat="server" Width="100%">
            <asp:Label ID="Label2" runat="server" Text="User CSV file:" Font-Names="Verdana" Font-Size="XX-Small"></asp:Label>
            <asp:FileUpload ID="FileUpload1" runat="server" Width="350px" Enabled="False" Font-Names="Verdana" Font-Size="XX-Small" />
            &nbsp;&nbsp;<asp:CheckBox ID="cbColumnHeadingsInRow1" runat="server" Text="column&nbsp;headings&nbsp;in&nbsp;row&nbsp;1"
                Enabled="False" Font-Names="Verdana" Font-Size="XX-Small" />&nbsp;
            <asp:Button ID="btnReadUsers" runat="server" OnClick="btnReadUsers_Click" Text="read user details"
                Width="140px" Enabled="False" Font-Names="Verdana" Font-Size="XX-Small" />&nbsp;<br /><asp:CheckBox ID="cbAutoGenerateUserId" runat="server" Text="autogenerate user id from First Name, Last Name separated by:" Font-Names="Verdana" Font-Size="XX-Small" />
            <asp:TextBox ID="tbUserIdSeparatorChars" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                Width="50px"></asp:TextBox></asp:Panel>
        <asp:Panel ID="pnlGridView" runat="server" Width="100%">
            <br />
            <asp:Button ID="btnCheckData" runat="server" Text="check data" OnClick="btnCheckData_Click"
                Enabled="False" />
            <asp:Button ID="btnUploadUsers" runat="server" Text="upload user details" OnClick="btnUploadUsers_Click"
                Enabled="False" /><br />
            <br />
            <asp:GridView ID="gvUserDetails" runat="server" Width="100%" Font-Names="Verdana" Font-Size="XX-Small"/>
        </asp:Panel>
        <asp:Panel ID="pnlMessage" runat="server" Width="100%">
            <asp:Label ID="lblMessage" runat="server" Font-Names="Verdana" Font-Size="XX-Small" /><br />
            <br />
            <br />
            &nbsp;</asp:Panel>
        <asp:Panel ID="pnlFinished" runat="server" Width="100%">
            <asp:Label ID="Label3" runat="server" Text="Results:" Font-Names="Verdana" Font-Size="XX-Small" />
            <asp:TextBox ID="tbResults" runat="server" Rows="10" TextMode="MultiLine" Width="100%" Font-Names="Verdana" Font-Size="XX-Small" />
            <br />
            <asp:Label ID="lblSuccess" runat="server" ForeColor="Green" Font-Names="Verdana" Font-Size="XX-Small"></asp:Label>
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
