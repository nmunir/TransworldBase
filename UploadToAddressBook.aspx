<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ import Namespace="System.IO" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">
    
    ' check if all selected fields are different (no dupes)
    ' check that country is selected, or defaulted
    ' show number of addresses read
    ' show number of addresses written
    ' later on, split large files into chunks and process sequentially
    ' use default country doesn't override country matching
    ' offer to show 1st line as prompt for column headers on line 1
    ' show if required field is not select in message box instead of erroring every missing instance
    
    Dim sFileName As String, sIntermediateFileName As String
    Dim sFilePrefix As String, sFileSuffix As String = ".csv"
    Dim lAddressKey As Long
    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsNumeric(Session("UserKey")) Then
            Server.Transfer("session_expired.aspx")
        End If
        If Not IsPostBack Then
            Call GetCountries()
            If Len(LookupAliasCountry("U.K")) > 0 Then  ' check alias table has been populated; this is called every time and assumes the upload function will not be used very often; this test & the button can be removed if necessary
                btnPopulateAliasTable.Visible = False
            End If
            pnlStart.Visible = True
            pnlCountryAliasing.Visible = False
            pnlGridView.Visible = False
            pnlMessage.Visible = False
            pnlHelp.Visible = False
            pnlDeleteAlias.Visible = False
            
            If Not Session("EditGAB") Then
                rblAddressDestination.SelectedIndex = 1
                rblAddressDestination.Items(0).Enabled = False
            End If
        End If
    End Sub
    
    Protected Sub btnReadAddresses_Click(ByVal sender As Object, ByVal e As System.EventArgs)
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
                WebMsgBox.Show("WARNING: The address file you are uploading contains " & nLineCount.ToString & " addresses.  To ensure successful uploading we recommend that you split the address file into multiple files of no more than 10,000 addresses per file.")
            Else
                sIntermediateFileName = Server.MapPath("") & "\" & sFilePrefix & "-2" & sFileSuffix
                Call RemoveEmbeddedLineBreaksAndCommas(sFileName, sIntermediateFileName)
                Dim dt As DataTable
                dt = DelimitFile(sIntermediateFileName, ",", cbColumnHeadingsInRow1.Checked)
                If dt.Rows.Count > 0 Then
                    Dim r As DataRow = dt.NewRow()
                    dt.Rows.InsertAt(r, 0)
                    GridView1.DataSource = dt
                    GridView1.DataBind()
                End If
                If GridView1.Rows.Count > 0 Then
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
    
    Private Sub AddDropDowns()
        Dim gvr As GridViewRow
        gvr = GridView1.Rows(0)
        Dim j As Integer
        For j = 0 To gvr.Cells.Count - 1
            Dim dd As New DropDownList
            dd.Font.Size = FontUnit.XXSmall
            dd.ID = "Select" & j.ToString.Trim
            'dd.AutoPostBack = True
            ' following list must match the array used to match selection during processing            
            dd.Items.Add("- nothing selected -")
            dd.Items.Add("- don't use this column -")
            dd.Items.Add("Short Code")
            dd.Items.Add("Addressee")
            dd.Items.Add("Addr Line 1")
            dd.Items.Add("Addr Line 2")
            dd.Items.Add("Addr Line 3")
            dd.Items.Add("Town/City")
            dd.Items.Add("County/State")
            dd.Items.Add("Post Code")
            dd.Items.Add("Country")
            dd.Items.Add("Attention of")
            dd.Items.Add("Telephone")
            dd.SelectedIndex = ExtractHiddenFieldValue(j)

            gvr.Cells(j).Controls.Add(dd)
            Dim cid As String = dd.ClientID
            dd.Attributes.Add("onClick", "HiddenField" & j.ToString.Trim & ".value=" & cid & ".selectedIndex; HiddenFieldChanged.value='TRUE'")
        Next
    End Sub

    Protected Function ExtractHiddenFieldValue(ByVal nIndex As Integer) As Integer
        ExtractHiddenFieldValue = 0
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
    
    Sub GetCountries()
        Dim oConn1 As New SqlConnection(gsConn)
        Dim oConn2 As New SqlConnection(gsConn)
        Dim oCmd1 As New SqlCommand("spASPNET_Country_GetCountries", oConn1)
        oCmd1.CommandType = CommandType.StoredProcedure
        Dim oCmd2 As New SqlCommand("spASPNET_Country_GetCountries", oConn2)
        oCmd2.CommandType = CommandType.StoredProcedure
        Try
            oConn1.Open()
            ddlCountries.DataSource = oCmd1.ExecuteReader()
            ddlCountries.DataTextField = "CountryName"
            ddlCountries.DataValueField = "CountryKey"
            ddlCountries.DataBind()
            oConn2.Open()
            ddlAliasCountries.DataSource = oCmd2.ExecuteReader()
            ddlAliasCountries.DataTextField = "CountryName"
            ddlAliasCountries.DataValueField = "CountryKey"
            ddlAliasCountries.DataBind()
            
        Catch ex As SqlException
            lblError.Text = ex.ToString
        End Try
        oConn1.Close()
        oConn2.Close()
    End Sub

    ' Dim sAddrMarshaller() As String = {"ShortCode", "Addressee", "Addr1", "Addr2", "Addr3", "City", "State", "PostCode", "Country", "Attn", "Tel"} ' note initialisation is for documentation purposes

    Sub AddNewAddress(ByVal sAddressMarshaller() As String)
        Dim lCountryKey As Long
        Dim bError As Boolean = False
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Address_Add", oConn)
        Dim oTrans As SqlTransaction
        oCmd.CommandType = CommandType.StoredProcedure
        lAddressKey = -1    ' the address key is returned by proc making insert so initialise it here
        lCountryKey = CLng(ddlCountries.SelectedItem.Value)
        Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
        paramCustomerKey.Value = Session("CustomerKey")
        oCmd.Parameters.Add(paramCustomerKey)
        Dim paramCode As SqlParameter = New SqlParameter("@Code", SqlDbType.NVarChar, 20)
        paramCode.Value = sAddressMarshaller(0)
        oCmd.Parameters.Add(paramCode)
        Dim paramCompany As SqlParameter = New SqlParameter("@Company", SqlDbType.NVarChar, 50)
        paramCompany.Value = sAddressMarshaller(1)
        oCmd.Parameters.Add(paramCompany)
        Dim paramAddr1 As SqlParameter = New SqlParameter("@Addr1", SqlDbType.NVarChar, 50)
        paramAddr1.Value = sAddressMarshaller(2)
        oCmd.Parameters.Add(paramAddr1)
        Dim paramparamAddr2 As SqlParameter = New SqlParameter("@Addr2", SqlDbType.NVarChar, 50)
        paramparamAddr2.Value = sAddressMarshaller(3)
        oCmd.Parameters.Add(paramparamAddr2)
        Dim paramparamAddr3 As SqlParameter = New SqlParameter("@Addr3", SqlDbType.NVarChar, 50)
        paramparamAddr3.Value = sAddressMarshaller(4)
        oCmd.Parameters.Add(paramparamAddr3)
        Dim paramTown As SqlParameter = New SqlParameter("@Town", SqlDbType.NVarChar, 50)
        paramTown.Value = sAddressMarshaller(5)
        oCmd.Parameters.Add(paramTown)
        Dim paramState As SqlParameter = New SqlParameter("@State", SqlDbType.NVarChar, 50)
        paramState.Value = sAddressMarshaller(6)
        oCmd.Parameters.Add(paramState)
        Dim paramPostCode As SqlParameter = New SqlParameter("@PostCode", SqlDbType.NVarChar, 50)
        paramPostCode.Value = sAddressMarshaller(7)
        oCmd.Parameters.Add(paramPostCode)
        Dim paramCountryKey As SqlParameter = New SqlParameter("@CountryKey", SqlDbType.Int, 4)
        paramCountryKey.Value = CInt(sAddressMarshaller(8))
        oCmd.Parameters.Add(paramCountryKey)
        Dim paramDefaultCommodityId As SqlParameter = New SqlParameter("@DefaultCommodityId", SqlDbType.NVarChar, 100)
        paramDefaultCommodityId.Value = ""
        oCmd.Parameters.Add(paramDefaultCommodityId)
        Dim paramDefaultSpecialInstructions As SqlParameter = New SqlParameter("@DefaultSpecialInstructions", SqlDbType.NVarChar, 100)
        paramDefaultSpecialInstructions.Value = ""
        oCmd.Parameters.Add(paramDefaultSpecialInstructions)
        Dim paramAttnOf As SqlParameter = New SqlParameter("@AttnOf", SqlDbType.NVarChar, 50)
        paramAttnOf.Value = sAddressMarshaller(9)
        oCmd.Parameters.Add(paramAttnOf)
        Dim paramTelephone As SqlParameter = New SqlParameter("@Telephone", SqlDbType.NVarChar, 50)
        paramTelephone.Value = sAddressMarshaller(10)
        oCmd.Parameters.Add(paramTelephone)
        Dim paramFax As SqlParameter = New SqlParameter("@Fax", SqlDbType.NVarChar, 50)
        paramFax.Value = ""
        oCmd.Parameters.Add(paramFax)
        Dim paramEmail As SqlParameter = New SqlParameter("@Email", SqlDbType.NVarChar, 50)
        paramEmail.Value = ""
        oCmd.Parameters.Add(paramEmail)
        Dim paramLastUpdatedByKey As SqlParameter = New SqlParameter("@LastUpdatedByKey", SqlDbType.Int, 4)
        paramLastUpdatedByKey.Value = Session("UserKey")
        oCmd.Parameters.Add(paramLastUpdatedByKey)
        Dim paramAddressKey As SqlParameter = New SqlParameter("@AddressKey", SqlDbType.Int, 4)
        paramAddressKey.Direction = ParameterDirection.Output
        oCmd.Parameters.Add(paramAddressKey)
        Try
            oConn.Open()
            oTrans = oConn.BeginTransaction(IsolationLevel.ReadCommitted, "AddRecord")
            oCmd.Connection = oConn
            oCmd.Transaction = oTrans
            oCmd.ExecuteNonQuery()
            oTrans.Commit()
            lAddressKey = paramAddressKey.Value
        Catch ex As SqlException
            oTrans.Rollback("AddRecord")
            bError = True
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
    End Sub

    Sub AddToSharedAddressBook()
        If lAddressKey > 0 Then
            Dim bError As Boolean = False
            Dim oConn As New SqlConnection(gsConn)
            Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Address_AddToGlobal", oConn)
            Dim oTrans As SqlTransaction
            oCmd.CommandType = CommandType.StoredProcedure
            Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
            paramCustomerKey.Value = Session("CustomerKey")
            oCmd.Parameters.Add(paramCustomerKey)
            Dim paramAddressKey As SqlParameter = New SqlParameter("@GlobalAddressKey", SqlDbType.Int, 4)
            paramAddressKey.Value = lAddressKey
            oCmd.Parameters.Add(paramAddressKey)
            Try
                oConn.Open()
                oTrans = oConn.BeginTransaction(IsolationLevel.ReadCommitted, "AddRecord")
                oCmd.Connection = oConn
                oCmd.Transaction = oTrans
                oCmd.ExecuteNonQuery()
                oTrans.Commit()
            Catch ex As SqlException
                oTrans.Rollback("AddRecord")
                bError = True
                'lblError.Text = ex.ToString
            Finally
                oConn.Close()
            End Try
            If Not bError Then
                'lblAddToGAB.Text = "Address added to Global Address Book"
            End If
        End If
    End Sub
    
    Sub AddToPersonalAddressBook()
        If lAddressKey > 0 Then
            Dim bError As Boolean = False
            Dim oConn As New SqlConnection(gsConn)
            Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Address_AddToPersonal", oConn)
            Dim oTrans As SqlTransaction
            oCmd.CommandType = CommandType.StoredProcedure
            Dim paramUserKey As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int, 4)
            paramUserKey.Value = Session("UserKey")
            oCmd.Parameters.Add(paramUserKey)
            Dim paramAddressKey As SqlParameter = New SqlParameter("@GlobalAddressKey", SqlDbType.Int, 4)
            paramAddressKey.Value = lAddressKey
            oCmd.Parameters.Add(paramAddressKey)
            Try
                oConn.Open()
                oTrans = oConn.BeginTransaction(IsolationLevel.ReadCommitted, "AddRecord")
                oCmd.Connection = oConn
                oCmd.Transaction = oTrans
                oCmd.ExecuteNonQuery()
                oTrans.Commit()
            Catch ex As SqlException
                oTrans.Rollback("AddRecord")
                bError = True
                'lblError.Text = ex.ToString
            Finally
                oConn.Close()
            End Try
            If Not bError Then
                'lblAddToMyAddressBook.Text = "Address added to My Address Book"
            End If
        Else
            ' couldn't make the insert as no address key
        End If
    End Sub
    
    Protected Sub OutputErrorMessage(ByVal sMessage As String) ' ensures any existing message is not overwritten
        If lblError.Text = "" Then
            lblError.Text = sMessage
        End If
    End Sub
    
    Protected Sub btnCheckData_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If SelectedFieldsUnique() Then
            lblMessage.Text = ""
            lblSuccess.Text = ""
            Call ProcessAddresses(bAddToDatabase:=False)
            If lblError.Text = "" Then
                lblSuccess.Text = "No errors found"
            End If
        Else
            WebMsgBox.Show("Two or more columns are mapped to the same address field.")
        End If
    End Sub

    Protected Sub btnUploadAddresses_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If SelectedFieldsUnique() Then
            lblMessage.Text = ""
            lblSuccess.Text = ""
            Call ProcessAddresses(bAddToDatabase:=False)
            If lblError.Text = "" Then
                Call ProcessAddresses(bAddToDatabase:=True)
            End If
            If lblError.Text = "" Then
                lblMessage.Text = GridView1.Rows.Count.ToString & " addresses added to your address book"
                pnlGridView.Visible = False
            End If
            If lblError.Text = "" Then
                lblSuccess.Text = "Addresses successfully loaded into Address Book"
            End If
        Else
            WebMsgBox.Show("Two or more columns are mapped to the same address field.")
        End If
    End Sub
    
    Protected Function SelectedFieldsUnique() As Boolean
        
        ' DO MSG BOX IN HERE; ALSO CHECK COUNTRY SELECTED; ALSO CHECK SOMETHING SELECTED
        Dim nCol As Integer
        Dim nMaxCols As Integer = GridView1.Rows(0).Cells.Count
        Dim nTargetIndex As Integer
        Dim nSelectedField(12) As Int16  ' dropdown fields: "", "", "ShortCode", "Addressee", "Addr1", "Addr2", "Addr3", "City", "State", "PostCode", "Country", "Attn", "Tel"
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
    
    Protected Sub ProcessAddresses(ByVal bAddToDatabase As Boolean)
        'Dim sSQL As String
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataSet As New DataSet
        
        Dim nRow As Integer, nMaxRows As Integer = GridView1.Rows.Count
        Dim nCol As Integer, nMaxCols As Integer = GridView1.Columns.Count
        Dim i As Integer, sVal As String, bCountryFound As Boolean
        Dim nTargetIndex As Integer
        ' following 2 lists must match the contents of the dropdown list box used to select target (sAddrMarshaller drops initial 2 items)
        Dim sTarget() As String = {"", "", "ShortCode", "Addressee", "Addr1", "Addr2", "Addr3", "City", "State", "PostCode", "Country", "Attn", "Tel"} ' this is the order of significant items presented in the dropdowns
        Dim sAddrMarshaller() As String = {"ShortCode", "Addressee", "Addr1", "Addr2", "Addr3", "City", "State", "PostCode", "Country", "Attn", "Tel"} ' note initialisation is for documentation purposes
        Dim sbErrors As New StringBuilder
        Dim sUnmatchedCountries As New StringCollection
        lblError.Text = ""
        lblErrorProxy.Text = ""
        ddlUnmatchedCountries.Items.Clear()
        
        nMaxRows = GridView1.Rows.Count
        nMaxCols = GridView1.Rows(0).Cells.Count
        
        If cbUseDefaultCountry.Checked = True Then
            If ddlCountries.SelectedIndex > 0 Then
                
            Else
            End If
        End If
        
        For nRow = 1 To nMaxRows - 1
            For i = sAddrMarshaller.GetLowerBound(0) To sAddrMarshaller.GetUpperBound(0)
                sAddrMarshaller(i) = ""
            Next
            bCountryFound = False
            For nCol = 0 To nMaxCols - 1
                nTargetIndex = ExtractHiddenFieldValue(nCol)
                sVal = HttpUtility.HtmlDecode(GridView1.Rows(nRow).Cells(nCol).Text).Trim
                If nTargetIndex > 0 Then 'a
                    If sTarget(nTargetIndex) = "Country" Then 'b
                        sVal = sVal.Trim.ToUpper
                        Dim sCountryName As String = sVal
                        If sCountryName <> "UK" Then
                            Dim sStop As String = sCountryName
                        End If
                        If Not sVal = "" Then
                            Dim sCountry As String = LookupAliasCountry(sVal.Trim.ToUpper)
                            If sCountry.Length > 0 Then
                                sCountry = LookupRealCountry(sCountry)
                            End If
                            If sCountry.Length > 0 Then
                                bCountryFound = True
                                sVal = sCountry
                            Else
                                sAddrMarshaller(8) = "NO MATCH" ' this could still get the default country but won't be added since an error message has been generated
                                If Not sUnmatchedCountries.Contains(sVal) Then
                                    sUnmatchedCountries.Add(sVal)
                                    sbErrors.Append("Could not match country '" & sVal & "<br />")
                                End If
                            End If
                        Else
                            ' no country
                        End If
                    End If 'b
                    sAddrMarshaller(nTargetIndex - 2) = sVal
                End If 'a
            Next
            If Not bCountryFound Then
                If cbUseDefaultCountry.Checked = True Then
                    If ddlCountries.SelectedIndex > 0 Then
                        sAddrMarshaller(8) = ddlCountries.SelectedValue.ToString
                        bCountryFound = True
                    Else
                        sbErrors.Append("Default country was required when processing row " & nRow.ToString & "but no default country was selected" & "<br />")
                    End If
                Else
                    sbErrors.Append("No country specified in row " & nRow.ToString & "<br />")
                End If
            End If
            If sAddrMarshaller(1) = "" Then
                sbErrors.Append("No addressee specified in row " & nRow.ToString & "<br />")
            End If
            
            If sAddrMarshaller(2) = "" Then
                sbErrors.Append("Required first line of address missing in row " & nRow.ToString & "<br />")
            End If

            If sAddrMarshaller(5) = "" Then
                sbErrors.Append("Required Town/City missing in row " & nRow.ToString & "<br />")
            End If

            If bAddToDatabase And bCountryFound And sbErrors.Length = 0 Then
                Call AddNewAddress(sAddrMarshaller)
                If rblAddressDestination.Items(0).Selected Then
                    Call AddToSharedAddressBook()
                Else
                    Call AddToPersonalAddressBook()
                End If
            End If
        Next
        If sUnmatchedCountries.Count > 0 Then
            pnlCountryAliasing.Visible = True
            For Each s As String In sUnmatchedCountries
                ddlUnmatchedCountries.Items.Add(s)
            Next
        Else
            pnlCountryAliasing.Visible = False
        End If
        If sbErrors.Length > 0 Then
            lblError.Text = sbErrors.ToString
            lblErrorProxy.Text = "Errors were found - see below"
        End If
    End Sub

    Function LookupRealCountry(ByVal sCountryName As String) As String
        Dim sCountryKey As String
        Dim sCountry As String = sCountryName.Replace("'", "''").Trim.ToUpper
        Dim sSQL As String = "SELECT * FROM Country WHERE CountryName = '" & sCountry & "'"
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter(sSQL, oConn)
        Dim oDataSet As New DataSet
        oAdapter.Fill(oDataSet, "Country")
        Dim dvSource As DataView = oDataSet.Tables("Country").DefaultView
        If dvSource.Count > 0 Then
            sCountryKey = dvSource(0).Item("CountryKey").ToString
            oDataSet.Clear()
            Return sCountryKey
        Else
            Return ""
        End If
    End Function
    
    Function LookupAliasCountry(ByVal sCountryName As String) As String
        Dim sRealName As String
        Dim sCountry As String = CleanUpCountryName(sCountryName)
        Dim sSQL As String = "SELECT * FROM CountryAliases WHERE MatchName = '" & sCountry & "' AND (CustomerKey = " & Session("CustomerKey") & " OR CustomerKey = 0)"
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter(sSQL, oConn)
        Dim oDataSet As New DataSet
        oAdapter.Fill(oDataSet, "Country")
        Dim dvSource As DataView = oDataSet.Tables("Country").DefaultView
        If dvSource.Count > 0 Then
            sRealName = dvSource(0).Item("RealName")
            oDataSet.Clear()
            Return sRealName
        Else
            Return ""
        End If
    End Function
    
    Protected Sub Page_PreRender(ByVal sender As Object, ByVal e As System.EventArgs)
        If GridView1.Rows.Count > 0 Then
            Call AddDropDowns()
        End If
    End Sub
    
    Protected Sub cbUseDefaultCountry_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If cbUseDefaultCountry.Checked = True Then
            ddlCountries.Enabled = True
        Else
            ddlCountries.Enabled = False
            ddlCountries.SelectedIndex = 0
        End If
    End Sub

    Protected Sub btnHelp_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If btnHelp.Text = "Show Help" Then
            pnlHelp.Visible = True
            btnHelp.Text = "Hide Help"
        Else
            pnlHelp.Visible = False
            btnHelp.Text = "Show Help"
        End If
    End Sub

    Protected Sub btnCreateAlias_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand
        If ddlAliasCountries.SelectedIndex > 0 And ddlUnmatchedCountries.Items.Count > 0 Then
            oConn.Open()
            Dim sCleanName = CleanUpCountryName(ddlUnmatchedCountries.SelectedItem.ToString)
            Dim sSQL As String = "INSERT INTO CountryAliases (AliasName, MatchName, RealName, ChangeDateTime, CustomerKey, UserKey) VALUES ('"
            sSQL = sSQL & ddlUnmatchedCountries.SelectedItem.ToString.Replace("'", "''") & "', '" & sCleanName & "', '" & ddlAliasCountries.SelectedItem.ToString.Replace("'", "''") & "', '" & Format(Now(), "dd-MMM-yyyy hh:mm:ss") & "', " & Session("CustomerKey") & ", " & Session("UserKey") & ")"
            oCmd = New SqlCommand(sSQL, oConn)
            Try
                oCmd.ExecuteNonQuery()
            Catch ex As Exception
            End Try
            ddlUnmatchedCountries.Items.RemoveAt(ddlUnmatchedCountries.SelectedIndex)
            ddlAliasCountries.SelectedIndex = 0
            If ddlUnmatchedCountries.Items.Count > 0 Then
                lblErrorProxy.Text = "Alias created - you have " & ddlUnmatchedCountries.Items.Count.ToString & " further alias(es) to create"
                lblError.Text = "Alias created - you have " & ddlUnmatchedCountries.Items.Count.ToString & " further alias(es) to create"
            Else
                lblErrorProxy.Text = "All aliases created - now revalidate your data"
                lblError.Text = "All aliases created - now revalidate your data"
            End If
        End If

    End Sub

    Protected Sub btnPopulateAliasTable_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim sRealName As String, sMatchName As String, sChar As Char
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand
        oConn.Open()
        For i As Integer = 1 To ddlCountries.Items.Count - 1
            sRealName = ddlCountries.Items(i).ToString
            sMatchName = CleanUpCountryName(sRealName)
            Dim sSQL As String = "INSERT INTO CountryAliases (AliasName, MatchName, RealName, ChangeDateTime, CustomerKey, UserKey) VALUES ('"
            sSQL = sSQL & sRealName.Replace("'", "''") & "', '" & sMatchName & "', '" & sRealName.Replace("'", "''") & "', '" & Format(Now(), "dd-MMM-yyyy hh:mm:ss") & "', 0, 0)"
            oCmd = New SqlCommand(sSQL, oConn)
            oCmd.ExecuteNonQuery()
        Next
    End Sub
    
    Protected Function CleanUpCountryName(ByVal sCountryName As String) As String
        Dim sChar As Char, sCleanName As String = ""
        For i As Integer = 1 To Len(sCountryName)
            sChar = Mid(sCountryName, i, 1).ToUpper
            If sChar >= "A" And sChar <= "Z" Then
                sCleanName = sCleanName & sChar
            End If
        Next
        Return sCleanName
    End Function

    Protected Sub btnManageAliases_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        pnlDeleteAlias.Visible = True
        pnlStart.Visible = False
        pnlCountryAliasing.Visible = False
        pnlGridView.Visible = False
        pnlMessage.Visible = False
        pnlHelp.Visible = False
        btnManageAliases.Visible = False
    End Sub

    Protected Sub btnDeleteAlias_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand
        oConn.Open()
        Dim sSQL As String = "DELETE FROM CountryAliases WHERE AliasName = '" & lbDeleteAlias.SelectedItem.Text.Replace("'", "''") & "' AND CustomerKey = " & Session("CustomerKey")
        oCmd = New SqlCommand(sSQL, oConn)
        Try
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
        End Try
        'SqlDataSource1.DataBind()
        lbDeleteAlias.DataBind()
        lblAliasLabel.Visible = False
        lblAlias.Visible = False
    End Sub

    Protected Sub lbDeleteAlias_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        lblAlias.Text = lbDeleteAlias.SelectedValue
        lblAliasLabel.Visible = True
        lblAlias.Visible = True
    End Sub
</script>

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Bulk Address Import</title>
</head>
<body>
    <form id="form1" runat="server" defaultfocus="FormUpload1" defaultbutton="btnReadAddresses" enctype="multipart/form-data">
    <div style="font-size: smaller; font-family: Arial, Verdana, Sans-Serif">
        <strong>
        BULK ADDRESS IMPORT &nbsp; &nbsp;
            <asp:Button ID="btnHelp" runat="server" Text="Show Help"
            Width="112px" OnClick="btnHelp_Click" /><br />
        </strong>
        <br />
        <asp:Panel ID="pnlHelp" runat="server" Width="100%" >
        The Bulk Address Import facility loads addresses into your shared or personal address book from a CSV (comma-separated
        variable) format file you supply. Follow the steps below to specify (a) the location of the address file and (b) how to use each part
        of the address.<br />
        <br />
        1. To import addresses from Excel you must first convert the data into CSV format.
        Open the spreadsheet, choose Save As, select
        CSV as the file type, specify a location
        and filename (eg C:\myaddresses.csv), then click Save. To import addresses from another source, eg Outlook, first ensure they are in a standard CSV format
        file.<br />
        <br />
        2. Click the <strong>Browse</strong> button to locate the address file on your local
        machine.<br />
        <br />
        3. If your data contains columns headings (ie the first row of data is the name
        or description of the column) click the <strong>Column Headings in Row 1</strong>
        check box.<br />
        <br />
        4. Click the <strong>Read Addresses</strong> button. The system reads and interprets your file, then displays
        the contents for you to check and confirm before it is entered into the Address
        Book. A message is displayed if there are problems reading
        the data.<br />
        <br />
        5. For each column of data you want to include in your address, choose the address
        field to associate with this column, using the dropdown list box at the top of each
        column.<br />
        <br />
        <strong>Specifying a Country<br />
        </strong>
        <br />
        For each address you must specify the country (eg U.K.).
        If your data does not include the destination country, you can set a default country to use
        for every address. If your data only includes the destination country when it differs from your native country, specify your own country as the Default
        Country; the system will use the default when no country is specified in your data.&nbsp; To use this feature click the <strong>Use Default Country</strong>
        check box then select the country from the dropdown list box.<br />
        <br />
        The country names in your data must match the country names in the dropdown list
        box (though they are case-insensitive). For example you must use U.K. rather than
        United Kingdom.<br />
        <br />
        6. Click <strong>Check Data</strong> if you want to check that the system can correctly process your
        data. The system checks your data and displays a message if a problem
        is found or required data is missing. Correct any errors and re-submit the data.<br />
        <br />
        7. Click <strong>Upload Addresses</strong> to copy your data to the Address Book.<br />
        <br />
        <br />
        </asp:Panel>

        <asp:Panel ID="pnlDeleteAlias" runat="server" Width="100%" >
            &nbsp;<strong>Alias Management</strong><br />
            <br />
            Select an alias, then click Delete Alias to remove.<br />
            <br />
            <asp:ListBox ID="lbDeleteAlias" runat="server" DataSourceID="SqlDataSource1" DataTextField="AliasName"
                DataValueField="RealName" AutoPostBack="True" OnSelectedIndexChanged="lbDeleteAlias_SelectedIndexChanged"></asp:ListBox><br />
            <br />
            <asp:Label ID="lblAliasLabel" runat="server" Visible="False">alias for: </asp:Label>
            <asp:Label ID="lblAlias" runat="server"></asp:Label><br />
            <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:AIMSRootConnectionString %>"
                SelectCommand="SELECT [CustomerKey], [AliasName], [RealName] FROM [CountryAliases] WHERE ([CustomerKey] = @CustomerKey) ORDER BY [AliasName]">
                <SelectParameters>
                    <asp:SessionParameter Name="CustomerKey" SessionField="CustomerKey" Type="Int32" />
                </SelectParameters>
            </asp:SqlDataSource>
            <br />
            <br />
            <asp:Button ID="btnDeleteAlias" runat="server" Text="Delete Alias" OnClick="btnDeleteAlias_Click" /></asp:Panel>
        
        <asp:Panel ID="pnlStart" runat="server" Width="100%" >
        Address&nbsp;CSV&nbsp;file: <asp:FileUpload ID="FileUpload1"
                        runat="server"
                        Width="350px" />
        &nbsp;&nbsp;<asp:CheckBox ID="cbColumnHeadingsInRow1" runat="server" Text="Column&nbsp;Headings&nbsp;in&nbsp;Row&nbsp;1" />&nbsp;
        <asp:Button ID="btnReadAddresses" runat="server" OnClick="btnReadAddresses_Click" Text="Read Addresses"
            Width="140px" />&nbsp;<br />
        &nbsp;<br />
        </asp:Panel>

        <asp:Panel ID="pnlCountryAliasing" runat="server" BackColor="#FFE0C0" BorderColor="#E0E0E0" BorderStyle="Solid" BorderWidth="1px" Width="100%"><strong>
            <br />
            &nbsp;Country Matching<br />
            <br />
        </strong>One or more countries in the data file could not be matched with a country
            name known by the system. For each unmatched
            country select the country name that should be used, then click Create Alias.<br />
            <br />
            &nbsp;For
            <asp:DropDownList ID="ddlUnmatchedCountries" runat="server">
            </asp:DropDownList>
            use
            <asp:DropDownList ID="ddlAliasCountries" runat="server">
            </asp:DropDownList>
            <asp:Button ID="btnCreateAlias" runat="server" OnClick="btnCreateAlias_Click" Text="Create Alias" />
            <asp:Button ID="btnPopulateAliasTable" runat="server" Text="Populate Alias Table" OnClick="btnPopulateAliasTable_Click" /><br />
            <br />
        </asp:Panel>

        <asp:Panel ID="pnlGridView" runat="server" Width="100%" >
        <table width="95%">
            <tr>
                <td style="width: 60px; height: 46px">
        Add&nbsp;Addresses&nbsp;To:</td>
                <td style="width: 374px; height: 46px">
                    <asp:RadioButtonList ID="rblAddressDestination" runat="server" RepeatDirection="Horizontal">
                        <asp:ListItem Selected="True">Shared Address Book</asp:ListItem>
                        <asp:ListItem>Personal Address Book</asp:ListItem>
                    </asp:RadioButtonList></td>
                <td style="height: 46px">
                    <asp:CheckBox ID="cbUseDefaultCountry"
            runat="server" AutoPostBack="True" OnCheckedChanged="cbUseDefaultCountry_CheckedChanged"
            Text="Use Default Country" />
        <asp:DropDownList ID="ddlCountries" runat="server" Enabled="False">
        </asp:DropDownList></td>
            </tr>
        </table>
        <br />
        <asp:Button ID="btnCheckData" runat="server" Text="Check Data" OnClick="btnCheckData_Click" />
        <asp:Button ID="btnUploadAddresses" runat="server" Text="Upload Addresses" OnClick="btnUploadAddresses_Click" />&nbsp;
            <asp:Label ID="lblErrorProxy" runat="server" ForeColor="Red"></asp:Label><br />
        <br />
            <asp:GridView ID="GridView1" runat="server" Font-Size="XX-Small" Width="100%" />
        </asp:Panel>

        <asp:Panel ID="pnlMessage" runat="server" Width="100%">
            <asp:Label ID="lblMessage" runat="server" Text=""></asp:Label>
        </asp:Panel>

        <asp:Panel ID="pnlFinished" runat="server" Width="100%">
            <br />
            <asp:Label ID="lblSuccess" runat="server" ForeColor="Green"></asp:Label>
            <br />
            <asp:Button ID="btnClose" runat="server" Text="Close" OnClientClick="window.close()" />
            &nbsp;&nbsp;
            <asp:Button ID="btnManageAliases" runat="server" OnClick="btnManageAliases_Click"
                Text="Manage Aliases" /></asp:Panel>
        <br />
        <asp:Label ID="lblError" runat="server" ForeColor="Red"></asp:Label>
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
