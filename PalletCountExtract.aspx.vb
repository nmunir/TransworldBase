Imports System.Data.SqlClient
Imports System.Data
Imports System.Data.DataSet
Imports System.IO
Imports System.Data.OleDb
Imports System.Collections.Generic

Partial Class PalletCountExtract

    ' TO DO
    ' sort out LastUpdatedBy

    Inherits System.Web.UI.Page

    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString()
    Private glstIgnoreStrings As List(Of String)
    Private gdictCustomerNameMapping As Dictionary(Of String, String)
    Private gdictMonths As New Dictionary(Of String, String)
    Private gdictYears As New Dictionary(Of String, String)

    Const STRING_NOT_MATCHED As String = "not matched!"
    Const STRING_UNPROCESSED As String = "UNPROCESSED"
    Const SPREADSHEET_PATH As String = "./PalletCountReport/"
    Const SPREADSHEET_BACKUP_PATH As String = "./PalletCountReport/backup/"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)
        If Not IsNumeric(Session("SiteKey")) Then
            Server.Transfer("session_expired.aspx")
            Exit Sub
        End If
        Call InitDateDictionaries()
        If Not IsPostBack Then
            Call AppLoad()
        End If
    End Sub

    Protected Sub AppLoad()
        Call InitStructures()
        Call GetSavedDataReport()
        Call BindMappingGrid()
        Call BindIgnoreGrid()
        Call SetDateManually(False)
        If GetFileToProcess() Then
            btnReadExcelFile.Enabled = True
            Call DeduceDateFromFilename()
        Else
            btnReadExcelFile.Enabled = False
        End If
        gvRawData.DataSource = Nothing
        gvRawData.DataBind()
        pnlDataIntegrityChecks.Visible = False
    End Sub

    Protected Sub GetSavedDataReport()
        If ExecuteQueryToDataTable("SELECT COUNT (*) FROM PalletUsage").Rows(0).Item(0) = 0 Then
            lblMostRecentSavedData.Text = "No data found"
            Exit Sub
        End If
        Dim nHighestYear As Int32 = ExecuteQueryToDataTable("SELECT MAX(Year) FROM PalletUsage").Rows(0).Item(0)
        For i As Int32 = 12 To 0 Step -1
            If ExecuteQueryToDataTable("SELECT COUNT (*) FROM PalletUsage WHERE Year = " & nHighestYear & " AND Month = " & i.ToString).Rows(0).Item(0) > 0 Then
                lblMostRecentSavedData.Text = "Most recent pallet count data extracted: <b>" & nHighestYear & "/" & i.ToString & "</b>"
                Exit For
            End If
        Next
        Dim dtMostRecentlyEntered As DateTime = ExecuteQueryToDataTable("SELECT MAX(LastUpdatedOn) FROM PalletUsage").Rows(0).Item(0)
        lblMostRecentSavedData.Text += "; last extraction: <b>" & dtMostRecentlyEntered.ToString("dd-MMM-yyyy") & "</b>"

        Dim oDataTable As DataTable
        oDataTable = ExecuteQueryToDataTable("SELECT DISTINCT Month, Year FROM PalletUsage ORDER BY Year, Month")
        lblSavedData.Text = "Database already contains data for period(s): "
        For Each dr As DataRow In oDataTable.Rows
            lblSavedData.Text += "<b>" & dr("Year") & "/" & dr("Month") & "</b>; "
        Next
        lblCustomers.Text = "Database already contains data for customer(s): "
        oDataTable = ExecuteQueryToDataTable("SELECT DISTINCT pu.CustomerKey, CustomerAccountCode FROM PalletUsage pu INNER JOIN Customer c ON pu.CustomerKey = c.CustomerKey ORDER BY CustomerAccountCode")
        For Each dr As DataRow In oDataTable.Rows
            lblCustomers.Text += "<b>" & dr("CustomerAccountCode") & "</b>; "
        Next
    End Sub

    Protected Sub InitDateDictionaries()
        gdictMonths.Add("jan", 1)
        gdictMonths.Add("feb", 2)
        gdictMonths.Add("mar", 3)
        gdictMonths.Add("apr", 4)
        gdictMonths.Add("may", 5)
        gdictMonths.Add("jun", 6)
        gdictMonths.Add("jul", 7)
        gdictMonths.Add("aug", 8)
        gdictMonths.Add("sep", 9)
        gdictMonths.Add("oct", 10)
        gdictMonths.Add("nov", 11)
        gdictMonths.Add("dec", 12)

        gdictYears.Add("2008", 2008)
        gdictYears.Add("08", 2008)
        gdictYears.Add("2009", 2009)
        gdictYears.Add("09", 2009)
        gdictYears.Add("2010", 2010)
        gdictYears.Add("10", 2010)
        gdictYears.Add("2011", 2011)
        gdictYears.Add("11", 2011)
        gdictYears.Add("2012", 2012)
        gdictYears.Add("12", 2012)
        gdictYears.Add("2013", 2013)
        gdictYears.Add("13", 2013)
        gdictYears.Add("2014", 2014)
        gdictYears.Add("14", 2014)
        gdictYears.Add("2015", 2015)
        gdictYears.Add("15", 2015)
        gdictYears.Add("2016", 2016)
        gdictYears.Add("16", 2016)
        gdictYears.Add("2017", 2017)
        gdictYears.Add("17", 2017)
        gdictYears.Add("2018", 2018)
        gdictYears.Add("18", 2018)
        gdictYears.Add("2019", 2019)
        gdictYears.Add("19", 2019)
        gdictYears.Add("2020", 2020)
        gdictYears.Add("20", 2020)
    End Sub

    Protected Sub InitStructures()
        glstIgnoreStrings = New List(Of String)
        Dim dt1 As DataTable = ExecuteQueryToDataTable("SELECT IgnoreString FROM PalletReportIgnoreStrings")
        For Each dr1 As DataRow In dt1.Rows
            glstIgnoreStrings.Add(dr1(0))
        Next
        gdictCustomerNameMapping = New Dictionary(Of String, String)()
        Dim dt2 As DataTable = ExecuteQueryToDataTable("SELECT ReportString, ClientName FROM PalletReportCustomerNameMapping ORDER BY LEN(ReportString) DESC")
        For Each dr2 As DataRow In dt2.Rows
            gdictCustomerNameMapping.Add(dr2(0), dr2(1))
        Next
        Dim oListItemCollection As ListItemCollection = ExecuteQueryToListItemCollection("SELECT CustomerAccountCode, CustomerKey FROM Customer WHERE CustomerStatusId = 'ACTIVE' AND DeletedFlag = 'N' AND ISNULL(AccountHandlerKey,0) > 0 ORDER BY CustomerAccountCode", "CustomerAccountCode", "CustomerKey")
        ddlAimsAccount.Items.Clear()
        ddlAimsAccount.Items.Add(New ListItem("- please select -", 0))
        For Each li As ListItem In oListItemCollection
            ddlAimsAccount.Items.Add(li)
        Next
        ddlAimsAccount.Items.Add(New ListItem(STRING_UNPROCESSED, 9999))
    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnReadExcelFile.Click
        Call ReadExcel()
        pnlDataIntegrityChecks.Visible = True
    End Sub

    Protected Sub SetDateManually(ByVal bState As Boolean)
        If bState Then
            cbSetDateManually.Text = "Use the following date:"
        Else
            cbSetDateManually.Text = "That's wrong!"
        End If
        ddlMonth.Visible = bState
        ddlMonth.SelectedIndex = 0
        ddlYear.Visible = bState
        ddlYear.SelectedIndex = 0
        btnSaveDate.Visible = bState
    End Sub

    Protected Function GetFileToProcess() As Boolean
        GetFileToProcess = False
        Dim sDirectory As String = MapPath(SPREADSHEET_PATH)
        Dim sTargetFile As String = String.Empty
        Dim collFiles As System.Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Computer.FileSystem.GetFiles(sDirectory)
        Dim nFileCount As Int32 = 0
        For Each sFilename As String In collFiles
            Dim sExtension As String = Path.GetExtension(sFilename)
            If sExtension = ".xls" Then
                nFileCount += 1
                sTargetFile = sFilename
            End If
        Next
        If nFileCount = 0 Then
            For Each sFilename As String In collFiles
                Dim sExtension As String = Path.GetExtension(sFilename)
                If sExtension = ".xlsx" Then
                    nFileCount += 1
                    sTargetFile = sFilename
                End If
            Next
        End If

        lblFileToProcess.Text = "No file found!"

        If nFileCount = 0 Then
            WebMsgBox.Show("No suitable files found.")
            Exit Function
        ElseIf nFileCount = 1 Then
            GetFileToProcess = True
            psFileName = sTargetFile
            lblFileToProcess.Text = psFileName
        ElseIf nFileCount > 1 Then
            WebMsgBox.Show("Two or more possible candidate files found. Please supply just one file for processing.")
            Exit Function
        End If
    End Function

    Protected Sub ReadExcel()
        Dim nParseFail As Int32 = 0
        Dim nUnprocessed As Int32 = 0
        Call InitStructures()
        Dim sTargetFile As String = psFileName

        'Dim sDirectory As String = MapPath("")
        'Dim sTargetFile As String = String.Empty
        'Dim collFiles As System.Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Computer.FileSystem.GetFiles(sDirectory)
        'Dim nFileCount As Int32 = 0
        'For Each sFilename As String In collFiles
        '    Dim sExtension As String = Path.GetExtension(sFilename)
        '    If sExtension = ".xls" Then
        '        nFileCount += 1
        '        sTargetFile = sFilename
        '    End If
        'Next
        'If nFileCount = 0 Then
        '    For Each sFilename As String In collFiles
        '        Dim sExtension As String = Path.GetExtension(sFilename)
        '        If sExtension = ".xlsx" Then
        '            nFileCount += 1
        '            sTargetFile = sFilename
        '        End If
        '    Next
        'End If
        'If nFileCount = 0 Then
        '    ' NO FILES FOUND
        '    Exit Sub
        'ElseIf nFileCount > 1 Then
        '    ' TOO MANY FILES
        '    Exit Sub
        'End If

        Dim connString As String = String.Empty
        If Path.GetExtension(sTargetFile) = ".xls" Then
            connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & sTargetFile & "';Extended Properties=Excel 8.0"
        Else
            connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & sTargetFile & "';Extended Properties=Excel 12.0"
        End If

        Dim oledbConn As OleDbConnection = New OleDbConnection(connString)

        Try
            oledbConn.Open()
            Dim cmd As OleDbCommand = New OleDbCommand("SELECT Account, Total, Account1, Total1 FROM [Display Report$]", oledbConn)
            Dim oleda As OleDbDataAdapter = New OleDbDataAdapter()
            oleda.SelectCommand = cmd
            Dim dtRaw As DataTable = New DataTable
            oleda.Fill(dtRaw)
            Dim dtInterim As New DataTable
            dtInterim.Columns.Add(New DataColumn("Account", GetType(String)))
            dtInterim.Columns.Add(New DataColumn("Total", GetType(Double)))
            dtInterim.Columns.Add(New DataColumn("CustomerAccountCode", GetType(String)))

            Dim drInterim As DataRow

            For Each dr As DataRow In dtRaw.Rows
                If Not IgnoreRow(dr("Account") & "") Then
                    drInterim = dtInterim.NewRow
                    drInterim("Account") = dr("Account")
                    drInterim("Total") = dr("Total")
                    dtInterim.Rows.Add(drInterim)
                End If
            Next

            For Each dr As DataRow In dtRaw.Rows
                If Not IgnoreRow(dr("Account1") & "") Then
                    drInterim = dtInterim.NewRow
                    drInterim("Account") = dr("Account1")
                    drInterim("Total") = dr("Total1")
                    dtInterim.Rows.Add(drInterim)
                End If
            Next

            For Each dr As DataRow In dtInterim.Rows
                For Each kv As KeyValuePair(Of String, String) In gdictCustomerNameMapping
                    If dr("Account").ToString.ToLower.Contains(kv.Key.ToLower) Then
                        dr("CustomerAccountCode") = kv.Value
                        Exit For
                    Else
                        dr("CustomerAccountCode") = STRING_NOT_MATCHED
                    End If
                Next
                If dr("CustomerAccountCode") = STRING_NOT_MATCHED Then
                    nParseFail += 1
                End If
                If dr("CustomerAccountCode") = STRING_UNPROCESSED Then
                    nUnprocessed += 1
                End If
            Next

            ViewState("EE_Data") = dtInterim
            gvRawData.DataSource = dtInterim
            gvRawData.DataBind()
        Catch ex As Exception
            lblMessage.Text = ex.Message
        Finally
            oledbConn.Close()
        End Try
        If nParseFail = 0 Then
            btnSaveData.Enabled = True
            Dim sMessage As String = "All customers matched okay!"
            If nUnprocessed > 0 Then
                If nUnprocessed = 1 Then
                    sMessage += " 1 customer was mapped to 'UNPROCESSED'."
                Else
                    sMessage += " " & nUnprocessed.ToString & " customers were mapped to 'UNPROCESSED'."
                End If
            End If
            WebMsgBox.Show(sMessage)
        Else
            btnSaveData.Enabled = False
            Dim sPlural As String = String.Empty
            If nParseFail > 1 Then
                sPlural = "s"
            End If
            WebMsgBox.Show(nParseFail & " customer " & sPlural & " could not be matched. Modify the Account Identification table until all customers are matched.")
        End If
    End Sub

    Protected Function IgnoreRow(ByVal sText As String) As Boolean
        If sText = String.Empty Then
            IgnoreRow = True
            Exit Function
        End If
        IgnoreRow = False
        For Each s As String In glstIgnoreStrings
            If sText.ToLower.Contains(s.ToLower) Then
                IgnoreRow = True
                Exit For
            End If
        Next
    End Function

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
            WebMsgBox.Show("Error in ExecuteQueryToListItemCollection executing: " & sQuery & " : " & ex.Message)
        Finally
            oConn.Close()
        End Try
        ExecuteQueryToListItemCollection = oListItemCollection
    End Function

    Protected Sub btnAddText_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAddAccountMatchingText.Click
        If tbAccountString.Text = String.Empty Then
            WebMsgBox.Show("Please specify some text.")
            Exit Sub
        End If
        If ddlAimsAccount.SelectedIndex = 0 Then
            WebMsgBox.Show("Please select an AIMS account.")
            Exit Sub
        End If
        Call ExecuteQueryToDataTable("DELETE FROM PalletReportCustomerNameMapping WHERE ReportString = '" & tbAccountString.Text.Replace("'", "''") & "'")
        Call ExecuteQueryToDataTable("INSERT INTO PalletReportCustomerNameMapping (ReportString, ClientName) VALUES ('" & tbAccountString.Text.Replace("'", "''") & "', '" & ddlAimsAccount.SelectedItem.Text & "')")
        Call BindMappingGrid()
        tbAccountString.Text = String.Empty
        ddlAimsAccount.SelectedIndex = 0
    End Sub

    Protected Sub BindMappingGrid()
        Dim oDataTable As DataTable = ExecuteQueryToDataTable("SELECT [id], ReportString, ClientName FROM PalletReportCustomerNameMapping ORDER BY ReportString")
        gvMapping.DataSource = oDataTable
        gvMapping.DataBind()
    End Sub

    Protected Sub BindIgnoreGrid()
        Dim oDataTable As DataTable = ExecuteQueryToDataTable("SELECT [id], IgnoreString FROM PalletReportIgnoreStrings ORDER BY IgnoreString")
        gvIgnoreText.DataSource = oDataTable
        gvIgnoreText.DataBind()
    End Sub

    Protected Sub lnkbtnRemoveMapping_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim lnkbtn As LinkButton = sender
        Call ExecuteQueryToDataTable("DELETE FROM PalletReportCustomerNameMapping WHERE [id] = " & lnkbtn.CommandArgument)
        Call BindMappingGrid()
    End Sub

    Protected Sub cbSetDateManually_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbSetDateManually.CheckedChanged
        Dim cb As CheckBox = sender
        Call SetDateManually(cb.Checked)
    End Sub

    Protected Sub DeduceDateFromFilename()
        Dim sFilename = Path.GetFileNameWithoutExtension(psFileName).ToLower
        For Each sMonth In gdictMonths
            If sFilename.Contains(sMonth.Key) Then
                pnMonth = sMonth.Value
                Exit For
            End If
        Next
        For Each sYear In gdictYears
            If sFilename.Contains(sYear.Key) Then
                pnYear = sYear.Value
                Exit For
            End If
        Next
        If pnMonth > 0 And pnYear > 0 Then
            'lblLegendThisSpreadsheetContainsDataFor.Visible = True
            lblDate.Text = pnYear.ToString & "/" & pnMonth.ToString
            lblDate.ForeColor = Drawing.Color.Empty
        Else
            lblLegendThisSpreadsheetContainsDataFor.Visible = False
            lblDate.Text = "Could not deduce the date for which this spreadsheet contains data from the filename."
            lblDate.ForeColor = Drawing.Color.Red
            Call SetDateManually(True)
        End If
    End Sub

    Protected Sub btnSaveDate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSaveDate.Click
        If ddlMonth.SelectedIndex > 0 And ddlYear.SelectedIndex > 0 Then
            pnMonth = ddlMonth.SelectedValue
            pnYear = ddlYear.SelectedValue
            lblDate.Text = pnMonth.ToString & "/" & pnYear.ToString
            lblLegendThisSpreadsheetContainsDataFor.Visible = True
            Call SetDateManually(False)
            cbSetDateManually.Checked = False
        Else
            WebMsgBox.Show("Please select month and year.")
        End If
    End Sub

    Protected Sub btnAddIgnoreText_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAddIgnoreText.Click
        If tbIgnoreText.Text = String.Empty Then
            WebMsgBox.Show("Please specify some text.")
            Exit Sub
        End If
        Call ExecuteQueryToDataTable("DELETE FROM PalletReportIgnoreStrings WHERE IgnoreString = '" & tbIgnoreText.Text.Replace("'", "''"))
        Call ExecuteQueryToDataTable("INSERT INTO PalletReportIgnoreStrings (IgnoreString) VALUES ('" & tbIgnoreText.Text.Replace("'", "''") & "')")
        Call BindIgnoreGrid()
        tbIgnoreText.Text = String.Empty
    End Sub

    Protected Sub lnkbtnRemoveIgnoreText_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim lnkbtn As LinkButton = sender
        Call ExecuteQueryToDataTable("DELETE FROM PalletReportIgnoreStrings WHERE [id] = " & lnkbtn.CommandArgument)
        Call BindIgnoreGrid()
    End Sub

    Protected Sub btnSaveData_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSaveData.Click
        If pnMonth = 0 Or pnYear = 0 Then
            WebMsgBox.Show("Please specify the month and year for which this spreadsheet holds data.")
        Else
            Call ExecuteQueryToDataTable("DELETE FROM PalletUsage WHERE Month = " & pnMonth & " AND Year = " & pnYear)
            Dim dtData As DataTable = ViewState("EE_Data")
            For Each dr As DataRow In dtData.Rows
                If Not (dr("CustomerAccountCode").ToString.Contains(STRING_NOT_MATCHED) Or dr("CustomerAccountCode").ToString.Contains(STRING_UNPROCESSED)) Then
                    Dim nCustomerKey As Int32 = ExecuteQueryToDataTable("SELECT CustomerKey FROM Customer WHERE CustomerAccountCode = '" & dr("CustomerAccountCode").ToString.Trim & "'").Rows(0).Item(0)
                    Dim sSQL As String = "INSERT INTO PalletUsage (CustomerKey, Month, Year, ReportIdentifier, Quantity, LastUpdatedBy, LastUpdatedOn) VALUES (" & nCustomerKey & ", " & pnMonth & ", " & pnYear & ", '" & dr("Account").ToString.Replace("'", "''") & "', " & dr("Total") & ", " & Session("UserKey") & ", GETDATE())"
                    Call ExecuteQueryToDataTable(sSQL)
                End If
            Next
            Dim guidGUID As Guid = Guid.NewGuid
            Dim sDirectory As String = MapPath(SPREADSHEET_BACKUP_PATH)
            Dim sFileNameWithoutExtension As String = Path.GetFileNameWithoutExtension(psFileName).ToLower
            Dim sFileExtension As String = Path.GetExtension(psFileName)
            Dim sBackupFilename As String = sDirectory & sFileNameWithoutExtension & "_" & guidGUID.ToString & sFileExtension
            My.Computer.FileSystem.MoveFile(psFileName, sBackupFilename)

            'My.Computer.FileSystem.DeleteFile(psFileName)
            WebMsgBox.Show("Data saved.  Spreadsheet backed up and deleted from \\SPRINT_DATA2\PalletCountReport.")
            lblFileToProcess.Text = "File removed"
            btnReadExcelFile.Enabled = False
            btnSaveData.Enabled = False
        End If
    End Sub

    Protected Sub gvRawData_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gvRawData.RowDataBound
        Dim gvr As GridViewRow = e.Row
        If gvr.RowType = DataControlRowType.DataRow Then
            If gvr.Cells(2).Text.Contains(STRING_NOT_MATCHED) Then
                gvr.Cells(2).ForeColor = Drawing.Color.Red
                gvr.Cells(2).Font.Bold = True
            End If
            If gvr.Cells(2).Text.Contains(STRING_UNPROCESSED) Then
                gvr.Cells(2).ForeColor = Drawing.ColorTranslator.FromOle(RGB(241, 60, 252))
                gvr.Cells(2).Font.Bold = True
            End If
        End If
    End Sub

    Protected Sub lnkbtnRecheckFile_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkbtnRecheckFile.Click
        Call AppLoad()
    End Sub

    Protected Sub lnkbtnRefreshSavedDataReport_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkbtnRefreshSavedDataReport.Click
        Call GetSavedDataReport()
    End Sub

    Property pnMonth() As Int32
        Get
            Dim o As Object = ViewState("EE_Month")
            If o Is Nothing Then
                Return 0
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Int32)
            ViewState("EE_Month") = Value
        End Set
    End Property

    Property pnYear() As Int32
        Get
            Dim o As Object = ViewState("EE_Year")
            If o Is Nothing Then
                Return 0
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Int32)
            ViewState("EE_Year") = Value
        End Set
    End Property

    Property psFileName() As String
        Get
            Dim o As Object = ViewState("EE_FileName")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("EE_FileName") = Value
        End Set
    End Property

    Protected Sub gvMapping_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gvMapping.RowDataBound
        Dim gvr As GridViewRow = e.Row
        If gvr.RowType = DataControlRowType.DataRow Then
            If gvr.Cells(2).Text.Contains(STRING_UNPROCESSED) Then
                gvr.Cells(2).ForeColor = Drawing.ColorTranslator.FromOle(RGB(241, 60, 252))
                gvr.Cells(2).Font.Bold = True
            End If
        End If
    End Sub
End Class
