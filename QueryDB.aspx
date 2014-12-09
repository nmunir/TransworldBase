<%@ Page Language="VB" Theme="AIMSDefault" ValidateRequest="false" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data.SqlTypes" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Collections.Generic" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<script runat="server">

    ' dynamic PARAM format
    ' #param1,Label,X,init value, notes (advisory text),#param1
    ' where X is parameter type for checking: T=textbox with any text, N=textbox, numeric only, D=dropdown list box with values
    
    ' TO DO
    ' when selecting another report, clear error label
    ' hide fields when re-hiding from check box
    ' enlarge key user-entered fields

    Const PARAM_LABEL As Integer = 1
    Const PARAM_TYPE As Integer = 2
    Const PARAM_INIT_VALUE As Integer = 3
    Const PARAM_NOTES As Integer = 4
    Const SECRET As String = "tw140rn"
   
   
    Dim gsConn As String = ConfigurationManager.AppSettings("AIMSRootConnectionString")
    Dim oDataTable As New DataTable

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsNumeric(Session("CustomerKey")) Then
            Server.Transfer("session_expired.aspx")
            Exit Sub
        End If
        If Not IsPostBack Then
            tbQuery.Attributes.Add("onkeypress", "return clickButton(event,'" + btnGo.ClientID + "')")
            tbRows.Attributes.Add("onkeypress", "return clickButton(event,'" + btnGo.ClientID + "')")
            Call HideAllPanels()
            Call PopulateQueryList(String.Empty)
            Call GetTags()
            Call InitCustomerDropdown()
        End If
        tbQuery.Focus()
        Call SetTitle()
    End Sub
   
    Protected Sub SetTitle()
        Dim sTitle As String = Session("SiteTitle")
        If sTitle <> String.Empty Then
            sTitle += " - "
        End If
        Page.Header.Title = sTitle & "Query the database"
    End Sub
   
    Protected Sub GetTags()
        Dim oListItemCollection As ListItemCollection = ExecuteQueryToListItemCollection("SELECT Tags FROM QueryDB WHERE ISNULL(IsDeleted,0) = 0", "Tags", "Tags")
        Dim lstTagList As New List(Of String)
        For Each liTags As ListItem In oListItemCollection
            Dim sTags() As String = liTags.Text.Split(" ")
            For Each sTag As String In sTags
                If Not lstTagList.Contains(sTag) Then
                    lstTagList.Add(sTag)
                End If
            Next
        Next
        lstTagList.Sort()
        ddlTags.Items.Clear()
        ddlTags.Items.Add(New ListItem("- please select -", 0))
        For Each sTag As String In lstTagList
            ddlTags.Items.Add(New ListItem(sTag, sTag))
        Next
    End Sub
   
    Protected Sub InitCustomerDropdown()
        Dim sSQL As String
        sSQL = "SELECT DISTINCT c.CustomerAccountCode, lp.CustomerKey FROM LogisticProduct lp INNER JOIN Customer c ON lp.CustomerKey = c.CustomerKey WHERE c.DeletedFlag = 'N' AND CustomerStatusId = 'ACTIVE' ORDER BY CustomerAccountCode"
        If cbIncludeAccountsWithNoProducts.Checked Then
            sSQL = "SELECT CustomerAccountCode, CustomerKey FROM Customer WHERE DeletedFlag = 'N' AND CustomerStatusId = 'ACTIVE' ORDER BY CustomerAccountCode"
        End If
        If cbIncludeSuspendedDeletedAccounts.Checked Then
            sSQL = "SELECT CustomerAccountCode, CustomerKey FROM Customer ORDER BY CustomerAccountCode"
        End If
        Dim oListItemCollection As ListItemCollection = ExecuteQueryToListItemCollection(sSQL, "CustomerAccountCode", "CustomerKey")
        ddlCustomer.Items.Clear()
        ddlCustomer.Items.Add(New ListItem("- please select -", 0))
        For Each li As ListItem In oListItemCollection
            ddlCustomer.Items.Add(li)
        Next
    End Sub
   
    Protected Sub PopulateQueryList(ByVal sTag As String)
        Dim oDataReader As SqlDataReader = Nothing
        Dim oConn As New SqlConnection(gsConn)
        Dim sSQL As String
        If sTag <> String.Empty Then
            sSQL = "SELECT * FROM QueryDB WHERE Tags LIKE '%" & sTag & "%' AND ISNULL(IsDeleted,0) = 0 ORDER BY Title"
        Else
            sSQL = "SELECT * FROM QueryDB WHERE ISNULL(IsDeleted,0) = 0 ORDER BY Title"
        End If
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            lbQueries.Items.Clear()
            If oDataReader.HasRows Then
                While oDataReader.Read
                    lbQueries.Items.Add(New ListItem(oDataReader("Title"), oDataReader("id")))
                End While
            Else
                WebMsgBox.Show("No queries defined")
            End If
        Catch ex As Exception
            WebMsgBox.Show("Error in PopulateQueryList: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
   
    Protected Sub HideAllPanels()
        pnlData.Visible = False
    End Sub

    Protected Function bIsAllowedQuery() As Boolean
        bIsAllowedQuery = False
        psQuery = tbQuery.Text.Trim
        If pnQueryLength > 0 AndAlso psQuery.Length = pnQueryLength Then
            bIsAllowedQuery = True
        End If
    End Function
   
    Protected Sub btnGo_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Page.Validate()
        If Page.IsValid Then
            psQuery = tbQuery.Text.Trim
            pnlData.Visible = True
            If psQuery <> String.Empty Then
                If bIsAllowedQuery() OrElse Not (psQuery.ToLower.Contains("update") Or psQuery.ToLower.Contains("insert") Or psQuery.ToLower.Contains("delete")) Then
                    psQuery = SubstituteCustomer(psQuery)
                    psQuery = SubstituteDate(psQuery)
                    psQuery = SubstituteFromDate(psQuery)
                    psQuery = SubstituteDateRange(psQuery)
                    psQuery = SubstituteParam1(psQuery)
                    psQuery = SubstituteParam2(psQuery)
                    lblActualQuery.Text = psQuery
                    gvDisplay.PageIndex = 0
                    If Not psQuery.Contains("1-jan-1900") Then
                        Call BindGrid()
                    End If
                Else
                    WebMsgBox.Show("SELECT queries only, thank you!" & " (" & psQuery.Length & ")")
                End If
            End If
            Call RecordUsage("display")
        End If
    End Sub

    Protected Sub GetUsage()
        Dim sSQL As String = "SELECT ISNULL(CountDisplay, 0) 'CountDisplay', ISNULL(CountExport, 0) 'CountExport' FROM QueryDB WHERE [id] = " & lbQueries.SelectedValue
        Dim dr As DataRow = ExecuteQueryToDataTable(sSQL).Rows(0)
        pnCountDisplay = dr(0)
        pnCountExport = dr(1)
    End Sub
    
    Protected Sub RecordUsage(sType As String)
        Dim sSQL As String
        Dim sFragment As String
        Call GetUsage()
        If sType = "display" Then
            pnCountDisplay += 1
            sFragment = ", CountDisplay = " & (pnCountDisplay).ToString
        Else
            pnCountExport += 1
            sFragment = ", CountExport = " & (pnCountExport).ToString
        End If
        sSQL = "UPDATE QueryDB SET LastUsedOn = GETDATE(), LastUsedBy = " & Session("UserKey") & sFragment & "  WHERE [id] = " & lbQueries.SelectedValue
        Call ExecuteQueryToDataTable(sSQL)
    End Sub
    
    Protected Sub btnDisplay_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call BindGrid()
    End Sub
   
    Protected Function SubstituteCustomer(ByVal sText As String) As String
        If trCustomer.Visible Then
            If ddlCustomer.SelectedIndex <= 0 Then
                WebMsgBox.Show("No customer selected")
                SubstituteCustomer = sText.Replace("#CustomerKey", 0)
            Else
                SubstituteCustomer = sText.Replace("#CustomerKey", ddlCustomer.SelectedValue)
            End If
        Else
            SubstituteCustomer = sText
        End If
    End Function
   
    Protected Function SubstituteFromDate(ByVal sText As String) As String
        If trFromDate.Visible Then
            If Not IsDate(tbFromDate.Text) Then
                WebMsgBox.Show("Date not recognised")
                SubstituteFromDate = sText.Replace("#FromDate", "1-jan-1900")
            Else
                SubstituteFromDate = sText.Replace("#FromDate", tbFromDate.Text & " 00:00:01")
            End If
        Else
            SubstituteFromDate = sText
        End If
    End Function
   
    Protected Function SubstituteDate(ByVal sText As String) As String  ' warning - this code is WRONG and should check the date RANGE but I've not removed it since I don't think any query uses it - CN 21NOV11
        If trDate.Visible Then
            If Not IsDate(tbDate.Text) Then
                WebMsgBox.Show("Date not recognised")
                SubstituteDate = sText.Replace("#Date", "1-jan-1900")
            Else
                SubstituteDate = sText.Replace("#Date", tbDate.Text)
            End If
        Else
            SubstituteDate = sText
        End If
    End Function
   
    Protected Function SubstituteDateRange(ByVal sText As String) As String
        Dim sTemp As String
        If trDateRange.Visible Then
            If Not (IsDate(tbStartDate.Text) And IsDate(tbEndDate.Text)) Then
                WebMsgBox.Show("FROM or TO date not recognised")
                sTemp = sText.Replace("#EndDate", "1-jan-1900")
                SubstituteDateRange = sTemp.Replace("#StartDate", "1-jan-1900")
            Else
                If Date.Parse(tbStartDate.Text) > Date.Parse(tbEndDate.Text) Then
                    WebMsgBox.Show("FROM date must precede TO date")
                    sTemp = sText.Replace("#EndDate", "1-jan-1900")
                    SubstituteDateRange = sTemp.Replace("#StartDate", "1-jan-1900")
                Else
                    sTemp = sText.Replace("#EndDate", tbEndDate.Text & " 23:59:59")
                    SubstituteDateRange = sTemp.Replace("#StartDate", tbStartDate.Text & " 00:00:01")
                End If
            End If
        Else
            SubstituteDateRange = sText
        End If
    End Function
   
    Protected Function SubstituteParam1(ByVal sText As String) As String
        If trParam1.Visible Then
            If tbParam1.Visible Then
                SubstituteParam1 = sText.Replace("#param1", tbParam1.Text)
            ElseIf ddlParam1.Visible Then
                SubstituteParam1 = sText.Replace("#param1", ddlParam1.SelectedValue)
            End If
        Else
            SubstituteParam1 = sText
        End If
    End Function
   
    Protected Function SubstituteParam2(ByVal sText As String) As String
        If trParam2.Visible Then
            If tbParam2.Visible Then
                SubstituteParam2 = sText.Replace("#param2", tbParam2.Text)
            ElseIf ddlParam2.Visible Then
                SubstituteParam2 = sText.Replace("#param2", ddlParam2.SelectedValue)
            End If
        Else
            SubstituteParam2 = sText
        End If
    End Function
   
    Protected Sub BindGrid()
        If psQuery <> String.Empty Then
            Dim oConn As New SqlConnection(gsConn)
            lblMessage.Text = String.Empty
            Try
                Dim sSQL As String = psQuery
                Dim oAdapter As New SqlDataAdapter(sSQL, oConn)
                If pnTimeout > 0 Then
                    oAdapter.SelectCommand.CommandTimeout = pnTimeout
                    Server.ScriptTimeout = pnTimeout
                Else
                    Server.ScriptTimeout = 90
                End If
                oAdapter.Fill(oDataTable)
                'If oDataTable.Rows.Count > 0 Then
                gvDisplay.PageSize = tbRows.Text
                gvDisplay.DataSource = oDataTable
                gvDisplay.DataBind()
                lblRowCount.Text = oDataTable.Rows.Count & " record(s)"
                ' End If
            Catch ex As Exception
                WebMsgBox.Show("Error in BindGrid: " & ex.ToString)
                lblMessage.Text = ex.ToString & " QUERY: " & psQuery
            Finally
                oConn.Close()
            End Try
        End If
    End Sub

    Protected Sub gvDisplay_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs)
        gvDisplay.PageIndex = e.NewPageIndex
        Call BindGrid()
    End Sub
   
    Protected Sub btnDo_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim c As String = tbDo.Text.Trim.ToLower
        If c.StartsWith("list") Then
            For Each f As String In My.Computer.FileSystem.GetFiles(Server.MapPath(""))
                lblList.Text = lblList.Text & f & "<br />"
            Next
            Exit Sub
        End If
        If c.StartsWith("read ") Then
            c = c.Substring(4)
            'lblError.Text = HttpUtility.HtmlEncode(My.Computer.FileSystem.ReadAllText(c))
            'lblError.Text.Replace(Environment.NewLine, "<br />")
            Try
                Using sr As StreamReader = New StreamReader(c)
                    Dim l As String
                    Do
                        l = sr.ReadLine
                        lblList.Text = lblList.Text & HttpUtility.HtmlEncode(l) & "<br />"
                    Loop Until l Is Nothing
                    sr.Close()
                End Using
            Catch ex As Exception
                lblList.Text = ex.Message
            End Try
            Exit Sub
        End If
    End Sub
   
    Protected Sub lbQueries_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        lblActualQuery.Text = String.Empty
        lblMessage.Text = String.Empty
        lblRowCount.Text = String.Empty
        lblQueryTitle.Text = lbQueries.SelectedItem.Text
        Call RetrieveQuery()
    End Sub
   
    Protected Sub RetrieveQuery()
        Dim oDataReader As SqlDataReader = Nothing
        Dim oConn As New SqlConnection(gsConn)
        Dim sSQL As String
        Dim oRegex As Regex
        Dim oMatch As Match
        sSQL = "SELECT Title, Tags, Preamble, Query, ISNULL(Timeout, 0) 'Timeout', ISNULL(Checksum, 0) 'Checksum', IsDeleted, CreatedOn, LastUpdatedOn, LastUsedOn, LastUsedBy FROM QueryDB WHERE [id] = " & lbQueries.SelectedValue
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            If oDataReader.HasRows Then
                oDataReader.Read()
                trInstructions.Visible = True
                lblInstructions.Text = oDataReader("Preamble")
                'pnCountDisplay = oDataReader("CountDisplay")
                'pnCountExport = oDataReader("CountExport")
                tbQuery.Text = oDataReader("Query")

                If cbShowQuerySource.Checked Then
                    trQuerySource.Visible = True
                    trActualQuery.Visible = True
                End If

                trCustomer.Visible = False
                If tbQuery.Text.ToLower.Contains("#customerkey") Then
                    trCustomer.Visible = True
                End If
               
                trParam1.Visible = False
                oRegex = New Regex("#param1.*#param1")
                If oRegex.IsMatch(tbQuery.Text) Then
                    oMatch = Regex.Match(tbQuery.Text, "#param1.*#param1")
                    Dim arrParam() = oMatch.Value.Split(",")
                    trParam1.Visible = True
                    lblParam1Legend.Text = arrParam(PARAM_LABEL)
                    lblParam1Notes.Text = arrParam(PARAM_NOTES)
                    tbQuery.Text = tbQuery.Text.Replace(oMatch.Value, "#param1")
                    Select Case arrParam(PARAM_TYPE)
                        Case "T"
                            tbParam1.Visible = True
                            tbParam1.Text = arrParam(PARAM_INIT_VALUE)
                            rfvParam1.ControlToValidate = "tbParam1"
                            rfvParam1.InitialValue = String.Empty
                        Case "N"
                            tbParam1.Visible = True
                            tbParam1.Text = arrParam(PARAM_INIT_VALUE)
                            rfvParam1.ControlToValidate = "tbParam1"
                            rfvParam1.InitialValue = String.Empty
                        Case "D"
                            tbParam1.Visible = False
                            ddlParam1.Visible = True
                            rfvParam1.ControlToValidate = "ddlParam1"
                            rfvParam1.InitialValue = "- please select -"
                            If arrParam(PARAM_INIT_VALUE) <> String.Empty Then
                                ddlParam1.Items.Clear()
                                'ddlParam1.Items.Add(New ListItem("- please select -", 0))
                                Dim arrListItems() As String = Split(arrParam(PARAM_INIT_VALUE), "|")
                                For Each sListItem As String In arrListItems
                                    Dim arrListItem() As String = Split(sListItem, "~")
                                    ddlParam1.Items.Add(New ListItem(arrListItem(0), arrListItem(1)))
                                Next
                            End If
                        Case Else
                    End Select
                End If

                trParam2.Visible = False
                oRegex = New Regex("#param2.*#param2")
                If oRegex.IsMatch(tbQuery.Text) Then
                    oMatch = Regex.Match(tbQuery.Text, "#param2.*#param2")
                    Dim arrParam() = oMatch.Value.Split(",")
                    trParam2.Visible = True
                    lblParam2Legend.Text = arrParam(PARAM_LABEL)
                    lblParam2Notes.Text = arrParam(PARAM_NOTES)
                    tbQuery.Text = tbQuery.Text.Replace(oMatch.Value, "#param2")
                    Select Case arrParam(PARAM_TYPE)
                        Case "T"
                            tbParam2.Visible = True
                            tbParam2.Text = arrParam(PARAM_INIT_VALUE)
                            ddlParam2.Visible = False
                            rfvParam2.ControlToValidate = "tbParam2"
                            rfvParam2.InitialValue = String.Empty
                        Case "N"
                            tbParam2.Visible = True
                            tbParam2.Text = arrParam(PARAM_INIT_VALUE)
                            ddlParam2.Visible = False
                            rfvParam2.ControlToValidate = "tbParam2"
                            rfvParam2.InitialValue = String.Empty
                        Case "D"
                            tbParam2.Visible = False
                            ddlParam2.Visible = True
                            rfvParam2.ControlToValidate = "ddlParam2"
                            rfvParam2.InitialValue = "- please select -"
                            If arrParam(PARAM_INIT_VALUE) <> String.Empty Then
                                ddlParam2.Items.Clear()
                                'ddlParam2.Items.Add(New ListItem("- please select -", 0))
                                Dim arrListItems() As String = Split(arrParam(PARAM_INIT_VALUE), "|")
                                For Each sListItem As String In arrListItems
                                    Dim arrListItem() As String = Split(sListItem, "~")
                                    ddlParam2.Items.Add(New ListItem(arrListItem(0), arrListItem(1)))
                                Next
                            End If
                        Case Else
                    End Select
                End If

                trDate.Visible = False
                If tbQuery.Text.ToLower.Contains("#date") Then
                    trDate.Visible = True
                End If

                trDateRange.Visible = False
                If tbQuery.Text.ToLower.Contains("#startdate") Then
                    trDateRange.Visible = True
                End If

                trFromDate.Visible = False
                If tbQuery.Text.ToLower.Contains("#fromdate") Then
                    trFromDate.Visible = True
                End If

                pnQueryLength = 0
                If Not IsDBNull(oDataReader("Checksum")) Then
                    pnQueryLength = oDataReader("Checksum")
                End If
                pnTimeout = 0
                If Not IsDBNull(oDataReader("Timeout")) Then
                    pnTimeout = oDataReader("Timeout")
                End If
                tbTags.Text = oDataReader("Tags")
                tbTitle.Text = oDataReader("Title")
                tbDescription.Text = oDataReader("Preamble")
                If Not IsDBNull(oDataReader("Timeout")) Then
                    tbTimeout.Text = oDataReader("Timeout")
                Else
                    tbTimeout.Text = 0
                End If
                
                If Not IsDBNull(oDataReader("Checksum")) Then
                    tbChecksum.Text = oDataReader("Checksum")
                Else
                    tbChecksum.Text = 0
                End If
                lnkbtnUpdateQuery.Enabled = True
                lnkbtnRemoveQuery.Enabled = True
            Else
                WebMsgBox.Show("Error - could not retrieve query")
            End If
        Catch ex As Exception
            WebMsgBox.Show("Error in RetrieveQuery: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
   
    Protected Sub ddlTags_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        If ddl.Items(0).Text.Contains("please select") Then
            ddl.Items.RemoveAt(0)
        End If
        Call PopulateQueryList(ddl.SelectedItem.Text)
    End Sub

    Protected Sub lnkbtnAddQuery_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TrimInputFields()
        If tbLocation.Text.Trim.Replace(" ", "") = SECRET Then
            If CountTitleInstances(tbTitle.Text) = 0 Then
                Dim sSQL As String = "INSERT INTO QueryDB (Title, Tags, Preamble, Query, Timeout, Checksum, IsDeleted, CreatedOn) VALUES ('" & tbTitle.Text.Replace("'", "''") & "', '" & tbTags.Text.Replace("'", "''") & "', '" & tbDescription.Text.Replace("'", "''") & "', '" & tbQuery.Text.Replace("'", "''") & "', " & tbTimeout.Text & ", " & tbChecksum.Text & ",0, GETDATE()) SELECT @@IDENTITY"
                Dim oDT As DataTable = ExecuteQueryToDataTable(sSQL)
                Dim nID As Int32 = oDT.Rows(0).Item(0)
                tbQuery.Text = String.Empty
                Call PopulateQueryList(String.Empty)
                For i As Int32 = 0 To lbQueries.Items.Count - 1
                    If lbQueries.Items(i).Value = nID Then
                        lbQueries.SelectedIndex = i
                        Exit For
                    End If
                Next
                Call RetrieveQuery()
                Call GetTags()
            Else
                WebMsgBox.Show("Query must be named uniquely. Please select an alternative name")
                tbTitle.Focus()
            End If
        Else
            WebMsgBox.Show("Enter valid location code")
            tbLocation.Focus()
        End If
    End Sub
   
    Protected Sub TrimInputFields()
        tbTags.Text = tbTags.Text.Trim
        tbTitle.Text = tbTitle.Text.Trim
        tbDescription.Text = tbDescription.Text.Trim
    End Sub
    
    Protected Function CountTitleInstances(sTitle As String) As Int32
        Dim sSQL As String = "SELECT * FROM QueryDB WHERE IsDeleted = 0 AND Title = '" & sTitle & "'"
        CountTitleInstances = ExecuteQueryToDataTable(sSQL).Rows.Count
    End Function
    
    Protected Sub lnkbtnUpdateQuery_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim nCurrentRecord = lbQueries.SelectedValue
        Call TrimInputFields()
        If tbLocation.Text.Trim.Replace(" ", "") = SECRET Then
            If CountTitleInstances(tbTitle.Text) < 2 Then
                Dim sSQL As String = "UPDATE QueryDb SET Title = '" & tbTitle.Text.Replace("'", "''") & "', Tags = '" & tbTags.Text.Replace("'", "''") & "', Preamble = '" & tbDescription.Text.Replace("'", "''") & "', Query = '" & tbQuery.Text.Replace("'", "''") & "', Timeout = " & tbTimeout.Text & ", Checksum = " & tbChecksum.Text & ", LastUpdatedOn = GETDATE() WHERE [id] = " & lbQueries.SelectedValue
                Call ExecuteQueryToDataTable(sSQL)
                tbQuery.Text = String.Empty
                tbTags.Text = String.Empty
                tbTitle.Text = String.Empty
                tbDescription.Text = String.Empty
                tbTimeout.Text = String.Empty
                tbChecksum.Text = String.Empty
                lnkbtnUpdateQuery.Enabled = False
                Call PopulateQueryList(String.Empty)
                For i As Int32 = 0 To lbQueries.Items.Count - 1
                    If lbQueries.Items(i).Value = nCurrentRecord Then
                        lbQueries.SelectedIndex = i
                        Exit For
                    End If
                Next
                Call RetrieveQuery()
                Call GetTags()
            Else
                WebMsgBox.Show("Query must be named uniquely. Please select an alternative name")
                tbTitle.Focus()
            End If
        Else
            WebMsgBox.Show("Enter valid location code")
            tbLocation.Focus()
        End If
    End Sub
    
    Protected Sub btnExport_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        psQuery = tbQuery.Text.Trim
        pnlData.Visible = True
        If psQuery <> String.Empty Then
            If bIsAllowedQuery() OrElse Not (psQuery.ToLower.Contains("update") Or psQuery.ToLower.Contains("insert") Or psQuery.ToLower.Contains("delete")) Then
                psQuery = SubstituteCustomer(psQuery)
                psQuery = SubstituteDate(psQuery)
                psQuery = SubstituteFromDate(psQuery)
                psQuery = SubstituteDateRange(psQuery)
                psQuery = SubstituteParam1(psQuery)
                psQuery = SubstituteParam2(psQuery)
                lblActualQuery.Text = psQuery
                If Not psQuery.Contains("1-jan-1900") Then
                    Call RecordUsage("export")
                    Call ExportResults()
                End If
            Else
                WebMsgBox.Show("SELECT queries only, thank you!" & " (" & psQuery.Length & ")")
            End If
        End If
    End Sub
   
    Protected Sub ExportResults()
        If psQuery <> String.Empty Then
            Dim oConn As New SqlConnection(gsConn)
            lblMessage.Text = String.Empty
            Try
                Dim sSQL As String = psQuery
                Dim oAdapter As New SqlDataAdapter(sSQL, oConn)
                If pnTimeout > 0 Then
                    oAdapter.SelectCommand.CommandTimeout = pnTimeout
                End If
                oAdapter.Fill(oDataTable)
                lblRowCount.Text = oDataTable.Rows.Count & " record(s)"
                Dim sCSVString As String = ConvertDataTableToCSVString(oDataTable)
                Call ExportCSVData(sCSVString)
            Catch ex As Exception
                lblMessage.Text = ex.Message
                'WebMsgBox.Show("Error in ExportResults: " & ex.ToString)
            Finally
                oConn.Close()
            End Try
        End If
    End Sub
   
    Public Function ConvertDataTableToCSVString(ByVal oDataTable As DataTable) As String
        Dim sbResult As New StringBuilder
        Dim oDataColumn As DataColumn
        Dim oDataRow As DataRow

        For Each oDataColumn In oDataTable.Columns         ' column headings in line 1
            sbResult.Append(oDataColumn.ColumnName)
            sbResult.Append(",")
        Next
        If sbResult.Length > 1 Then
            sbResult.Length = sbResult.Length - 1
        End If
        sbResult.Append(Environment.NewLine)
        Dim s2 As String
        For Each oDataRow In oDataTable.Rows
            For Each s As Object In oDataRow.ItemArray
                Try
                    s2 = s
                Catch
                    s2 = String.Empty
                End Try
                s2 = s2.Replace(Environment.NewLine, " ")
                sbResult.Append(s2.Replace(",", " "))
                sbResult.Append(",")
            Next
            sbResult.Length = sbResult.Length - 1
            sbResult.Append(Environment.NewLine)
        Next oDataRow

        If Not sbResult Is Nothing Then
            Return sbResult.ToString()
        Else
            Return String.Empty
        End If
    End Function
   
    Private Sub ExportCSVData(ByVal sCSVString As String)
        Response.Clear()
        Response.AddHeader("Content-Disposition", "attachment;filename=" & "QueryResult.csv")
        Response.ContentType = "text/csv"
        'Response.ContentType = "application/vnd.ms-excel"
   
        Dim eEncoding As Encoding = Encoding.GetEncoding("Windows-1252")
        Dim eUnicode As Encoding = Encoding.Unicode
        Dim byUnicodeBytes As Byte() = eUnicode.GetBytes(sCSVString)
        Dim byEncodedBytes As Byte() = Encoding.Convert(eUnicode, eEncoding, byUnicodeBytes)
        Response.BinaryWrite(byEncodedBytes)
        Response.End()
        ' Response.Flush()
    End Sub

    Protected Sub cbShowQuerySource_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim cb As CheckBox = sender
        If cb.Checked Then
            trQuerySource.Visible = True
            trActualQuery.Visible = True
            tbLocation.Focus()
        Else
            trQuerySource.Visible = False
            trActualQuery.Visible = False
            trFields.Visible = False
        End If
    End Sub
   
    Protected Sub cbIncludeAccountsWithNoProducts_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call InitCustomerDropdown()
    End Sub

    Protected Sub cbIncludeSuspendedDeletedAccounts_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call InitCustomerDropdown()
    End Sub
   
    Protected Sub lnkbtnShowFields_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If tbLocation.Text.Trim.ToLower.Replace(" ", "") = SECRET Then
            trFields.Visible = True
            lnkbtnAddQuery.Visible = True
            lnkbtnUpdateQuery.Visible = True
        Else
            WebMsgBox.Show("Enter valid location code")
            tbLocation.Focus()
        End If
    End Sub
    
    Protected Sub lnkbtnRefreshQueryList_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call PopulateQueryList(String.Empty)
        Call GetTags()
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

    'Protected Function ExecuteNonQuery(ByVal sQuery As String) As Boolean
    '    ExecuteNonQuery = False
    '    Dim oConn As New SqlConnection(gsConn)
    '    Dim oCmd As SqlCommand
    '    Try
    '        oConn.Open()
    '        oCmd = New SqlCommand(sQuery, oConn)
    '        oCmd.ExecuteNonQuery()
    '        ExecuteNonQuery = True
    '    Catch ex As Exception
    '        WebMsgBox.Show("Error in ExecuteNonQuery executing: " & sQuery & " : " & ex.Message)
    '    Finally
    '        oConn.Close()
    '    End Try
    'End Function

    Property pnQueryLength() As Integer
        Get
            Dim o As Object = ViewState("QD_QueryLength")
            If o Is Nothing Then
                Return 0
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Integer)
            ViewState("QD_QueryLength") = Value
        End Set
    End Property

    Property pnTimeout() As Integer
        Get
            Dim o As Object = ViewState("QD_Timeout")
            If o Is Nothing Then
                Return 0
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Integer)
            ViewState("QD_Timeout") = Value
        End Set
    End Property

    Property pnCountDisplay() As Integer
        Get
            Dim o As Object = ViewState("QD_CountDisplay")
            If o Is Nothing Then
                Return 0
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Integer)
            ViewState("QD_CountDisplay") = Value
        End Set
    End Property

    Property pnCountExport() As Integer
        Get
            Dim o As Object = ViewState("QD_CountExport")
            If o Is Nothing Then
                Return 0
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Integer)
            ViewState("QD_CountExport") = Value
        End Set
    End Property

    Property psQuery() As String
        Get
            Dim o As Object = ViewState("QD_Query")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("QD_Query") = Value
        End Set
    End Property

    Protected Sub lnkbtnClearFields_Click(sender As Object, e As System.EventArgs)
        Call ClearFields()
    End Sub
    
    Protected Sub ClearFields()
        tbTags.Text = String.Empty
        tbTitle.Text = String.Empty
        tbDescription.Text = String.Empty
        tbTimeout.Text = "0"
        tbChecksum.Text = "0"
        tbQuery.Text = String.Empty
    End Sub
    
    Protected Sub lnkbtnRemoveQuery_Click(sender As Object, e As System.EventArgs)
        Dim sSQL As String = "UPDATE QueryDB SET IsDeleted = 1 WHERE [id] = " & lbQueries.SelectedValue
        Call ExecuteQueryToDataTable(sSQL)
        Call PopulateQueryList(String.Empty)
        lnkbtnRemoveQuery.Enabled = False
        Call ClearFields()
    End Sub
    
    Protected Sub lnkbtnMoreOptions_Click(sender As Object, e As System.EventArgs)
        Dim lb As LinkButton = sender
        If lb.Text.Contains("more") Then
            lb.Text = "less..."
            cbIncludeAccountsWithNoProducts.Visible = True
            cbIncludeSuspendedDeletedAccounts.Visible = True
        Else
            lb.Text = "more..."
            cbIncludeAccountsWithNoProducts.Visible = False
            cbIncludeSuspendedDeletedAccounts.Visible = False
        End If
    End Sub
    
</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Query the database</title>
    <style type="text/css">
        .style1
        {
            width: 10%;
        }
        .style2
        {
            width: 90%;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
    <div style="font-size: xx-small; font-family: Verdana">
        <main:Header ID="ctlHeader" runat="server"></main:Header>
        <table width="100%" cellpadding="0" cellspacing="0">
            <tr class="bar_accounthandler">
                <td style="width: 50%; white-space: nowrap">
                </td>
                <td style="width: 50%; white-space: nowrap" align="right">
                </td>
            </tr>
        </table>
        <table width="95%">
            <tr>
                <td colspan="2">
                    <strong>
                    <asp:Label ID="lblTitle" runat="server" Font-Names="Verdana" Font-Size="Small">Database Query</asp:Label>
                    </strong>
                </td>
            </tr>
            <tr>
                <td align="right" style="width: 10%" valign="top">
                    Tags:
                </td>
                <td style="width: 90%">
                    <asp:DropDownList ID="ddlTags" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        AutoPostBack="True" OnSelectedIndexChanged="ddlTags_SelectedIndexChanged">
                    </asp:DropDownList>
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                    <asp:CheckBox ID="cbShowQuerySource" runat="server" AutoPostBack="True" OnCheckedChanged="cbShowQuerySource_CheckedChanged"
                        Text="edit query" />
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                    <asp:LinkButton ID="lnkbtnRefreshQueryList" runat="server" OnClick="lnkbtnRefreshQueryList_Click">refresh query list</asp:LinkButton>
                </td>
            </tr>
            <tr>
                <td align="right" style="width: 10%" valign="top">
                    Available queries:
                </td>
                <td style="width: 90%">
                    <asp:ListBox ID="lbQueries" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Rows="10" Width="100%" AutoPostBack="True" OnSelectedIndexChanged="lbQueries_SelectedIndexChanged">
                    </asp:ListBox>
                </td>
            </tr>
            <tr id="trInstructions" runat="server" visible="false">
                <td align="right" style="width: 10%" valign="top">
                    Description:
                </td>
                <td style="width: 90%">
                    <asp:Label ID="lblQueryTitle" runat="server" Font-Names="Verdana" 
                        Font-Size="Small" Font-Bold="True" ForeColor="#003399"></asp:Label>
                &nbsp;<asp:Label ID="lblInstructions" runat="server" Font-Names="Verdana" 
                        Font-Size="X-Small" Font-Bold="False" ForeColor="#003399"></asp:Label>
                </td>
            </tr>
            <tr id="trCustomer" runat="server" visible="false">
                <td align="right" style="width: 10%" valign="top">
                    Customer:
                </td>
                <td style="width: 90%">
                    <asp:DropDownList ID="ddlCustomer" runat="server" Font-Names="Verdana" Font-Size="Small" />
                    &nbsp;&nbsp;&nbsp;&nbsp;
                    <asp:LinkButton ID="lnkbtnMoreOptions" runat="server" Font-Size="XX-Small" 
                        onclick="lnkbtnMoreOptions_Click">more...</asp:LinkButton>
                    <asp:CheckBox ID="cbIncludeAccountsWithNoProducts" runat="server" Text="include accounts with no products"
                        AutoPostBack="True" 
                        OnCheckedChanged="cbIncludeAccountsWithNoProducts_CheckedChanged" 
                        Visible="False" />
                    <asp:CheckBox ID="cbIncludeSuspendedDeletedAccounts" runat="server" Text="include suspended &amp; deleted accounts"
                        AutoPostBack="True" 
                        OnCheckedChanged="cbIncludeSuspendedDeletedAccounts_CheckedChanged" 
                        Visible="False" />
                </td>
            </tr>
            <tr id="trParam1" runat="server" visible="false">
                <td align="right" style="width: 10%" valign="top">
                    <asp:Label ID="lblParam1Legend" runat="server"></asp:Label>
                </td>
                <td style="width: 90%">
                    <asp:TextBox ID="tbParam1" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Width="200px" Visible="False"></asp:TextBox>
                    <asp:DropDownList ID="ddlParam1" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Visible="False">
                    </asp:DropDownList>
                    <asp:Label ID="lblParam1Notes" runat="server"></asp:Label>
                    <asp:RequiredFieldValidator ID="rfvParam1" runat="server" ControlToValidate="tbParam1"
                        ErrorMessage="required!" Font-Bold="True"></asp:RequiredFieldValidator>
                </td>
            </tr>
            <tr id="trParam2" runat="server" visible="false">
                <td align="right" style="width: 10%" valign="top">
                    <asp:Label ID="lblParam2Legend" runat="server"></asp:Label>
                </td>
                <td style="width: 90%">
                    <asp:TextBox ID="tbParam2" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Width="200px" Visible="False"></asp:TextBox>
                    <asp:DropDownList ID="ddlParam2" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Visible="False">
                    </asp:DropDownList>
                    <asp:Label ID="lblParam2Notes" runat="server"></asp:Label>
                    <asp:RequiredFieldValidator ID="rfvParam2" runat="server" ControlToValidate="tbParam2"
                        ErrorMessage="required!" Font-Bold="True"></asp:RequiredFieldValidator>
                </td>
            </tr>
            <tr id="trDateRange" runat="server" visible="false">
                <td align="right" style="width: 10%" valign="middle">
                    Date range:
                </td>
                <td style="width: 90%" valign="middle">
                    From:
                    <asp:TextBox ID="tbStartDate" runat="server" Font-Names="Verdana" Font-Size="Small"
                        Width="90"></asp:TextBox>
                    <a id="imgCalendarButton1" runat="server" href="javascript:;" onclick="window.open('./PopupCalendar4.aspx?textbox=tbStartDate','cal','width=300,height=305,left=270,top=180')"
                        visible="true">
                        <img id="Img1" runat="server" alt="" border="0" ie:visible="true" src="./images/SmallCalendar.gif"
                            visible="false" /></a> <span id="spnDateExample1" runat="server" class="informational light"
                                visible="true">(eg 12-Jan-2011)</span> &nbsp;&nbsp; <span class="informational dark">
                                    To:</span>
                    <asp:TextBox ID="tbEndDate" runat="server" Font-Names="Verdana" Font-Size="Small"
                        Width="90" />
                    <a id="imgCalendarButton2" runat="server" href="javascript:;" onclick="window.open('./PopupCalendar4.aspx?textbox=tbEndDate','cal','width=300,height=305,left=270,top=180')"
                        visible="true">
                        <img id="Img2" runat="server" alt="" border="0" ie:visible="true" src="./images/SmallCalendar.gif"
                            visible="false" /></a> <span id="spnDateExample2" runat="server" class="informational light"
                                visible="true">(eg 12-Jan-2012)</span>
                </td>
            </tr>
            <tr id="trFromDate" runat="server" visible="false">
                <td align="right" style="width: 10%" valign="middle">
                    From Date:
                </td>
                <td style="width: 90%" valign="middle">
                    <asp:TextBox ID="tbFromDate" runat="server" Font-Names="Verdana" Font-Size="Small"
                        Width="90" />
                    <a id="A1" runat="server" href="javascript:;" onclick="window.open('./PopupCalendar4.aspx?textbox=tbFromDate','cal','width=300,height=305,left=270,top=180')"
                        visible="true">
                        <img id="Img33" runat="server" alt="" border="0" ie:visible="true" src="./images/SmallCalendar.gif"
                            visible="false" /></a> <span id="spnDateExample33" runat="server" class="informational light"
                                visible="true">(eg 05-Jan-2012)</span> to present
                </td>
            </tr>
            <tr id="trDate" runat="server" visible="false">
                <td align="right" style="width: 10%" valign="middle">
                    Date:
                </td>
                <td style="width: 90%" valign="middle">
                    <asp:TextBox ID="tbDate" runat="server" Font-Names="Verdana" Font-Size="Small"
                        Width="90" />
                    <a id="imgCalendarButton3" runat="server" href="javascript:;" onclick="window.open('./PopupCalendar4.aspx?textbox=tbDate','cal','width=300,height=305,left=270,top=180')"
                        visible="true">
                        <img id="Img3" runat="server" alt="" border="0" ie:visible="true" src="./images/SmallCalendar.gif"
                            visible="false" /></a> <span id="spnDateExample3" runat="server" class="informational light"
                                visible="true">(eg 12-Jan-2012)</span>
                </td>
            </tr>
            <tr id="trQuerySource" runat="server" visible="false">
                <td align="right" valign="top" class="style1">
                    Query source:
                </td>
                <td class="style2">
                    <asp:TextBox ID="tbQuery" runat="server" Width="95%" MaxLength="1000" Font-Names="Verdana"
                        Font-Size="XX-Small" Rows="3" TextMode="MultiLine" Height="87px"></asp:TextBox><br />
                    <asp:Label ID="Label3" runat="server">(admin use only)</asp:Label>
                    &nbsp;<asp:Label ID="Label1" runat="server">Location:</asp:Label>
                    <asp:TextBox ID="tbLocation" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Width="80px"></asp:TextBox>
                    <asp:LinkButton ID="lnkbtnShowFields" runat="server" CausesValidation="False" OnClick="lnkbtnShowFields_Click">show fields</asp:LinkButton>
                    &nbsp;<asp:LinkButton ID="lnkbtnClearFields" runat="server" 
                        onclick="lnkbtnClearFields_Click">clear fields</asp:LinkButton>
&nbsp;<asp:LinkButton ID="lnkbtnAddQuery" runat="server" OnClick="lnkbtnAddQuery_Click"
                        CausesValidation="False" Visible="False">add&nbsp;query</asp:LinkButton>
                    &nbsp;<asp:LinkButton ID="lnkbtnUpdateQuery" runat="server" Enabled="False" OnClick="lnkbtnUpdateQuery_Click"
                        Visible="False">update query</asp:LinkButton>
                    &nbsp;<asp:LinkButton ID="lnkbtnRemoveQuery" runat="server" 
                        OnClientClick='return confirm("Are you sure you want to remove this query?");' 
                        onclick="lnkbtnRemoveQuery_Click" Enabled="False">remove query</asp:LinkButton>
                </td>
            </tr>
            <tr id="trFields" runat="server" visible="false">
                <td align="right" style="width: 10%" valign="top">
                </td>
                <td style="width: 90%">
                    Tags:<asp:TextBox ID="tbTags" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Width="90">ALL </asp:TextBox>
                    &nbsp;Title:<asp:TextBox ID="tbTitle" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Width="169px"></asp:TextBox>
                    &nbsp;Description:<asp:TextBox ID="tbDescription" runat="server" Font-Names="Verdana"
                        Font-Size="XX-Small" Width="174px"></asp:TextBox>
                    &nbsp;Timeout:<asp:TextBox ID="tbTimeout" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Width="40px">0</asp:TextBox>&nbsp;Checksum:<asp:TextBox ID="tbChecksum" runat="server"
                            Font-Names="Verdana" Font-Size="XX-Small" Width="45px">0</asp:TextBox>
                </td>
            </tr>
            <tr id="trActualQuery" runat="server" visible="false">
                <td align="right" style="width: 10%; height: 14px" valign="top">
                    Actual query:
                </td>
                <td style="width: 90%; height: 14px">
                    <asp:Label ID="lblActualQuery" runat="server" Font-Names="Verdana" Font-Size="XX-Small"></asp:Label>
                </td>
            </tr>
            <tr id="trFileSys" runat="server" visible="false">
                <td align="right" valign="top">
                    FileSys:
                </td>
                <td>
                    <asp:TextBox ID="tbDo" runat="server" Font-Names="Verdana" Font-Size="XX-Small"></asp:TextBox>
                    <asp:Button ID="btnDo" runat="server" OnClick="btnDo_Click" Text="do" /><br />
                    <asp:Label ID="lblList" runat="server"></asp:Label><br />
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td>
                    <asp:Button ID="btnGo" runat="server" Text="display results" OnClick="btnGo_Click"
                        Width="200px" />
                    &nbsp;<asp:Label ID="Label5" runat="server">show</asp:Label>
                    &nbsp;<asp:TextBox ID="tbRows" runat="server" Width="31px" Font-Names="Verdana" Font-Size="XX-Small">25</asp:TextBox>
                    <asp:Label ID="Label4" runat="server">rows/page</asp:Label>
                    &nbsp;<asp:RegularExpressionValidator ID="revRows" runat="server" ControlToValidate="tbRows"
                        ErrorMessage="must be numeric" ValidationExpression="\d*" Font-Bold="True"></asp:RegularExpressionValidator>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                    <asp:Button ID="btnExport" runat="server" OnClick="btnExport_Click" Text="export results to excel"
                        Width="200px" />
                &nbsp;<asp:Label ID="lblRowCount" runat="server" Font-Size="Small"></asp:Label>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td>
                    <asp:Label ID="lblMessage" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small"
                        ForeColor="Red"></asp:Label>
                </td>
            </tr>
        </table>
    </div>
    <asp:Panel ID="pnlData" runat="server" Width="100%">
        <asp:GridView ID="gvDisplay" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
            CellPadding="3" EnableViewState="False" AllowPaging="True" OnPageIndexChanging="gvDisplay_PageIndexChanging"
            Width="100%">
            <EmptyDataTemplate>
                no records found
            </EmptyDataTemplate>
            <AlternatingRowStyle BackColor="WhiteSmoke" />
        </asp:GridView>
    </asp:Panel>
    </form>
</body>
</html>
