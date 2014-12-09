<%@ Page Language="VB" Theme="AIMSDefault" ValidateRequest="false" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.IO" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

    Const MAX_POSTCODE_LENGTH As Int32 = 10
    
    Const COUNTRY_CODE_CANADA As Int32 = 38
    Const COUNTRY_CODE_USA As Int32 = 223
    Const COUNTRY_CODE_USA_NYC As Int32 = 256

    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString

    Sub Page_Load(ByVal Source As Object, ByVal E As EventArgs)
        Server.ScriptTimeout = 3600
        Page.MaintainScrollPositionOnPostBack = True
    End Sub

    Protected Sub GetNextEntry()
        lblLegendPostcodeTooLong.Visible = False
        lblLegendPostcodeRequired.Visible = False
        lblLegendCountyStateRegionRequired.Visible = False

        Dim sSQL As String = String.Empty
        sSQL = "SELECT TOP 1 CustomerAccountCode, CountryName, gab.* FROM GlobalAddressBook gab INNER JOIN Customer c ON gab.CustomerKey = c.CustomerKey INNER JOIN Country ctry ON gab.CountryKey = ctry.CountryKey WHERE CustomerStatusId = 'ACTIVE' AND gab.DeletedFlag = 0 AND NOT (ISNULL(gab.DefaultCommodityId,'') LIKE '%DONE%' OR ISNULL(gab.DefaultCommodityId,'') LIKE '%PENDING%')"
        If rbCanadaOnly.Checked Then
            sSQL = sSQL & " AND gab.CountryKey = 38"
        ElseIf rbUSOnly.Checked Then
            sSQL = sSQL & " AND (gab.CountryKey = 223 OR gab.CountryKey = 256)"
        End If
        
        sSQL = sSQL & " ORDER BY CustomerAccountCode"
        Dim odt As DataTable = ExecuteQueryToDataTable(sSQL)
        'Dim odt As DataTable = ExecuteQueryToDataTable("SELECT TOP 1 CustomerAccountCode, CountryName, gab.* FROM GlobalAddressBook gab INNER JOIN Customer c ON gab.CustomerKey = c.CustomerKey INNER JOIN Country ctry ON gab.CountryKey = ctry.CountryKey WHERE CustomerStatusId = 'ACTIVE' AND gab.DeletedFlag = 0 AND ISNULL(gab.DefaultCommodityId,'') = '' AND gab.CountryKey = 38 ORDER BY CustomerAccountCode")
        
        If odt.Rows.Count = 1 Then
            Dim dr As DataRow = odt.Rows(0)
            lblCustomerAccountCode.Text = dr("CustomerAccountCode") & String.Empty
            tbCompany.Text = dr("Company") & String.Empty
            tbContactName.Text = dr("AttnOf") & String.Empty
            tbAddr1.Text = dr("Addr1") & String.Empty
            tbAddr2.Text = dr("Addr2") & String.Empty
            tbAddr3.Text = dr("Addr3") & String.Empty
            tbTownCity.Text = dr("Town") & String.Empty
            tbCountyState.Text = dr("State") & String.Empty
            tbPostcode.Text = dr("PostCode") & String.Empty
            tbCountry.Text = dr("CountryName") & String.Empty
            
            pnKey = dr("Key") & String.Empty
            
            Call SetCountry(dr("CountryKey"), tbCountyState.Text)

            lblCompany.Text = dr("Company") & String.Empty
            lblContactName.Text = dr("AttnOf") & String.Empty
            lblAddr1.Text = dr("Addr1") & String.Empty
            lblAddr2.Text = dr("Addr2") & String.Empty
            lblAddr3.Text = dr("Addr3") & String.Empty
            lblTownCity.Text = dr("Town") & String.Empty
            lblCountyState.Text = dr("State") & String.Empty
            lblPostcode.Text = dr("PostCode") & String.Empty
            tbCountry.Text = dr("CountryName") & String.Empty

            hidCountryKey.Value = dr("CountryKey")

            tbNote.Text = dr("DefaultSpecialInstructions") & String.Empty
        Else
            lblCustomerAccountCode.Text = String.Empty
            tbCompany.Text = String.Empty
            tbContactName.Text = String.Empty
            tbAddr1.Text = String.Empty
            tbAddr2.Text = String.Empty
            tbAddr3.Text = String.Empty
            tbTownCity.Text = String.Empty
            tbCountyState.Text = String.Empty
            tbPostcode.Text = String.Empty
            tbCountry.Text = String.Empty
            
            pnKey = 0
            
            lblCompany.Text = String.Empty
            lblContactName.Text = String.Empty
            lblAddr1.Text = String.Empty
            lblAddr2.Text = String.Empty
            lblAddr3.Text = String.Empty
            lblTownCity.Text = String.Empty
            lblCountyState.Text = String.Empty
            lblPostcode.Text = String.Empty
            tbCountry.Text = String.Empty
            hidCountryKey.Value = 0
            tbNote.Text = String.Empty
        End If
        'If hidCountryKey.Value = COUNTRY_CODE_CANADA Or hidCountryKey.Value = COUNTRY_CODE_USA Then
        '    rfvRegion.Enabled = True
        'End If
    End Sub

    Protected Sub SetCountry(nCountryKey As Int32, sStateOrProvince As String)
        If nCountryKey = COUNTRY_CODE_USA Then
            Call SetCountryUSA(sStateOrProvince)
        ElseIf nCountryKey = COUNTRY_CODE_USA_NYC Then
            Call SetCountryUSANewYorkCity()
        ElseIf nCountryKey = COUNTRY_CODE_CANADA Then
            Call SetCountryCanada(sStateOrProvince)
        Else
            Call SetCountryOther()
        End If
    End Sub
    
    Protected Sub SetCountryOther()
        Call HideCountryRelatedControls()
        tbCountyState.Visible = True
        lblLegendRegion.Text = "County / Region"
        lblLegendRegion.ForeColor = Drawing.Color.Blue
        tbCountyState.Text = String.Empty
    End Sub
    
    Protected Sub SetCountryUSA(sState As String)
        Call HideCountryRelatedControls()
        ddlUSStatesCanadianProvinces.Visible = True
        lblLegendRegion.Text = "State"
        lblLegendRegion.ForeColor = Drawing.Color.Red
        Call PopulateUSStatesDropdown()
        If sState <> String.Empty Then
            For i As Int32 = 0 To ddlUSStatesCanadianProvinces.Items.Count - 1
                If ddlUSStatesCanadianProvinces.Items(i).Text.ToLower = sState.ToLower Or ddlUSStatesCanadianProvinces.Items(i).Text.ToLower.Contains(sState.ToLower) Then
                    ddlUSStatesCanadianProvinces.SelectedIndex = i
                    Exit For
                End If
            Next
        End If
    End Sub
    
    Protected Sub SetCountryUSANewYorkCity()
        Call HideCountryRelatedControls()
        lblLegendNewYorkCity.Visible = True
        lblLegendRegion.Text = "State"
        lblLegendRegion.ForeColor = Drawing.Color.Red
    End Sub
    
    Protected Sub SetCountryCanada(sProvince As String)
        Call HideCountryRelatedControls()
        ddlUSStatesCanadianProvinces.Visible = True
        lblLegendRegion.Text = "Province"
        lblLegendRegion.ForeColor = Drawing.Color.Red
        Call PopulateCanadianProvincesDropdown()
        If sProvince <> String.Empty Then
            For i As Int32 = 0 To ddlUSStatesCanadianProvinces.Items.Count - 1
                If ddlUSStatesCanadianProvinces.Items(i).Text.ToLower = sProvince.ToLower Or ddlUSStatesCanadianProvinces.Items(i).Text.ToLower.Contains(sProvince.ToLower) Then
                    ddlUSStatesCanadianProvinces.SelectedIndex = i
                    Exit For
                
                End If
            Next
        End If
    End Sub
    
    Protected Sub HideCountryRelatedControls()
        ddlUSStatesCanadianProvinces.Visible = False
        lblLegendNewYorkCity.Visible = False
        tbCountyState.Visible = False
    End Sub
    
    Protected Sub PopulateUSStatesDropdown()
        Dim olic As ListItemCollection = ExecuteQueryToListItemCollection("SELECT StateName + ' (' + StateAbbreviation + ')' sn, StateAbbreviation sa FROM US_States ORDER BY StateName", "sn", "sa")
        ddlUSStatesCanadianProvinces.Items.Clear()
        ddlUSStatesCanadianProvinces.Items.Add(New ListItem("- please select -", 0))
        For Each li As ListItem In olic
            ddlUSStatesCanadianProvinces.Items.Add(New ListItem(li.Text, li.Value))
        Next
    End Sub
    
    Protected Sub PopulateCanadianProvincesDropdown()
        Dim olic As ListItemCollection = ExecuteQueryToListItemCollection("SELECT ProvinceName + ' (' + ProvinceAbbreviation + ')' pn, ProvinceAbbreviation pa FROM CanadianProvinces ORDER BY ProvinceName", "pn", "pa")
        ddlUSStatesCanadianProvinces.Items.Clear()
        ddlUSStatesCanadianProvinces.Items.Add(New ListItem("- please select -", 0))
        For Each li As ListItem In olic
            ddlUSStatesCanadianProvinces.Items.Add(New ListItem(li.Text, li.Value))
        Next
    End Sub

    Protected Sub SaveEntry()
        If pnKey > 0 Then
            Dim sbSQL As New StringBuilder
            sbSQL.Append("UPDATE GlobalAddressBook ")
            sbSQL.Append("SET ")
            sbSQL.Append("Company = ")
            sbSQL.Append("'")
            sbSQL.Append(tbCompany.Text.Replace("'", "''"))
            sbSQL.Append("'")
            sbSQL.Append(",")

            sbSQL.Append("AttnOf = ")
            sbSQL.Append("'")
            sbSQL.Append(tbContactName.Text.Replace("'", "''"))
            sbSQL.Append("'")
            sbSQL.Append(",")

            sbSQL.Append("Addr1 = ")
            sbSQL.Append("'")
            sbSQL.Append(tbAddr1.Text.Replace("'", "''"))
            sbSQL.Append("'")
            sbSQL.Append(",")

            sbSQL.Append("Addr2 = ")
            sbSQL.Append("'")
            sbSQL.Append(tbAddr2.Text.Replace("'", "''"))
            sbSQL.Append("'")
            sbSQL.Append(",")

            sbSQL.Append("Addr3 = ")
            sbSQL.Append("'")
            sbSQL.Append(tbAddr3.Text.Replace("'", "''"))
            sbSQL.Append("'")
            sbSQL.Append(",")

            sbSQL.Append("Town = ")
            sbSQL.Append("'")
            sbSQL.Append(tbTownCity.Text.Replace("'", "''"))
            sbSQL.Append("'")
            sbSQL.Append(",")

            sbSQL.Append("State = ")
            sbSQL.Append("'")
            If tbCountyState.Visible Then
                sbSQL.Append(tbCountyState.Text.Replace("'", "''"))
            ElseIf ddlUSStatesCanadianProvinces.Visible AndAlso ddlUSStatesCanadianProvinces.SelectedIndex > 0 Then
                sbSQL.Append(ddlUSStatesCanadianProvinces.SelectedItem.Text)
            ElseIf lblLegendNewYorkCity.Visible Then
                sbSQL.Append("NY")
            End If
            sbSQL.Append("'")
            sbSQL.Append(",")

            sbSQL.Append("Postcode = ")
            sbSQL.Append("'")
            sbSQL.Append(tbPostcode.Text.Replace("'", "''"))
            sbSQL.Append("'")
            sbSQL.Append(" ")

            sbSQL.Append("WHERE [Key] = ")
            sbSQL.Append(pnKey.ToString)

            Call ExecuteQueryToDataTable(sbSQL.ToString)
        End If
    End Sub

    Protected Sub MarkEntry(sText As String)
        Dim sSQL As String = "UPDATE GlobalAddressBook SET DefaultCommodityId = '" & sText.Replace("'", "''") & "' WHERE [Key] = " & pnKey
        Call ExecuteQueryToDataTable(sSQL)
    End Sub

    Protected Sub AddNote(sText As String)
        Dim sSQL As String = "UPDATE GlobalAddressBook SET DefaultSpecialInstructions = '" & sText.Replace("'", "''") & "' WHERE [Key] = " & pnKey
        Call ExecuteQueryToDataTable(sSQL)
    End Sub

    Protected Sub btnGetNext_Click(sender As Object, e As System.EventArgs)
        Call GetNextEntry()
    End Sub

    Protected Function PostcodeLengthIsValid() As Boolean
        tbPostcode.Text = tbPostcode.Text.Trim
        If tbPostcode.Text.Length > MAX_POSTCODE_LENGTH Then
            PostcodeLengthIsValid = False
            lblLegendPostcodeTooLong.Visible = True
        Else
            PostcodeLengthIsValid = True
            lblLegendPostcodeTooLong.Visible = False
        End If
    End Function
    
    Protected Function PostcodeIsPresent() As Boolean
        tbPostcode.Text = tbPostcode.Text.Trim
        If tbPostcode.Text.Length = 0 Then
            PostcodeIsPresent = False
            lblLegendPostcodeRequired.Visible = True
        Else
            PostcodeIsPresent = True
            lblLegendPostcodeRequired.Visible = False
        End If
    End Function
        
    Protected Sub btnSaveAndMarkDone_Click(sender As Object, e As System.EventArgs)
        If pnKey = 0 Then
            Exit Sub
        End If
        
        Page.Validate()
        If Not (Page.IsValid And PostcodeLengthIsValid() And PostcodeIsPresent()) Then
            Exit Sub
        End If
        If hidCountryKey.Value = COUNTRY_CODE_CANADA Or hidCountryKey.Value = COUNTRY_CODE_USA Then
            If ddlUSStatesCanadianProvinces.SelectedIndex = 0 Then
                lblLegendCountyStateRegionRequired.Visible = True
                Exit Sub
            Else
                lblLegendCountyStateRegionRequired.Visible = False
            End If
        End If


        Call SaveEntry()
        Call MarkEntry(Format(Date.Now, "yyyyMMddhhmmss") & " DONE")
        If tbNote.Text <> String.Empty Then
            Call AddNote(tbNote.Text)
        End If
        Call GetNextEntry()
    End Sub

    Protected Sub btnMarkDone_Click(sender As Object, e As System.EventArgs)
        Page.Validate()
        If Not Page.IsValid Then
            Exit Sub
        End If

        Call MarkEntry(Format(Date.Now, "yyyyMMddhhmmss") & " DONE")
        If tbNote.Text <> String.Empty Then
            Call AddNote(tbNote.Text)
        End If
        Call GetNextEntry()
    End Sub

    Protected Sub btnMarkPending_Click(sender As Object, e As System.EventArgs)
        Call MarkEntry(Format(Date.Now, "yyyyMMddhhmmss") & " PENDING")
        If tbNote.Text <> String.Empty Then
            Call AddNote(tbNote.Text)
        End If
        Call GetNextEntry()
    End Sub

    Protected Sub lnkbtnCopyCustomer_Click(sender As Object, e As System.EventArgs)
        tbClipboard.Text = tbCompany.Text
    End Sub

    Protected Sub lnkbtnPasteCustomer_Click(sender As Object, e As System.EventArgs)
        tbCompany.Text = tbClipboard.Text
    End Sub

    Protected Sub lnkbtnCopyContactName_Click(sender As Object, e As System.EventArgs)
        tbClipboard.Text = tbContactName.Text
    End Sub

    Protected Sub lnkbtnPasteContactName_Click(sender As Object, e As System.EventArgs)
        tbContactName.Text = tbClipboard.Text
    End Sub

    Protected Sub lnkbtnCopyAddr1_Click(sender As Object, e As System.EventArgs)
        tbClipboard.Text = tbAddr1.Text
    End Sub

    Protected Sub lnkbtnPasteAddr1_Click(sender As Object, e As System.EventArgs)
        tbAddr1.Text = tbClipboard.Text
    End Sub

    Protected Sub lnkbtnCopyAddr2_Click(sender As Object, e As System.EventArgs)
        tbClipboard.Text = tbAddr2.Text
    End Sub

    Protected Sub lnkbtnPasteAddr2_Click(sender As Object, e As System.EventArgs)
        tbAddr2.Text = tbClipboard.Text
    End Sub

    Protected Sub lnkbtnCopyAddr3_Click(sender As Object, e As System.EventArgs)
        tbClipboard.Text = tbAddr3.Text
    End Sub

    Protected Sub lnkbtnPasteAddr3_Click(sender As Object, e As System.EventArgs)
        tbAddr3.Text = tbClipboard.Text
    End Sub

    Protected Sub lnkbtnCopyTownCity_Click(sender As Object, e As System.EventArgs)
        tbClipboard.Text = tbTownCity.Text
    End Sub

    Protected Sub lnkbtnPasteTownCity_Click(sender As Object, e As System.EventArgs)
        tbTownCity.Text = tbClipboard.Text
    End Sub

    Protected Sub lnkbtnCopyCountyState_Click(sender As Object, e As System.EventArgs)
        tbClipboard.Text = tbCountyState.Text
    End Sub

    Protected Sub lnkbtnPasteCountyState_Click(sender As Object, e As System.EventArgs)
        tbCountyState.Text = tbClipboard.Text
    End Sub

    Protected Sub lnkbtnCopyPostcode_Click(sender As Object, e As System.EventArgs)
        tbClipboard.Text = tbPostcode.Text
    End Sub

    Protected Sub lnkbtnPastPostcode_Click(sender As Object, e As System.EventArgs)
        tbPostcode.Text = tbClipboard.Text
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

    Property pnKey() As Int32
        Get
            Dim o As Object = ViewState("CAB_Key")
            If o Is Nothing Then
                Return -1
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Int32)
            ViewState("CAB_Key") = Value
        End Set
    End Property

    Property psQuery() As String
        Get
            Dim o As Object = ViewState("CAB_Query")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("CAB_Query") = Value
        End Set
    End Property

    Protected Sub lnkbtnClearCustomer_Click(sender As Object, e As System.EventArgs)
        tbCompany.Text = String.Empty
    End Sub

    Protected Sub lnkbtnClearContactName_Click(sender As Object, e As System.EventArgs)
        tbContactName.Text = String.Empty
    End Sub

    Protected Sub lnkbtnClearAddr1_Click(sender As Object, e As System.EventArgs)
        tbAddr1.Text = String.Empty
    End Sub

    Protected Sub lnkbtnClearAddr2_Click(sender As Object, e As System.EventArgs)
        tbAddr2.Text = String.Empty
    End Sub

    Protected Sub lnkbtnClearAddr3_Click(sender As Object, e As System.EventArgs)
        tbAddr3.Text = String.Empty
    End Sub

    Protected Sub lnkbtnClearTownCity_Click(sender As Object, e As System.EventArgs)
        tbTownCity.Text = String.Empty
    End Sub

    Protected Sub lnkbtnClearCountyState_Click(sender As Object, e As System.EventArgs)
        tbCountyState.Text = String.Empty
    End Sub

    Protected Sub lnkbtnClearPostcode_Click(sender As Object, e As System.EventArgs)
        tbPostcode.Text = String.Empty
    End Sub
    
    Protected Sub lnkbtnClearAllPendingEntries_Click(sender As Object, e As System.EventArgs)
        Dim sSQL As String = "UPDATE GlobalAddressBook SET DefaultCommodityId = '' WHERE DefaultCommodityId LIKE '%pending%'"
        Call ExecuteQueryToDataTable(sSQL)
    End Sub
    
    Protected Sub lnkbtnGetStats_Click(sender As Object, e As System.EventArgs)
        lblStats.Text = "ENTRIES UNCHECKED: <b>" & ExecuteQueryToDataTable("SELECT COUNT (*) FROM GlobalAddressBook gab INNER JOIN Customer c ON gab.CustomerKey = c.CustomerKey WHERE CustomerStatusId = 'ACTIVE' AND gab.DeletedFlag = 0 AND ISNULL(DefaultCommodityId,'') = ''").Rows(0).Item(0)
        lblStats.Text = lblStats.Text & "</b> ENTRIES CHECKED: <b>" & ExecuteQueryToDataTable("SELECT COUNT (*) FROM GlobalAddressBook gab INNER JOIN Customer c ON gab.CustomerKey = c.CustomerKey WHERE CustomerStatusId = 'ACTIVE' AND gab.DeletedFlag = 0 AND ISNULL(gab.DefaultCommodityId,'') LIKE '%DONE%'").Rows(0).Item(0)
        lblStats.Text = lblStats.Text & "</b> ENTRIES PENDING: <b>" & ExecuteQueryToDataTable("SELECT COUNT (*) FROM GlobalAddressBook gab INNER JOIN Customer c ON gab.CustomerKey = c.CustomerKey WHERE CustomerStatusId = 'ACTIVE' AND gab.DeletedFlag = 0 AND ISNULL(gab.DefaultCommodityId,'') LIKE '%PENDING%'").Rows(0).Item(0)
        lblStats.Text = lblStats.Text & "</b>"
    End Sub
    
    Protected Sub btnClassifyEntries_Click(sender As Object, e As System.EventArgs)
        pnlGroup.Visible = True
        pnlIndividual.Visible = False
    End Sub

    Protected Sub btnEditEntries_Click(sender As Object, e As System.EventArgs)
        pnlGroup.Visible = False
        pnlIndividual.Visible = True
    End Sub
    
    Protected Sub btnAny_Click(sender As Object, e As System.EventArgs)
        psQuery = "SELECT TOP 10 gab.[key] 'recno', gab.Company, gab.Addr1  'Addr 1', ISNULL(gab.Addr2,'') 'Addr 2', ISNULL(gab.Addr3,'') 'Addr 3', gab.Town, gab.State 'Region', gab.PostCode FROM GlobalAddressBook gab INNER JOIN Customer c ON gab.CustomerKey = c.CustomerKey INNER JOIN Country ctry ON gab.CountryKey = ctry.CountryKey WHERE CustomerStatusId = 'ACTIVE' AND gab.DeletedFlag = 0 AND LEN(ISNULL(State,'')) > 1 AND LEN(ISNULL(PostCode,'')) > 2  AND NOT (ISNULL(gab.DefaultCommodityId,'') LIKE '%DONE%' OR ISNULL(gab.DefaultCommodityId,'') LIKE '%PENDING%')"
        Call BindGrid()
    End Sub

    Protected Sub btnCanada_Click(sender As Object, e As System.EventArgs)
        psQuery = "SELECT TOP 10 gab.[key] 'recno', gab.Company, gab.Addr1  'Addr 1', ISNULL(gab.Addr2,'') 'Addr 2', ISNULL(gab.Addr3,'') 'Addr 3', gab.Town, gab.State 'Province', gab.PostCode FROM GlobalAddressBook gab INNER JOIN Customer c ON gab.CustomerKey = c.CustomerKey INNER JOIN Country ctry ON gab.CountryKey = ctry.CountryKey WHERE CustomerStatusId = 'ACTIVE' AND gab.DeletedFlag = 0 AND gab.CountryKey IN (38) AND LEN(ISNULL(State,'')) > 1 AND LEN(ISNULL(PostCode,'')) > 2  AND NOT (ISNULL(gab.DefaultCommodityId,'') LIKE '%DONE%' OR ISNULL(gab.DefaultCommodityId,'') LIKE '%PENDING%')"
        Call BindGrid()
    End Sub

    Protected Sub btnUS_Click(sender As Object, e As System.EventArgs)
        psQuery = "SELECT TOP 10 gab.[key] 'recno', gab.Company, gab.Addr1  'Addr 1', ISNULL(gab.Addr2,'') 'Addr 2', ISNULL(gab.Addr3,'') 'Addr 3', gab.Town, gab.State, gab.PostCode 'Zip code' FROM GlobalAddressBook gab INNER JOIN Customer c ON gab.CustomerKey = c.CustomerKey INNER JOIN Country ctry ON gab.CountryKey = ctry.CountryKey WHERE CustomerStatusId = 'ACTIVE' AND gab.DeletedFlag = 0 AND gab.CountryKey IN (223, 256) AND LEN(ISNULL(State,'')) > 1 AND LEN(ISNULL(PostCode,'')) > 2  AND NOT (ISNULL(gab.DefaultCommodityId,'') LIKE '%DONE%' OR ISNULL(gab.DefaultCommodityId,'') LIKE '%PENDING%')"
        Call BindGrid()
    End Sub

    Protected Sub BindGrid()
        gvEntries.DataSource = ExecuteQueryToDataTable(psQuery)
        gvEntries.DataBind()
    End Sub
    
    Protected Sub btnGridDone_Click(sender As Object, e As System.EventArgs)
        Dim b As Button = sender
        pnKey = b.CommandArgument
        Call MarkEntry(Format(Date.Now, "yyyyMMddhhmmss") & " DONE")
        Call BindGrid()
    End Sub

    Protected Sub btnGridPending_Click(sender As Object, e As System.EventArgs)
        Dim b As Button = sender
        pnKey = b.CommandArgument
        Call MarkEntry(Format(Date.Now, "yyyyMMddhhmmss") & " PENDING")
        Call BindGrid()
    End Sub
    
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title></title>
        <link href="sprint.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="frmUserIdApplication" runat="server">
    &nbsp;<asp:Label ID="lblLegendCleanAddressBooks" runat="server" Font-Bold="True"
        Font-Names="Verdana" Font-Size="X-Small" Text="Clean Address Books" />
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    <asp:Button ID="btnClassifyEntries" runat="server" OnClick="btnClassifyEntries_Click"
        Text="classify entries" />
    &nbsp;
    <asp:Button ID="btnEditEntries" runat="server" Text="edit entries" OnClick="btnEditEntries_Click" />
    &nbsp;<asp:Label ID="Label15" runat="server" 
                        
        Text="Use 'classify entries' first to mark good entries. Then clean the remaining entries using 'edit entries'" />
                <br /><br />
    <asp:Panel ID="pnlGroup" runat="server" Visible="False" Width="100%">
        &nbsp;&nbsp;<asp:Button ID="btnAny" runat="server" Text="ANY" Width="100px" OnClick="btnAny_Click" />
        &nbsp;<asp:Button ID="btnCanada" runat="server" Text="CANADA" Width="100px" onclick="btnCanada_Click" />
        &nbsp;<asp:Button ID="btnUS" runat="server" Text="US" Width="100px" onclick="btnUS_Click" />
        &nbsp;<br />
        &nbsp;
        <asp:GridView ID="gvEntries" runat="server" CellPadding="2" Font-Names="Arial" Font-Size="Small" Width="98%" EnableModelValidation="True">
            <AlternatingRowStyle BackColor="#99FFCC" />
            <Columns>
                <asp:TemplateField>
                    <ItemTemplate>
                        <asp:Button ID="btnGridDone" runat="server" OnClick="btnGridDone_Click" CommandArgument='<%# Container.DataItem("recno")%>' Text="done" Height="60px" Width="60px" />
                    </ItemTemplate>
                    <ItemStyle Height="60px" />
                </asp:TemplateField>
                <asp:TemplateField>
                    <ItemTemplate>
                         <asp:Button ID="btnGridPending" runat="server" Text="pending" OnClick="btnGridPending_Click" CommandArgument='<%# Container.DataItem("recno")%>' Height="60px" Width="60px" />
                    </ItemTemplate>
                </asp:TemplateField>
            </Columns>
            <RowStyle BackColor="#FFFFCC" />
        </asp:GridView>
        <br />
        INSTRUCTIONS<br />
        <br />
        1. Do US and CANADA first, then ANY.<br />
        <br />
        2. For each type, click the button to fill the grid with entries. Check whether 
        the addresses look okay or not. Canada and US addresses must have a valid postal 
        / zip code and must show the Province / State.<br />
        <br />
        3.&nbsp; Click PENDING on each entry that looks bad. That will remove it from the 
        list for later examination.<br />
        <br />
        4.&nbsp; Click DONE on each entry that looks good. That will remove itfrom the list 
        and mark it as &#39;complete&#39;.<br />
    </asp:Panel>
    <asp:Panel ID="pnlIndividual" runat="server" Visible="True" Width="100%">
        <table style="width: 98%;">
            <tr>
                <td style="width: 15%;">
                    &nbsp;
                </td>
                <td align="right" colspan="2" style="width: 35%;">
                    <asp:Label ID="Label13" runat="server" Text="Customer:" />
                </td>
                <td style="width: 25%;">
                    <asp:Label ID="lblCustomerAccountCode" runat="server" Font-Bold="True" />
                </td>
                <td style="width: 25%;">
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td align="right">
                    <asp:Button ID="btnGetNext" runat="server" OnClick="btnGetNext_Click" Text="get next entry" />
                </td>
                <td>
                    <asp:RadioButton ID="rbUnmarked" runat="server" Checked="True" GroupName="NextEntryType"
                        Text="any unmarked" />
                    <asp:RadioButton ID="rbCanadaOnly" runat="server" GroupName="NextEntryType" Text="CANADA only" />
                    <asp:RadioButton ID="rbUSOnly" runat="server" GroupName="NextEntryType" Text="US only" />
                </td>
                <td align="right">
                    <asp:Label ID="Label4" runat="server" Text="Clipboard:" />
                </td>
                <td colspan="2">
                    <asp:TextBox ID="tbClipboard" runat="server" Width="98%" />
                </td>
            </tr>
            <tr>
                <td>
                    &nbsp;
                </td>
                <td colspan="2">
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td align="right">
                    <asp:Label ID="Label3" runat="server" Font-Bold="True" ForeColor="Red" Text="Company:" />
                </td>
                <td colspan="2">
                    <asp:TextBox ID="tbCompany" runat="server" MaxLength="50" Width="98%" />
                </td>
                <td>
                    <asp:LinkButton ID="lnkbtnCopyCustomer" runat="server" OnClick="lnkbtnCopyCustomer_Click">copy</asp:LinkButton>
                    &nbsp;
                    <asp:LinkButton ID="lnkbtnPasteCustomer" runat="server" OnClick="lnkbtnPasteCustomer_Click">paste</asp:LinkButton>
                    &nbsp;
                    <asp:LinkButton ID="lnkbtnClearCustomer" runat="server" OnClick="lnkbtnClearCustomer_Click">clear</asp:LinkButton>
                </td>
                <td>
                    <asp:Label ID="lblCompany" runat="server" />
                </td>
            </tr>
            <tr>
                <td align="right">
                    <asp:Label ID="Label2" runat="server" Text="Contact Name:" />
                </td>
                <td colspan="2">
                    <asp:TextBox ID="tbContactName" runat="server" MaxLength="50" Width="98%" />
                </td>
                <td>
                    <asp:LinkButton ID="lnkbtnCopyContactName" runat="server" OnClick="lnkbtnCopyContactName_Click">copy</asp:LinkButton>
                    &nbsp;
                    <asp:LinkButton ID="lnkbtnPasteContactName" runat="server" OnClick="lnkbtnPasteContactName_Click">paste</asp:LinkButton>
                    &nbsp;
                    <asp:LinkButton ID="lnkbtnClearContactName" runat="server" OnClick="lnkbtnClearContactName_Click">clear</asp:LinkButton>
                </td>
                <td>
                    <asp:Label ID="lblContactName" runat="server" />
                </td>
            </tr>
            <tr>
                <td align="right">
                    <asp:Label ID="Label5" runat="server" Font-Bold="True" ForeColor="Red" Text="Addr1:" />
                </td>
                <td colspan="2">
                    <asp:TextBox ID="tbAddr1" runat="server" MaxLength="50" Width="98%" />
                </td>
                <td>
                    <asp:LinkButton ID="lnkbtnCopyAddr1" runat="server" OnClick="lnkbtnCopyAddr1_Click">copy</asp:LinkButton>
                    &nbsp;
                    <asp:LinkButton ID="lnkbtnPasteAddr1" runat="server" OnClick="lnkbtnPasteAddr1_Click">paste</asp:LinkButton>
                    &nbsp;
                    <asp:LinkButton ID="lnkbtnClearAddr1" runat="server" OnClick="lnkbtnClearAddr1_Click">clear</asp:LinkButton>
                </td>
                <td>
                    <asp:Label ID="lblAddr1" runat="server" />
                </td>
            </tr>
            <tr>
                <td align="right">
                    <asp:Label ID="Label6" runat="server" Text="Addr2:" />
                </td>
                <td colspan="2">
                    <asp:TextBox ID="tbAddr2" runat="server" MaxLength="50" Width="98%" />
                </td>
                <td>
                    <asp:LinkButton ID="lnkbtnCopyAddr2" runat="server" OnClick="lnkbtnCopyAddr2_Click">copy</asp:LinkButton>
                    &nbsp;
                    <asp:LinkButton ID="lnkbtnPasteAddr2" runat="server" OnClick="lnkbtnPasteAddr2_Click">paste</asp:LinkButton>
                    &nbsp;
                    <asp:LinkButton ID="lnkbtnClearAddr2" runat="server" OnClick="lnkbtnClearAddr2_Click">clear</asp:LinkButton>
                </td>
                <td>
                    <asp:Label ID="lblAddr2" runat="server" />
                </td>
            </tr>
            <tr>
                <td align="right">
                    <asp:Label ID="Label7" runat="server" Text="Addr3:" />
                </td>
                <td colspan="2">
                    <asp:TextBox ID="tbAddr3" runat="server" MaxLength="50" Width="98%" />
                </td>
                <td>
                    <asp:LinkButton ID="lnkbtnCopyAddr3" runat="server" OnClick="lnkbtnCopyAddr3_Click">copy</asp:LinkButton>
                    &nbsp;
                    <asp:LinkButton ID="lnkbtnPasteAddr3" runat="server" OnClick="lnkbtnPasteAddr3_Click">paste</asp:LinkButton>
                    &nbsp;
                    <asp:LinkButton ID="lnkbtnClearAddr3" runat="server" OnClick="lnkbtnClearAddr3_Click">clear</asp:LinkButton>
                </td>
                <td>
                    <asp:Label ID="lblAddr3" runat="server" />
                </td>
            </tr>
            <tr>
                <td align="right">
                    <asp:Label ID="Label8" runat="server" Font-Bold="True" ForeColor="Red" Text="Town/City:" />
                </td>
                <td colspan="2">
                    <asp:TextBox ID="tbTownCity" runat="server" MaxLength="50" Width="98%" />
                </td>
                <td>
                    <asp:LinkButton ID="lnkbtnCopyTownCity" runat="server" OnClick="lnkbtnCopyTownCity_Click">copy</asp:LinkButton>
                    &nbsp;
                    <asp:LinkButton ID="lnkbtnPasteTownCity" runat="server" OnClick="lnkbtnPasteTownCity_Click">paste</asp:LinkButton>
                    &nbsp;
                    <asp:LinkButton ID="lnkbtnClearTownCity" runat="server" OnClick="lnkbtnClearTownCity_Click">clear</asp:LinkButton>
                </td>
                <td>
                    <asp:Label ID="lblTownCity" runat="server" />
                </td>
            </tr>
            <tr>
                <td align="right">
                    <asp:Label ID="lblLegendCountyStateRegionRequired" runat="server" Font-Bold="True"
                        Font-Names="Verdana" Font-Size="Small" ForeColor="#CC3300" Text="REQD &gt;" Visible="False"></asp:Label>
                    &nbsp;<asp:Label ID="lblLegendRegion" runat="server" Text="County/State/Region:" />
                </td>
                <td colspan="2">
                    <asp:DropDownList ID="ddlUSStatesCanadianProvinces" runat="server" Visible="False"
                        Width="98%" />
                    <asp:TextBox ID="tbCountyState" runat="server" MaxLength="50" Width="98%" />
                    <asp:Label ID="lblLegendNewYorkCity" runat="server" Text="NEW YORK (NY)" Visible="False"></asp:Label>
                </td>
                <td>
                    <asp:LinkButton ID="lnkbtnCopyCountyState" runat="server" OnClick="lnkbtnCopyCountyState_Click">copy</asp:LinkButton>
                    &nbsp;
                    <asp:LinkButton ID="lnkbtnPasteCountyState" runat="server" OnClick="lnkbtnPasteCountyState_Click">paste</asp:LinkButton>
                    &nbsp;
                    <asp:LinkButton ID="lnkbtnClearCountyState" runat="server" OnClick="lnkbtnClearCountyState_Click">clear</asp:LinkButton>
                </td>
                <td>
                    <asp:Label ID="lblCountyState" runat="server" />
                </td>
            </tr>
            <tr>
                <td align="right">
                    <asp:Label ID="lblLegendPostcodeTooLong" runat="server" Font-Bold="True" Font-Names="Verdana"
                        Font-Size="Small" ForeColor="#CC3300" Text="TOO LONG &gt;" Visible="False"></asp:Label>
                    &nbsp;<asp:Label ID="lblLegendPostcodeRequired" runat="server" Font-Bold="True" Font-Names="Verdana"
                        Font-Size="Small" ForeColor="#CC3300" Text="REQD &gt;" Visible="False"></asp:Label>
                    &nbsp;<asp:Label ID="lblLegendPostcode" runat="server" Font-Bold="True" ForeColor="Red"
                        Text="Post code:" />
                </td>
                <td colspan="2">
                    <asp:TextBox ID="tbPostcode" runat="server" MaxLength="10" Width="98%" />
                </td>
                <td>
                    <asp:LinkButton ID="lnkbtnCopyPostcode" runat="server" OnClick="lnkbtnCopyPostcode_Click">copy</asp:LinkButton>
                    &nbsp;
                    <asp:LinkButton ID="lnkbtnPastePostcode" runat="server" OnClick="lnkbtnPastPostcode_Click">paste</asp:LinkButton>
                    &nbsp;
                    <asp:LinkButton ID="lnkbtnClearPostcode" runat="server" OnClick="lnkbtnClearPostcode_Click">clear</asp:LinkButton>
                </td>
                <td>
                    <asp:Label ID="lblPostcode" runat="server" />
                </td>
            </tr>
            <tr>
                <td align="right">
                    <asp:Label ID="Label11" runat="server" Text="Country:" />
                </td>
                <td colspan="2">
                    <asp:TextBox ID="tbCountry" runat="server" Width="98%" />
                </td>
                <td>
                    <asp:HiddenField ID="hidCountryKey" runat="server" />
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td align="right">
                    &nbsp;
                </td>
                <td colspan="2">
                    &nbsp;&nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td align="right">
                    <asp:Label ID="lblLegendNote" runat="server" Text="Add Note:" />
                </td>
                <td colspan="2">
                    <asp:TextBox ID="tbNote" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        MaxLength="50" Width="98%" />
                </td>
                <td>
                    Add reminder text to the entry for later reference. Optional. Any text you enter
                    only ever appears here.
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td align="right">
                    &nbsp;
                </td>
                <td colspan="2">
                    <asp:Button ID="btnSaveAndMarkDone" runat="server" Font-Bold="True" OnClick="btnSaveAndMarkDone_Click"
                        Text="validate, save, mark 'done', get next entry" Width="300px" />
                </td>
                <td>
                    Validates the entry, marks as &quot;DONE&quot;, saves the changes, gets next entry
                    that is not marked as &quot;DONE&quot; or &quot;PENDING&quot;.
                </td>
                <td>
                    Use this for normal corrections.
                </td>
            </tr>
            <tr>
                <td align="right">
                    &nbsp;
                </td>
                <td colspan="2">
                    <asp:Button ID="btnMark" runat="server" OnClick="btnMarkDone_Click" Text="mark 'done', get next entry"
                        Width="300px" />
                </td>
                <td>
                    Marks the entry as &quot;DONE&quot; without validating it or saving any changes,
                    gets next entry that is not marked as &quot;DONE&quot; or &quot;PENDING&quot;.
                </td>
                <td>
                    Use when you know the entry is correct, or cannot be corrected, but it may fail
                    validation.
                </td>
            </tr>
            <tr>
                <td align="right">
                    &nbsp;
                </td>
                <td colspan="2">
                    <asp:Button ID="btnMarkPending" runat="server" OnClick="btnMarkPending_Click" Text="mark 'pending',  get next entry"
                        Width="300px" />
                </td>
                <td>
                    Marks the entry as &quot;PENDING&quot; so it doesn&#39;t get retrieved by &#39;get
                    next entry&#39;.
                </td>
                <td>
                    Use when you want to come back to this entry later.
                </td>
            </tr>
            <tr>
                <td align="right">
                    &nbsp;
                </td>
                <td colspan="2">
                    <asp:LinkButton ID="lnkbtnClearAllPendingEntries" runat="server" OnClick="lnkbtnClearAllPendingEntries_Click"
                        OnClientClick="return confirm(&quot;Are you sure you want to clear PENDING from all entries?&quot;);">clear &#39;PENDING&#39; from all entries</asp:LinkButton>
                    &nbsp;
                    <asp:LinkButton ID="lnkbtnGetStats" runat="server" OnClick="lnkbtnGetStats_Click">get stats</asp:LinkButton>
                </td>
                <td>
                    Sets all entries previously marked as &#39;PENDING&#39; to be available for update.<br />
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td align="right">
                    &nbsp;
                </td>
                <td colspan="2">
                    <asp:Label ID="lblStats" runat="server" Font-Bold="False"></asp:Label>
                </td>
                <td>
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td align="right">
                    &nbsp;
                </td>
                <td colspan="2">
                    <asp:Label ID="Label14" runat="server" Text="NOTES: Post code / zip code is required for ALL entries. State / region is required for US and CANADA." />
                </td>
                <td>
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
        </table>
        <br />
        &nbsp;<br />
        <asp:Label ID="lblError" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
            ForeColor="Red" />
        <br />
        <asp:Label ID="lblResults" runat="server"></asp:Label>
        <br />
        <br />
    </asp:Panel>
    </form>
    <script language="JavaScript" type="text/javascript" src="wz_tooltip.js"></script>
    <script language="JavaScript" type="text/javascript" src="library_functions.js"></script>
</body>
</html>