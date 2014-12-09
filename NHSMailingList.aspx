<%@ Page Language="VB" Theme="AIMSDefault" MaintainScrollPositionOnPostback="true" ValidateRequest="false" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Data.SqlTypes" %>
<%@ Import Namespace="System.Drawing.Image" %>
<%@ Import Namespace="System.Drawing.Color" %>
<%@ Import Namespace="System.Globalization" %>
<%@ Import Namespace="System.Threading" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%@ Import Namespace="System.Net" %>

<script runat="server">

    ' TO DO
    ' integrate help, esp quick start help

    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    Const ACCOUNT_CODE As String = "COURI11111"
    Const LICENSE_KEY As String = "RA61-XZ94-CT55-FH67"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsNumeric(Session("CustomerKey")) Then
            Server.Transfer("session_expired.aspx")
            Exit Sub
        End If
        If Not IsPostBack Then
            Call GetLookupSeed()
            Call DisableAllDateFields()
            tbQuickStart.Focus()
            Call UpdateStats()
        End If
        Call SetTitle()
        tbPostcode.Attributes.Add("onkeypress", "return clickButton(event,'" + btnFindAddress.ClientID + "')")
        'tbQuickStart.Attributes.Add("onkeypress", "return clickButton(event,'" + btnGoQuick.ClientID + "')")
    End Sub
    
    Protected Sub SetTitle()
        Dim sTitle As String = Session("SiteTitle")
        If sTitle <> String.Empty Then
            sTitle += " - "
        End If
        Page.Header.Title = sTitle & "NHS Mailing List"
    End Sub
   
    Protected Sub lnkbtnTitleMs_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TitleMs()
        tbFirstName.Focus()
    End Sub
    
    Protected Sub TitleMs()
        tbTitle.Text = "Ms"
        rblGenderFemale.Checked = True
    End Sub

    Protected Sub lnkbtnTitleMrs_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TitleMrs()
        tbFirstName.Focus()
    End Sub
    
    Protected Sub TitleMrs()
        tbTitle.Text = "Mrs"
        rblGenderFemale.Checked = True
    End Sub

    Protected Sub lnkbtnTitleMr_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TitleMr()
        tbFirstName.Focus()
    End Sub
    
    Protected Sub TitleMr()
        tbTitle.Text = "Mr"
        rblGenderMale.Checked = True
    End Sub
    
    Protected Sub ClearForm()
        tbQuickStart.Text = String.Empty
        tbTitle.Text = String.Empty
        tbFirstName.Text = String.Empty
        tbLastName.Text = String.Empty
        rblGenderFemale.Checked = False
        rblGenderMale.Checked = False
        tbAddr1.Text = String.Empty
        tbAddr2.Text = String.Empty
        tbAddr3.Text = String.Empty
        tbTown.Text = String.Empty
        tbCounty.Text = String.Empty
        tbPostcode.Text = String.Empty
        
        cb4269Gujarati.Checked = False
        cb3717_3984Gujarati.Checked = False
        cb4269Urdu.Checked = False
        cb3717_3984Urdu.Checked = False
        cb4269Mandarin.Checked = False
        cb3717_3984Mandarin.Checked = False
        cb4269Polish.Checked = False
        cb3717_3984Polish.Checked = False
        cb4269French.Checked = False
        cb3717_3984French.Checked = False
        cb4269Farsi.Checked = False
        cb3717_3984Farsi.Checked = False
        cb4269Spanish.Checked = False
        cb3717_3984Spanish.Checked = False
        cb3717_3984English.Checked = False
        cb4269EnglishLP.Checked = False
        cb3717_3984EnglishLP.Checked = False
        cb4269EnglishBraille.Checked = False
        cb3717_3984EnglishBraille.Checked = False
        cb4269EasyRead.Checked = False
        cb4269EnglishAudio.Checked = False

        cb3717_3984English.Font.Bold = False
        cb4269EnglishBraille.Font.Bold = False
        cb3717_3984EnglishBraille.Font.Bold = False
        cb4269EnglishLP.Font.Bold = False
        cb3717_3984EnglishLP.Font.Bold = False
        cb4269EasyRead.Font.Bold = False
        cb3717_3984EasyRead.Font.Bold = False
        cb4269EnglishAudio.Font.Bold = False
        cb3717_3984EnglishLAudio.Font.Bold = False
        cb4269Gujarati.Font.Bold = False
        cb3717_3984Gujarati.Font.Bold = False
        cb4269Urdu.Font.Bold = False
        cb3717_3984Urdu.Font.Bold = False
        cb4269Mandarin.Font.Bold = False
        cb3717_3984Mandarin.Font.Bold = False
        cb4269Polish.Font.Bold = False
        cb3717_3984Polish.Font.Bold = False
        cb4269French.Font.Bold = False
        cb3717_3984French.Font.Bold = False
        cb4269Farsi.Font.Bold = False
        cb3717_3984Farsi.Font.Bold = False
        cb4269Spanish.Font.Bold = False
        cb3717_3984Spanish.Font.Bold = False
        cbOptOut.Font.Bold = False
        cb3716CRSEasyRead.Font.Bold = False

        ddlPCT.SelectedIndex = 0
        
        cbBypassPostcodeValidation.Visible = False
        cbBypassPostcodeValidation.Checked = False
        'pnHouseNumber = 0
        psHouseNumber = String.Empty
        tfvPostcode.EnableClientScript = True
        tfvPostcode.Enabled = True
        btnBypassPostcode.Visible = False
        lblLegendPostcode.ForeColor = Red
        tbQuickStart.Focus()
    End Sub
    
    Protected Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Page.Validate()
        Dim sProduct As String = IsProductSelected()
        If sProduct = String.Empty Then
            WebMsgBox.Show("You must select at least one NHS product.")
            Exit Sub
        End If
        tbPostcode.Text = tbPostcode.Text.ToUpper.Trim
        If Not Regex.IsMatch(tbPostcode.Text, "(GIR 0AA)|((([A-Z-[QVX]][0-9][0-9]?)|(([A-Z-[QVX]][A-Z-[IJZ]][0-9][0-9]?)|(([A-Z-[QVX]][0-9][A-HJKSTUW])|([A-Z-[QVX]][A-Z-[IJZ]][0-9][ABEHMNPRVWXY])))) [0-9][A-Z-[CIKMOV]]{2})") Then
            If Not cbBypassPostcodeValidation.Visible Then
                cbBypassPostcodeValidation.Visible = True
                WebMsgBox.Show("The postcode you entered (" & tbPostcode.Text & ") does not appear to be valid. Either correct it, or click the 'bypass postcode validation' check box next to the save button.")
                Exit Sub
            ElseIf Not cbBypassPostcodeValidation.Checked Then
                WebMsgBox.Show("The postcode you entered (" & tbPostcode.Text & ") does not appear to be valid. Either correct it, or click the 'bypass postcode validation' check box next to the save button.")
                Exit Sub
            End If
        End If
        If Page.IsValid Then
            'Call ExecuteNonQuery(CreateEntry.ToString)
            sProduct = IsProductSelected(bClearProduct:=True)
            Do While sProduct <> String.Empty
                Call ExecuteNonQuery(CreateEntry(sProduct).ToString)
                sProduct = IsProductSelected(bClearProduct:=True)
            Loop
            lblMessage.Text = "Created entry. (LastName = " & tbLastName.Text & ", Addr1 = " & tbAddr1.Text & ", Town = " & tbTown.Text & ", PCT = " & ddlPCT.SelectedItem.Text & ")"
            Call ClearForm()
        End If
        lnkbtnResetPCTLookupSeed.Visible = False
        Call UpdateStats()
    End Sub
    
    Protected Function IsProductSelected(Optional ByVal bClearProduct As Boolean = False) As String
        IsProductSelected = String.Empty
        If cb3717_3984English.Checked Then
            IsProductSelected = "4269,English"
            If bClearProduct Then
                cb3717_3984English.Checked = False
            End If
            Exit Function
        End If
        If cb4269Gujarati.Checked Then
            IsProductSelected = "4269,Gujarati"
            If bClearProduct Then
                cb4269Gujarati.Checked = False
            End If
            Exit Function
        End If
        If cb3717_3984Gujarati.Checked Then
            IsProductSelected = "3717/3984,Gujarati"
            If bClearProduct Then
                cb3717_3984Gujarati.Checked = False
            End If
            Exit Function
        End If
        If cb4269Urdu.Checked Then
            IsProductSelected = "4269,Urdu"
            If bClearProduct Then
                cb4269Urdu.Checked = False
            End If
            Exit Function
        End If
        If cb3717_3984Urdu.Checked Then
            IsProductSelected = "3717/3984,Urdu"
            If bClearProduct Then
                cb3717_3984Urdu.Checked = False
            End If
            Exit Function
        End If
        If cb4269Mandarin.Checked Then
            IsProductSelected = "4269,Mandarin"
            If bClearProduct Then
                cb4269Mandarin.Checked = False
            End If
            Exit Function
        End If
        If cb3717_3984Mandarin.Checked Then
            IsProductSelected = "3717/3984,Mandarin"
            If bClearProduct Then
                cb3717_3984Mandarin.Checked = False
            End If
            Exit Function
        End If
        If cb4269Polish.Checked Then
            IsProductSelected = "4269,Polish"
            If bClearProduct Then
                cb4269Polish.Checked = False
            End If
            Exit Function
        End If
        If cb3717_3984Polish.Checked Then
            IsProductSelected = "3717/3984,Polish"
            If bClearProduct Then
                cb3717_3984Polish.Checked = False
            End If
            Exit Function
        End If
        If cb4269French.Checked Then
            IsProductSelected = "4269,French"
            If bClearProduct Then
                cb4269French.Checked = False
            End If
            Exit Function
        End If
        If cb3717_3984French.Checked Then
            IsProductSelected = "3717/3984,French"
            If bClearProduct Then
                cb3717_3984French.Checked = False
            End If
            Exit Function
        End If
        If cb4269Farsi.Checked Then
            IsProductSelected = "4269,Farsi"
            If bClearProduct Then
                cb4269Farsi.Checked = False
            End If
            Exit Function
        End If
        If cb3717_3984Farsi.Checked Then
            IsProductSelected = "3717/3984,Farsi"
            If bClearProduct Then
                cb3717_3984Farsi.Checked = False
            End If
            Exit Function
        End If
        If cb4269Spanish.Checked Then
            IsProductSelected = "4269,Spanish"
            If bClearProduct Then
                cb4269Spanish.Checked = False
            End If
            Exit Function
        End If
        If cb3717_3984Spanish.Checked Then
            IsProductSelected = "3717/3984,Spanish"
            If bClearProduct Then
                cb3717_3984Spanish.Checked = False
            End If
            Exit Function
        End If
        If cb4269EnglishBraille.Checked Then
            IsProductSelected = "4269,English Braille"
            If bClearProduct Then
                cb4269EnglishBraille.Checked = False
            End If
            Exit Function
        End If
        If cb3717_3984EnglishBraille.Checked Then
            IsProductSelected = "4269,English Braille"
            If bClearProduct Then
                cb3717_3984EnglishBraille.Checked = False
            End If
            Exit Function
        End If
        If cb4269EnglishLP.Checked Then
            IsProductSelected = "3717/3984,English Large Print"
            If bClearProduct Then
                cb4269EnglishLP.Checked = False
            End If
            Exit Function
        End If
        If cb3717_3984EnglishLP.Checked Then
            IsProductSelected = "4269,English Large Print"
            If bClearProduct Then
                cb3717_3984EnglishLP.Checked = False
            End If
            Exit Function
        End If
        If cb4269EasyRead.Checked Then
            IsProductSelected = "4269,English Easy Read"
            If bClearProduct Then
                cb4269EasyRead.Checked = False
            End If
            Exit Function
        End If
        If cb3717_3984EasyRead.Checked Then
            IsProductSelected = "3717/3984,English Easy Read"
            If bClearProduct Then
                cb3717_3984EasyRead.Checked = False
            End If
            Exit Function
        End If
        If cb4269EnglishAudio.Checked Then
            IsProductSelected = "4269,English Audio CD"
            If bClearProduct Then
                cb4269EnglishAudio.Checked = False
            End If
            Exit Function
        End If
        If cb3717_3984EnglishLAudio.Checked Then
            IsProductSelected = "3717/3984,English Audio CD"
            If bClearProduct Then
                cb3717_3984EnglishLAudio.Checked = False
            End If
            Exit Function
        End If
        If cbOptOut.Checked Then
            IsProductSelected = "OPT OUT,OPT OUT"
            If bClearProduct Then
                cbOptOut.Checked = False
            End If
            Exit Function
        End If
        If cb3716CRSEasyRead.Checked Then
            IsProductSelected = "3716,CRS Easy Read"
            If bClearProduct Then
                cb3716CRSEasyRead.Checked = False
            End If
            Exit Function
        End If
    End Function
    
    Protected Function CreateEntry(ByVal sProductInfo As String) As StringBuilder
        Dim sProduct() As String = sProductInfo.Split(",")
        Dim sProductCode As String = sProduct(0)
        Dim sProductLanguage As String = sProduct(1)
        Dim sbSQL As New StringBuilder
        sbSQL.Append("")
        sbSQL.Append("INSERT INTO NHSMailingList (InputDate, Title, FirstName, LastName, Gender, Addr1, Addr2, Addr3, Town, County, PostCode, ContactType, Product, Language, [4269Gujarati], [3717_3984Gujarati], [4269Urdu], [3717_3984Urdu], [4269Mandarin], [3717_3984Mandarin], [4269Polish], [3717_3984Polish], [4269French], [3717_3984French], [4269Farsi], [3717_3984Farsi], [4269Spanish], [3717_3984Spanish], [3717_3984English], [4269EnglishLP], [3717_3984EnglishLP], [4269EnglishBraille], [3717_3984EnglishBraille], [4269EasyRead], [3717_3984EasyRead], [4269EnglishAudio], [3717_3984EnglishAudio], OptOut, [3716CRSEasyRead], PCT, LastUpdatedOn, LastUpdatedBy) VALUES (GETDATE(), ")
        sbSQL.Append("'")
        sbSQL.Append(tbTitle.Text.Replace("'", "''"))
        sbSQL.Append("'")
        sbSQL.Append(", ")
        sbSQL.Append("'")
        sbSQL.Append(tbFirstName.Text.Replace("'", "''"))
        sbSQL.Append("'")
        sbSQL.Append(", ")
        sbSQL.Append("'")
        sbSQL.Append(tbLastName.Text.Replace("'", "''"))
        sbSQL.Append("'")
        sbSQL.Append(", ")
        sbSQL.Append("'")
        If rblGenderFemale.Checked Then
            sbSQL.Append("F")
        ElseIf rblGenderMale.Checked Then
            sbSQL.Append("M")
        Else
            sbSQL.Append("U")
        End If
        sbSQL.Append("'")
        sbSQL.Append(", ")
        sbSQL.Append("'")
        sbSQL.Append(tbAddr1.Text.Replace("'", "''"))
        sbSQL.Append("'")
        sbSQL.Append(", ")
        sbSQL.Append("'")
        sbSQL.Append(tbAddr2.Text.Replace("'", "''"))
        sbSQL.Append("'")
        sbSQL.Append(", ")
        sbSQL.Append("'")
        sbSQL.Append(tbAddr3.Text.Replace("'", "''"))
        sbSQL.Append("'")
        sbSQL.Append(", ")
        sbSQL.Append("'")
        sbSQL.Append(tbTown.Text.Replace("'", "''"))
        sbSQL.Append("'")
        sbSQL.Append(", ")
        sbSQL.Append("'")
        sbSQL.Append(tbCounty.Text.Replace("'", "''"))
        sbSQL.Append("'")
        sbSQL.Append(", ")
        sbSQL.Append("'")
        sbSQL.Append(tbPostcode.Text.Replace("'", "''"))
        sbSQL.Append("'")
        sbSQL.Append(", ")
        sbSQL.Append("'")
        If rblContactMethodOrderForm.Checked Then
            sbSQL.Append("O")
        Else
            sbSQL.Append("T")
        End If
        sbSQL.Append("'")
        sbSQL.Append(", ")
        sbSQL.Append("'")
        sbSQL.Append(sProductCode.Replace("'", "''"))
        sbSQL.Append("'")
        sbSQL.Append(", ")
        sbSQL.Append("'")
        sbSQL.Append(sProductLanguage.Replace("'", "''"))
        sbSQL.Append("'")
        sbSQL.Append(", ")
        If cb4269Gujarati.Checked Then
            sbSQL.Append("1")
        Else
            sbSQL.Append("0")
        End If
        sbSQL.Append(", ")
        If cb3717_3984Gujarati.Checked Then
            sbSQL.Append("1")
        Else
            sbSQL.Append("0")
        End If
        sbSQL.Append(", ")
        If cb4269Urdu.Checked Then
            sbSQL.Append("1")
        Else
            sbSQL.Append("0")
        End If
        sbSQL.Append(", ")
        If cb3717_3984Urdu.Checked Then
            sbSQL.Append("1")
        Else
            sbSQL.Append("0")
        End If
        sbSQL.Append(", ")
        If cb4269Mandarin.Checked Then
            sbSQL.Append("1")
        Else
            sbSQL.Append("0")
        End If
        sbSQL.Append(", ")
        If cb3717_3984Mandarin.Checked Then
            sbSQL.Append("1")
        Else
            sbSQL.Append("0")
        End If
        sbSQL.Append(", ")
        If cb4269Polish.Checked Then
            sbSQL.Append("1")
        Else
            sbSQL.Append("0")
        End If
        sbSQL.Append(", ")
        If cb3717_3984Polish.Checked Then
            sbSQL.Append("1")
        Else
            sbSQL.Append("0")
        End If
        sbSQL.Append(", ")
        If cb4269French.Checked Then
            sbSQL.Append("1")
        Else
            sbSQL.Append("0")
        End If
        sbSQL.Append(", ")
        If cb3717_3984French.Checked Then
            sbSQL.Append("1")
        Else
            sbSQL.Append("0")
        End If
        sbSQL.Append(", ")
        If cb4269Farsi.Checked Then
            sbSQL.Append("1")
        Else
            sbSQL.Append("0")
        End If
        sbSQL.Append(", ")
        If cb3717_3984Farsi.Checked Then
            sbSQL.Append("1")
        Else
            sbSQL.Append("0")
        End If
        sbSQL.Append(", ")
        If cb4269Spanish.Checked Then
            sbSQL.Append("1")
        Else
            sbSQL.Append("0")
        End If
        sbSQL.Append(", ")
        If cb3717_3984Spanish.Checked Then
            sbSQL.Append("1")
        Else
            sbSQL.Append("0")
        End If
        sbSQL.Append(", ")
        If cb3717_3984English.Checked Then
            sbSQL.Append("1")
        Else
            sbSQL.Append("0")
        End If
        sbSQL.Append(", ")
        If cb4269EnglishLP.Checked Then
            sbSQL.Append("1")
        Else
            sbSQL.Append("0")
        End If
        sbSQL.Append(", ")
        If cb3717_3984EnglishLP.Checked Then
            sbSQL.Append("1")
        Else
            sbSQL.Append("0")
        End If
        sbSQL.Append(", ")
        If cb4269EnglishBraille.Checked Then
            sbSQL.Append("1")
        Else
            sbSQL.Append("0")
        End If
        sbSQL.Append(", ")
        If cb3717_3984EnglishBraille.Checked Then
            sbSQL.Append("1")
        Else
            sbSQL.Append("0")
        End If
        sbSQL.Append(", ")
        If cb4269EasyRead.Checked Then
            sbSQL.Append("1")
        Else
            sbSQL.Append("0")
        End If
        sbSQL.Append(", ")
        If cb3717_3984EasyRead.Checked Then
            sbSQL.Append("1")
        Else
            sbSQL.Append("0")
        End If
        sbSQL.Append(", ")
        If cb4269EnglishAudio.Checked Then
            sbSQL.Append("1")
        Else
            sbSQL.Append("0")
        End If
        sbSQL.Append(", ")
        If cb3717_3984EnglishLAudio.Checked Then
            sbSQL.Append("1")
        Else
            sbSQL.Append("0")
        End If
        sbSQL.Append(", ")
        If cbOptOut.Checked Then
            sbSQL.Append("1")
        Else
            sbSQL.Append("0")
        End If
        sbSQL.Append(", ")
        If cb3716CRSEasyRead.Checked Then
            sbSQL.Append("1")
        Else
            sbSQL.Append("0")
        End If
        sbSQL.Append(", ")
        sbSQL.Append("'")
        sbSQL.Append(ddlPCT.SelectedValue)
        sbSQL.Append("'")
        sbSQL.Append(", ")
        sbSQL.Append("GETDATE(), ")
        sbSQL.Append(Session("UserKey"))
        sbSQL.Append(")")
        CreateEntry = sbSQL
    End Function
    
    Protected Sub DisableAllDateFields()
        tbDateSince.Enabled = False
        tbDateFrom.Enabled = False
        tbDateTo.Enabled = False
    End Sub
    
    Protected Sub rblExportEntriesSince_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call DisableAllDateFields()
        tbDateSince.Enabled = True
        btnExport.Enabled = True
        tbDateSince.Focus()
        tbDateFrom.Enabled = False
        tbDateFrom.Text = String.Empty
        tbDateTo.Enabled = False
        tbDateTo.Text = String.Empty
    End Sub

    Protected Sub rblExportEntriesBetween_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call DisableAllDateFields()
        tbDateFrom.Enabled = True
        tbDateTo.Enabled = True
        btnExport.Enabled = True
        tbDateSince.Enabled = False
        tbDateSince.Text = String.Empty
        tbDateFrom.Focus()
    End Sub

    Protected Sub rblExportAllEntries_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call DisableAllDateFields()
        btnExport.Enabled = True
        btnExport.Focus()
        tbDateFrom.Enabled = False
        tbDateFrom.Text = String.Empty
        tbDateTo.Enabled = False
        tbDateTo.Text = String.Empty
        tbDateSince.Enabled = False
        tbDateSince.Text = String.Empty
    End Sub
    
    Protected Sub lnkbtnLastSunday_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim dtDay As Date = Date.Today.AddDays(-1)
        While Not dtDay.DayOfWeek = DayOfWeek.Sunday
            dtDay = dtDay.AddDays(-1)
        End While
        tbDateSince.Text = dtDay.ToString("dd-MMM-yyyy")
        rblExportAllEntries.Checked = False
        rblExportEntriesBetween.Checked = False
        rblExportEntriesSince.Checked = True
        tbDateSince.Enabled = True
        btnExport.Enabled = True
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

    Protected Sub lnkbtnClearForm_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ClearForm()
    End Sub
    
    Protected Sub lnkbtnPCT_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim lb As LinkButton = sender
        Dim sLegend As String = lb.Text
        For i As Integer = 1 To ddlPCT.Items.Count - 1
            If ddlPCT.Items(i).Text = sLegend Then
                ddlPCT.SelectedIndex = i
                Exit For
            End If
        Next
    End Sub
    
    Protected Sub btnExport_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim dtDate1 As Date
        Dim dtDate2 As Date
        If rblExportEntriesSince.Checked Then
            dtDate2 = Date.Parse("1-Jan-2049")
            If IsDate(tbDateSince.Text) Then
                dtDate1 = Date.Parse(tbDateSince.Text)
                If dtDate1 > Today Then
                    WebMsgBox.Show("You have specified a date in the future.")
                    Exit Sub
                End If
            Else
                WebMsgBox.Show("Please specify a valid date.")
                Exit Sub
            End If
        ElseIf rblExportEntriesBetween.Checked Then
            If IsDate(tbDateFrom.Text) Then
                dtDate1 = Date.Parse(tbDateFrom.Text)
            Else
                WebMsgBox.Show("Please specify a valid FROM date.")
                Exit Sub
            End If
            If IsDate(tbDateTo.Text) Then
                dtDate2 = Date.Parse(tbDateTo.Text)
            Else
                WebMsgBox.Show("Please specify a valid TO date.")
                Exit Sub
            End If
            If dtDate1 > dtDate2 Then
                WebMsgBox.Show("TO date cannot precede FROM date.")
                Exit Sub
            End If
        Else
            dtDate1 = Date.Parse("1-Jan-2000")
            dtDate2 = Date.Parse("1-Jan-2049")
        End If
        Call ExportData(dtDate1:=dtDate1, dtDate2:=dtDate2)
    End Sub
    
    Protected Function SafeDate(ByVal dt As Date) As String
        Dim arrMonths() As String = {"", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"}
        SafeDate = dt.Day & "-" & arrMonths(dt.Month) & "-" & dt.Year
    End Function
    
    Protected Sub ExportData(ByVal dtDate1 As Date, ByVal dtDate2 As Date)
        'Dim sSQL As String = "SELECT * FROM NHSMailingList WHERE InputDate >= '" & SafeDate(dtDate1) & "' AND InputDate <= '" & SafeDate(dtDate2) & "' ORDER BY InputDate"
        Dim sSQL As String = "SELECT InputDate 'Input Date', nml.Title, nml.FirstName 'First Name', nml.LastName 'Last Name', Gender, Addr1 'Addr 1', Addr2 'Addr 2', Addr3 'Addr 3', Town, County, Postcode 'Post code', ContactType 'Contact Type', Product, Language, PCT, UserId 'Entered By' FROM NHSMailingList nml INNER JOIN UserProfile up ON nml.LastUpdatedBy = up.[Key] WHERE InputDate >= '" & SafeDate(dtDate1) & "' AND InputDate <= '" & SafeDate(dtDate2) & "' ORDER BY InputDate"
        Dim oDataTable As DataTable = ExecuteQueryToDataTable(sSQL)
        If oDataTable.Rows.Count > 0 Then
            Dim sResponseValue As New StringBuilder
            Response.ContentType = "text/csv"
            sResponseValue.Append("attachment; filename=""NHSMailingList.csv""")
            Response.AddHeader("Content-Disposition", "attachment; filename=""NHSMailingList.csv""")

            Dim IgnoredItems As New ArrayList
            IgnoredItems.Add("")   ' add name of any field that is not to be output
            
            'IgnoredItems.Add("id")
            'IgnoredItems.Add("4269Gujarati")
            'IgnoredItems.Add("3717_3984Gujarati")
            'IgnoredItems.Add("4269Urdu")
            'IgnoredItems.Add("3717_3984Urdu")
            'IgnoredItems.Add("4269Mandarin")
            'IgnoredItems.Add("3717_3984Mandarin")
            'IgnoredItems.Add("4269Polish")
            'IgnoredItems.Add("3717_3984Polish")
            'IgnoredItems.Add("4269French")
            'IgnoredItems.Add("3717_3984French")
            'IgnoredItems.Add("4269Farsi")
            'IgnoredItems.Add("3717_3984Farsi")
            'IgnoredItems.Add("4269Spanish")
            'IgnoredItems.Add("3717_3984Spanish")
            'IgnoredItems.Add("3717_3984English")
            'IgnoredItems.Add("4269EnglishLP")
            'IgnoredItems.Add("3717_3984EnglishLP")
            'IgnoredItems.Add("4269EnglishBraille")
            'IgnoredItems.Add("3717_3984EnglishBraille")
            'IgnoredItems.Add("4269EasyRead")
            'IgnoredItems.Add("3717_3984EasyRead")
            'IgnoredItems.Add("4269EnglishAudio")
            'IgnoredItems.Add("3717_3984EnglishAudio")
            'IgnoredItems.Add("OptOut")
            'IgnoredItems.Add("3716CRSEasyRead")
            'IgnoredItems.Add("LastUpdatedOn")
            'IgnoredItems.Add("LastUpdatedBy")

            For Each c As DataColumn In oDataTable.Columns
                If Not IgnoredItems.Contains(c.ColumnName) Then
                    Response.Write(c.ColumnName)
                    Response.Write(",")
                End If
            Next
            Response.Write(vbCrLf)

            Dim sItem As String
            For Each dr As DataRow In oDataTable.Rows
                For Each c As DataColumn In oDataTable.Columns
                    If Not IgnoredItems.Contains(c.ColumnName) Then
                        sItem = (dr(c.ColumnName).ToString)
                        sItem = sItem.Replace(ControlChars.Quote, ControlChars.Quote & ControlChars.Quote)
                        sItem = ControlChars.Quote & sItem & ControlChars.Quote
                        Response.Write(sItem)
                        Response.Write(",")
                    End If
                Next
                Response.Write(vbCrLf)
            Next
            Response.End()
        End If
    End Sub
    
    Protected Sub btnFindAddress_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call FindAddress()
        Call ProcessPCTLookupPage()
    End Sub

    Protected Sub lnkbtnAddrLookupCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call CancelLookup()
        tbPostcode.Focus()
    End Sub

    Protected Sub CancelLookup()
        trIdentity.Visible = True
        trProducts.Visible = True
        trAddress1.Visible = True
        trAddress2.Visible = True
        trTownCity.Visible = True
        trPCT.Visible = True
        trControls.Visible = True
        trPostcodeLookupResults.Visible = False
        lblLookupError.Text = String.Empty
    End Sub
    
    Protected Sub lbLookupResults_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        SelectAddress(lbLookupResults.SelectedValue)
    End Sub
    
    Protected Sub SelectAddress(ByVal sAddressId As String)
        Dim objLookup As New uk.co.postcodeanywhere.services.LookupUK
        Dim objAddressResults As uk.co.postcodeanywhere.services.AddressResults
        Dim objAddress As uk.co.postcodeanywhere.services.Address

        objAddressResults = objLookup.FetchAddress(sAddressId, _
           uk.co.postcodeanywhere.services.enLanguage.enLanguageEnglish, _
           uk.co.postcodeanywhere.services.enContentType.enContentStandardAddress, _
           ACCOUNT_CODE, LICENSE_KEY, "")
        objLookup.Dispose()
        
        trIdentity.Visible = True
        trAddress1.Visible = True
        trAddress2.Visible = True
        trTownCity.Visible = True
        trProducts.Visible = True
        trPCT.Visible = True
        trControls.Visible = True
        trPostcodeLookupResults.Visible = False

        btnSave.Focus()
        
        If objAddressResults.IsError Then
            lblLookupError.Text = objAddressResults.ErrorMessage
        Else
            objAddress = objAddressResults.Results(0)

            tbLastName.Text = tbLastName.Text.Trim
            If tbLastName.Text = String.Empty Then
                If objAddress.OrganisationName.Trim <> String.Empty Then
                    tbLastName.Text = objAddress.OrganisationName
                End If
            End If
            tbAddr1.Text = objAddress.Line1
            tbAddr2.Text = objAddress.Line2
            tbAddr3.Text = objAddress.Line3
            tbTown.Text = objAddress.PostTown
            tbCounty.Text = objAddress.County
            tbPostcode.Text = objAddress.Postcode

            Dim sHostIPAddress As String = Server.HtmlEncode(Request.UserHostName)
            If Not ExecuteNonQuery("INSERT INTO PostCodeLookup (ClientIPAddress, LookupDateTime, CustomerKey, CostInUnits) VALUES ('" & sHostIPAddress & "', GETDATE(), " & 0 & ", 1)") Then
                WebMsgBox.Show("Error in lbLookupResults_SelectedIndexChanged, Could not log lookup")
            End If
        End If
    End Sub
    
    Protected Sub FindAddress()
        trIdentity.Visible = False
        trAddress1.Visible = False
        trAddress2.Visible = False
        trTownCity.Visible = False
        trProducts.Visible = False
        trPCT.Visible = False
        trControls.Visible = False
        trPostcodeLookupResults.Visible = True

        lblLookupError.Text = String.Empty
        tbPostcode.Text = tbPostcode.Text.Trim.ToUpper

        Dim objLookup As New uk.co.postcodeanywhere.services.LookupUK
        Dim objInterimResults As uk.co.postcodeanywhere.services.InterimResults
        Dim objInterimResult As uk.co.postcodeanywhere.services.InterimResult

        objInterimResults = objLookup.ByPostcode(tbPostcode.Text, ACCOUNT_CODE, LICENSE_KEY, "")
        objLookup.Dispose()
        
        If objInterimResults.IsError OrElse objInterimResults.Results Is Nothing OrElse objInterimResults.Results.GetLength(0) = 0 Then
            lblLookupError.Visible = True
            lbLookupResults.Visible = False
            lblSelectADestination.Visible = False
            lblLookupError.Text = objInterimResults.ErrorMessage
            If lblLookupError.Text.Trim = String.Empty Then
                lblLookupError.Text = "<br />No results found for this post code"
            Else
                lblLookupError.Text = "<br />" & lblLookupError.Text
            End If
            btnBypassPostcode.Visible = True
        Else
            lblLookupError.Visible = False
            lbLookupResults.Visible = True
            lblSelectADestination.Visible = True
            lbLookupResults.Items.Clear()
            Dim sHouseNumber As String = psHouseNumber & " "
            Dim nHouseNumberMatches As Integer = 0
            Dim sMatchedId As String = String.Empty
            If Not objInterimResults.Results Is Nothing Then
                For Each objInterimResult In objInterimResults.Results
                    If objInterimResult.Description.ToLower.StartsWith(sHouseNumber) Then
                        nHouseNumberMatches += 1
                        sMatchedId = objInterimResult.Id
                    End If
                    lbLookupResults.Items.Add(New ListItem(objInterimResult.Description, objInterimResult.Id))
                Next
            End If
            Dim sHostIPAddress As String = Server.HtmlEncode(Request.UserHostName)
            If Not ExecuteNonQuery("INSERT INTO PostCodeLookup (ClientIPAddress, LookupDateTime, CustomerKey, CostInUnits) VALUES ('" & sHostIPAddress & "', GETDATE(), " & 0 & ", 0)") Then
                WebMsgBox.Show("Error in lnkbtnFindAddress_Click, could not log lookup")
            End If
            If nHouseNumberMatches = 1 Then
                Call SelectAddress(sMatchedId)
            Else
                lbLookupResults.Focus()
            End If
        End If
    End Sub

    Protected Sub lnkbtnShowMostRecentEntries_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call GetMostRecentEntries()
    End Sub
        
    Protected Sub GetMostRecentEntries()
        Dim oDataTable As DataTable
        Dim sSQL As String = "SELECT TOP 10 InputDate 'Input Date', Title, FirstName 'First Name', LastName 'Last Name', Gender, Addr1 'Addr 1', Addr2 'Addr 2', Addr3 'Addr 3', Town, County, PostCode 'Post Code', ContactType 'Contact Type', Product, Language, PCT, LastUpdatedBy 'Entered By' FROM NHSMailingList ORDER BY InputDate DESC"
        oDataTable = ExecuteQueryToDataTable(sSQL)
        gvRecentEntries.DataSource = oDataTable
        gvRecentEntries.DataBind()
        gvRecentEntries.Visible = True
        lnkbtnRefresh.Visible = True
        lnkbtnHideMostRecentEntries.Visible = True
        lnkbtnShowMostRecentEntries.Visible = False
    End Sub
    
    Protected Sub lnkbtnRefresh_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call GetMostRecentEntries()
    End Sub

    Protected Sub lnkbtnHideMostRecentEntries_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        gvRecentEntries.Visible = False
        lnkbtnHideMostRecentEntries.Visible = False
        lnkbtnRefresh.Visible = False
        lnkbtnShowMostRecentEntries.Visible = True
    End Sub
    
    Protected Function GetLookupSeed() As String
        GetLookupSeed = String.Empty
        Dim sPage As String = RetrievePCTLookupPage()
        Dim sTemp As String
        Dim nStartPos As Integer
        Dim sName As String
        Dim sValue As String

        nStartPos = sPage.IndexOf("hidden") + 13
        sTemp = sPage.Substring(nStartPos, 50)
        sName = sTemp.Substring(1, 5)
        sValue = sTemp.Substring(15, 25)
        psPCTLookupSeed = sName & "=" & sValue
    End Function
    
    Protected Function RetrievePCTLookupPage() As String
        Dim sURL As String = "https://www.ndtms.org.uk/emids/cgi-bin/ons_locale.cgi"
        RetrievePCTLookupPage = String.Empty
        Dim wr As System.Net.WebRequest
        Try
            wr = WebRequest.Create(sURL)
            Dim resp As HttpWebResponse = wr.GetResponse()
            Dim sr As New StreamReader(resp.GetResponseStream)
            RetrievePCTLookupPage = sr.ReadToEnd().Trim
            sr.Close()
        Catch ex As Exception
            WebMsgBox.Show("Error in RetrievePCTLookupPage: (" & sURL & " ): " & ex.Message)
        End Try
    End Function

    Protected Function LookupPCT() As String
        Dim sURL As String = "https://www.ndtms.org.uk/emids/cgi-bin/ons_locale.cgi"
        LookupPCT = String.Empty
        Dim wr As System.Net.WebRequest
        Try
            Dim sPostData As String = psPCTLookupSeed & "&newpct=on&pc=" & tbPostcode.Text.Trim.Replace(" ", "")
            Dim encoding As ASCIIEncoding = New ASCIIEncoding
            Dim data() As Byte = encoding.GetBytes(sPostData)
            
            wr = WebRequest.Create(sURL)
            wr.Method = "POST"
            wr.ContentType = "application/x-www-form-urlencoded"
            wr.ContentLength = data.Length
            
            Dim os As System.IO.Stream = wr.GetRequestStream()
            os.Write(data, 0, data.Length)
            os.Close()
            
            Dim resp As HttpWebResponse = wr.GetResponse()
            Dim sr As New StreamReader(resp.GetResponseStream)
            LookupPCT = sr.ReadToEnd().Trim
            sr.Close()
        Catch ex As Exception
            WebMsgBox.Show("Error in LookupPCT: (" & sURL & " ): " & ex.Message)
        End Try
    End Function
  
    Protected Sub ProcessPCTLookupPage()
        Dim sResult As String = LookupPCT()
        If sResult.Contains("Probability") Then
            Dim nStartPos As Integer
            Dim nEndPos As Integer
            Dim sPCTCode As String
            Try
                nStartPos = sResult.IndexOf("Probability")
                nEndPos = sResult.Length
                sResult = sResult.Substring(nStartPos, nEndPos - nStartPos)
                nStartPos = 46
                nEndPos = sResult.IndexOf("%")
                sResult = sResult.Substring(nStartPos, nEndPos - nStartPos)
                nStartPos = sResult.IndexOf("/td") + 11
                nEndPos = sResult.IndexOf("align") - 6
                sResult = sResult.Substring(nStartPos, nEndPos - nStartPos)
                nEndPos = sResult.IndexOf("</td")
                sResult = sResult.Substring(0, nEndPos)
                lblPCT.Text = "PCT: " & sResult
                nStartPos = sResult.IndexOf("(") + 1
                nEndPos = sResult.IndexOf(")")
                sPCTCode = sResult.Substring(nStartPos, nEndPos - nStartPos)
            Catch
                sPCTCode = "XXX"
            End Try
            If sPCTCode.Length = 3 Then
                Dim bMatched As Boolean = False
                For i As Integer = 1 To ddlPCT.Items.Count - 1
                    If ddlPCT.Items(i).Value = sPCTCode Then
                        ddlPCT.SelectedIndex = i
                        bMatched = True
                        Exit For
                    End If
                Next
                If Not bMatched Then
                    Call SendMail("MISSING PCT", "chris.newport@sprintexpress.co.uk", "Could not parse PCT page for postcode " & tbPostcode.Text, sResult, sResult)
                    lblPCT.Text = "Failed searching for PCT."
                    btnBypassPostcode.Visible = True
                End If
            Else
                lblPCT.Text = "Inexplicable error - PCT code length <> 3 (" & sPCTCode & ")"
            End If
        Else
            'lblPCT.Text = "Could not match this postcode to a PCT (" & psPCTLookupSeed & ")"
            lblPCT.Text = "Could not match this postcode to a PCT."
            lnkbtnResetPCTLookupSeed.Visible = True
            btnBypassPostcode.Visible = True
        End If
        tbQuickStart.Text = String.Empty
    End Sub

    Protected Sub btnGoQuick_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ProcessQuickStart()
    End Sub
    
    Protected Sub ProcessQuickStart()
        Dim arrQS() As String
        Dim sUnprocessed As String = String.Empty
        Dim sSeparator As String = String.Empty
        Dim sQuickStart As String = tbQuickStart.Text.Trim
        If sQuickStart.Contains(";") Then
            sSeparator = ";"
        ElseIf sQuickStart.Contains(",") Then
            sSeparator = ","
        ElseIf sQuickStart.Contains(".") Then
            sSeparator = "."
        ElseIf sQuickStart.Contains(" ") Then
            sSeparator = " "
        End If
        If sSeparator = String.Empty Then
            sQuickStart = String.Empty
            Exit Sub
        End If
        If sQuickStart.Length > 0 Then
            arrQS = sQuickStart.Split(sSeparator)
            If arrQS.Length >= 3 Then
                Dim nElementCount As Integer = 1
                For Each sElement As String In arrQS
                    If Not ProcessElement(sElement, nElementCount) Then
                        sUnprocessed += sElement & ";"
                        nElementCount += 1
                    End If
                Next
                arrQS = sUnprocessed.Split(";")
                If arrQS.Length >= 3 Then
                    tbFirstName.Text = UpcaseFirstLetter(arrQS(0).ToLower)
                    tbLastName.Text = UpcaseFirstLetter(arrQS(1).ToLower)
                    tbPostcode.Text = arrQS(2).ToUpper
                    Call FindAddress()
                    Call ProcessPCTLookupPage()
                End If
            Else
                sQuickStart = String.Empty
                Exit Sub
            End If
        Else
            tbFirstName.Focus()
        End If
    End Sub
    
    Protected Function ProcessElement(ByVal sElement As String, ByVal nElementCount As Integer) As Boolean
        ProcessElement = False
        sElement = sElement.Trim.ToLower
        If sElement = String.Empty Then
            ProcessElement = True
            Exit Function
        End If
        If sElement = "mr" Then
            Call TitleMr()
            ProcessElement = True
            Exit Function
        End If
        If sElement = "mrs" Then
            Call TitleMrs()
            ProcessElement = True
            Exit Function
        End If
        If sElement = "ms" Or sElement = "miss" Then
            Call TitleMs()
            ProcessElement = True
            Exit Function
        End If
        If sElement = "3en" Then
            cb3717_3984English.Checked = True
            cb3717_3984English.Font.Bold = True
            ProcessElement = True
            Exit Function
        End If
        If sElement = "4br" Then
            cb4269EnglishBraille.Checked = True
            cb4269EnglishBraille.Font.Bold = True
            ProcessElement = True
            Exit Function
        End If
        If sElement = "3br" Then
            cb3717_3984EnglishBraille.Checked = True
            cb3717_3984EnglishBraille.Font.Bold = True
            ProcessElement = True
            Exit Function
        End If
        If sElement = "4lp" Then
            cb4269EnglishLP.Checked = True
            cb4269EnglishLP.Font.Bold = True
            ProcessElement = True
            Exit Function
        End If
        If sElement = "3lp" Then
            cb3717_3984EnglishLP.Checked = True
            cb3717_3984EnglishLP.Font.Bold = True
            ProcessElement = True
            Exit Function
        End If
        If sElement = "4er" Then
            cb4269EasyRead.Checked = True
            cb4269EasyRead.Font.Bold = True
            ProcessElement = True
            Exit Function
        End If
        If sElement = "3er" Then
            cb3717_3984EasyRead.Checked = True
            cb3717_3984EasyRead.Font.Bold = True
            ProcessElement = True
            Exit Function
        End If
        If sElement = "4au" Then
            cb4269EnglishAudio.Checked = True
            cb4269EnglishAudio.Font.Bold = True
            ProcessElement = True
            Exit Function
        End If
        If sElement = "3au" Then
            cb3717_3984EnglishLAudio.Checked = True
            cb3717_3984EnglishLAudio.Font.Bold = True
            ProcessElement = True
            Exit Function
        End If
        If sElement = "4gu" Then
            cb4269Gujarati.Checked = True
            cb4269Gujarati.Font.Bold = True
            ProcessElement = True
            Exit Function
        End If
        If sElement = "3gu" Then
            cb3717_3984Gujarati.Checked = True
            cb3717_3984Gujarati.Font.Bold = True
            ProcessElement = True
            Exit Function
        End If
        If sElement = "4ur" Then
            cb4269Urdu.Checked = True
            cb4269Urdu.Font.Bold = True
            ProcessElement = True
            Exit Function
        End If
        If sElement = "3ur" Then
            cb3717_3984Urdu.Checked = True
            cb3717_3984Urdu.Font.Bold = True
            ProcessElement = True
            Exit Function
        End If
        If sElement = "4ma" Then
            cb4269Mandarin.Checked = True
            cb4269Mandarin.Font.Bold = True
            ProcessElement = True
            Exit Function
        End If
        If sElement = "3ma" Then
            cb3717_3984Mandarin.Checked = True
            cb3717_3984Mandarin.Font.Bold = True
            ProcessElement = True
            Exit Function
        End If
        If sElement = "4po" Then
            cb4269Polish.Checked = True
            cb4269Polish.Font.Bold = True
            ProcessElement = True
            Exit Function
        End If
        If sElement = "3po" Then
            cb3717_3984Polish.Checked = True
            cb3717_3984Polish.Font.Bold = True
            ProcessElement = True
            Exit Function
        End If
        If sElement = "4fr" Then
            cb4269French.Checked = True
            cb4269French.Font.Bold = True
            ProcessElement = True
            Exit Function
        End If
        If sElement = "3fr" Then
            cb3717_3984French.Checked = True
            cb3717_3984French.Font.Bold = True
            ProcessElement = True
            Exit Function
        End If
        If sElement = "4fa" Then
            cb4269Farsi.Checked = True
            cb4269Farsi.Font.Bold = True
            ProcessElement = True
            Exit Function
        End If
        If sElement = "3fa" Then
            cb3717_3984Farsi.Checked = True
            cb3717_3984Farsi.Font.Bold = True
            ProcessElement = True
            Exit Function
        End If
        If sElement = "4sp" Then
            cb4269Spanish.Checked = True
            cb4269Spanish.Font.Bold = True
            ProcessElement = True
            Exit Function
        End If
        If sElement = "3sp" Then
            cb3717_3984Spanish.Checked = True
            cb3717_3984Spanish.Font.Bold = True
            ProcessElement = True
            Exit Function
        End If
        If sElement = "opt" Then
            cbOptOut.Checked = True
            cbOptOut.Font.Bold = True
            ProcessElement = True
            Exit Function
        End If
        If sElement = "crs" Then
            cb3716CRSEasyRead.Checked = True
            cb3716CRSEasyRead.Font.Bold = True
            ProcessElement = True
            Exit Function
        End If
        If sElement = "ord" Then
            rblContactMethodTelephone.Checked = False
            rblContactMethodOrderForm.Checked = True
            ProcessElement = True
            Exit Function
        End If
        If sElement = "tel" Then
            rblContactMethodTelephone.Checked = True
            ProcessElement = True
            Exit Function
        End If
        If IsNumeric(sElement.Substring(0, 1)) Then
            psHouseNumber = sElement.ToLower
            ProcessElement = True
            Exit Function
        End If
    End Function

    Function UpcaseFirstLetter(ByVal a As String) As String
        UpcaseFirstLetter = String.Empty
        a = a.Trim
        If a.Length > 0 Then
            UpcaseFirstLetter = a.Substring(0, 1).ToUpper
        End If
        If a.Length > 1 Then
            UpcaseFirstLetter += a.Substring(1, a.Length - 1)
        End If
    End Function
    
    Property psPCTLookupSeed() As String
        Get
            Dim o As Object = ViewState("PCTLookupSeed")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("PCTLookupSeed") = Value
        End Set
    End Property

    Protected Sub lnkbtnResetPCTLookupSeed_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call GetLookupSeed()
    End Sub
    
    Protected Sub UpdateStats()
        Dim sSQL As String = "SELECT COUNT(DISTINCT FirstName + LastName + Addr1 + Addr2 + Postcode) FROM NHSMailingList WHERE LastUpdatedOn >= '" & Today.ToString("dd-MMM-yyyy") & " 00:00:00' AND LastUpdatedBy = " & Session("UserKey")
        Dim oDataTable As DataTable = ExecuteQueryToDataTable(sSQL)
        lblStats.Text = "Entered Today: " & oDataTable.Rows(0).Item(0)
    End Sub
    
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
    
    Property psHouseNumber() As String
        Get
            Dim o As Object = ViewState("ML_HouseNumber")
            If o Is Nothing Then
                Return String.Empty
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("ML_HouseNumber") = Value
        End Set
    End Property

    Protected Sub btnBypassPostcode_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call CancelLookup()
        tbPostcode.Text = String.Empty
        tfvPostcode.EnableClientScript = False
        tfvPostcode.Enabled = False
        cbBypassPostcodeValidation.Visible = True
        cbBypassPostcodeValidation.Checked = True
        lblLegendPostcode.ForeColor = Black
        btnBypassPostcode.Visible = False
        tbAddr1.Focus()
    End Sub
    
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>NHS Mailing List</title>
</head>
<body>
    <form id="Form1" runat="Server">
    <main:Header ID="ctlHeader" runat="server"></main:Header>
    <table style="width: 100%" cellpadding="0" cellspacing="0">
        <tr class="bar_reports">
            <td style="width: 50%; white-space: nowrap">
            </td>
            <td style="width: 50%; white-space: nowrap" align="right">
            </td>
        </tr>
    </table>
    <asp:Panel ID="pnlMain" runat="server" Font-Names="Verdana" Font-Size="X-Small" Width="100%">
        <table style="width: 100%">
            <tr>
                <td>
                    <asp:Label ID="Label1" runat="server" Font-Bold="True" Font-Names="Verdana" 
                        Font-Size="X-Small" >NHS Mailing List - Create Entry</asp:Label>
                    &nbsp;&nbsp;&nbsp;&nbsp;<asp:Label ID="Label19" runat="server" Font-Bold="True" Font-Italic="True" 
                        Font-Names="Verdana" Font-Size="XX-Small" Text="Quick start:" />
                    &nbsp;<asp:TextBox ID="tbQuickStart" runat="server" Font-Names="Verdana" 
                        Font-Size="X-Small" MaxLength="50" Width="494px" BackColor="#FFFFCC" />
                    &nbsp;<asp:Button ID="btnGoQuick" runat="server" Text="go" OnClick="btnGoQuick_Click" CausesValidation="False" />
                    <br />
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                    &nbsp;&nbsp;&nbsp;<asp:Label 
                        ID="Label20" runat="server" Font-Bold="True" Font-Italic="True" 
                        Font-Names="Verdana" Font-Size="XX-Small" 
                        
                        Text="([mr;] first name; last name; house number; post code; xyz RETURN)" />
                    &nbsp;&nbsp;
                    </td>
            </tr>
        </table>
        <table style="width: 100%">
            <tr id="trIdentity" runat="server" visible="true">
                <td style="width: 5%" />
                <td>
                    <asp:LinkButton ID="lnkbtnTitleMs" runat="server" Font-Names="Verdana" Font-Size="XX-Small" OnClick="lnkbtnTitleMs_Click" CausesValidation="False">Ms</asp:LinkButton>
                    <asp:LinkButton ID="lnkbtnTitleMrs" runat="server" Font-Names="Verdana" Font-Size="XX-Small" OnClick="lnkbtnTitleMrs_Click" CausesValidation="False">Mrs</asp:LinkButton>
                    <asp:LinkButton ID="lnkbtnTitleMr" runat="server" Font-Names="Verdana" Font-Size="XX-Small" OnClick="lnkbtnTitleMr_Click" CausesValidation="False">Mr</asp:LinkButton>
                    &nbsp;<asp:Label ID="Label3" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small" Text="Title:" />
                    &nbsp;<asp:TextBox ID="tbTitle" runat="server" Font-Names="Verdana" Font-Size="XX-Small" MaxLength="20" Width="80px" />
                    &nbsp;<asp:Label ID="Label4" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small" Text="First name:" />
                    &nbsp;<asp:TextBox ID="tbFirstName" runat="server" Font-Names="Verdana" Font-Size="XX-Small" MaxLength="50" Width="200px" />
                    &nbsp;<asp:Label ID="Label5" runat="server" Font-Bold="False" 
                        Font-Names="Verdana" Font-Size="XX-Small" Text="Last name:" ForeColor="Red" />
                    &nbsp;<asp:TextBox ID="tbLastName" runat="server" Font-Names="Verdana" Font-Size="XX-Small" MaxLength="50" Width="200px" />
                    &nbsp;<asp:RequiredFieldValidator ID="rfvLastName" runat="server" 
                        ControlToValidate="tbLastName" ErrorMessage="###" Font-Bold="True" 
                        Font-Names="Verdana" Font-Size="XX-Small" SetFocusOnError="True"/>
                    &nbsp;<asp:RadioButton ID="rblGenderMale" runat="server" Font-Names="Verdana" Font-Size="XX-Small" GroupName="gender" Text="Male" />
                    <asp:RadioButton ID="rblGenderFemale" runat="server" Font-Names="Verdana" Font-Size="XX-Small" GroupName="gender" Text="Female" />
                </td>
            </tr>
            <tr>
                <td />
                <td>
                    <asp:Label ID="lblLegendPostcode" runat="server" Font-Bold="True" 
                        Font-Names="Verdana" Font-Size="XX-Small" Text="Post code:" ForeColor="Red" />
                    &nbsp;<asp:TextBox ID="tbPostcode" runat="server" Font-Names="Verdana" Font-Size="XX-Small" MaxLength="50" Width="234px" Height="23px" />
                    &nbsp;<asp:RequiredFieldValidator ID="tfvPostcode" runat="server" ControlToValidate="tbPostcode" ErrorMessage="###" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small" SetFocusOnError="True" />
                    &nbsp;<asp:Button ID="btnFindAddress" runat="server" Text="find address (ALT + 1)" 
                        onclick="btnFindAddress_Click" CausesValidation="False" AccessKey="1" />
                    &nbsp;<asp:Button ID="btnBypassPostcode" runat="server" AccessKey="8" 
                        Text="BYPASS POSTCODE (ALT + 8)" Visible="False" 
                        OnClick="btnBypassPostcode_Click" CausesValidation="False" />
                    <asp:Label ID="lblLookupError" runat="server" Font-Names="Verdana" Font-Size="XX-Small" ForeColor="Red" Visible="False" />
                    &nbsp;<asp:Label ID="lblPCT" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small" />
                    &nbsp;<asp:LinkButton ID="lnkbtnResetPCTLookupSeed" runat="server" 
                        Font-Names="Verdana" Font-Size="XX-Small" 
                        onclick="lnkbtnResetPCTLookupSeed_Click" Visible="False" 
                        CausesValidation="False">reset PCT seed</asp:LinkButton>
                </td>
            </tr>
            <tr id="trPostcodeLookupResults" runat="server" visible="false">
                <td />
                <td valign="top">
                    <asp:Label ID="lblSelectADestination" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small" Text="Select a destination:" />
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:LinkButton 
                        ID="lnkbtnAddrLookupCancel" runat="server" Font-Names="Verdana" 
                        Font-Size="XX-Small" onclick="lnkbtnAddrLookupCancel_Click" AccessKey="9" 
                        CausesValidation="False">cancel (ALT + 9)</asp:LinkButton>
                    <br />
                    <asp:ListBox ID="lbLookupResults" runat="server" AutoPostBack="True" 
                        Font-Names="Verdana" Font-Size="XX-Small" Height="250px" 
                        OnSelectedIndexChanged="lbLookupResults_SelectedIndexChanged" Width="408px">
                    </asp:ListBox>
                    &nbsp;</td>
            </tr>
            <tr id="trAddress1" runat="server" visible="true">
                <td />
                <td>
                    <asp:Label ID="Label2" runat="server" Font-Bold="False" Font-Names="Verdana" 
                        Font-Size="XX-Small" Text="Address 1:" ForeColor="Red" />
                    &nbsp;<asp:TextBox ID="tbAddr1" runat="server" Font-Names="Verdana" Font-Size="XX-Small" MaxLength="50" Width="350px" />
                    &nbsp;<asp:RequiredFieldValidator ID="rfvAddr1" runat="server" ControlToValidate="tbAddr1" ErrorMessage="###" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small" SetFocusOnError="True"/>
                </td>
            </tr>
            <tr id="trAddress2" runat="server" visible="true">
                <td />
                <td>
                    <asp:Label ID="Label6" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small" Text="Address 2:" />
                    &nbsp;<asp:TextBox ID="tbAddr2" runat="server" Font-Names="Verdana" Font-Size="XX-Small" MaxLength="50" Width="350px" />
                    &nbsp;<asp:Label ID="Label8" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small" Text="Address 3:" />
                    &nbsp;<asp:TextBox ID="tbAddr3" runat="server" Font-Names="Verdana" Font-Size="XX-Small" MaxLength="50" Width="350px" />
                </td>
            </tr>
            <tr id="trTownCity" runat="server" visible="true">
                <td />
                <td>
                    <asp:Label ID="Label7" runat="server" Font-Bold="False" Font-Names="Verdana" 
                        Font-Size="XX-Small" Text="Town/City:" ForeColor="Red" />
                    &nbsp;<asp:TextBox ID="tbTown" runat="server" Font-Names="Verdana" 
                        Font-Size="XX-Small" MaxLength="50" Width="150px" />
                    &nbsp;<asp:RequiredFieldValidator ID="rfvTown" runat="server" ControlToValidate="tbTown" ErrorMessage="###" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small" SetFocusOnError="True"/>
                    &nbsp;<asp:Label ID="Label15" runat="server" Font-Bold="False" Font-Names="Verdana" 
                        Font-Size="XX-Small" Text="County:" />
                    &nbsp;<asp:TextBox ID="tbCounty" runat="server" Font-Names="Verdana" 
                        Font-Size="XX-Small" MaxLength="50" Width="150px" />
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                    <asp:Label ID="Label13" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small" Text="Contact:" />
                    &nbsp;<asp:RadioButton ID="rblContactMethodOrderForm" runat="server" 
                        Font-Names="Verdana" Font-Size="XX-Small" GroupName="contact" 
                        Text="order form &lt;b&gt;ord&lt;/b&gt;" Checked="True" />
                    <asp:RadioButton ID="rblContactMethodTelephone" runat="server" 
                        Font-Names="Verdana" Font-Size="XX-Small" GroupName="contact" 
                        Text="telephone &lt;b&gt;tel&lt;/b&gt;" />
                </td>
            </tr>
            <tr>
                <td />
                <td>
                </td>
            </tr>
            <tr id="trProducts" runat="server" visible="true">
                <td />
                <td>
                    <table style="width:100%">
                        <tr>
                            <td style="width:266px">
                                <asp:Label ID="Label17" runat="server" Font-Bold="True" Font-Names="Verdana" 
                                    Font-Size="XX-Small" Text="SUMMARY CARE RECORD LEAFLET" />
                            </td>
                            <td>
                                <asp:Label ID="Label18" runat="server" Font-Bold="True" Font-Names="Verdana" 
                                    Font-Size="XX-Small" 
                                    Text="NHS CARE RECORD GUARANTEE / CONFIDENTIALITY LEAFLETS" />
                            </td>
                        </tr>
                        <tr>
                            <td>
                                &nbsp;
                            </td>
                            <td>
                                <asp:CheckBox ID="cb3717_3984English" runat="server" Font-Names="Verdana" 
                                    Font-Size="XX-Small" Text="3717/3984 English  &lt;b&gt;3en&lt;/b&gt;" />
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <asp:CheckBox ID="cb4269EnglishBraille" runat="server" Font-Names="Verdana" 
                                    Font-Size="XX-Small" Text="4269 English Braille  &lt;b&gt;4br&lt;/b&gt;" />
                            </td>
                            <td>
                                <asp:CheckBox ID="cb3717_3984EnglishBraille" runat="server" 
                                    Font-Names="Verdana" Font-Size="XX-Small" 
                                    Text="3717/3984 English Braille &lt;b&gt; 3br&lt;/b&gt;" />
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <asp:CheckBox ID="cb4269EnglishLP" runat="server" Font-Names="Verdana" 
                                    Font-Size="XX-Small" 
                                    Text="4269 English Large Print &lt;b&gt; 4lp&lt;/b&gt;" />
                            </td>
                            <td>
                                <asp:CheckBox ID="cb3717_3984EnglishLP" runat="server" Font-Names="Verdana" 
                                    Font-Size="XX-Small" 
                                    Text="3717/3984 English Large Print  &lt;b&gt;3lp&lt;/b&gt;" />
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <asp:CheckBox ID="cb4269EasyRead" runat="server" Font-Names="Verdana" 
                                    Font-Size="XX-Small" Text="4269 Easy Read  &lt;b&gt;4er&lt;/b&gt;" />
                            </td>
                            <td>
                                <asp:CheckBox ID="cb3717_3984EasyRead" runat="server" Font-Names="Verdana" 
                                    Font-Size="XX-Small" Text="3717/3984 Easy Read  &lt;b&gt;3er&lt;/b&gt;" />
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <asp:CheckBox ID="cb4269EnglishAudio" runat="server" Font-Names="Verdana" 
                                    Font-Size="XX-Small" 
                                    Text="4269 English Audio CD  &lt;b&gt;4au&lt;/b&gt;" />
                            </td>
                            <td>
                                <asp:CheckBox ID="cb3717_3984EnglishLAudio" runat="server" Font-Names="Verdana" 
                                    Font-Size="XX-Small" 
                                    Text="3717/3984 English Audio CD  &lt;b&gt;3au&lt;/b&gt;" />
                            </td>
                        </tr>
                        <tr>
                            <td style="width:266px">
                                <asp:CheckBox ID="cb4269Gujarati" runat="server" Font-Names="Verdana" 
                                    Font-Size="XX-Small" Text="4269 Gujarati  &lt;b&gt;4gu&lt;/b&gt;" />
                            </td>
                            <td>
                                <asp:CheckBox ID="cb3717_3984Gujarati" runat="server" Font-Names="Verdana" 
                                    Font-Size="XX-Small" Text="3717/3984 Gujarati  &lt;b&gt;3gu&lt;/b&gt;" />
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <asp:CheckBox ID="cb4269Urdu" runat="server" Font-Names="Verdana" 
                                    Font-Size="XX-Small" Text="4269 Urdu  &lt;b&gt;4ur&lt;/b&gt;" />
                            </td>
                            <td>
                                <asp:CheckBox ID="cb3717_3984Urdu" runat="server" Font-Names="Verdana" 
                                    Font-Size="XX-Small" Text="3717/3984 Urdu  &lt;b&gt;3ur&lt;/b&gt;" />
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <asp:CheckBox ID="cb4269Mandarin" runat="server" Font-Names="Verdana" 
                                    Font-Size="XX-Small" Text="4269 Mandarin  &lt;b&gt;4ma&lt;/b&gt;" />
                            </td>
                            <td>
                                <asp:CheckBox ID="cb3717_3984Mandarin" runat="server" Font-Names="Verdana" 
                                    Font-Size="XX-Small" Text="3717/3984 Mandarin  &lt;b&gt;3ma&lt;/b&gt;" />
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <asp:CheckBox ID="cb4269Polish" runat="server" Font-Names="Verdana" 
                                    Font-Size="XX-Small" Text="4269 Polish  &lt;b&gt;4po&lt;/b&gt;" />
                            </td>
                            <td>
                                <asp:CheckBox ID="cb3717_3984Polish" runat="server" Font-Names="Verdana" 
                                    Font-Size="XX-Small" Text="3717/3984 Polish  &lt;b&gt;3po&lt;/b&gt;" />
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <asp:CheckBox ID="cb4269French" runat="server" Font-Names="Verdana" 
                                    Font-Size="XX-Small" Text="4269 French  &lt;b&gt;4fr&lt;/b&gt;" />
                            </td>
                            <td>
                                <asp:CheckBox ID="cb3717_3984French" runat="server" Font-Names="Verdana" 
                                    Font-Size="XX-Small" Text="3717/3984 French  &lt;b&gt;3fr&lt;/b&gt;" />
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <asp:CheckBox ID="cb4269Farsi" runat="server" Font-Names="Verdana" 
                                    Font-Size="XX-Small" Text="4269 Farsi  &lt;b&gt;4fa&lt;/b&gt;" />
                            </td>
                            <td>
                                <asp:CheckBox ID="cb3717_3984Farsi" runat="server" Font-Names="Verdana" 
                                    Font-Size="XX-Small" Text="3717/3984 Farsi  &lt;b&gt;3fa&lt;/b&gt;" />
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <asp:CheckBox ID="cb4269Spanish" runat="server" Font-Names="Verdana" 
                                    Font-Size="XX-Small" Text="4269 Spanish  &lt;b&gt;4sp&lt;/b&gt;" />
                            </td>
                            <td>
                                <asp:CheckBox ID="cb3717_3984Spanish" runat="server" Font-Names="Verdana" 
                                    Font-Size="XX-Small" Text="3717/3984 Spanish  &lt;b&gt;3sp&lt;/b&gt;" />
                            </td>
                        </tr>
                        <tr>
                            <td colspan="2">
                                <asp:CheckBox ID="cb3716CRSEasyRead" runat="server" Font-Names="Verdana" 
                                    Font-Size="XX-Small" 
                                    
                                    Text="Easy read picture vsn (ref 3716 NHS Care Recs Service)  &lt;b&gt;crs&lt;/b&gt;" />
                            </td>
                            <tr>
                                <td colspan="2">
                                    <asp:CheckBox ID="cbOptOut" runat="server" Font-Names="Verdana" 
                                        Font-Size="XX-Small" 
                                        
                                        
                                        Text="Please send me information about what to do if I do not want a summary card record created for me.  &lt;b&gt;opt&lt;/b&gt;" />
                                </td>
                            </tr>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr id="trPCT" runat="server" visible="true">
                <td />
                <td>
                    <asp:Label ID="Label12" runat="server" Font-Bold="False" Font-Names="Verdana" 
                        Font-Size="XX-Small" Text="PCT:" ForeColor="Red" />
                    &nbsp;<asp:DropDownList ID="ddlPCT" runat="server" Font-Names="Verdana" Font-Size="XX-Small">
                        <asp:ListItem Selected="True" Value="XXX">- please select -</asp:ListItem>
                        <asp:ListItem Value="5HG">Ashton, Leigh and Wigan PCT</asp:ListItem>
                        <asp:ListItem Value="5C2">Barking and Dagenham PCT</asp:ListItem>
                        <asp:ListItem Value="5A9">Barnet PCT</asp:ListItem>
                        <asp:ListItem Value="5JE">Barnsley PCT</asp:ListItem>
                        <asp:ListItem Value="5ET">Bassetlaw PCT</asp:ListItem>
                        <asp:ListItem Value="5FL">Bath and North East Somerset PCT</asp:ListItem>
                        <asp:ListItem Value="5P2">Bedfordshire PCT</asp:ListItem>
                        <asp:ListItem Value="5QG">Berkshire East PCT</asp:ListItem>
                        <asp:ListItem Value="5QF">Berkshire West PCT</asp:ListItem>
                        <asp:ListItem Value="TAK">Bexley PCT</asp:ListItem>
                        <asp:ListItem Value="5PG">Birmingham East and North PCT</asp:ListItem>
                        <asp:ListItem Value="5CC">Blackburn With Darwen PCT</asp:ListItem>
                        <asp:ListItem Value="5HP">Blackpool PCT</asp:ListItem>
                        <asp:ListItem Value="5HQ">Bolton PCT</asp:ListItem>
                        <asp:ListItem Value="5QN">Bournemouth and Poole Teaching PCT</asp:ListItem>
                        <asp:ListItem Value="5NY">Bradford and Airedale Teaching PCT</asp:ListItem>
                        <asp:ListItem Value="5K5">Brent Teaching PCT</asp:ListItem>
                        <asp:ListItem Value="5LQ">Brighton and Hove City PCT</asp:ListItem>
                        <asp:ListItem Value="5QJ">Bristol PCT</asp:ListItem>
                        <asp:ListItem Value="5A7">Bromley PCT</asp:ListItem>
                        <asp:ListItem Value="5QD">Buckinghamshire PCT</asp:ListItem>
                        <asp:ListItem Value="5JX">Bury PCT</asp:ListItem>
                        <asp:ListItem Value="5J6">Calderdale PCT</asp:ListItem>
                        <asp:ListItem Value="5PP">Cambridgeshire PCT</asp:ListItem>
                        <asp:ListItem Value="5K7">Camden PCT</asp:ListItem>
                        <asp:ListItem Value="5NP">Central and Eastern Cheshire PCT</asp:ListItem>
                        <asp:ListItem Value="5NG">Central Lancashire PCT</asp:ListItem>
                        <asp:ListItem Value="5C3">City and Hackney Teaching PCT</asp:ListItem>
                        <asp:ListItem Value="5QP">Cornwall and Isles Of Scilly PCT</asp:ListItem>
                        <asp:ListItem Value="5ND">County Durham PCT</asp:ListItem>
                        <asp:ListItem Value="5MD">Coventry Teaching PCT</asp:ListItem>
                        <asp:ListItem Value="5K9">Croydon PCT</asp:ListItem>
                        <asp:ListItem Value="5NE">Cumbria Teaching PCT</asp:ListItem>
                        <asp:ListItem Value="5J9">Darlington PCT</asp:ListItem>
                        <asp:ListItem Value="5N7">Derby City PCT</asp:ListItem>
                        <asp:ListItem Value="5N6">Derbyshire County PCT</asp:ListItem>
                        <asp:ListItem Value="5QQ">Devon PCT</asp:ListItem>
                        <asp:ListItem Value="5N5">Doncaster PCT</asp:ListItem>
                        <asp:ListItem Value="5QM">Dorset PCT</asp:ListItem>
                        <asp:ListItem Value="5PE">Dudley PCT</asp:ListItem>
                        <asp:ListItem Value="5HX">Ealing PCT</asp:ListItem>
                        <asp:ListItem Value="5P3">East and North Hertfordshire PCT</asp:ListItem>
                        <asp:ListItem Value="5NH">East Lancashire Teaching PCT</asp:ListItem>
                        <asp:ListItem Value="5NW">East Riding Of Yorkshire PCT</asp:ListItem>
                        <asp:ListItem Value="5P7">East Sussex Downs and Weald PCT</asp:ListItem>
                        <asp:ListItem Value="5QA">Eastern and Coastal Kent PCT</asp:ListItem>
                        <asp:ListItem Value="5C1">Enfield PCT</asp:ListItem>
                        <asp:ListItem Value="5KF">Gateshead PCT</asp:ListItem>
                        <asp:ListItem Value="5QH">Gloucestershire PCT</asp:ListItem>
                        <asp:ListItem Value="5PR">Great Yarmouth and Waveney PCT</asp:ListItem>
                        <asp:ListItem Value="5A8">Greenwich Teaching PCT</asp:ListItem>
                        <asp:ListItem Value="5NM">Halton and St Helens PCT</asp:ListItem>
                        <asp:ListItem Value="5H1">Hammersmith and Fulham PCT</asp:ListItem>
                        <asp:ListItem Value="5QC">Hampshire PCT</asp:ListItem>
                        <asp:ListItem Value="5C9">Haringey Teaching PCT</asp:ListItem>
                        <asp:ListItem Value="5K6">Harrow PCT</asp:ListItem>
                        <asp:ListItem Value="5D9">Hartlepool PCT</asp:ListItem>
                        <asp:ListItem Value="5P8">Hastings and Rother PCT</asp:ListItem>
                        <asp:ListItem Value="5A4">Havering PCT</asp:ListItem>
                        <asp:ListItem Value="5MX">Heart Of Birmingham Teaching PCT</asp:ListItem>
                        <asp:ListItem Value="5CN">Herefordshire PCT</asp:ListItem>
                        <asp:ListItem Value="5NQ">Heywood, Middleton and Rochdale PCT</asp:ListItem>
                        <asp:ListItem Value="5AT">Hillingdon PCT</asp:ListItem>
                        <asp:ListItem Value="5HY">Hounslow PCT</asp:ListItem>
                        <asp:ListItem Value="5NX">Hull Teaching PCT</asp:ListItem>
                        <asp:ListItem Value="5QT">Isle Of Wight NHS PCT</asp:ListItem>
                        <asp:ListItem Value="5K8">Islington PCT</asp:ListItem>
                        <asp:ListItem Value="5LA">Kensington and Chelsea PCT</asp:ListItem>
                        <asp:ListItem Value="5A5">Kingston PCT</asp:ListItem>
                        <asp:ListItem Value="5N2">Kirklees PCT</asp:ListItem>
                        <asp:ListItem Value="5J4">Knowsley PCT</asp:ListItem>
                        <asp:ListItem Value="5LD">Lambeth PCT</asp:ListItem>
                        <asp:ListItem Value="5N1">Leeds PCT</asp:ListItem>
                        <asp:ListItem Value="5PC">Leicester City PCT</asp:ListItem>
                        <asp:ListItem Value="5PA">Leicestershire County and Rutland PCT</asp:ListItem>
                        <asp:ListItem Value="5LF">Lewisham PCT</asp:ListItem>
                        <asp:ListItem Value="5N9">Lincolnshire Teaching PCT</asp:ListItem>
                        <asp:ListItem Value="5NL">Liverpool PCT</asp:ListItem>
                        <asp:ListItem Value="5GC">Luton PCT</asp:ListItem>
                        <asp:ListItem Value="5NT">Manchester PCT</asp:ListItem>
                        <asp:ListItem Value="5L3">Medway PCT</asp:ListItem>
                        <asp:ListItem Value="5PX">Mid Essex PCT</asp:ListItem>
                        <asp:ListItem Value="5KM">Middlesbrough PCT</asp:ListItem>
                        <asp:ListItem Value="5CQ">Milton Keynes PCT</asp:ListItem>
                        <asp:ListItem Value="5D7">Newcastle PCT</asp:ListItem>
                        <asp:ListItem Value="5C5">Newham PCT</asp:ListItem>
                        <asp:ListItem Value="5PQ">Norfolk PCT</asp:ListItem>
                        <asp:ListItem Value="5PW">North East Essex PCT</asp:ListItem>
                        <asp:ListItem Value="5NF">North Lancashire Teaching PCT</asp:ListItem>
                        <asp:ListItem Value="5EF">North Lincolnshire PCT</asp:ListItem>
                        <asp:ListItem Value="5M8">North Somerset PCT</asp:ListItem>
                        <asp:ListItem Value="5PH">North Staffordshire PCT</asp:ListItem>
                        <asp:ListItem Value="5D8">North Tyneside PCT</asp:ListItem>
                        <asp:ListItem Value="5NV">North Yorkshire and York PCT</asp:ListItem>
                        <asp:ListItem Value="5PD">Northamptonshire Teaching PCT</asp:ListItem>
                        <asp:ListItem Value="TAC">Northumberland Care Trust</asp:ListItem>
                        <asp:ListItem Value="5EM">Nottingham City PCT</asp:ListItem>
                        <asp:ListItem Value="5N8">Nottinghamshire County Teaching PCT</asp:ListItem>
                        <asp:ListItem Value="5J5">Oldham PCT</asp:ListItem>
                        <asp:ListItem Value="5QE">Oxfordshire PCT</asp:ListItem>
                        <asp:ListItem Value="5PN">Peterborough PCT</asp:ListItem>
                        <asp:ListItem Value="5F1">Plymouth Teaching PCT</asp:ListItem>
                        <asp:ListItem Value="5FE">Portsmouth City Teaching PCT</asp:ListItem>
                        <asp:ListItem Value="5NA">Redbridge PCT</asp:ListItem>
                        <asp:ListItem Value="5QR">Redcar and Cleveland PCT</asp:ListItem>
                        <asp:ListItem Value="5M6">Richmond and Twickenham PCT</asp:ListItem>
                        <asp:ListItem Value="5H8">Rotherham PCT</asp:ListItem>
                        <asp:ListItem Value="5F5">Salford PCT</asp:ListItem>
                        <asp:ListItem Value="5PF">Sandwell PCT</asp:ListItem>
                        <asp:ListItem Value="5NJ">Sefton PCT</asp:ListItem>
                        <asp:ListItem Value="5N4">Sheffield PCT</asp:ListItem>
                        <asp:ListItem Value="5M2">Shropshire County PCT</asp:ListItem>
                        <asp:ListItem Value="5QL">Somerset PCT</asp:ListItem>
                        <asp:ListItem Value="5M1">South Birmingham PCT</asp:ListItem>
                        <asp:ListItem Value="5P1">South East Essex PCT</asp:ListItem>
                        <asp:ListItem Value="5A3">South Gloucestershire PCT</asp:ListItem>
                        <asp:ListItem Value="5PK">South Staffordshire PCT</asp:ListItem>
                        <asp:ListItem Value="5KG">South Tyneside PCT</asp:ListItem>
                        <asp:ListItem Value="5PY">South West Essex PCT</asp:ListItem>
                        <asp:ListItem Value="5L1">Southampton City PCT</asp:ListItem>
                        <asp:ListItem Value="5LE">Southwark PCT</asp:ListItem>
                        <asp:ListItem Value="5F7">Stockport PCT</asp:ListItem>
                        <asp:ListItem Value="5E1">Stockton-on-Tees Teaching PCT</asp:ListItem>
                        <asp:ListItem Value="5PJ">Stoke On Trent PCT</asp:ListItem>
                        <asp:ListItem Value="5PT">Suffolk PCT</asp:ListItem>
                        <asp:ListItem Value="5KL">Sunderland Teaching PCT</asp:ListItem>
                        <asp:ListItem Value="5P5">Surrey PCT</asp:ListItem>
                        <asp:ListItem Value="5M7">Sutton and Merton PCT</asp:ListItem>
                        <asp:ListItem Value="5K3">Swindon PCT</asp:ListItem>
                        <asp:ListItem Value="5LH">Tameside and Glossop PCT</asp:ListItem>
                        <asp:ListItem Value="5MK">Telford and Wrekin PCT</asp:ListItem>
                        <asp:ListItem Value="5C4">Tower Hamlets PCT</asp:ListItem>
                        <asp:ListItem Value="5NR">Trafford PCT</asp:ListItem>
                        <asp:ListItem Value="5N3">Wakefield District PCT</asp:ListItem>
                        <asp:ListItem Value="5M3">Walsall Teaching PCT</asp:ListItem>
                        <asp:ListItem Value="5NC">Waltham Forest PCT</asp:ListItem>
                        <asp:ListItem Value="5LG">Wandsworth PCT</asp:ListItem>
                        <asp:ListItem Value="5J2">Warrington PCT</asp:ListItem>
                        <asp:ListItem Value="5PM">Warwickshire PCT</asp:ListItem>
                        <asp:ListItem Value="5PV">West Essex PCT</asp:ListItem>
                        <asp:ListItem Value="5P4">West Hertfordshire PCT</asp:ListItem>
                        <asp:ListItem Value="5P9">West Kent PCT</asp:ListItem>
                        <asp:ListItem Value="5P6">West Sussex PCT</asp:ListItem>
                        <asp:ListItem Value="5NN">Western Cheshire PCT</asp:ListItem>
                        <asp:ListItem Value="5LC">Westminster PCT</asp:ListItem>
                        <asp:ListItem Value="5QK">Wiltshire PCT</asp:ListItem>
                        <asp:ListItem Value="5NK">Wirral PCT</asp:ListItem>
                        <asp:ListItem Value="5MV">Wolverhampton City PCT</asp:ListItem>
                        <asp:ListItem Value="5PL">Worcestershire PCT</asp:ListItem>
                    </asp:DropDownList>
                    &nbsp;<asp:RequiredFieldValidator ID="rfvPCT" runat="server" 
                        ControlToValidate="ddlPCT" ErrorMessage="###" Font-Bold="True" 
                        Font-Names="Verdana" Font-Size="XX-Small" InitialValue="XXX" 
                        SetFocusOnError="True" />
                    &nbsp;&nbsp;&nbsp;
                    <asp:HyperLink ID="HyperLink1" runat="server" Font-Names="Verdana" 
                        Font-Size="XX-Small" 
                        NavigateUrl="https://www.ndtms.org.uk/emids/cgi-bin/ons_locale.cgi" 
                        Target="_blank">PCT lookup web page</asp:HyperLink>
                    </td>
            </tr>
            <tr ID="trControls" runat="server" visible="true">
                <td />
                <td>
                    <asp:Button ID="btnSave" runat="server" AccessKey="0" EnableTheming="False" 
                        OnClick="btnSave_Click" Text="save (ALT + 0)" Width="219px" />
                    &nbsp;&nbsp;<asp:CheckBox ID="cbBypassPostcodeValidation" runat="server" 
                        Font-Names="Verdana" Font-Size="XX-Small" Text="bypass postcode validation (ALT + 2)" 
                        Visible="False" AccessKey="2" />
                    &nbsp;&nbsp;<asp:LinkButton ID="lnkbtnClearForm" runat="server" 
                        CausesValidation="False" Font-Names="Verdana" Font-Size="XX-Small" 
                        OnClick="lnkbtnClearForm_Click" AccessKey="3">clear form (ALT + 3)</asp:LinkButton>
                    &nbsp;&nbsp;
                    <asp:Label ID="lblStats" runat="server" Font-Bold="True" Font-Names="Verdana" 
                        Font-Size="XX-Small" ForeColor="#CC9900" Text="Entered Today:" />
                </td>
            </tr>
            <tr>
                <td />
                <td>
                    <asp:Label ID="lblMessage" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="X-Small" ForeColor="#006600" />
                </td>
            </tr>
                <tr>
                    <td />
                        <td>
                            <asp:LinkButton ID="lnkbtnShowMostRecentEntries" runat="server" 
                                CausesValidation="False" Font-Names="Verdana" Font-Size="XX-Small" 
                                onclick="lnkbtnShowMostRecentEntries_Click">show most recent entries</asp:LinkButton>
                            &nbsp;<asp:LinkButton ID="lnkbtnRefresh" runat="server" Font-Names="Verdana" 
                                Font-Size="XX-Small" onclick="lnkbtnRefresh_Click" Visible="False" 
                                CausesValidation="False">refresh</asp:LinkButton>
                            &nbsp;<asp:LinkButton ID="lnkbtnHideMostRecentEntries" runat="server" 
                                Font-Names="Verdana" Font-Size="XX-Small" Visible="False" 
                                onclick="lnkbtnHideMostRecentEntries_Click" CausesValidation="False">hide most recent 
                            entries</asp:LinkButton>
                        </td>
                    </tr>
                    <tr>
                        <td />
                            <td>
                                &nbsp;</td>
                        </tr>
        </table>
    </asp:Panel>
    <asp:Panel ID="pnlRecentEntries" runat="server" HorizontalAlign="Center">
                <asp:GridView ID="gvRecentEntries" runat="server" CellPadding="2" Font-Names="Verdana" Font-Size="XX-Small" Width="99%" >
                </asp:GridView>
    </asp:Panel>
    <asp:Panel ID="pnlPreamble" runat="server" Font-Names="Verdana" Font-Size="X-Small" Width="100%">
        <hr />
        <table style="width: 100%">
            <tr>
                <td style="width: 50%">
                    <asp:Label ID="lblScreenTitle" runat="server" Font-Bold="True" 
                        Font-Names="Verdana" Font-Size="X-Small" 
                        Text="NHS Mailing List - Export Data" />
                </td>
                <td style="width: 50%" align="right">
                    &nbsp;&nbsp;
                </td>
            </tr>
        </table>
        <table style="width: 100%">
            <tr>
                <td style="width: 5%">
                </td>
                <td style="width: 95%">
                    <asp:Label ID="Label14" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small" Text="Export" />
                    <asp:RadioButton ID="rblExportEntriesSince" runat="server" Font-Names="Verdana" Font-Size="XX-Small" GroupName="export" Text="entries since..." AutoPostBack="True" OnCheckedChanged="rblExportEntriesSince_CheckedChanged" />
                    <asp:TextBox ID="tbDateSince" runat="server" Font-Names="Verdana" Font-Size="XX-Small" MaxLength="20" Width="80px" />
                    &nbsp;<asp:LinkButton ID="lnkbtnLastSunday" runat="server" Font-Names="Verdana" Font-Size="XX-Small" OnClick="lnkbtnLastSunday_Click" CausesValidation="False">last 
                    weekend</asp:LinkButton>&nbsp;
                    <asp:RadioButton ID="rblExportEntriesBetween" runat="server" Font-Names="Verdana" Font-Size="XX-Small" GroupName="export" Text="entries between..." AutoPostBack="True" OnCheckedChanged="rblExportEntriesBetween_CheckedChanged" />
                    <asp:TextBox ID="tbDateFrom" runat="server" Font-Names="Verdana" Font-Size="XX-Small" MaxLength="20" Width="80px" />&nbsp;<asp:Label ID="lblAndDate" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small" Text="and" />
                    &nbsp;<asp:TextBox ID="tbDateTo" runat="server" Font-Names="Verdana" Font-Size="XX-Small" MaxLength="20" Width="80px" />
                    &nbsp;<asp:Label ID="Label16" runat="server" Font-Bold="False" Font-Names="Verdana" 
                        Font-Size="XX-Small" ForeColor="Gray" Text="(eg 16-Dec-2009)" />
                    &nbsp;<asp:RadioButton ID="rblExportAllEntries" runat="server" AutoPostBack="True" 
                        Font-Names="Verdana" Font-Size="XX-Small" GroupName="export" 
                        OnCheckedChanged="rblExportAllEntries_CheckedChanged" Text="all entries" />
                    &nbsp;<asp:Button ID="btnExport" runat="server" Text="go" Enabled="False" 
                        CausesValidation="False" OnClick="btnExport_Click" />
                </td>
            </tr>
        </table>
    </asp:Panel>
    </form>
</body>
</html>
