<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>
<%@ import Namespace="System.Collections.Generic" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

    ' NOTE: AllocateManagementFee, which divided the management fee equally between each category, has been superseded by the system that divides it according to the amount of business done by each category
    
    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    Private goStorageDataTable As DataTable = New DataTable()
    Private goReceptionDataTable As DataTable = New DataTable()
    Dim gsMonthNames() As String = {"", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"}
    
    Const RECORD_TYPE_STORAGE_COUNT As Integer = 1
    Const RECORD_TYPE_MANAGEMENT_FEE As Integer = 2
    Const RECORD_TYPE_PICK_CHARGES As Integer = 3
    Const RECORD_TYPE_SHIPPING_CHARGE As Integer = 4                          ' not used here but used in batch job
    Const RECORD_TYPE_RECEPTION_COUNT As Integer = 5
    
    Const CUSTOMER_KEY_KODDFIS As Integer = 541
    Const CUSTOMER_KEY As Integer = CUSTOMER_KEY_KODDFIS

    Protected Sub Page_Load(ByVal Source As Object, ByVal E As EventArgs)
        If Not IsNumeric(Session("CustomerKey")) Then
            Server.Transfer("session_expired.aspx")
            Exit Sub
        End If
        If Not Page.IsPostBack Then
            Call CreateDataTables()
            
            gvStorage.DataSource = goStorageDataTable
            Call GetMostRecentStorageData()
            gvStorage.EditIndex = 1
            gvStorage.DataBind()
            gvStorage.Rows(1).RowState = DataControlRowState.Edit

            gvReception.DataSource = goReceptionDataTable
            Call GetMostRecentReceptionData()
            gvReception.EditIndex = 1
            gvReception.DataBind()
            gvReception.Rows(1).RowState = DataControlRowState.Edit
        End If
    End Sub
    
    Protected Sub SetTitle()
        Dim sTitle As String = Session("SiteTitle")
        If sTitle <> String.Empty Then
            sTitle += " - "
        End If
        Page.Header.Title = sTitle & "Kodak Monthly Pallet Storage & Reception Record"
    End Sub
    
    Protected Sub CreateDataTables()
        If IsNothing(ViewState("KS_StorageDataTable")) Then
            goStorageDataTable = New DataTable()
            'goStorageDataTable.Columns.Add(New DataColumn("Accessories", GetType(Double)))
            goStorageDataTable.Columns.Add(New DataColumn("DigitalCameras", GetType(Double)))
            goStorageDataTable.Columns.Add(New DataColumn("Frames", GetType(Double)))
            goStorageDataTable.Columns.Add(New DataColumn("PocketVideoCameras", GetType(Double)))
            goStorageDataTable.Columns.Add(New DataColumn("FieldMerchandising", GetType(Double)))
            goStorageDataTable.Columns.Add(New DataColumn("FilmSUC", GetType(Double)))
            goStorageDataTable.Columns.Add(New DataColumn("Inkjet", GetType(Double)))
            goStorageDataTable.Columns.Add(New DataColumn("KioskDryLab", GetType(Double)))
            goStorageDataTable.Columns.Add(New DataColumn("KodakExpress", GetType(Double)))
            goStorageDataTable.Rows.Add()
            goStorageDataTable.Rows.Add()
            ViewState("KS_StorageDataTable") = goStorageDataTable
        End If

        If IsNothing(ViewState("KS_ReceptionDataTable")) Then
            goReceptionDataTable = New DataTable()
            'goReceptionDataTable.Columns.Add(New DataColumn("Accessories", GetType(Double)))
            goReceptionDataTable.Columns.Add(New DataColumn("DigitalCameras", GetType(Double)))
            goReceptionDataTable.Columns.Add(New DataColumn("Frames", GetType(Double)))
            goReceptionDataTable.Columns.Add(New DataColumn("PocketVideoCameras", GetType(Double)))
            goReceptionDataTable.Columns.Add(New DataColumn("FieldMerchandising", GetType(Double)))
            goReceptionDataTable.Columns.Add(New DataColumn("FilmSUC", GetType(Double)))
            goReceptionDataTable.Columns.Add(New DataColumn("Inkjet", GetType(Double)))
            goReceptionDataTable.Columns.Add(New DataColumn("KioskDryLab", GetType(Double)))
            goReceptionDataTable.Columns.Add(New DataColumn("KodakExpress", GetType(Double)))
            goReceptionDataTable.Rows.Add()
            goReceptionDataTable.Rows.Add()
            ViewState("KS_ReceptionDataTable") = goReceptionDataTable
        End If
    End Sub
        
    Protected Function bRecordExists(ByVal sMonth As Integer, ByVal sYear As Integer, ByVal nRecordType As Integer) As Boolean
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataReader As SqlDataReader = Nothing
        Dim sSQL As String = "SELECT * FROM ClientData_KODDFIS_AllocatedCharges WHERE Month = " & sMonth & " AND Year = " & sYear & " AND RecordType = " & nRecordType
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        bRecordExists = False
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            If oDataReader.HasRows Then
                bRecordExists = True
            End If
        Catch ex As Exception
            WebMsgBox.Show("Error in bRecordExists: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Function
    
    Protected Sub GetMostRecentStorageData()
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataReader As SqlDataReader = Nothing
        Dim sSQL As String = "SELECT TOP 1 * FROM ClientData_KODDFIS_AllocatedCharges WHERE RecordType = " & RECORD_TYPE_STORAGE_COUNT & " ORDER BY Year, Month DESC"
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            If oDataReader.HasRows Then
                oDataReader.Read()
                'oDataTable.Rows(0).Item("Accessories") = oDataReader("Accessories")
                goStorageDataTable.Rows(0).Item("DigitalCameras") = oDataReader("DigitalCameras")
                goStorageDataTable.Rows(0).Item("Frames") = oDataReader("Frames")
                goStorageDataTable.Rows(0).Item("PocketVideoCameras") = oDataReader("PocketVideoCameras")
                goStorageDataTable.Rows(0).Item("FieldMerchandising") = oDataReader("FieldMerchandising")
                goStorageDataTable.Rows(0).Item("FilmSUC") = oDataReader("FilmSUC")
                goStorageDataTable.Rows(0).Item("Inkjet") = oDataReader("Inkjet")
                goStorageDataTable.Rows(0).Item("KioskDryLab") = oDataReader("KioskDryLab")
                goStorageDataTable.Rows(0).Item("KodakExpress") = oDataReader("KodakExpress")
                Dim nYear As Integer = oDataReader("Year")
                Dim nMonth As Integer = oDataReader("Month")
                lblLastDataEntered.Text = "Last data entered was for " & gsMonthNames(nMonth) & " " & nYear.ToString
                If nMonth = 12 Then
                    nMonth = 0
                    nYear += 1
                End If
                nMonth += 1
                For i As Integer = 1 To ddlMonth.Items.Count - 1
                    If ddlMonth.Items(i).Value = nMonth Then
                        ddlMonth.SelectedIndex = i
                        Exit For
                    End If
                Next
                For i As Integer = 1 To ddlYear.Items.Count - 1
                    If ddlYear.Items(i).Value = nYear Then
                        ddlYear.SelectedIndex = i
                        Exit For
                    End If
                Next
                gvStorage.Focus()
            Else
                lblLastDataEntered.Text = "No storage data has been entered previously"
                ddlMonth.Focus()
            End If
        Catch ex As Exception
            WebMsgBox.Show("Error in GetMostRecentStorageData: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub GetMostRecentReceptionData()
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataReader As SqlDataReader = Nothing
        Dim sSQL As String = "SELECT TOP 1 * FROM ClientData_KODDFIS_AllocatedCharges WHERE RecordType = " & RECORD_TYPE_RECEPTION_COUNT & " ORDER BY Year, Month DESC"
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            If oDataReader.HasRows Then
                oDataReader.Read()
                'oDataTable.Rows(0).Item("Accessories") = oDataReader("Accessories")
                goReceptionDataTable.Rows(0).Item("DigitalCameras") = oDataReader("DigitalCameras")
                goReceptionDataTable.Rows(0).Item("Frames") = oDataReader("Frames")
                goReceptionDataTable.Rows(0).Item("PocketVideoCameras") = oDataReader("PocketVideoCameras")
                goReceptionDataTable.Rows(0).Item("FieldMerchandising") = oDataReader("FieldMerchandising")
                goReceptionDataTable.Rows(0).Item("FilmSUC") = oDataReader("FilmSUC")
                goReceptionDataTable.Rows(0).Item("Inkjet") = oDataReader("Inkjet")
                goReceptionDataTable.Rows(0).Item("KioskDryLab") = oDataReader("KioskDryLab")
                goReceptionDataTable.Rows(0).Item("KodakExpress") = oDataReader("KodakExpress")
                Dim nYear As Integer = oDataReader("Year")
                Dim nMonth As Integer = oDataReader("Month")
                lblLastDataEntered.Text = "Last data entered was for " & gsMonthNames(nMonth) & " " & nYear.ToString
                If nMonth = 12 Then
                    nMonth = 0
                    nYear += 1
                End If
                nMonth += 1
                For i As Integer = 1 To ddlMonth.Items.Count - 1
                    If ddlMonth.Items(i).Value = nMonth Then
                        ddlMonth.SelectedIndex = i
                        Exit For
                    End If
                Next
                For i As Integer = 1 To ddlYear.Items.Count - 1
                    If ddlYear.Items(i).Value = nYear Then
                        ddlYear.SelectedIndex = i
                        Exit For
                    End If
                Next
                'gvReception.Focus()
            Else
                lblLastDataEntered.Text = "No reception data has been entered previously"
                ddlMonth.Focus()
            End If
        Catch ex As Exception
            WebMsgBox.Show("Error in GetMostRecentReceptionData: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Page.Validate()
        If Page.IsValid Then
            If ddlMonth.SelectedValue > 0 AndAlso ddlYear.SelectedValue > 0 Then
                Call AllocateManagementFee()
                If SaveStorageData() And SaveReceptionData() Then
                    WebMsgBox.Show("Storage and reception data for " & ddlMonth.SelectedItem.Text & " " & ddlYear.SelectedItem.Text & " has been saved. Thank you.")
                End If
            Else
                WebMsgBox.Show("Please select month and year")
            End If
        End If
    End Sub

    Protected Sub AllocateManagementFee()
        Const CATEGORY_COUNT As Integer = 8
        Dim dblManagementFeeAllocation As Double = GetManagementFee(ddlMonth.SelectedValue, ddlYear.SelectedValue) / CATEGORY_COUNT
        Dim sbSQL As New StringBuilder
        If bRecordExists(ddlMonth.SelectedValue, ddlYear.SelectedValue, RECORD_TYPE_MANAGEMENT_FEE) Then
            sbSQL.Append("UPDATE ClientData_KODDFIS_AllocatedCharges SET ")
            'sbSQL.Append("Accessories = ")
            'sbSQL.Append(dblManagementFeeAllocation)
            'sbSQL.Append(", ")
            sbSQL.Append("DigitalCameras = ")
            sbSQL.Append(dblManagementFeeAllocation)
            sbSQL.Append(", ")
            sbSQL.Append("Frames = ")
            sbSQL.Append(dblManagementFeeAllocation)
            sbSQL.Append(", ")
            sbSQL.Append("PocketVideoCameras = ")
            sbSQL.Append(dblManagementFeeAllocation)
            sbSQL.Append(", ")
            sbSQL.Append("FieldMerchandising = ")
            sbSQL.Append(dblManagementFeeAllocation)
            sbSQL.Append(", ")
            sbSQL.Append("FilmSUC = ")
            sbSQL.Append(dblManagementFeeAllocation)
            sbSQL.Append(", ")
            sbSQL.Append("Inkjet = ")
            sbSQL.Append(dblManagementFeeAllocation)
            sbSQL.Append(", ")
            sbSQL.Append("KioskDrylab = ")
            sbSQL.Append(dblManagementFeeAllocation)
            sbSQL.Append(", ")
            sbSQL.Append("KodakExpress = ")
            sbSQL.Append(dblManagementFeeAllocation)
            sbSQL.Append(", ")
            sbSQL.Append("DateLastChanged = GETDATE() ")
            sbSQL.Append(", ")
            sbSQL.Append("LastChangedBy = 0")
            sbSQL.Append(" WHERE ")
            sbSQL.Append("Month = ")
            sbSQL.Append(ddlMonth.SelectedValue)
            sbSQL.Append(" AND ")
            sbSQL.Append(" Year = ")
            sbSQL.Append(ddlYear.SelectedValue)
            sbSQL.Append(" AND ")
            sbSQL.Append(" RecordType = ")
            sbSQL.Append(RECORD_TYPE_MANAGEMENT_FEE)
        Else
            'sbSQL.Append("INSERT INTO ClientData_KODDFIS_AllocatedCharges (Year, Month, RecordType, Accessories, DigitalCameras, Frames, PocketVideoCameras, FieldMerchandising, FilmSUC, Inkjet, KioskDrylab, KodakExpress, DateCreated, DateLastChanged, CreatedBy, LastChangedBy) VALUES (")
            sbSQL.Append("INSERT INTO ClientData_KODDFIS_AllocatedCharges (Year, Month, RecordType, DigitalCameras, Frames, PocketVideoCameras, FieldMerchandising, FilmSUC, Inkjet, KioskDrylab, KodakExpress, DateCreated, DateLastChanged, CreatedBy, LastChangedBy) VALUES (")
            sbSQL.Append(ddlYear.SelectedValue)
            sbSQL.Append(", ")
            sbSQL.Append(ddlMonth.SelectedValue)
            sbSQL.Append(", ")
            sbSQL.Append(RECORD_TYPE_MANAGEMENT_FEE)
            sbSQL.Append(", ")
            'sbSQL.Append(dblManagementFeeAllocation)
            'sbSQL.Append(", ")
            sbSQL.Append(dblManagementFeeAllocation)
            sbSQL.Append(", ")
            sbSQL.Append(dblManagementFeeAllocation)
            sbSQL.Append(", ")
            sbSQL.Append(dblManagementFeeAllocation)
            sbSQL.Append(", ")
            sbSQL.Append(dblManagementFeeAllocation)
            sbSQL.Append(", ")
            sbSQL.Append(dblManagementFeeAllocation)
            sbSQL.Append(", ")
            sbSQL.Append(dblManagementFeeAllocation)
            sbSQL.Append(", ")
            sbSQL.Append(dblManagementFeeAllocation)
            sbSQL.Append(", ")
            sbSQL.Append(dblManagementFeeAllocation)
            sbSQL.Append(", ")
            sbSQL.Append("GETDATE(), ")
            sbSQL.Append("GETDATE(), ")
            sbSQL.Append("0, ")
            sbSQL.Append("0")
            sbSQL.Append(")")
        End If
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand(sbSQL.ToString, oConn)
        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            WebMsgBox.Show("Error in AllocateManagementFee: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Function GetManagementFee(ByVal nMonth As Integer, ByVal nYear As Integer) As Double
        GetManagementFee = 0
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataReader As SqlDataReader = Nothing
        Dim sSQL As String = "SELECT MonthlyManagementFee, Year FROM CustomerFees WHERE CustomerKey = " & CUSTOMER_KEY & " AND ((Year = " & ddlYear.SelectedValue & " AND Month = " & ddlMonth.SelectedValue & ") OR (Year = 0 AND Month = 0))"
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            If oDataReader.HasRows Then
                oDataReader.Read()
                GetManagementFee = oDataReader("MonthlyManagementFee")
                If oDataReader("Year") = 0 Then
                    If oDataReader.Read() Then ' try another read in case year/month specific record present
                        GetManagementFee = oDataReader("MonthlyManagementFee")
                    Else
                        Call ExecuteNonQuery("INSERT INTO CustomerFees (CustomerKey, Year, Month, MonthlyManagementFee, FirstItemPickFee, AdditionalItemPickFee, PalletReceptionFee, PalletWeeklyFee) SELECT " & CUSTOMER_KEY & ", " & ddlYear.SelectedValue & ", " & ddlMonth.SelectedValue & ", MonthlyManagementFee, FirstItemPickFee, AdditionalItemPickFee, PalletReceptionFee, PalletWeeklyFee FROM CustomerFees WHERE Year = 0 AND Month = 0 AND CustomerKey = " & CUSTOMER_KEY)
                    End If
                End If
            Else
                WebMsgBox.Show("Error - no Management Fee specified for this customer")
            End If
        Catch ex As Exception
            WebMsgBox.Show("Error in GetManagementFee: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Function
    
    Protected Function SaveStorageData() As Boolean
        SaveStorageData = True
        Dim sbSQL As New StringBuilder
        'Dim dblAccessories As Double
        Dim dblDigitalCameras As Double
        Dim dblFrames As Double
        Dim dblPocketVideoCameras As Double
        Dim dblFieldMerchandising As Double
        Dim dblFilmSUC As Double
        Dim dblInkjet As Double
        Dim dblKioskDryLab As Double
        Dim dblKodakExpress As Double
        Dim gvr As GridViewRow = gvStorage.Rows(1)
        
        Dim tb As TextBox
        'tb = gvr.FindControl("tbAccessories")
        'If IsNumeric(tb.Text) Then
        '    dblAccessories = CDbl(tb.Text)
        'End If
        tb = gvr.FindControl("tbDigitalCameras")
        If IsNumeric(tb.Text) Then
            dblDigitalCameras = CDbl(tb.Text)
        End If
        tb = gvr.FindControl("tbFrames")
        If IsNumeric(tb.Text) Then
            dblFrames = CDbl(tb.Text)
        End If
        tb = gvr.FindControl("tbPocketVideoCameras")
        If IsNumeric(tb.Text) Then
            dblPocketVideoCameras = CDbl(tb.Text)
        End If
        tb = gvr.FindControl("tbFieldMerchandising")
        If IsNumeric(tb.Text) Then
            dblFieldMerchandising = CDbl(tb.Text)
        End If
        tb = gvr.FindControl("tbFilmSUC")
        If IsNumeric(tb.Text) Then
            dblFilmSUC = CDbl(tb.Text)
        End If
        tb = gvr.FindControl("tbInkjet")
        If IsNumeric(tb.Text) Then
            dblInkjet = CDbl(tb.Text)
        End If
        tb = gvr.FindControl("tbKioskDrylab")
        If IsNumeric(tb.Text) Then
            dblKioskDryLab = CDbl(tb.Text)
        End If
        tb = gvr.FindControl("tbKodakExpress")
        If IsNumeric(tb.Text) Then
            dblKodakExpress = CDbl(tb.Text)
        End If
        If bRecordExists(ddlMonth.SelectedValue, ddlYear.SelectedValue, RECORD_TYPE_STORAGE_COUNT) Then
            sbSQL.Append("UPDATE ClientData_KODDFIS_AllocatedCharges SET ")
            'sbSQL.Append("Accessories = ")
            'sbSQL.Append(dblAccessories)
            'sbSQL.Append(", ")
            sbSQL.Append("DigitalCameras = ")
            sbSQL.Append(dblDigitalCameras)
            sbSQL.Append(", ")
            sbSQL.Append("Frames = ")
            sbSQL.Append(dblFrames)
            sbSQL.Append(", ")
            sbSQL.Append("PocketVideoCameras = ")
            sbSQL.Append(dblPocketVideoCameras)
            sbSQL.Append(", ")
            sbSQL.Append("FieldMerchandising = ")
            sbSQL.Append(dblFieldMerchandising)
            sbSQL.Append(", ")
            sbSQL.Append("FilmSUC = ")
            sbSQL.Append(dblFilmSUC)
            sbSQL.Append(", ")
            sbSQL.Append("Inkjet = ")
            sbSQL.Append(dblInkjet)
            sbSQL.Append(", ")
            sbSQL.Append("KioskDrylab = ")
            sbSQL.Append(dblKioskDryLab)
            sbSQL.Append(", ")
            sbSQL.Append("KodakExpress = ")
            sbSQL.Append(dblKodakExpress)
            sbSQL.Append(", ")
            sbSQL.Append("DateLastChanged = GETDATE() ")
            sbSQL.Append(", ")
            sbSQL.Append("LastChangedBy = 0")
            sbSQL.Append(" WHERE ")
            sbSQL.Append("Month = ")
            sbSQL.Append(ddlMonth.SelectedValue)
            sbSQL.Append(" AND ")
            sbSQL.Append(" Year = ")
            sbSQL.Append(ddlYear.SelectedValue)
            sbSQL.Append(" AND ")
            sbSQL.Append(" RecordType = ")
            sbSQL.Append(RECORD_TYPE_STORAGE_COUNT)
        Else
            'sbSQL.Append("INSERT INTO ClientData_KODDFIS_AllocatedCharges (Year, Month, RecordType, Accessories, DigitalCameras, Frames, PocketVideoCameras, FieldMerchandising, FilmSUC, Inkjet, KioskDrylab, KodakExpress, DateCreated, DateLastChanged, CreatedBy, LastChangedBy) VALUES (")
            sbSQL.Append("INSERT INTO ClientData_KODDFIS_AllocatedCharges (Year, Month, RecordType, DigitalCameras, Frames, PocketVideoCameras, FieldMerchandising, FilmSUC, Inkjet, KioskDrylab, KodakExpress, DateCreated, DateLastChanged, CreatedBy, LastChangedBy) VALUES (")
            sbSQL.Append(ddlYear.SelectedValue)
            sbSQL.Append(", ")
            sbSQL.Append(ddlMonth.SelectedValue)
            sbSQL.Append(", ")
            sbSQL.Append(RECORD_TYPE_STORAGE_COUNT)
            sbSQL.Append(", ")
            'sbSQL.Append(dblAccessories)
            'sbSQL.Append(", ")
            sbSQL.Append(dblDigitalCameras)
            sbSQL.Append(", ")
            sbSQL.Append(dblFrames)
            sbSQL.Append(", ")
            sbSQL.Append(dblPocketVideoCameras)
            sbSQL.Append(", ")
            sbSQL.Append(dblFieldMerchandising)
            sbSQL.Append(", ")
            sbSQL.Append(dblFilmSUC)
            sbSQL.Append(", ")
            sbSQL.Append(dblInkjet)
            sbSQL.Append(", ")
            sbSQL.Append(dblKioskDryLab)
            sbSQL.Append(", ")
            sbSQL.Append(dblKodakExpress)
            sbSQL.Append(", ")
            sbSQL.Append("GETDATE(), ")
            sbSQL.Append("GETDATE(), ")
            sbSQL.Append("0, ")
            sbSQL.Append("0")
            sbSQL.Append(")")
        End If
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand(sbSQL.ToString, oConn)
        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            SaveStorageData = False
            WebMsgBox.Show("Error in SaveStorageData: " & ex.Message)
        Finally
            oConn.Close()
            'WebMsgBox.Show("Thank you. Kodak storage costs for " & ddlMonth.SelectedItem.Text & " " & ddlYear.SelectedItem.Text & " have been saved.")
        End Try
    End Function
    
    Protected Function SaveReceptionData() As Boolean
        SaveReceptionData = True
        Dim sbSQL As New StringBuilder
        'Dim dblAccessories As Double
        Dim dblDigitalCameras As Double
        Dim dblFrames As Double
        Dim dblPocketVideoCameras As Double
        Dim dblFieldMerchandising As Double
        Dim dblFilmSUC As Double
        Dim dblInkjet As Double
        Dim dblKioskDryLab As Double
        Dim dblKodakExpress As Double
        Dim gvr As GridViewRow = gvReception.Rows(1)
        
        Dim tb As TextBox
        'tb = gvr.FindControl("tbAccessories")
        'If IsNumeric(tb.Text) Then
        '    dblAccessories = CDbl(tb.Text)
        'End If
        tb = gvr.FindControl("tbDigitalCameras")
        If IsNumeric(tb.Text) Then
            dblDigitalCameras = CDbl(tb.Text)
        End If
        tb = gvr.FindControl("tbFrames")
        If IsNumeric(tb.Text) Then
            dblFrames = CDbl(tb.Text)
        End If
        tb = gvr.FindControl("tbPocketVideoCameras")
        If IsNumeric(tb.Text) Then
            dblPocketVideoCameras = CDbl(tb.Text)
        End If
        tb = gvr.FindControl("tbFieldMerchandising")
        If IsNumeric(tb.Text) Then
            dblFieldMerchandising = CDbl(tb.Text)
        End If
        tb = gvr.FindControl("tbFilmSUC")
        If IsNumeric(tb.Text) Then
            dblFilmSUC = CDbl(tb.Text)
        End If
        tb = gvr.FindControl("tbInkjet")
        If IsNumeric(tb.Text) Then
            dblInkjet = CDbl(tb.Text)
        End If
        tb = gvr.FindControl("tbKioskDrylab")
        If IsNumeric(tb.Text) Then
            dblKioskDryLab = CDbl(tb.Text)
        End If
        tb = gvr.FindControl("tbKodakExpress")
        If IsNumeric(tb.Text) Then
            dblKodakExpress = CDbl(tb.Text)
        End If
        If bRecordExists(ddlMonth.SelectedValue, ddlYear.SelectedValue, RECORD_TYPE_RECEPTION_COUNT) Then
            sbSQL.Append("UPDATE ClientData_KODDFIS_AllocatedCharges SET ")
            'sbSQL.Append("Accessories = ")
            'sbSQL.Append(dblAccessories)
            'sbSQL.Append(", ")
            sbSQL.Append("DigitalCameras = ")
            sbSQL.Append(dblDigitalCameras)
            sbSQL.Append(", ")
            sbSQL.Append("Frames = ")
            sbSQL.Append(dblFrames)
            sbSQL.Append(", ")
            sbSQL.Append("PocketVideoCameras = ")
            sbSQL.Append(dblPocketVideoCameras)
            sbSQL.Append(", ")
            sbSQL.Append("FieldMerchandising = ")
            sbSQL.Append(dblFieldMerchandising)
            sbSQL.Append(", ")
            sbSQL.Append("FilmSUC = ")
            sbSQL.Append(dblFilmSUC)
            sbSQL.Append(", ")
            sbSQL.Append("Inkjet = ")
            sbSQL.Append(dblInkjet)
            sbSQL.Append(", ")
            sbSQL.Append("KioskDrylab = ")
            sbSQL.Append(dblKioskDryLab)
            sbSQL.Append(", ")
            sbSQL.Append("KodakExpress = ")
            sbSQL.Append(dblKodakExpress)
            sbSQL.Append(", ")
            sbSQL.Append("DateLastChanged = GETDATE() ")
            sbSQL.Append(", ")
            sbSQL.Append("LastChangedBy = 0")
            sbSQL.Append(" WHERE ")
            sbSQL.Append("Month = ")
            sbSQL.Append(ddlMonth.SelectedValue)
            sbSQL.Append(" AND ")
            sbSQL.Append(" Year = ")
            sbSQL.Append(ddlYear.SelectedValue)
            sbSQL.Append(" AND ")
            sbSQL.Append(" RecordType = ")
            sbSQL.Append(RECORD_TYPE_RECEPTION_COUNT)
        Else
            'sbSQL.Append("INSERT INTO ClientData_KODDFIS_AllocatedCharges (Year, Month, RecordType, Accessories, DigitalCameras, Frames, PocketVideoCameras, FieldMerchandising, FilmSUC, Inkjet, KioskDrylab, KodakExpress, DateCreated, DateLastChanged, CreatedBy, LastChangedBy) VALUES (")
            sbSQL.Append("INSERT INTO ClientData_KODDFIS_AllocatedCharges (Year, Month, RecordType, DigitalCameras, Frames, PocketVideoCameras, FieldMerchandising, FilmSUC, Inkjet, KioskDrylab, KodakExpress, DateCreated, DateLastChanged, CreatedBy, LastChangedBy) VALUES (")
            sbSQL.Append(ddlYear.SelectedValue)
            sbSQL.Append(", ")
            sbSQL.Append(ddlMonth.SelectedValue)
            sbSQL.Append(", ")
            sbSQL.Append(RECORD_TYPE_RECEPTION_COUNT)
            sbSQL.Append(", ")
            'sbSQL.Append(dblAccessories)
            'sbSQL.Append(", ")
            sbSQL.Append(dblDigitalCameras)
            sbSQL.Append(", ")
            sbSQL.Append(dblFrames)
            sbSQL.Append(", ")
            sbSQL.Append(dblPocketVideoCameras)
            sbSQL.Append(", ")
            sbSQL.Append(dblFieldMerchandising)
            sbSQL.Append(", ")
            sbSQL.Append(dblFilmSUC)
            sbSQL.Append(", ")
            sbSQL.Append(dblInkjet)
            sbSQL.Append(", ")
            sbSQL.Append(dblKioskDryLab)
            sbSQL.Append(", ")
            sbSQL.Append(dblKodakExpress)
            sbSQL.Append(", ")
            sbSQL.Append("GETDATE(), ")
            sbSQL.Append("GETDATE(), ")
            sbSQL.Append("0, ")
            sbSQL.Append("0")
            sbSQL.Append(")")
        End If
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand(sbSQL.ToString, oConn)
        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            SaveReceptionData = False
            WebMsgBox.Show("Error in SaveReceptionData: " & ex.Message)
        Finally
            oConn.Close()
            'WebMsgBox.Show("Thank you. Kodak reception costs for " & ddlMonth.SelectedItem.Text & " " & ddlYear.SelectedItem.Text & " have been saved.")
        End Try
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
    
    ' FROM BEGINNING OF GVDATA <Columns> section...
    
    '    <asp:TemplateField HeaderText="Accessories">
    '    <EditItemTemplate>
    '        <asp:TextBox ID="tbAccessories" runat="server" Font-Names="Verdana" Font-Size="XX-Small"/>
    '        <br />
    '        <asp:RangeValidator ID="rvAccessories" ControlToValidate="tbAccessories" Type="Currency" MinimumValue="0" MaximumValue="999999" runat="server" ErrorMessage="not a number!!!!"/>
    '    </EditItemTemplate>
    '    <ItemTemplate>
    '        <asp:Label ID="lblAccessories" runat="server" Text='<%# Container.DataItem("Accessories")%>'/>
    '    </ItemTemplate>
    '</asp:TemplateField>


</script>

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Kodak Monthly Pallet Storage & Reception Record</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
      <main:Header id="ctlHeader" runat="server"></main:Header>
        <table width="100%" cellpadding="0" cellspacing="0">
            <tr class="bar_accounthandler">
                <td style="width:50%; white-space:nowrap">
                </td>
                <td style="width:50%; white-space:nowrap" align="right">
                </td>
            </tr>
        </table>
        <asp:Label ID="Label1" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small"
            Text="Kodak Monthly Pallet Storage & Reception Record"></asp:Label><br />
        <br />
        <asp:Label ID="Label4" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small"
            ForeColor="Maroon" Text="INSTRUCTIONS:<br /><br />1. Ensure you have selected the correct month and year for the data you want to save.<br /><br />2. Enter the month's data into the relevant boxes. The data displayed above the boxes, if any, is the most recent data previously entered - usually last month's data.<br /><br />3. Click the <b>save</b> button. You should see a message box confirming that the data has been saved successfully.<br /><br />NOTE: You can re-enter the data for a month if required, but before doing so check that the data has not already been seen by Kodak. Alert the Account Handler for the Kodak account if you're unsure about this."></asp:Label><br />
        <br />
        <asp:Label ID="Label3" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small"
            Text="Month:"></asp:Label>
        <asp:DropDownList ID="ddlMonth" runat="server" Font-Names="Verdana" Font-Size="XX-Small">
            <asp:ListItem Value="0">- please select -</asp:ListItem>
            <asp:ListItem Value="1">January</asp:ListItem>
            <asp:ListItem Value="2">February</asp:ListItem>
            <asp:ListItem Value="3">March</asp:ListItem>
            <asp:ListItem Value="4">April</asp:ListItem>
            <asp:ListItem Value="5">May</asp:ListItem>
            <asp:ListItem Value="6">June</asp:ListItem>
            <asp:ListItem Value="7">July</asp:ListItem>
            <asp:ListItem Value="8">August</asp:ListItem>
            <asp:ListItem Value="9">September</asp:ListItem>
            <asp:ListItem Value="10">October</asp:ListItem>
            <asp:ListItem Value="11">November</asp:ListItem>
            <asp:ListItem Value="12">December</asp:ListItem>
        </asp:DropDownList>
        <asp:Label ID="Label2" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small"
            Text="Year:"></asp:Label>
        <asp:DropDownList ID="ddlYear" runat="server" Font-Names="Verdana" Font-Size="XX-Small">
            <asp:ListItem Value="0">- please select -</asp:ListItem>
            <asp:ListItem>2008</asp:ListItem>
            <asp:ListItem>2009</asp:ListItem>
            <asp:ListItem>2010</asp:ListItem>
            <asp:ListItem>2011</asp:ListItem>
            <asp:ListItem>2012</asp:ListItem>
            <asp:ListItem>2013</asp:ListItem>
            <asp:ListItem>2014</asp:ListItem>
            <asp:ListItem>2015</asp:ListItem>
            <asp:ListItem>2016</asp:ListItem>
            <asp:ListItem>2017</asp:ListItem>
            <asp:ListItem>2018</asp:ListItem>
            <asp:ListItem>2019</asp:ListItem>
            <asp:ListItem>2020</asp:ListItem>
        </asp:DropDownList>
        &nbsp; &nbsp; &nbsp;
        <asp:Label ID="lblLastDataEntered" runat="server" Font-Bold="False" Font-Names="Verdana"
            Font-Size="XX-Small"></asp:Label><br />
        <br />
        <asp:Label ID="Label5" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small"
            Text="Pallet storage:"></asp:Label><br />
        <asp:GridView ID="gvStorage" runat="server" AutoGenerateColumns="False" CellPadding="2"
            Font-Names="Verdana" Font-Size="XX-Small" Width="100%">
            <Columns>
                <asp:TemplateField HeaderText="Digital Cameras">
                    <EditItemTemplate>
                        <asp:TextBox ID="tbDigitalCameras" runat="server" Font-Names="Verdana" Font-Size="XX-Small"></asp:TextBox>
                        <br />
                        <asp:RangeValidator ID="rvDigitalCameras" ControlToValidate="tbDigitalCameras" Type="Currency" MinimumValue="0" MaximumValue="999999" runat="server" ErrorMessage="not a number!!!!"/>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="lblDigitalCameras" runat="server" Text='<%# Container.DataItem("DigitalCameras")%>'/>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Frames">
                    <EditItemTemplate>
                        <asp:TextBox ID="tbFrames" runat="server" Font-Names="Verdana" Font-Size="XX-Small"/>
                        <br />
                        <asp:RangeValidator ID="rvFrames" ControlToValidate="tbFrames" Type="Currency" MinimumValue="0" MaximumValue="999999" runat="server" ErrorMessage="not a number!!!!"/>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="lblFrames" runat="server" Text='<%# Container.DataItem("Frames")%>'/>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Pocket Video Cameras">
                    <EditItemTemplate>
                        <asp:TextBox ID="tbPocketVideoCameras" runat="server" Font-Names="Verdana" Font-Size="XX-Small"/>
                        <br />
                        <asp:RangeValidator ID="rvPocketVideoCameras" ControlToValidate="tbPocketVideoCameras" Type="Currency" MinimumValue="0" MaximumValue="999999" runat="server" ErrorMessage="not a number!!!!"/>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="lblPocketVideoCameras" runat="server" Text='<%# Container.DataItem("PocketVideoCameras")%>'/>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Field Merchandising">
                    <EditItemTemplate>
                        <asp:TextBox ID="tbFieldMerchandising" runat="server" Font-Names="Verdana" Font-Size="XX-Small"/>
                        <br />
                        <asp:RangeValidator ID="rvFieldMerchandising" ControlToValidate="tbFieldMerchandising" Type="Currency" MinimumValue="0" MaximumValue="999999" runat="server" ErrorMessage="not a number!!!!"/>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="lblFieldMerchandising" runat="server" Text='<%# Container.DataItem("FieldMerchandising")%>'/>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Film &amp; SUC">
                    <EditItemTemplate>
                        <asp:TextBox ID="tbFilmSUC" runat="server" Font-Names="Verdana" Font-Size="XX-Small"/>
                        <br />
                        <asp:RangeValidator ID="rvFilmSUC" ControlToValidate="tbFilmSUC" Type="Currency" MinimumValue="0" MaximumValue="999999" runat="server" ErrorMessage="not a number!!!!"/>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="lblFilmSUC" runat="server" Text='<%# Container.DataItem("FilmSUC")%>'/>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Inkjet">
                    <EditItemTemplate>
                        <asp:TextBox ID="tbInkjet" runat="server" Font-Names="Verdana" Font-Size="XX-Small"/>
                        <br />
                        <asp:RangeValidator ID="rvInkjet" ControlToValidate="tbInkjet" Type="Currency" MinimumValue="0" MaximumValue="999999" runat="server" ErrorMessage="not a number!!!!"/>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="lblInkjet" runat="server" Text='<%# Container.DataItem("Inkjet")%>'/>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Kiosk &amp; Dry Lab">
                    <EditItemTemplate>
                        <asp:TextBox ID="tbKioskDryLab" runat="server" Font-Names="Verdana" Font-Size="XX-Small"/>
                        <br />
                        <asp:RangeValidator ID="rvKioskDryLab" ControlToValidate="tbKioskDryLab" Type="Currency" MinimumValue="0" MaximumValue="999999" runat="server" ErrorMessage="not a number!!!!"/>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="lblKioskDryLab" runat="server" Text='<%# Container.DataItem("KioskDryLab")%>'/>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Kodak Express">
                    <EditItemTemplate>
                        <asp:TextBox ID="tbKodakExpress" runat="server" Font-Names="Verdana" Font-Size="XX-Small"/>
                        <br />
                        <asp:RangeValidator ID="rvKodakExpress" ControlToValidate="tbKodakExpress" Type="Currency" MinimumValue="0" MaximumValue="999999" runat="server" ErrorMessage="not a number!!!!"/>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="lblKodakExpress" runat="server" Text='<%# Container.DataItem("KodakExpress")%>'/>
                    </ItemTemplate>
                </asp:TemplateField>
            </Columns>
        </asp:GridView>
        <br />
        <asp:Label ID="Label6" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small" Text="Pallet reception:"></asp:Label><br />
        <asp:GridView ID="gvReception" runat="server" AutoGenerateColumns="False" CellPadding="2" Font-Names="Verdana" Font-Size="XX-Small" Width="100%">
            <Columns>
                <asp:TemplateField HeaderText="Digital Cameras">
                    <EditItemTemplate>
                        <asp:TextBox ID="tbDigitalCameras" runat="server" Font-Names="Verdana" Font-Size="XX-Small"></asp:TextBox>
                        <br />
                        <asp:RangeValidator ID="rvDigitalCameras" ControlToValidate="tbDigitalCameras" Type="Currency" MinimumValue="0" MaximumValue="999999" runat="server" ErrorMessage="not a number!!!!"/>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="lblDigitalCameras" runat="server" Text='<%# Container.DataItem("DigitalCameras")%>'/>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Frames">
                    <EditItemTemplate>
                        <asp:TextBox ID="tbFrames" runat="server" Font-Names="Verdana" Font-Size="XX-Small"/>
                        <br />
                        <asp:RangeValidator ID="rvFrames" ControlToValidate="tbFrames" Type="Currency" MinimumValue="0" MaximumValue="999999" runat="server" ErrorMessage="not a number!!!!"/>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="lblFrames" runat="server" Text='<%# Container.DataItem("Frames")%>'/>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Pocket Video Cameras">
                    <EditItemTemplate>
                        <asp:TextBox ID="tbPocketVideoCameras" runat="server" Font-Names="Verdana" Font-Size="XX-Small"/>
                        <br />
                        <asp:RangeValidator ID="rvPocketVideoCameras" ControlToValidate="tbPocketVideoCameras" Type="Currency" MinimumValue="0" MaximumValue="999999" runat="server" ErrorMessage="not a number!!!!"/>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="lblPocketVideoCameras" runat="server" Text='<%# Container.DataItem("PocketVideoCameras")%>'/>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Field Merchandising">
                    <EditItemTemplate>
                        <asp:TextBox ID="tbFieldMerchandising" runat="server" Font-Names="Verdana" Font-Size="XX-Small"/>
                        <br />
                        <asp:RangeValidator ID="rvFieldMerchandising" ControlToValidate="tbFieldMerchandising" Type="Currency" MinimumValue="0" MaximumValue="999999" runat="server" ErrorMessage="not a number!!!!"/>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="lblFieldMerchandising" runat="server" Text='<%# Container.DataItem("FieldMerchandising")%>'/>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Film &amp; SUC">
                    <EditItemTemplate>
                        <asp:TextBox ID="tbFilmSUC" runat="server" Font-Names="Verdana" Font-Size="XX-Small"/>
                        <br />
                        <asp:RangeValidator ID="rvFilmSUC" ControlToValidate="tbFilmSUC" Type="Currency" MinimumValue="0" MaximumValue="999999" runat="server" ErrorMessage="not a number!!!!"/>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="lblFilmSUC" runat="server" Text='<%# Container.DataItem("FilmSUC")%>'/>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Inkjet">
                    <EditItemTemplate>
                        <asp:TextBox ID="tbInkjet" runat="server" Font-Names="Verdana" Font-Size="XX-Small"/>
                        <br />
                        <asp:RangeValidator ID="rvInkjet" ControlToValidate="tbInkjet" Type="Currency" MinimumValue="0" MaximumValue="999999" runat="server" ErrorMessage="not a number!!!!"/>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="lblInkjet" runat="server" Text='<%# Container.DataItem("Inkjet")%>'/>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Kiosk &amp; Dry Lab">
                    <EditItemTemplate>
                        <asp:TextBox ID="tbKioskDryLab" runat="server" Font-Names="Verdana" Font-Size="XX-Small"/>
                        <br />
                        <asp:RangeValidator ID="rvKioskDryLab" ControlToValidate="tbKioskDryLab" Type="Currency" MinimumValue="0" MaximumValue="999999" runat="server" ErrorMessage="not a number!!!!"/>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="lblKioskDryLab" runat="server" Text='<%# Container.DataItem("KioskDryLab")%>'/>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Kodak Express">
                    <EditItemTemplate>
                        <asp:TextBox ID="tbKodakExpress" runat="server" Font-Names="Verdana" Font-Size="XX-Small"/>
                        <br />
                        <asp:RangeValidator ID="rvKodakExpress" ControlToValidate="tbKodakExpress" Type="Currency" MinimumValue="0" MaximumValue="999999" runat="server" ErrorMessage="not a number!!!!"/>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="lblKodakExpress" runat="server" Text='<%# Container.DataItem("KodakExpress")%>'/>
                    </ItemTemplate>
                </asp:TemplateField>
            </Columns>
        </asp:GridView>
        <br />
        <asp:Button ID="btnSave" runat="server" Text="save" OnClick="btnSave_Click" Width="150px" />
        &nbsp;
        &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
        <asp:Button ID="btnCloseWindow" runat="server" OnClientClick="window.close()" Text="close window" /></div>
    </form>
</body>
</html>