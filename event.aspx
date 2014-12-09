<%@ Page Language="VB" Theme="AIMSDefault"  %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

    Const MIN_CHANGE_PERIOD_DAYS As Integer = 5
    Const COUNTRY_KEY_UK As Integer = 222
    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    Dim sGUID As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsPostBack Then
            sGUID = Request.QueryString("GUID")
            If sGUID = String.Empty Then
                WebMsgBox.Show("Error in request - please check the parameters you supplied")
            Else
                Select Case GetEventFromGUID()
                    Case 0
                        WebMsgBox.Show("No event found - if the event you are requesting finished more than one month ago it may have expired from the system")
                    Case 1
                        WebMsgBox.Show("This event has been deleted")
                    Case 2
                        Call ShowEvent()
                End Select
            End If
        End If
    End Sub
    
    Protected Sub HideAllPanels()
        pnlEvent.Visible = False
        pnlMessage.Visible = False
    End Sub
    
    Protected Sub ShowEvent()
        Call HideAllPanels()
        pnlEvent.Visible = True
        Dim dtDeliveryDate As Date = Date.Parse(lblDeliveryDate.Text)
        Dim nDeliveryInterval As Integer = DateDiff(DateInterval.Day, Date.Now, dtDeliveryDate)
        If nDeliveryInterval <= MIN_CHANGE_PERIOD_DAYS Then
            btnSaveChanges.Visible = False
            cbDifferentCollectionAddress.AutoPostBack = False
            
            cbDifferentCollectionAddress.Enabled = False
            lblOnlineChangesMessage.Text = "No changes can be accepted online as there are " & MIN_CHANGE_PERIOD_DAYS.ToString & " days or fewer remaining until delivery. Contact Customer Services to request changes."
            lblOnlineChangesMessage.Visible = True
        End If
    End Sub
    
    Protected Sub ShowMessage()
        Call HideAllPanels()
        pnlMessage.Visible = True
    End Sub
    
    Protected Function GetEventFromGUID() As Integer
        Dim nDDLIndex As Integer
        Call InitCountryDropdowns()
        GetEventFromGUID = 0
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataTable1 As New DataTable
        Dim oAdapter1 As New SqlDataAdapter("spASPNET_CalendarManaged_GetEventByGUID2", oConn)

        oAdapter1.SelectCommand.CommandType = CommandType.StoredProcedure
        oAdapter1.SelectCommand.Parameters.Add(New SqlParameter("@AccessGUID", SqlDbType.VarChar, 30))
        oAdapter1.SelectCommand.Parameters("@AccessGUID").Value = sGUID
        
        Try
            oConn.Open()
            oAdapter1.Fill(oDataTable1)
            If oDataTable1.Rows.Count > 0 Then
                Dim dr As DataRow = oDataTable1.Rows(0)
                If Not IsDBNull(dr("IsDeleted")) Then
                    If dr("IsDeleted") = 1 Then
                        GetEventFromGUID = 1
                    Else
                        GetEventFromGUID = 2
                    End If
                Else
                    GetEventFromGUID = 2
                End If
                ' need to check 1 and only 1 row present
                pnEventId = dr("id")
                lblEventName.Text = dr("EventName")
                lblDeliveryDate.Text = dr("Delivery Date")
                lblCollectionDate.Text = dr("Collection Date")
                tbContactName.Text = dr("ContactName")
                tbContactPhone.Text = dr("ContactPhone")
                tbContactMobile.Text = dr("ContactMobile")
                
                Dim sContactName2 As String
                Dim sContactPhone2 As String
                Dim sContactMobile2 As String
            
                If Not IsDBNull(dr("ContactName2")) Then
                    sContactName2 = dr("ContactName2").ToString.Trim
                Else
                    sContactName2 = String.Empty
                End If
                If Not IsDBNull(dr("ContactPhone2")) Then
                    sContactPhone2 = dr("ContactPhone2").ToString.Trim
                Else
                    sContactPhone2 = String.Empty
                End If
                If Not IsDBNull(dr("ContactMobile2")) Then
                    sContactMobile2 = dr("ContactMobile2").ToString.Trim
                Else
                    sContactMobile2 = String.Empty
                End If
            
                tbCMContactName2.Text = sContactName2
                tbCMContactPhone2.Text = sContactPhone2
                tbCMContactMobile2.Text = sContactMobile2
                If String.IsNullOrEmpty(sContactName2) And String.IsNullOrEmpty(sContactPhone2) And String.IsNullOrEmpty(sContactMobile2) Then
                    Call SetContact2FieldsVisibility(False)
                Else
                    Call SetContact2FieldsVisibility(True)
                End If

                tbEventAddress1.Text = dr("EventAddress1")
                tbEventAddress2.Text = dr("EventAddress2")
                tbEventAddress3.Text = dr("EventAddress3")
                tbTown.Text = dr("Town")
                tbPostcode.Text = dr("Postcode")

                Dim nCountryKey As Integer
                If Not IsDBNull(dr("CountryKey")) Then
                    nCountryKey = dr("CountryKey")
                Else
                    nCountryKey = COUNTRY_KEY_UK
                End If

                If nCountryKey = COUNTRY_KEY_UK Then
                    trCMCountry.Visible = False
                    lnkbtnCMAddressOutsideUK.Visible = True
                Else
                    trCMCountry.Visible = True
                    lnkbtnCMAddressOutsideUK.Visible = False
                End If

                For nDDLIndex = 1 To ddlCMCountry.Items.Count - 1
                    If ddlCMCountry.Items(nDDLIndex).Value = nCountryKey Then
                        ddlCMCountry.SelectedIndex = nDDLIndex
                        Exit For
                    End If
                Next
            
                Dim sTemp As String = dr("DeliveryTime")
                For nDDLIndex = 0 To ddlDeliveryTime.Items.Count - 1
                    If ddlDeliveryTime.Items(nDDLIndex).Text = sTemp Then
                        ddlDeliveryTime.SelectedIndex = nDDLIndex
                        Exit For
                    End If
                Next
                tbPreciseDeliveryPoint.Text = dr("PreciseDeliveryPoint")
                tbPreciseCollectionPoint.Text = dr("PreciseCollectionPoint")
                sTemp = dr("CollectionTime")
                For nDDLIndex = 0 To ddlCollectionTime.Items.Count - 1
                    If ddlCollectionTime.Items(nDDLIndex).Text = sTemp Then
                        ddlCollectionTime.SelectedIndex = nDDLIndex
                        Exit For
                    End If
                Next
                If Not IsDBNull(dr("CustomerReference")) Then
                    tbCustomerReference.Text = dr("CustomerReference")
                Else
                    tbCustomerReference.Text = String.Empty
                End If
                tbSpecialInstructions.Text = dr("SpecialInstructions")
                lblBookedBy.Text = dr("username")
                lblBookedOn.Text = dr("BookedOn")
                
                tbCollectionAddress1.Text = String.Empty
                tbCollectionAddress2.Text = String.Empty
                tbCollectionTown.Text = String.Empty
                tbCollectionPostcode.Text = String.Empty
                
                If Not IsDBNull(dr("DifferentCollectionAddress")) Then
                    If dr("DifferentCollectionAddress") = True Then
                        trCollection1.Visible = True
                        trCollection2.Visible = True
                        cbDifferentCollectionAddress.Checked = True
                        If Not IsDBNull(dr("CollectionAddress1")) Then
                            tbCollectionAddress1.Text = dr("CollectionAddress1")
                        End If
                        If Not IsDBNull(dr("CollectionAddress1")) Then
                            tbCollectionAddress2.Text = dr("CollectionAddress2")
                        End If
                        If Not IsDBNull(dr("CollectionAddress1")) Then
                            tbCollectionTown.Text = dr("CollectionTown")
                        End If
                        If Not IsDBNull(dr("CollectionAddress1")) Then
                            tbCollectionPostcode.Text = dr("CollectionPostCode")
                        End If
                    Else
                        trCollection1.Visible = False
                        trCollection2.Visible = False
                        cbDifferentCollectionAddress.Checked = False
                    End If
                End If
            
                Dim oAdapter2 As New SqlDataAdapter("spASPNET_CalendarManaged_GetEventItemsById", oConn)
            
                oAdapter2.SelectCommand.CommandType = CommandType.StoredProcedure
                oAdapter2.SelectCommand.Parameters.Add(New SqlParameter("@EventId", SqlDbType.Int))
                oAdapter2.SelectCommand.Parameters("@EventId").Value = pnEventId
                Dim oDataTable2 As New DataTable
                oAdapter2.Fill(oDataTable2)
                gvItems.DataSource = oDataTable2
                gvItems.DataBind()
            
                If gvItems.Rows.Count = 1 Then
                    lblLegendProduct.Text = "Product:"
                Else
                    lblLegendProduct.Text = "Products:"
                End If
                Call GetNotes()
            End If
        Catch ex As Exception
            WebMsgBox.Show(ex.Message)
        Finally
            oConn.Close()
        End Try
    End Function

    Protected Sub btnSaveChanges_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Page.Validate("CalendarManaged")
        If Page.IsValid Then
            Call SaveChanges()
            WebMsgBox.Show("Your changes were saved.")
        Else
            WebMsgBox.Show("One or more fields were incorrect or not supplied. Please correct the information and resubmit.")
        End If
    End Sub
    
    Protected Sub SaveChanges()
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_CalendarManaged_UpdateEvent3", oConn)
        oCmd.CommandType = CommandType.StoredProcedure

        Dim paramEventId As SqlParameter = New SqlParameter("@EventId", SqlDbType.Int)
        paramEventId.Value = pnEventId
        oCmd.Parameters.Add(paramEventId)

        Dim paramContactName As SqlParameter = New SqlParameter("@ContactName", SqlDbType.VarChar, 50)
        paramContactName.Value = tbContactName.Text
        oCmd.Parameters.Add(paramContactName)

        Dim paramContactPhone As SqlParameter = New SqlParameter("@ContactPhone", SqlDbType.VarChar, 50)
        paramContactPhone.Value = tbContactPhone.Text
        oCmd.Parameters.Add(paramContactPhone)

        Dim paramContactMobile As SqlParameter = New SqlParameter("@ContactMobile", SqlDbType.VarChar, 50)
        paramContactMobile.Value = tbContactMobile.Text
        oCmd.Parameters.Add(paramContactMobile)

        Dim paramContactName2 As SqlParameter = New SqlParameter("@ContactName2", SqlDbType.VarChar, 50)
        paramContactName2.Value = tbCMContactName2.Text
        oCmd.Parameters.Add(paramContactName2)

        Dim paramContactPhone2 As SqlParameter = New SqlParameter("@ContactPhone2", SqlDbType.VarChar, 50)
        paramContactPhone2.Value = tbCMContactPhone2.Text
        oCmd.Parameters.Add(paramContactPhone2)

        Dim paramContactMobile2 As SqlParameter = New SqlParameter("@ContactMobile2", SqlDbType.VarChar, 50)
        paramContactMobile2.Value = tbCMContactMobile2.Text
        oCmd.Parameters.Add(paramContactMobile2)

        Dim paramEventAddress1 As SqlParameter = New SqlParameter("@EventAddress1", SqlDbType.VarChar, 50)
        paramEventAddress1.Value = tbEventAddress1.Text
        oCmd.Parameters.Add(paramEventAddress1)

        Dim paramEventAddress2 As SqlParameter = New SqlParameter("@EventAddress2", SqlDbType.VarChar, 50)
        paramEventAddress2.Value = tbEventAddress2.Text
        oCmd.Parameters.Add(paramEventAddress2)

        Dim paramEventAddress3 As SqlParameter = New SqlParameter("@EventAddress3", SqlDbType.VarChar, 50)
        paramEventAddress3.Value = tbEventAddress3.Text
        oCmd.Parameters.Add(paramEventAddress3)

        Dim paramTown As SqlParameter = New SqlParameter("@Town", SqlDbType.VarChar, 50)
        paramTown.Value = tbTown.Text
        oCmd.Parameters.Add(paramTown)

        Dim paramPostcode As SqlParameter = New SqlParameter("@Postcode", SqlDbType.VarChar, 50)
        paramPostcode.Value = tbPostcode.Text
        oCmd.Parameters.Add(paramPostcode)

        Dim paramCountryKey As SqlParameter = New SqlParameter("@CountryKey", SqlDbType.Int)
        paramCountryKey.Value = ddlCMCountry.SelectedValue
        oCmd.Parameters.Add(paramCountryKey)

        Dim paramDeliveryTime As SqlParameter = New SqlParameter("@DeliveryTime", SqlDbType.VarChar, 50)
        paramDeliveryTime.Value = ddlDeliveryTime.SelectedItem.Text
        oCmd.Parameters.Add(paramDeliveryTime)

        Dim paramPreciseDeliveryPoint As SqlParameter = New SqlParameter("@PreciseDeliveryPoint", SqlDbType.VarChar, 100)
        paramPreciseDeliveryPoint.Value = tbPreciseDeliveryPoint.Text
        oCmd.Parameters.Add(paramPreciseDeliveryPoint)

        Dim paramDifferentCollectionAddress As SqlParameter = New SqlParameter("@DifferentCollectionAddress", SqlDbType.Bit)
        paramDifferentCollectionAddress.Value = cbDifferentCollectionAddress.Checked
        oCmd.Parameters.Add(paramDifferentCollectionAddress)

        Dim paramCollectionAddress1 As SqlParameter = New SqlParameter("@CollectionAddress1", SqlDbType.NVarChar, 50)
        paramCollectionAddress1.Value = tbCollectionAddress1.Text
        oCmd.Parameters.Add(paramCollectionAddress1)

        Dim paramCollectionAddress2 As SqlParameter = New SqlParameter("@CollectionAddress2", SqlDbType.NVarChar, 50)
        paramCollectionAddress2.Value = tbCollectionAddress2.Text
        oCmd.Parameters.Add(paramCollectionAddress2)

        Dim paramCollectionTown As SqlParameter = New SqlParameter("@CollectionTown", SqlDbType.NVarChar, 50)
        paramCollectionTown.Value = tbCollectionTown.Text
        oCmd.Parameters.Add(paramCollectionTown)

        Dim paramCollectionPostcode As SqlParameter = New SqlParameter("@CollectionPostcode", SqlDbType.NVarChar, 50)
        paramCollectionPostcode.Value = tbCollectionPostcode.Text
        oCmd.Parameters.Add(paramCollectionPostcode)

        Dim paramCollectionCountryKey As SqlParameter = New SqlParameter("@CollectionCountryKey", SqlDbType.Int)
        paramCollectionCountryKey.Value = ddlCMCollectionCountry.SelectedValue
        oCmd.Parameters.Add(paramCollectionCountryKey)

        Dim paramCollectionTime As SqlParameter = New SqlParameter("@CollectionTime", SqlDbType.VarChar, 50)
        paramCollectionTime.Value = ddlCollectionTime.SelectedItem.Text
        oCmd.Parameters.Add(paramCollectionTime)

        Dim paramPreciseCollectionPoint As SqlParameter = New SqlParameter("@PreciseCollectionPoint", SqlDbType.VarChar, 100)
        paramPreciseCollectionPoint.Value = tbPreciseCollectionPoint.Text
        oCmd.Parameters.Add(paramPreciseCollectionPoint)

        Dim paramSpecialInstructions As SqlParameter = New SqlParameter("@SpecialInstructions", SqlDbType.VarChar, 200)
        paramSpecialInstructions.Value = tbSpecialInstructions.Text
        oCmd.Parameters.Add(paramSpecialInstructions)

        Dim paramCustomerReference As SqlParameter = New SqlParameter("@CustomerReference", SqlDbType.NVarChar, 100)
        paramCustomerReference.Value = tbCustomerReference.Text
        oCmd.Parameters.Add(paramCustomerReference)

        Dim paramUpdatedBy As SqlParameter = New SqlParameter("@UpdatedBy", SqlDbType.Int)
        'paramUpdatedBy.Value = Session("UserKey")
        paramUpdatedBy.Value = 0
        oCmd.Parameters.Add(paramUpdatedBy)

        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As SqlException
            WebMsgBox.Show(ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub lnkbtnShowHideNotes_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ToggleNotesGrid()
    End Sub
    
    Protected Sub ToggleNotesGrid()
        If lnkbtnShowHideNotes.Text.ToLower.Contains("hide") Then
            lnkbtnShowHideNotes.Text = "show notes"
            trNotes.Visible = False
        Else
            lnkbtnShowHideNotes.Text = "hide notes"
            trNotes.Visible = True
        End If
    End Sub
    
    Protected Sub gvNotes_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs)
        Dim gv As GridView = sender
        gv.PageIndex = e.NewPageIndex
        Call GetNotes()
    End Sub

    Protected Sub GetNotes()
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataTable As New DataTable
        Dim oAdapter As New SqlDataAdapter("spASPNET_CalendarManaged_GetEventNotes", oConn)

        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure

        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@EventId", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@EventId").Value = pnEventId
        
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerVisibleOnly", SqlDbType.Bit))
        oAdapter.SelectCommand.Parameters("@CustomerVisibleOnly").Value = 1
        
        Try
            oConn.Open()
            oAdapter.Fill(oDataTable)
            gvNotes.DataSource = oDataTable
            gvNotes.DataBind()
        Catch ex As Exception
            WebMsgBox.Show(ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub lnkbtnRefreshNotes_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call GetNotes()
        If lnkbtnShowHideNotes.Text.ToLower.Contains("show") Then
            Call ToggleNotesGrid()
        End If
    End Sub
    
    Protected Sub gvNotes_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs)
        Dim gvrea As GridViewRowEventArgs = e
        Dim row As GridViewRow = gvrea.Row
        If row.Cells.Count >= 3 Then
            row.Cells(3).Visible = False  ' hide Customer Visible flag
        End If
    End Sub

    Protected Sub cbDifferentCollectionAddress_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim cb As CheckBox = sender
        If cb.Checked Then
            trCollection1.Visible = True
            trCollection2.Visible = True
            rfvCollectionAddress1.Enabled = True
            rfvCollectionTown.Enabled = True
            rfvCollectionPostCode.Enabled = True
        Else
            trCollection1.Visible = False
            trCollection2.Visible = False
            tbCollectionAddress1.Text = String.Empty
            tbCollectionAddress2.Text = String.Empty
            tbCollectionTown.Text = String.Empty
            tbCollectionPostcode.Text = String.Empty
            rfvCollectionAddress1.Enabled = False
            rfvCollectionTown.Enabled = False
            rfvCollectionPostCode.Enabled = False
        End If
    End Sub
    
    Protected Sub lnkbtnCMAddSecondContact_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call SetContact2FieldsVisibility(True)
        tbCMContactName2.Focus()
    End Sub

    Protected Sub lnkbtnCMAddressOutsideUK_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        trCMCountry.Visible = True
        ddlCMCountry.SelectedIndex = 0
        ddlCMCountry.Focus()
    End Sub

    Protected Sub lnkbtnCMCollectionAddressOutsideUK_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        trCMCollectionCountry.Visible = True
        ddlCMCollectionCountry.SelectedIndex = 0
        ddlCMCollectionCountry.Focus()
    End Sub
    
    Protected Sub lnkbtnCMRemoveSecondContact_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        tbCMContactName2.Text = String.Empty
        tbCMContactPhone2.Text = String.Empty
        tbCMContactMobile2.Text = String.Empty
        Call SetContact2FieldsVisibility(False)
    End Sub
    
    Protected Sub SetContact2FieldsVisibility(ByVal bVisibility As Boolean)
        lnkbtnCMAddSecondContact.Visible = Not bVisibility
        rfvCMContactName2.Visible = bVisibility
        lblLegendCMContactName2.Visible = bVisibility
        tbCMContactName2.Visible = bVisibility
        rfvCMContactMobile2.Visible = bVisibility
        lblLegendCMContactMobile2.Visible = bVisibility
        tbCMContactMobile2.Visible = bVisibility
        trCMContactPhone2.Visible = bVisibility
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

    Protected Sub InitCountryDropdowns()
        If ddlCMCountry.Items.Count = 0 Or ddlCMCollectionCountry.Items.Count = 0 Then
            Dim sSQL As String = "SELECT SUBSTRING(CountryName,1,25) 'CountryName', CountryKey FROM Country WHERE DeletedFlag = 0 ORDER BY CountryName"
            Dim oListItemCollection As ListItemCollection = ExecuteQueryToListItemCollection(sSQL, "CountryName", "CountryKey")
            ddlCMCountry.Items.Clear()
            ddlCMCollectionCountry.Items.Clear()
            ddlCMCountry.Items.Add(New ListItem("- please select -", 0))
            ddlCMCollectionCountry.Items.Add(New ListItem("- please select -", 0))
            For Each li As ListItem In oListItemCollection
                ddlCMCountry.Items.Add(li)
                ddlCMCollectionCountry.Items.Add(li)
            Next
        End If
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
            WebMsgBox.Show("Error in ExecuteQueryToListItemCollection executing: " & sQuery & " : " & ex.Message)
        Finally
            oConn.Close()
        End Try
        ExecuteQueryToListItemCollection = oListItemCollection
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
            WebMsgBox.Show("Error in ExecuteNonQuery executing " & sQuery & " : " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Function
    
    Property pnEventId() As Integer
        Get
            Dim o As Object = ViewState("EV_EventId")
            If o Is Nothing Then
                Return 0
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Integer)
            ViewState("EV_EventId") = Value
        End Set
    End Property
    
</script>

<html xmlns=" http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>Event</title>
    <link rel="stylesheet" type="text/css" href="~/Reports.css" />
    <link href="elog.css" rel="STYLESHEET" type="text/css" />
    <link href="tabs.css" rel="STYLESHEET" type="text/css" />
    <link href="CS_Style.css" rel="stylesheet" type="text/css" />
</head>
<body style="font-family: Verdana">
    <form id="form1" runat="server">
    <div>
        <asp:Panel ID="pnlEvent" runat="server" Width="100%" Visible="False">
        <table style="width: 100%">
            <tr>
                <td style="width: 1%">
                </td>
                <td align="right" style="width: 16%">
                </td>
                <td style="width: 33%">
                </td>
                <td style="width: 16%">
                </td>
                <td style="width: 33%">
                </td>
                <td style="width: 1%">
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td colspan="2">
                    &nbsp;<asp:Label ID="lblLegendEvent" runat="server" Font-Bold="True" Text="Event Details:"></asp:Label></td>
                <td>
                </td>
                <td style="width: 450px">
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td align="right">
                    <asp:Label ID="Label2" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Event Name:"></asp:Label></td>
                <td>
                    <asp:Label ID="lblEventName" runat="server" Font-Bold="True" Font-Names="Verdana"
                        Font-Size="X-Small"></asp:Label></td>
                <td colspan="2">
                    <asp:Label ID="Label16" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="Booked by"></asp:Label>
                    <asp:Label ID="lblBookedBy" runat="server" Font-Names="Verdana" Font-Size="XX-Small"></asp:Label><asp:Label
                        ID="Label17" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="&nbsp;on"></asp:Label>
                    <asp:Label ID="lblBookedOn" runat="server" Font-Names="Verdana" Font-Size="XX-Small"></asp:Label></td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td align="right">
                    <asp:RequiredFieldValidator ID="rfvContactName" runat="server" ControlToValidate="tbContactName"
                        ErrorMessage="#" ValidationGroup="CalendarManaged" Width="8px" 
                        Font-Names="Verdana" Font-Size="XX-Small"></asp:RequiredFieldValidator>
                    <asp:Label ID="Label3" runat="server" Font-Names="Verdana" Font-Size="XX-Small" ForeColor="Red"
                        Text="Contact Name:"></asp:Label></td>
                <td>
                    <asp:TextBox ID="tbContactName" runat="server" BackColor="LightGoldenrodYellow" Font-Names="Verdana"
                        Font-Size="XX-Small" MaxLength="50" Width="300px"></asp:TextBox></td>
                <td colspan="2">
                    <asp:Label ID="Label20" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="Delivery date:"></asp:Label>
                    <asp:Label ID="lblDeliveryDate" runat="server" Font-Bold="True" Font-Names="Verdana"
                        Font-Size="XX-Small"></asp:Label>
                    &nbsp;
                    <asp:Label ID="Label26" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="Collection date:"/>
                    <asp:Label ID="lblCollectionDate" runat="server" Font-Bold="True" Font-Names="Verdana"
                        Font-Size="XX-Small"/>
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td align="right">
                    <asp:RequiredFieldValidator ID="rfvContactPhone" runat="server" ControlToValidate="tbContactPhone"
                        ErrorMessage="#" ValidationGroup="CalendarManaged" Font-Names="Verdana" 
                        Font-Size="XX-Small"/>
                    <asp:Label ID="Label4" runat="server" Font-Names="Verdana" Font-Size="XX-Small" ForeColor="Red"
                        Text="Contact Phone:"/>
                </td>
                <td>
                    <asp:TextBox ID="tbContactPhone" runat="server" BackColor="LightGoldenrodYellow"
                        Font-Names="Verdana" Font-Size="XX-Small" MaxLength="50" Width="300px"/>
                </td>
                <td align="right">
                    <asp:Label ID="Label5" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Contact Mobile:"/>
                </td>
                <td style="width: 450px">
                    <asp:TextBox ID="tbContactMobile" runat="server" BackColor="LightGoldenrodYellow"
                        Font-Names="Verdana" Font-Size="XX-Small" MaxLength="50" Width="300px"/>
                </td>
                <td>
                </td>
            </tr>
                <tr>
                    <td>
                        &nbsp;</td>
                    <td align="right">
                        <asp:RequiredFieldValidator ID="rfvCMContactName2" runat="server" ControlToValidate="tbCMContactName2" ErrorMessage="#" Font-Names="Verdana" Font-Size="XX-Small" ValidationGroup="CalendarManaged" Visible="False"/>
                        <asp:Label ID="lblLegendCMContactName2" runat="server" Font-Names="Verdana" Font-Size="XX-Small" ForeColor="Red" Text="Contact name 2:" Visible="False"/>
                    </td>
                    <td>
                        <asp:LinkButton ID="lnkbtnCMAddSecondContact" runat="server" Font-Names="Verdana" Font-Size="XX-Small" onclick="lnkbtnCMAddSecondContact_Click">add 2nd contact</asp:LinkButton>                    
                        <asp:TextBox ID="tbCMContactName2" runat="server" Font-Names="Verdana" 
                            Font-Size="XX-Small" MaxLength="50" Width="300px" Visible="False" 
                            BackColor="LightGoldenrodYellow" />
                    </td>
                    <td align="right">
                        <asp:RequiredFieldValidator ID="rfvCMContactMobile2" runat="server" ControlToValidate="tbCMContactMobile2" ErrorMessage="#" Font-Names="Verdana" Font-Size="XX-Small" ValidationGroup="CalendarManaged" />
                        <asp:Label ID="lblLegendCMContactMobile2" runat="server" Font-Names="Verdana" 
                            Font-Size="XX-Small" ForeColor="Red" Text="Contact mobile 2:"/>
                    </td>
                    <td style="width: 450px">
                        <asp:TextBox ID="tbCMContactMobile2" runat="server" Font-Names="Verdana" 
                            Font-Size="XX-Small" MaxLength="50" Width="300px" 
                            BackColor="LightGoldenrodYellow" />
                    </td>
                    <td>
                        &nbsp;
                    </td>
                </tr>
                <tr ID="trCMContactPhone2" runat="server" visible="false">
                    <td>
                        &nbsp;</td>
                    <td align="right">
                        <asp:RequiredFieldValidator ID="rfvCMContactPhone2" runat="server" ControlToValidate="tbCMContactPhone2" ErrorMessage="#" Font-Names="Verdana" Font-Size="XX-Small" ValidationGroup="CalendarManaged" />
                        <asp:Label ID="Label15axa0" runat="server" Font-Names="Verdana" Font-Size="XX-Small" ForeColor="Red" Text="Contact phone 2:"/>
                    </td>
                    <td>
                        <asp:TextBox ID="tbCMContactPhone2" runat="server" Font-Names="Verdana" 
                            Font-Size="XX-Small" MaxLength="50" Width="300px" 
                            BackColor="LightGoldenrodYellow" />
                    </td>
                    <td align="right">
                        &nbsp;</td>
                    <td style="width: 450px">
                        <asp:LinkButton ID="lnkbtnCMRemoveSecondContact" runat="server" 
                            OnClientClick='return confirm("Are you sure you want to remove the 2nd contact?");' 
                            onclick="lnkbtnCMRemoveSecondContact_Click" Font-Names="Verdana" 
                            Font-Size="XX-Small">remove 2nd contact</asp:LinkButton>
                    </td>
                    <td>
                        &nbsp;</td>
                </tr>
            <tr>
                <td>
                </td>
                <td align="right">
                    <asp:RequiredFieldValidator ID="rfvEventAddress1" runat="server" ControlToValidate="tbEventAddress1"
                        ErrorMessage="#" ValidationGroup="CalendarManaged" Font-Names="Verdana" 
                        Font-Size="XX-Small"/>
                    <asp:Label ID="Label6" runat="server" Font-Names="Verdana" Font-Size="XX-Small" ForeColor="Red"
                        Text="Event Address 1:"/>
                </td>
                <td colspan="3">
                    <asp:TextBox ID="tbEventAddress1" runat="server" BackColor="LightGoldenrodYellow"
                        Font-Names="Verdana" Font-Size="XX-Small" MaxLength="50" Width="400px"/>
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td align="right">
                    <asp:Label ID="Label7" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Event Address 2:"/>
                </td>
                <td>
                    <asp:TextBox ID="tbEventAddress2" runat="server" BackColor="LightGoldenrodYellow"
                        Font-Names="Verdana" Font-Size="XX-Small" MaxLength="50" Width="400px"/>
                </td>
                <td align="right">
                    <asp:Label ID="Label8" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Event Address 3:"/>
                </td>
                <td>
                    <asp:TextBox ID="tbEventAddress3" runat="server" BackColor="LightGoldenrodYellow"
                        Font-Names="Verdana" Font-Size="XX-Small" MaxLength="50" Width="400px"/>
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td align="right">
                    <asp:RequiredFieldValidator ID="rfvTown" runat="server" ControlToValidate="tbTown"
                        ErrorMessage="#" ValidationGroup="CalendarManaged" Font-Names="Verdana" 
                        Font-Size="XX-Small"/>
                    <asp:Label ID="Label9" runat="server" Font-Names="Verdana" Font-Size="XX-Small" ForeColor="Red"
                        Text="Town:"/>
                </td>
                <td>
                    <asp:TextBox ID="tbTown" runat="server" BackColor="LightGoldenrodYellow" Font-Names="Verdana"
                        Font-Size="XX-Small" MaxLength="50" Width="300px"/>
                </td>
                <td align="right">
                    <asp:RequiredFieldValidator ID="rfvPostCode" runat="server" ControlToValidate="tbPostcode"
                        ErrorMessage="#" ValidationGroup="CalendarManaged" Font-Names="Verdana" 
                        Font-Size="XX-Small"/>
                    <asp:Label ID="Label18" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        ForeColor="Red" Text="Postcode:"/>
                </td>
                <td style="width: 450px">
                    <asp:TextBox ID="tbPostcode" runat="server" BackColor="LightGoldenrodYellow" Font-Names="Verdana"
                        Font-Size="XX-Small" MaxLength="50" Width="150px"/>
                    &nbsp;<asp:LinkButton ID="lnkbtnCMAddressOutsideUK" runat="server" 
                        Font-Names="Verdana" Font-Size="XX-Small" 
                        onclick="lnkbtnCMAddressOutsideUK_Click">addr outside UK</asp:LinkButton>
                </td>
                <td>
                </td>
            </tr>
                <tr id="trCMCountry" runat="server" visible="false">
                    <td>
                        &nbsp;</td>
                    <td align="right">
                        <asp:RequiredFieldValidator ID="rfvCMCountry" runat="server" 
                            ControlToValidate="ddlCMCountry" ErrorMessage="#" Font-Names="Verdana" 
                            Font-Size="XX-Small" InitialValue="0" ValidationGroup="CalendarManaged" />
                        <asp:Label ID="Label38axa0" runat="server" Font-Names="Verdana" 
                            Font-Size="XX-Small" ForeColor="Red" Text="Country:" />
                    </td>
                    <td>
                        <asp:DropDownList ID="ddlCMCountry" runat="server" Font-Names="Verdana" 
                            Font-Size="XX-Small" ForeColor="Navy" TabIndex="8" Width="250px">
                        </asp:DropDownList>
                    </td>
                    <td align="right">
                        &nbsp;</td>
                    <td style="width: 450px">
                        &nbsp;</td>
                    <td>
                        &nbsp;</td>
                </tr>
            <tr>
                <td>
                </td>
                <td align="right">
                    <asp:Label ID="Label10" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        ForeColor="Red" Text="Delivery Time:"/>
                </td>
                <td>
                    <asp:DropDownList ID="ddlDeliveryTime" runat="server" BackColor="LightGoldenrodYellow"
                        Font-Names="Verdana" Font-Size="XX-Small">
                        <asp:ListItem>9.00am</asp:ListItem>
                        <asp:ListItem>10.30am</asp:ListItem>
                        <asp:ListItem>12.00 noon</asp:ListItem>
                    </asp:DropDownList>
                </td>
                <td align="right">
                    <asp:RequiredFieldValidator ID="rfvPreciseDeliveryPoint" runat="server" ControlToValidate="tbPreciseDeliveryPoint"
                        ErrorMessage="#" ValidationGroup="CalendarManaged" Font-Names="Verdana" 
                        Font-Size="XX-Small"/>
                    <asp:Label ID="Label12" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        ForeColor="Red" Text="Exact Delivery Point:"/>
                </td>
                <td style="width: 450px">
                    <asp:TextBox ID="tbPreciseDeliveryPoint" runat="server" BackColor="LightGoldenrodYellow"
                        Font-Names="Verdana" Font-Size="XX-Small" MaxLength="100" Width="100%"/>
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td align="right">
                    <asp:Label ID="Label13" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        ForeColor="Red" Text="Collection Time:"/>
                </td>
                <td>
                    <asp:DropDownList ID="ddlCollectionTime" runat="server" BackColor="LightGoldenrodYellow"
                        Font-Names="Verdana" Font-Size="XX-Small">
                        <asp:ListItem>9.00am - 10.00am</asp:ListItem>
                        <asp:ListItem>10.00am - 11.00am</asp:ListItem>
                        <asp:ListItem>11.00am - 12.00 noon</asp:ListItem>
                        <asp:ListItem>12.00 noon - 1.00pm</asp:ListItem>
                        <asp:ListItem>1.00pm - 2.00pm</asp:ListItem>
                        <asp:ListItem>2.00pm - 3.00pm</asp:ListItem>
                        <asp:ListItem>3.00pm - 4.00pm</asp:ListItem>
                        <asp:ListItem>4.00pm - 5.00pm</asp:ListItem>
                        <asp:ListItem>5.00pm - 6.00pm</asp:ListItem>
                        <asp:ListItem>Other - contact Transworld</asp:ListItem>
                    </asp:DropDownList>
                </td>
                <td align="right">
                    <asp:RequiredFieldValidator ID="rfvPreciseCollectionPoint" runat="server" ControlToValidate="tbPreciseCollectionPoint"
                        ErrorMessage="#" ValidationGroup="CalendarManaged" Font-Names="Verdana" 
                        Font-Size="XX-Small"/>
                    <asp:Label ID="Label14" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        ForeColor="Red" Text="Exact Collection Point:"/>
                </td>
                <td style="width: 450px">
                    <asp:TextBox ID="tbPreciseCollectionPoint" runat="server" BackColor="LightGoldenrodYellow"
                        Font-Names="Verdana" Font-Size="XX-Small" MaxLength="100" Width="100%"/>
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                    &nbsp;</td>
                <td align="right">
                    &nbsp;</td>
                <td>
                    <asp:CheckBox ID="cbDifferentCollectionAddress" runat="server" 
                        OnCheckedChanged="cbDifferentCollectionAddress_CheckedChanged" 
                        AutoPostBack="True" Text="collect from a different address" 
                        Font-Names="Verdana" Font-Size="XX-Small" />
                </td>
                <td align="right">
                    &nbsp;</td>
                <td style="width: 450px">
                    &nbsp;</td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr id="trCollection1" runat="server" visible="false">
                <td>
                    &nbsp;</td>
                <td align="right">
                    <asp:RequiredFieldValidator ID="rfvCollectionAddress1" runat="server" 
                        ControlToValidate="tbCollectionAddress1" ErrorMessage="#" 
                        ValidationGroup="CalendarManaged" Font-Names="Verdana" 
                        Font-Size="XX-Small"/>
                    <asp:Label ID="Label28" runat="server" Font-Names="Verdana" 
                        Font-Size="XX-Small" ForeColor="Red" Text="Collection Addr 1:"/>
                </td>
                <td>
                    <asp:TextBox ID="tbCollectionAddress1" runat="server" 
                        BackColor="LightGoldenrodYellow" Font-Names="Verdana" Font-Size="XX-Small" 
                        MaxLength="50" Width="400px"/>
                </td>
                <td align="right">
                    <asp:Label ID="Label27" runat="server" Font-Names="Verdana" 
                        Font-Size="XX-Small" Text="Collection Addr 2:"/>
                </td>
                <td style="width: 450px">
                    <asp:TextBox ID="tbCollectionAddress2" runat="server" 
                        BackColor="LightGoldenrodYellow" Font-Names="Verdana" Font-Size="XX-Small" 
                        MaxLength="50" Width="400px"/>
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr id="trCollection2" runat="server" visible="false">
                <td>
                    &nbsp;</td>
                <td align="right">
                    <asp:RequiredFieldValidator ID="rfvCollectionTown" runat="server" 
                        ControlToValidate="tbCollectionTown" ErrorMessage="#" 
                        ValidationGroup="CalendarManaged" Font-Names="Verdana" 
                        Font-Size="XX-Small"/>
                    <asp:Label ID="Label29" runat="server" Font-Names="Verdana" 
                        Font-Size="XX-Small" ForeColor="Red" Text="Collection Town:"/>
                </td>
                <td>
                    <asp:TextBox ID="tbCollectionTown" runat="server" 
                        BackColor="LightGoldenrodYellow" Font-Names="Verdana" Font-Size="XX-Small" 
                        MaxLength="50" Width="300px"/>
                </td>
                <td align="right">
                    <asp:RequiredFieldValidator ID="rfvCollectionPostCode" runat="server" 
                        ControlToValidate="tbCollectionPostCode" ErrorMessage="#" 
                        ValidationGroup="CalendarManaged" Font-Names="Verdana" 
                        Font-Size="XX-Small"/>
                    <asp:Label ID="Label30" runat="server" Font-Names="Verdana" 
                        Font-Size="XX-Small" ForeColor="Red" Text="Collection Postcode:"/>
                </td>
                <td style="width: 450px">
                    <asp:TextBox ID="tbCollectionPostcode" runat="server" 
                        BackColor="LightGoldenrodYellow" Font-Names="Verdana" Font-Size="XX-Small" 
                        MaxLength="50" Width="150px"/>
                    &nbsp;<asp:LinkButton ID="lnkbtnCMCollectionAddressOutsideUK" runat="server" 
                        Font-Names="Verdana" Font-Size="XX-Small" 
                        onclick="lnkbtnCMCollectionAddressOutsideUK_Click">addr outside UK</asp:LinkButton>
                </td>
                <td>
                    &nbsp;</td>
            </tr>
                <tr ID="trCMCollectionCountry" runat="server" visible="false">
                    <td>
                        &nbsp;</td>
                    <td align="right">
                        <asp:RequiredFieldValidator ID="rfvCMCollectionCountry" runat="server" ControlToValidate="ddlCMCollectionCountry" ErrorMessage="#" Font-Names="Verdana" Font-Size="XX-Small" InitialValue="0" ValidationGroup="CalendarManaged" />
                        <asp:Label ID="Label38axa1" runat="server" Font-Names="Verdana" Font-Size="XX-Small" ForeColor="Red" Text="Country:" />
                    </td>
                    <td>
                        <asp:DropDownList ID="ddlCMCollectionCountry" runat="server" Font-Names="Verdana" Font-Size="XX-Small" ForeColor="Navy" TabIndex="8" Width="250px"/>
                    </td>
                    <td align="right">
                        &nbsp;</td>
                    <td style="width: 450px">
                        &nbsp;</td>
                    <td>
                        &nbsp;</td>
                </tr>
            <tr>
                <td>
                    &nbsp;</td>
                <td align="right">
                    <asp:Label ID="Label31" runat="server" Font-Names="Verdana" 
                        Font-Size="XX-Small" Text="Customer Ref:"/>
                </td>
                <td colspan="3">
                    <asp:TextBox ID="tbCustomerReference" runat="server" 
                        BackColor="LightGoldenrodYellow" Font-Names="Verdana" Font-Size="XX-Small" 
                        MaxLength="50" Width="300px"/>
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td>
                </td>
                <td align="right">
                    <asp:Label ID="Label15" runat="server" Text="Special Instructions:" 
                        Font-Names="Verdana" Font-Size="XX-Small"></asp:Label></td>
                <td colspan="3">
                    <asp:TextBox ID="tbSpecialInstructions" runat="server" 
                        BackColor="LightGoldenrodYellow" Font-Names="Verdana" Font-Size="XX-Small" 
                        MaxLength="180" Width="99%"></asp:TextBox>
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td align="right">
                    <asp:Label ID="lblLegendProduct" runat="server" Text="Product:" 
                        Font-Names="Verdana" Font-Size="X-Small"></asp:Label>
                </td>
                <td colspan="3">
                    <asp:GridView ID="gvItems" runat="server" CellPadding="2" Width="100%" 
                        Font-Names="Verdana" Font-Size="XX-Small">
                    </asp:GridView>
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td align="right">
                </td>
                <td colspan="3">
                    <asp:Button ID="btnSaveChanges" runat="server" OnClick="btnSaveChanges_Click" 
                        Text="save changes" />
                    <asp:Label ID="lblOnlineChangesMessage" runat="server" 
                        Text="No changes can be accepted online as there are 5 days or fewer remaining until delivery. Contact Customer Services to request changes." 
                        Visible="False" Font-Names="Verdana" Font-Size="XX-Small" ForeColor="Red"></asp:Label>
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td colspan="4">
                    <hr />
                </td>
                <td>
                </td>
            </tr>
            <tr ID="trNotes" runat="server">
                <td>
                </td>
                <td align="right">
                </td>
                <td colspan="3">
                    <asp:Label ID="Label25" runat="server" Font-Bold="True" Font-Names="Verdana" 
                        Font-Size="XX-Small" Text="Notes:"></asp:Label>
                    <asp:GridView ID="gvNotes" runat="server" AllowPaging="True" CellPadding="2" OnPageIndexChanging="gvNotes_PageIndexChanging" OnRowDataBound="gvNotes_RowDataBound" PageSize="6" Width="100%" Font-Names="Verdana" Font-Size="XX-Small">
                        <EmptyDataTemplate>
                            no notes
                        </EmptyDataTemplate>
                        <PagerStyle Font-Names="Verdana" Font-Size="Small" HorizontalAlign="Center" />
                    </asp:GridView>
                    &nbsp; </td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td align="right">
                </td>
                <td colspan="3">
                    <asp:LinkButton ID="lnkbtnShowHideNotes" runat="server" 
                        OnClick="lnkbtnShowHideNotes_Click" Font-Names="Verdana" 
                        Font-Size="XX-Small">hide notes</asp:LinkButton>
                    <asp:LinkButton ID="lnkbtnRefreshNotes" runat="server" 
                        OnClick="lnkbtnRefreshNotes_Click" Font-Names="Verdana" 
                        Font-Size="XX-Small">refresh notes</asp:LinkButton>
                </td>
                <td>
                </td>
            </tr>
        </table>
        </asp:Panel>
        <asp:Panel ID="pnlMessage" runat="server" Width="100%" Visible="True">
            <asp:Label ID="lblMessage" runat="server" Font-Names="Verdana" Font-Size="X-Small"/><br />
            <br />
            &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
            <asp:Button ID="btnClose" runat="server" OnClientClick="javascript:window.close();"
                Text="close" />
        &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
        &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
        </asp:Panel>
        <br />
        <asp:HyperLink ID="HyperLink1" runat="server" 
            NavigateUrl="~/event.aspx?GUID=cfd60842-496e-4656-a8b4-ff33e3" Visible="False">Call myself</asp:HyperLink></div>
    </form>
</body>
</html>