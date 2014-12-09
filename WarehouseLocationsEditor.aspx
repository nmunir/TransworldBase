<%@ Page Language="VB" Theme="AIMSDefault" ValidateRequest="false" %>

<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="Telerik.Web.UI" %>
<%@ Register Assembly="Telerik.Web.UI" Namespace="Telerik.Web.UI" TagPrefix="telerik" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<script runat="server">

    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsNumeric(Session("UserKey")) Then
            Response.RedirectLocation = "http:/my.transworld.eu.com/common/session_expired.aspx"
            Server.Transfer("session_expired.aspx")
        End If
        If Not IsPostBack Then
            Call PopulateWarehouseDropdown()
            ddlWarehouse.Focus()
        End If
        Call SetTitle()
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
        Page.Header.Title = sTitle & "Locations Editor"
    End Sub
   
    Protected Sub PopulateWarehouseDropdown()
        Dim sSQL As String = "SELECT WarehouseID, WarehouseKey FROM Warehouse WHERE DeletedFlag = 'N' ORDER BY WarehouseID"
        Dim dtWarehouses As DataTable = ExecuteQueryToDataTable(sSQL)
        ddlWarehouse.Items.Clear()
        ddlWarehouseBSE.Items.Clear()
        ddlWarehouse.Items.Add(New ListItem("- please select -", 0))
        ddlWarehouseBSE.Items.Add(New ListItem("- please select -", 0))
        For Each drWarehouse As DataRow In dtWarehouses.Rows
            ddlWarehouse.Items.Add(New ListItem(drWarehouse("WarehouseID"), drWarehouse("WarehouseKey")))
            ddlWarehouseBSE.Items.Add(New ListItem(drWarehouse("WarehouseID"), drWarehouse("WarehouseKey")))
        Next
    End Sub
    
    Protected Sub btnTestCreateLocations_Click(sender As Object, e As System.EventArgs)
        Dim sResult As String = CreateLocations(bTestOnly:=True)
    End Sub

    Protected Sub btnCreateLocations_Click(sender As Object, e As System.EventArgs)
        Dim sResult As String = CreateLocations(bTestOnly:=False)
    End Sub
    
    Protected Function bIsValidLocationID(ByRef sLocationID As String) As Boolean
        Return Regex.IsMatch(sLocationID, "^[a-zA-Z0-9\-_~\s]*$")
    End Function
    
    Protected Function LocationExists(nWarehouseKey As Int32, sWarehouseRackId As String, sWarehouseSectionId As String, sWarehouseBayId As String) As Boolean
        LocationExists = False
        Dim dtEntity As DataTable
        'Dim nWarehouseKey As Int32 = ExecuteQueryToDataTable("SELECT WarehouseKey FROM Warehouse WHERE WarehouseID = '" & sWarehouseId & "'").Rows(0).Item(0)
        dtEntity = ExecuteQueryToDataTable("SELECT WarehouseRackKey FROM WarehouseRack WHERE WarehouseKey = " & nWarehouseKey & " AND WarehouseRackId = '" & sWarehouseRackId & "'")
        If dtEntity.Rows.Count > 0 Then
            Dim nWarehouseRackKey As Int32 = dtEntity.Rows(0).Item(0)
            dtEntity = ExecuteQueryToDataTable("SELECT WarehouseSectionKey FROM WarehouseSection WHERE WarehouseRackKey = " & nWarehouseRackKey & " AND WarehouseSectionId = '" & sWarehouseSectionId & "'")
            If dtEntity.Rows.Count > 0 Then
                Dim nWarehouseSectionKey As Int32 = dtEntity.Rows(0).Item(0)
                dtEntity = ExecuteQueryToDataTable("SELECT WarehouseBayKey FROM WarehouseBay WHERE WarehouseSectionKey = " & nWarehouseSectionKey & " AND WarehouseBayId = '" & sWarehouseBayId & "'")
                If dtEntity.Rows.Count > 0 Then
                    LocationExists = True
                End If
            End If
        End If
    End Function
    
    Protected Function CreateLocations(bTestOnly As Boolean) As String
        ' also need to check for duplicates in lists
        Dim sSQL As String = String.Empty
        CreateLocations = String.Empty
        Dim bLocationsExist As Boolean = False
        Dim bDuplicatesFound As Boolean = False
        Dim bInvalidLocationIDFound As Boolean = False
        Dim bOverlengthLocationIDFound As Boolean = False
        Dim lstRacks As List(Of String) = LocationList(tbRacks)
        Dim lstSections As List(Of String) = LocationList(tbSections)
        Dim lstBays As List(Of String) = LocationList(tbBays)
        
        If lstRacks.Count = 0 Then
            WebMsgBox.Show("You must define at least one rack.")
            Exit Function
        End If
        If lstSections.Count = 0 Then
            WebMsgBox.Show("You must define at least one section.")
            Exit Function
        End If
        If lstBays.Count = 0 Then
            WebMsgBox.Show("You must define at least one bay.")
            Exit Function
        End If
        
        Dim ssRacks As New SortedSet(Of String)
        Dim ssSections As New SortedSet(Of String)
        Dim ssBays As New SortedSet(Of String)
        
        For Each sRack As String In lstRacks
            If Not bIsValidLocationID(sRack) Then
                JournalAdd("Detected invalid character(s) in rack name (only numbers & letters allowed): " & sRack)
                bInvalidLocationIDFound = True
            End If
            If sRack.Length > 20 Then
                JournalAdd("Detected overlength rack name (max length 20 chars): " & sRack)
                bOverlengthLocationIDFound = True
            End If
            Try
                ssRacks.Add(sRack)
            Catch ex As Exception
                JournalAdd("Detected duplicate rack entry: " & sRack)
                bDuplicatesFound = True
            End Try
            For Each sSection As String In lstSections
                If Not bIsValidLocationID(sSection) Then
                    JournalAdd("Detected invalid character(s) in section name (only numbers & letters allowed): " & sSection)
                    bInvalidLocationIDFound = True
                End If
                If sSection.Length > 20 Then
                    JournalAdd("Detected overlength section name (max length 20 chars): " & sSection)
                    bOverlengthLocationIDFound = True
                End If
                Try
                    ssSections.Add(sSection)
                Catch ex As Exception
                    JournalAdd("Detected duplicate section entry: " & sRack)
                    bDuplicatesFound = True
                End Try
                For Each sBay As String In lstBays
                    If Not bIsValidLocationID(sBay) Then
                        JournalAdd("Detected invalid character(s) in bay name (only numbers & letters allowed): " & sBay)
                        bInvalidLocationIDFound = True
                    End If
                    If sBay.Length > 20 Then
                        JournalAdd("Detected overlength bay name (max length 20 chars): " & sBay)
                        bOverlengthLocationIDFound = True
                    End If
                    Try
                        ssBays.Add(sSection)
                    Catch ex As Exception
                        JournalAdd("Detected duplicate bay entry: " & sBay)
                        bDuplicatesFound = True
                    End Try
                    If LocationExists(ddlWarehouse.SelectedValue, sRack, sSection, sBay) Then
                        JournalAdd("Location already exists: " & ddlWarehouse.SelectedItem.Text & "," & sRack & ", " & sSection & ", " & sBay)
                        bLocationsExist = True
                    End If
                Next
            Next
        Next
        JournalAdd("Validation complete.")
        JournalAdd("")
        
        If bLocationsExist Then
            WebMsgBox.Show("One or more locations already exist - see Journal for details.")
        ElseIf bDuplicatesFound Then
            WebMsgBox.Show("Duplicate names detected - see Journal for details.")
        ElseIf bInvalidLocationIDFound Then
            WebMsgBox.Show("Detected invalid character(s) in location name - see Journal for details.")
        ElseIf bOverlengthLocationIDFound Then
            WebMsgBox.Show("Detected overlength location name - see Journal for details.")
        Else
            For Each sRack As String In lstRacks
                If bTestOnly Then
                    JournalAdd("In warehouse " & ddlWarehouse.SelectedItem.Text & ", would have created new rack: " & sRack)
                Else
                    sSQL = "INSERT WarehouseRack(WarehouseKey, WarehouseRackId, DeletedFlag, LastUpdatedByKey, LastUpdatedOn) VALUES (" & ddlWarehouse.SelectedValue & ", '" & sRack & "', 'N', " & Session("UserKey") & ", GETDATE())"
                    Call ExecuteQueryToDataTable(sSQL)
                    JournalAdd("In warehouse " & ddlWarehouse.SelectedItem.Text & ", created new rack: " & sRack)
                End If
                Dim nRackKey As Int32 = GetRackKeyFromRackID(ddlWarehouse.SelectedValue, sRack)
                For Each sSection As String In lstSections
                    If bTestOnly Then
                        JournalAdd("In warehouse " & ddlWarehouse.SelectedItem.Text & ", rack " & sRack & ", would have created new section: " & sSection)
                    Else
                        sSQL = "INSERT WarehouseSection (WarehouseRackKey, WarehouseSectionId, DeletedFlag, LastUpdatedByKey, LastUpdatedOn) VALUES (" & nRackKey & ", '" & sSection & "', 'N', " & Session("UserKey") & ", GETDATE())"
                        Call ExecuteQueryToDataTable(sSQL)
                        JournalAdd("In warehouse " & ddlWarehouse.SelectedItem.Text & ", rack " & sRack & ", created new section: " & sSection)
                    End If
                    Dim nSectionKey As Int32 = GetSectionKeyFromSectionID(nRackKey, sSection)
                    For Each sBay As String In lstBays
                        If bTestOnly Then
                            JournalAdd("In warehouse " & ddlWarehouse.SelectedItem.Text & ", rack " & sRack & ", section " & sSection & ", would have created new bay: " & sBay & ", bay size " & ddlBaySize.SelectedItem.Text.Replace("- undefined -", "undefined"))
                        Else
                            sSQL = "INSERT WarehouseBay(WarehouseSectionKey, WarehouseBayId, WarehouseBaySize, DeletedFlag, LastUpdatedByKey, LastUpdatedOn) VALUES (" & nSectionKey & ", '" & sBay & "', " & ddlBaySize.SelectedValue & ", 'N', " & Session("UserKey") & ", GETDATE())"
                            Call ExecuteQueryToDataTable(sSQL)
                            JournalAdd("In warehouse " & ddlWarehouse.SelectedItem.Text & ", rack " & sRack & ", section " & sSection & ", created new bay: " & sBay & ", bay size " & ddlBaySize.SelectedItem.Text.Replace("- undefined -", "undefined"))
                        End If
                    Next
                Next
            Next
            If bTestOnly Then
                JournalAdd("Test complete.")
                JournalAdd("")
            Else
                JournalAdd("Locations created.")
                JournalAdd("")
            End If
        End If
    End Function
    
    Protected Function GetRackKeyFromRackID(nWarehouseKey As Int32, sRackID As String) As Int32
        GetRackKeyFromRackID = 0
        Try
            GetRackKeyFromRackID = ExecuteQueryToDataTable("SELECT WarehouseRackKey FROM WarehouseRack WHERE WarehouseKey = " & nWarehouseKey & " AND WarehouseRackId = '" & sRackID & "'").Rows(0).Item(0)
        Catch ex As Exception
        End Try
    End Function
    
    Protected Function GetSectionKeyFromSectionID(nRackKey As Int32, sSectionID As String) As Int32
        GetSectionKeyFromSectionID = 0
        Try
            GetSectionKeyFromSectionID = ExecuteQueryToDataTable("SELECT WarehouseSectionKey FROM WarehouseSection WHERE WarehouseRackKey = " & nRackKey & " AND WarehouseSectionId = '" & sSectionID & "'").Rows(0).Item(0)
        Catch ex As Exception
        End Try
    End Function

    Protected Function LocationList(tbTextbox As TextBox) As List(Of String)
        Dim lstLocationList As New List(Of String)
        Dim sLocations As String = tbTextbox.Text
        Dim arrLocations() As String
        sLocations = sLocations.Trim
        sLocations = sLocations.Replace(",", " ")
        sLocations = sLocations.Replace("  ", " ")
        arrLocations = sLocations.Split(" ")
        For Each sLocation In arrLocations
            If sLocation <> String.Empty Then
                lstLocationList.Add(sLocation.Replace("~", " "))
            End If
        Next
        LocationList = lstLocationList
    End Function
    
    Protected Sub JournalAdd(sJournalEntry As String)
        tbJournal.Text = tbJournal.Text & sJournalEntry & Environment.NewLine
    End Sub
    
    Protected Sub ddlWarehouse_SelectedIndexChanged(sender As Object, e As System.EventArgs)
        Dim ddl As DropDownList = sender
        If ddl.SelectedValue > 0 Then
            btnTestCreateLocations.Enabled = True
            btnCreateLocations.Enabled = True
        Else
            btnTestCreateLocations.Enabled = False
            btnCreateLocations.Enabled = False
        End If
    End Sub
    
    Protected Sub ddlWarehouseBSE_SelectedIndexChanged(sender As Object, e As System.EventArgs)
        Call InitRackDropdown()
        Call ClearSectionDropdown()
        Call ClearBayDropdown()
        Call ClearSizeDropdown()
        ddlSizeBSE.SelectedIndex = 0
        lblLegendSaved.Visible = False
        ddlRackBSE.Focus()
    End Sub

    Protected Sub InitRackDropdown()
        ddlRackBSE.Items.Clear()
        Dim sSQL As String = "SELECT WarehouseRackId, WarehouseRackKey FROM WarehouseRack WHERE DeletedFlag = 'N' AND WarehouseKey = " & ddlWarehouseBSE.SelectedValue & " ORDER BY WarehouseRackId"
        Dim oListItemCollection As ListItemCollection = ExecuteQueryToListItemCollection(sSQL, "WarehouseRackId", "WarehouseRackKey")
        ddlRackBSE.Items.Add(New ListItem("- please select -", 0))
        For Each li As ListItem In oListItemCollection
            ddlRackBSE.Items.Add(li)
        Next
    End Sub
    
    Protected Sub ddlRackBSE_SelectedIndexChanged(sender As Object, e As System.EventArgs)
        Call InitSectionDropdown()
        Call ClearBayDropdown()
        Call ClearSizeDropdown()
        ddlSizeBSE.SelectedIndex = 0
        lblLegendSaved.Visible = False
        btnSaveBaySize.Enabled = False
        ddlSectionBSE.Focus()
    End Sub

    Protected Sub InitSectionDropdown()
        ddlSectionBSE.Items.Clear()
        Dim sSQL As String = "SELECT WarehouseSectionId, WarehouseSectionKey FROM WarehouseSection WHERE DeletedFlag = 'N' AND WarehouseRackKey = " & ddlRackBSE.SelectedValue & " ORDER BY WarehouseSectionId"
        Dim oListItemCollection As ListItemCollection = ExecuteQueryToListItemCollection(sSQL, "WarehouseSectionId", "WarehouseSectionKey")
        ddlSectionBSE.Items.Add(New ListItem("- please select -", 0))
        For Each li As ListItem In oListItemCollection
            ddlSectionBSE.Items.Add(li)
        Next
    End Sub

    Protected Sub ddlSectionBSE_SelectedIndexChanged(sender As Object, e As System.EventArgs)
        Call InitBayDropdown()
        Call ClearSizeDropdown()
        ddlSizeBSE.SelectedIndex = 0
        lblLegendSaved.Visible = False
        btnSaveBaySize.Enabled = False
        ddlSectionBSE.Focus()
    End Sub

    Protected Sub InitBayDropdown()
        ddlBayBSE.Items.Clear()
        Dim sSQL As String = "SELECT WarehouseBayId, WarehouseBayKey FROM WarehouseBay WHERE DeletedFlag = 'N' AND WarehouseSectionKey = " & ddlSectionBSE.SelectedValue & " ORDER BY WarehouseBayId"
        Dim oListItemCollection As ListItemCollection = ExecuteQueryToListItemCollection(sSQL, "WarehouseBayId", "WarehouseBayKey")
        ddlBayBSE.Items.Add(New ListItem("- please select -", 0))
        For Each li As ListItem In oListItemCollection
            ddlBayBSE.Items.Add(li)
        Next
    End Sub

    Protected Sub ddlBayBSE_SelectedIndexChanged(sender As Object, e As System.EventArgs)
        Dim nBaySize As Int32 = ExecuteQueryToDataTable("SELECT ISNULL(WarehouseBaySize, 0) FROM WarehouseBay WHERE WarehouseBayKey = " & ddlBayBSE.SelectedValue).Rows(0).Item(0)
        For i As Int32 = 0 To ddlSizeBSE.Items.Count - 1
            If ddlSizeBSE.Items(i).Value = nBaySize Then
                ddlSizeBSE.SelectedIndex = i
                ddlSizeBSE.Enabled = True
                Exit For
            End If
        Next
        lblLegendSaved.Visible = False
        btnSaveBaySize.Enabled = False
        ddlSizeBSE.Focus()
    End Sub

    Protected Sub ClearRackDropdown()
        If ddlRackBSE.Items.Count > 0 Then
            'ddlRackBSE.SelectedIndex = 0
            ddlRackBSE.Items.Clear()
        End If
    End Sub
    
    Protected Sub ClearSectionDropdown()
        If ddlSectionBSE.Items.Count > 0 Then
            'ddlSectionBSE.SelectedIndex = 0
            ddlSectionBSE.Items.Clear()
        End If
    End Sub
    
    Protected Sub ClearBayDropdown()
        If ddlBayBSE.Items.Count > 0 Then
            'ddlBayBSE.SelectedIndex = 0
            ddlBayBSE.Items.Clear()
        End If
    End Sub
    
    Protected Sub ClearSizeDropdown()
        ddlSizeBSE.SelectedIndex = 0
        ddlSizeBSE.Enabled = False
    End Sub
    
    Protected Sub ddlSizeBSE_SelectedIndexChanged(sender As Object, e As System.EventArgs)
        btnSaveBaySize.Enabled = True
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

    Protected Sub btnSaveBaySize_Click(sender As Object, e As System.EventArgs)
        Dim sSQL As String
        sSQL = "UPDATE WarehouseBay SET WarehouseBaySize = " & ddlSizeBSE.SelectedValue & " WHERE WarehouseBayKey = " & ddlBayBSE.SelectedValue
        Call ExecuteQueryToDataTable(sSQL)
        lblLegendSaved.Visible = True
    End Sub
    
</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <style type="text/css">
        .style1
        {
            font-family: Arial, Helvetica, sans-serif;
            font-size: xx-small;
            padding-left: 10px;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
    <main:Header ID="ctlHeader" runat="server" />
    <asp:PlaceHolder ID="PlaceHolderForScriptManager" runat="server" />
    &nbsp;<asp:Label ID="lblLegendLocationsEditor" runat="server" Font-Names="Verdana"
        Font-Size="Small" Font-Bold="True" Text="Warehouse Locations Editor" />
    &nbsp;<asp:Panel ID="pnlUpdate" Width="100%" runat="server">
        &nbsp;<table style="width: 98%">
            <tr>
                <td style="width: 10%">
                </td>
                <td style="width: 25%">
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:Label ID="lblLegendWarehouse" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="Warehouse:"></asp:Label>
                </td>
                <td>
                    <asp:DropDownList ID="ddlWarehouse" runat="server" AutoPostBack="True" Font-Size="X-Small"
                        OnSelectedIndexChanged="ddlWarehouse_SelectedIndexChanged">
                    </asp:DropDownList>
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
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
                <td>
                    <asp:Label ID="lblLegendRacks" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="Racks:" />
                </td>
                <td>
                    <asp:TextBox ID="tbRacks" runat="server" TextMode="MultiLine" Width="100%" Rows="3"
                        Font-Names="Verdana" Font-Size="XX-Small"></asp:TextBox>
                </td>
                <td class="style1">
                    Enter the names of the racks you want to create, separated by space or comma. You
                    can use any alphanumeric (0-9, a-z, A-Z) character, plus hyphens and underscores.
                    To put a space within a name, eg LEFT WALL, enter LEFT~WALL, using the <b>tilde</b>
                    character &#39;<b>~</b>&#39; (which is normally found above the # on your keyboard)
                    to represent a space character.
                </td>
            </tr>
            <tr>
                <td>
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
                <td>
                    <asp:Label ID="lblLegendSections" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="Sections:" />
                </td>
                <td>
                    <asp:TextBox ID="tbSections" runat="server" TextMode="MultiLine" Width="100%" Rows="3"
                        Font-Names="Verdana" Font-Size="XX-Small" />
                </td>
                <td class="style1">
                    Enter the names of the racks you want to create in each of the racks above, separated
                    by space or comma. You can use any alphanumeric (0-9, a-z, A-Z) character, plus
                    hyphens and underscores. To put a space within a name, eg LEFT WALL, enter LEFT~WALL,
                    using the <b>tilde</b> character &#39;<b>~</b>&#39; (which is normally found above
                    the # on your keyboard) to represent a space character.
                </td>
            </tr>
            <tr>
                <td>
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
                <td>
                    <asp:Label ID="lblLegendBays" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="Bays:" />
                </td>
                <td>
                    <asp:TextBox ID="tbBays" runat="server" TextMode="MultiLine" Width="100%" Rows="3"
                        Font-Names="Verdana" Font-Size="XX-Small" />
                </td>
                <td class="style1">
                    Enter the names of the bays you want to create in each of the sections above, separated
                    by space or comma. You can use any alphanumeric (0-9, a-z, A-Z) character, plus
                    hyphens and underscores. To put a space within a name, eg LEFT WALL, enter LEFT~WALL,
                    using the <b>tilde</b> character &#39;<b>~</b>&#39; (which is normally found above
                    the # on your keyboard) to represent a space character.
                </td>
            </tr>
            <tr>
                <td>
                    <asp:Label ID="lblLegendBaysize" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="Bay Size:" />
                </td>
                <td>
                    <asp:DropDownList ID="ddlBaySize" runat="server" AutoPostBack="True" Font-Size="X-Small">
                        <asp:ListItem Value="0">- undefined -</asp:ListItem>
                        <asp:ListItem Value="50">0.5</asp:ListItem>
                        <asp:ListItem Value="100" Selected="True">1</asp:ListItem>
                        <asp:ListItem Value="150">1.5</asp:ListItem>
                        <asp:ListItem Value="200">2</asp:ListItem>
                        <asp:ListItem Value="250">2.5</asp:ListItem>
                        <asp:ListItem Value="300">3</asp:ListItem>
                        <asp:ListItem Value="350">3.5</asp:ListItem>
                        <asp:ListItem Value="400">4</asp:ListItem>
                        <asp:ListItem Value="450">4.5</asp:ListItem>
                        <asp:ListItem Value="500">5</asp:ListItem>
                    </asp:DropDownList>
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td>
                    &nbsp;
                </td>
                <td colspan="2">
                    &nbsp;<asp:Button ID="btnTestCreateLocations" runat="server" Enabled="False" Text="Test..."
                        Width="150px" OnClick="btnTestCreateLocations_Click" />
                    &nbsp;<asp:Button ID="btnCreateLocations" runat="server" Enabled="False" Text="Create Locations"
                        Width="150px" OnClick="btnCreateLocations_Click" />
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td>
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:Label ID="lblLegendJournal" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="Journal:"></asp:Label>
                    <br />
                </td>
                <td colspan="2">
                    <asp:TextBox ID="tbJournal" runat="server" Rows="10" TextMode="MultiLine" Width="100%"
                        Font-Names="Verdana" Font-Size="XX-Small"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td colspan="2">
                    <hr />
                </td>
            </tr>
            <tr>
                <td colspan="3">
                    <asp:Label ID="lblLegendLocationsEditor0" runat="server" Font-Bold="True" Font-Names="Verdana"
                        Font-Size="Small" Text="Bay Size Editor" />
                </td>
            </tr>
            <tr>
                <td>
                    &nbsp;
                </td>
                <td colspan="2">
                    <asp:Label ID="lblLegendWarehouseBSE" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="W:" />
                    <asp:DropDownList ID="ddlWarehouseBSE" runat="server" AutoPostBack="True" Font-Size="X-Small"
                        OnSelectedIndexChanged="ddlWarehouseBSE_SelectedIndexChanged" />
                    &nbsp;<asp:Label ID="lblLegendWarehouseRackBSE" runat="server" Font-Names="Verdana"
                        Font-Size="XX-Small" Text="R:" />
                    <asp:DropDownList ID="ddlRackBSE" runat="server" AutoPostBack="True" Font-Size="X-Small"
                        OnSelectedIndexChanged="ddlRackBSE_SelectedIndexChanged" />
                    <asp:Label ID="lblLegendWarehouseSectionBSE" runat="server" Font-Names="Verdana"
                        Font-Size="XX-Small" Text="S:" />
                    <asp:DropDownList ID="ddlSectionBSE" runat="server" AutoPostBack="True" Font-Size="X-Small"
                        OnSelectedIndexChanged="ddlSectionBSE_SelectedIndexChanged" />
                    <asp:Label ID="lblLegendWarehouseBayBSE" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="B:" />
                    <asp:DropDownList ID="ddlBayBSE" runat="server" AutoPostBack="True" Font-Size="X-Small"
                        OnSelectedIndexChanged="ddlBayBSE_SelectedIndexChanged" />
                </td>
            </tr>
            <tr>
                <td>
                    &nbsp;
                </td>
                <td>
                    <asp:Label ID="lblLegendSizeBSE" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="Size:" />
                    <asp:DropDownList ID="ddlSizeBSE" runat="server" AutoPostBack="True" Font-Size="X-Small"
                        OnSelectedIndexChanged="ddlSizeBSE_SelectedIndexChanged">
                        <asp:ListItem Value="0">- undefined -</asp:ListItem>
                        <asp:ListItem Value="50">0.5</asp:ListItem>
                        <asp:ListItem Selected="True" Value="100">1</asp:ListItem>
                        <asp:ListItem Value="150">1.5</asp:ListItem>
                        <asp:ListItem Value="200">2</asp:ListItem>
                        <asp:ListItem Value="250">2.5</asp:ListItem>
                        <asp:ListItem Value="300">3</asp:ListItem>
                        <asp:ListItem Value="350">3.5</asp:ListItem>
                        <asp:ListItem Value="400">4</asp:ListItem>
                        <asp:ListItem Value="450">4.5</asp:ListItem>
                        <asp:ListItem Value="500">5</asp:ListItem>
                    </asp:DropDownList>
                    &nbsp;
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td valign="middle">
                    <asp:Button ID="btnSaveBaySize" runat="server" Enabled="False" Text="Save" Width="200px"
                        OnClick="btnSaveBaySize_Click" />
                    &nbsp;<asp:Label ID="lblLegendSaved" runat="server" Font-Bold="True" Font-Names="Verdana"
                        Font-Size="Small" ForeColor="#009933" Text="saved" Visible="False"></asp:Label>
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                    &nbsp;
                </td>
                <td valign="middle">
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td>
                    &nbsp;</td>
                <td colspan="2" valign="middle">
                    <hr />
                </td>
            </tr>
            <tr style="font-family: Verdana; font-size: x-small">
                <td>
                    &nbsp;
                </td>
                <td colspan="2" valign="middle">
                    <b>INSTRUCTIONS (Locations Editor):<br />
                    </b>
                    <br />
                    1.&nbsp; Using the Warehouse drop down box, select the warehouse to which you want
                    to add locations.<br />
                    <br />
                    2.&nbsp; Enter the names of the Racks, Sections &amp; Bays you want to create.&nbsp;
                    Separate each name with a space or comma.<br />
                    <br />
                    Example:<br />
                    <br />
                    RACK_A&nbsp; RACK_B&nbsp; RACK_C<br />
                    SECTION_A&nbsp; SECTION_B&nbsp; SECTION_C<br />
                    BAY_A&nbsp; BAY_B&nbsp; BAY_C<br />
                    <br />
                    ...will create <b>27</b> locations within the warehouse you have selected:<br />
                    <br />
                    <i>RACK_A | SECTION_A | BAY_A<br />
                        RACK_A | SECTION_A | BAY_B<br />
                        RACK_A | SECTION_A | BAY_C<br />
                        RACK_A | SECTION_B | BAY_A<br />
                        RACK_A | SECTION_B | BAY_B<br />
                        RACK_A | SECTION_B | BAY_C<br />
                        RACK_A | SECTION_C | BAY_A<br />
                        RACK_A | SECTION_C | BAY_B<br />
                        RACK_A | SECTION_C | BAY_C<br />
                        RACK_B | SECTION_A | BAY_A<br />
                        RACK_B | SECTION_A | BAY_B</i><br />
                    <br />
                    etc.<br />
                    <br />
                    3.&nbsp; Select the Bay Size from the <b>Bay Size</b> drop down box.<br />
                    <br />
                    4.&nbsp; Click the <b>Test</b> button. The utility shows you what locations will
                    be created when you click the <b>Create Locations</b> button.<br />
                    <br />
                    5.&nbsp; Click the <b>Create Locations</b> button to create the locations.<br />
                    <br />
                    Check the <b>Journal</b> window to see progress and error messages.<br />
                    <br />
                    <b>
                        <br />
                        INSTRUCTIONS (Bay Size Editor):<br />
                    </b>
                    <br />
                    1.&nbsp; Select the bay you want to edit by using the <b>W</b>(arehouse), <b>R</b>(ack),
                    <b>S</b>(ection) and <b>B</b>(ay) drop down boxes.<br />
                    <br />
                    2.&nbsp; From the <b>Size</b> drop down box select the new size you want to assign to 
                    the bay.<br />
                    <br />
                    3.&nbsp; Click the <b>Save</b> button.<br />
                    <br />
                    [end]
                </td>
            </tr>
        </table>
    </asp:Panel>
    </form>
</body>
</html>
