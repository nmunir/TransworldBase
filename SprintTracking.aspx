<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>
<%@ import Namespace="System.Globalization" %>
<%@ import Namespace="System.Threading" %>
<%@ import Namespace="System.Collections.Generic" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">
    
    Dim gsConn As String = ConfigurationManager.AppSettings("AIMSRootConnectionString")
    Dim gnCustomerKey As Integer
    Dim dicCustomerFromGUID As New Dictionary(Of String, Integer)

    Dim gnRequestType As Integer
    Const CUSTOMER_ACCOUNT_CODE_ADOF As Integer = 185
    Const REQUEST_TYPE_AWB_DETAIL = 0
    Const REQUEST_TYPE_LAST_50 = 1
    Const REQUEST_TYPE_DATE_RANGE = 2
    Const REQUEST_TYPE_FROM_START_DATE = 3
    Const REQUEST_TYPE_OUTSTANDING_PODS = 4
    Const REQUEST_TYPE_SEARCH = 5
    
    Const MIN_SEARCH_STRING_LENGTH = 3
    
    Sub Page_Load(ByVal Source As Object, ByVal E As EventArgs)
        Dim bError As Boolean = False
        Dim sTitle As String = ConfigLib.GetConfigItem_AppTitle()

        Thread.CurrentThread.CurrentCulture = New CultureInfo("en-GB", False)

        gnRequestType = REQUEST_TYPE_AWB_DETAIL
        pnlConsignmentDetail.Visible = True
        pnlConsignments.Visible = True
        pnlParameterError.Visible = False

        If Request.QueryString("GUID") Is Nothing Then
            Call ReturnError("Missing GUID")
            Exit Sub
        End If
        
        Call InitCustomerFromGUIDDictionary()
        If dicCustomerFromGUID.ContainsKey(Request.QueryString("GUID").Trim.ToLower) Or (Request.QueryString("GUID") = "185") Or (Request.QueryString("GUID") = "ADOF") Then
            'gnCustomerKey = CUSTOMER_ACCOUNT_CODE_ADOF
            gnCustomerKey = dicCustomerFromGUID.Item(Request.QueryString("GUID").Trim.ToLower)
        Else
            Call ReturnError("Unrecognised GUID")
            Exit Sub
        End If
        
        If Request.QueryString("AWB") Is Nothing Then
            gnRequestType = REQUEST_TYPE_LAST_50
        End If
        
        If Not Request.QueryString("StartDate") Is Nothing Then
            If Not IsDate(Request.QueryString("StartDate")) Then
                Call ReturnError("Start Date is not a recognisable date")
                Exit Sub
                gnRequestType = REQUEST_TYPE_FROM_START_DATE
            End If
            If Not Request.QueryString("EndDate") Is Nothing Then
                If Not IsDate(Request.QueryString("EndDate")) Then
                    Call ReturnError("End Date is not a recognisable date")
                    Exit Sub
                End If
                gnRequestType = REQUEST_TYPE_DATE_RANGE
            End If
        End If
        If Not Request.QueryString("PODs") Is Nothing Then
            gnRequestType = REQUEST_TYPE_OUTSTANDING_PODS
        End If
        
        If Not Request.QueryString("Search") Is Nothing Then
            gnRequestType = REQUEST_TYPE_SEARCH
            If Request.QueryString("Search").Length < MIN_SEARCH_STRING_LENGTH Then
                Call ReturnError("Search string must be at least " & MIN_SEARCH_STRING_LENGTH & "characters long")
                Exit Sub
            End If
        End If

        Select Case gnRequestType
            Case REQUEST_TYPE_AWB_DETAIL
                psAWB = Request.QueryString("AWB")
                pnlConsignments.Visible = False
                Call ShowConsignment()
            Case REQUEST_TYPE_LAST_50, REQUEST_TYPE_DATE_RANGE, REQUEST_TYPE_FROM_START_DATE, REQUEST_TYPE_OUTSTANDING_PODS, REQUEST_TYPE_SEARCH
                pnlConsignmentDetail.Visible = False
                Call ShowConsignmentList()
        End Select
    End Sub
    
    Sub ReturnError(ByVal sErrorMessage As String)
        lblParameterErrorMessage.Text = sErrorMessage
        pnlConsignmentDetail.Visible = False
        pnlConsignments.Visible = False
        pnlParameterError.Visible = True
    End Sub
    
    Sub InitCustomerFromGUIDDictionary()
        dicCustomerFromGUID.Add("BA126AD7-2166-11D1-B1D0-00805FC1270E".ToLower, 185)
    End Sub
    
    Sub ShowConsignment()
        If Not psAWB = String.Empty Then
            ResetForm()
            ' lblError.Text = ""
            Dim sCnorName As String = String.Empty : Dim sCnorAddr1 = String.Empty : Dim sCnorAddr2 = String.Empty : Dim sCnorAddr3 As String = String.Empty : Dim sCnorTownCounty As String = String.Empty : Dim sCnorPostCodeCountry As String = String.Empty, sCnorContact As String = String.Empty
            Dim sCneeName As String = String.Empty : Dim sCneeAddr1 As String = String.Empty : Dim sCneeAddr2 As String = String.Empty : Dim sCneeAddr3 As String = String.Empty : Dim sCneeTownCounty As String = String.Empty : Dim sCneePostCodeCountry As String = String.Empty : Dim sCneeContact As String = String.Empty
            Dim oDataReader As SqlDataReader
            Dim oConn As New SqlConnection(gsConn)
            Dim oCmd As New SqlCommand("spASPNET_Tracking_GetConsignmentFromAWBForAnonymousAccess", oConn)
            oCmd.CommandType = CommandType.StoredProcedure
            Dim oParam As New SqlParameter("@AWB", SqlDbType.NVarChar, 50)
            oCmd.Parameters.Add(oParam)
            oParam.Value = psAWB
            Try
                oConn.Open()
                oDataReader = oCmd.ExecuteReader()
                oDataReader.Read()
                lblConsignment.Text = psAWB
                If Not IsDBNull(oDataReader("Key")) Then plConsignmentKey = oDataReader("Key")
                If Not IsDBNull(oDataReader("CustomerKey")) Then gnCustomerKey = oDataReader("CustomerKey")
                If Not IsDBNull(oDataReader("CreatedOn")) Then lblDate.Text = Format(oDataReader("CreatedOn"), "dd.MM.yy")
                If Not IsDBNull(oDataReader("CnorName")) Then sCnorName = oDataReader("CnorName")
                If Not IsDBNull(oDataReader("CnorAddr1")) Then sCnorAddr1 = oDataReader("CnorAddr1")
                If Not IsDBNull(oDataReader("CnorAddr2")) Then sCnorAddr2 = oDataReader("CnorAddr2")
                If Not IsDBNull(oDataReader("CnorAddr3")) Then sCnorAddr3 = oDataReader("CnorAddr3")
                If Not IsDBNull(oDataReader("CnorTown")) Then sCnorTownCounty = oDataReader("CnorTown")
                If Not IsDBNull(oDataReader("CnorState")) Then sCnorTownCounty &= "  " & oDataReader("CnorState")
                If Not IsDBNull(oDataReader("CnorPostCode")) Then sCnorPostCodeCountry = oDataReader("CnorPostCode")
                If Not IsDBNull(oDataReader("CnorCountryName")) Then sCnorPostCodeCountry &= "  " & oDataReader("CnorCountryName")
                If Not IsDBNull(oDataReader("CnorCtcName")) Then sCnorContact = oDataReader("CnorCtcName")
                If Not IsDBNull(oDataReader("CnorTel")) Then sCnorContact &= "  " & oDataReader("CnorTel")
                If Not IsDBNull(oDataReader("CneeName")) Then sCneeName = oDataReader("CneeName")
                If Not IsDBNull(oDataReader("CneeAddr1")) Then sCneeAddr1 = oDataReader("CneeAddr1")
                If Not IsDBNull(oDataReader("CneeAddr2")) Then sCneeAddr2 = oDataReader("CneeAddr2")
                If Not IsDBNull(oDataReader("CneeAddr3")) Then sCneeAddr3 = oDataReader("CneeAddr3")
                If Not IsDBNull(oDataReader("CneeTown")) Then sCneeTownCounty = oDataReader("CneeTown")
                If Not IsDBNull(oDataReader("CneeState")) Then sCneeTownCounty &= "  " & oDataReader("CneeState")
                If Not IsDBNull(oDataReader("CneePostCode")) Then sCneePostCodeCountry = oDataReader("CneePostCode")
                If Not IsDBNull(oDataReader("CneeCountryName")) Then sCneePostCodeCountry &= "  " & oDataReader("CneeCountryName")
                If Not IsDBNull(oDataReader("CneeCtcName")) Then sCneeContact = oDataReader("CneeCtcName")
                If Not IsDBNull(oDataReader("CneeTel")) Then sCneeContact &= "  " & oDataReader("CneeTel")
                If Not IsDBNull(oDataReader("Weight")) AndAlso oDataReader("Weight") <> "0" Then lblWeight.Text = oDataReader("Weight")
                If Not IsDBNull(oDataReader("NOP")) AndAlso oDataReader("NOP") <> "0" Then lblNOP.Text = oDataReader("NOP")
                If Not IsDBNull(oDataReader("SpecialInstructions")) Then lblSpclInstructions.Text = oDataReader("SpecialInstructions")
                If Not IsDBNull(oDataReader("ShippingInformation")) Then lblPackingNote.Text = oDataReader("ShippingInformation")
                If Not IsDBNull(oDataReader("Description")) Then lblContents.Text = oDataReader("Description")
                If Not IsDBNull(oDataReader("ValForCustoms")) AndAlso oDataReader("ValForCustoms") > 0 Then lblCustomsValue.Text = oDataReader("ValForCustoms")
                If Not IsDBNull(oDataReader("CustomerRef1")) Then lblCustRef1.Text = oDataReader("CustomerRef1")
                If Not IsDBNull(oDataReader("CustomerRef2")) Then lblCustRef2.Text = oDataReader("CustomerRef2")
                If Not IsDBNull(oDataReader("Misc1")) Then lblCustRef3.Text = oDataReader("Misc1")
                If Not IsDBNull(oDataReader("Misc2")) Then lblCustRef4.Text = oDataReader("Misc2")
                If Not IsDBNull(oDataReader("PODDate")) Then lblPODDate.Text = oDataReader("PODDate")
                If Not IsDBNull(oDataReader("PODName")) Then lblPODName.Text = oDataReader("PODName")
                If Not IsDBNull(oDataReader("PODTime")) Then lblPODTime.Text = oDataReader("PODTime")
                oDataReader.Close()
            Catch ex As SqlException
                Call ReturnError("Database error" & ex.Message)
                Exit Sub
            Finally
                oConn.Close()
            End Try
            'If dicCustomerFromGUID.TryGetValue(Request.QueryString("GUID").Trim.ToLower, nCustomerKey) Then
            'Else
            'End If
            
            If plConsignmentKey > 0 And gnCustomerKey = dicCustomerFromGUID.Item(Request.QueryString("GUID").Trim.ToLower) Then
                lblCnorAddr1.Text = sCnorName : lblCnorAddr2.Text = sCnorAddr1 : lblCnorAddr3.Text = sCnorAddr2 : lblCnorAddr4.Text = sCnorAddr3 : lblCnorAddr5.Text = sCnorTownCounty : lblCnorAddr6.Text = sCnorPostCodeCountry : lblCnorAddr7.Text = sCnorContact
                lblCneeAddr1.Text = sCneeName : lblCneeAddr2.Text = sCneeAddr1 : lblCneeAddr3.Text = sCneeAddr2 : lblCneeAddr4.Text = sCneeAddr3 : lblCneeAddr5.Text = sCneeTownCounty : lblCneeAddr6.Text = sCneePostCodeCountry : lblCneeAddr7.Text = sCneeContact
                Call GetTracking()
            Else
                pnlConsignmentDetail.Visible = False
                pnlParameterError.Visible = True
                lblParameterErrorMessage.Text = "No matching consignment found"
            End If
        End If
    End Sub
    
    Sub GetTracking()
        If plConsignmentKey > 0 Then
            ' lblError.Text = ""
            Dim oConn As New SqlConnection(gsConn)
            Dim oDataSet As New DataSet()
            Dim oAdapter As New SqlDataAdapter("spASPNET_Consignment_GetTracking", oConn)
            Try
                oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ConsignmentKey", SqlDbType.Int))
                oAdapter.SelectCommand.Parameters("@ConsignmentKey").Value = plConsignmentKey
                oAdapter.Fill(oDataSet, "Tracking")
                Dim Source As DataView = oDataSet.Tables("Tracking").DefaultView
                If Source.Count > 0 Then
                    dgTracking.DataSource = Source
                    dgTracking.DataBind()
                    dgTracking.Visible = True
                Else
                    dgTracking.Visible = False
                End If
            Catch ex As SqlException
                Call ReturnError("Database error" & ex.Message)
                Exit Sub
            Finally
                oConn.Close()
            End Try
        End If
    End Sub
    
    Function BindStockItems(ByVal SortField As String) As Integer
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataSet As New DataSet()
        Dim oAdapter As New SqlDataAdapter("spASPNET_LogisticBooking_GetMovementsWithVals", oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ConsignmentKey", SqlDbType.Int, 4))
        oAdapter.SelectCommand.Parameters("@ConsignmentKey").Value = plConsignmentKey
        Try
            oAdapter.Fill(oDataSet, "Movements")
            Dim Source As DataView = oDataSet.Tables("Movements").DefaultView
            Source.Sort = SortField
            If Source.Count > 0 Then
                dgBookingItems.DataSource = Source
                dgBookingItems.DataBind()
                dgBookingItems.Visible = True
            Else
                dgBookingItems.Visible = False
            End If
            BindStockItems = CInt(Source.Count)
        Catch ex As SqlException
            Call ReturnError("Database error" & ex.Message)
            Exit Function
        Finally
            oConn.Close()
        End Try
    End Function
    
    Sub ShowConsignmentList()
        lblAWBList.Text = ""
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataTable As New DataTable()
        Dim oAdapter As New SqlDataAdapter
        gnCustomerKey = dicCustomerFromGUID.Item(Request.QueryString("GUID").Trim.ToLower)
        Select Case gnRequestType
            Case REQUEST_TYPE_LAST_50
                oAdapter.SelectCommand = New SqlCommand("spASPNET_Consignment_TrackLast50AllForAnonymousAccess", oConn)
                oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
                oAdapter.SelectCommand.Parameters("@CustomerKey").Value = gnCustomerKey
            Case REQUEST_TYPE_DATE_RANGE
                Dim nConsignmentCount = GetConsignmentCount()
                If nConsignmentCount > 100 Then
                    Call ReturnError("Query would return " & nConsignmentCount & " records but no more 100 may be returned - please reduce the interval between dates")
                    Exit Sub
                Else
                    oAdapter.SelectCommand = New SqlCommand("spASPNET_Consignment_SearchByDateAnonymousAccess", oConn)
                    oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
                    oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@FromDate", SqlDbType.SmallDateTime))
                    oAdapter.SelectCommand.Parameters("@FromDate").Value = Request.QueryString("StartDate")
                    oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ToDate", SqlDbType.SmallDateTime))
                    oAdapter.SelectCommand.Parameters("@ToDate").Value = Request.QueryString("EndDate")
                    oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
                    oAdapter.SelectCommand.Parameters("@CustomerKey").Value = gnCustomerKey
                End If
            Case REQUEST_TYPE_FROM_START_DATE
                oAdapter.SelectCommand = New SqlCommand("spASPNET_Consignment_SearchFromDateAnonymousAccess", oConn)
                oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@FromDate", SqlDbType.SmallDateTime))
                oAdapter.SelectCommand.Parameters("@FromDate").Value = Request.QueryString("StartDate")
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
                oAdapter.SelectCommand.Parameters("@CustomerKey").Value = gnCustomerKey
            Case REQUEST_TYPE_OUTSTANDING_PODS
                oAdapter.SelectCommand = New SqlCommand("spASPNET_Consignment_TrackOutstandingPODs", oConn)
                oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
                oAdapter.SelectCommand.Parameters("@CustomerKey").Value = gnCustomerKey
            Case REQUEST_TYPE_SEARCH
                oAdapter.SelectCommand = New SqlCommand("spASPNET_Consignment_SearchAll2", oConn)
                oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SearchCriteria", SqlDbType.NVarChar, 50))
                oAdapter.SelectCommand.Parameters("@SearchCriteria").Value = Request.QueryString("Search")
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
                oAdapter.SelectCommand.Parameters("@CustomerKey").Value = gnCustomerKey
        End Select
        Try
            oAdapter.Fill(oDataTable)
            If oDataTable.Rows.Count > 0 Then
                dgConsignments.DataSource = oDataTable
                dgConsignments.DataBind()
                dgConsignments.Visible = True
            Else
                dgConsignments.Visible = False
                lblAWBList.Text = "no records found"
            End If
        Catch ex As SqlException
            Call ReturnError("Database error" & ex.Message)
            Exit Sub
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Function GetConsignmentCount() As Integer
        GetConsignmentCount = 0
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataTable As New DataTable()
        Dim oAdapter As New SqlDataAdapter("spASPNET_Consignment_SearchByDateGetCountForAnonymousAccess", oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@FromDate", SqlDbType.SmallDateTime))
        oAdapter.SelectCommand.Parameters("@FromDate").Value = Request.QueryString("StartDate")
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ToDate", SqlDbType.SmallDateTime))
        oAdapter.SelectCommand.Parameters("@ToDate").Value = Request.QueryString("EndDate")
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@CustomerKey").Value = gnCustomerKey
        Try
            oAdapter.Fill(oDataTable)
            GetConsignmentCount = oDataTable.Rows(0).Item(0)
        Catch ex As SqlException
            Call ReturnError("Database error" & ex.Message)
            Exit Function
        Finally
            oConn.Close()
        End Try
    End Function
    
    Protected Sub ResetForm()
        lblDate.Text = ""
        lblCnorAddr1.Text = "" : lblCnorAddr2.Text = "" : lblCnorAddr3.Text = "" : lblCnorAddr4.Text = "" : lblCnorAddr5.Text = ""
        lblCneeAddr1.Text = "" : lblCneeAddr2.Text = "" : lblCneeAddr3.Text = "" : lblCneeAddr4.Text = "" : lblCneeAddr5.Text = ""
        lblWeight.Text = "" : lblNOP.Text = "" : lblSpclInstructions.Text = "" : lblContents.Text = "" : lblCustRef1.Text = "" : lblCustRef2.Text = ""
        lblPODDate.Text = "" : lblPODName.Text = "" : lblPODTime.Text = ""
    End Sub

    Property psAWB() As String
        Get
            Dim o As Object = ViewState("ST_AWB")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("ST_AWB") = Value
        End Set
    End Property
    
    Property plConsignmentKey() As Long
        Get
            Dim o As Object = ViewState("ST_ConsignmentKey")
            If o Is Nothing Then
                Return -1
            End If
            Return CLng(o)
        End Get
        Set(ByVal Value As Long)
            ViewState("ST_ConsignmentKey") = Value
        End Set
    End Property

    Property psFilterName() As String
        Get
            Dim o As Object = ViewState("ST_FilterName")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("ST_FilterName") = Value
        End Set
    End Property
    
</script>

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Transworld Consignment Tracking</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:Panel ID="pnlConsignmentDetail" runat="server" Width="100%">
            <table id="tabConsignmentDetail1" style="font-size: xx-small; width: 100%; font-family: Verdana">
                <tr>
                    <td style="width: 450px; white-space: nowrap" valign="middle">
                        <asp:Label ID="xyz" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="X-Small"
                            ForeColor="Gray">Consignment:</asp:Label>
                        &nbsp; &nbsp;<asp:Label ID="lblConsignment" runat="server" Font-Names="Verdana" Font-Size="X-Small"
                            ForeColor="Red"></asp:Label>
                        &nbsp; &nbsp; &nbsp; &nbsp;<asp:Label ID="l0011" runat="server" Font-Size="X-Small"
                            ForeColor="Gray">Date:</asp:Label>
                        &nbsp; &nbsp;<asp:Label ID="lblDate" runat="server" Font-Names="Verdana" Font-Size="X-Small"
                            ForeColor="Red"></asp:Label>
                    </td>
                    <td align="right" style="white-space: nowrap" valign="middle">
                    </td>
                    <td align="right" style="width: 40px; white-space: nowrap" valign="middle">
                    </td>
                </tr>
            </table>
            <br />
            <table id="tabConsignmentDetail2" style="font-size: xx-small; width: 100%; color: navy;
                font-family: Verdana">
                <tr>
                    <td style="width: 15%">
                    </td>
                    <td style="width: 35%">
                    </td>
                    <td style="width: 10%">
                    </td>
                    <td style="width: 40%">
                    </td>
                </tr>
                <tr>
                    <td>
                        <asp:Label ID="l0001" runat="server" Font-Bold="true" ForeColor="Gray">From:</asp:Label>
                    </td>
                    <td>
                        <asp:Label ID="lblCnorAddr1" runat="server"></asp:Label>
                    </td>
                    <td>
                        <asp:Label ID="l0002" runat="server" Font-Bold="true" ForeColor="Gray">To:</asp:Label>
                    </td>
                    <td>
                        <asp:Label ID="lblCneeAddr1" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td>
                        <asp:Label ID="lblCnorAddr2" runat="server"></asp:Label>
                    </td>
                    <td>
                    </td>
                    <td>
                        <asp:Label ID="lblCneeAddr2" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td>
                        <asp:Label ID="lblCnorAddr3" runat="server"></asp:Label>
                    </td>
                    <td>
                    </td>
                    <td>
                        <asp:Label ID="lblCneeAddr3" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td>
                        <asp:Label ID="lblCnorAddr4" runat="server"></asp:Label>
                    </td>
                    <td>
                    </td>
                    <td>
                        <asp:Label ID="lblCneeAddr4" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td>
                        <asp:Label ID="lblCnorAddr5" runat="server"></asp:Label>
                    </td>
                    <td>
                    </td>
                    <td>
                        <asp:Label ID="lblCneeAddr5" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td>
                        <asp:Label ID="lblCnorAddr6" runat="server"></asp:Label>
                    </td>
                    <td>
                    </td>
                    <td>
                        <asp:Label ID="lblCneeAddr6" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td>
                        <asp:Label ID="lblCnorAddr7" runat="server"></asp:Label>
                    </td>
                    <td>
                    </td>
                    <td>
                        <asp:Label ID="lblCneeAddr7" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td>
                        <asp:Label ID="l0003" runat="server" Font-Bold="true" ForeColor="Gray">NOP:</asp:Label>
                    </td>
                    <td>
                        <asp:Label ID="lblNOP" runat="server"></asp:Label>
                    </td>
                    <td>
                        <asp:Label ID="l0004" runat="server" Font-Bold="true" ForeColor="Gray">Weight:</asp:Label>
                    </td>
                    <td>
                        <asp:Label ID="lblWeight" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td>
                        <asp:Label ID="l0005" runat="server" Font-Bold="true" ForeColor="Gray">Contents:</asp:Label>
                    </td>
                    <td>
                        <asp:Label ID="lblContents" runat="server"></asp:Label>
                    </td>
                    <td>
                        <asp:Label ID="lblLegendCustomsValue" runat="server" Font-Bold="true" ForeColor="Gray">Value:</asp:Label>
                    </td>
                    <td>
                        <asp:Label ID="lblCustomsValue" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td>
                        <asp:Label ID="l0007" runat="server" Font-Bold="true" ForeColor="Gray">Spcl Instr:</asp:Label>
                    </td>
                    <td colspan="3">
                        <asp:Label ID="lblSpclInstructions" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td>
                        <asp:Label ID="Label3" runat="server" Font-Bold="true" ForeColor="Gray">Packing Note:</asp:Label>
                    </td>
                    <td colspan="3">
                        <asp:Label ID="lblPackingNote" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr id="trCustRefRow1" runat="server">
                    <td>
                        <asp:Label ID="l0008" runat="server" Font-Bold="true" ForeColor="Gray">Ref 1:</asp:Label>
                    </td>
                    <td>
                        <asp:Label ID="lblCustRef1" runat="server"></asp:Label>
                    </td>
                    <td>
                        <asp:Label ID="l0009" runat="server" Font-Bold="true" ForeColor="Gray">Ref 2:</asp:Label>
                    </td>
                    <td>
                        <asp:Label ID="lblCustRef2" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td>
                        <asp:Label ID="lblLegendCustRef3" runat="server" Font-Bold="true" ForeColor="Gray"
                            Text="Ref 3:"></asp:Label>
                    </td>
                    <td>
                        <asp:Label ID="lblCustRef3" runat="server"></asp:Label>
                    </td>
                    <td>
                        <asp:Label ID="lblLegendCustRef4" runat="server" Font-Bold="true" ForeColor="Gray"
                            Text="Ref 4:"></asp:Label>
                    </td>
                    <td>
                        <asp:Label ID="lblCustRef4" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td>
                        <asp:Label ID="l0010" runat="server" Font-Bold="true" ForeColor="Gray">Received By:</asp:Label>
                    </td>
                    <td colspan="3">
                        <asp:Label ID="lblPODDate" runat="server" ForeColor="Red"></asp:Label>
                        &nbsp;<asp:Label ID="lblPODName" runat="server" ForeColor="Red"></asp:Label>
                        &nbsp;<asp:Label ID="lblPODTime" runat="server" ForeColor="Red"></asp:Label>
                    </td>
                </tr>
            </table>
            <hr />
            <asp:Table ID="Table1" runat="Server" Width="100%">
                <asp:TableRow>
                    <asp:TableCell Wrap="False">
                        <asp:Label ID="Label2" runat="server" Font-Bold="True" Font-Names="Arial" Font-Size="X-Small"
                            ForeColor="Gray" Text="Tracking Events"></asp:Label>
                    </asp:TableCell>
                </asp:TableRow>
            </asp:Table>
            <asp:DataGrid ID="dgTracking" runat="server" AutoGenerateColumns="False" Font-Names="Verdana"
                Font-Size="XX-Small" GridLines="None" Visible="False" Width="100%">
                <FooterStyle Wrap="False" />
                <HeaderStyle Font-Names="Verdana" Wrap="False" />
                <PagerStyle BackColor="Silver" Font-Bold="True" Font-Names="Verdana" Font-Size="X-Small"
                    ForeColor="Blue" HorizontalAlign="Center" Mode="NumericPages" Wrap="False" />
                <Columns>
                    <asp:BoundColumn DataField="Time" DataFormatString="{0:dd.MM.yy HH:mm}" HeaderText="Time">
                        <HeaderStyle Font-Bold="true" ForeColor="Gray" Width="15%" Wrap="False" />
                        <ItemStyle ForeColor="Navy" VerticalAlign="Top" Wrap="False" />
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="Location" HeaderText="Location">
                        <HeaderStyle Font-Bold="true" ForeColor="Gray" HorizontalAlign="Left" Width="10%"
                            Wrap="False" />
                        <ItemStyle ForeColor="Navy" VerticalAlign="Top" Wrap="False" />
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="Description" HeaderText="Description">
                        <HeaderStyle Font-Bold="true" ForeColor="Gray" Width="75%" Wrap="False" />
                        <ItemStyle ForeColor="Navy" VerticalAlign="Top" />
                    </asp:BoundColumn>
                </Columns>
            </asp:DataGrid>
            <br />
            <asp:Panel ID="pnlBookingDetail" runat="server" Visible="false" Width="100%">
                <hr />
                <asp:Table ID="Table3" runat="Server" Width="100%">
                    <asp:TableRow>
                        <asp:TableCell Wrap="False">
                            <asp:Label ID="lblGrandTotal" runat="server" Font-Bold="True" Font-Names="Arial"
                                Font-Size="X-Small" ForeColor="Gray" Text="Item(s) Booked"></asp:Label>
                        </asp:TableCell>
                    </asp:TableRow>
                </asp:Table>
                <asp:DataGrid ID="dgBookingItems" runat="server" AutoGenerateColumns="False" CellPadding="2"
                    CellSpacing="-1" Font-Names="Verdana" Font-Size="XX-Small" GridLines="None" Width="100%">
                    <HeaderStyle Font-Bold="True" />
                    <AlternatingItemStyle ForeColor="WhiteSmoke" />
                    <ItemStyle ForeColor="#0000C0" />
                    <Columns>
                        <asp:BoundColumn DataField="ProductCode" HeaderText="Code">
                            <HeaderStyle ForeColor="Gray" Wrap="False" />
                            <ItemStyle ForeColor="Navy" VerticalAlign="Top" Wrap="False" />
                        </asp:BoundColumn>
                        <asp:BoundColumn DataField="ProductDate" HeaderText="Prod Date">
                            <HeaderStyle ForeColor="Gray" Wrap="False" />
                            <ItemStyle ForeColor="Navy" VerticalAlign="Top" Wrap="False" />
                        </asp:BoundColumn>
                        <asp:BoundColumn DataField="ProductDescription" HeaderText="Description">
                            <HeaderStyle ForeColor="Gray" Wrap="False" />
                            <ItemStyle ForeColor="Navy" VerticalAlign="Top" />
                        </asp:BoundColumn>
                        <asp:BoundColumn DataField="ItemsOut" DataFormatString="{0:#,##0}" HeaderText="Qty">
                            <HeaderStyle ForeColor="Gray" HorizontalAlign="Right" Wrap="False" />
                            <ItemStyle ForeColor="Navy" HorizontalAlign="Right" VerticalAlign="Top" Wrap="False" />
                        </asp:BoundColumn>
                        <asp:BoundColumn DataField="UnitValue" DataFormatString="{0:#,##0.00}" HeaderText="Unit Cost">
                            <HeaderStyle ForeColor="Gray" HorizontalAlign="Right" Wrap="False" />
                            <ItemStyle ForeColor="Navy" HorizontalAlign="Right" VerticalAlign="Top" />
                        </asp:BoundColumn>
                        <asp:BoundColumn DataField="Cost" DataFormatString="{0:#,##0.00}" HeaderText="Total Cost">
                            <HeaderStyle ForeColor="Gray" HorizontalAlign="Right" Wrap="False" />
                            <ItemStyle ForeColor="Navy" HorizontalAlign="Right" VerticalAlign="Top" />
                        </asp:BoundColumn>
                    </Columns>
                </asp:DataGrid>
            <asp:Label ID="lblLegendGrandTotal" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                Text="Total cost for all items:"></asp:Label>
            <asp:Label ID="Label1" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Total cost for all items:"></asp:Label>
            </asp:Panel>
            </asp:Panel>
        <asp:Panel id="pnlConsignments" Width="100%" runat="server" visible="False">            
            <br />
            <asp:DataGrid id="dgConsignments" runat="server" Width="100%" Font-Names="Verdana" Font-Size="XX-Small" ShowFooter="True" GridLines="None" AutoGenerateColumns="False" Visible="False">
                <FooterStyle wrap="False"></FooterStyle>
                <HeaderStyle font-names="Verdana" wrap="False"></HeaderStyle>
                <AlternatingItemStyle backcolor="White"></AlternatingItemStyle>
                <ItemStyle backcolor="LightGray"></ItemStyle>
                <Columns>
                    <asp:BoundColumn DataField="Key" visible="False">
                        <HeaderStyle></HeaderStyle>
                        <ItemStyle></ItemStyle>
                    </asp:BoundColumn>
                    <asp:TemplateColumn>
                        <ItemStyle wrap="False" horizontalalign="Left"></ItemStyle>
                        <ItemTemplate>
                        </ItemTemplate>
                    </asp:TemplateColumn>
                    <asp:BoundColumn DataField="AWB" HeaderText="Consignment">
                        <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                        <ItemStyle wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="CreatedOn" HeaderText="Created On" DataFormatString="{0:dd/MM/yy HH:mm}">
                        <HeaderStyle font-bold="True" wrap="False" horizontalalign="Left" forecolor="Gray"></HeaderStyle>
                        <ItemStyle wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="BookedBy" HeaderText="Booked By">
                        <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="CneeName" HeaderText="To">
                        <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                        <ItemStyle wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="CneeTown" HeaderText="City">
                        <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                        <ItemStyle wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="CountryName" SortExpression="CountryName" HeaderText="Country">
                        <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="POD" HeaderText="P.O.D.">
                        <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                    </asp:BoundColumn>
                </Columns>
            </asp:DataGrid>
            <asp:Label id="lblAWBList" runat="server" font-names="Verdana" font-size="X-Small" forecolor="Blue"></asp:Label>
        </asp:Panel>
        <asp:Panel ID="pnlParameterError" Visible="False" Font-Names="Verdana" runat="server" Width="100%">
        <br />
            <asp:Label ID="lblParameterErrorMessage" runat="server" Text=""></asp:Label>
            <br />
        </asp:Panel>
    </div>
    </form>
</body>
</html>
