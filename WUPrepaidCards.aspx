<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.IO" %>
<%@ Register Assembly="Telerik.Web.UI" Namespace="Telerik.Web.UI" TagPrefix="telerik" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ Import Namespace="FileHelpers" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<script runat="server">

    'IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ClientData_WU_PrePaidCardMapping]') AND type in (N'U'))
    'DROP TABLE [dbo].[ClientData_WU_PrePaidCardMapping]
    'GO

    'CREATE TABLE [dbo].[ClientData_WU_PrePaidCardMapping](
    '[id] [int] IDENTITY(1,1) NOT NULL,
    '[ConsignmentKey] [int] NULL,
    '[ProxyNo] [varchar](50) NOT NULL,
    '[BarcodeNo] [varchar](50) NOT NULL,
    '[IsScanned] [bit] NOT NULL,
    '[LoadDateTime] [smalldatetime] NULL,
    '[ScanDateTime] [smalldatetime] NULL,
    '[UserKey] [int] NOT NULL
    ') ON [PRIMARY]
    'GO

    'GRANT  SELECT ,  INSERT ,  DELETE ,  UPDATE  ON [dbo].[ClientData_WU_PrePaidCardMapping] TO [LogisticsUserRole]
    'GO

    'GRANT  SELECT ,  INSERT ,  DELETE ,  UPDATE  ON [dbo].[ClientData_WU_PrePaidCardMapping] TO [LogisticsAdminRole]
    'GO
    
    Const ITEMS_PER_REQUEST As Integer = 30
    Const PREPAID_CARD_KEY = 51865

    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsNumeric(Session("UserKey")) Then
            'Server.Transfer("session_expired.aspx")
        End If
        If Not IsPostBack Then
            tbBarcode.Attributes.Add("onkeypress", "return clickButton(event,'" + btnGo.ClientID + "')")
            Call SetTitle()
            Call PopulateAWBDropdown()
            Call CreateWUBarcodesFolder()
            ddlConsignment.Focus()
            'tbBarcode.Focus()
        End If
    End Sub
    
    Protected Sub SetTitle()
        Dim sTitle As String = Session("SiteTitle")
        If sTitle <> String.Empty Then
            sTitle += " - "
        End If
        Page.Header.Title = sTitle & "Western Union Pre-paid Card (WUF/194) Pack Capture"
    End Sub

    Protected Sub PopulateAWBDropdown()
        ddlConsignment.Items.Clear()
        Dim sSQL As String = "SELECT TOP 30 CAST(REPLACE(CONVERT(VARCHAR(11), c.CreatedOn, 106), ' ', '-') AS varchar(20)) + '  ' + CAST(ConsignmentKey AS varchar(10)) + '  ' + c.CneeName + ' ' + c.CneeAddr1 + ' ' + CneePostcode 'Consignment', ConsignmentKey FROM LogisticMovement lm INNER JOIN Consignment c ON lm.ConsignmentKey = c.[key] WHERE LogisticProductKey = 51865 AND ConsignmentKey IS NOT NULL AND ItemsOut > 0 ORDER BY LogisticMovementKey DESC"
        Dim dtConsignment As DataTable = ExecuteQueryToDataTable(sSQL)
        If dtConsignment.Rows.Count = 0 Then
            WebMsgBox.Show("No consignments found.")
        Else
            ddlConsignment.Items.Add(New ListItem("- please select - ", 0))
            For Each dr As DataRow In dtConsignment.Rows
                ddlConsignment.Items.Add(New ListItem(dr("Consignment"), dr("ConsignmentKey")))
            Next
        End If
    End Sub
    
    Private Sub CreateWUBarcodesFolder()
        Dim sPath As String = Server.MapPath("~/")
        If Not Directory.Exists(sPath & "\WUPrepaidCardBarcodes") Then
            Directory.CreateDirectory(sPath & "\WUPrepaidCardBarcodes")
        End If
    End Sub

    Protected Sub btnUploadBarcodesCSVFile_Click(sender As Object, e As System.EventArgs)
        Try
            If ruWUBarcodesFileUpload.UploadedFiles.Count > 0 Then
                lblNoResults.Visible = False
                psFileName = ruWUBarcodesFileUpload.UploadedFiles(0).GetName
                lblFileName.Text = ruWUBarcodesFileUpload.UploadedFiles(0).GetName
                tbJournal.Text = String.Empty
                Call WriteToJournal("Uploaded " & ruWUBarcodesFileUpload.UploadedFiles(0).GetName & " @ " & Format(Date.Now, "d-MMM-yyyy hh:mm:ss"))
                Call ProcessBarcodesFile()
            Else
                Call WriteToJournal("Nothing uploaded")
            End If
        Catch ex As Exception
            lblNoResults.Text = ex.Message.ToString()
            Call WriteToJournal(ex.Message.ToString())
        End Try
    End Sub

    Protected Sub ProcessBarcodesFile()
        Dim sMessage As String = String.Empty
        Dim sFilePath As String = Server.MapPath("~/WUPrepaidCardBarcodes/" & psFileName)
        Dim BarcodeFileEntries As BarcodeFileEntry()
        Dim engine As New FileHelperEngine(GetType(BarcodeFileEntry))
        Try
            BarcodeFileEntries = DirectCast(engine.ReadFile(sFilePath), BarcodeFileEntry())
        Catch ex As Exception
            Call WriteToJournal("Could not read file: " & ex.Message)
            Exit Sub
        End Try
        Dim bError As Boolean = False
        For Each oEntry As BarcodeFileEntry In BarcodeFileEntries
            Dim sProxyNo As String = oEntry.ProxyNo
            Dim sBarcodeNo As String = oEntry.BarcodeNo
            Dim sExistingBarcodeNo As String = GetProxyNumber(sProxyNo)
            If sExistingBarcodeNo <> String.Empty Then
                If IsNumeric(sExistingBarcodeNo) Then
                    Call WriteToJournal("WARNING: Proxy No. " & sProxyNo & " (" & sBarcodeNo & ") already loaded (associated with Barcode No: " & sExistingBarcodeNo & ")")
                Else
                    Call WriteToJournal("Duplicate proxy number detected.")
                End If
                bError = True
            End If
        Next
        If Not bError Then
            Dim nRecordCount As Int32 = 0
            For Each oEntry As BarcodeFileEntry In BarcodeFileEntries
                Call AddEntry(oEntry.ProxyNo, oEntry.BarcodeNo)
                nRecordCount += 1
            Next
            sMessage = "Added " & nRecordCount.ToString & " record(s)."
        Else
            sMessage = "One or more errors detected. The file was not processed."
        End If
        Call WriteToJournal(sMessage)
        WebMsgBox.Show(sMessage)
    End Sub
    
    Protected Sub AddEntry(sProxyNo As String, sBarcodeNo As String)
        Dim sSQL As String = "INSERT INTO ClientData_WU_PrePaidCardMapping (ProxyNo, BarcodeNo, IsScanned, LoadDateTime, UserKey) VALUES ('" & sProxyNo & "', '" & sBarcodeNo & "', 0, GETDATE(), 0)"
        Call ExecuteQueryToDataTable(sSQL)
    End Sub
    
    Protected Function GetProxyNumber(sProxyNo As String) As String
        GetProxyNumber = String.Empty
        Dim sSQL As String = "SELECT BarcodeNo FROM ClientData_WU_PrePaidCardMapping WHERE ProxyNo = '" & sProxyNo & "'"
        Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
        If dt.Rows.Count > 0 Then
            GetProxyNumber = dt.Rows(0).Item(0)
        ElseIf dt.Rows.Count > 1 Then
            GetProxyNumber = "Duplicate proxy number detected (" & sProxyNo & ")."
        End If
    End Function
    
    Protected Sub WriteToJournal(sMessage As String)
        tbJournal.Text += Date.Now & " " & sMessage.Replace("\n", "") & Environment.NewLine
    End Sub
    
    <DelimitedRecord(",")> _
    <IgnoreFirst(1)> _
    Public Class BarcodeFileEntry
        <FieldQuoted("""", QuoteMode.OptionalForBoth)> _
        Public ProxyNo As String
        
        <FieldQuoted("""", QuoteMode.OptionalForBoth)> _
        Public BarcodeNo As String
    End Class

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

    Property psFileName() As String
        Get
            Dim o As Object = ViewState("WU_Prepaid_Cards_Barcodes_FileName")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("WU_Prepaid_Cards_Barcodes_FileName") = Value
        End Set
    End Property

    Protected Sub SetMaintenanceVisibility(bVisible As Boolean)
        trMaintenance01.Visible = bVisible
        trMaintenance02.Visible = bVisible
        trMaintenance03.Visible = bVisible
    End Sub
    
    Protected Sub lnkbtnMaintenance_Click(sender As Object, e As System.EventArgs)
        Dim lnkbtn As LinkButton = sender
        If lnkbtn.Text.Contains("hide") Then
            lnkbtn.Text = "maintenance"
            Call SetMaintenanceVisibility(False)
        Else
            lnkbtn.Text = "hide maintenance"
            Call SetMaintenanceVisibility(True)
        End If
    End Sub

    Protected Sub btnGo_Click(sender As Object, e As System.EventArgs)
        tbBarcode.Text = tbBarcode.Text.Trim
        Dim nBarcodeLength As Int32 = tbBarcode.Text.Length
        If nBarcodeLength > 0 Then
            Dim sMessage As String = String.Empty
            If Not (nBarcodeLength = 4) Or (nBarcodeLength = 13) Then
                If nBarcodeLength = 1 Then
                    sMessage = "Unexpected barcode length (" & nBarcodeLength.ToString & " char).\n\n Please enter a full 13 digit barcode or the last 4 digits, eg ""2345""."
                Else
                    sMessage = "Unexpected barcode length (" & nBarcodeLength.ToString & " chars).\n\n Please enter a full 13 digit barcode or the last 4 digits, eg ""2345""."
                End If
                'WebMsgBox.Show(sMessage)
                'WriteToJournal(sMessage)
            Else
                Dim sSQL As String
                If nBarcodeLength = 13 Then
                    sSQL = "SELECT TOP 1 * FROM ClientData_WU_PrePaidCardMapping WHERE BarcodeNo = '" & tbBarcode.Text & "'"
                Else
                    sSQL = "SELECT TOP 1 * FROM ClientData_WU_PrePaidCardMapping WHERE BarcodeNo LIKE '%" & tbBarcode.Text & "'"
                End If
                Dim dtBarcode As DataTable = ExecuteQueryToDataTable(sSQL)
                If dtBarcode.Rows.Count = 0 Then
                    sMessage = "Could not match this barcode (" & tbBarcode.Text & ") against an expected barcode."
                    'WebMsgBox.Show(sMessage)
                    'WriteToJournal(sMessage)

                Else
                    If dtBarcode.Rows(0).Item("IsScanned") <> 0 Then
                        sMessage = "This barcode (" & tbBarcode.Text & ") has already been scanned (scan logged " & dtBarcode.Rows(0).Item("ScanDateTime") & ")."
                        'WebMsgBox.Show()
                    Else
                        If nBarcodeLength = 13 Then
                            sSQL = "UPDATE ClientData_WU_PrePaidCardMapping SET ConsignmentKey = " & ddlConsignment.SelectedValue & ", IsScanned = 1, ScanDateTime = GETDATE() WHERE BarcodeNo = '" & tbBarcode.Text & "' SELECT @@ROWCOUNT"
                        Else
                            sSQL = "UPDATE ClientData_WU_PrePaidCardMapping SET ConsignmentKey = " & ddlConsignment.SelectedValue & ", IsScanned = 1, ScanDateTime = GETDATE() WHERE BarcodeNo LIKE '%" & tbBarcode.Text & "' SELECT @@ROWCOUNT"
                        End If
                        dtBarcode = ExecuteQueryToDataTable(sSQL)
                        sMessage = "Barcode " & tbBarcode.Text & " associated with consignment " & ddlConsignment.SelectedValue & ". " & dtBarcode.Rows(0).Item(0) & " proxy records matched."
                    End If
                End If
            End If
            WebMsgBox.Show(sMessage)
            WriteToJournal(sMessage)
            tbBarcode.Text = String.Empty
            tbBarcode.Focus()
        End If
    End Sub
    
    Protected Sub lnkbtnUnscanned_Click(sender As Object, e As System.EventArgs)
        Call ExcelReport("unscanned")
    End Sub

    Protected Sub lnkbtnSscanned_Click(sender As Object, e As System.EventArgs)
        Call ExcelReport("scanned")
    End Sub
    
    Protected Sub ExcelReport(sType As String)
        Dim sSQL As String
        Dim sFilename As String
        Dim sHeader As String
        If sType.Contains("un") Then
            sSQL = "SELECT '' 'ConsignmentKey', ProxyNo, BarcodeNo, ISNULL(CAST(REPLACE(CONVERT(VARCHAR(11), LoadDateTime, 106), ' ', '-') AS varchar(20)),'(never)') 'LoadDate', ISNULL(CAST(REPLACE(CONVERT(VARCHAR(11), ScanDateTime, 106), ' ', '-') AS varchar(20)),'(never)') 'ScanDate' FROM ClientData_WU_PrePaidCardMapping WHERE IsScanned = 0 ORDER BY [id]"
            sFilename = "WU_PrePaid_Cards_ProxyNos_UNSCANNED.csv"
            sHeader = "Western Union Prepaid Card Proxy Nos - UNSCANNED"
        Else
            sSQL = "SELECT CAST(ConsignmentKey AS varchar(10)) 'ConsignmentKey', ProxyNo, BarcodeNo, ISNULL(CAST(REPLACE(CONVERT(VARCHAR(11), LoadDateTime, 106), ' ', '-') AS varchar(20)),'(never)') 'LoadDate', ISNULL(CAST(REPLACE(CONVERT(VARCHAR(11), ScanDateTime, 106), ' ', '-') AS varchar(20)),'(never)') 'ScanDate' FROM ClientData_WU_PrePaidCardMapping WHERE IsScanned = 1 ORDER BY [id]"
            sFilename = "WU_PrePaid_Cards_ProxyNos_SCANNED.csv"
            sHeader = "Western Union Prepaid Card Proxy Nos - SCANNED"
        End If
        Dim dtReportData As DataTable = ExecuteQueryToDataTable(sSQL)
        If dtReportData.Rows.Count = 0 Then
            WebMsgBox.Show("No data found.")
        Else
        
            Response.Clear()
            Response.ContentType = "text/csv"
            Response.AddHeader("Content-Disposition", "attachment; filename=" & sFilename)
        
            Dim sbResponse As New StringBuilder

            sbResponse.Append(sHeader)
            sbResponse.Append(",")
            sbResponse.Append("")
            sbResponse.Append(",")
            sbResponse.Append("")
            sbResponse.Append(",")
            sbResponse.Append("")
            sbResponse.Append(",")
            sbResponse.Append("")
            sbResponse.Append(vbCrLf)

            sbResponse.Append("++ Consignment ++")
            sbResponse.Append(",")
            sbResponse.Append("++ Proxy No ++")
            sbResponse.Append(",")
            sbResponse.Append("++ Barcode No ++")
            sbResponse.Append(",")
            sbResponse.Append("++ Load Date ++")
            sbResponse.Append(",")
            sbResponse.Append("++ Scan Date ++")
            sbResponse.Append(vbCrLf)

            For Each dr As DataRow In dtReportData.Rows
                sbResponse.Append("# " & dr("ConsignmentKey"))
                sbResponse.Append(",")
                sbResponse.Append("# " & dr("ProxyNo"))
                sbResponse.Append(",")
                sbResponse.Append("# " & dr("BarcodeNo"))
                sbResponse.Append(",")
                sbResponse.Append(dr("LoadDate"))
                sbResponse.Append(",")
                sbResponse.Append(dr("ScanDate"))
                sbResponse.Append(vbCrLf)
            Next
            Response.Write(sbResponse.ToString)
            Response.End()
        End If
    End Sub
    
    Public Function ExecuteStoredProcedureToDataTable(ByVal sp_name As String, Optional ByVal IListPrams As List(Of SqlParameter) = Nothing) As DataTable
        Dim oDataTable As New DataTable
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter(sp_name, oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        If Not IListPrams Is Nothing AndAlso IListPrams.Count > 0 Then
            oAdapter.SelectCommand.Parameters.AddRange(IListPrams.ToArray)
        End If
        Try
            oAdapter.Fill(oDataTable)
        Catch ex As Exception
            WebMsgBox.Show(ex.Message.ToString())
        End Try
        ExecuteStoredProcedureToDataTable = oDataTable
    End Function
    

    Protected Sub ddlConsignment_SelectedIndexChanged(sender As Object, e As System.EventArgs)
        Dim ddl As DropDownList = sender
        If ddl.SelectedIndex > 0 Then
            tbBarcode.Enabled = True
            tbBarcode.Focus()
        Else
            tbBarcode.Enabled = False
            ddlConsignment.Focus()
        End If
    End Sub
    
</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <main:Header ID="ctlHeader" runat="server" />
    <asp:ScriptManager ID="ScriptManager1" runat="server" />
    <div>
        <table style="width: 100%">
            <tr>
                <td colspan="3">
                    &nbsp;
                    <asp:Label ID="lblLegendTitle" runat="server" Font-Bold="True" Font-Names="Verdana"
                        Font-Size="Small">Western Union Prepaid Cards (WUF/194)</asp:Label>
                    &nbsp;
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td align="right">
                    <asp:Label ID="lblLegendScanPackBarcode1" runat="server" Font-Names="Verdana" 
                        Font-Size="X-Small">AWB:</asp:Label>
                </td>
                <td valign="middle">
                    <asp:DropDownList ID="ddlConsignment" runat="server" Font-Names="Verdana" 
                        Font-Size="Small" 
                        onselectedindexchanged="ddlConsignment_SelectedIndexChanged" 
                        AutoPostBack="True"/>
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td align="right">
                    <asp:Label ID="lblLegendScanPackBarcode" runat="server" Font-Names="Verdana" Font-Size="X-Small">Barcode (or last 4 digits):</asp:Label>
                </td>
                <td valign="middle">
                    <asp:TextBox ID="tbBarcode" runat="server" Width="300px" Font-Names="Verdana" 
                        Font-Size="Small" Enabled="False" />
                    &nbsp;<asp:Button ID="btnGo" runat="server" Text="go" onclick="btnGo_Click" />
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td>
                    &nbsp;</td>
                <td valign="middle">
                    &nbsp;</td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td colspan="3">
                            &nbsp;<asp:Label ID="lblLegendExcelReports" runat="server" Font-Names="Verdana"
                        Font-Size="X-Small">Excel reports:</asp:Label>
                &nbsp;<asp:LinkButton ID="lnkbtnUnscanned" runat="server" 
                        Font-Names="Verdana" Font-Size="XX-Small" onclick="lnkbtnUnscanned_Click">unscanned proxy nos / barcodes</asp:LinkButton>

                &nbsp;
                            <asp:LinkButton ID="lnkbtnSscanned" runat="server" Font-Names="Verdana" 
                                Font-Size="XX-Small" onclick="lnkbtnSscanned_Click">scanned proxy nos / barcodes</asp:LinkButton>

                </td>
            </tr>
            <tr>
                <td colspan="3">
                    &nbsp;<asp:Label ID="lblLegendJournal" runat="server" Font-Names="Verdana" 
                        Font-Size="XX-Small">Session Journal:</asp:Label>
                    <br />
                    &nbsp;<asp:TextBox ID="tbJournal" runat="server" Width="96%" Rows="6" 
                        TextMode="MultiLine" Font-Names="Courier New, Verdana" Font-Size="X-Small"/>
                </td>
            </tr>
            <tr>
                <td colspan="3">
                    <hr />
                </td>
            </tr>
            <tr>
                <td>
                    &nbsp;<asp:LinkButton ID="lnkbtnMaintenance" runat="server" Font-Names="Verdana"
                        Font-Size="XX-Small" OnClick="lnkbtnMaintenance_Click">maintenance</asp:LinkButton>
                </td>
                <td>
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr id="trMaintenance01" runat="server" visible="false">
                <td>
                    &nbsp;<asp:Label ID="lblLegendScanPackBarcode0" runat="server" Font-Names="Verdana"
                        Font-Size="X-Small">Select the Proxy Number / Barcode CSV filename:</asp:Label>
                </td>
                <td>
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr id="trMaintenance02" runat="server" visible="false">
                <td>
                    &nbsp;<telerik:RadUpload ID="ruWUBarcodesFileUpload" Width="100%" TargetFolder="~/WUPrepaidCardBarcodes"
                        AllowedFileExtensions=".csv,.txt" ToolTip="Select the Proxy Number / Barcode CSV filename"
                        MaxFileInputsCount="1" OverwriteExistingFiles="True" runat="server" BackColor="#FFE7CE"
                        ControlObjectsVisibility="None" />
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
                    <label>
                        &nbsp;<asp:Label ID="lblFileName" runat="server" Font-Bold="True" Font-Names="Verdana"
                            Font-Size="Small" />
                        <asp:Label ID="lblNoResults" runat="server" Visible="false" Text="No file uploaded."
                            Font-Names="Verdana" Font-Size="X-Small" Font-Bold="true" />
                    </label>
                </td>
                <td>
                    &nbsp;</td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr id="trMaintenance03" runat="server" visible="false">
                <td>
                    &nbsp;<asp:Button ID="btnUploadBarcodesCSVFile" runat="server" OnClick="btnUploadBarcodesCSVFile_Click"
                        Text="Upload Barcodes CSV File" />

                </td>
                <td>

                &nbsp;
                            
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
        </table>
    </div>
    </form>
</body>
</html>
