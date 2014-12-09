<%@ Page Language="VB" Theme="AIMSDefault" ValidateRequest="false" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ Register TagPrefix="ComponentArt" Namespace="ComponentArt.Web.UI" Assembly="ComponentArt.Web.UI" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%@ Import Namespace="Microsoft.Win32" %>
<%@ Import Namespace="System.IO" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<script runat="server">

    Const AUDIT_TRAIL_FILENAME_VOSTRO As String = "C:/temp/CSN/CSNProcessor.html"
    Const AUDIT_TRAIL_FILENAME_PROD As String = "D:/CourierSoftware/masterlog/CSNProcessor/CSNProcessor.html"
    
    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsNumeric(Session("UserKey")) Then
            Server.Transfer("session_expired.aspx")
        End If
        If Not IsPostBack Then
            Call RefreshAuditTrail()
        End If
    End Sub
    
    Protected Sub btnConsignmentsShipped_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ConsignmentsShipped()
    End Sub

    Protected Sub ConsignmentsShipped()
        Dim dt As DataTable = ExecuteQueryToDataTable("SELECT * FROM ClientData_CSN_ConsignmentsShipped ORDER BY [id]")
        gvData.DataSource = dt
        gvData.DataBind()
        gvData.Visible = True
        lnkbtnHide.Visible = True
        lblTableName.Text = "ClientData_CSN_ConsignmentsShipped"
    End Sub
    
    Protected Sub btnControl_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowControl()
    End Sub

    Protected Sub ShowControl()
        Dim dt As DataTable = ExecuteQueryToDataTable("SELECT * FROM ClientData_CSN_Control ORDER BY [id]")
        gvData.DataSource = dt
        gvData.DataBind()
        gvData.Visible = True
        lnkbtnHide.Visible = True
        lblTableName.Text = "ClientData_CSN_Control"
    End Sub

    Protected Sub btnInventoryAdjustments_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim dt As DataTable = ExecuteQueryToDataTable("SELECT * FROM ClientData_CSN_InventoryAdjustments ORDER BY [id]")
        gvData.DataSource = dt
        gvData.DataBind()
        gvData.Visible = True
        lnkbtnHide.Visible = True
        lblTableName.Text = "ClientData_CSN_InventoryAdjustments"
    End Sub

    Protected Sub btnProductList_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim dt As DataTable = ExecuteQueryToDataTable("SELECT * FROM ClientData_CSN_ProductList ORDER BY [id]")
        gvData.DataSource = dt
        gvData.DataBind()
        gvData.Visible = True
        lnkbtnHide.Visible = True
        lblTableName.Text = "ClientData_CSN_ProductList"
    End Sub

    Protected Sub btnVPOs_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim dt As DataTable = ExecuteQueryToDataTable("SELECT * FROM ClientData_CSN_VendorPurchaseOrders ORDER BY [id]")
        gvData.DataSource = dt
        gvData.DataBind()
        gvData.Visible = True
        lnkbtnHide.Visible = True
        lblTableName.Text = "ClientData_CSN_VendorPurchaseOrders"
    End Sub

    Protected Sub btnVPOReceipts_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim dt As DataTable = ExecuteQueryToDataTable("SELECT * FROM ClientData_CSN_VPOReceipts ORDER BY [id]")
        gvData.DataSource = dt
        gvData.DataBind()
        gvData.Visible = True
        lnkbtnHide.Visible = True
        lblTableName.Text = "ClientData_CSN_VPOReceipts"
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

    Protected Sub lnkbtnRefreshAuditTrail_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call RefreshAuditTrail()
    End Sub

    Protected Sub RefreshAuditTrail()
        Dim sr As StreamReader
        Dim sFilename As String = String.Empty
        Try
            If Server.MachineName.ToLower.Contains("sprint") Then
                sFilename = AUDIT_TRAIL_FILENAME_PROD
            ElseIf Server.MachineName.ToLower.Contains("vostro") Then
                sFilename = AUDIT_TRAIL_FILENAME_VOSTRO
            Else
                WebMsgBox.Show("Cannot identify audit trail file to display")
                Exit Sub
            End If
            sr = New StreamReader(sFilename)
            lblAuditTrail.Text = sr.ReadToEnd()
            sr.Close()
            sr.Dispose()
            If lblAuditTrail.Text.Trim = String.Empty Then
                lblAuditTrail.Text = "(audit trail empty)"
            End If
        Catch Except As Exception
            lblAuditTrail.Text = "Audit trail not found"
        End Try
    End Sub
    
    Protected Sub lnkbtnClearAuditTrail_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim sFilename As String = String.Empty
        Try
            If Server.MachineName.ToLower.Contains("sprint") Then
                sFilename = AUDIT_TRAIL_FILENAME_PROD
            ElseIf Server.MachineName.ToLower.Contains("vostro") Then
                sFilename = AUDIT_TRAIL_FILENAME_VOSTRO
            Else
                WebMsgBox.Show("Cannot identify audit trail file to display")
                Exit Sub
            End If
            File.Delete(sFilename)
            Dim w As New StreamWriter(sFilename)
            w.WriteLine()
            w.Close()
            lblAuditTrail.Text = "(audit trail empty)"
        Catch ex As Exception
            WebMsgBox.Show("Could not clear audit trail " & sFilename & " (" & ex.Message & ")")
        End Try
    End Sub
    
    Protected Sub lnkbtnClearAllControlValues_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ExecuteQueryToDataTable("UPDATE ClientData_CSN_Control SET ControlValue = ''")
        Call ShowControl()
    End Sub
    
    Protected Sub ClearTable(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim lb As LinkButton = sender
        Call ExecuteQueryToDataTable("DELETE FROM " & lb.CommandArgument)
    End Sub
    
    Protected Sub ClearControlField(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim lb As LinkButton = sender
        Call ExecuteQueryToDataTable("UPDATE ClientData_CSN_Control SET ControlValue = '' WHERE ControlName = '" & lb.CommandArgument & "'")
        Call ShowControl()
    End Sub
    ''WebMsgBox.Show(Format("30/sep/2011 00:00:00", "dd-MMM-yyyy"))
    'WebMsgBox.Show(Format(Date.Parse("30/09/2011 00:00:00"), "dd-MMM-yyyy HH:mm:ss"))
    'Date.TryParse()
    'WebMsgBox.Show(IsDate("30/09/2011 00:00:00"))
    
    Protected Function GetRegistryValue(ByVal Hive As RegistryHive, ByVal Key As String, ByVal ValueName As String) As String
        Dim objParent As RegistryKey = Nothing
        Dim objSubkey As RegistryKey = Nothing
        Dim sAns As String
        Dim ErrInfo As String = String.Empty

        Select Case Hive
            Case RegistryHive.ClassesRoot
                objParent = Registry.ClassesRoot
            Case RegistryHive.CurrentConfig
                objParent = Registry.CurrentConfig
            Case RegistryHive.CurrentUser
                objParent = Registry.CurrentUser
            Case RegistryHive.LocalMachine
                objParent = Registry.LocalMachine
            Case RegistryHive.PerformanceData
                objParent = Registry.PerformanceData
            Case RegistryHive.Users
                objParent = Registry.Users
        End Select

        Try
            objSubkey = objParent.OpenSubKey(Key)
            'if can't be found, object is not initialized
            If Not objSubkey Is Nothing Then
                sAns = (objSubkey.GetValue(ValueName))
            End If
        Catch ex As Exception
            sAns = "Error"
            'ErrInfo = ex.Message
        Finally
            If ErrInfo = "" And sAns = "" Then
                sAns = "No value found for requested registry key"
            End If
        End Try
        GetRegistryValue = sAns
        
    End Function
    
    Protected Sub lnkbtnStaticConfiguration_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim sWayfairRegistryKeys() As String = {"CSNControllerProcessUserName", "CSNControllerProcessUserNameHelp", "CSNCustomerKey", "Debug", "DebugHelp", "FilesReceivedArchiveDirectory", "FilesSentArchiveDirectory", "FTPDirectoryInboundItemList", "FTPDirectoryInboundSalesOrders", "FTPDirectoryInboundVPOs", "FTPDirectoryOutboundASNs", "FTPDirectoryOutboundInventory", "FTPDirectoryOutboundInventoryAdjustments", "FTPDirectoryOutboundInvoices", "FTPDirectoryOutboundVPOReceipts", "FTPHelp1", "FTPHelp2", "FTPHost", "FTPPassword", "InventoryBalanceReportIntervalMins", "NextConsignmentCostReport", "NextInventoryBalanceReportDate", "NextInventoryBalanceReportMins", "NextInventoryReport", "NextReceiptReport", "OperationsEndTimeMins", "OperationsEndTimeMinsHelp", "OperationsStartTimeMins", "OperationsStartTimeMinsHelp", "PollIntervalSecs", "PollRunImmediate", "PollRunImmediateHelp", "RunMode", "RunModeHelp", "SupportEmailAddr", "SupportEmailAddrHelp"}
        Dim dtStaticConfiguration As New DataTable
        dtStaticConfiguration.Columns.Add(New DataColumn("Name", GetType(String)))
        dtStaticConfiguration.Columns.Add(New DataColumn("Value", GetType(String)))

        For Each s As String In sWayfairRegistryKeys
            Dim drStaticConfiguration As DataRow = dtStaticConfiguration.NewRow
            drStaticConfiguration("Name") = s
            drStaticConfiguration("Value") = GetRegistryValue(RegistryHive.LocalMachine, "SOFTWARE\CourierSoftware\CSN", s)
            dtStaticConfiguration.Rows.Add(drStaticConfiguration)
        Next
        
        gvData.DataSource = dtStaticConfiguration
        gvData.DataBind()
        gvData.Visible = True
        lnkbtnHide.Visible = True
        lblTableName.Text = "Static Configuration"
    End Sub
    
    Protected Sub btnRun_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Registry.SetValue("HKEY_LOCAL_MACHINE\SOFTWARE\CourierSoftware\CSN", "PollRunImmediate", "RUN_NOW!")
    End Sub
    
    Protected Sub lnkbtnHide_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        gvData.DataSource = Nothing
        gvData.Visible = False
        lnkbtnHide.Visible = False
        lblTableName.Text = String.Empty
    End Sub
    
    Protected Sub lnkbtnToggleTableSource_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim lb As LinkButton = sender
        If lb.Text.ToLower.Contains("show") Then
            lb.Text = "hide table source"
            pnlTableSource.Visible = True
        Else
            lb.Text = "show table source"
            pnlTableSource.Visible = False
        End If
    End Sub
    
    Protected Sub btnUnlock_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If tbPassword.Text.Trim.ToLower = "tw140rn" Then
            tbPassword.Visible = False
            btnUnlock.Visible = False
            lblLegendUnlocked.Visible = True
            btnRun.Enabled = True
            lnkbtnClearConsignmentsShipped.Enabled = True
            lnkbtnClearAllControlValues.Enabled = True
            lnkbtnClearNextConsignmentCostReportField.Enabled = True
            lnkbtnClearLastItemListReceivedField.Enabled = True
            lnkbtnClearLastVPOListReceivedField.Enabled = True
            lnkbtnClearLastVPOReceiptsSentField.Enabled = True
            lnkbtnClearLastCostAcctInvsSentField.Enabled = True
            lnkbtnClearLastInventoryAdjstsSentField.Enabled = True
            lnkbtnClearInventoryAdjustments.Enabled = True
            lnkbtnClearProductList.Enabled = True
            lnkbtnClearVPOs.Enabled = True
            lnkbtnClearVPOReceipts.Enabled = True
            lnkbtnClearAuditTrail.Enabled = True
            lnkbtnToggleTableSource.Visible = True
        Else
            WebMsgBox.Show("Incorrect password! Not unlocked!")
        End If
    End Sub
    
</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>CSN </title>
    </head>
<body>
    <form id="form1" runat="server">
    <main:Header ID="ctlHeader" runat="server" />
    <table width="100%" cellpadding="0" cellspacing="0">
        <tr class="bar_accounthandler">
            <td style="width: 50%; white-space: nowrap">
            </td>
            <td style="width: 50%; white-space: nowrap" align="right">
            </td>
        </tr>
    </table>
    <asp:Label ID="Label4" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Wayfair Dashboard" Font-Bold="True" />
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    <asp:LinkButton ID="lnkbtnStaticConfiguration" runat="server" Font-Names="Verdana" Font-Size="XX-Small" onclick="lnkbtnStaticConfiguration_Click">static configuration</asp:LinkButton>
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    <asp:Button ID="btnRun" runat="server" onclick="btnRun_Click" Text="Run Processing" Enabled="False" />
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    <asp:TextBox ID="tbPassword" runat="server" Font-Names="Verdana" Font-Size="XX-Small" MaxLength="10" TextMode="Password" Width="60px"></asp:TextBox>
&nbsp;<asp:Button ID="btnUnlock" runat="server" onclick="btnUnlock_Click" Text="Unlock" />
    <asp:Label ID="lblLegendUnlocked" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small" ForeColor="Red" Text="UNLOCKED" Visible="False"></asp:Label>
    <br />
    <br />
    <asp:Button ID="btnConsignmentsShipped" runat="server" Text="Consignments Shipped" Width="180px" OnClick="btnConsignmentsShipped_Click" />
    &nbsp;<asp:LinkButton ID="lnkbtnClearConsignmentsShipped" runat="server" Font-Names="Verdana" Font-Size="XX-Small" CommandArgument="ClientData_CSN_ConsignmentsShipped" onclick="ClearTable" Enabled="False">clear consignments shipped</asp:LinkButton>
    <br />
    <asp:Button ID="btnControl" runat="server" Text="Control" Width="180px" OnClick="btnControl_Click" />
    &nbsp;<asp:LinkButton ID="lnkbtnClearAllControlValues" runat="server" Font-Names="Verdana" Font-Size="XX-Small" onclick="lnkbtnClearAllControlValues_Click" Enabled="False">clear all control values</asp:LinkButton>
    <br />
    <asp:LinkButton ID="lnkbtnClearNextConsignmentCostReportField" runat="server" Font-Names="Verdana" Font-Size="XX-Small" CommandArgument="NEXT_CONSIGNMENT_COST_REPORT" onclick="ClearControlField" Enabled="False">clear NEXT_CONSIGNMENT_COST_REPORT</asp:LinkButton>
    &nbsp; <asp:LinkButton ID="lnkbtnClearLastItemListReceivedField" runat="server" Font-Names="Verdana" Font-Size="XX-Small" CommandArgument="LAST_ITEM_LIST_RECEIVED" onclick="ClearControlField" Enabled="False">clear LAST_ITEM_LIST_RECEIVED</asp:LinkButton>
    &nbsp; <asp:LinkButton ID="lnkbtnClearLastVPOListReceivedField" runat="server" Font-Names="Verdana" Font-Size="XX-Small" CommandArgument="LAST_VPO_LIST_RECEIVED" onclick="ClearControlField" Enabled="False">clear LAST_VPO_LIST_RECEIVED</asp:LinkButton>
    <br />
    <asp:LinkButton ID="lnkbtnClearLastVPOReceiptsSentField" runat="server" Font-Names="Verdana" Font-Size="XX-Small" CommandArgument="LAST_VPO_RECEIPTS_SENT" onclick="ClearControlField" Enabled="False">clear LAST_VPO_RECEIPTS_SENT</asp:LinkButton>
    &nbsp; <asp:LinkButton ID="lnkbtnClearLastCostAcctInvsSentField" runat="server" Font-Names="Verdana" Font-Size="XX-Small" CommandArgument="LAST_COST_ACCT_INVS_SENT" onclick="ClearControlField" Enabled="False">clear LAST_COST_ACCT_INVS_SENT</asp:LinkButton>
    &nbsp; <asp:LinkButton ID="lnkbtnClearLastInventoryAdjstsSentField" runat="server" Font-Names="Verdana" Font-Size="XX-Small" CommandArgument="LAST_INVENTORY_ADJSTS_SENT" onclick="ClearControlField" Enabled="False">clear LAST_INVENTORY_ADJSTS_SENT</asp:LinkButton>
    <br />
    <asp:Button ID="btnInventoryAdjustments" runat="server" Text="Inventory Adjustments" Width="180px" OnClick="btnInventoryAdjustments_Click" />
    &nbsp;<asp:LinkButton ID="lnkbtnClearInventoryAdjustments" runat="server" Font-Names="Verdana" Font-Size="XX-Small" CommandArgument="ClientData_CSN_InventoryAdjustments" onclick="ClearTable" Enabled="False">clear inventory adjustments</asp:LinkButton>
    <br />
    <asp:Button ID="btnProductList" runat="server" Text="Product List" Width="180px" OnClick="btnProductList_Click" />
    &nbsp;<asp:LinkButton ID="lnkbtnClearProductList" runat="server" Font-Names="Verdana" Font-Size="XX-Small" CommandArgument="ClientData_CSN_ProductList" onclick="ClearTable" Enabled="False">clear product list</asp:LinkButton>
    <br />
    <asp:Button ID="btnVPOs" runat="server" Text="VPOs" Width="180px" OnClick="btnVPOs_Click" />
    &nbsp;<asp:LinkButton ID="lnkbtnClearVPOs" runat="server" Font-Names="Verdana" Font-Size="XX-Small" CommandArgument="ClientData_CSN_VendorPurchaseOrders" onclick="ClearTable" Enabled="False">clear VPOs</asp:LinkButton>
    <br />
    <asp:Button ID="btnVPOReceipts" runat="server" Text="VPO Receipts" Width="180px" OnClick="btnVPOReceipts_Click" />
    &nbsp;<asp:LinkButton ID="lnkbtnClearVPOReceipts" runat="server" Font-Names="Verdana" Font-Size="XX-Small" CommandArgument="ClientData_CSN_VPOReceipts" onclick="ClearTable" Enabled="False">clear VPO receipts</asp:LinkButton>
    <br />
    <asp:Label ID="lblTableName" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small" />
    &nbsp;<asp:LinkButton ID="lnkbtnHide" runat="server" Font-Names="Verdana" Font-Size="XX-Small" onclick="lnkbtnHide_Click" Visible="False">hide</asp:LinkButton>
    <br />
    <asp:GridView ID="gvData" runat="server" CellPadding="2" Width="100%" Font-Names="Verdana" Font-Size="XX-Small">
        <EmptyDataTemplate>
            <div style="text-align: center">
                <asp:Label ID="Label13" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small" ForeColor="Red" Text="table is empty" />
            </div>
        </EmptyDataTemplate>
    </asp:GridView>
    <p>
        <asp:LinkButton ID="lnkbtnRefreshAuditTrail" runat="server" Font-Names="Verdana" Font-Size="XX-Small" OnClick="lnkbtnRefreshAuditTrail_Click">refresh audit trail</asp:LinkButton>
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <asp:LinkButton ID="lnkbtnClearAuditTrail" runat="server" Font-Names="Verdana" Font-Size="XX-Small" OnClick="lnkbtnClearAuditTrail_Click" Enabled="False">clear audit trail</asp:LinkButton>
    &nbsp;</p>
    <p>
        <asp:Label ID="lblAuditTrail" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small" Width="100%" />
    </p>
    <asp:LinkButton ID="lnkbtnToggleTableSource" runat="server" Font-Names="Verdana" Font-Size="XX-Small" onclick="lnkbtnToggleTableSource_Click" Visible="False">show table source</asp:LinkButton>
&nbsp;<asp:Panel ID="pnlTableSource" runat="server" Visible="False">
        <span id="internal-source-marker_0.7638791683836861" style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">USE Logistics</span><br />
        <br />
        <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id(N&#39;[dbo].[ClientData_CSN_ProductList]&#39;) AND OBJECTPROPERTY(id, N&#39;IsUserTable&#39;) = 1)</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">DROP TABLE [dbo].[ClientData_CSN_ProductList]</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">GO</span><br />
        <br />
        <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">CREATE TABLE [dbo].[ClientData_CSN_ProductList](</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [id] [int] IDENTITY(1,1) NOT NULL,</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [SUID] [varchar](55) NOT NULL,</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [PartNo] [varchar](55) NOT NULL,</span><br />
        <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [SKUDescription] [varchar](55) NOT NULL,</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [OptionSetDescription] [varchar](55) NOT NULL,</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [TrueSupplierSUID] [varchar](55) NOT NULL,</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [TrueSupplierName] [varchar](55) NOT NULL,</span><br />
        <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [VendorPartNo] [varchar](55) NOT NULL,</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [CSNSKU] [varchar](55) NOT NULL,</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [Inactive] [int] NOT NULL,</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [Discontinued] [int] NOT NULL,</span><br />
        <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [CreatedOn] [smalldatetime] NOT NULL,</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [CreatedBy] [int] NOT NULL,</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">CONSTRAINT [PK_ClientData_CSN_ProductList] PRIMARY KEY CLUSTERED</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">(</span><br />
        <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [PartNo] ASC</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">)WITH (PAD_INDEX &nbsp;= OFF, STATISTICS_NORECOMPUTE &nbsp;= OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS &nbsp;= ON, ALLOW_PAGE_LOCKS &nbsp;= ON) ON [PRIMARY]</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">) ON [PRIMARY]</span><br />
        <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">GO</span><br />
        <br />
        <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">GRANT &nbsp;SELECT , &nbsp;INSERT , &nbsp;DELETE , &nbsp;UPDATE &nbsp;ON [dbo].[ClientData_CSN_ProductList] &nbsp;TO [LogisticsUserRole]</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">GO</span><br />
        <br />
        <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">GRANT &nbsp;SELECT , &nbsp;INSERT , &nbsp;DELETE , &nbsp;UPDATE &nbsp;ON [dbo].[ClientData_CSN_ProductList] &nbsp;TO [LogisticsAdminRole]</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">GO</span><br />
        <br />
        <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id(N&#39;[dbo].[ClientData_CSN_VendorPurchaseOrders]&#39;) AND OBJECTPROPERTY(id, N&#39;IsUserTable&#39;) = 1)</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">DROP TABLE [dbo].[ClientData_CSN_VendorPurchaseOrders]</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">GO</span><br />
        <br />
        <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">CREATE TABLE [dbo].[ClientData_CSN_VendorPurchaseOrders](</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [id] [int] IDENTITY(1,1) NOT NULL,</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [CreatedOn] [smalldatetime] NOT NULL,</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [SupplierID] [varchar](55) NOT NULL,</span><br />
        <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [SupplierName] [varchar](55) NOT NULL,</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [SPONumber] [varchar](30) NOT NULL,</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [SPOSentDate] [smalldatetime] NOT NULL,</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [SPOEstShipDate] [smalldatetime] NOT NULL,</span><br />
        <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [SupplierPartNumber] [varchar](55) NOT NULL,</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [CSNSKU] [varchar](55) NOT NULL,</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [SKUDescription] [varchar](55) NOT NULL,</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [WSCost] [money] NOT NULL,</span><br />
        <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [QuantityOrdered] [int] NOT NULL,</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [TotalQuantityReceived] [int] NOT NULL,</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [RemainingOnOrder] [int] NOT NULL,</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [Closed] [char](1) NOT NULL,</span><br />
        <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [CreatedBy] [int] NOT NULL</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">) ON [PRIMARY]</span><br />
        <br />
        <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">GRANT &nbsp;SELECT , &nbsp;INSERT , &nbsp;DELETE , &nbsp;UPDATE &nbsp;ON [dbo].[ClientData_CSN_VendorPurchaseOrders] &nbsp;TO [LogisticsUserRole]</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">GO</span><br />
        <br />
        <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">GRANT &nbsp;SELECT , &nbsp;INSERT , &nbsp;DELETE , &nbsp;UPDATE &nbsp;ON [dbo].[ClientData_CSN_VendorPurchaseOrders] &nbsp;TO [LogisticsAdminRole]</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">GO</span><br />
        <br />
        <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id(N&#39;[dbo].[ClientData_CSN_VPOReceipts]&#39;) AND OBJECTPROPERTY(id, N&#39;IsUserTable&#39;) = 1)</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">DROP TABLE [dbo].[ClientData_CSN_VPOReceipts]</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">GO</span><br />
        <br />
        <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">CREATE TABLE [dbo].[ClientData_CSN_VPOReceipts](</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [id] [int] IDENTITY(1,1) NOT NULL,</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [DateReceived] [smalldatetime] NOT NULL,</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [PONumber] [varchar](30) NOT NULL,</span><br />
        <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [SupplierID] [varchar](30) NOT NULL,</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [SupplierPartNumber] [varchar](30) NOT NULL,</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [QuantityReceived] [int] NOT NULL,</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [CreatedBy] [int] NOT NULL</span><br />
        <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">) ON [PRIMARY]</span><br />
        <br />
        <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">GRANT &nbsp;SELECT , &nbsp;INSERT , &nbsp;DELETE , &nbsp;UPDATE &nbsp;ON [dbo].[ClientData_CSN_VPOReceipts] &nbsp;TO [LogisticsUserRole]</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">GO</span><br />
        <br />
        <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">GRANT &nbsp;SELECT , &nbsp;INSERT , &nbsp;DELETE , &nbsp;UPDATE &nbsp;ON [dbo].[ClientData_CSN_VPOReceipts] &nbsp;TO [LogisticsAdminRole]</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">GO</span><br />
        <br />
        <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id(N&#39;[dbo].[ClientData_CSN_InventoryAdjustments]&#39;) AND OBJECTPROPERTY(id, N&#39;IsUserTable&#39;) = 1)</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">DROP TABLE [dbo].[ClientData_CSN_InventoryAdjustments]</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">GO</span><br />
        <br />
        <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">CREATE TABLE [dbo].[ClientData_CSN_InventoryAdjustments](</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [id] [int] IDENTITY(1,1) NOT NULL,</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [CreatedOn] [smalldatetime] NOT NULL,</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [SupplierID] [varchar](55) NOT NULL,</span><br />
        <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [PartNumber] [varchar](55) NOT NULL,</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [AdjustmentQuantity] [int] NOT NULL,</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [SPONumber] [varchar](30) NOT NULL,</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [Comment] [varchar](200) NOT NULL,</span><br />
        <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [CreatedBy] [int] NOT NULL</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">) ON [PRIMARY]</span><br />
        <br />
        <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">GRANT &nbsp;SELECT , &nbsp;INSERT , &nbsp;DELETE , &nbsp;UPDATE &nbsp;ON [dbo].[ClientData_CSN_InventoryAdjustments] &nbsp;TO [LogisticsUserRole]</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">GO</span><br />
        <br />
        <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">GRANT &nbsp;SELECT , &nbsp;INSERT , &nbsp;DELETE , &nbsp;UPDATE &nbsp;ON [dbo].[ClientData_CSN_InventoryAdjustments] &nbsp;TO [LogisticsAdminRole]</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">GO</span><br />
        <br />
        <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id(N&#39;[dbo].[ClientData_CSN_AuditTrail]&#39;) AND OBJECTPROPERTY(id, N&#39;IsUserTable&#39;) = 1)</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">DROP TABLE [dbo].[ClientData_CSN_AuditTrail]</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">GO</span><br />
        <br />
        <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">CREATE TABLE [dbo].[ClientData_CSN_AuditTrail](</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [id] [int] IDENTITY(1,1) NOT NULL,</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [CreatedOn] [smalldatetime] NOT NULL,</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [Code] [varchar](30) NOT NULL,</span><br />
        <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [Description] [varchar](1000) NOT NULL,</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [CreatedBy] [int] NOT NULL</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">) ON [PRIMARY]</span><br />
        <br />
        <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">GRANT &nbsp;SELECT , &nbsp;INSERT , &nbsp;DELETE , &nbsp;UPDATE &nbsp;ON [dbo].[ClientData_CSN_AuditTrail] &nbsp;TO [LogisticsUserRole]</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">GO</span><br />
        <br />
        <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">GRANT &nbsp;SELECT , &nbsp;INSERT , &nbsp;DELETE , &nbsp;UPDATE &nbsp;ON [dbo].[ClientData_CSN_AuditTrail] &nbsp;TO [LogisticsAdminRole]</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">GO</span><br />
        <br />
        <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id(N&#39;[dbo].[ClientData_CSN_ConsignmentsShipped]&#39;) AND OBJECTPROPERTY(id, N&#39;IsUserTable&#39;) = 1)</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">DROP TABLE [dbo].[ClientData_CSN_ConsignmentsShipped]</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">GO</span><br />
        <br />
        <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">CREATE TABLE [dbo].[ClientData_CSN_ConsignmentsShipped](</span><br />
        <br />
        <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [id] [int] IDENTITY(1,1) NOT NULL,</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [ConsignmentKey] [int] NOT NULL,</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [CreatedOn] [smalldatetime] NOT NULL</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">) ON [PRIMARY]</span><br />
        <br />
        <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">GRANT &nbsp;SELECT , &nbsp;INSERT , &nbsp;DELETE , &nbsp;UPDATE &nbsp;ON [dbo].[ClientData_CSN_ConsignmentsShipped] &nbsp;TO [LogisticsUserRole]</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">GO</span><br />
        <br />
        <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">GRANT &nbsp;SELECT , &nbsp;INSERT , &nbsp;DELETE , &nbsp;UPDATE &nbsp;ON [dbo].[ClientData_CSN_ConsignmentsShipped] &nbsp;TO [LogisticsAdminRole]</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">GO</span><br />
        <br />
        <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id(N&#39;[dbo].[ClientData_CSN_Control]&#39;) AND OBJECTPROPERTY(id, N&#39;IsUserTable&#39;) = 1)</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">DROP TABLE [dbo].[ClientData_CSN_Control]</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">GO</span><br />
        <br />
        <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">CREATE TABLE [dbo].[ClientData_CSN_Control](</span><br />
        <br />
        <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [id] [int] IDENTITY(1,1) NOT NULL,</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [ControlName] [varchar](50) NOT NULL,</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">&nbsp;&nbsp;&nbsp; [ControlValue] [varchar](100) NULL</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">) ON [PRIMARY]</span><br />
        <br />
        <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">GRANT &nbsp;SELECT , &nbsp;INSERT , &nbsp;DELETE , &nbsp;UPDATE &nbsp;ON [dbo].[ClientData_CSN_Control] &nbsp;TO [LogisticsUserRole]</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">GO</span><br />
        <br />
        <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">GRANT &nbsp;SELECT , &nbsp;INSERT , &nbsp;DELETE , &nbsp;UPDATE &nbsp;ON [dbo].[ClientData_CSN_Control] &nbsp;TO [LogisticsAdminRole]</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">GO</span><br />
        <br />
        <br />
        <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">-- WORKER_PROCESS_INTERVAL</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">-- NEXT_CONSIGNMENT_COST_REPORT</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">-- NEXT_INVENTORY_REPORT</span><br /> <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">-- INVENTORY_REPORT_INTERVAL</span><br />
        <span style="font-size:11pt;font-family:Arial;color:#000000;background-color:transparent;font-weight:normal;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;">-- NEXT_RECEIPT_REPORT</span><br />
        <br />
    </asp:Panel>
    </form>
</body>
</html>
