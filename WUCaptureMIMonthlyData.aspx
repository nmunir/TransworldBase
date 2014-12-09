<%@ Page Language="VB" Theme="AIMSDefault" %>

<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ Register Assembly="Telerik.Web.UI" Namespace="Telerik.Web.UI" TagPrefix="telerik" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<script runat="server">

    Const ITEMS_PER_REQUEST As Integer = 30

    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        'If Not IsNumeric(Session("UserKey")) Then
        '    Server.Transfer("session_expired.aspx")
        'End If
        If Not IsPostBack Then
            Call SetTitle()
        End If
    End Sub

    Protected Sub SetTitle()
        Dim sTitle As String = Session("SiteTitle")
        If sTitle <> String.Empty Then
            sTitle += " - "
        End If
        Page.Header.Title = sTitle & "Capture Management Information Monthly Data"
    End Sub

    Protected Function ExecuteStoredProcedureToDataTable(ByVal sp_name As String, Optional ByVal IListPrams As List(Of SqlParameter) = Nothing) As DataTable
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
    
    Protected Function ExecuteQueryToDataTable(ByVal sQuery As String) As DataTable
        Dim oDataTable As New DataTable
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter(sQuery, oConn)
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

    Protected Sub CreateEntry()
        Dim sSQL As String
        If rbUK.Checked Then
            sSQL = "INSERT INTO ClientData_WU_MIMonthlyReport (Year, Month, Country, VisibleToClient, OrderBreakdownOperations, OrderBreakdownMarketing, OrderBreakdownFININT, OrderBreakdownCosta, OrderBreakdownPrePaid, StorageCostsOperations, StorageCostsMarketing, StorageCostsFININT, StorageCostsCosta, StorageCostsPrePaid, LogisticsCostsCourierOperations, LogisticsCostsCourierMarketing, LogisticsCostsCourierFININT, LogisticsCostsCourierCosta, LogisticsCostsPrepaid, LogisticsCostsMailFulfilment, LogisticsCostsAdHocFulfilment, ServiceFeesPickFeesOperations, ServiceFeesPickFeesMarketing, ServiceFeesPickFeesFININT, ServiceFeesPickFeesCosta, ServiceFeesPickFeesPrePaid, ServiceFeesGoodsInOperations, ServiceFeesGoodsInMarketing, ServiceFeesGoodsInFININT, ServiceFeesGoodsInCosta, ServiceFeesGoodsInPrePaid, ServiceFeesDestructionFeesOperations, ServiceFeesDestructionFeesMarketing, ServiceFeesDestructionFeesFININT, ServiceFeesDestructionFeesCosta, ServiceFeesDestructionFeesPrePaid, ServiceFeesManagementFee, InternalNotes, ClientNotes, LastUpdateOn, LastUpdatedBy) VALUES (" & ddlYear.SelectedValue & ", " & ddlMonth.SelectedValue & ", 'UK', 0,  -1,-1,-1,-1,-1,-1,-1,-1,-1,-1,  -1,-1,-1,-1,-1,-1,-1,-1,-1,-1,   -1,-1,-1,-1,-1,-1,-1,-1,-1,-1,  -1,-1, -1,  '', '', GETDATE(), 0)"
        Else
            sSQL = "INSERT INTO ClientData_WU_MIMonthlyReport (Year, Month, Country, VisibleToClient, OrderBreakdownOperations, OrderBreakdownMarketing, OrderBreakdownFININT, OrderBreakdownCosta, OrderBreakdownPrePaid, StorageCostsOperations, StorageCostsMarketing, StorageCostsFININT, StorageCostsCosta, StorageCostsPrePaid, LogisticsCostsCourierOperations, LogisticsCostsCourierMarketing, LogisticsCostsCourierFININT, LogisticsCostsCourierCosta, LogisticsCostsPrepaid, LogisticsCostsMailFulfilment, LogisticsCostsAdHocFulfilment, ServiceFeesPickFeesOperations, ServiceFeesPickFeesMarketing, ServiceFeesPickFeesFININT, ServiceFeesPickFeesCosta, ServiceFeesPickFeesPrePaid, ServiceFeesGoodsInOperations, ServiceFeesGoodsInMarketing, ServiceFeesGoodsInFININT, ServiceFeesGoodsInCosta, ServiceFeesGoodsInPrePaid, ServiceFeesDestructionFeesOperations, ServiceFeesDestructionFeesMarketing, ServiceFeesDestructionFeesFININT, ServiceFeesDestructionFeesCosta, ServiceFeesDestructionFeesPrePaid, ServiceFeesManagementFee, InternalNotes, ClientNotes, LastUpdateOn, LastUpdatedBy) VALUES (" & ddlYear.SelectedValue & ", " & ddlMonth.SelectedValue & ", 'IRELAND', 0,  -1,-1,-1,-1,-1,-1,-1,-1,-1,-1,  -1,-1,-1,-1,-1,-1,-1,-1,-1,-1,   -1,-1,-1,-1,-1,-1,-1,-1,-1,-1,  -1,-1, -1,  '', '', GETDATE(), 0)"
        End If
        Call ExecuteQueryToDataTable(sSQL)
    End Sub
    
    Protected Sub CreateEntry(sCountry As String)
        Dim sSQL As String
        sSQL = "INSERT INTO ClientData_WU_MIMonthlyReport (Year, Month, Country, VisibleToClient, OrderBreakdownOperations, OrderBreakdownMarketing, OrderBreakdownFININT, OrderBreakdownCosta, OrderBreakdownPrePaid, StorageCostsOperations, StorageCostsMarketing, StorageCostsFININT, StorageCostsCosta, StorageCostsPrePaid, LogisticsCostsCourierOperations, LogisticsCostsCourierMarketing, LogisticsCostsCourierFININT, LogisticsCostsCourierCosta, LogisticsCostsPrepaid, LogisticsCostsMailFulfilment, LogisticsCostsAdHocFulfilment, ServiceFeesPickFeesOperations, ServiceFeesPickFeesMarketing, ServiceFeesPickFeesFININT, ServiceFeesPickFeesCosta, ServiceFeesPickFeesPrePaid, ServiceFeesGoodsInOperations, ServiceFeesGoodsInMarketing, ServiceFeesGoodsInFININT, ServiceFeesGoodsInCosta, ServiceFeesGoodsInPrePaid, ServiceFeesDestructionFeesOperations, ServiceFeesDestructionFeesMarketing, ServiceFeesDestructionFeesFININT, ServiceFeesDestructionFeesCosta, ServiceFeesDestructionFeesPrePaid, ServiceFeesManagementFee, InternalNotes, ClientNotes, LastUpdateOn, LastUpdatedBy) VALUES (" & ddlYear.SelectedValue & ", " & ddlMonth.SelectedValue & ", '" & sCountry & "', 0,  -1,-1,-1,-1,-1,-1,-1,-1,-1,-1,  -1,-1,-1,-1,-1,-1,-1,-1,-1,-1,   -1,-1,-1,-1,-1,-1,-1,-1,-1,-1,  -1,-1, -1, '', '', GETDATE(), 0)"
        Call ExecuteQueryToDataTable(sSQL)
    End Sub
    
    Protected Function UnselectedCountry() As String
        If rbUK.Checked Then
            UnselectedCountry = "IRELAND"
        Else
            UnselectedCountry = "UK"
        End If
    End Function
    
    Protected Sub btnUpdate_Click(sender As Object, e As System.EventArgs)
        If RetrieveEntry() Is Nothing Then
            Call CreateEntry()
        End If

        If RetrieveEntry(UnselectedCountry) Is Nothing Then
            Call CreateEntry(UnselectedCountry)
        End If

        Dim sbSQL As New StringBuilder
        sbSQL.Append("UPDATE ClientData_WU_MIMonthlyReport ")
        sbSQL.Append("SET ")
        
        sbSQL.Append("VisibleToClient")
        sbSQL.Append(" = ")

        If cbVisibleToClient.Checked Then
            sbSQL.Append("1")
        Else
            sbSQL.Append("0")
        End If
        sbSQL.Append(", ")
        
        sbSQL.Append("OrderBreakdownOperations")
        sbSQL.Append(" = ")
        If rntbOrderBreakdownOperations.Text = String.Empty Then
            sbSQL.Append("-1")
        Else
            sbSQL.Append(rntbOrderBreakdownOperations.Text.Replace("'", "''"))
        End If
        sbSQL.Append(", ")

        sbSQL.Append("OrderBreakdownMarketing")
        sbSQL.Append(" = ")
        If rntbOrderBreakdownMarketing.Text = String.Empty Then
            sbSQL.Append("-1")
        Else
            sbSQL.Append(rntbOrderBreakdownMarketing.Text.Replace("'", "''"))
        End If
        sbSQL.Append(", ")

        sbSQL.Append("OrderBreakdownFININT")
        sbSQL.Append(" = ")
        If rntbOrderBreakdownFININT.Text = String.Empty Then
            sbSQL.Append("-1")
        Else
            sbSQL.Append(rntbOrderBreakdownFININT.Text.Replace("'", "''"))
        End If
        sbSQL.Append(", ")

        sbSQL.Append("OrderBreakdownCosta")
        sbSQL.Append(" = ")
        If rntbOrderBreakdownCosta.Text = String.Empty Then
            sbSQL.Append("-1")
        Else
            sbSQL.Append(rntbOrderBreakdownCosta.Text.Replace("'", "''"))
        End If
        sbSQL.Append(", ")

        sbSQL.Append("OrderBreakdownPrePaid")
        sbSQL.Append(" = ")
        If rntbOrderBreakdownPrePaid.Text = String.Empty Then
            sbSQL.Append("-1")
        Else
            sbSQL.Append(rntbOrderBreakdownPrePaid.Text.Replace("'", "''"))
        End If
        sbSQL.Append(", ")

        sbSQL.Append("StorageCostsOperations")
        sbSQL.Append(" = ")
        If rntbStorageCostsOperations.Text = String.Empty Then
            sbSQL.Append("-1")
        Else
            sbSQL.Append(rntbStorageCostsOperations.Text.Replace("'", "''"))
        End If
        sbSQL.Append(", ")

        sbSQL.Append("StorageCostsMarketing")
        sbSQL.Append(" = ")
        If rntbStorageCostsMarketing.Text = String.Empty Then
            sbSQL.Append("-1")
        Else
            sbSQL.Append(rntbStorageCostsMarketing.Text.Replace("'", "''"))
        End If
        sbSQL.Append(", ")

        sbSQL.Append("StorageCostsFININT")
        sbSQL.Append(" = ")
        If rntbStorageCostsFININT.Text = String.Empty Then
            sbSQL.Append("-1")
        Else
            sbSQL.Append(rntbStorageCostsFININT.Text.Replace("'", "''"))
        End If
        sbSQL.Append(", ")

        sbSQL.Append("StorageCostsCosta")
        sbSQL.Append(" = ")
        If rntbStorageCostsCosta.Text = String.Empty Then
            sbSQL.Append("-1")
        Else
            sbSQL.Append(rntbStorageCostsCosta.Text.Replace("'", "''"))
        End If
        sbSQL.Append(", ")

        sbSQL.Append("StorageCostsPrePaid")
        sbSQL.Append(" = ")
        If rntbStorageCostsPrePaid.Text = String.Empty Then
            sbSQL.Append("-1")
        Else
            sbSQL.Append(rntbStorageCostsPrePaid.Text.Replace("'", "''"))
        End If
        sbSQL.Append(", ")

        sbSQL.Append("LogisticsCostsCourierOperations")
        sbSQL.Append(" = ")
        If rntbLogisticsCostsCourierOperations.Text = String.Empty Then
            sbSQL.Append("-1")
        Else
            sbSQL.Append(rntbLogisticsCostsCourierOperations.Text.Replace("'", "''"))
        End If
        sbSQL.Append(", ")

        sbSQL.Append("LogisticsCostsCourierMarketing")
        sbSQL.Append(" = ")
        If rntbLogisticsCostsCourierMarketing.Text = String.Empty Then
            sbSQL.Append("-1")
        Else
            sbSQL.Append(rntbLogisticsCostsCourierMarketing.Text.Replace("'", "''"))
        End If
        sbSQL.Append(", ")

        sbSQL.Append("LogisticsCostsCourierFININT")
        sbSQL.Append(" = ")
        If rntbLogisticsCostsCourierFININT.Text = String.Empty Then
            sbSQL.Append("-1")
        Else
            sbSQL.Append(rntbLogisticsCostsCourierFININT.Text.Replace("'", "''"))
        End If
        sbSQL.Append(", ")

        sbSQL.Append("LogisticsCostsCourierCosta")
        sbSQL.Append(" = ")
        If rntbLogisticsCostsCourierCosta.Text = String.Empty Then
            sbSQL.Append("-1")
        Else
            sbSQL.Append(rntbLogisticsCostsCourierCosta.Text.Replace("'", "''"))
        End If
        sbSQL.Append(", ")

        sbSQL.Append("LogisticsCostsPrepaid")
        sbSQL.Append(" = ")
        If rntbLogisticsCostsPrepaid.Text = String.Empty Then
            sbSQL.Append("-1")
        Else
            sbSQL.Append(rntbLogisticsCostsPrepaid.Text.Replace("'", "''"))
        End If
        sbSQL.Append(", ")

        sbSQL.Append("LogisticsCostsMailFulfilment")
        sbSQL.Append(" = ")
        If rntbLogisticsCostsMailFulfilment.Text = String.Empty Then
            sbSQL.Append("-1")
        Else
            sbSQL.Append(rntbLogisticsCostsMailFulfilment.Text.Replace("'", "''"))
        End If
        sbSQL.Append(", ")

        sbSQL.Append("LogisticsCostsAdHocFulfilment")
        sbSQL.Append(" = ")
        If rntbLogisticsCostsAdHocFulfilment.Text = String.Empty Then
            sbSQL.Append("-1")
        Else
            sbSQL.Append(rntbLogisticsCostsAdHocFulfilment.Text.Replace("'", "''"))
        End If
        sbSQL.Append(", ")

        sbSQL.Append("ServiceFeesPickFeesOperations")
        sbSQL.Append(" = ")
        If rntbServiceFeesPickFeesOperations.Text = String.Empty Then
            sbSQL.Append("-1")
        Else
            sbSQL.Append(rntbServiceFeesPickFeesOperations.Text.Replace("'", "''"))
        End If
        sbSQL.Append(", ")

        sbSQL.Append("ServiceFeesPickFeesMarketing")
        sbSQL.Append(" = ")
        If rntbServiceFeesPickFeesMarketing.Text = String.Empty Then
            sbSQL.Append("-1")
        Else
            sbSQL.Append(rntbServiceFeesPickFeesMarketing.Text.Replace("'", "''"))
        End If
        sbSQL.Append(", ")

        sbSQL.Append("ServiceFeesPickFeesFININT")
        sbSQL.Append(" = ")
        If rntbServiceFeesPickFeesFININT.Text = String.Empty Then
            sbSQL.Append("-1")
        Else
            sbSQL.Append(rntbServiceFeesPickFeesFININT.Text.Replace("'", "''"))
        End If
        sbSQL.Append(", ")

        sbSQL.Append("ServiceFeesPickFeesCosta")
        sbSQL.Append(" = ")
        If rntbServiceFeesPickFeesCosta.Text = String.Empty Then
            sbSQL.Append("-1")
        Else
            sbSQL.Append(rntbServiceFeesPickFeesCosta.Text.Replace("'", "''"))
        End If
        sbSQL.Append(", ")

        sbSQL.Append("ServiceFeesPickFeesPrePaid")
        sbSQL.Append(" = ")
        If rntbServiceFeesPickFeesPrePaid.Text = String.Empty Then
            sbSQL.Append("-1")
        Else
            sbSQL.Append(rntbServiceFeesPickFeesPrePaid.Text.Replace("'", "''"))
        End If
        sbSQL.Append(", ")

        sbSQL.Append("ServiceFeesGoodsInOperations")
        sbSQL.Append(" = ")
        If rntbServiceFeesGoodsInOperations.Text = String.Empty Then
            sbSQL.Append("-1")
        Else
            sbSQL.Append(rntbServiceFeesGoodsInOperations.Text.Replace("'", "''"))
        End If
        sbSQL.Append(", ")

        sbSQL.Append("ServiceFeesGoodsInMarketing")
        sbSQL.Append(" = ")
        If rntbServiceFeesGoodsInMarketing.Text = String.Empty Then
            sbSQL.Append("-1")
        Else
            sbSQL.Append(rntbServiceFeesGoodsInMarketing.Text.Replace("'", "''"))
        End If
        sbSQL.Append(", ")

        sbSQL.Append("ServiceFeesGoodsInFININT")
        sbSQL.Append(" = ")
        If rntbServiceFeesGoodsInFININT.Text = String.Empty Then
            sbSQL.Append("-1")
        Else
            sbSQL.Append(rntbServiceFeesGoodsInFININT.Text.Replace("'", "''"))
        End If
        sbSQL.Append(", ")

        sbSQL.Append("ServiceFeesGoodsInCosta")
        sbSQL.Append(" = ")
        If rntbServiceFeesGoodsInCosta.Text = String.Empty Then
            sbSQL.Append("-1")
        Else
            sbSQL.Append(rntbServiceFeesGoodsInCosta.Text.Replace("'", "''"))
        End If
        sbSQL.Append(", ")

        sbSQL.Append("ServiceFeesGoodsInPrePaid")
        sbSQL.Append(" = ")
        If rntbServiceFeesGoodsInPrePaid.Text = String.Empty Then
            sbSQL.Append("-1")
        Else
            sbSQL.Append(rntbServiceFeesGoodsInPrePaid.Text.Replace("'", "''"))
        End If
        sbSQL.Append(", ")

        sbSQL.Append("ServiceFeesDestructionFeesOperations")
        sbSQL.Append(" = ")
        If rntbServiceFeesDestructionFeesOperations.Text = String.Empty Then
            sbSQL.Append("-1")
        Else
            sbSQL.Append(rntbServiceFeesDestructionFeesOperations.Text.Replace("'", "''"))
        End If
        sbSQL.Append(", ")

        sbSQL.Append("ServiceFeesDestructionFeesMarketing")
        sbSQL.Append(" = ")
        If rntbServiceFeesDestructionFeesMarketing.Text = String.Empty Then
            sbSQL.Append("-1")
        Else
            sbSQL.Append(rntbServiceFeesDestructionFeesMarketing.Text.Replace("'", "''"))
        End If
        sbSQL.Append(", ")

        sbSQL.Append("ServiceFeesDestructionFeesFININT")
        sbSQL.Append(" = ")
        If rntbServiceFeesDestructionFeesFININT.Text = String.Empty Then
            sbSQL.Append("-1")
        Else
            sbSQL.Append(rntbServiceFeesDestructionFeesFININT.Text.Replace("'", "''"))
        End If
        sbSQL.Append(", ")

        sbSQL.Append("ServiceFeesDestructionFeesCosta")
        sbSQL.Append(" = ")
        If rntbServiceFeesDestructionFeesCosta.Text = String.Empty Then
            sbSQL.Append("-1")
        Else
            sbSQL.Append(rntbServiceFeesDestructionFeesCosta.Text.Replace("'", "''"))
        End If
        sbSQL.Append(", ")

        sbSQL.Append("ServiceFeesDestructionFeesPrePaid")
        sbSQL.Append(" = ")
        If rntbServiceFeesDestructionFeesPrePaid.Text = String.Empty Then
            sbSQL.Append("-1")
        Else
            sbSQL.Append(rntbServiceFeesDestructionFeesPrePaid.Text.Replace("'", "''"))
        End If
        sbSQL.Append(", ")

        sbSQL.Append("ServiceFeesManagementFee")
        sbSQL.Append(" = ")
        If rntbServiceFeesManagementFee.Text = String.Empty Then
            sbSQL.Append("-1")
        Else
            sbSQL.Append(rntbServiceFeesManagementFee.Text.Replace("'", "''"))
        End If
        sbSQL.Append(", ")

        sbSQL.Append("InternalNotes")
        sbSQL.Append(" = '")
        sbSQL.Append(reInternallyVisibleNotes.Content.TrimEnd.Replace("'", "''"))
        sbSQL.Append("', ")

        sbSQL.Append("ClientNotes")
        sbSQL.Append(" = '")
        sbSQL.Append(reClientVisibleNotes.Content.TrimEnd.Replace("'", "''"))
        sbSQL.Append("', ")

        sbSQL.Append("LastUpdateOn")
        sbSQL.Append(" = ")
        sbSQL.Append("GETDATE()")
        sbSQL.Append(", ")

        sbSQL.Append("LastUpdatedBy")
        sbSQL.Append(" = ")
        sbSQL.Append("0")

        sbSQL.Append(" ")
        
        sbSQL.Append("WHERE Year = ")
        sbSQL.Append(ddlYear.SelectedValue)


        sbSQL.Append(" AND Month = ")
        sbSQL.Append(ddlMonth.SelectedValue)
        
        sbSQL.Append(" AND Country = '")

        If rbUK.Checked Then
            sbSQL.Append("UK")
        Else
            sbSQL.Append("IRELAND")
        End If
        sbSQL.Append("'")

        sbSQL.Append(" ")
        sbSQL.Append("UPDATE ClientData_WU_MIMonthlyReport ")
        sbSQL.Append("SET ")
        sbSQL.Append("ServiceFeesManagementFee")
        sbSQL.Append(" = ")
        If rntbServiceFeesManagementFee.Text = String.Empty Then
            sbSQL.Append("-1")
        Else
            sbSQL.Append(rntbServiceFeesManagementFee.Text.Replace("'", "''"))
        End If
        sbSQL.Append(" ")
        
        sbSQL.Append("WHERE Year = ")
        sbSQL.Append(ddlYear.SelectedValue)

        sbSQL.Append(" AND Month = ")
        sbSQL.Append(ddlMonth.SelectedValue)
        
        sbSQL.Append(" AND Country = '")

        sbSQL.Append(UnselectedCountry)
        sbSQL.Append("'")

        Call ExecuteQueryToDataTable(sbSQL.ToString)
        lblMessage.Text = "Updated MI report for " & ddlMonth.SelectedItem.Text & " " & ddlYear.SelectedItem.Text & " - "
        If rbUK.Checked Then
            lblMessage.Text = lblMessage.Text & " UK - "
        Else
            lblMessage.Text = lblMessage.Text & " IRELAND - "
        End If
        If cbVisibleToClient.Checked Then
            lblMessage.Text &= "VISIBLE TO CLIENT"
            lblMessage.ForeColor = Drawing.Color.Green
        Else
            lblMessage.Text &= "NOT VISIBLE TO CLIENT"
            lblMessage.ForeColor = Drawing.Color.Red
        End If
        ddlYear.SelectedIndex = 0
        ddlMonth.SelectedIndex = 0
        rbUK.Checked = False
        rbIreland.Checked = False
        ddlYear.Enabled = True
        Call ClearForm()
        tabData.Visible = False
        ddlYear.Focus()
    End Sub

    Protected Sub btnCancel_Click(sender As Object, e As System.EventArgs)
        Call ClearForm()
        ddlYear.SelectedIndex = 0
        ddlMonth.SelectedIndex = 0
        ddlYear.Enabled = True
        ddlMonth.Enabled = False
        rbUK.Checked = False
        rbIreland.Checked = False
        rbUK.Enabled = False
        rbIreland.Enabled = False
        tabData.Visible = False
        ddlYear.Focus()
    End Sub
    
    Protected Sub ClearForm()
        rntbOrderBreakdownOperations.Text = String.Empty
        rntbOrderBreakdownMarketing.Text = String.Empty
        rntbOrderBreakdownFININT.Text = String.Empty
        rntbOrderBreakdownCosta.Text = String.Empty
        rntbOrderBreakdownPrePaid.Text = String.Empty

        rntbStorageCostsOperations.Text = String.Empty
        rntbStorageCostsMarketing.Text = String.Empty
        rntbStorageCostsFININT.Text = String.Empty
        rntbStorageCostsCosta.Text = String.Empty
        rntbStorageCostsPrePaid.Text = String.Empty

        rntbLogisticsCostsCourierOperations.Text = String.Empty
        rntbLogisticsCostsCourierMarketing.Text = String.Empty
        rntbLogisticsCostsCourierFININT.Text = String.Empty
        rntbLogisticsCostsCourierCosta.Text = String.Empty
        rntbLogisticsCostsPrepaid.Text = String.Empty
        rntbLogisticsCostsMailFulfilment.Text = String.Empty
        rntbLogisticsCostsAdHocFulfilment.Text = String.Empty

        rntbServiceFeesPickFeesOperations.Text = String.Empty
        rntbServiceFeesPickFeesMarketing.Text = String.Empty
        rntbServiceFeesPickFeesFININT.Text = String.Empty
        rntbServiceFeesPickFeesCosta.Text = String.Empty
        rntbServiceFeesPickFeesPrePaid.Text = String.Empty

        rntbServiceFeesGoodsInOperations.Text = String.Empty
        rntbServiceFeesGoodsInMarketing.Text = String.Empty
        rntbServiceFeesGoodsInFININT.Text = String.Empty
        rntbServiceFeesGoodsInCosta.Text = String.Empty
        rntbServiceFeesGoodsInPrePaid.Text = String.Empty

        rntbServiceFeesDestructionFeesOperations.Text = String.Empty
        rntbServiceFeesDestructionFeesMarketing.Text = String.Empty
        rntbServiceFeesDestructionFeesFININT.Text = String.Empty
        rntbServiceFeesDestructionFeesCosta.Text = String.Empty
        rntbServiceFeesDestructionFeesPrePaid.Text = String.Empty

        rntbServiceFeesManagementFee.Text = String.Empty
        
        reInternallyVisibleNotes.Content = String.Empty
        reClientVisibleNotes.Content = String.Empty

        cbVisibleToClient.Checked = False
        lnkbtnNotes.Text = "show internal & client-visible notes"
        trNotes.Visible = False
        lblMonth.Text = ""
    End Sub
  
    ' rntbOrderBreakdownOperations.Text = String.Empty
    'rntbOrderBreakdownMarketing.Text
    'rntbOrderBreakdownFININT.Text
    'rntbOrderBreakdownCosta.Text
    'rntbOrderBreakdownPrePaid.Text

    'rntbStorageCostsOperations.Text
    'rntbStorageCostsMarketing.Text
    'rntbStorageCostsFININT.Text
    'rntbStorageCostsCosta.Text
    'rntbStorageCostsPrePaid.Text

    'rntbLogisticsCostsCourierOperations.Text
    'rntbLogisticsCostsCourierMarketing.Text
    'rntbLogisticsCostsCourierFININT.Text
    'rntbLogisticsCostsCourierCosta.Text
    'rntbLogisticsCostsPrepaid.Text
    'rntbLogisticsCostsMailFulfilment.Text
    'rntbLogisticsCostsAdHocFulfilment.Text

    'rntbServiceFeesPickFeesOperations.Text
    'rntbServiceFeesPickFeesMarketing.Text
    'rntbServiceFeesPickFeesFININT.Text
    'rntbServiceFeesPickFeesCosta.Text
    'rntbServiceFeesPickFeesPrePaid.Text

    'rntbServiceFeesGoodsInOperations.Text
    'rntbServiceFeesGoodsInMarketing.Text
    'rntbServiceFeesGoodsInFININT.Text
    'rntbServiceFeesGoodsInCosta.Text
    'rntbServiceFeesGoodsInPrePaid.Text

    'rntbServiceFeesDestructionFeesOperations.Text
    'rntbServiceFeesDestructionFeesMarketing.Text
    'rntbServiceFeesDestructionFeesFININT.Text
    'rntbServiceFeesDestructionFeesCosta.Text
    'rntbServiceFeesDestructionFeesPrePaid.Text

    'rntbServiceFeesManagementFee
    
    Protected Sub ddlYear_SelectedIndexChanged(sender As Object, e As System.EventArgs)
        Dim ddl As DropDownList = sender
        If ddl.SelectedValue > 0 Then
            ddlMonth.Enabled = True
            ddlMonth.Focus()
        Else
            ddlMonth.Enabled = False
        End If
        ddlMonth.SelectedIndex = 0
        Call ClearForm()
    End Sub

    Protected Sub ddlMonth_SelectedIndexChanged(sender As Object, e As System.EventArgs)
        Dim ddl As DropDownList = sender
        If ddl.SelectedValue > 0 Then
            'tabData.Visible = True
            'Dim dr As DataRow = RetrieveEntry()
            'If dr IsNot Nothing Then
            '    Call PopulateForm(dr)
            'Else
            '    Call ClearForm()
            'End If
            rbUK.Enabled = True
            rbIreland.Enabled = True
        End If
        If ddl.SelectedIndex > 0 Then
            rbUK.Enabled = True
            rbIreland.Enabled = True
            'ddlYear.Enabled = False
            'ddlMonth.Enabled = False
            'rntbOrderBreakdownOperations.Focus()
        End If
        'lblMessage.Text = String.Empty
        'lblMonth.Text = " - " & ddlMonth.SelectedItem.Text & " " & ddlYear.SelectedItem.Text & " "
        'If rbUK.Checked Then
        '    lblMonth.Text = lblMonth.Text & "UK"
        'Else
        '    lblMonth.Text = lblMonth.Text & "IRELAND"
        'End If
    End Sub

    Protected Sub PopulateForm(dr As DataRow)
        Dim nVisibleToClient As Int32 = dr("VisibleToClient")
        If nVisibleToClient = 1 Then
            cbVisibleToClient.Checked = True
        Else
            cbVisibleToClient.Checked = False
        End If

        If dr("OrderBreakdownOperations") = -1 Then
            rntbOrderBreakdownOperations.Text = String.Empty
        Else
            rntbOrderBreakdownOperations.Text = dr("OrderBreakdownOperations")
        End If

        If dr("OrderBreakdownMarketing") = -1 Then
            rntbOrderBreakdownMarketing.Text = String.Empty
        Else
            rntbOrderBreakdownMarketing.Text = dr("OrderBreakdownMarketing")
        End If

        If dr("OrderBreakdownFININT") = -1 Then
            rntbOrderBreakdownFININT.Text = String.Empty
        Else
            rntbOrderBreakdownFININT.Text = dr("OrderBreakdownFININT")
        End If

        If dr("OrderBreakdownCosta") = -1 Then
            rntbOrderBreakdownCosta.Text = String.Empty
        Else
            rntbOrderBreakdownCosta.Text = dr("OrderBreakdownCosta")
        End If

        If dr("OrderBreakdownPrePaid") = -1 Then
            rntbOrderBreakdownPrePaid.Text = String.Empty
        Else
            rntbOrderBreakdownPrePaid.Text = dr("OrderBreakdownPrePaid")
        End If

        If dr("StorageCostsOperations") = -1 Then
            rntbStorageCostsOperations.Text = String.Empty
        Else
            rntbStorageCostsOperations.Text = dr("StorageCostsOperations")
        End If

        If dr("StorageCostsMarketing") = -1 Then
            rntbStorageCostsMarketing.Text = String.Empty
        Else
            rntbStorageCostsMarketing.Text = dr("StorageCostsMarketing")
        End If

        If dr("StorageCostsFININT") = -1 Then
            rntbStorageCostsFININT.Text = String.Empty
        Else
            rntbStorageCostsFININT.Text = dr("StorageCostsFININT")
        End If

        If dr("StorageCostsCosta") = -1 Then
            rntbStorageCostsCosta.Text = String.Empty
        Else
            rntbStorageCostsCosta.Text = dr("StorageCostsCosta")
        End If

        If dr("StorageCostsPrePaid") = -1 Then
            rntbStorageCostsPrePaid.Text = String.Empty
        Else
            rntbStorageCostsPrePaid.Text = dr("StorageCostsPrePaid")
        End If

        If dr("LogisticsCostsCourierOperations") = -1 Then
            rntbLogisticsCostsCourierOperations.Text = String.Empty
        Else
            rntbLogisticsCostsCourierOperations.Text = dr("LogisticsCostsCourierOperations")
        End If

        If dr("LogisticsCostsCourierMarketing") = -1 Then
            rntbLogisticsCostsCourierMarketing.Text = String.Empty
        Else
            rntbLogisticsCostsCourierMarketing.Text = dr("LogisticsCostsCourierMarketing")
        End If

        If dr("LogisticsCostsCourierFININT") = -1 Then
            rntbLogisticsCostsCourierFININT.Text = String.Empty
        Else
            rntbLogisticsCostsCourierFININT.Text = dr("LogisticsCostsCourierFININT")
        End If

        If dr("LogisticsCostsCourierCosta") = -1 Then
            rntbLogisticsCostsCourierCosta.Text = String.Empty
        Else
            rntbLogisticsCostsCourierCosta.Text = dr("LogisticsCostsCourierCosta")
        End If

        If dr("LogisticsCostsPrepaid") = -1 Then
            rntbLogisticsCostsPrepaid.Text = String.Empty
        Else
            rntbLogisticsCostsPrepaid.Text = dr("LogisticsCostsPrepaid")
        End If

        If dr("LogisticsCostsMailFulfilment") = -1 Then
            rntbLogisticsCostsMailFulfilment.Text = String.Empty
        Else
            rntbLogisticsCostsMailFulfilment.Text = dr("LogisticsCostsMailFulfilment")
        End If

        If dr("LogisticsCostsAdHocFulfilment") = -1 Then
            rntbLogisticsCostsAdHocFulfilment.Text = String.Empty
        Else
            rntbLogisticsCostsAdHocFulfilment.Text = dr("LogisticsCostsAdHocFulfilment")
        End If

        If dr("ServiceFeesPickFeesOperations") = -1 Then
            rntbServiceFeesPickFeesOperations.Text = String.Empty
        Else
            rntbServiceFeesPickFeesOperations.Text = dr("ServiceFeesPickFeesOperations")
        End If

        If dr("ServiceFeesPickFeesMarketing") = -1 Then
            rntbServiceFeesPickFeesMarketing.Text = String.Empty
        Else
            rntbServiceFeesPickFeesMarketing.Text = dr("ServiceFeesPickFeesMarketing")
        End If

        If dr("ServiceFeesPickFeesFININT") = -1 Then
            rntbServiceFeesPickFeesFININT.Text = String.Empty
        Else
            rntbServiceFeesPickFeesFININT.Text = dr("ServiceFeesPickFeesFININT")
        End If

        If dr("ServiceFeesPickFeesCosta") = -1 Then
            rntbServiceFeesPickFeesCosta.Text = String.Empty
        Else
            rntbServiceFeesPickFeesCosta.Text = dr("ServiceFeesPickFeesCosta")
        End If

        If dr("ServiceFeesPickFeesPrePaid") = -1 Then
            rntbServiceFeesPickFeesPrePaid.Text = String.Empty
        Else
            rntbServiceFeesPickFeesPrePaid.Text = dr("ServiceFeesPickFeesPrePaid")
        End If

        If dr("ServiceFeesGoodsInOperations") = -1 Then
            rntbServiceFeesGoodsInOperations.Text = String.Empty
        Else
            rntbServiceFeesGoodsInOperations.Text = dr("ServiceFeesGoodsInOperations")
        End If

        If dr("ServiceFeesGoodsInMarketing") = -1 Then
            rntbServiceFeesGoodsInMarketing.Text = String.Empty
        Else
            rntbServiceFeesGoodsInMarketing.Text = dr("ServiceFeesGoodsInMarketing")
        End If

        If dr("ServiceFeesGoodsInFININT") = -1 Then
            rntbServiceFeesGoodsInFININT.Text = String.Empty
        Else
            rntbServiceFeesGoodsInFININT.Text = dr("ServiceFeesGoodsInFININT")
        End If

        If dr("ServiceFeesGoodsInCosta") = -1 Then
            rntbServiceFeesGoodsInCosta.Text = String.Empty
        Else
            rntbServiceFeesGoodsInCosta.Text = dr("ServiceFeesGoodsInCosta")
        End If

        If dr("ServiceFeesGoodsInPrePaid") = -1 Then
            rntbServiceFeesGoodsInPrePaid.Text = String.Empty
        Else
            rntbServiceFeesGoodsInPrePaid.Text = dr("ServiceFeesGoodsInPrePaid")
        End If

        If dr("ServiceFeesDestructionFeesOperations") = -1 Then
            rntbServiceFeesDestructionFeesOperations.Text = String.Empty
        Else
            rntbServiceFeesDestructionFeesOperations.Text = dr("ServiceFeesDestructionFeesOperations")
        End If

        If dr("ServiceFeesDestructionFeesMarketing") = -1 Then
            rntbServiceFeesDestructionFeesMarketing.Text = String.Empty
        Else
            rntbServiceFeesDestructionFeesMarketing.Text = dr("ServiceFeesDestructionFeesMarketing")
        End If

        If dr("ServiceFeesDestructionFeesFININT") = -1 Then
            rntbServiceFeesDestructionFeesFININT.Text = String.Empty
        Else
            rntbServiceFeesDestructionFeesFININT.Text = dr("ServiceFeesDestructionFeesFININT")
        End If

        If dr("ServiceFeesDestructionFeesCosta") = -1 Then
            rntbServiceFeesDestructionFeesCosta.Text = String.Empty
        Else
            rntbServiceFeesDestructionFeesCosta.Text = dr("ServiceFeesDestructionFeesCosta")
        End If

        If dr("ServiceFeesDestructionFeesPrePaid") = -1 Then
            rntbServiceFeesDestructionFeesPrePaid.Text = String.Empty
        Else
            rntbServiceFeesDestructionFeesPrePaid.Text = dr("ServiceFeesDestructionFeesPrePaid")
        End If

        If dr("ServiceFeesManagementFee") = -1 Then
            rntbServiceFeesManagementFee.Text = String.Empty
        Else
            rntbServiceFeesManagementFee.Text = dr("ServiceFeesManagementFee")
        End If

        reInternallyVisibleNotes.Content = dr("InternalNotes")
        reClientVisibleNotes.Content = dr("ClientNotes")
        
        If reInternallyVisibleNotes.Content <> String.Empty Or reClientVisibleNotes.Content <> String.Empty Then
            lnkbtnNotes.Text = "hide internal & client-visible notes"
            trNotes.Visible = True
        End If
    End Sub
    
    Protected Function RetrieveEntry() As DataRow
        RetrieveEntry = Nothing
        Dim sSQL As String = "SELECT * FROM ClientData_WU_MIMonthlyReport WHERE Year = " & ddlYear.SelectedValue & " AND Month = " & ddlMonth.SelectedValue & " AND Country = '"
        If rbUK.Checked Then
            sSQL &= "UK'"
        Else
            sSQL &= "IRELAND'"
        End If
       
        Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
        If dt.Rows.Count = 1 Then
            RetrieveEntry = dt.Rows(0)
        End If
    End Function
    
    Protected Function RetrieveEntry(sCountry As String) As DataRow
        RetrieveEntry = Nothing
        Dim sSQL As String = "SELECT * FROM ClientData_WU_MIMonthlyReport WHERE Year = " & ddlYear.SelectedValue & " AND Month = " & ddlMonth.SelectedValue & " AND Country = '" & sCountry & "'"
        Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
        If dt.Rows.Count = 1 Then
            RetrieveEntry = dt.Rows(0)
        End If
    End Function

    Protected Sub lnkbtnNotes_Click(sender As Object, e As System.EventArgs)
        Dim lb As LinkButton = sender
        If lb.Text.Contains("show") Then
            lb.Text = "hide internal & client-visible notes"
            trNotes.Visible = True
        Else
            lb.Text = "show internal & client-visible notes"
            trNotes.Visible = False
        End If
    End Sub
    
    Protected Sub rbUK_CheckedChanged(sender As Object, e As System.EventArgs)
        Dim rb As RadioButton = sender
        If rb.Checked Then
            Call CountrySelected()
        End If
    End Sub

    Protected Sub rbIreland_CheckedChanged(sender As Object, e As System.EventArgs)
        Dim rb As RadioButton = sender
        If rb.Checked Then
            Call CountrySelected()
        End If
    End Sub
    
    Protected Sub CountrySelected()
        tabData.Visible = True
        Dim dr As DataRow = RetrieveEntry()
        If dr IsNot Nothing Then
            Call PopulateForm(dr)
        Else
            Call ClearForm()
        End If
        ddlYear.Enabled = False
        ddlMonth.Enabled = False
        rbUK.Enabled = False
        rbIreland.Enabled = False
        
        lblMessage.Text = String.Empty
        lblMonth.Text = " - " & ddlMonth.SelectedItem.Text & " " & ddlYear.SelectedItem.Text & " "
        If rbUK.Checked Then
            lblMonth.Text = lblMonth.Text & "UK"
        Else
            lblMonth.Text = lblMonth.Text & "IRELAND"
        End If
        rntbOrderBreakdownOperations.Focus()
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
    <div style="font-size: small; font-family: Verdana">
        <strong>&nbsp;Western Union Management Information Monthly Data
        <asp:Label ID="lblMonth" runat="server" Font-Names="Verdana" Font-Size="Small"
            Font-Bold="True" />
        </strong><br />
        &nbsp;<br />
        &nbsp;Year:
        <asp:DropDownList ID="ddlYear" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlYear_SelectedIndexChanged">
            <asp:ListItem Value="0">- please select -</asp:ListItem>
            <asp:ListItem>2013</asp:ListItem>
            <asp:ListItem>2014</asp:ListItem>
            <asp:ListItem>2015</asp:ListItem>
            <asp:ListItem>2016</asp:ListItem>
            <asp:ListItem>2017</asp:ListItem>
            <asp:ListItem>2018</asp:ListItem>
            <asp:ListItem>2019</asp:ListItem>
            <asp:ListItem>2020</asp:ListItem>
        </asp:DropDownList>
        &nbsp;Month:
        <asp:DropDownList ID="ddlMonth" runat="server" AutoPostBack="True" Enabled="False"
            OnSelectedIndexChanged="ddlMonth_SelectedIndexChanged">
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
        &nbsp;&nbsp;
        <asp:RadioButton ID="rbUK" runat="server" GroupName="Country" 
            Text="UK" AutoPostBack="True" oncheckedchanged="rbUK_CheckedChanged" 
            Enabled="False" />
        <asp:RadioButton ID="rbIreland" runat="server" GroupName="Country" 
            Text="IRELAND" AutoPostBack="True" 
            oncheckedchanged="rbIreland_CheckedChanged" Enabled="False" />
&nbsp;&nbsp;
        <asp:Label ID="lblMessage" runat="server" Font-Names="Verdana" Font-Size="Small"
            Font-Bold="True" />
        <br />
        <br />
        <table id="tabData" runat="server" visible="false" style="width: 100%">
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
                <td align="right">
                    <asp:Label ID="lblLegendAgentID" runat="server" Font-Names="Verdana" Font-Size="Small"
                        Text="ORDERS" Font-Bold="True" />
                </td>
                <td>
                    &nbsp;
                </td>
                <td align="right">
                    <asp:Label ID="Label41" runat="server" Font-Names="Verdana" Font-Size="Small" Text="STORAGE COSTS"
                        Font-Bold="True" />
                </td>
                <td>
                    &nbsp;
                </td>
                <td align="right">
                    <asp:Label ID="Label42" runat="server" Font-Names="Verdana" Font-Size="Small" Text="LOGISTICS COSTS"
                        Font-Bold="True" />
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td align="right">
                    <asp:Label ID="Label43" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Operations:" />
                </td>
                <td>
                    <telerik:RadNumericTextBox ID="rntbOrderBreakdownOperations" runat="server" Font-Bold="True"
                        Type="Number" Font-Size="Small" MaxValue="100000" MinValue="0" ShowSpinButtons="True"
                        Width="150px">
                        <NumberFormat DecimalDigits="0" />
                    </telerik:RadNumericTextBox>
                </td>
                <td align="right">
                    <asp:Label ID="Label47" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Operations:" />
                </td>
                <td>
                    <telerik:RadNumericTextBox ID="rntbStorageCostsOperations" runat="server" Font-Bold="True"
                        Font-Size="Small" MaxValue="100000" MinValue="0" ShowSpinButtons="True" Width="150px"
                        Type="Currency">
                        <NumberFormat ZeroPattern="£n" />
                    </telerik:RadNumericTextBox>
                </td>
                <td align="right">
                    <asp:Label ID="Label54" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Courier Operations:" />
                </td>
                <td>
                    <telerik:RadNumericTextBox ID="rntbLogisticsCostsCourierOperations" runat="server"
                        Font-Bold="True" Font-Size="Small" MaxValue="100000" MinValue="0" ShowSpinButtons="True"
                        Width="150px" Type="Currency">
                        <NumberFormat ZeroPattern="£n" />
                    </telerik:RadNumericTextBox>
                </td>
            </tr>
            <tr>
                <td align="right">
                    <asp:Label ID="Label44" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Marketing:" />
                </td>
                <td>
                    <telerik:RadNumericTextBox ID="rntbOrderBreakdownMarketing" runat="server" Font-Bold="True"
                        Type="Number" Font-Size="Small" MaxValue="100000" MinValue="0" ShowSpinButtons="True"
                        Width="150px">
                        <NumberFormat DecimalDigits="0" />
                    </telerik:RadNumericTextBox>
                </td>
                <td align="right">
                    <asp:Label ID="Label48" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Marketing:" />
                </td>
                <td>
                    <telerik:RadNumericTextBox ID="rntbStorageCostsMarketing" runat="server" Font-Bold="True"
                        Font-Size="Small" MaxValue="100000" MinValue="0" ShowSpinButtons="True" Width="150px"
                        Type="Currency">
                        <NumberFormat ZeroPattern="£n" />
                    </telerik:RadNumericTextBox>
                </td>
                <td align="right">
                    <asp:Label ID="Label55" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Courier Marketing:" />
                </td>
                <td>
                    <telerik:RadNumericTextBox ID="rntbLogisticsCostsCourierMarketing" runat="server"
                        Font-Bold="True" Font-Size="Small" MaxValue="100000" MinValue="0" ShowSpinButtons="True"
                        Width="150px" Type="Currency">
                        <NumberFormat ZeroPattern="£n" />
                    </telerik:RadNumericTextBox>
                </td>
            </tr>
            <tr>
                <td align="right">
                    <asp:Label ID="Label45" runat="server" Font-Names="Verdana" Font-Size="Small" Text="FININT:" />
                </td>
                <td>
                    <telerik:RadNumericTextBox ID="rntbOrderBreakdownFININT" runat="server" Font-Bold="True"
                        Type="Number" Font-Size="Small" MaxValue="100000" MinValue="0" ShowSpinButtons="True"
                        Width="150px">
                        <NumberFormat DecimalDigits="0" />
                    </telerik:RadNumericTextBox>
                </td>
                <td align="right">
                    <asp:Label ID="Label49" runat="server" Font-Names="Verdana" Font-Size="Small" Text="FININT:" />
                </td>
                <td>
                    <telerik:RadNumericTextBox ID="rntbStorageCostsFININT" runat="server" Font-Bold="True"
                        Font-Size="Small" MaxValue="100000" MinValue="0" ShowSpinButtons="True" Width="150px"
                        Type="Currency">
                        <NumberFormat ZeroPattern="£n" />
                    </telerik:RadNumericTextBox>
                </td>
                <td align="right">
                    <asp:Label ID="Label56" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Courier FININT:" />
                </td>
                <td>
                    <telerik:RadNumericTextBox ID="rntbLogisticsCostsCourierFININT" runat="server" Font-Bold="True"
                        Font-Size="Small" MaxValue="100000" MinValue="0" ShowSpinButtons="True" Width="150px"
                        Type="Currency">
                        <NumberFormat ZeroPattern="£n" />
                    </telerik:RadNumericTextBox>
                </td>
            </tr>
            <tr>
                <td align="right">
                    <asp:Label ID="Label50" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Costa:" />
                </td>
                <td>
                    <telerik:RadNumericTextBox ID="rntbOrderBreakdownCosta" runat="server" Font-Bold="True"
                        Type="Number" Font-Size="Small" MaxValue="100000" MinValue="0" ShowSpinButtons="True"
                        Width="150px">
                        <NumberFormat DecimalDigits="0" />
                    </telerik:RadNumericTextBox>
                </td>
                <td align="right">
                    <asp:Label ID="Label51" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Costa:" />
                </td>
                <td>
                    <telerik:RadNumericTextBox ID="rntbStorageCostsCosta" runat="server" Font-Bold="True"
                        Font-Size="Small" MaxValue="100000" MinValue="0" ShowSpinButtons="True" Width="150px"
                        Type="Currency">
                        <NumberFormat ZeroPattern="£n" />
                    </telerik:RadNumericTextBox>
                </td>
                <td align="right">
                    <asp:Label ID="Label57" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Courier Costa:" />
                </td>
                <td>
                    <telerik:RadNumericTextBox ID="rntbLogisticsCostsCourierCosta" runat="server" Font-Bold="True"
                        Font-Size="Small" MaxValue="100000" MinValue="0" ShowSpinButtons="True" Width="150px"
                        Type="Currency">
                        <NumberFormat ZeroPattern="£n" />
                    </telerik:RadNumericTextBox>
                </td>
            </tr>
            <tr>
                <td align="right">
                    <asp:Label ID="Label52" runat="server" Font-Names="Verdana" Font-Size="Small" Text="PrePaid:" />
                </td>
                <td>
                    <telerik:RadNumericTextBox ID="rntbOrderBreakdownPrePaid" runat="server" Font-Bold="True"
                        Type="Number" Font-Size="Small" MaxValue="100000" MinValue="0" ShowSpinButtons="True"
                        Width="150px">
                        <NumberFormat DecimalDigits="0" />
                    </telerik:RadNumericTextBox>
                </td>
                <td align="right">
                    <asp:Label ID="Label53" runat="server" Font-Names="Verdana" Font-Size="Small" Text="PrePaid:" />
                </td>
                <td>
                    <telerik:RadNumericTextBox ID="rntbStorageCostsPrePaid" runat="server" Font-Bold="True"
                        Font-Size="Small" MaxValue="100000" MinValue="0" ShowSpinButtons="True" Width="150px"
                        Type="Currency">
                        <NumberFormat ZeroPattern="£n" />
                    </telerik:RadNumericTextBox>
                </td>
                <td align="right">
                    <asp:Label ID="Label58" runat="server" Font-Names="Verdana" Font-Size="Small" Text="PrePaid:" />
                </td>
                <td>
                    <telerik:RadNumericTextBox ID="rntbLogisticsCostsPrepaid" runat="server" Font-Bold="True"
                        Font-Size="Small" MaxValue="100000" MinValue="0" ShowSpinButtons="True" Width="150px"
                        Type="Currency">
                        <NumberFormat ZeroPattern="£n" />
                    </telerik:RadNumericTextBox>
                </td>
            </tr>
            <tr>
                <td align="right">
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
                <td align="right">
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
                <td align="right">
                    <asp:Label ID="Label59" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Mail Fulfilment:" />
                </td>
                <td>
                    <telerik:RadNumericTextBox ID="rntbLogisticsCostsMailFulfilment" runat="server" Font-Bold="True"
                        Font-Size="Small" MaxValue="100000" MinValue="0" ShowSpinButtons="True" Width="150px"
                        Type="Currency">
                        <NumberFormat ZeroPattern="£n" />
                    </telerik:RadNumericTextBox>
                </td>
            </tr>
            <tr>
                <td align="right">
                </td>
                <td>
                </td>
                <td align="right">
                </td>
                <td>
                </td>
                <td align="right">
                    <asp:Label ID="Label60" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Ad Hoc Fulfilment:" />
                </td>
                <td>
                    <telerik:RadNumericTextBox ID="rntbLogisticsCostsAdHocFulfilment" runat="server"
                        Font-Bold="True" Font-Size="Small" MaxValue="100000" MinValue="0" ShowSpinButtons="True"
                        Width="150px" Type="Currency">
                        <NumberFormat ZeroPattern="£n" />
                    </telerik:RadNumericTextBox>
                </td>
            </tr>
            <tr>
                <td align="right">
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
                <td align="right">
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
                <td align="right">
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td align="right">
                    <asp:Label ID="Label61" runat="server" Font-Names="Verdana" Font-Size="Small" Text="SERVICE FEES - PICK FEES"
                        Font-Bold="True" />
                </td>
                <td>
                    &nbsp;
                </td>
                <td align="right">
                    <asp:Label ID="Label63" runat="server" Font-Names="Verdana" Font-Size="Small" Text="SERVICE FEES - GOODS IN"
                        Font-Bold="True" />
                </td>
                <td>
                    &nbsp;
                </td>
                <td align="right">
                    <asp:Label ID="Label64" runat="server" Font-Names="Verdana" Font-Size="Small" Text="SERVICE FEES - DESTRUCTION FEES"
                        Font-Bold="True" />
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td align="right">
                    <asp:Label ID="Label65" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Operations:" />
                </td>
                <td style="margin-left: 40px">
                    <telerik:RadNumericTextBox ID="rntbServiceFeesPickFeesOperations" runat="server"
                        Font-Bold="True" Font-Size="Small" MaxValue="100000" MinValue="0" ShowSpinButtons="True"
                        Width="150px" Type="Currency">
                        <NumberFormat ZeroPattern="£n" />
                    </telerik:RadNumericTextBox>
                </td>
                <td align="right">
                    <asp:Label ID="Label70" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Operations:" />
                </td>
                <td>
                    <telerik:RadNumericTextBox ID="rntbServiceFeesGoodsInOperations" runat="server" Font-Bold="True"
                        Font-Size="Small" MaxValue="100000" MinValue="0" ShowSpinButtons="True" Width="150px"
                        Type="Currency">
                        <NumberFormat ZeroPattern="£n" />
                    </telerik:RadNumericTextBox>
                </td>
                <td align="right">
                    <asp:Label ID="Label71" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Operations:" />
                </td>
                <td>
                    <telerik:RadNumericTextBox ID="rntbServiceFeesDestructionFeesOperations" runat="server"
                        Font-Bold="True" Font-Size="Small" MaxValue="100000" MinValue="0" ShowSpinButtons="True"
                        Width="150px" Type="Currency">
                        <NumberFormat ZeroPattern="£n" />
                    </telerik:RadNumericTextBox>
                </td>
            </tr>
            <tr>
                <td align="right">
                    <asp:Label ID="Label66" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Marketing:" />
                </td>
                <td>
                    <telerik:RadNumericTextBox ID="rntbServiceFeesPickFeesMarketing" runat="server" Font-Bold="True"
                        Font-Size="Small" MaxValue="100000" MinValue="0" ShowSpinButtons="True" Width="150px"
                        Type="Currency">
                        <NumberFormat ZeroPattern="£n" />
                    </telerik:RadNumericTextBox>
                </td>
                <td align="right">
                    <asp:Label ID="Label72" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Marketing:" />
                </td>
                <td>
                    <telerik:RadNumericTextBox ID="rntbServiceFeesGoodsInMarketing" runat="server" Font-Bold="True"
                        Font-Size="Small" MaxValue="100000" MinValue="0" ShowSpinButtons="True" Width="150px"
                        Type="Currency">
                        <NumberFormat ZeroPattern="£n" />
                    </telerik:RadNumericTextBox>
                </td>
                <td align="right">
                    <asp:Label ID="Label73" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Marketing:" />
                </td>
                <td>
                    <telerik:RadNumericTextBox ID="rntbServiceFeesDestructionFeesMarketing" runat="server"
                        Font-Bold="True" Font-Size="Small" MaxValue="100000" MinValue="0" ShowSpinButtons="True"
                        Width="150px" Type="Currency">
                        <NumberFormat ZeroPattern="£n" />
                    </telerik:RadNumericTextBox>
                </td>
            </tr>
            <tr>
                <td align="right">
                    <asp:Label ID="Label67" runat="server" Font-Names="Verdana" Font-Size="Small" Text="FININT:" />
                </td>
                <td>
                    <telerik:RadNumericTextBox ID="rntbServiceFeesPickFeesFININT" runat="server" Font-Bold="True"
                        Font-Size="Small" MaxValue="100000" MinValue="0" ShowSpinButtons="True" Width="150px"
                        Type="Currency">
                        <NumberFormat ZeroPattern="£n" />
                    </telerik:RadNumericTextBox>
                </td>
                <td align="right">
                    <asp:Label ID="Label74" runat="server" Font-Names="Verdana" Font-Size="Small" Text="FININT:" />
                </td>
                <td>
                    <telerik:RadNumericTextBox ID="rntbServiceFeesGoodsInFININT" runat="server" Font-Bold="True"
                        Font-Size="Small" MaxValue="100000" MinValue="0" ShowSpinButtons="True" Width="150px"
                        Type="Currency">
                        <NumberFormat ZeroPattern="£n" />
                    </telerik:RadNumericTextBox>
                </td>
                <td align="right">
                    <asp:Label ID="Label75" runat="server" Font-Names="Verdana" Font-Size="Small" Text="FININT:" />
                </td>
                <td>
                    <telerik:RadNumericTextBox ID="rntbServiceFeesDestructionFeesFININT" runat="server"
                        Font-Bold="True" Font-Size="Small" MaxValue="100000" MinValue="0" ShowSpinButtons="True"
                        Width="150px" Type="Currency">
                        <NumberFormat ZeroPattern="£n" />
                    </telerik:RadNumericTextBox>
                </td>
            </tr>
            <tr>
                <td align="right">
                    <asp:Label ID="Label68" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Costa:" />
                </td>
                <td>
                    <telerik:RadNumericTextBox ID="rntbServiceFeesPickFeesCosta" runat="server" Font-Bold="True"
                        Font-Size="Small" MaxValue="100000" MinValue="0" ShowSpinButtons="True" Width="150px"
                        Type="Currency">
                        <NumberFormat ZeroPattern="£n" />
                    </telerik:RadNumericTextBox>
                </td>
                <td align="right">
                    <asp:Label ID="Label76" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Costa:" />
                </td>
                <td>
                    <telerik:RadNumericTextBox ID="rntbServiceFeesGoodsInCosta" runat="server" Font-Bold="True"
                        Font-Size="Small" MaxValue="100000" MinValue="0" ShowSpinButtons="True" Width="150px"
                        Type="Currency">
                        <NumberFormat ZeroPattern="£n" />
                    </telerik:RadNumericTextBox>
                </td>
                <td align="right">
                    <asp:Label ID="Label77" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Costa:" />
                </td>
                <td>
                    <telerik:RadNumericTextBox ID="rntbServiceFeesDestructionFeesCosta" runat="server"
                        Font-Bold="True" Font-Size="Small" MaxValue="100000" MinValue="0" ShowSpinButtons="True"
                        Width="150px" Type="Currency">
                        <NumberFormat ZeroPattern="£n" />
                    </telerik:RadNumericTextBox>
                </td>
            </tr>
            <tr>
                <td align="right">
                    <asp:Label ID="Label69" runat="server" Font-Names="Verdana" Font-Size="Small" Text="PrePaid:" />
                </td>
                <td>
                    <telerik:RadNumericTextBox ID="rntbServiceFeesPickFeesPrePaid" runat="server" Font-Bold="True"
                        Font-Size="Small" MaxValue="100000" MinValue="0" ShowSpinButtons="True" Width="150px"
                        Type="Currency">
                        <NumberFormat ZeroPattern="£n" />
                    </telerik:RadNumericTextBox>
                </td>
                <td align="right">
                    <asp:Label ID="Label78" runat="server" Font-Names="Verdana" Font-Size="Small" Text="PrePaid:" />
                </td>
                <td>
                    <telerik:RadNumericTextBox ID="rntbServiceFeesGoodsInPrePaid" runat="server" Font-Bold="True"
                        Font-Size="Small" MaxValue="100000" MinValue="0" ShowSpinButtons="True" Width="150px"
                        Type="Currency">
                        <NumberFormat ZeroPattern="£n" />
                    </telerik:RadNumericTextBox>
                </td>
                <td align="right">
                    <asp:Label ID="Label79" runat="server" Font-Names="Verdana" Font-Size="Small" Text="PrePaid:" />
                </td>
                <td>
                    <telerik:RadNumericTextBox ID="rntbServiceFeesDestructionFeesPrePaid" runat="server"
                        Font-Bold="True" Font-Size="Small" MaxValue="100000" MinValue="0" ShowSpinButtons="True"
                        Width="150px" Type="Currency">
                        <NumberFormat ZeroPattern="£n" />
                    </telerik:RadNumericTextBox>
                </td>
            </tr>
            <tr>
                <td align="right">
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
                <td align="right">
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
                <td align="right">
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td align="right">
                    <asp:Label ID="Label82" runat="server" Font-Names="Verdana" Font-Size="Small" Text="SERVICE FEES - "
                        Font-Bold="True" />
                    <asp:Label ID="Label83" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Management Fee:" />
                </td>
                <td>
                    <telerik:RadNumericTextBox ID="rntbServiceFeesManagementFee" runat="server" Font-Bold="True"
                        Font-Size="Small" MaxValue="100000" MinValue="0" ShowSpinButtons="True" Width="150px"
                        Type="Currency">
                        <NumberFormat ZeroPattern="£n" />
                    </telerik:RadNumericTextBox>
                </td>
                <td align="right">
                    &nbsp;</td>
                <td>
                    &nbsp;</td>
                <td align="right">
                    &nbsp;</td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td align="right">
                    &nbsp;</td>
                <td>
                    &nbsp;</td>
                <td align="right">
                    &nbsp;</td>
                <td>
                    &nbsp;</td>
                <td align="right">
                    &nbsp;</td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td align="right">
                    <asp:Label ID="Label10" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Visible to client:" />
                </td>
                <td>
                    <asp:CheckBox ID="cbVisibleToClient" runat="server" />
                </td>
                <td align="right">
                    &nbsp;
                </td>
                <td>
                    <asp:LinkButton ID="lnkbtnNotes" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        OnClick="lnkbtnNotes_Click">show internal &amp; client-visible notes</asp:LinkButton>
                </td>
                <td align="right">
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr id="trNotes" runat="server" visible="false">
                <td align="right" colspan="6">
                    <table style="width: 100%">
                        <tr>
                            <td style="width: 2%">
                                &nbsp;
                            </td>
                            <td style="width: 47%" align="left">
                                <asp:Label ID="Label80" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Internal Notes:" />
                            </td>
                            <td style="width: 2%">
                                &nbsp;
                            </td>
                            <td style="width: 47%" align="left">
                                <asp:Label ID="Label81" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Client-Visible Notes:" />
                            </td>
                            <td style="width: 2%">
                                &nbsp;
                            </td>
                        </tr>
                        <tr>
                            <td>
                                &nbsp;
                            </td>
                            <td>
                                <telerik:RadEditor ID="reInternallyVisibleNotes" runat="server" Width="100%">
                                </telerik:RadEditor>
                            </td>
                            <td>
                                &nbsp;
                            </td>
                            <td>
                                <telerik:RadEditor ID="reClientVisibleNotes" runat="server" Width="100%">
                                </telerik:RadEditor>
                            </td>
                            <td>
                                &nbsp;
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td align="right">
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
                <td align="right">
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
                <td align="right">
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
                    <asp:Button ID="btnUpdate" runat="server" Text="Update" Width="150px" OnClick="btnUpdate_Click" />
                    &nbsp;<asp:Button ID="btnCancel" runat="server" Text="Cancel" OnClick="btnCancel_Click" />
                </td>
                <td>
                    &nbsp;
                </td>
                <td align="right">
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
                <td>
                    &nbsp;
                </td>
                <td align="right">
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
                <td align="right">
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
        </table>
        <br />
    <p>
        &nbsp;
        NOTES:</p>
    <p>
        &nbsp;
        1.&nbsp; This form differentiates between blank fields and fields containing 
        zero. You can leave fields empty when no data is available.</p>
    </div>
    </form>
</body>
</html>
