<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ import Namespace="System.IO" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient " %>
<%@ import Namespace="System.Collections.Generic" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" " http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

    Const WUIRE_CUSTOMER_KEY As Int32 = 686
    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsNumeric(Session("CustomerKey")) Then
            Server.Transfer("session_expired.aspx")
            Exit Sub
        End If
        If IsPostBack Then
            'Call HideAllPanels()
        End If
        Call SetTitle()
    End Sub

    Protected Sub SetTitle()
        Dim sTitle As String = Session("SiteTitle")
        If sTitle <> String.Empty Then
            sTitle += " - "
        End If
        Page.Header.Title = sTitle & "Western Union IRELAND Serial Numbers"
    End Sub
    
    Protected Sub HideAllPanels()
        pnlInstructions.Visible = False
        pnlEnterNumbers.Visible = False
        pnlConsignments.Visible = False
    End Sub
    
    Protected Sub lnkbtnShowConsignmentsWithoutSerialNumbers_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        gvConsignments.PageIndex = 0
        gvConsignments.DataSource = GetConsignments(bAll:=False)
        gvConsignments.DataBind()
        pbShowAll = False
        tbConsignmentNumber.Text = String.Empty
        tbSerialNo.Text = String.Empty
    End Sub

    Protected Sub lnkbtnShowAllConsignments_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        gvConsignments.PageIndex = 0
        gvConsignments.DataSource = GetConsignments(bAll:=True)
        gvConsignments.DataBind()
        pbShowAll = True
        tbConsignmentNumber.Text = String.Empty
        tbSerialNo.Text = String.Empty
    End Sub
    
    Protected Function GetConsignments(ByVal bAll As Boolean) As DataTable
        Dim sbSQL As New StringBuilder
        sbSQL.Append("SELECT DISTINCT fsn.[id] 'Record', c.[key] 'Consignment', lp.ProductCode + ' ' + lp.ProductDescription 'Product', fsn.BookNumber 'Book No', fsn.FirstPageNumber '1st Page No', CneeName + ', ' + CneeAddr1 + ', ' + CneeTown + ' ' + CneePostCode 'Consignee', c.CreatedOn 'Date', c.StateId 'Status' ")
        sbSQL.Append("FROM Consignment c INNER JOIN LogisticMovement lm ON c.[key] = lm.ConsignmentKey ")
        sbSQL.Append("LEFT OUTER JOIN ClientData_WUIRE_SerialNumbers fsn ON c.[key] = fsn.ConsignmentKey ")
        sbSQL.Append("LEFT OUTER JOIN LogisticProduct lp ON fsn.LogisticProductKey = lp.LogisticProductKey ")
        sbSQL.Append("WHERE (c.CustomerKey = " & WUIRE_CUSTOMER_KEY & ") ")
        If cbLast30Days.Checked Then
            sbSQL.Append("AND (c.CreatedOn >= (GETDATE()-30)) ")
        End If
        sbSQL.Append("AND lm.LogisticProductKey IN ")
        sbSQL.Append("(SELECT LogisticProductKey FROM LogisticProduct WHERE SerialNumbersFlag = 'Y' AND CustomerKey = " & WUIRE_CUSTOMER_KEY & ") ")
        If Not bAll Then
            sbSQL.Append("AND (fsn.BookNumber IS NULL OR fsn.FirstPageNumber IS NULL) ")
        End If
        sbSQL.Append("ORDER BY c.[key]")
        GetConsignments = ExecuteQueryToDataTable(sbSQL.ToString)
        pnlConsignments.Visible = True
    End Function
    
    Protected Sub btnAdd_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim b As Button = sender
        Call HideAllPanels()
        tbBookNumber.Text = String.Empty
        tbFirstPageNumber.Text = String.Empty
        lblConsignment.Text = b.CommandArgument
        lblConsignee.Text = b.CommandName
        Dim oListItemCollection As ListItemCollection = ExecuteQueryToListItemCollection("SELECT DISTINCT lp.ProductCode + ' ' + lp.ProductDescription 'P', lp.LogisticProductKey 'K' FROM LogisticMovement lm INNER JOIN LogisticProduct lp ON lp.LogisticProductKey = lm.LogisticProductKey WHERE SerialNumbersFlag = 'Y' AND ConsignmentKey = " & b.CommandArgument, "P", "K")
        ddlProduct.Items.Clear()
        For Each li As ListItem In oListItemCollection
            ddlProduct.Items.Add(New ListItem(li.Text, li.Value))
        Next
        If ddlProduct.Items.Count = 1 Then
            ddlProduct.Enabled = False
        Else
            ddlProduct.Enabled = True
        End If
        pnlEnterNumbers.Visible = True
        tbBookNumber.Focus()
    End Sub
    
    Protected Sub SaveNumbersInsert()
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand
        Dim sSQL As String = "INSERT INTO ClientData_WUIRE_SerialNumbers (ConsignmentKey, LogisticProductKey, BookNumber, FirstPageNumber, LastUpdatedOn, LastUpdatedBy) VALUES ("
        sSQL += lblConsignment.Text & ", " & ddlProduct.SelectedValue & "," & tbBookNumber.Text & ", " & tbFirstPageNumber.Text & ", GETDATE(), " & Session("UserKey") & ")"
        Try
            oConn.Open()
            oCmd = New SqlCommand(sSQL, oConn)
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            WebMsgBox.Show("Error in SaveNumbersInsert: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub SaveNumbersUpdate()
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand
        Dim sSQL As String = "UPDATE ClientData_WUIRE_SerialNumbers "
        sSQL += "SET BookNumber = " & tbBookNumber.Text & ", FirstPageNumber = " & tbFirstPageNumber.Text & ", "
        sSQL += "LastUpdatedOn = GETDATE(), LastUpdatedBy = " & Session("UserKey") & " "
        sSQL += "WHERE ConsignmentKey = " & lblConsignment.Text
        Try
            oConn.Open()
            oCmd = New SqlCommand(sSQL, oConn)
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            WebMsgBox.Show("Error in SaveNumbersUpdate: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Function NumberExists(ByVal sType As String, ByVal nNumber As Integer) As Boolean
        NumberExists = False
        Dim oDataReader As SqlDataReader = Nothing
        Dim oConn As New SqlConnection(gsConn)
        Dim sSQL As String = "SELECT * FROM ClientData_WUIRE_SerialNumbers WHERE " & sType & " = " & nNumber
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            NumberExists = oDataReader.HasRows
        Catch ex As Exception
            WebMsgBox.Show("Error in NumberExists: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Function
    
    Protected Function RetrieveSavedNumbers(ByVal sConsignmentNo As String) As Boolean
        RetrieveSavedNumbers = False
        tbBookNumber.Text = String.Empty
        tbFirstPageNumber.Text = String.Empty
        Dim oDataReader As SqlDataReader = Nothing
        Dim oConn As New SqlConnection(gsConn)
        Dim sSQL As String = "SELECT * FROM ClientData_WUIRE_SerialNumbers WHERE ConsignmentKey = " & sConsignmentNo
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            If oDataReader.HasRows Then
                RetrieveSavedNumbers = True
                oDataReader.Read()
                If Not IsDBNull(oDataReader("BookNumber")) Then
                    tbBookNumber.Text = oDataReader("BookNumber")
                End If
                If Not IsDBNull(oDataReader("FirstPageNumber")) Then
                    tbFirstPageNumber.Text = oDataReader("FirstPageNumber")
                End If
            End If
        Catch ex As Exception
            WebMsgBox.Show("Error in RetrieveSavedNumbers: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Function
    
    Protected Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        cbOverrideValidation.Visible = False
        cbOverrideValidation.Checked = False
        Call HideAllPanels()
        pnlConsignments.Visible = True
    End Sub
    
    Protected Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call SaveNumbers()
    End Sub
    
    Protected Function NumberMatches() As Boolean
        Return True
        
        NumberMatches = False
        Dim sBookUniqueNumber As String = tbFirstPageNumber.Text.Trim.Substring(0, 1)
        ' HSBC
        If ddlProduct.SelectedValue = 34234 And sBookUniqueNumber = "2" Then
            NumberMatches = True
            Exit Function
        End If
        ' NATWEST
        If ddlProduct.SelectedValue = 34235 And sBookUniqueNumber = "5" Then
            NumberMatches = True
            Exit Function
        End If
        ' ULSTER
        If ddlProduct.SelectedValue = 34236 And (sBookUniqueNumber = "9" Or sBookUniqueNumber = "8") Then
            NumberMatches = True
            Exit Function
        End If
        ' RBS
        If ddlProduct.SelectedValue = 34237 And (sBookUniqueNumber = "8" Or sBookUniqueNumber = "9") Then
            NumberMatches = True
            Exit Function
        End If
    End Function
    
    Protected Function SaveNumbers() As Boolean
        SaveNumbers = False
        If IsNumeric(tbBookNumber.Text) AndAlso IsNumeric(tbFirstPageNumber.Text) Then
            'If NumberExists("BookNumber", tbBookNumber.Text) Then
            '    WebMsgBox.Show("Book number " & tbBookNumber.Text & " has already been used - NOT SAVED!")
            '    pnlEnterNumbers.Visible = True
            '    Exit Function
            'End If
            If NumberExists("FirstPageNumber", tbFirstPageNumber.Text) Then
                WebMsgBox.Show("First page number " & tbFirstPageNumber.Text & " has already been used - NOT SAVED!")
                pnlEnterNumbers.Visible = True
                Exit Function
            End If
            If Not cbOverrideValidation.Checked Then
                If Not NumberMatches() Then
                    cbOverrideValidation.Visible = True
                    WebMsgBox.Show("The first page number entered does not appear to match the type of book.\n\n Click the override validation check box if you really want to use this number.")
                    pnlEnterNumbers.Visible = True
                    Exit Function
                End If
            Else
                cbOverrideValidation.Visible = False
                cbOverrideValidation.Checked = False
            End If
            If pbRecordExists Then
                Call SaveNumbersUpdate()
            Else
                Call SaveNumbersInsert()
            End If
            Call HideAllPanels()
        
            gvConsignments.DataSource = GetConsignments(bAll:=pbShowAll)
            gvConsignments.DataBind()
        Else
            WebMsgBox.Show("Both fields must be numeric - NOT SAVED!")
        End If
        pnlConsignments.Visible = True
        SaveNumbers = True
    End Function
    
    Protected Sub gvConsignments_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs)
        If tbConsignmentNumber.Text = String.Empty Then
            gvConsignments.DataSource = GetConsignments(bAll:=pbShowAll)
        Else
            gvConsignments.DataSource = GetSingleConsignment(tbConsignmentNumber.Text)
        End If
        gvConsignments.Visible = True
        pnlConsignments.Visible = True
        gvConsignments.PageIndex = e.NewPageIndex
        gvConsignments.DataBind()
    End Sub
    
    Protected Sub btnRemove_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim b As Button = sender
        If b.CommandArgument <> String.Empty Then
            Call RemoveRecord(b.CommandArgument)
        End If
    End Sub
    
    Protected Sub RemoveRecord(ByVal nRecord As Integer)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand
        Dim sSQL As String = "DELETE FROM ClientData_WUIRE_SerialNumbers WHERE [id] = " & nRecord
        Try
            oConn.Open()
            oCmd = New SqlCommand(sSQL, oConn)
            oCmd.ExecuteNonQuery()

            gvConsignments.DataSource = GetConsignments(bAll:=pbShowAll)
            gvConsignments.DataBind()
        Catch ex As Exception
            WebMsgBox.Show("Error in RemoveRecord: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub btnSaveAndAddAnother_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If SaveNumbers() Then
            tbBookNumber.Text = String.Empty
            tbFirstPageNumber.Text = String.Empty
            Call HideAllPanels()
            pnlEnterNumbers.Visible = True
            tbBookNumber.Focus()
        End If
    End Sub
    
    Property pbRecordExists() As Boolean
        Get
            Dim o As Object = ViewState("FSM_RecordExists")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("FSM_RecordExists") = Value
        End Set
    End Property
  
    Property pbShowAll() As Boolean
        Get
            Dim o As Object = ViewState("FSM_ShowAll")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("FSM_ShowAll") = Value
        End Set
    End Property
  
    Protected Sub btnFindConsignment_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        tbSerialNo.Text = String.Empty
        Dim sConsignmentNumber As String = tbConsignmentNumber.Text.Trim
        Dim oDataTable As DataTable
        gvConsignments.PageIndex = 0
        oDataTable = GetSingleConsignment(sConsignmentNumber)
        If oDataTable.Rows.Count > 0 Then
            gvConsignments.DataSource = oDataTable
            gvConsignments.DataBind()
            pnlConsignments.Visible = True
        Else
            Dim sMessage As String = "No WUIRE consignment found matching consignment number " & sConsignmentNumber
            If cbLast30Days.Checked Then
                sMessage += " in the last 30 days"
            End If
            WebMsgBox.Show(sMessage)
            tbConsignmentNumber.Text = String.Empty
        End If
    End Sub
    
    Protected Function GetSingleConsignment(ByVal sConsignmentNumber As String) As DataTable
        Dim sbSQL As New StringBuilder
        sbSQL.Append("SELECT DISTINCT fsn.[id] 'Record', c.[key] 'Consignment', lp.ProductCode + ' ' + lp.ProductDescription 'Product', fsn.BookNumber 'Book No', fsn.FirstPageNumber '1st Page No', CneeName + ', ' + CneeAddr1 + ', ' + CneeTown + ' ' + CneePostCode 'Consignee', c.CreatedOn 'Date', c.StateId 'Status' ")
        sbSQL.Append("FROM Consignment c INNER JOIN LogisticMovement lm ON c.[key] = lm.ConsignmentKey ")
        sbSQL.Append("LEFT OUTER JOIN ClientData_WUIRE_SerialNumbers fsn ON c.[key] = fsn.ConsignmentKey ")
        sbSQL.Append("LEFT OUTER JOIN LogisticProduct lp ON fsn.LogisticProductKey = lp.LogisticProductKey ")
        sbSQL.Append("WHERE(c.CustomerKey = " & WUIRE_CUSTOMER_KEY & ") ")
        sbSQL.Append("AND c.AWB = '" & sConsignmentNumber & "' ")
        sbSQL.Append("ORDER BY c.[key]")
        GetSingleConsignment = ExecuteQueryToDataTable(sbSQL.ToString)
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

    Protected Sub gvConsignments_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs)
        Dim gv As GridView = sender
        Dim gvr As GridViewRow = e.Row
        If gvr.RowType = DataControlRowType.DataRow Then
            Dim tc As TableCell = gvr.Cells(0)
            Dim b As Button = gvr.FindControl("btnAdd")
            Dim sConsignmentNumber = b.CommandArgument
            If OrderContainsMultipleSerialNumberItems(sConsignmentNumber) Then
                Dim lbl As Label = gvr.FindControl("lblMultipleSerialNumberItems")
                lbl.Visible = True
            End If
        End If
    End Sub
    
    Protected Function OrderContainsMultipleSerialNumberItems(ByVal sConsignmentNumber As String) As Boolean
        Dim oDataTable As DataTable = ExecuteQueryToDataTable("SELECT DISTINCT lm.LogisticProductKey, lm.ItemsOut FROM LogisticMovement lm INNER JOIN LogisticProduct lp ON lp.LogisticProductKey = lm.LogisticProductKey WHERE lp.SerialNumbersFlag = 'Y' AND ConsignmentKey = " & sConsignmentNumber)
        If oDataTable.Rows.Count > 1 Then
            OrderContainsMultipleSerialNumberItems = True
        Else
            OrderContainsMultipleSerialNumberItems = False
        End If
    End Function

    Protected Function gvConsignmentGetContents(ByVal DataItem As Object) As String
        Dim sConsignmentContents As String = String.Empty
        Dim oListItemCollection As ListItemCollection = ExecuteQueryToListItemCollection("SELECT DISTINCT lp.ProductCode, lm.ItemsOut FROM LogisticMovement lm INNER JOIN LogisticProduct lp ON lp.LogisticProductKey = lm.LogisticProductKey WHERE ConsignmentKey = " & DataBinder.Eval(DataItem, "Consignment"), "ProductCode", "ItemsOut")
        For Each li As ListItem In oListItemCollection
            sConsignmentContents += li.Text & " (" & li.Value & "); "
        Next
        gvConsignmentGetContents = sConsignmentContents
    End Function

    Protected Sub btnFindSerialNo_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        tbConsignmentNumber.Text = String.Empty
        tbSerialNo.Text = tbSerialNo.Text.Trim
        If tbSerialNo.Text <> String.Empty Then
            If IsNumeric(tbSerialNo.Text) Then
                Dim sbSQL As New StringBuilder
                sbSQL.Append("SELECT DISTINCT fsn.[id] 'Record', c.[key] 'Consignment', lp.ProductCode + ' ' + lp.ProductDescription 'Product', fsn.BookNumber 'Book No', fsn.FirstPageNumber '1st Page No', CneeName + ', ' + CneeAddr1 + ', ' + CneeTown + ' ' + CneePostCode 'Consignee', c.CreatedOn 'Date', c.StateId 'Status' ")
                sbSQL.Append("FROM ClientData_WUIRE_SerialNumbers fsn ")
                sbSQL.Append("INNER JOIN Consignment c ON c.[key] = fsn.ConsignmentKey ")
                sbSQL.Append("INNER JOIN LogisticMovement lm ON c.[key] = lm.ConsignmentKey ")
                sbSQL.Append("LEFT OUTER JOIN LogisticProduct lp ON fsn.LogisticProductKey = lp.LogisticProductKey ")
                sbSQL.Append("WHERE fsn.BookNumber = " & tbSerialNo.Text & " OR fsn.FirstPageNumber = " & tbSerialNo.Text & " ")
                If cbLast30Days.Checked Then
                    sbSQL.Append("AND (c.CreatedOn >= (GETDATE()-30)) ")
                End If
                sbSQL.Append("ORDER BY c.[key]")
                Dim oDataTable As DataTable = ExecuteQueryToDataTable(sbSQL.ToString)
                If oDataTable.Rows.Count = 0 Then
                    Dim sMessage As String = "No WUIRE consignment found matching serial number " & tbSerialNo.Text
                    If cbLast30Days.Checked Then
                        sMessage += " in the last 30 days"
                    End If
                    WebMsgBox.Show(sMessage)
                    tbSerialNo.Text = String.Empty
                Else
                    gvConsignments.PageIndex = 0
                    gvConsignments.DataSource = oDataTable
                    gvConsignments.DataBind()
                    pnlConsignments.Visible = True
                End If
            Else
                WebMsgBox.Show("Please enter a number")
                tbSerialNo.Text = String.Empty
            End If
        Else
            WebMsgBox.Show("Please enter a number")
        End If
    End Sub
    
</script>

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>File Upload</title>
    <style type="text/css">
        .style1
        {
            width: 30%;
        }
    </style>
    </head>
<body style="background-color:LightGreen">
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
        <table style="width: 100%; font-size: xx-small; font-family: Verdana;" border="0">
            <tr>
                <td align="left" class="style1">
        &nbsp;
        <asp:Label ID="lblTitle" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="Small"
            Text="Western Union IRELAND Serial Numbers"></asp:Label></td>
                <td style="width: 60%">
                    &nbsp;</td>
                <td style="width: 20%">
                </td>
            </tr>
        </table>
        <br />
        <asp:Panel ID="pnlEnterNumbers" runat="server" Visible="false" Width="100%" Font-Names="Verdana">
        <table style="width: 100%; font-size: xx-small; font-family: Verdana;" border="0">
            <tr>
                <td style="width: 10%">
                </td>
                <td style="width: 80%">
                </td>
                <td style="width: 10%">
                </td>
            </tr>
            <tr>
                <td align="right">
                    Consignment:
                </td>
                <td>
                    <asp:Label ID="lblConsignment" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small"/>
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td align="right">Consignee:</td>
                <td>
                    <asp:Label ID="lblConsignee" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small"/>
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td align="right">
                    Product:&nbsp;</td>
                <td>
                    <asp:DropDownList ID="ddlProduct" runat="server" Font-Names="Verdana" Font-Size="XX-Small">
                    </asp:DropDownList></td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td align="right">
                </td>
                <td>
                    &nbsp;</td>
                <td>
                </td>
            </tr>
            <tr>
                <td align="right">
                    Book number:</td>
                <td>
                    <asp:TextBox ID="tbBookNumber" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="200px"/>
                    &nbsp;<asp:RegularExpressionValidator ID="revBookNumber" ValidationExpression="\d\d\d\d\d\d\d\d" 
                        runat="server" ControlToValidate="tbBookNumber" Font-Size="XX-Small" 
                        Font-Names="Arial"> # (expected 8 digits)</asp:RegularExpressionValidator>
                    &nbsp;<asp:RequiredFieldValidator ID="rfvBookNumber" runat="server" 
                        ControlToValidate="tbBookNumber" ErrorMessage="required!"/>
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td align="right">
                    &nbsp;
                </td>
                <td>
                    &nbsp;</td>
                <td>
                </td>
            </tr>
            <tr>
                <td align="right" style="height: 26px">
                    First page number:<br />
                </td>
                <td style="height: 26px">
                    <asp:TextBox ID="tbFirstPageNumber" runat="server" Font-Names="Verdana" 
                        Font-Size="XX-Small" Width="200px"></asp:TextBox>
                    &nbsp;<asp:RegularExpressionValidator ID="revFirstPageNumber" 
                        ValidationExpression="\d\d\d\d\d\d" runat="server" 
                        ControlToValidate="tbFirstPageNumber" Font-Size="XX-Small" Font-Names="Arial"> 
                    # (expected 6 digits)</asp:RegularExpressionValidator>
                    &nbsp;<asp:RequiredFieldValidator ID="rfvFirstPageNumber" runat="server" 
                        ControlToValidate="tbFirstPageNumber" ErrorMessage="required!"></asp:RequiredFieldValidator>
                    <asp:CheckBox ID="cbOverrideValidation" runat="server" Text="override validation" Visible="False" /></td>
                <td style="height: 26px">
                </td>
            </tr>
            <tr>
                <td align="right">
                </td>
                <td style="white-space:nowrap">
                    <asp:Button ID="btnSave" runat="server" Text="save" Width="100px" onclick="btnSave_Click" />
                    <asp:Button ID="btnSaveAndAddAnother" runat="server" Text="save & add another number for this consignment" OnClick="btnSaveAndAddAnother_Click" />
                    &nbsp;<asp:Button ID="btnCancel" runat="server" onclick="btnCancel_Click" Text="cancel" Width="100px" CausesValidation="False" />
                </td>
                <td>
                </td>
            </tr>
        </table>
        </asp:Panel>

        <asp:Panel ID="pnlConsignments" runat="server" Width="100%">
            <strong>
                &nbsp;</strong><asp:LinkButton ID="lnkbtnShowConsignmentsWithoutSerialNumbers" 
                runat="server" Font-Names="Verdana" Font-Size="XX-Small" 
                onclick="lnkbtnShowConsignmentsWithoutSerialNumbers_Click" 
                Text="show consignments w/o serial nos" />
            &nbsp;
            <asp:LinkButton ID="lnkbtnShowAllConsignments" runat="server" Font-Names="Verdana" Font-Size="XX-Small" onclick="lnkbtnShowAllConsignments_Click" Text="show all consignments"/>
            &nbsp;
            <asp:CheckBox ID="cbLast30Days" runat="server" Checked="True" Font-Names="Verdana" Font-Size="XX-Small" Text="search last 30 days only" />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <asp:Label ID="Label3" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="find c'sgnmnt:"/>
            <asp:TextBox ID="tbConsignmentNumber" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="100px" />
            <asp:Button ID="btnFindConsignment" runat="server" OnClick="btnFindConsignment_Click" Text="go" />
            &nbsp;
            <asp:Label ID="Label4" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="find serial no:"/>
            <asp:TextBox ID="tbSerialNo" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="100px" />
            <asp:Button ID="btnFindSerialNo" runat="server" Text="go" onclick="btnFindSerialNo_Click" />
            <br />
            <table style="width: 100%">
                <tr>
                    <td style="width: 100%">
                        <asp:GridView ID="gvConsignments" runat="server" CellPadding="2" 
                            Font-Names="Verdana" Font-Size="XX-Small" Width="95%" 
                            EmptyDataText="no entries found" AllowPaging="True" 
                            OnPageIndexChanging="gvConsignments_PageIndexChanging" AutoGenerateColumns="False" OnRowDataBound="gvConsignments_RowDataBound" ShowFooter="True">
                            <PagerSettings Position="TopAndBottom" />
                            <Columns>
                                <asp:TemplateField>
                                    <ItemTemplate>
                                        <asp:Button ID="btnAdd" runat="server" CommandArgument='<%# DataBinder.Eval(Container.DataItem,"Consignment") %>' CommandName='<%# DataBinder.Eval(Container.DataItem,"Consignee") %>' Text="add" onclick="btnAdd_Click" />
                                        <asp:Button ID="btnRemove" runat="server" CommandArgument='<%# DataBinder.Eval(Container.DataItem,"Record") %>' Text="remove" OnClick="btnRemove_Click" />
                                        <asp:Label ID="lblMultipleSerialNumberItems" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="Medium" ForeColor="#C04000" Text="M" Visible="False"/>
                                    </ItemTemplate>
                                    <ItemStyle Wrap="False" />
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Consignment" SortExpression="Consignment">
                                    <ItemTemplate>
                                        <asp:Label ID="Label0" runat="server" Text='<%# DataBinder.Eval(Container.DataItem,"Consignment") %>' ToolTip='<%# gvConsignmentGetContents(Container.DataItem) %>'/>
                                    </ItemTemplate>
                                    <ItemStyle Wrap="False" />
                                </asp:TemplateField>
                                <asp:BoundField DataField="Product" HeaderText="Product" />
                                <asp:BoundField DataField="Book No" HeaderText="Book No" ReadOnly="True" SortExpression="Book No" >
                                    <ItemStyle Wrap="False" />
                                </asp:BoundField>
                                <asp:BoundField DataField="1st Page No" HeaderText="1st Page No" ReadOnly="True"
                                    SortExpression="1st Page No" >
                                    <ItemStyle Wrap="False" />
                                </asp:BoundField>
                                <asp:BoundField DataField="Consignee" HeaderText="Consignee" ReadOnly="True" SortExpression="Consignee" >
                                    <ItemStyle Wrap="False" />
                                </asp:BoundField>
                                <asp:BoundField DataField="Date" HeaderText="Date" ReadOnly="True" SortExpression="Date" >
                                    <ItemStyle Wrap="False" />
                                </asp:BoundField>
                                <asp:BoundField DataField="Status" HeaderText="Status" ReadOnly="True" SortExpression="Status" >
                                    <ItemStyle Wrap="False" />
                                </asp:BoundField>
                            </Columns>
                            <PagerStyle HorizontalAlign="Center" />
                            <EmptyDataTemplate>
                                <asp:Label ID="Label1" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="no consignments found"/>
                            </EmptyDataTemplate>    
                        </asp:GridView>
                        <br />
                        <asp:Label ID="lblNote" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="NOTE 1: Hover the mouse over the consignment number to see a summary of the consignment contents (product code & quantity)<br /> NOTE 2: Where <font size='medium' color='#C04000'><b>M</b></font> is displayed next to a consignment, this indicates the consignment contains more than one type of serial numbered product"></asp:Label></td>
                </tr>
            </table>
        </asp:Panel>
        <asp:Panel runat="server" ID="pnlInstructions" Visible="false" Width="100%">
            <strong>
            &nbsp;<asp:Label ID="lblInstructions" runat="server" Font-Bold="True" 
                Font-Names="Verdana" Font-Size="Small" Text="Instructions"></asp:Label>
            </strong><table width="100%" >
                <tr>
                    <td style="width:5%">
                        &nbsp;</td>
                    <td style="width:90%">
                        &nbsp;</td>
                    <td style="width:5%">
                        &nbsp;</td>
                </tr>
                <tr>
                    <td>
                        &nbsp;</td>
                    <td>
                        <asp:Label ID="Label2" runat="server" Font-Names="Verdana" Font-Size="XX-Small" 
                            
                            Text="1. Click the &lt;b&gt;Browse&lt;/b&gt; button, then browse to the file you want to upload. The file location will be displayed in the &lt;b&gt;Filename&lt;/b&gt; box."></asp:Label>
                        <br />
                        <br />
                    </td>
                    <td>
                        &nbsp;</td>
                </tr>
            </table>
        </asp:Panel>
        </div>
    </form>
    <script language="JavaScript" type="text/javascript" src="wz_tooltip.js"></script>
</body>
</html>
