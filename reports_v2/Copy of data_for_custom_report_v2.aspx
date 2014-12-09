<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>
<script runat="server">

    '   NEW HEADERLESS VERSION FOR COMMON REPORTING FACILITY - CN
    '   Custom Reports
    
    ' N O T E S

    ' XMLDataSource xdsTables lists all tables (XPATH: dataDefinitions/tables/table)
    ' XMLDataSource xdsFields lists all fields in a table (XPATH: //table[@friendlyName='whatever']/*)

    ' Repeater rptTableData bound to xdsTables harvests hidDescription = #XPath("@description"), hidViewNameToDescription = #XPath("@ignoreCustomerKey") and hidViewName = XPath("@viewName")
    ' and puts them into hshtblViewNameToDescription to supply table description text and hshtblViewNameToDescription to supply logic for WHERE clause for retrieving non client-specific data such as Country, Currency

    ' ddlTables bound to xdsTables shows friendlyName & indexes viewName; on index change rebinds xdsFields to table friendlyName
    ' gvFields shows friendly field names in table and has HiddenField containing real field name, to build SQL statement
    ' hshtblRealToFriendly maps the (real) column names from the DataView to friendlyNames, and is built whenever data is retrieved; 

    '   TO ADD A NEW TABLE...
    '   1.  Create a view on that table, omitting any fields never to be supplied by this mechanism
    '       View *must* have a reference to CustomerKey
    '   2.  Add the relevant <table> and <column> entries to DataDefinitions.xml

    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString()
    Const DROPDOWN_TABLE_INITIAL_STRING As String = "- select a data source -"
    
    Protected Sub Page_Load(ByVal Source As Object, ByVal E As EventArgs)
        If Not IsNumeric(Session("CustomerKey")) Then
            Server.Transfer("../session_expired.aspx")
        End If
        If Not IsPostBack Then
            
            ' F O R  D E B U G G I N G
            'If Not IsNumeric(Session("CustomerKey")) Then
            ' Session("CustomerKey") = 16
            'End If

        End If
    End Sub
    
    Protected Sub ddlTables_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim sTableName As String
        If ddlTables.Items(0).Text = DROPDOWN_TABLE_INITIAL_STRING Then
            ddlTables.Items.Remove(DROPDOWN_TABLE_INITIAL_STRING)
        End If
        
        sTableName = ddlTables.SelectedItem.Text
        xdsFields.XPath = "//table[@friendlyName='" & sTableName & "']/*"
        
        gvFields.DataBind()

        lblTableDescription.Text = phshtblViewNameToDescription.Item(ddlTables.SelectedValue)
        tblRetrieveDataOptions.Visible = True
        btnRetrieveData.Visible = True
    End Sub

    Protected Sub btnRetrieveData_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim sbSQL As New StringBuilder
        Dim oGridViewRow As GridViewRow
        Dim oCheckBox As CheckBox
        Dim oHiddenField As HiddenField
        Dim oTableCell As TableCell
        
        Dim bDataItemSelected As Boolean = False
        
        Dim hshtblRealToFriendly As New Hashtable  ' index friendly names by real names

        Dim oConn As SqlConnection = New SqlConnection
        If lblError.Text <> String.Empty Then
            lblError.Text = String.Empty
        End If
        oConn.ConnectionString = gsConn
        ' oConn.ConnectionString = ConfigurationManager.ConnectionStrings("AIMSRootConnectionString").ConnectionString
        ' oConn.ConnectionString = ConfigurationManager.ConnectionStrings("LogisticsConnectionString").ConnectionString

        sbSQL.Append("SELECT ")
        If cbLimitOutput.Checked = True Then
            sbSQL.Append("TOP 10 ")
        End If
        For Each oGridViewRow In gvFields.Rows
            oCheckBox = oGridViewRow.FindControl("cbSelect")
            If oCheckBox.Checked Then
                oHiddenField = oGridViewRow.FindControl("hidFieldName")
                oTableCell = oGridViewRow.Cells(2)
                
                sbSQL.Append(oHiddenField.Value)
                sbSQL.Append(", ")
                
                Dim sBrackets As String = "[]"
                hshtblRealToFriendly.Add(oHiddenField.Value.Trim(sBrackets.ToCharArray), oTableCell.Text)
                
                bDataItemSelected = True
            End If
        Next
        
        If Not bDataItemSelected Then
            ' MsgBox("Please select at least one item of data to retrieve", , "Nothing selected!")
            Exit Sub
        End If
        
        sbSQL.Remove(sbSQL.Length - 2, 2)
        sbSQL.Append(" FROM ")
        sbSQL.Append(ddlTables.SelectedValue)

        Dim bIgnoreFlag As Boolean = CBool(phshtblViewNameToIgnoreCustomerKey.Item(ddlTables.SelectedValue))
        If Not bIgnoreFlag Then
            sbSQL.Append(" WHERE CustomerKey = ")
            sbSQL.Append(Session("CustomerKey").ToString)
        End If
        
        Dim sSQL As String = sbSQL.ToString
        Dim oAdapter As New SqlDataAdapter(sSQL, oConn)
        Dim oDataSet As New DataSet()
        Try
            oAdapter.Fill(oDataSet, "thisTable")
            Dim Source As DataView = oDataSet.Tables("thisTable").DefaultView
            If Source.Count > 0 Then
                Response.Clear()
                Response.ContentType = rblFileExtension.Items(rblFileExtension.SelectedIndex).Value
                'Response.ContentType = "Application/x-msexcel"
                
                Dim sResponseValue As New StringBuilder
                sResponseValue.Append("attachment; filename=""")
                sResponseValue.Append(ddlTables.SelectedItem)

                sResponseValue.Append(rblFileExtension.Items(rblFileExtension.SelectedIndex).Text)
                sResponseValue.Append("""")
                
                Response.AddHeader("Content-Disposition", sResponseValue.ToString)
                'Response.AddHeader("Content-Disposition", "attachment; filename=products.csv")
    
                Dim r As DataRowView
                Dim c As DataColumn
                Dim sItem As String
    
                Dim IgnoredItems As New ArrayList
                'IgnoredItems.Add("")   ' add name of any field that is not to be output
                
                If cbIncludeFieldNames.Checked = True Then
                    For Each c In Source.Table.Columns
                        If Not IgnoredItems.Contains(c.ColumnName) Then
                            Response.Write(hshtblRealToFriendly(c.ColumnName))
                            Response.Write(",")
                        End If
                    Next
                    Response.Write(vbCrLf)
                End If
    
                For Each r In Source
                    For Each c In Source.Table.Columns
                        If Not IgnoredItems.Contains(c.ColumnName) Then
                            sItem = (r(c.ColumnName).ToString)
                            sItem = sItem.Replace(ControlChars.Quote, ControlChars.Quote & ControlChars.Quote)
                            sItem = ControlChars.Quote & sItem & ControlChars.Quote
                            Response.Write(sItem)
                            Response.Write(",")
                        End If
                    Next
                    Response.Write(vbCrLf)
                Next
                Response.End()
            Else
                lblError.Text = "No data found"
            End If
        Catch ex As SqlException
            lblError.Text = ex.ToString
        Finally
            oConn.Dispose()
        End Try
    End Sub

    Protected Sub lnkbtnSelectAll_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim oGridViewRow As GridViewRow
        Dim oCheckBox As CheckBox
        For Each oGridViewRow In gvFields.Rows
            oCheckBox = oGridViewRow.FindControl("cbSelect")
            oCheckBox.Checked = True
        Next
    End Sub

    Protected Sub lnkbtnClearAll_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim oGridViewRow As GridViewRow
        Dim oCheckBox As CheckBox
        For Each oGridViewRow In gvFields.Rows
            oCheckBox = oGridViewRow.FindControl("cbSelect")
            oCheckBox.Checked = False
        Next
    End Sub

    Protected Sub rptTableData_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.RepeaterItemEventArgs)
        Dim hidDescription As HiddenField
        Dim hidViewName As HiddenField
        Dim hidIgnoreCustomerKey As HiddenField
        Dim sValue1 As String
        Dim sValue2 As String

        Dim x As RepeaterItemEventArgs = e
        hidDescription = x.Item.FindControl("hidDescription")
        hidViewName = x.Item.FindControl("hidViewName")
        hidIgnoreCustomerKey = x.Item.FindControl("hidIgnoreCustomerKey")
        sValue1 = hidDescription.Value
        sValue2 = hidViewName.Value
        
        Dim oHashTable As Hashtable
        Try
            oHashTable = phshtblViewNameToDescription
            oHashTable.Add(hidViewName.Value, hidDescription.Value)
            phshtblViewNameToDescription = oHashTable
            
            oHashTable = phshtblViewNameToIgnoreCustomerKey
            oHashTable.Add(hidViewName.Value, hidIgnoreCustomerKey.Value)
            phshtblViewNameToIgnoreCustomerKey = oHashTable
        Catch
        End Try
    End Sub

    Protected Sub ddlTables_DataBound(ByVal sender As Object, ByVal e As System.EventArgs)
        ddlTables.Items.Insert(0, DROPDOWN_TABLE_INITIAL_STRING)
    End Sub

    Property psViewName() As String
        Get
            Dim o As Object = ViewState("ViewName")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("ViewName") = Value
        End Set
    End Property
    
    Property phshtblViewNameToDescription() As Hashtable
        Get
            Dim o As Object = ViewState("ViewNameToDescription")
            If o Is Nothing Then
                Return New Hashtable
            End If
            Return CType(o, Hashtable)
        End Get
        Set(ByVal Value As Hashtable)
            ViewState("ViewNameToDescription") = Value
        End Set
    End Property
    
    Property phshtblViewNameToIgnoreCustomerKey() As Hashtable
        Get
            Dim o As Object = ViewState("ViewNameToIgnoreCustomerKey")
            If o Is Nothing Then
                Return New Hashtable
            End If
            Return CType(o, Hashtable)
        End Get
        Set(ByVal Value As Hashtable)
            ViewState("ViewNameToIgnoreCustomerKey") = Value
        End Set
    End Property
    
</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Data for Custom Reports</title>
    <link rel="stylesheet" type="text/css" href="../css/sprint.css" />
</head>
<body>
    <form id="frmDataForCustomReports" runat="server">
            <asp:Table id="Table4" runat="server" width="100%">
                <asp:TableRow>
                    <asp:TableCell VerticalAlign="Bottom" width="0%"></asp:TableCell>
                    <asp:TableCell Wrap="False" width="50%">
                        <asp:Label ID="Label1"
                                   runat="server"
                                   forecolor="silver"
                                   font-size="Small"
                                   font-bold="True"
                                   font-names="Arial">Data for Custom Reports</asp:Label><br />
                    </asp:TableCell>
                    <asp:TableCell Wrap="False" HorizontalAlign="Right" width="50%"></asp:TableCell>
                </asp:TableRow>
            </asp:Table>
        &nbsp;<asp:Repeater ID="rptTableData"
        
                      runat="server"
                      DataSourceID="xdsTables"
                      OnItemDataBound="rptTableData_ItemDataBound">
                <ItemTemplate>
                  <asp:HiddenField runat="server"
                             ID="hidDescription"
                             Value='<%#XPath("@description")%>'>
                  </asp:HiddenField>
                  <asp:HiddenField runat="server"
                             ID="hidViewName"
                             Value='<%#XPath("@viewName")%>'>
                  </asp:HiddenField>
                  <asp:HiddenField runat="server"
                             ID="hidIgnoreCustomerKey"
                             Value='<%#XPath("@ignoreCustomerKey")%>'>
                  </asp:HiddenField>
                </ItemTemplate>
        </asp:Repeater>
        <table width="100%">
            <tr>
                <td style="width: 25%; height: 43px">
                    <asp:DropDownList ID="ddlTables"
                                      runat="server"
                                      AutoPostBack="True"
                                      DataSourceID="xdsTables"
                                      DataTextField="friendlyName"
                                      DataValueField="viewName"
                                      OnSelectedIndexChanged="ddlTables_SelectedIndexChanged"
                                      OnDataBound="ddlTables_DataBound" Font-Names="Verdana" 
                        Font-Size="XX-Small">
                    </asp:DropDownList>
                  </td>
                <td width="5%" style="height: 43px">
                   </td>
                <td width="75%" valign="top" style="height: 43px">
                    <asp:Label ID="lblTableDescription"
                               runat="server"
                               Font-Names="Verdana,Sans-Serif"
                               Font-Size="XX-Small"
                               ></asp:Label></td>
            </tr>
        </table>
        <br />
        <asp:GridView ID="gvFields"
                      runat="server"
                      AutoGenerateColumns="False"
                      CellPadding="4"
                      DataSourceID="xdsFields"
                      ForeColor="Black"
                      GridLines="Vertical"
                      BackColor="White"
                      BorderColor="#DEDFDE"
                      BorderStyle="None"
                      BorderWidth="1px"
                      Font-Names="verdana, sans-serif"
                      Font-Size="XX-Small" Width="100%" ShowFooter="true">
            <FooterStyle BackColor="#EEEEEE" />
            <Columns>
                <asp:TemplateField>
                    <ItemTemplate>
                        <asp:CheckBox ID="cbSelect" runat="server" Font-Names="Verdana" 
                            Font-Size="XX-Small" />
                        <asp:HiddenField ID="hidFieldName"  runat="server" Value='<%# Eval("name") %>' />
                    </ItemTemplate>
                    <FooterTemplate>
                        <asp:LinkButton ID="lnkbtnSelectAll" runat="server" OnClick="lnkbtnSelectAll_Click">Select&nbsp;All</asp:LinkButton>
                        <asp:LinkButton ID="lnkbtnClearAll" runat="server" OnClick="lnkbtnClearAll_Click">Clear&nbsp;All</asp:LinkButton>
                    </FooterTemplate>
                </asp:TemplateField>
                <asp:BoundField DataField="name" HeaderText="name" SortExpression="name" Visible="False" />
                <asp:BoundField DataField="friendlyname" HeaderText="Data Item" SortExpression="friendlyname" >
                </asp:BoundField>
                <asp:BoundField DataField="description" HeaderText="Description" SortExpression="description" />
            </Columns>
            <RowStyle BackColor="#eeeeee" />
            <SelectedRowStyle BackColor="#CE5D5A" Font-Bold="True" ForeColor="White" />
            <PagerStyle BackColor="#F7F7DE" ForeColor="Black" HorizontalAlign="Right" />
            <HeaderStyle BackColor="#6B696B" Font-Bold="True" ForeColor="White" />
            <AlternatingRowStyle BackColor="White" />
        </asp:GridView>
        <table runat="server" id="tblRetrieveDataOptions" visible="false">
            <tr>
                <td style="width: 186px; height: 22px;">
                    <asp:CheckBox ID="cbLimitOutput" runat="server" Font-Size="XX-Small" Text="Return 10 items max."
                        ToolTip="Limit the number of items returned to 10" Width="153px" 
                        Font-Names="Verdana" /></td>
                <td style="width: 190px; height: 22px;">
                    <asp:CheckBox ID="cbIncludeFieldNames" runat="server" Checked="True" Font-Size="XX-Small"
                        Text="Include field names" ToolTip="Show field names in row 1" 
                        Width="184px" Font-Names="Verdana" /></td>
                <td style="width: 400px; height: 22px;">
                    <asp:Label ID="lblSaveAs" runat="server" Font-Size="XX-Small" Text="Save as" 
                        Font-Names="Verdana"></asp:Label>&nbsp;<asp:RadioButtonList ID="rblFileExtension" runat="server" Font-Size="XX-Small" RepeatDirection="Horizontal" RepeatLayout="Flow" ToolTip=".csv - Excel; .doc - Word; .txt - Notepad">
                        <asp:ListItem Selected="True" Value="text/csv">.csv </asp:ListItem>
                        <asp:ListItem Value="Application/x-msword">.doc </asp:ListItem>
                        <asp:ListItem Value="text/plain">.txt </asp:ListItem>
                    </asp:RadioButtonList></td>
            </tr>
        </table>
        &nbsp;<br />
        &nbsp;
        <asp:Button ID="btnRetrieveData" runat="server" Text=" retrieve data " 
                Visible="false" OnClick="btnRetrieveData_Click" />
        <br />
        <br />
        <br />
            <asp:Label ID="lblReportGeneratedDateTime" Visible="false" runat="server" Text="" font-size="XX-Small" font-names="Verdana, Sans-Serif" forecolor="Green"></asp:Label><br />
        <asp:Label id="lblError" runat="server" font-size="XX-Small" font-names="Arial" forecolor="red"></asp:Label>&nbsp;
        <br />
        
        <asp:XmlDataSource ID="xdsTables" runat="server" DataFile="~/DataDefinitions.xml"
            XPath="dataDefinitions/tables/table"></asp:XmlDataSource>
            
        <asp:XmlDataSource ID="xdsFields" runat="server" DataFile="~/DataDefinitions.xml"
            XPath="//table[@friendlyName='whatever']/*"></asp:XmlDataSource>
            
    </form>
</body>
</html>
