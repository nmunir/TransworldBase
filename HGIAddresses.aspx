<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Collections.Generic" %>

<script runat="server">

    Const COUNTRY_KEY_UK As Int32 = 222
    Const MAX_RECORDS_RETURNED As Int32 = 500
    Const CUSTOMER_HGIWF As Int32 = 726
    Const CUSTOMER_HGIIT As Int32 = 727
    Const CUSTOMER_HGIOPT As Int32 = 783
    
    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString()
    
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsPostBack Then
            Call SetTitle()
            tbToConsignmentNo.Text = ExecuteQueryToDataTable("SELECT MAX([key]) FROM Consignment").Rows(0).Item(0)
            tbFromConsignmentNo.Focus()
        End If
    End Sub

    Protected Sub SetTitle()
        Dim sTitle As String = Session("SiteTitle")
        If sTitle <> String.Empty Then
            sTitle += " - "
        End If
        Page.Header.Title = sTitle & "Reports"
    End Sub

    Protected Sub HideAllPanels()
        pnlAddresses.Visible = False
    End Sub
    
    Protected Sub btnShowAddresses_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Page.Validate()
        If Page.IsValid Then
            If CInt(tbFromConsignmentNo.Text) > CInt(tbToConsignmentNo.Text) Then
                WebMsgBox.Show("You specified a range with a start greater than the end.")
                Exit Sub
            End If
            Dim sSQL As String = "SELECT COUNT ([key]) FROM Consignment WHERE [key] >= " & tbFromConsignmentNo.Text & " AND [key] <= " & tbToConsignmentNo.Text & " AND (CustomerKey = 726 OR CustomerKey = 727 OR CustomerKey = 783) AND StateId <> 'CANCELLED'"
            Dim nRecordCount As Int32 = ExecuteQueryToDataTable(sSQL).Rows(0).Item(0)
            If nRecordCount > MAX_RECORDS_RETURNED Then
                WebMsgBox.Show("The range of consignment numbers you entered would return more than " & MAX_RECORDS_RETURNED & " records (" & nRecordCount.ToString & "). Please specify a narrower range.")
            Else
                Call RetrieveConsignments()
            End If
        End If
    End Sub

    Protected Sub RetrieveConsignments()
        'Dim sSQL As String = "SELECT * FROM Consignment WHERE [key] >= " & tbFromConsignmentNo.Text & " AND [key] <= " & tbToConsignmentNo.Text & " AND (CustomerKey = 726 OR CustomerKey = 727) AND StateId <> 'CANCELLED' ORDER BY [key]"
        Dim sSQL As String = "SELECT [key], CreatedOn, CustomerKey, ISNULL(CustomerRef1,'') 'CustomerRef1', ISNULL(CustomerRef2,'') 'CustomerRef2', ISNULL(Misc1,'') 'Misc1', ISNULL(Misc2,'') 'Misc2', ISNULL(CneeCtcName,'') 'CneeCtcName', ISNULL(CneeName,'') 'CneeName', CneeAddr1, CneeAddr2, ISNULL(CneeAddr3,'') 'CneeAddr3', CneeTown, CneeState, CneePostCode, CneeCountryKey FROM Consignment WHERE [key] >= " & tbFromConsignmentNo.Text & " AND [key] <= " & tbToConsignmentNo.Text & " AND (CustomerKey = 726 OR CustomerKey = 727 OR CustomerKey = 783) AND StateId <> 'CANCELLED' ORDER BY [key]"
		'webmsgbox.show(ssql)
		'exit sub
        Dim oDT As DataTable = ExecuteQueryToDataTable(sSQL)
        If oDT.Rows.Count > 0 Then
            Response.Clear()
            Response.ContentType = "text/csv"
            Dim sResponseValue As New StringBuilder
            sResponseValue.Append("attachment; filename=""")
            sResponseValue.Append(tbFromConsignmentNo.Text)
            sResponseValue.Append(" - ")
            sResponseValue.Append(tbToConsignmentNo.Text)
            sResponseValue.Append(".csv")
            sResponseValue.Append("""")
            Response.AddHeader("Content-Disposition", sResponseValue.ToString)

            For Each dr As DataRow In oDT.Rows
                Dim lstFields As New List(Of String)
                lstFields.Add(dr("key"))
                lstFields.Add(dr("CreatedOn"))
                If dr("CustomerKey") = 727 Then
                    lstFields.Add("IT")
                ElseIf dr("CustomerKey") = 726 Then
                    lstFields.Add("OEIC/UT")
                Else
                    lstFields.Add("OPT")
                End If
                lstFields.Add(dr("CustomerRef1") & String.Empty)
                lstFields.Add(dr("CustomerRef2") & String.Empty)
                lstFields.Add(dr("Misc1") & String.Empty)
                lstFields.Add(dr("Misc2") & String.Empty)
                lstFields.Add(GetConsignmentContents(dr("key")))
                lstFields.Add(dr("CneeCtcName") & String.Empty)
                lstFields.Add(dr("CneeName") & String.Empty)
                lstFields.Add(dr("CneeAddr1") & String.Empty)
                lstFields.Add(dr("CneeAddr2") & String.Empty)
                lstFields.Add(dr("CneeAddr3") & String.Empty)
                lstFields.Add(dr("CneeTown") & String.Empty)
                lstFields.Add(dr("CneeState") & String.Empty)
                lstFields.Add(dr("CneePostCode") & String.Empty)
                If dr("CneeCountryKey") = COUNTRY_KEY_UK Then
                    lstFields.Add("U.K.")
                Else
                    lstFields.Add(ExecuteQueryToDataTable("SELECT CountryName FROM Country WHERE CountryKey = " & dr("CneeCountryKey")).Rows(0).Item(0))
                End If
                    
                'lstFields.Add(dr(""))
                'lstFields.Add(dr(""))
                'lstFields.Add(dr(""))
                'lstFields.Add(dr(""))
                
                For Each s As String In lstFields
                    Dim sItem As String = s.Replace(ControlChars.Quote, ControlChars.Quote & ControlChars.Quote)
                    sItem = ControlChars.Quote & sItem & ControlChars.Quote
                    Response.Write(sItem)
                    Response.Write(",")
                Next
                Response.Write(vbCrLf)
            Next
            Response.End()
        Else
            WebMsgBox.Show("No matching consignment in this range.")
        End If
    End Sub
    
    Protected Function GetConsignmentContents(ByVal sConsignmentKey As String) As String
        GetConsignmentContents = String.Empty
        Dim sSQL As String = "SELECT ProductCode + ' ' + ProductDescription + ' (Qty: ' + CAST(ItemsOut AS varchar(6)) + ') ' FROM LogisticMovement lm INNER JOIN LogisticProduct lp ON lm.LogisticProductKey = lp.LogisticProductKey WHERE lm.ConsignmentKey = " & sConsignmentKey
        Dim oDT As DataTable
        oDT = ExecuteQueryToDataTable(sSQL)
        For Each dr As DataRow In oDT.Rows
            GetConsignmentContents += dr(0)
        Next
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

</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <link href="sprint.css" rel="stylesheet" type="text/css" />
    <style type="text/css">
        .style1
        {
            height: 30px;
        }
    </style>
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
        <asp:Panel ID="pnlAddresses" runat="server" Width="100%">
            &nbsp; <strong><span style="font-size: 10pt; color: #000080">&nbsp;HGI Addresses</span></strong><table style="width: 100%">
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td align="right" style="width: 20%">
                    </td>
                    <td style="width: 29%">
                    </td>
                    <td style="width: 20%">
                    </td>
                    <td align="right" style="width: 29%">
                    </td>
                    <td style="width: 1%">
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td>
                        &nbsp;</td>
                    <td align="left" colspan="3" style="white-space: nowrap">
                        <asp:Label ID="Label2" runat="server" Text="Start consignment #" Font-Names="Verdana" Font-Size="XX-Small"
                            Font-Bold="False" />
                        &nbsp;<asp:TextBox ID="tbFromConsignmentNo" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="65px"></asp:TextBox>
                        <a ID="aHelpFromOrderNo" runat="server" onmouseover="return escape('')" style="color: gray; cursor: help" visible="false">&nbsp;?&nbsp;</a>
                        <asp:Label ID="Label3" runat="server" Text="End consignment #" Font-Names="Verdana" Font-Size="XX-Small"
                            Font-Bold="False" />
                        &nbsp;<asp:TextBox ID="tbToConsignmentNo" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="65px"></asp:TextBox>
                        <a ID="aHelpToOrderNo" runat="server" onmouseover="return escape('')" style="color: gray; cursor: help" visible="false">&nbsp;?&nbsp;</a>
                        <asp:RangeValidator ID="rvFromOrderNo" runat="server" ControlToValidate="tbFromConsignmentNo"
                            ErrorMessage="invalid start #" MaximumValue="9999999" MinimumValue="1" Type="Integer"
                            ValidationGroup="ByOrderNo"></asp:RangeValidator>
                        <asp:RangeValidator ID="RangeValidator1" runat="server" ControlToValidate="tbToConsignmentNo"
                            ErrorMessage="invalid end #" MaximumValue="9999999" MinimumValue="1" Type="Integer"
                            ValidationGroup="ByOrderNo"></asp:RangeValidator>
                        <asp:RequiredFieldValidator ID="rfvFromOrderNo" runat="server" ControlToValidate="tbFromConsignmentNo"
                            ErrorMessage="start # required" ValidationGroup="ByOrderNo"></asp:RequiredFieldValidator>
                        <asp:RequiredFieldValidator ID="rfvToOrderNo" runat="server" ControlToValidate="tbToConsignmentNo"
                            ErrorMessage="end # required" ValidationGroup="ByOrderNo"></asp:RequiredFieldValidator></td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td class="style1">
                        </td>
                    <td class="style1">
                        </td>
                    <td align="left" colspan="3" style="white-space: nowrap" class="style1">
                        <asp:Button ID="btnShowAddresses" runat="server" onclick="btnShowAddresses_Click" Text="show addresses" Width="170px" />
                    </td>
                    <td class="style1">
                        </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td colspan="4">
                        &nbsp;</td>
                    <td>
                    </td>
                </tr>
            </table>
        </asp:Panel>
    </form>
</body>
</html>
