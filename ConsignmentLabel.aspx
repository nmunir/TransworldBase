<%@ Page Language="VB" %>
<%@ Register TagPrefix="barcode2" Assembly="Barcode, Version=1.0.5.40001, Culture=neutral, PublicKeyToken=6dc438ab78a525b3" Namespace="Lesnikowski.Web" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>
<script runat="server">

    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString()

    Protected Sub Page_Load(ByVal Source As Object, ByVal E As EventArgs)
        If Not IsPostBack Then
            Dim lConsignmentKey As Long = Request.QueryString("Key")
            lblConsignmentNumber.Text = lConsignmentKey
            barcode.Number = lConsignmentKey.ToString
            GetConsignmentFromKey(lConsignmentKey)
        End If
    End Sub
    
    Protected Sub GetConsignmentFromKey(ByVal lConsignmentKey As Long)
        If lConsignmentKey > 0 Then
            ResetForm()
            Dim sCnorAddr1 As String = String.Empty, sCnorAddr2 As String = String.Empty, sCnorAddr3 As String = String.Empty, sCnorAddr4 As String = String.Empty, sCnorAddr5 As String = String.Empty
            Dim sCneeAddr1 As String = String.Empty, sCneeAddr2 As String = String.Empty, sCneeAddr3 As String = String.Empty, sCneeAddr4 As String = String.Empty, sCneeAddr5 As String = String.Empty
            Dim oDataReader As SqlDataReader
            Dim oConn As New SqlConnection(gsConn)
            Dim oCmd As New SqlCommand("spASPNET_Consignment_GetAWBDetailsFromAWB", oConn)
            oCmd.CommandType = CommandType.StoredProcedure
            Dim oParam As New SqlParameter("@AWB", SqlDbType.NVarChar, 50)
            oCmd.Parameters.Add(oParam)
            oParam.Value = CStr(lConsignmentKey)
            Try
                oConn.Open()
                oDataReader = oCmd.ExecuteReader()
                oDataReader.Read()
                If Not IsDBNull(oDataReader("CreatedOn")) Then
                    lblDate.Text = Format(oDataReader("CreatedOn"), "dd/MM/yyyy HH:mm")
                End If
                If Not IsDBNull(oDataReader("CnorName")) Then
                    sCnorAddr1 = oDataReader("CnorName")
                End If
                If Not IsDBNull(oDataReader("CustomerAccountCode")) Then
                    lblCustAcctCode.Text = oDataReader("CustomerAccountCode")
                End If
                If Not IsDBNull(oDataReader("CnorAddr1")) Then
                    sCnorAddr2 = oDataReader("CnorAddr1")
                End If
                If Not IsDBNull(oDataReader("CnorAddr2")) Then
                    sCnorAddr2 &= "  " & oDataReader("CnorAddr2")
                End If
                If Not IsDBNull(oDataReader("CnorTown")) Then
                    sCnorAddr3 = oDataReader("CnorTown")
                End If
                If Not IsDBNull(oDataReader("CnorState")) Then
                    sCnorAddr3 &= "  " & oDataReader("CnorState")
                End If
                If Not IsDBNull(oDataReader("CnorPostCode")) Then
                    sCnorAddr3 &= "  " & oDataReader("CnorPostCode")
                End If
                If Not IsDBNull(oDataReader("CnorCountryName")) Then
                    sCnorAddr4 = oDataReader("CnorCountryName")
                End If
                If Not IsDBNull(oDataReader("CnorCtcName")) Then
                    sCnorAddr5 = oDataReader("CnorCtcName")
                End If
                If Not IsDBNull(oDataReader("CnorTel")) Then
                    sCnorAddr5 &= "  " & oDataReader("CnorTel")
                End If
    
                If Not IsDBNull(oDataReader("CneeName")) Then
                    sCneeAddr1 = oDataReader("CneeName")
                End If
                If Not IsDBNull(oDataReader("CneeAddr1")) Then
                    sCneeAddr2 = oDataReader("CneeAddr1")
                End If
                If Not IsDBNull(oDataReader("CneeAddr2")) Then
                    sCneeAddr2 &= "  " & oDataReader("CneeAddr2")
                End If
                If Not IsDBNull(oDataReader("CneeTown")) Then
                    sCneeAddr3 = oDataReader("CneeTown")
                End If
                If Not IsDBNull(oDataReader("CneeState")) Then
                    sCneeAddr3 &= "  " & oDataReader("CneeState")
                End If
                If Not IsDBNull(oDataReader("CneePostCode")) Then
                    sCneeAddr3 &= "  " & oDataReader("CneePostCode")
                End If
                If Not IsDBNull(oDataReader("CneeCountryName")) Then
                    sCneeAddr4 = oDataReader("CneeCountryName")
                End If
                If Not IsDBNull(oDataReader("CneeCtcName")) Then
                    sCneeAddr5 = oDataReader("CneeCtcName")
                End If
                If Not IsDBNull(oDataReader("CneeTel")) Then
                    sCneeAddr5 &= " Tel:" & oDataReader("CneeTel")
                End If
                If Not IsDBNull(oDataReader("Weight")) Then
                    If oDataReader("Weight") <> "0" Then
                        lblWeight.Text = oDataReader("Weight")
                    End If
                End If
                If Not IsDBNull(oDataReader("NOP")) Then
                    If oDataReader("NOP") <> "0" Then
                        lblNOP.Text = oDataReader("NOP")
                    End If
                End If
                If Not IsDBNull(oDataReader("SpecialInstructions")) Then
                    lblSpclInstructions.Text = oDataReader("SpecialInstructions")
                End If
                If Not IsDBNull(oDataReader("ValForCustoms")) Then
                    If oDataReader("ValForCustoms") = "0" Then
                        'lblValForCustoms.Text = "NIL"
                    Else
                        'lblValForCustoms.Text = Format(oDataReader("ValForCustoms"),"#,###.00")
                    End If
                End If
                If Not IsDBNull(oDataReader("Description")) Then
                    'lblContents.Text = oDataReader("Description")
                End If
                If Not IsDBNull(oDataReader("CustomerRef1")) Then
                    lblShippersRef1.Text = oDataReader("CustomerRef1")
                End If
                If Not IsDBNull(oDataReader("CustomerRef2")) Then
                    lblShippersRef2.Text = oDataReader("CustomerRef2")
                End If
                oDataReader.Close()
            Catch ex As SqlException
                'lblError.Text = ex.ToString
            Finally
                oConn.Close()
            End Try
            lblCneeAddr1.Text = sCneeAddr1
            lblCneeAddr2.Text = sCneeAddr2
            lblCneeAddr3.Text = sCneeAddr3
            lblCneeAddr4.Text = sCneeAddr4
            lblCneeAddr5.Text = sCneeAddr5
        End If
    End Sub
    
    Protected Sub ResetForm()
        lblDate.Text = ""
        lblCneeAddr1.Text = ""
        lblCneeAddr2.Text = ""
        lblCneeAddr3.Text = ""
        lblCneeAddr4.Text = ""
        lblCneeAddr5.Text = ""
        lblWeight.Text = ""
        lblNOP.Text = ""
        lblSpclInstructions.Text = ""
        'lblValForCustoms.Text = ""
        'lblContents.Text = ""
    End Sub

</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Consignment Label</title>
</head>
<body>
    <form id="frmConsignmentLabel"runat="server">
        <asp:Table id="tabPlainPaperAWB1" runat="server" Font-Size="X-Small" Font-Names="Verdana" Width="300px" font-bold="True">
            <asp:TableRow>
                <asp:TableCell Wrap="False">
                    <asp:Label runat="server" font-size="Small">Transworld</asp:Label>
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell Wrap="False">
                    <asp:Label runat="server" id="lblDate" font-size="XX-Small"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell Wrap="False">
                    <asp:Label runat="server" id="lblConsignmentNumber" font-size="Large"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>
        </asp:Table>
        <table  style=" font-size:x-small; font-family:Verdana; width:300px">
            <tr>
                <td valign="top", style="white-space:nowrap">
                    <asp:Label runat="server" id="lblCustAcctCode" font-size="X-Small" font-bold="True"></asp:Label>
                </td>
                <td valign="top" align="right" style="white-space:nowrap">
                    <barcode2:BarcodeControl ID="barcode" Symbology="Code128" XDpi="300" YDpi="300" NarrowBarWidth="2" Height="50" IsNumberVisible="false" runat="server"/>
                </td>
            </tr>
        </table>
        <asp:Table id="tabPlainPaperAWB3" runat="server" Font-Size="X-Small" Font-Names="Verdana" Width="300px" font-bold="True">
            <asp:TableRow>
                <asp:TableCell>
                    <hr />
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell>
                    <asp:Label runat="server" id="lblCneeAddr1"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell>
                    <asp:Label runat="server" id="lblCneeAddr2"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell>
                    <asp:Label runat="server" id="lblCneeAddr3"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell>
                    <asp:Label runat="server" id="lblCneeAddr4"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell>
                    <asp:Label runat="server" id="lblCneeAddr5"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell>
                    <hr />
                </asp:TableCell>
            </asp:TableRow>
        </asp:Table>
        <asp:Table id="tabPlainPaperAWB4" runat="server" Font-Size="X-Small" Font-Names="Verdana" Width="300px" font-bold="True">
            <asp:TableRow>
                <asp:TableCell Width="80">
                    <asp:Label runat="server">Spcl Instr: </asp:Label>
                </asp:TableCell>
                <asp:TableCell Width="220">
                    <asp:Label runat="server" id="lblSpclInstructions"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell ColumnSpan="2">
                    <hr />
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell>
                    <asp:Label runat="server">NOP: </asp:Label>
                </asp:TableCell>
                <asp:TableCell>
                    <asp:Label runat="server" id="lblNOP"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell>
                    <asp:Label runat="server">Weight: </asp:Label>
                </asp:TableCell>
                <asp:TableCell>
                    <asp:Label runat="server" id="lblWeight"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell>
                    <asp:Label runat="server">Ref 1: </asp:Label>
                </asp:TableCell>
                <asp:TableCell>
                    <asp:Label runat="server" id="lblShippersRef1"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell>
                    <asp:Label runat="server">Ref 2: </asp:Label>
                </asp:TableCell>
                <asp:TableCell>
                    <asp:Label runat="server" id="lblShippersRef2"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>
        </asp:Table>
    </form>
</body>
</html>
