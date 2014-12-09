<%@ Page Language="VB" %>
<%@ Register TagPrefix="barcode" Namespace="Xheo.ByteSize.Barcode.Web" Assembly="Xheo.ByteSize.Barcode" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>
<script runat="server">

    ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '   Copyright Jonathan Hare January 2004
    '   Courier Bookings: part of the web interface to Stock Manager
    ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '
    '       MASTER COPY
    '
    ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    
    
    
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '~~~~~~~~~~~~~ Page Load ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    
    
    
    Sub Page_Load(Source As Object, E As EventArgs)
        If Not IsPostBack Then
            Dim lConsignmentKey As Long = Request.QueryString("Key")
            lblConsignmentNumber.Text =lConsignmentKey
            Barcode.Value = lConsignmentKey
            Barcode.Barwidth = 0.05
            Barcode.IntegralWidth = True
            Barcode.Dpi = 300
            GetConsignmentFromKey(lConsignmentKey)
        End If
    End Sub
    
    
    Sub GetConsignmentFromKey(lConsignmentKey As Long)
        If lConsignmentKey > 0 Then
            ResetForm()
            Dim sCnorAddr1, sCnorAddr2, sCnorAddr3, sCnorAddr4, sCnorAddr5 As String
            Dim sCneeAddr1, sCneeAddr2, sCneeAddr3, sCneeAddr4, sCneeAddr5 As String
            Dim oDataReader As SqlDataReader
            Dim sConn As String = ConfigurationSettings.AppSettings("ConnectionString")
            Dim oConn As New SqlConnection(sConn)
            Dim oCmd As New SQLCommand("spASPNET_Consignment_GetAWBDetailsFromAWB",oConn)
            oCmd.CommandType = CommandType.StoredProcedure
            Dim oParam As New SQLParameter("@AWB", SqlDbType.NVarChar, 50)
            oCmd.Parameters.Add(oParam)
            oParam.Value = CStr(lConsignmentKey)
            Try
                oConn.Open()
                oDataReader = oCmd.ExecuteReader()
                oDataReader.Read()
                If Not IsDBNull(oDataReader("CreatedOn")) Then
                    lblDate.Text = Format(oDataReader("CreatedOn"),"dd/MM/yyyy HH:mm")
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
                If Not IsDBNull(oDataReader ("CnorAddr2")) Then
                    sCnorAddr2 &= "  " & oDataReader("CnorAddr2")
                End If
                If Not IsDBNull(oDataReader("CnorTown")) Then
                    sCnorAddr3 = oDataReader("CnorTown")
                End If
                If Not IsDBNull(oDataReader ("CnorState")) Then
                    sCnorAddr3 &= "  " & oDataReader("CnorState")
                End If
                If Not IsDBNull(oDataReader("CnorPostCode")) Then
                    sCnorAddr3 &= "  " & oDataReader("CnorPostCode")
                End If
                If Not IsDBNull(oDataReader ("CnorCountryName")) Then
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
                If Not IsDBNull(oDataReader ("CneeAddr2")) Then
                    sCneeAddr2 &= "  " & oDataReader("CneeAddr2")
                End If
                If Not IsDBNull(oDataReader("CneeTown")) Then
                    sCneeAddr3 = oDataReader("CneeTown")
                End If
                If Not IsDBNull(oDataReader ("CneeState")) Then
                    sCneeAddr3 &= "  " & oDataReader("CneeState")
                End If
                If Not IsDBNull(oDataReader("CneePostCode")) Then
                    sCneeAddr3 &= "  " & oDataReader("CneePostCode")
                End If
                If Not IsDBNull(oDataReader ("CneeCountryName")) Then
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
            Catch ex As SQLException
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
    
    Sub ResetForm()
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
<html>
<head>
</head>
<body>
    <form runat="server">
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
        <asp:Table id="tabPlainPaperAWB2" runat="server" Font-Size="X-Small" Font-Names="Verdana" Width="300px">
            <asp:TableRow>
                <asp:TableCell Wrap="False" VerticalAlign="Top">
                    <asp:Label runat="server" id="lblCustAcctCode" font-size="X-Small" font-bold="True"></asp:Label>
                </asp:TableCell>
                <asp:TableCell Wrap="False" VerticalAlign="Top" HorizontalAlign="Right">
                    <barcode:Code128 runat="server" BarcodeAlignment="MiddleRight" ID="Barcode" Height="50px"></barcode:Code128>
                </asp:TableCell>
            </asp:TableRow>
        </asp:Table>
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
