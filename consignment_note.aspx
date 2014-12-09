<%@ Page Language="VB" %>
<%@ Register TagPrefix="barcode" Namespace="Xheo.ByteSize.Barcode.Web" Assembly="Xheo.ByteSize.Barcode, Version=1.1.5000.0, Culture=neutral, PublicKeyToken=798276055709c98a" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>
<script runat="server">

    '       MASTER COPY - Copyright Jonathan Hare January 2004
    
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
            Dim sConn As String = ConfigurationManager.AppSettings("ConnectionString")
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
                    lblDate1.Text = Format(oDataReader("CreatedOn"),"dd.MM.yyyy")
                    lblDate2.Text = Format(oDataReader("CreatedOn"),"HH:mm")
                End If
                If Not IsDBNull(oDataReader("CnorName")) Then
                    sCnorAddr1 = oDataReader("CnorName")
                End If
                If Not IsDBNull(oDataReader("CustomerAccountCode")) Then
                    lblCustomerAccountCode.Text = "[" & oDataReader("CustomerAccountCode") & "]"
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
                    sCneeAddr5 &= "  " & oDataReader("CneeTel")
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
                        lblValForCustoms.Text = "NIL"
                    Else
                        lblValForCustoms.Text = Format(oDataReader("ValForCustoms"),"#,###.00")
                        If Not IsDBNull(oDataReader("ValForCustomsCurrency")) Then
                            lblValForCustomsCurrency.Text = oDataReader("ValForCustomsCurrency")
                        End If
                    End If
                End If
    
                If Not IsDBNull(oDataReader("ValForIns")) Then
                    If oDataReader("ValForIns") = "0" Then
                        lblValForInsurance.Text = "NIL"
                    Else
                        lblValForInsurance.Text = Format(oDataReader("ValForIns"),"#,###.00")
                        If Not IsDBNull(oDataReader("ValForInsCurrency")) Then
                            lblValForInsuranceCurrency.Text  = oDataReader("ValForInsCurrency")
                        End If
                    End If
                End If
    
                If Not IsDBNull(oDataReader("Description")) Then
                    lblContents.Text = oDataReader("Description")
                End If
                If Not IsDBNull(oDataReader("CustomerRef1")) Then
                    lblShippersRef1.Text = oDataReader("CustomerRef1")
                End If
                If Not IsDBNull(oDataReader("CustomerRef2")) Then
                    lblShippersRef2.Text = oDataReader("CustomerRef2")
                End If
                lblPrinted1.Text = Format(now,"dd MMM yyyy")
                lblPrinted2.Text = Format(now,"HH:mm")
                oDataReader.Close()
            Catch ex As SQLException
                lblError.Text = ex.ToString
            Finally
                oConn.Close()
            End Try
    
            lblCnorAddr1.Text = sCnorAddr1
            lblCnorAddr2.Text = sCnorAddr2
            lblCnorAddr3.Text = sCnorAddr3
            lblCnorAddr4.Text = sCnorAddr4
            lblCnorAddr5.Text = sCnorAddr5
    
            lblCneeAddr1.Text = sCneeAddr1
            lblCneeAddr2.Text = sCneeAddr2
            lblCneeAddr3.Text = sCneeAddr3
            lblCneeAddr4.Text = sCneeAddr4
            lblCneeAddr5.Text = sCneeAddr5
    
        End If
    End Sub
    
    Sub ResetForm()
        lblDate1.Text = ""
        lblDate2.Text = ""
        lblCnorAddr1.Text = ""
        lblCnorAddr2.Text = ""
        lblCnorAddr3.Text = ""
        lblCnorAddr4.Text = ""
        lblCnorAddr5.Text = ""
        lblCneeAddr1.Text = ""
        lblCneeAddr2.Text = ""
        lblCneeAddr3.Text = ""
        lblCneeAddr4.Text = ""
        lblCneeAddr5.Text = ""
        lblWeight.Text = ""
        lblNOP.Text = ""
        lblSpclInstructions.Text = ""
        lblValForCustoms.Text = ""
        lblContents.Text = ""
    End Sub

</script>
<html>
<head>
<title>
Print Consignment Note
</title>
</head>
<body>
    <form runat="server">

        <asp:Table id="tabPlainPaperAWB1" runat="server" Width="620px" Font-Names="Verdana" Font-Size="XX-Small">
            <asp:TableRow>
                <asp:TableCell VerticalAlign="Top" Wrap="False" Width="400px">
                    <asp:Label runat="server" font-size="Large" font-bold="True">Consignment</asp:Label>&nbsp;&nbsp;&nbsp;<asp:Label runat="server" id="lblConsignmentNumber" font-size="Large" font-bold="True"></asp:Label>
                </asp:TableCell>
                <asp:TableCell HorizontalAlign="Right" Wrap="False" Width="220px" VerticalAlign="Top">
                    <barcode:Code128 runat="server" Height="50px" ID="Barcode"></barcode:Code128>
                </asp:TableCell>
            </asp:TableRow>
        </asp:Table>
        <asp:Table id="tabPlainPaperAWB2" runat="server" Width="620px" Font-Names="Verdana" Font-Size="XX-Small">
            <asp:TableRow>
                <asp:TableCell Width="425px">
                    <asp:Label runat="server" font-size="XX-Small" font-names="Verdana" font-underline="True">Attach a copy of this consignment to each parcel</asp:Label>
                </asp:TableCell>
                <asp:TableCell Wrap="False" Width="195px">

                </asp:TableCell>
            </asp:TableRow>
        </asp:Table>
        <asp:Table id="tabPlainPaperAWB3" runat="server" Width="620px" Font-Names="Verdana" Font-Size="XX-Small">
            <asp:TableRow>
                <asp:TableCell Width="325px">
                    <br/>
                    <asp:Label runat="server" font-bold="True">Transworld</asp:Label>
                </asp:TableCell>
                <asp:TableCell Width="100px" HorizontalAlign="Right">
                    <br/>
                    <asp:Label runat="server">Dated:</asp:Label>&nbsp;&nbsp;
                </asp:TableCell>
                <asp:TableCell Width="195px">
                    <br/>
                    <asp:Label runat="server" id="lblDate1" font-bold="True"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell>
                    <asp:Label runat="server">The Mercury Centre, Central Way</asp:Label>
                </asp:TableCell>
                <asp:TableCell HorizontalAlign="Right">
                    <asp:Label runat="server">Timed:</asp:Label>&nbsp;&nbsp;
                </asp:TableCell>
                <asp:TableCell>
                    <asp:Label runat="server" id="lblDate2" font-bold="True"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell>
                    <asp:Label runat="server">Feltham, Middx, TW14 0RN</asp:Label>
                </asp:TableCell>
                <asp:TableCell HorizontalAlign="Right">
                    <asp:Label runat="server">Shipper's Ref 1:</asp:Label>&nbsp;&nbsp;
                </asp:TableCell>
                <asp:TableCell>
                    <asp:Label runat="server" id="lblShippersRef1" font-bold="True"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell>
                    <asp:Label runat="server">UNITED KINGDOM</asp:Label>
                </asp:TableCell>
                <asp:TableCell HorizontalAlign="Right">
                    <asp:Label runat="server">Shipper's Ref 2:</asp:Label>&nbsp;&nbsp;
                </asp:TableCell>
                <asp:TableCell>
                    <asp:Label runat="server" id="lblShippersRef2" font-bold="True"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell>
                    <asp:Label runat="server">T: 44(0)20 8751 1111   F: 44(0)20 8890 9090</asp:Label>
                </asp:TableCell>
                <asp:TableCell HorizontalAlign="Right">
                    <asp:Label runat="server">Pieces:</asp:Label>&nbsp;&nbsp;
                </asp:TableCell>
                <asp:TableCell>
                    <asp:Label runat="server" id="lblNOP" font-bold="True"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell>
                    <asp:Label runat="server">E: account.managers@transworld.eu.com</asp:Label>
                </asp:TableCell>
                <asp:TableCell HorizontalAlign="Right">
                    <asp:Label runat="server">Weight (Kgs):</asp:Label>&nbsp;&nbsp;
                </asp:TableCell>
                <asp:TableCell>
                    <asp:Label runat="server" id="lblWeight" font-bold="True"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell ColumnSpan="3">
                    <br/>
                    <hr />
                </asp:TableCell>
            </asp:TableRow>
        </asp:Table>
        <asp:Table id="tabPlainPaperAWB4" runat="server" Width="620px" Font-Names="Verdana" Font-Size="XX-Small">
            <asp:TableRow>
                <asp:TableCell Width="310px">
                    <asp:Label runat="server" font-bold="True">Consignor</asp:Label>&nbsp;<asp:Label runat="server" id="lblCustomerAccountCode" font-bold="True"></asp:Label>
                    <br/>
                    <br/>
                </asp:TableCell>
                <asp:TableCell Width="310px">
                    <asp:Label runat="server" font-bold="True">Consignee</asp:Label>
                    <br/>
                    <br/>
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell>
                    <asp:Label runat="server" id="lblCnorAddr1"></asp:Label>
                </asp:TableCell>
                <asp:TableCell>
                    <asp:Label runat="server" id="lblCneeAddr1"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell>
                    <asp:Label runat="server" id="lblCnorAddr2"></asp:Label>
                </asp:TableCell>
                <asp:TableCell>
                    <asp:Label runat="server" id="lblCneeAddr2"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell>
                    <asp:Label runat="server" id="lblCnorAddr3"></asp:Label>
                </asp:TableCell>
                <asp:TableCell>
                    <asp:Label runat="server" id="lblCneeAddr3"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell>
                    <asp:Label runat="server" id="lblCnorAddr4"></asp:Label>
                </asp:TableCell>
                <asp:TableCell>
                    <asp:Label runat="server" id="lblCneeAddr4"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell>
                    <asp:Label runat="server" id="lblCnorAddr5"></asp:Label>
                    <br/>
                    <br/>
                </asp:TableCell>
                <asp:TableCell>
                    <asp:Label runat="server" id="lblCneeAddr5"></asp:Label>
                    <br/>
                    <br/>
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell ColumnSpan="2">
                    <hr />
                </asp:TableCell>
            </asp:TableRow>
        </asp:Table>
        <asp:Table id="tabPlainPaperAWB5" runat="server" Width="620px" Font-Names="Verdana" Font-Size="XX-Small">
            <asp:TableRow>
                <asp:TableCell Width="175px">
                    <asp:Label runat="server">Contents:</asp:Label>
                </asp:TableCell>
                <asp:TableCell Width="445px">
                    <asp:Label runat="server" id="lblContents" font-bold="True"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell>
                    <asp:Label runat="server">Customs Value:</asp:Label>
                </asp:TableCell>
                <asp:TableCell>
                    <asp:Label runat="server" id="lblValForCustoms" font-bold="True"></asp:Label>
                    &nbsp;<asp:Label runat="server" id="lblValForCustomsCurrency" font-bold="True"></asp:Label>                    
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell>
                    <asp:Label runat="server">Insurance Value:</asp:Label>
                </asp:TableCell>
                <asp:TableCell>
                    <asp:Label runat="server" id="lblValForInsurance" font-bold="True"></asp:Label>
                    &nbsp;<asp:Label runat="server" id="lblValForInsuranceCurrency" font-bold="True"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell>
                    <asp:Label runat="server">Special Instructions:</asp:Label>
                </asp:TableCell>
                <asp:TableCell>
                    <asp:Label runat="server" id="lblSpclInstructions" font-bold="True"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>
        </asp:Table>
        <asp:Table id="tabPlainPaperAWB6" runat="server" Width="620px" Font-Names="Verdana" Font-Size="XX-Small">
            <asp:TableRow>
                <asp:TableCell Width="420px">
                    <br/>
                    <br/>
                    <br/>
                    <asp:Label runat="server" ForeColor="gray">Powered by courier software (www.couriersoftware.co.uk)</asp:Label>
                </asp:TableCell>
                <asp:TableCell Width="200px" HorizontalAlign="Right">
                    <br/>
                    <br/>
                    <br/>
                    <asp:Label runat="server">Printed on</asp:Label>&nbsp;<asp:Label runat="server" id="lblPrinted1"></asp:Label>&nbsp;<asp:Label runat="server">at</asp:Label>&nbsp;<asp:Label runat="server" id="lblPrinted2"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>
        </asp:Table>
        <asp:Table id="tabPlainPaperAWB7" runat="server" Width="620px" Font-Names="Verdana" Font-Size="XX-Small">
            <asp:TableRow>
                <asp:TableCell Width="310px" HorizontalAlign="Right">
                    <br/>
                    <br/>
                    <br/>
                    <asp:Label runat="server">- - - -    fold here    - - - - - - - - - - - - - - - - - - - - - - - -</asp:Label>
                </asp:TableCell>
                <asp:TableCell Width="310px">
                    <br/>
                    <br/>
                    <br/>
                    <asp:Label runat="server">- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - </asp:Label>
                </asp:TableCell>
            </asp:TableRow>
        </asp:Table>

        <asp:Label id="lblError" runat="server" font-size="X-Small" font-names="Verdana" forecolor="#00C000"></asp:Label>
    </form>
</body>
</html>