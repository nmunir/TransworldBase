<%@ Page Language="VB" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>
<script runat="server">

    Sub Page_Load(Source As Object, E As EventArgs)
    
        If Not IsPostBack Then
    
    
            If Not IsNumeric(Session("CustomerKey")) Then
    
                Server.Transfer("session_expired.aspx")
    
            ElseIf Session("CustomerKey") = 0 Then
    
                'We shouldn't be here without a valid CustomerKey
                Server.Transfer("error.aspx")
    
            End If
    
            ViewState("Referrer") = Request.UrlReferrer
    
            'Get Product detail
            Dim oDataReader As SqlDataReader
            Dim sConn As String = ConfigurationSettings.AppSettings("ConnectionString")
            Dim oConn As New SqlConnection(sConn)
            Dim oCmd As New SqlCommand("spASPNET_Product_GetProductFromKey",oConn)
    
            oCmd.CommandType = CommandType.StoredProcedure
            Dim oParam As SqlParameter = oCmd.Parameters.Add("@ProductKey",SqlDbType.Int,4)
            oParam.Value = CLng(Request.QueryString("ProductKey"))
    
            Try
    
                oConn.Open()
                oDataReader = oCmd.ExecuteReader()
                oDataReader.Read()
    
                If IsDBNull(oDataReader ("ProductCode")) Then
                    lblProductCode.Text = ""
                Else
                    lblProductCode.Text = oDataReader ("ProductCode")
                End If
    
                If IsDBNull(oDataReader ("LanguageId")) Then
                    lblLanguage.Text = ""
                Else
                    lblLanguage.Text = oDataReader ("LanguageId")
                End If
    
                If IsDBNull(oDataReader ("ProductDate")) Then
                    lblProductDate.Text = ""
                Else
                    lblProductDate.Text = oDataReader ("ProductDate")
                End If
    
                If IsDBNull(oDataReader ("ItemsPerBox")) Then
                    lblItemsPerBox.Text = ""
                Else
                    lblItemsPerBox.Text = oDataReader ("ItemsPerBox")
                End If
    
                If IsDBNull(oDataReader ("ProductDepartmentId")) Then
                    lblDepartment.Text = ""
                Else
                    lblDepartment.Text = oDataReader ("ProductDepartmentId")
                End If
    
                If IsDBNull(oDataReader ("MinimumStockLevel")) Then
                    lblMinStockLevel.Text = ""
                Else
                    lblMinStockLevel.Text = oDataReader ("MinimumStockLevel")
                End If
    
                If IsDBNull(oDataReader ("ProductDescription")) Then
                    lblDescription.Text = ""
                Else
                    lblDescription.Text = oDataReader ("ProductDescription")
                End If
    
                If IsDBNull(oDataReader ("UnitValue")) Then
                    lblUnitValue.Text = ""
                Else
                    lblUnitValue.Text = oDataReader ("UnitValue")
                End If
    
                If IsDBNull(oDataReader ("SerialNumbersFlag")) Then
                    chkProspectusNumbers.Checked = "False"
                ElseIf oDataReader ("SerialNumbersFlag") = "Y" Then
                    chkProspectusNumbers.Checked = "True"
                ElseIf oDataReader ("SerialNumbersFlag") = "N" Then
                    chkProspectusNumbers.Checked = "False"
                End If
    
                If IsDBNull(oDataReader ("UnitWeightGrams")) Then
                    lblUnitWeight.Text = ""
                Else
                    lblUnitWeight.Text = oDataReader ("UnitWeightGrams")
                End If
    
    
            Catch ex As SqlException
    
                lblError.Text = ex.ToString
    
            End Try
    
            oConn.Close()
    
        End If
    
    End Sub
    
    
    Sub GoBack_Click(sender As Object, e As EventArgs)
    
        Dim sLastUrl As String = ViewState("Referrer").ToString()
        Response.Redirect(sLastUrl)
    
    
    End Sub

</script>
<html>
<head>
</head>
<body style="FONT-FAMILY: arial">
    <h2>
    </h2>
    <form runat="server">
        <p>
            <main:Header id="ctlHeader" runat="server"></main:Header>
            <br />
            <asp:Table id="Table2" runat="server" Width="100%">
                <asp:TableRow>
                    <asp:TableCell Width="23%">
                        <asp:Image id="Image1" runat="server" ImageUrl="./images/icon_property.gif"></asp:Image>
                        &nbsp; 
                        <asp:Label runat="server" ID="Label1" Font-Size="X-Small" Font-Names="Arial">Product Detail</asp:Label>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow>
                    <asp:TableCell Text="
                        &lt;hr /&gt;
                    "></asp:TableCell>
                </asp:TableRow>
            </asp:Table>
            <br />
        </p>
        <p>
            <asp:Table id="Table1" runat="server" Width="100%" Font-Names="Arial" Font-Size="X-Small">
                <asp:TableRow>
                    <asp:TableCell BackColor="LightBlue" Width="25%">
                        &nbsp;
                        <asp:Label runat="server" ForeColor="#0000C0" ID="Label2">Product Code:</asp:Label>
                    </asp:TableCell>
                    <asp:TableCell BackColor="LightGray" Width="25%">
                        &nbsp;
                        <asp:Label runat="server" ID="lblProductCode"></asp:Label>
                    </asp:TableCell>
                    <asp:TableCell BackColor="LightBlue" Width="25%">
                        &nbsp;
                        <asp:Label runat="server" ForeColor="#0000C0" ID="Label3">Language:</asp:Label>
                    </asp:TableCell>
                    <asp:TableCell BackColor="LightGray" Width="25%">
                        &nbsp;
                        <asp:Label runat="server" ID="lblLanguage"></asp:Label>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow>
                    <asp:TableCell BackColor="LightBlue">
                        &nbsp;
                        <asp:Label runat="server" ForeColor="#0000C0" ID="Label4">Product Date:</asp:Label>
                    </asp:TableCell>
                    <asp:TableCell BackColor="LightGray">
                        &nbsp;
                        <asp:Label runat="server" ID="lblProductDate"></asp:Label>
                    </asp:TableCell>
                    <asp:TableCell BackColor="LightBlue">
                        &nbsp;
                        <asp:Label runat="server" ForeColor="#0000C0" ID="Label5">Items per box:</asp:Label>
                    </asp:TableCell>
                    <asp:TableCell BackColor="LightGray">
                        &nbsp;
                        <asp:Label runat="server" ID="lblItemsPerBox"></asp:Label>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow>
                    <asp:TableCell BackColor="LightBlue">
                        &nbsp;
                        <asp:Label runat="server" ForeColor="#0000C0" ID="Label6">Department:</asp:Label>
                    </asp:TableCell>
                    <asp:TableCell BackColor="LightGray">
                        &nbsp;
                        <asp:Label runat="server" ID="lblDepartment"></asp:Label>
                    </asp:TableCell>
                    <asp:TableCell BackColor="LightBlue">
                        &nbsp;
                        <asp:Label runat="server" ForeColor="#0000C0" ID="Label7">Product Code:</asp:Label>
                    </asp:TableCell>
                    <asp:TableCell BackColor="LightGray">
                        &nbsp;
                        <asp:Label runat="server" ID="lblMinStockLevel"></asp:Label>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow>
                    <asp:TableCell BackColor="LightBlue">
                        &nbsp;
                        <asp:Label runat="server" ForeColor="#0000C0" ID="Label8">Description:</asp:Label>
                    </asp:TableCell>
                    <asp:TableCell BackColor="LightGray" ColumnSpan="3">
                        &nbsp;
                        <asp:Label runat="server" ID="lblDescription"></asp:Label>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow>
                    <asp:TableCell BackColor="LightBlue">
                        &nbsp;
                        <asp:Label runat="server" ForeColor="#0000C0" ID="Label9">Unit Value:</asp:Label>
                    </asp:TableCell>
                    <asp:TableCell BackColor="LightGray">
                        &nbsp;
                        <asp:Label runat="server" ID="lblUnitValue"></asp:Label>
                    </asp:TableCell>
                    <asp:TableCell BackColor="LightBlue">
                        &nbsp;
                        <asp:Label runat="server" ForeColor="#0000C0" ID="Label10">Prospectus Numbers:</asp:Label>
                    </asp:TableCell>
                    <asp:TableCell BackColor="LightGray">
                        &nbsp;
                        <asp:CheckBox runat="server" ID="chkProspectusNumbers"></asp:CheckBox>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow>
                    <asp:TableCell BackColor="LightBlue">
                        &nbsp;
                        <asp:Label runat="server" ForeColor="#0000C0" ID="Label11">Unit Weight:</asp:Label>
                    </asp:TableCell>
                    <asp:TableCell BackColor="LightGray">
                        &nbsp;
                        <asp:Label runat="server" ID="lblUnitWeight"></asp:Label>
                    </asp:TableCell>
                    <asp:TableCell></asp:TableCell>
                    <asp:TableCell></asp:TableCell>
                </asp:TableRow>
            </asp:Table>
        </p>
        <p align="right">
            <asp:Label id="Label12" runat="server" Font-Names="Verdana" Font-Size="XX-Small" ForeColor="Green">use your browser's back button to go back</asp:Label>
        </p>
        <p align="left">
            <asp:Label id="lblError" runat="server" Font-Names="Arial" Font-Size="X-Small" ForeColor="Red"></asp:Label>
        </p>
        <p>
        </p>
        <p>
        </p>
    </form>
</body>
</html>
