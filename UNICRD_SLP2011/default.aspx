<%@ Page Language="VB" Theme="AIMSDefault"  %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>

<script runat="server">

    Dim sSendEmailTo As String = "marilyn.quinn@transworld.eu.com"
    Dim sSendWarningEmailTo As String = "marilyn.quinn@transworld.eu.com"
    Dim sSentFromEmail As String = "automailer@transworld.eu.com"
    Dim sSubject As String = "Web form order: USLP Progress Report 2011"
    Dim sWarningSubject As String = "WARNING: Order for > 2000 copies of Unilever USLP Progress Report 2011"
    Dim sProduct As String = "USLP Progress Report 2011 - Product code - SUS LIVING PR 2011"
    Dim lCustomerKey As Long = 43
    
    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        txtFirstName.Focus()
    End Sub
    
    Protected Sub btnReset_Click(ByVal sender As Object, ByVal e As EventArgs)
        txtFirstName.Text = ""
        txtLastName.Text = ""
        txtJobTitle.Text = ""
        txtDepartment.Text = ""
        txtStreet.Text = ""
        txtPostcode.Text = ""
        txtTown.Text = ""
        ddlCountry.SelectedIndex = 0
        txtEmail.Text = ""
        txtTelephone.Text = ""
        txtQuantity.Text = ""
        txtComments.Text = ""
    End Sub
    
    Protected Sub btnSubmit_Click(ByVal sender As Object, ByVal e As EventArgs)
        Call StoreRequest()
        Call SubmitRequest()
        Response.Redirect("thank_you.aspx")
    End Sub
    
    Protected Sub StoreRequest()
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Adhoc_Fulfilment_AddEntry", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
    
        Try
            oCmd.Parameters.Add(New SqlParameter("@RecordType", SqlDbType.VarChar, 50))
            oCmd.Parameters("@RecordType").Value = "5LEVERS"

            oCmd.Parameters.Add(New SqlParameter("@FirstName", SqlDbType.VarChar, 50))
            oCmd.Parameters("@FirstName").Value = txtFirstName.Text

            oCmd.Parameters.Add(New SqlParameter("@LastName", SqlDbType.VarChar, 50))
            oCmd.Parameters("@LastName").Value = txtLastName.Text

            oCmd.Parameters.Add(New SqlParameter("@JobTitle", SqlDbType.VarChar, 50))
            oCmd.Parameters("@JobTitle").Value = txtJobTitle.Text

            oCmd.Parameters.Add(New SqlParameter("@Department", SqlDbType.VarChar, 50))
            oCmd.Parameters("@Department").Value = txtDepartment.Text

            oCmd.Parameters.Add(New SqlParameter("@Company", SqlDbType.VarChar, 50))
            oCmd.Parameters("@Company").Value = txtCompany.Text

            oCmd.Parameters.Add(New SqlParameter("@AddressLine1", SqlDbType.VarChar, 50))
            oCmd.Parameters("@AddressLine1").Value = txtStreet.Text

            oCmd.Parameters.Add(New SqlParameter("@AddressLine2", SqlDbType.VarChar, 50))
            oCmd.Parameters("@AddressLine2").Value = txtStreet2.Text

            oCmd.Parameters.Add(New SqlParameter("@Town", SqlDbType.VarChar, 50))
            oCmd.Parameters("@Town").Value = txtTown.Text

            oCmd.Parameters.Add(New SqlParameter("@State", SqlDbType.VarChar, 50))
            oCmd.Parameters("@State").Value = String.Empty

            oCmd.Parameters.Add(New SqlParameter("@Postcode", SqlDbType.VarChar, 50))
            oCmd.Parameters("@Postcode").Value = txtPostcode.Text

            oCmd.Parameters.Add(New SqlParameter("@Country", SqlDbType.VarChar, 50))
            oCmd.Parameters("@Country").Value = ddlCountry.SelectedItem.Text

            oCmd.Parameters.Add(New SqlParameter("@CountryCode", SqlDbType.Int))
            oCmd.Parameters("@CountryCode").Value = 0

            oCmd.Parameters.Add(New SqlParameter("@EmailAddr", SqlDbType.VarChar, 50))
            oCmd.Parameters("@EmailAddr").Value = txtEmail.Text

            oCmd.Parameters.Add(New SqlParameter("@Telephone", SqlDbType.VarChar, 50))
            oCmd.Parameters("@Telephone").Value = txtTelephone.Text

            oCmd.Parameters.Add(New SqlParameter("@Quantity1", SqlDbType.Int))
            oCmd.Parameters("@Quantity1").Value = CInt(txtQuantity.Text)

            oCmd.Parameters.Add(New SqlParameter("@Quantity2", SqlDbType.Int))
            oCmd.Parameters("@Quantity2").Value = 0

            oCmd.Parameters.Add(New SqlParameter("@Comments", SqlDbType.VarChar, 200))
            oCmd.Parameters("@Comments").Value = txtComments.Text

            oCmd.Parameters.Add(New SqlParameter("@CustomStr1", SqlDbType.VarChar, 50))
            oCmd.Parameters("@CustomStr1").Value = String.Empty

            oCmd.Parameters.Add(New SqlParameter("@CustomStr2", SqlDbType.VarChar, 50))
            oCmd.Parameters("@CustomStr2").Value = String.Empty

            oCmd.Parameters.Add(New SqlParameter("@CustomInt1", SqlDbType.Int))
            oCmd.Parameters("@CustomInt1").Value = 0

            oCmd.Parameters.Add(New SqlParameter("@CustomInt2", SqlDbType.Int))
            oCmd.Parameters("@CustomInt2").Value = 0

            oCmd.Parameters.Add(New SqlParameter("@CustomBit1", SqlDbType.Bit))
            oCmd.Parameters("@CustomBit1").Value = 0

            oCmd.Parameters.Add(New SqlParameter("@CustomBit2", SqlDbType.Bit))
            oCmd.Parameters("@CustomBit2").Value = 0

            oConn.Open()
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            WebMsgBox.Show("StoreRequest: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub SubmitRequest()
        If IsValid Then
            Dim sFirstName As String
            Dim sLastName As String
            Dim sJobTitle As String
            Dim sDepartment As String
            Dim sAddr1 As String
            Dim sAddr2 As String
            Dim sTown As String
            Dim sZipCOde As String
            Dim cCountry As String
            Dim sEmail As String
            Dim sTelephone As String
            Dim sQuantity As String
            Dim sComments As String
    
            Dim sBody As String = ""
            Dim dteNow As DateTime
    
            dteNow = DateTime.Now()
    
            sFirstName = txtFirstName.Text
            sLastName = txtLastName.Text.ToString
            sJobTitle = txtJobTitle.Text
            sDepartment = txtDepartment.Text
            sAddr1 = txtStreet.Text
            sAddr2 = txtStreet2.Text
            sTown = txtTown.Text
            sZipCOde = txtPostcode.Text
            cCountry = ddlCountry.SelectedItem.Text
            sEmail = txtEmail.Text
            sTelephone = txtTelephone.Text
            sQuantity = txtQuantity.Text
            sComments = txtComments.Text
    
            sBody &= "<br />" & vbNewLine
            sBody &= "Ordered on:           " & dteNow.ToString("F") & "<br />" & vbNewLine
            sBody &= "<br />" & vbNewLine
            sBody &= "First Name:           " & sFirstName & "<br />" & vbNewLine
            sBody &= "Last Name:            " & sLastName & "<br />" & vbNewLine
            sBody &= "Job Title:            " & sJobTitle & "<br />" & vbNewLine
            sBody &= "Department:           " & sDepartment & "<br />" & vbNewLine
            sBody &= "Addr 1:               " & sAddr1 & "<br />" & vbNewLine
            sBody &= "Addr 2:               " & sAddr2 & "<br />" & vbNewLine
            sBody &= "Town/City:            " & sTown & "<br />" & vbNewLine
            sBody &= "Zip Code:             " & sZipCOde & "<br />" & vbNewLine
            sBody &= "Country:              " & cCountry & "<br />" & vbNewLine
            sBody &= "Email:                " & sEmail & "<br />" & vbNewLine
            sBody &= "Telephone:            " & sTelephone & "<br />" & vbNewLine
            sBody &= "Brochure Quantity:    " & sQuantity & "<br />" & vbNewLine
            sBody &= "Comments:             " & sComments & "<br />" & vbNewLine
            sBody &= "<br />" & vbNewLine
            sBody &= "<br />" & vbNewLine
            sBody &= "Product:              " & sProduct & "<br />" & vbNewLine
    
            Call SendHTMLEmail("UNI_SLP_11_REQUEST", lCustomerKey, 0, 0, 0, sSendEmailTo, sSubject, sBody, sBody, 0)
            If CInt(txtQuantity.Text) > 2000 Then
                Call SendHTMLEmail("UNI_SLP_11_WARNING", lCustomerKey, 0, 0, 0, sSendWarningEmailTo, sWarningSubject, sBody, sBody, 0)
            End If
        End If
    End Sub

    Protected Sub SendHTMLEmail(ByVal sType As String, ByVal nCustomerKey As Integer, ByVal nStockBookingKey As Integer, ByVal nConsignmentKey As Integer, ByVal nProductKey As Integer, ByVal sRecipient As String, ByVal sSubject As String, ByVal sBodyText As String, ByVal sBodyHTML As String, ByVal nQueuedBy As Integer)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Email_AddToQueue", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
    
        Try
            oCmd.Parameters.Add(New SqlParameter("@MessageId", SqlDbType.NVarChar, 20))
            oCmd.Parameters("@MessageId").Value = sType
    
            oCmd.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int, 4))
            oCmd.Parameters("@CustomerKey").Value = nCustomerKey
    
            oCmd.Parameters.Add(New SqlParameter("@StockBookingKey", SqlDbType.Int, 4))
            oCmd.Parameters("@StockBookingKey").Value = nStockBookingKey
    
            oCmd.Parameters.Add(New SqlParameter("@ConsignmentKey", SqlDbType.Int, 4))
            oCmd.Parameters("@ConsignmentKey").Value = nConsignmentKey
    
            oCmd.Parameters.Add(New SqlParameter("@ProductKey", SqlDbType.Int, 4))
            oCmd.Parameters("@ProductKey").Value = nProductKey
    
            oCmd.Parameters.Add(New SqlParameter("@To", SqlDbType.NVarChar, 100))
            oCmd.Parameters("@To").Value = sRecipient
    
            oCmd.Parameters.Add(New SqlParameter("@Subject", SqlDbType.NVarChar, 60))
            oCmd.Parameters("@Subject").Value = sSubject
    
            oCmd.Parameters.Add(New SqlParameter("@BodyText", SqlDbType.NText))
            oCmd.Parameters("@BodyText").Value = sBodyText
    
            oCmd.Parameters.Add(New SqlParameter("@BodyHTML", SqlDbType.NText))
            oCmd.Parameters("@BodyHTML").Value = sBodyHTML
    
            oCmd.Parameters.Add(New SqlParameter("@QueuedBy", SqlDbType.Int, 4))
            oCmd.Parameters("@QueuedBy").Value = nQueuedBy
    
            oConn.Open()
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            WebMsgBox.Show("SendHTMLEmail: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
    
</script>
<html xmlns="http://www.w3.org/1999/xhtml" >
<head>
<title>Unilever Sustainable Living Plan Progress Report 2011</title>
    <style type="text/css">
        .style1
        {
            width: 425px;
            height: 605px;
        }
    </style>
</head>
<body>
    <form id="SDO2008" runat="server">
      <asp:Label id="Label1" runat="server" font-names="Tahoma" 
          forecolor="LightSeaGreen" font-size="Large" 
          Text="Unilever Sustainable Living Plan Progress Report 2011"/>
      <table style="width: 100%; font-size: xx-small; font-family: Verdana;">
            <tr>
                <td style="width: 1%">
                </td>
                <td style="width: 25%">
                </td>
                <td style="width: 24%">
                </td>
                <td style="width: 25%">
                </td>
                <td style="width: 24%">
                </td>
                <td style="width: 1%">
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td style="font-size: x-small; font-family: Arial" colspan="4">
                    Available from April 2012. To order copies please enter your name, full address, job title and department.<br />
                    <br />
                    Please indicate in the comments box below the reason for your order - this helps
                    us judge how the report is being used - and if you would prefer an express courier
                    or mail delivery service. Please refer orders for more than 2,000 copies to
                    Marilyn Quinn (+44) (0)208 751 1111 email: marilyn.quinn@transworld.eu.com<br />
                    <br /><br />
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td>
                </td>
                <td>
                </td>
                <td>
                </td>
                <td>
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td rowspan="14" align="center" valign="middle">
                    <img alt="" class="style1" src="images/USLP2011_internal.jpg" /></td>
                <td align="right" style="color: red">
                    First Name:</td>
                <td>
                        <asp:TextBox runat="server" ID="txtFirstName" Font-Size="XX-Small" MaxLength="50" Width="200px"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="rfvFirstname" runat="server" ControlToValidate="txtFirstName" Font-Names="Arial"> required!</asp:RequiredFieldValidator>
                </td>
                <td>
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td style="height: 23px">
                </td>
                <td align="right" style="color: red; height: 23px;">
                    Last Name:</td>
                <td style="height: 23px">
                        <asp:TextBox runat="server" ID="txtLastName" Font-Size="XX-Small" MaxLength="50" Width="200px"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="rfvLastname" runat="server" ControlToValidate="txtLastName" Font-Names="Arial"> required!</asp:RequiredFieldValidator>
                </td>
                <td style="height: 23px">
                </td>
                <td style="height: 23px">
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td align="right" style="color: red">
                    Job Title:</td>
                <td>
                        <asp:TextBox runat="server" ID="txtJobTitle" Font-Size="XX-Small" MaxLength="50" Width="200px"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="RequiredFieldValidator3" runat="server" ControlToValidate="txtJobTitle" Font-Names="Arial"> required!</asp:RequiredFieldValidator>
                </td>
                <td>
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td align="right" style="color: red">
                    Department:</td>
                <td>
                        <asp:TextBox runat="server" ID="txtDepartment" Font-Size="XX-Small" MaxLength="50" Width="200px"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="rfvDepartment" runat="server" ControlToValidate="txtDepartment" Font-Names="Arial"> required!</asp:RequiredFieldValidator>
                </td>
                <td>
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td align="right" style="color: red">
                    Company:
                </td>
                <td>
                        <asp:TextBox runat="server" ID="txtCompany" Font-Size="XX-Small" MaxLength="50" Width="200px"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="rfvCompany" runat="server" ControlToValidate="txtCompany" Font-Names="Arial"> required!</asp:RequiredFieldValidator>
                </td>
                <td>
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td align="right" style="color: red">
                    Address Line 1:</td>
                <td>
                        <asp:TextBox runat="server" ID="txtStreet" Font-Size="XX-Small" MaxLength="50" Width="200px"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="rfvAddrLine1" runat="server" ControlToValidate="txtStreet" Font-Names="Arial"> required!</asp:RequiredFieldValidator>
                </td>
                <td>
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td align="right">
                    Address Line 2:
                </td>
                <td>
                        <asp:TextBox runat="server" ID="txtStreet2" Font-Size="XX-Small" MaxLength="50" Width="200px"></asp:TextBox>
                </td>
                <td>
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td align="right" style="color: red">
                    Town/City:</td>
                <td>
                        <asp:TextBox runat="server" ID="txtTown" Font-Size="XX-Small" MaxLength="50" Width="200px"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="rfvTown" runat="server" ControlToValidate="txtTown" Font-Names="Arial"> required!</asp:RequiredFieldValidator>
                </td>
                <td>
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td align="right" style="color: red">
                    Post Code / Zip Code:</td>
                <td>
                        <asp:TextBox runat="server" ID="txtPostcode" Font-Size="XX-Small" MaxLength="50" Width="200px"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="rfvPostcode" runat="server" ControlToValidate="txtPostcode" Font-Names="Arial"> required!</asp:RequiredFieldValidator>
                </td>
                <td>
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td align="right" style="color: red">
                    Country:</td>
                <td colspan="2">
                        <asp:DropDownList runat="server" ID="ddlCountry" Font-Size="XX-Small">
                            <asp:ListItem Value="-- please select --">-- please select --</asp:ListItem>
                            <asp:ListItem Value="AFGHANISTAN">AFGHANISTAN</asp:ListItem>
                            <asp:ListItem Value="ALBANIA">ALBANIA</asp:ListItem>
                            <asp:ListItem Value="ALGERIA">ALGERIA</asp:ListItem>
                            <asp:ListItem Value="AMERICAN SAMOA">AMERICAN SAMOA</asp:ListItem>
                            <asp:ListItem Value="ANDORRA">ANDORRA</asp:ListItem>
                            <asp:ListItem Value="ANGOLA">ANGOLA</asp:ListItem>
                            <asp:ListItem Value="ANGUILLA">ANGUILLA</asp:ListItem>
                            <asp:ListItem Value="ANTARCTICA">ANTARCTICA</asp:ListItem>
                            <asp:ListItem Value="ANTIGUA AND BARBUDA">ANTIGUA AND BARBUDA</asp:ListItem>
                            <asp:ListItem Value="ARGENTINA">ARGENTINA</asp:ListItem>
                            <asp:ListItem Value="ARMENIA">ARMENIA</asp:ListItem>
                            <asp:ListItem Value="ARUBA">ARUBA</asp:ListItem>
                            <asp:ListItem Value="AUSTRALIA">AUSTRALIA</asp:ListItem>
                            <asp:ListItem Value="AUSTRIA">AUSTRIA</asp:ListItem>
                            <asp:ListItem Value="AZERBAIJAN">AZERBAIJAN</asp:ListItem>
                            <asp:ListItem Value="BAHAMAS">BAHAMAS</asp:ListItem>
                            <asp:ListItem Value="BAHRAIN">BAHRAIN</asp:ListItem>
                            <asp:ListItem Value="BANGLADESH">BANGLADESH</asp:ListItem>
                            <asp:ListItem Value="BARBADOS">BARBADOS</asp:ListItem>
                            <asp:ListItem Value="BELARUS">BELARUS</asp:ListItem>
                            <asp:ListItem Value="BELGIUM">BELGIUM</asp:ListItem>
                            <asp:ListItem Value="BELIZE">BELIZE</asp:ListItem>
                            <asp:ListItem Value="BENIN">BENIN</asp:ListItem>
                            <asp:ListItem Value="BERMUDA">BERMUDA</asp:ListItem>
                            <asp:ListItem Value="BHUTAN">BHUTAN</asp:ListItem>
                            <asp:ListItem Value="BOLIVIA">BOLIVIA</asp:ListItem>
                            <asp:ListItem Value="BOSNIA AND HERZEGOWINA">BOSNIA AND HERZEGOWINA</asp:ListItem>
                            <asp:ListItem Value="BOTSWANA">BOTSWANA</asp:ListItem>
                            <asp:ListItem Value="BOUVET ISLAND">BOUVET ISLAND</asp:ListItem>
                            <asp:ListItem Value="BRAZIL">BRAZIL</asp:ListItem>
                            <asp:ListItem Value="BRITISH INDIAN OCEAN TERRITORY">BRITISH INDIAN OCEAN TERRITORY</asp:ListItem>
                            <asp:ListItem Value="BRUNEI DARUSSALAM">BRUNEI DARUSSALAM</asp:ListItem>
                            <asp:ListItem Value="BULGARIA">BULGARIA</asp:ListItem>
                            <asp:ListItem Value="BURKINA FASO">BURKINA FASO</asp:ListItem>
                            <asp:ListItem Value="BURUNDI">BURUNDI</asp:ListItem>
                            <asp:ListItem Value="CAMBODIA">CAMBODIA</asp:ListItem>
                            <asp:ListItem Value="CAMEROON">CAMEROON</asp:ListItem>
                            <asp:ListItem Value="CANADA">CANADA</asp:ListItem>
                            <asp:ListItem Value="CAPE VERDE">CAPE VERDE</asp:ListItem>
                            <asp:ListItem Value="CAYMAN ISLANDS">CAYMAN ISLANDS</asp:ListItem>
                            <asp:ListItem Value="CENTRAL AFRICAN REPUBLIC">CENTRAL AFRICAN REPUBLIC</asp:ListItem>
                            <asp:ListItem Value="CHAD">CHAD</asp:ListItem>
                            <asp:ListItem Value="CHILE">CHILE</asp:ListItem>
                            <asp:ListItem Value="CHINA">CHINA</asp:ListItem>
                            <asp:ListItem Value="CHRISTMAS ISLAND">CHRISTMAS ISLAND</asp:ListItem>
                            <asp:ListItem Value="COCOS (KEELING) ISLANDS">COCOS (KEELING) ISLANDS</asp:ListItem>
                            <asp:ListItem Value="COLOMBIA">COLOMBIA</asp:ListItem>
                            <asp:ListItem Value="COMOROS">COMOROS</asp:ListItem>
                            <asp:ListItem Value="CONGO">CONGO</asp:ListItem>
                            <asp:ListItem Value="COOK ISLANDS">COOK ISLANDS</asp:ListItem>
                            <asp:ListItem Value="COSTA RICA">COSTA RICA</asp:ListItem>
                            <asp:ListItem Value="COTE D'IVOIRE">COTE D'IVOIRE</asp:ListItem>
                            <asp:ListItem Value="CROATIA">CROATIA</asp:ListItem>
                            <asp:ListItem Value="CUBA">CUBA</asp:ListItem>
                            <asp:ListItem Value="CYPRUS">CYPRUS</asp:ListItem>
                            <asp:ListItem Value="CZECH REPUBLIC">CZECH REPUBLIC</asp:ListItem>
                            <asp:ListItem Value="DENMARK">DENMARK</asp:ListItem>
                            <asp:ListItem Value="DJIBOUTI">DJIBOUTI</asp:ListItem>
                            <asp:ListItem Value="DOMINICA">DOMINICA</asp:ListItem>
                            <asp:ListItem Value="DOMINICAN REPUBLIC">DOMINICAN REPUBLIC</asp:ListItem>
                            <asp:ListItem Value="EAST TIMOR">EAST TIMOR</asp:ListItem>
                            <asp:ListItem Value="ECUADOR">ECUADOR</asp:ListItem>
                            <asp:ListItem Value="EGYPT">EGYPT</asp:ListItem>
                            <asp:ListItem Value="EL SALVADOR">EL SALVADOR</asp:ListItem>
                            <asp:ListItem Value="EQUATORIAL GUINEA">EQUATORIAL GUINEA</asp:ListItem>
                            <asp:ListItem Value="ERITREA">ERITREA</asp:ListItem>
                            <asp:ListItem Value="ESTONIA">ESTONIA</asp:ListItem>
                            <asp:ListItem Value="ETHIOPIA">ETHIOPIA</asp:ListItem>
                            <asp:ListItem Value="FALKLAND ISLANDS (MALVINAS)">FALKLAND ISLANDS (MALVINAS)</asp:ListItem>
                            <asp:ListItem Value="FAROE ISLANDS">FAROE ISLANDS</asp:ListItem>
                            <asp:ListItem Value="FIJI">FIJI</asp:ListItem>
                            <asp:ListItem Value="FINLAND">FINLAND</asp:ListItem>
                            <asp:ListItem Value="FRANCE">FRANCE</asp:ListItem>
                            <asp:ListItem Value="FRANCE, METROPOLITAN">FRANCE, METROPOLITAN</asp:ListItem>
                            <asp:ListItem Value="FRENCH GUIANA">FRENCH GUIANA</asp:ListItem>
                            <asp:ListItem Value="FRENCH POLYNESIA">FRENCH POLYNESIA</asp:ListItem>
                            <asp:ListItem Value="FRENCH SOUTHERN TERRITORIES">FRENCH SOUTHERN TERRITORIES</asp:ListItem>
                            <asp:ListItem Value="GABON">GABON</asp:ListItem>
                            <asp:ListItem Value="GAMBIA">GAMBIA</asp:ListItem>
                            <asp:ListItem Value="GEORGIA">GEORGIA</asp:ListItem>
                            <asp:ListItem Value="GERMANY">GERMANY</asp:ListItem>
                            <asp:ListItem Value="GHANA">GHANA</asp:ListItem>
                            <asp:ListItem Value="GIBRALTAR">GIBRALTAR</asp:ListItem>
                            <asp:ListItem Value="GREECE">GREECE</asp:ListItem>
                            <asp:ListItem Value="GREENLAND">GREENLAND</asp:ListItem>
                            <asp:ListItem Value="GRENADA">GRENADA</asp:ListItem>
                            <asp:ListItem Value="GUADELOUPE">GUADELOUPE</asp:ListItem>
                            <asp:ListItem Value="GUAM">GUAM</asp:ListItem>
                            <asp:ListItem Value="GUATEMALA">GUATEMALA</asp:ListItem>
                            <asp:ListItem Value="GUINEA">GUINEA</asp:ListItem>
                            <asp:ListItem Value="GUINEA-BISSAU">GUINEA-BISSAU</asp:ListItem>
                            <asp:ListItem Value="GUYANA">GUYANA</asp:ListItem>
                            <asp:ListItem Value="HAITI">HAITI</asp:ListItem>
                            <asp:ListItem Value="HEARD AND MC DONALD ISLANDS">HEARD AND MC DONALD ISLANDS</asp:ListItem>
                            <asp:ListItem Value="HONDURAS">HONDURAS</asp:ListItem>
                            <asp:ListItem Value="HONG KONG">HONG KONG</asp:ListItem>
                            <asp:ListItem Value="HUNGARY">HUNGARY</asp:ListItem>
                            <asp:ListItem Value="ICELAND">ICELAND</asp:ListItem>
                            <asp:ListItem Value="INDIA">INDIA</asp:ListItem>
                            <asp:ListItem Value="INDONESIA">INDONESIA</asp:ListItem>
                            <asp:ListItem Value="IRAN (ISLAMIC REPUBLIC OF)">IRAN (ISLAMIC REPUBLIC OF)</asp:ListItem>
                            <asp:ListItem Value="IRAQ">IRAQ</asp:ListItem>
                            <asp:ListItem Value="IRELAND">IRELAND</asp:ListItem>
                            <asp:ListItem Value="ISRAEL">ISRAEL</asp:ListItem>
                            <asp:ListItem Value="ITALY">ITALY</asp:ListItem>
                            <asp:ListItem Value="JAMAICA">JAMAICA</asp:ListItem>
                            <asp:ListItem Value="JAPAN">JAPAN</asp:ListItem>
                            <asp:ListItem Value="JORDAN">JORDAN</asp:ListItem>
                            <asp:ListItem Value="KAZAKHSTAN">KAZAKHSTAN</asp:ListItem>
                            <asp:ListItem Value="KENYA">KENYA</asp:ListItem>
                            <asp:ListItem Value="KIRIBATI">KIRIBATI</asp:ListItem>
                            <asp:ListItem Value="KOREA, DEMOCRATIC PEOPLE'S REPUBLIC OF">KOREA, DEMOCRATIC PEOPLE'S REPUBLIC OF</asp:ListItem>
                            <asp:ListItem Value="KOREA, REPUBLIC OF">KOREA, REPUBLIC OF</asp:ListItem>
                            <asp:ListItem Value="KUWAIT">KUWAIT</asp:ListItem>
                            <asp:ListItem Value="KYRGYZSTAN">KYRGYZSTAN</asp:ListItem>
                            <asp:ListItem Value="LAO PEOPLE'S DEMOCRATIC REPUBLIC">LAO PEOPLE'S DEMOCRATIC REPUBLIC</asp:ListItem>
                            <asp:ListItem Value="LATVIA">LATVIA</asp:ListItem>
                            <asp:ListItem Value="LEBANON">LEBANON</asp:ListItem>
                            <asp:ListItem Value="LESOTHO">LESOTHO</asp:ListItem>
                            <asp:ListItem Value="LIBERIA">LIBERIA</asp:ListItem>
                            <asp:ListItem Value="LIBYAN ARAB JAMAHIRIYA">LIBYAN ARAB JAMAHIRIYA</asp:ListItem>
                            <asp:ListItem Value="LIECHTENSTEIN">LIECHTENSTEIN</asp:ListItem>
                            <asp:ListItem Value="LITHUANIA">LITHUANIA</asp:ListItem>
                            <asp:ListItem Value="LUXEMBOURG">LUXEMBOURG</asp:ListItem>
                            <asp:ListItem Value="MACAU">MACAU</asp:ListItem>
                            <asp:ListItem Value="MACEDONIA, THE FORMER YUGOSLAV REPUBLIC OF">MACEDONIA, THE FORMER YUGOSLAV REPUBLIC OF</asp:ListItem>
                            <asp:ListItem Value="MADAGASCAR">MADAGASCAR</asp:ListItem>
                            <asp:ListItem Value="MALAWI">MALAWI</asp:ListItem>
                            <asp:ListItem Value="MALAYSIA">MALAYSIA</asp:ListItem>
                            <asp:ListItem Value="MALDIVES">MALDIVES</asp:ListItem>
                            <asp:ListItem Value="MALI">MALI</asp:ListItem>
                            <asp:ListItem Value="MALTA">MALTA</asp:ListItem>
                            <asp:ListItem Value="MARSHALL ISLANDS">MARSHALL ISLANDS</asp:ListItem>
                            <asp:ListItem Value="MARTINIQUE">MARTINIQUE</asp:ListItem>
                            <asp:ListItem Value="MAURITANIA">MAURITANIA</asp:ListItem>
                            <asp:ListItem Value="MAURITIUS">MAURITIUS</asp:ListItem>
                            <asp:ListItem Value="MAYOTTE">MAYOTTE</asp:ListItem>
                            <asp:ListItem Value="MEXICO">MEXICO</asp:ListItem>
                            <asp:ListItem Value="MICRONESIA, FEDERATED STATES OF">MICRONESIA, FEDERATED STATES OF</asp:ListItem>
                            <asp:ListItem Value="MOLDOVA, REPUBLIC OF">MOLDOVA, REPUBLIC OF</asp:ListItem>
                            <asp:ListItem Value="MONACO">MONACO</asp:ListItem>
                            <asp:ListItem Value="MONGOLIA">MONGOLIA</asp:ListItem>
                            <asp:ListItem Value="MONTSERRAT">MONTSERRAT</asp:ListItem>
                            <asp:ListItem Value="MOROCCO">MOROCCO</asp:ListItem>
                            <asp:ListItem Value="MOZAMBIQUE">MOZAMBIQUE</asp:ListItem>
                            <asp:ListItem Value="MYANMAR">MYANMAR</asp:ListItem>
                            <asp:ListItem Value="NAMIBIA">NAMIBIA</asp:ListItem>
                            <asp:ListItem Value="NAURU">NAURU</asp:ListItem>
                            <asp:ListItem Value="NEPAL">NEPAL</asp:ListItem>
                            <asp:ListItem Value="NETHERLANDS">NETHERLANDS</asp:ListItem>
                            <asp:ListItem Value="NETHERLANDS ANTILLES">NETHERLANDS ANTILLES</asp:ListItem>
                            <asp:ListItem Value="NEW CALEDONIA">NEW CALEDONIA</asp:ListItem>
                            <asp:ListItem Value="NEW ZEALAND">NEW ZEALAND</asp:ListItem>
                            <asp:ListItem Value="NICARAGUA">NICARAGUA</asp:ListItem>
                            <asp:ListItem Value="NIGER">NIGER</asp:ListItem>
                            <asp:ListItem Value="NIGERIA">NIGERIA</asp:ListItem>
                            <asp:ListItem Value="NIUE">NIUE</asp:ListItem>
                            <asp:ListItem Value="NORFOLK ISLAND">NORFOLK ISLAND</asp:ListItem>
                            <asp:ListItem Value="NORTHERN MARIANA ISLANDS">NORTHERN MARIANA ISLANDS</asp:ListItem>
                            <asp:ListItem Value="NORWAY">NORWAY</asp:ListItem>
                            <asp:ListItem Value="OMAN">OMAN</asp:ListItem>
                            <asp:ListItem Value="PAKISTAN">PAKISTAN</asp:ListItem>
                            <asp:ListItem Value="PALAU">PALAU</asp:ListItem>
                            <asp:ListItem Value="PANAMA">PANAMA</asp:ListItem>
                            <asp:ListItem Value="PAPUA NEW GUINEA">PAPUA NEW GUINEA</asp:ListItem>
                            <asp:ListItem Value="PARAGUAY">PARAGUAY</asp:ListItem>
                            <asp:ListItem Value="PERU">PERU</asp:ListItem>
                            <asp:ListItem Value="PHILIPPINES">PHILIPPINES</asp:ListItem>
                            <asp:ListItem Value="PITCAIRN">PITCAIRN</asp:ListItem>
                            <asp:ListItem Value="POLAND">POLAND</asp:ListItem>
                            <asp:ListItem Value="PORTUGAL">PORTUGAL</asp:ListItem>
                            <asp:ListItem Value="PUERTO RICO">PUERTO RICO</asp:ListItem>
                            <asp:ListItem Value="QATAR">QATAR</asp:ListItem>
                            <asp:ListItem Value="REUNION">REUNION</asp:ListItem>
                            <asp:ListItem Value="ROMANIA">ROMANIA</asp:ListItem>
                            <asp:ListItem Value="RUSSIAN FEDERATION">RUSSIAN FEDERATION</asp:ListItem>
                            <asp:ListItem Value="RWANDA">RWANDA</asp:ListItem>
                            <asp:ListItem Value="SAINT KITTS AND NEVIS">SAINT KITTS AND NEVIS</asp:ListItem>
                            <asp:ListItem Value="SAINT LUCIA">SAINT LUCIA</asp:ListItem>
                            <asp:ListItem Value="SAINT VINCENT AND THE GRENADINES">SAINT VINCENT AND THE GRENADINES</asp:ListItem>
                            <asp:ListItem Value="SAMOA">SAMOA</asp:ListItem>
                            <asp:ListItem Value="SAN MARINO">SAN MARINO</asp:ListItem>
                            <asp:ListItem Value="SAO TOME AND PRINCIPE">SAO TOME AND PRINCIPE</asp:ListItem>
                            <asp:ListItem Value="SAUDI ARABIA">SAUDI ARABIA</asp:ListItem>
                            <asp:ListItem Value="SENEGAL">SENEGAL</asp:ListItem>
                            <asp:ListItem Value="SEYCHELLES">SEYCHELLES</asp:ListItem>
                            <asp:ListItem Value="SIERRA LEONE">SIERRA LEONE</asp:ListItem>
                            <asp:ListItem Value="SINGAPORE">SINGAPORE</asp:ListItem>
                            <asp:ListItem Value="SLOVAKIA (SLOVAK REPUBLIC)">SLOVAKIA (SLOVAK REPUBLIC)</asp:ListItem>
                            <asp:ListItem Value="SLOVENIA">SLOVENIA</asp:ListItem>
                            <asp:ListItem Value="SOLOMON ISLANDS">SOLOMON ISLANDS</asp:ListItem>
                            <asp:ListItem Value="SOMALIA">SOMALIA</asp:ListItem>
                            <asp:ListItem Value="SOUTH AFRICA">SOUTH AFRICA</asp:ListItem>
                            <asp:ListItem Value="SOUTH GEORGIA AND THE SOUTH SANDWICH ISLANDS">SOUTH GEORGIA AND THE SOUTH SANDWICH ISLANDS</asp:ListItem>
                            <asp:ListItem Value="SPAIN">SPAIN</asp:ListItem>
                            <asp:ListItem Value="SRI LANKA">SRI LANKA</asp:ListItem>
                            <asp:ListItem Value="ST. HELENA">ST. HELENA</asp:ListItem>
                            <asp:ListItem Value="ST. PIERRE AND MIQUELON">ST. PIERRE AND MIQUELON</asp:ListItem>
                            <asp:ListItem Value="SUDAN">SUDAN</asp:ListItem>
                            <asp:ListItem Value="SURINAME">SURINAME</asp:ListItem>
                            <asp:ListItem Value="SVALBARD AND JAN MAYEN ISLANDS">SVALBARD AND JAN MAYEN ISLANDS</asp:ListItem>
                            <asp:ListItem Value="SWAZILAND">SWAZILAND</asp:ListItem>
                            <asp:ListItem Value="SWEDEN">SWEDEN</asp:ListItem>
                            <asp:ListItem Value="SWITZERLAND">SWITZERLAND</asp:ListItem>
                            <asp:ListItem Value="SYRIAN ARAB REPUBLIC">SYRIAN ARAB REPUBLIC</asp:ListItem>
                            <asp:ListItem Value="TAIWAN, REPUBLIC OF CHINA">TAIWAN, REPUBLIC OF CHINA</asp:ListItem>
                            <asp:ListItem Value="TAJIKISTAN">TAJIKISTAN</asp:ListItem>
                            <asp:ListItem Value="TANZANIA, UNITED REPUBLIC OF">TANZANIA, UNITED REPUBLIC OF</asp:ListItem>
                            <asp:ListItem Value="THAILAND">THAILAND</asp:ListItem>
                            <asp:ListItem Value="TOGO">TOGO</asp:ListItem>
                            <asp:ListItem Value="TOKELAU">TOKELAU</asp:ListItem>
                            <asp:ListItem Value="TONGA">TONGA</asp:ListItem>
                            <asp:ListItem Value="TRINIDAD AND TOBAGO">TRINIDAD AND TOBAGO</asp:ListItem>
                            <asp:ListItem Value="TUNISIA">TUNISIA</asp:ListItem>
                            <asp:ListItem Value="TURKEY">TURKEY</asp:ListItem>
                            <asp:ListItem Value="TURKMENISTAN">TURKMENISTAN</asp:ListItem>
                            <asp:ListItem Value="TURKS AND CAICOS ISLANDS">TURKS AND CAICOS ISLANDS</asp:ListItem>
                            <asp:ListItem Value="TUVALU">TUVALU</asp:ListItem>
                            <asp:ListItem Value="UGANDA">UGANDA</asp:ListItem>
                            <asp:ListItem Value="UKRAINE">UKRAINE</asp:ListItem>
                            <asp:ListItem Value="UNITED ARAB EMIRATES">UNITED ARAB EMIRATES</asp:ListItem>
                            <asp:ListItem Value="UNITED KINGDOM">UNITED KINGDOM</asp:ListItem>
                            <asp:ListItem Value="UNITED STATES OF AMERICA">UNITED STATES OF AMERICA</asp:ListItem>
                            <asp:ListItem Value="URUGUAY">URUGUAY</asp:ListItem>
                            <asp:ListItem Value="UZBEKISTAN">UZBEKISTAN</asp:ListItem>
                            <asp:ListItem Value="VANUATU">VANUATU</asp:ListItem>
                            <asp:ListItem Value="VATICAN CITY STATE (HOLY SEE)">VATICAN CITY STATE (HOLY SEE)</asp:ListItem>
                            <asp:ListItem Value="VENEZUELA">VENEZUELA</asp:ListItem>
                            <asp:ListItem Value="VIETNAM">VIETNAM</asp:ListItem>
                            <asp:ListItem Value="VIRGIN ISLANDS (BRITISH)">VIRGIN ISLANDS (BRITISH)</asp:ListItem>
                            <asp:ListItem Value="VIRGIN ISLANDS (U.S.)">VIRGIN ISLANDS (U.S.)</asp:ListItem>
                            <asp:ListItem Value="WALLIS AND FUTUNA ISLANDS">WALLIS AND FUTUNA ISLANDS</asp:ListItem>
                            <asp:ListItem Value="WESTERN SAHARA">WESTERN SAHARA</asp:ListItem>
                            <asp:ListItem Value="YEMEN">YEMEN</asp:ListItem>
                            <asp:ListItem Value="YUGOSLAVIA">YUGOSLAVIA</asp:ListItem>
                            <asp:ListItem Value="ZAIRE">ZAIRE</asp:ListItem>
                            <asp:ListItem Value="ZAMBIA">ZAMBIA</asp:ListItem>
                            <asp:ListItem Value="ZIMBABWE">ZIMBABWE</asp:ListItem>
                        </asp:DropDownList>
                        <asp:CompareValidator ID="cvCountry" runat="server" ValueToCompare="-- please select --" Operator="NotEqual" Font-Names="Arial" ControlToValidate="ddlCountry"> required!</asp:CompareValidator>
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td align="right" style="color: red">
                    Email:</td>
                <td colspan="2">
                        <asp:TextBox runat="server" ID="txtEmail" Font-Size="XX-Small" MaxLength="50" Width="200px"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="rfvEmail" runat="server" ControlToValidate="txtEmail" Font-Names="Arial"> required!</asp:RequiredFieldValidator><asp:RegularExpressionValidator ID="revEmail" runat="server" ErrorMessage="Not a valid email address!" ControlToValidate="txtEmail"
                                ValidationExpression="\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*" Font-Names="Arial" Font-Size="XX-Small"/></td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td align="right" style="color: red">
                    Telephone:</td>
                <td>
                        <asp:TextBox runat="server" ID="txtTelephone" Font-Size="XX-Small" MaxLength="50" Width="200px"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="rfvTelephone" runat="server" ControlToValidate="txtEmail" Font-Names="Arial"> required!</asp:RequiredFieldValidator>
                </td>
                <td>
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td align="right" style="color: red">
                    Quantity Required:</td>
                <td colspan="2">
                        <asp:TextBox runat="server" ID="txtQuantity" Font-Size="XX-Small" MaxLength="6" Width="80px"></asp:TextBox>
                    <asp:RequiredFieldValidator ID="rfvQuantityRequired" runat="server" ControlToValidate="txtQuantity"
                        Font-Names="Arial"> required!</asp:RequiredFieldValidator><asp:RangeValidator
                            ID="rvQuantityRequired" runat="server" ControlToValidate="txtQuantity" ErrorMessage="Must be a number between 1 and 10000, digits only"
                            Font-Names="Arial" MaximumValue="10000" MinimumValue="1" Type="Integer"></asp:RangeValidator></td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td align="right" style="color: red">
                    Comments:</td>
                <td colspan="2" valign="middle">
                        <asp:TextBox runat="server" TextMode="MultiLine" Width="350px" ID="txtComments" Font-Size="XX-Small" MaxLength="200" Rows="4" Font-Names="Verdana"></asp:TextBox>
                    <asp:RequiredFieldValidator ID="rfvComments" runat="server" ControlToValidate="txtComments"
                        Font-Names="Arial"> required!</asp:RequiredFieldValidator></td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td>
                </td>
                <td>
                </td>
                <td>
                        <asp:Button runat="server" ID="btnSubmit" Text="submit" OnClick="btnSubmit_Click"></asp:Button>&nbsp;&nbsp;<asp:Button runat="server" ID="btnReset" CausesValidation="False" Text="clear" OnClick="btnReset_Click"></asp:Button>
                </td>
                <td>
                </td>
                <td>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
