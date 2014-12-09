<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">
    Dim gsConn As String = ConfigurationManager.AppSettings("AIMSRootConnectionString")
    Dim oCmd As SqlCommand
    Dim sGUID As String
    Dim oDataTable As New DataTable

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        sGUID = Request.QueryString("GUID")
        If sGUID <> String.Empty Then
            Call FetchAuthRequest()
            Call ProcessAuthRequest()
            tbQtyAuthorised.Focus()
        End If
        Call SetStyleSheet()
    End Sub
    
    Protected Sub SetStyleSheet()
        Dim hlCSSLink As New HtmlLink
        hlCSSLink.Href = Session("StyleSheetPath")
        hlCSSLink.Attributes.Add("rel", "stylesheet")
        hlCSSLink.Attributes.Add("type", "text/css")
        Page.Header.Controls.Add(hlCSSLink)
    End Sub

    Protected Sub FetchAuthRequest()
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter("spASPNET_Product_GetAuthorisationByGUID", oConn)
        Try
            oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@GUID", SqlDbType.VarChar, 20))
            oAdapter.SelectCommand.Parameters("@GUID").Value = sGUID

            oAdapter.Fill(oDataTable)
        Catch
        End Try
    End Sub

    Protected Sub ProcessAuthRequest()
        If oDataTable.Rows.Count > 0 Then
            If Not IsDBNull(oDataTable.Rows(0).Item("Granted")) Then
                lblMessage.Text = "This authorisation request has already been processed"
                Call ShowMessage()
            Else
                Call ShowAuth()
            End If
        Else
            lblMessage.Text = "Authorisation request not found"
            Call ShowMessage()
        End If
        dvAuthorise.DataSource = oDataTable
        dvAuthorise.DataBind()
    End Sub
    
    Protected Sub ShowAuth()
        pnlAuthorise.Visible = True
        pnlMessage.Visible = False
    End Sub
    
    Protected Sub ShowMessage()
        pnlMessage.Visible = True
        pnlAuthorise.Visible = False
    End Sub
    
    Protected Sub btnAuthorise_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim oConn As New SqlConnection(gsConn)
        Dim hidAuthoriseId As HiddenField
        
        lblValidationError.Text = String.Empty
        If Not IsNumeric(tbQtyAuthorised.Text) OrElse CInt(tbQtyAuthorised.Text) < 0 Then
            lblValidationError.Text = "Please enter a value for quantity authorised"
        End If
        If Not (tbDuration.Text = String.Empty Or tbDuration.Text.ToLower = "unlimited" Or (IsNumeric(tbDuration.Text) AndAlso CInt(tbDuration.Text) >= 0)) Then
            lblValidationError.Text = "Please enter a value for duration or leave blank to indicate unlimited"
        End If
        If lblValidationError.Text <> String.Empty Then
            Exit Sub
        End If
        
        hidAuthoriseId = CType(dvAuthorise.Rows(6).Controls(1).Controls.Item(1), HiddenField)

        Dim dtExpiryDate As DateTime
        If IsNumeric(tbDuration.Text) Then
            Dim tsDuration As TimeSpan = TimeSpan.FromDays(CInt(tbDuration.Text))
            dtExpiryDate = Now() + tsDuration
        Else
            Dim tsDuration As TimeSpan = TimeSpan.FromDays(5000)
            dtExpiryDate = Now() + tsDuration
        End If
        Dim sExpiryDate As String = dtExpiryDate.ToString

        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Product_SetAuthorisation", oConn)
        oCmd.CommandType = CommandType.StoredProcedure

        Dim paramAuthorisationKey As SqlParameter = New SqlParameter("@AuthorisationKey", SqlDbType.Int)
        paramAuthorisationKey.Value = hidAuthoriseId.Value
        oCmd.Parameters.Add(paramAuthorisationKey)
                
        Dim paramResult As SqlParameter = New SqlParameter("@Result", SqlDbType.Bit)
        paramResult.Value = 1
        oCmd.Parameters.Add(paramResult)
                
        Dim paramQuantityAuthorised As SqlParameter = New SqlParameter("@Quantity", SqlDbType.Int)
        paramQuantityAuthorised.Value = tbQtyAuthorised.Text
        oCmd.Parameters.Add(paramQuantityAuthorised)
                
        Dim paramExpiryDate As SqlParameter = New SqlParameter("@Expiry", SqlDbType.SmallDateTime)
        paramExpiryDate.Value = sExpiryDate
        oCmd.Parameters.Add(paramExpiryDate)

        Dim sMessage As String = tbMessage.Text.Trim
        Dim paramMessage As SqlParameter = New SqlParameter("@Message", SqlDbType.VarChar, 4000)
        If sMessage <> String.Empty Then
            paramMessage.Value = sMessage
        Else
            paramMessage.Value = System.Data.SqlTypes.SqlString.Null
        End If
        oCmd.Parameters.Add(paramMessage)

        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
            
            lblMessage.Text = "Authorisation complete - thank you"
            Call ShowMessage()
            
        Catch ex As SqlException
            WebMsgBox.Show("Unable to set authorisation status - aborting")
        End Try
    End Sub
    
    Protected Sub btnDecline_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim oConn As New SqlConnection(gsConn)
        Dim hidAuthoriseId As HiddenField
        hidAuthoriseId = CType(dvAuthorise.Rows(6).Controls(1).Controls.Item(1), HiddenField)

        Dim dtExpiryDate As DateTime = Now()
        Dim sExpiryDate As String = dtExpiryDate.ToString

        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Product_SetAuthorisation", oConn)
        oCmd.CommandType = CommandType.StoredProcedure

        Dim paramAuthorisationKey As SqlParameter = New SqlParameter("@AuthorisationKey", SqlDbType.Int)
        paramAuthorisationKey.Value = hidAuthoriseId.Value
        oCmd.Parameters.Add(paramAuthorisationKey)
                
        Dim paramResult As SqlParameter = New SqlParameter("@Result", SqlDbType.Bit)
        paramResult.Value = 0
        oCmd.Parameters.Add(paramResult)
                
        Dim paramQuantityAuthorised As SqlParameter = New SqlParameter("@Quantity", SqlDbType.Int)
        paramQuantityAuthorised.Value = 0
        oCmd.Parameters.Add(paramQuantityAuthorised)
                
        Dim paramExpiryDate As SqlParameter = New SqlParameter("@Expiry", SqlDbType.SmallDateTime)
        paramExpiryDate.Value = sExpiryDate
        oCmd.Parameters.Add(paramExpiryDate)

        Dim sMessage As String = tbMessage.Text.Trim
        Dim paramMessage As SqlParameter = New SqlParameter("@Message", SqlDbType.VarChar, 4000)
        If sMessage <> String.Empty Then
            paramMessage.Value = sMessage
        Else
            paramMessage.Value = System.Data.SqlTypes.SqlString.Null
        End If
        oCmd.Parameters.Add(paramMessage)

        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
            
            lblMessage.Text = "Authorisation declined - An email has been sent to the requester"
            Call ShowMessage()
            
        Catch ex As SqlException
            WebMsgBox.Show("Unable to set authorisation status - aborting")
        End Try
    End Sub
    
</script>

<html xmlns=" http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>Quick Authorise</title>
</head>
<body style="font-family: Verdana">
    <form id="form1" runat="server">
    <div>
        <strong>Quick Authorise</strong><br />
            <br />
        <asp:Panel ID="pnlAuthorise" runat="server" Width="100%" Visible="False">
            <asp:DetailsView ID="dvAuthorise" runat="server" Height="40px" Width="496px" AutoGenerateRows="False" Font-Names="Verdana" Font-Size="X-Small" DefaultMode="Edit">
                <Fields>
                    <asp:BoundField DataField="UserID" HeaderText="UserID" ReadOnly="True" SortExpression="UserID" />
                    <asp:BoundField DataField="FirstName" HeaderText="First Name" ReadOnly="True" SortExpression="FirstName" />
                    <asp:BoundField DataField="LastName" HeaderText="Last Name" ReadOnly="True" SortExpression="LastName" />
                    <asp:BoundField DataField="ProductCode" HeaderText="Product Code" ReadOnly="True"
                        SortExpression="ProductCode" />
                    <asp:BoundField DataField="ProductDate" HeaderText="Product Date" ReadOnly="True"
                        SortExpression="ProductDate" />
                    <asp:BoundField DataField="ProductDescription" HeaderText="Description" ReadOnly="True"
                        SortExpression="ProductDescription" />
                    <asp:TemplateField HeaderText="Qty Requested">
                        <ItemTemplate>
                            <asp:HiddenField ID="hidAuthoriseId" Value='<%# Eval("id") %>' runat="server" />
                            <asp:Label ID="lblQtyRequested" Text='<%# Eval("RequestedQty") %>' runat="server"></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>
                </Fields>
            </asp:DetailsView>
            <br />
            <table>
                <tr>
                    <td style="width: 100px; font-size: x-small; font-family: Verdana;">
                        Qty authorised:</td>
                    <td style="width: 100px; font-size: x-small; font-family: Verdana;">
                        <asp:TextBox ID="tbQtyAuthorised" runat="server" Width="42px" BackColor="#FFFFC0"></asp:TextBox></td>
                    <td style="width: 100px; font-size: x-small; font-family: Verdana;">
                        Duration of authorisation:</td>
                    <td style="width: 100px; font-size: x-small; font-family: Verdana;">
                        <asp:TextBox ID="tbDuration" Text="unlimited" runat="server" Width="55px" BackColor="#FFFFC0"></asp:TextBox>
                        (days)</td>
                </tr>
            </table>
        <strong style="font-size: xx-small">
            <br />
            Message to recipient (optional):<br />
        </strong>
        <asp:TextBox ID="tbMessage" runat="server" Height="88px" TextMode="MultiLine" Width="384px"></asp:TextBox><br />
            <br />
            <asp:Button ID="btnAuthorise" runat="server" Text="Grant Authorisation" OnClick="btnAuthorise_Click" Width="312px" /><strong>
            </strong>
        <asp:Button ID="btnDecline" runat="server" Text="Decline Authorisation" OnClick="btnDecline_Click" /><br />
            <br />
            <asp:Label ID="lblValidationError" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                ForeColor="Red"></asp:Label><br />
        </asp:Panel>
        <asp:Panel ID="pnlMessage" runat="server" Width="100%" Visible="False">
            <asp:Label ID="lblMessage" runat="server" Font-Names="Verdana" Font-Size="X-Small"></asp:Label><br />
            <br />
            &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
            <asp:Button ID="btnClose" runat="server" OnClientClick="javascript:window.close();"
                Text="Close" /></asp:Panel>
        <br />
        <asp:HyperLink ID="HyperLink1" runat="server" NavigateUrl="~/authorise.aspx?GUID=305afeed-820a-407f-a" Visible="False">Call myself</asp:HyperLink></div>
    </form>
</body>
</html>