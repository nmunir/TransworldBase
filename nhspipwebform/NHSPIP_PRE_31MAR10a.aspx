<%@ Page Language="VB" Theme="NHSPIP" MaintainScrollPositionOnPostback="true" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Data.SqlTypes" %>
<%@ Import Namespace="System.Drawing.Image" %>
<%@ Import Namespace="System.Drawing.Color" %>
<%@ Import Namespace="System.Globalization" %>
<%@ Import Namespace="System.Threading" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%@ Import Namespace="System.Net" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">
    
    Dim gbIsCreatingProductsAndAccounts As Boolean

    ' need to resolve identification of PCT in user accounts to use PCT codes rather than names
    
    ' TO TEST
    ' Check permissioning for new account and for new product
    
    ' TO DO
    ' permissioning existing users for newly created products
    ' alignment of fields in confirmation page

    ' SET VIEW ZERO STOCK ITEMS !!!!!!!!!!!!!!!!!!!!!
    
    ' Const NHSPIP_CUSTOMER_KEY As Integer = 580
    Const COUNTRY_KEY_UK As Integer = 222

    Const EMAIL_ADDRESS_NHS As String = "scr.comms@nhs.net"
    Const EMAIL_ADDRESS_ACCOUNT_HANDLER As String = "m.quinn@sprintexpress.co.uk"
    Const EMAIL_ADDRESS_SUPPORT As String = "chris.newport@sprintexpress.co.uk"
    Const EMAIL_ADDRESS_ALERTS As String = EMAIL_ADDRESS_SUPPORT & "," & EMAIL_ADDRESS_NHS
    Const EMAIL_ADDRESS_ERRORS As String = EMAIL_ADDRESS_SUPPORT
    
    Const LOG_ENTRY_TYPE_ORDER As String = "ORDER"
    Const LOG_ENTRY_TYPE_PRODUCT As String = "PRODUCT"
    Const LOG_ENTRY_TYPE_ACCOUNT As String = "ACCOUNT"

    Const MESSAGE_TYPE_NHSPIP_NEW_ACCOUNT As String = "NHSPIP NEW ACCOUNT"
    Const MESSAGE_TYPE_NHSPIP_WEBFORM_ERROR As String = "NHSPIP WEBFORM ERROR"
    Const MESSAGE_TYPE_NHSPIP_CONFIRM_ORDER As String = "NHSPIP CONFIRM ORDER"
    Const MESSAGE_TYPE_NHSPIP_ALERT As String = "NHSPIP ALERT"

    Const NHSPIP_CUSTOMER_KEY As Integer = 580
    Const NHSPIPTEST_CUSTOMER_KEY As Integer = 16

    Const NHSPIP_WEBFORM_CONTROL_USERID As String = "NHSPIPPCTWebform"
    Const NHSPIPTEST_WEBFORM_CONTROL_USERID As String = "NHSPIPTESTPCTWebform"
    
    Dim gsWebFormControlUserId As String = NHSPIP_WEBFORM_CONTROL_USERID
    Dim gnCustomerKey As Integer = NHSPIP_CUSTOMER_KEY
    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsPostBack Then
        End If
        'If Server.MachineName.ToLower.Contains("chrisn") Or Server.MachineName.ToLower.Contains("vostro") Then
        gbIsCreatingProductsAndAccounts = True
        'End If
        Response.Cache.SetCacheability(HttpCacheability.NoCache)
    End Sub
   
    Protected Function GetWebformUserKey() As Integer
        Dim oDataTable As DataTable = ExecuteQueryToDataTable("SELECT [key] FROM UserProfile WHERE UserId = '" & gsWebFormControlUserId & "'")
        GetWebformUserKey = oDataTable.Rows(0).Item(0)
    End Function
    
    Protected Sub DisplayProductInfo()
        Dim oDataTable As DataTable = ExecuteQueryToDataTable("SELECT LogisticProductKey, ProductCode, ProductDescription, MaxGrabQty FROM LogisticProduct lp INNER JOIN UserProductProfile up ON  lp.LogisticProductKey = up.ProductKey WHERE lp.ArchiveFlag = 'N' AND lp.DeletedFlag = 'N' AND AbleToPick = 1 AND up.UserKey = " & GetWebformUserKey() & " ORDER BY AdRotatorText")
        gvProductInfo.DataSource = oDataTable
        gvProductInfo.DataBind()
    End Sub
   
    Protected Sub HideAllPanels()
        pnlIntro.Visible = False
        pnlDataEntry.Visible = False
        pnlFinalCheck.Visible = False
        pnlComplete.Visible = False
    End Sub
   
    Protected Sub ShowMaterialsReservation()
        Call HideAllPanels()
        pnlDataEntry.Visible = True
        Call DisplayProductInfo()
        tbFirstName.Focus()
    End Sub
   
    Protected Function Highlight(ByVal sText As String) As String
        Highlight = "<font color=""blue"">" & sText & "</font>"
    End Function

    Protected Function IsOrderingAtLeastOneProduct() As Boolean
        IsOrderingAtLeastOneProduct = False
        For Each gvr As GridViewRow In gvProductInfo.Rows
            If gvr.RowType = DataControlRowType.DataRow Then
                Dim tb As TextBox = gvr.FindControl("tbQty")
                tb.Text = tb.Text.Trim
                If tb.Text <> String.Empty Then
                    If IsNumeric(tb.Text) Then
                        Dim nQty As Integer = CInt(tb.Text)
                        If nQty > 0 Then
                            IsOrderingAtLeastOneProduct = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        Next
    End Function
    
    Protected Function NL(ByVal sType As String) As String
        If sType = "HTML" Then
            NL = "<br/>"
        Else
            NL = vbNewLine
        End If
    End Function
    
    Protected Function BuildAccountCreatedConfirmation(ByVal sType As String) As String
        Dim sbOrderConfirmation As New StringBuilder
        sbOrderConfirmation.Append(Highlight("NEW ACCOUNT CREATED"))
        sbOrderConfirmation.Append(NL(sType))
        sbOrderConfirmation.Append(NL(sType))
        sbOrderConfirmation.Append("An account has been created for you on the Sprint NHS PIP web site.")
        sbOrderConfirmation.Append(NL(sType))
        sbOrderConfirmation.Append(NL(sType))
        sbOrderConfirmation.Append("User name: ")
        sbOrderConfirmation.Append(tbNewUserName.Text)
        sbOrderConfirmation.Append(NL(sType))
        sbOrderConfirmation.Append("Password: ")
        sbOrderConfirmation.Append(tbNewPassword.Text)
        sbOrderConfirmation.Append(NL(sType))
        sbOrderConfirmation.Append(NL(sType))
        sbOrderConfirmation.Append("Visit http://www.sprintexpress.co.uk/nhspip/" & " to access your account.")
        sbOrderConfirmation.Append(NL(sType))
        sbOrderConfirmation.Append(NL(sType))
        sbOrderConfirmation.Append(NL(sType))
        sbOrderConfirmation.Append("Please do not reply to this email, as replies are not monitored.")
        BuildAccountCreatedConfirmation = sbOrderConfirmation.ToString
    End Function
    
    Protected Function BuildOrderEmailConfirmation(ByVal sType As String, ByVal nOrderNo As Integer) As String
        Dim sbOrderConfirmation As New StringBuilder
        sbOrderConfirmation.Append(Highlight("Thank you for your order, which is now being processed as per the ordering guidance."))
        sbOrderConfirmation.Append(NL(sType))
        sbOrderConfirmation.Append(NL(sType))
        sbOrderConfirmation.Append(Highlight("Your order may take up to 4 weeks to become available for draw down from the Sprint NHS PIP web site. Thank you for your order, which is now being processed as per the ordering guidance."))
        sbOrderConfirmation.Append(NL(sType))
        sbOrderConfirmation.Append(NL(sType))
        sbOrderConfirmation.Append(Highlight("Please note the short code for your PCT/Trust is " & ddlPCTs.SelectedValue & "."))
        sbOrderConfirmation.Append(NL(sType))
        sbOrderConfirmation.Append(NL(sType))
        sbOrderConfirmation.Append(Highlight("ORDER SUMMARY"))
        sbOrderConfirmation.Append(NL(sType))
        sbOrderConfirmation.Append(NL(sType))
        sbOrderConfirmation.Append(Highlight("Order ref: #"))
        sbOrderConfirmation.Append(nOrderNo.ToString)
        sbOrderConfirmation.Append(Highlight(" received "))
        sbOrderConfirmation.Append(Now.ToString("dd-MMM-yyyy hh:mm"))
        sbOrderConfirmation.Append(Highlight(" from "))
        sbOrderConfirmation.Append(tbFirstName.Text)
        sbOrderConfirmation.Append(" ")
        sbOrderConfirmation.Append(tbLastName.Text)
        sbOrderConfirmation.Append(" for PCT/Trust ")
        sbOrderConfirmation.Append(ddlPCTs.SelectedItem.Text)
        sbOrderConfirmation.Append(NL(sType))
        sbOrderConfirmation.Append(NL(sType))
        If rbRequireAccount.Checked Then
            sbOrderConfirmation.Append(Highlight("Your account user name is: " & tbNewUserName.Text & " ."))
            sbOrderConfirmation.Append(NL(sType))
            sbOrderConfirmation.Append(Highlight("Your password is: " & tbNewPassword.Text & " ."))
        ElseIf rbHaveAccount.Checked Then
            sbOrderConfirmation.Append(Highlight("Your account user name is: " & tbUserName.Text & " ."))
        Else
            sbOrderConfirmation.Append(Highlight("You did not request any account action."))
        End If
        sbOrderConfirmation.Append(NL(sType))
        sbOrderConfirmation.Append(NL(sType))
        sbOrderConfirmation.Append(Highlight("ITEMS ORDERED"))
        sbOrderConfirmation.Append(NL(sType))
        sbOrderConfirmation.Append(NL(sType))
        
        For Each gvr As GridViewRow In gvProductInfo.Rows
            If gvr.RowType = DataControlRowType.DataRow Then
                Dim tb As TextBox = gvr.FindControl("tbQty")
                tb.Text = tb.Text.Trim
                If tb.Text <> String.Empty Then
                    If IsNumeric(tb.Text) Then
                        Dim nQty As Integer = CInt(tb.Text)
                        If nQty > 0 Then
                            Dim tcProductCode As TableCell = gvr.Cells(0)
                            sbOrderConfirmation.Append(tcProductCode.Text)
                            sbOrderConfirmation.Append(Highlight(" - Qty: "))
                            sbOrderConfirmation.Append(nQty)
                            sbOrderConfirmation.Append("<br />")
                        End If
                    Else
                        WebMsgBox.Show("Internal error - expected a number - please report this to Sprint IT development (it@sprintexpress.co.uk)")
                    End If
                End If
            End If
        Next
        sbOrderConfirmation.Append(NL(sType))
        sbOrderConfirmation.Append("To draw down your order when it becomes available please visit http://www.sprintexpress.co.uk/nhspip")
        sbOrderConfirmation.Append(NL(sType))
        sbOrderConfirmation.Append(NL(sType))
        sbOrderConfirmation.Append("If you have any queries regarding your order please email NHS Connecting for Health SCR Communications Team at scr.comms@nhs.net. When enquiring please state your name, organisation, date of order and order reference. Please do not reply to this email, as replies are not monitored.")
        sbOrderConfirmation.Append(NL(sType))
        BuildOrderEmailConfirmation = sbOrderConfirmation.ToString
    End Function
    
    Protected Sub BuildOrderDisplayConfirmation()
        Dim sbOrderConfirmation As New StringBuilder
        sbOrderConfirmation.Append(Highlight("Your name: "))
        sbOrderConfirmation.Append(tbFirstName.Text)
        sbOrderConfirmation.Append(" ")
        sbOrderConfirmation.Append(tbLastName.Text)
        sbOrderConfirmation.Append("<br />")
        sbOrderConfirmation.Append(Highlight("Your email address: "))
        sbOrderConfirmation.Append(tbEmail.Text)
        sbOrderConfirmation.Append("<br />")
        sbOrderConfirmation.Append(Highlight("Your PCT: "))
        sbOrderConfirmation.Append(ddlPCTs.SelectedItem.Text)
        sbOrderConfirmation.Append("<br />")
        sbOrderConfirmation.Append(Highlight("Requested: "))
        sbOrderConfirmation.Append(Now.ToString("dd-MMM-yyyy hh:mm"))
        sbOrderConfirmation.Append("<br />")
        sbOrderConfirmation.Append("<br />")
        If rbHaveAccount.Checked Then
            sbOrderConfirmation.Append(Highlight("You already have an account on the Sprint NHS PIP system. Your user name is "))
            sbOrderConfirmation.Append(tbUserName.Text)
            sbOrderConfirmation.Append(Highlight("."))
        End If
        If rbRequireAccount.Checked Then
            sbOrderConfirmation.Append(Highlight("You have requested an account on the Sprint NHS PIP system.  Your user name will be "))
            sbOrderConfirmation.Append(tbNewUserName.Text)
            sbOrderConfirmation.Append(Highlight(" and your password will be "))
            sbOrderConfirmation.Append(tbNewPassword.Text)
            sbOrderConfirmation.Append(Highlight(" ."))
        End If
        If rbAccountNoActionRequired.Checked Then
            sbOrderConfirmation.Append(Highlight("You do not currently have an account on the Sprint NHS PIP system and you have not requested one."))
        End If
        sbOrderConfirmation.Append("<br />")
        sbOrderConfirmation.Append("<br />")
        sbOrderConfirmation.Append(Highlight("You are ordering the following products:"))
        sbOrderConfirmation.Append("<br />")
        sbOrderConfirmation.Append("<br />")
        For Each gvr As GridViewRow In gvProductInfo.Rows
            If gvr.RowType = DataControlRowType.DataRow Then
                Dim tb As TextBox = gvr.FindControl("tbQty")
                tb.Text = tb.Text.Trim
                If tb.Text <> String.Empty Then
                    If IsNumeric(tb.Text) Then
                        Dim nQty As Integer = CInt(tb.Text)
                        If nQty > 0 Then
                            Dim tcProductCode As TableCell = gvr.Cells(0)
                            sbOrderConfirmation.Append(tcProductCode.Text)
                            sbOrderConfirmation.Append(Highlight(" - Qty: "))
                            sbOrderConfirmation.Append(nQty)
                            sbOrderConfirmation.Append("<br />")
                        End If
                    Else
                        WebMsgBox.Show("Internal error - expected a number - please report this to Sprint IT development (it@sprintexpress.co.uk)")
                    End If
                End If
            End If
        Next
        sbOrderConfirmation.Append("<br />")
        sbOrderConfirmation.Append("<br />")
        tbInformation.Text = tbInformation.Text.Trim
        If tbInformation.Text <> String.Empty Then
            Dim sInformation As String = tbInformation.Text.Replace(vbNewLine, "<br />")
            sbOrderConfirmation.Append(Highlight("Additional information provided:"))
            sbOrderConfirmation.Append("<br />")
            sbOrderConfirmation.Append(sInformation)
            sbOrderConfirmation.Append("<br />")
            sbOrderConfirmation.Append("<br />")
        End If
        sbOrderConfirmation.Append(Highlight("Please note that it may take up to 4 weeks for these products to become available."))
        sbOrderConfirmation.Append("<br />")
        sbOrderConfirmation.Append("<br />")
        sbOrderConfirmation.Append(Highlight("Now click the SUBMIT ORDER button to submit your order. We will send a confirmation to your email address."))
        sbOrderConfirmation.Append("<br />")
        sbOrderConfirmation.Append("<br />")
        sbOrderConfirmation.Append(Highlight("Thank you."))
        sbOrderConfirmation.Append("<br />")
        sbOrderConfirmation.Append("<br />")
        lblOrderSummary.Text = sbOrderConfirmation.ToString
    End Sub
    
    Protected Sub ShowFinalCheck()
        Call HideAllPanels()
        Call BuildOrderDisplayConfirmation()
        pnlFinalCheck.Visible = True
    End Sub
   
    Protected Sub ShowOrderComplete()
        Call HideAllPanels()
        pnlComplete.Visible = True
    End Sub
   
    Protected Sub InitPCTList(ByVal sType As String)
        Dim oListItemCollection As ListItemCollection = ExecuteQueryToListItemCollection("SELECT PCTName, PCTAbbreviation FROM NHSPCTs WHERE IsVirtual = 0 AND Type = '" & sType & "' ORDER BY PCTName", "PCTName", "PCTAbbreviation")
        Dim sOrganisationType As String = String.Empty
        If ddlPCTs.Items.Count > 0 Then
            For i As Integer = ddlPCTs.Items.Count - 1 To 0 Step -1
                ddlPCTs.Items.RemoveAt(i)
            Next
        End If
        Select Case sType
            Case "PCT"
                sOrganisationType = "Primary Care Trust"
            Case "SHA"
                sOrganisationType = "Strategic Health Authority"
            Case "RO"
                sOrganisationType = "Regional Office"
            Case "CARE"
                sOrganisationType = "Care Trust"
            Case "TRUST"
                sOrganisationType = "Trust"
        End Select
        ddlPCTs.Items.Add(New ListItem("Select your " & sOrganisationType, 0))
        For Each li As ListItem In oListItemCollection
            ddlPCTs.Items.Add(New ListItem(li.Text, li.Value))
        Next
        If sType = "PCT" Then
            ddlPCTs.Items.Add(New ListItem("CONNECTING FOR HEALTH (admin use only)", "CFH"))
        End If
        ddlPCTs.SelectedIndex = 0
    End Sub
    
    Protected Sub btnMaterialsReservation_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowMaterialsReservation()
    End Sub
   
    Protected Sub SendMail(ByVal sType As String, ByVal sRecipientList As String, ByVal sSubject As String, ByVal sBodyText As String, ByVal sBodyHTML As String)
        Dim sSendTo() As String = sRecipientList.Split(",")
        For Each s As String In sSendTo
            If s.Trim <> String.Empty Then
                Dim oConn As New SqlConnection(gsConn)
                Try
                    Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Email_AddToQueue", oConn)
                    oCmd.CommandType = CommandType.StoredProcedure

                    oCmd.Parameters.Add(New SqlParameter("@MessageId", SqlDbType.NVarChar, 20))
                    oCmd.Parameters("@MessageId").Value = sType
   
                    oCmd.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
                    oCmd.Parameters("@CustomerKey").Value = gnCustomerKey
   
                    oCmd.Parameters.Add(New SqlParameter("@StockBookingKey", SqlDbType.Int))
                    oCmd.Parameters("@StockBookingKey").Value = 0
   
                    oCmd.Parameters.Add(New SqlParameter("@ConsignmentKey", SqlDbType.Int))
                    oCmd.Parameters("@ConsignmentKey").Value = 0
   
                    oCmd.Parameters.Add(New SqlParameter("@ProductKey", SqlDbType.Int))
                    oCmd.Parameters("@ProductKey").Value = 0
   
                    oCmd.Parameters.Add(New SqlParameter("@To", SqlDbType.NVarChar, 100))
                    oCmd.Parameters("@To").Value = s
   
                    oCmd.Parameters.Add(New SqlParameter("@Subject", SqlDbType.NVarChar, 60))
                    oCmd.Parameters("@Subject").Value = sSubject
   
                    oCmd.Parameters.Add(New SqlParameter("@BodyText", SqlDbType.NText))
                    oCmd.Parameters("@BodyText").Value = sBodyText
   
                    oCmd.Parameters.Add(New SqlParameter("@BodyHTML", SqlDbType.NText))
                    oCmd.Parameters("@BodyHTML").Value = sBodyHTML
   
                    oCmd.Parameters.Add(New SqlParameter("@QueuedBy", SqlDbType.Int))
                    oCmd.Parameters("@QueuedBy").Value = 0
   
                    oConn.Open()
                    oCmd.ExecuteNonQuery()
                Catch ex As Exception
                    WebMsgBox.Show("Error in SendMail: " & ex.Message)
                Finally
                    oConn.Close()
                End Try
            End If
        Next
    End Sub

    Protected Function ValidateProductSelection() As Boolean
        ValidateProductSelection = True
        For Each gvr As GridViewRow In gvProductInfo.Rows
            If gvr.RowType = DataControlRowType.DataRow Then
                Dim tcProductCode As TableCell = gvr.Cells(0)
                Dim tcMaxGrab As TableCell = gvr.Cells(2)
                Dim tb As TextBox = gvr.FindControl("tbQty")
                Dim nMaxGrab As Integer = CInt(tcMaxGrab.Text)
                If tb.Text.Trim <> String.Empty Then
                    Dim nQty As Integer = CInt(tb.Text)
                    If nMaxGrab > 0 Then
                        If nQty > nMaxGrab Then
                            WebMsgBox.Show("You have requested more than the allowed quantity for product " & tcProductCode.Text & ".")
                            ValidateProductSelection = False
                            Exit Function
                        End If
                    End If
                End If
            End If
        Next
    End Function
    
    Protected Sub btnFinalCheck_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If IsAllRequiredInfoEntered() Then
            If IsOrderingAtLeastOneProduct() Or rbRequireAccount.Checked Then
                If ValidateProductSelection() Then
                    Call ShowFinalCheck()
                End If
            Else
                WebMsgBox.Show("You have not selected any products, or requested account creation.")
            End If
        End If
    End Sub

    Protected Sub lnkbtnCheckAccountExists_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call AccountExists(bSilent:=False)
    End Sub
    
    Protected Function AccountExists(ByVal bSilent As Boolean) As Boolean
        AccountExists = False
        tbUserName.Text = tbUserName.Text.Trim
        If tbUserName.Text <> String.Empty Then
            If UserNameIsAvailable(tbUserName.Text) Then
                WebMsgBox.Show("Sorry, we could not find that user name.")
                tbUserName.Focus()
            Else
                If Not bSilent Then
                    WebMsgBox.Show("Okay, we found that user name.")
                End If
                AccountExists = True
            End If
        Else
            WebMsgBox.Show("Please enter your user name.")
            tbUserName.Focus()
        End If
        
    End Function

    Protected Sub lnkbtnCheckUserNameAvailable_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call CheckUserNameAvailable(bsilent:=False)
    End Sub

    Protected Function CheckUserNameAvailable(ByVal bSilent As Boolean) As Boolean
        CheckUserNameAvailable = False
        tbNewUserName.Text = tbNewUserName.Text.Trim
        If tbNewUserName.Text <> String.Empty Then
            If UserNameIsAvailable(tbNewUserName.Text) Then
                If Not bSilent Then
                    WebMsgBox.Show("Okay, this user name is available.")
                End If
                CheckUserNameAvailable = True
            Else
                WebMsgBox.Show("Sorry, this user name is not available.")
            End If
        Else
            WebMsgBox.Show("Please enter your user name.")
        End If
    End Function
    
    Protected Function NewAccountDetailsProvided() As Boolean
        NewAccountDetailsProvided = False
        If tbNewPassword.Text.Length <> tbNewPassword.Text.Trim.Length Then
            WebMsgBox.Show("Please remove leading and trailing spaces from your password to avoid confusion.")
            Exit Function
        End If
        If CheckUserNameAvailable(bSilent:=True) Then
            If tbNewPassword.Text.Length >= 6 Then
                Dim nDigitCount As Integer = 0
                Dim nLowerCaseCount As Integer = 0
                Dim nUpperCaseCount As Integer = 0
                Dim nRepeatedChars As Integer = 0
                Dim encoding As New System.Text.ASCIIEncoding()
                Dim bytePassword() As Byte = encoding.GetBytes(tbNewPassword.Text)
                For Each b As Byte In bytePassword
                    Dim c As String = Chr(b)
                    If c >= "0" And c <= "9" Then
                        nDigitCount += 1
                    End If
                    If c >= "a" And c <= "z" Then
                        nLowerCaseCount += 1
                    End If
                    If c >= "A" And c <= "Z" Then
                        nUpperCaseCount += 1
                    End If
                    If tbNewPassword.Text.IndexOf(c) <> tbNewPassword.Text.LastIndexOf(c) Then
                        nRepeatedChars += 1
                    End If
                Next
                If nDigitCount = 0 Then
                    WebMsgBox.Show("Your password must contain at least one digit.")
                    tbNewPassword.Focus()
                    Exit Function
                End If
                If nLowerCaseCount = 0 Then
                    WebMsgBox.Show("Your password must contain at least one lower case character.")
                    tbNewPassword.Focus()
                    Exit Function
                End If
                If nUpperCaseCount = 0 Then
                    WebMsgBox.Show("Your password must contain at least one upper case character.")
                    tbNewPassword.Focus()
                    Exit Function
                End If
                If nRepeatedChars > 0 Then
                    WebMsgBox.Show("Your password must not repeat any characters.")
                    tbNewPassword.Focus()
                    Exit Function
                End If
                NewAccountDetailsProvided = True
            Else
                WebMsgBox.Show("Your password must be at least 6 characters long. Click the password help link for more information on setting a strong password.")
            End If
        End If
    End Function
    
    Protected Function IsRequiredNonAccountInfoEntered() As Boolean
        IsRequiredNonAccountInfoEntered = False
        If IsNameEntered() AndAlso IsPCTSelected() AndAlso IsValidEmailAddressProvided() Then
            IsRequiredNonAccountInfoEntered = True
        End If
    End Function

    Protected Function IsAllRequiredInfoEntered() As Boolean
        IsAllRequiredInfoEntered = False
        If IsRequiredNonAccountInfoEntered() Then
            If rbHaveAccount.Checked AndAlso AccountExists(bSilent:=True) Then
                IsAllRequiredInfoEntered = True
                Exit Function
            End If
            If rbRequireAccount.Checked AndAlso NewAccountDetailsProvided() Then
                IsAllRequiredInfoEntered = True
                Exit Function
            End If
            If rbAccountNoActionRequired.Checked Then
                IsAllRequiredInfoEntered = True
                Exit Function
            End If
            WebMsgBox.Show("Please select an account option.")
        End If
    End Function

    Protected Function bIsValidEmailAddress(ByRef sEmailAddr As String) As Boolean
        sEmailAddr = Trim$(sEmailAddr)
        Return Regex.IsMatch(sEmailAddr, "^[\w\.\-]+@[a-zA-Z0-9\-]+(\.[a-zA-Z0-9\-]{1,})*(\.[a-zA-Z]{2,3}){1,2}$")
    End Function

    Protected Function IsNameEntered() As Boolean
        If tbFirstName.Text.Length > 0 And tbLastName.Text.Length > 1 Then
            IsNameEntered = True
        Else
            IsNameEntered = False
            WebMsgBox.Show("Please enter your name.")
            If tbFirstName.Text.Length = 0 Then
                tbFirstName.Focus()
            Else
                tbLastName.Focus()
            End If
        End If
    End Function
    
    Protected Function IsPCTSelected() As Boolean
        If ddlPCTs.SelectedIndex > 0 OrElse ddlPCTs.SelectedValue <> "0" Then
            IsPCTSelected = True
        Else
            IsPCTSelected = False
            If Not divOrganisation.Visible Then
                WebMsgBox.Show("Please select your Organisation Type.")
            Else
                WebMsgBox.Show("Please select your Organisation.")
            End If
            ddlPCTs.Focus()
        End If
    End Function
    
    Protected Function IsValidEmailAddressProvided() As Boolean
        IsValidEmailAddressProvided = False
        tbEmail.Text = tbEmail.Text.Trim
        tbEmailConfirm.Text = tbEmailConfirm.Text.Trim
        If tbEmail.Text <> String.Empty Then
            If tbEmail.Text = tbEmailConfirm.Text Then
                If bIsValidEmailAddress(tbEmail.Text) Then
                    IsValidEmailAddressProvided = True
                    Exit Function
                Else
                    WebMsgBox.Show("This does not appear to be a valid email address!")
                End If
            Else
                WebMsgBox.Show("Your email address does not match the re-typed address!")
            End If
        Else
            WebMsgBox.Show("Please enter your email address.")
        End If
        tbEmail.Focus()
    End Function
    
    Protected Sub HideAllAccountRows()
        trVerifyAccount.Visible = False
        trCreateAccount.Visible = False
    End Sub
    
    Protected Sub rbHaveAccount_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim rb As RadioButton = sender
        If rb.Checked Then
            If IsRequiredNonAccountInfoEntered() Then
                Call HideAllAccountRows()
                trVerifyAccount.Visible = True
                tbUserName.Focus()
            Else
                rb.Checked = False
            End If
        End If
    End Sub

    Protected Function UserNameIsAvailable(ByVal sUserName As String) As Boolean
        Dim oDataTable As DataTable = ExecuteQueryToDataTable("SELECT UserId FROM UserProfile WHERE UserId = '" & sUserName.Replace("'", "''") & "'")
        If oDataTable.Rows.Count = 0 Then
            UserNameIsAvailable = True
        Else
            UserNameIsAvailable = False
        End If
    End Function
    
    Protected Sub rbRequireAccount_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim rb As RadioButton = sender
        If rb.Checked Then
            If IsRequiredNonAccountInfoEntered() Then
                Call HideAllAccountRows()
                trCreateAccount.Visible = True
                tbNewUserName.Text = tbEmail.Text
                tbNewUserName.Enabled = False
                tbNewPassword.Focus()
            Else
                rb.Checked = False
            End If
        End If
    End Sub

    Protected Sub rbUseEmailAddress_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        tbNewUserName.Enabled = False
        tbNewUserName.Text = tbEmail.Text
    End Sub

    Protected Sub rbSetUserName_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        tbNewUserName.Enabled = True
        tbNewUserName.Text = String.Empty
        tbNewUserName.Focus()
    End Sub
    
    Protected Sub rbAccountNoActionRequired_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call HideAllAccountRows()
    End Sub
    
    Protected Function AddNewProduct(ByVal sLogisticProductKey As Integer, ByVal sPCTAbbreviation As String) As Integer
        AddNewProduct = 0
        Dim oConn As New SqlConnection(gsConn)
        Dim oTrans As SqlTransaction
        
        Dim oDataTable As DataTable = ExecuteQueryToDataTable("SELECT * FROM LogisticProduct WHERE LogisticProductKey = " & sLogisticProductKey)
        Dim dr As DataRow = oDataTable.Rows(0)
        
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Product_AddWithAccessControl8", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
  
        Dim paramUserKey As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int, 4)
        paramUserKey.Value = 0
        oCmd.Parameters.Add(paramUserKey)

        Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
        paramCustomerKey.Value = gnCustomerKey
        oCmd.Parameters.Add(paramCustomerKey)
  
        Dim paramProductCode As SqlParameter = New SqlParameter("@ProductCode", SqlDbType.NVarChar, 25)
        paramProductCode.Value = dr("ProductCode")
        oCmd.Parameters.Add(paramProductCode)

        Dim paramProductDate As SqlParameter = New SqlParameter("@ProductDate", SqlDbType.NVarChar, 10)
        paramProductDate.Value = sPCTAbbreviation
        oCmd.Parameters.Add(paramProductDate)
  
        Dim paramMinimumStockLevel As SqlParameter = New SqlParameter("@MinimumStockLevel", SqlDbType.Int, 4)
        paramMinimumStockLevel.Value = 0
        oCmd.Parameters.Add(paramMinimumStockLevel)
        
        Dim paramDescription As SqlParameter = New SqlParameter("@ProductDescription", SqlDbType.NVarChar, 300)
        paramDescription.Value = dr("ProductDescription") & " (" & sPCTAbbreviation & ")"
        oCmd.Parameters.Add(paramDescription)

        Dim paramItemsPerBox As SqlParameter = New SqlParameter("@ItemsPerBox", SqlDbType.Int, 4)
        paramItemsPerBox.Value = dr("ItemsPerBox")
        oCmd.Parameters.Add(paramItemsPerBox)
        
        Dim paramCategory As SqlParameter = New SqlParameter("@ProductCategory", SqlDbType.NVarChar, 50)
        paramCategory.Value = String.Empty
        oCmd.Parameters.Add(paramCategory)
        
        Dim paramSubCategory As SqlParameter = New SqlParameter("@SubCategory", SqlDbType.NVarChar, 50)
        paramSubCategory.Value = String.Empty
        oCmd.Parameters.Add(paramSubCategory)

        Dim paramSubCategory2 As SqlParameter = New SqlParameter("@SubCategory2", SqlDbType.NVarChar, 50)
        paramSubCategory2.Value = String.Empty
        oCmd.Parameters.Add(paramSubCategory2)

        Dim paramUnitValue As SqlParameter = New SqlParameter("@UnitValue", SqlDbType.Money, 8)
        paramUnitValue.Value = dr("UnitValue")
        oCmd.Parameters.Add(paramUnitValue)

        Dim paramUnitValue2 As SqlParameter = New SqlParameter("@UnitValue2", SqlDbType.Money, 8)
        paramUnitValue2.Value = 0
        oCmd.Parameters.Add(paramUnitValue2)

        Dim paramLanguage As SqlParameter = New SqlParameter("@LanguageId", SqlDbType.NVarChar, 20)
        paramLanguage.Value = dr("LanguageId")
        oCmd.Parameters.Add(paramLanguage)

        Dim paramDepartment As SqlParameter = New SqlParameter("@ProductDepartmentId", SqlDbType.NVarChar, 20)
        paramDepartment.Value = String.Empty
        oCmd.Parameters.Add(paramDepartment)
        
        Dim paramWeight As SqlParameter = New SqlParameter("@UnitWeightGrams", SqlDbType.Int, 4)
        paramWeight.Value = dr("UnitWeightGrams")
        oCmd.Parameters.Add(paramWeight)

        Dim paramStockOwnedByKey As SqlParameter = New SqlParameter("@StockOwnedByKey", SqlDbType.Int, 4)
        paramStockOwnedByKey.Value = 0
        oCmd.Parameters.Add(paramStockOwnedByKey)

        Dim paramMisc1 As SqlParameter = New SqlParameter("@Misc1", SqlDbType.NVarChar, 50)
        paramMisc1.Value = sPCTAbbreviation
        oCmd.Parameters.Add(paramMisc1)
        
        Dim paramMisc2 As SqlParameter = New SqlParameter("@Misc2", SqlDbType.NVarChar, 50)
        paramMisc2.Value = dr("ProductDate")
        oCmd.Parameters.Add(paramMisc2)

        Dim paramArchive As SqlParameter = New SqlParameter("@ArchiveFlag", SqlDbType.NVarChar, 1)
        paramArchive.Value = "N"
        oCmd.Parameters.Add(paramArchive)
      
        Dim paramStatus As SqlParameter = New SqlParameter("@Status", SqlDbType.TinyInt)
        paramStatus.Value = 1                     ' 1= created, unpopulated
        oCmd.Parameters.Add(paramStatus)

        Dim paramExpiryDate As SqlParameter = New SqlParameter("@ExpiryDate", SqlDbType.SmallDateTime)
        paramExpiryDate.Value = Nothing
        oCmd.Parameters.Add(paramExpiryDate)

        Dim paramReplenishmentDate As SqlParameter = New SqlParameter("@ReplenishmentDate", SqlDbType.SmallDateTime)
        paramReplenishmentDate.Value = Nothing
        oCmd.Parameters.Add(paramReplenishmentDate)
      
        Dim paramSerialNumbers As SqlParameter = New SqlParameter("@SerialNumbersFlag", SqlDbType.NVarChar, 1)
        paramSerialNumbers.Value = "N"
        oCmd.Parameters.Add(paramSerialNumbers)

        Dim paramAdRotatorText As SqlParameter = New SqlParameter("@AdRotatorText", SqlDbType.NVarChar, 120)
        paramAdRotatorText.Value = String.Empty
        oCmd.Parameters.Add(paramAdRotatorText)

        Dim paramWebsiteAdRotatorFlag As SqlParameter = New SqlParameter("@WebsiteAdRotatorFlag", SqlDbType.Bit)
        paramWebsiteAdRotatorFlag.Value = 0
        oCmd.Parameters.Add(paramWebsiteAdRotatorFlag)

        Dim paramNotes As SqlParameter = New SqlParameter("@Notes", SqlDbType.NVarChar, 1000)
        paramNotes.Value = "Awaiting stock as at " & Date.Now.ToString("dd-MMM-yyyy hh:mm")
        oCmd.Parameters.Add(paramNotes)

        Dim paramViewOnWebForm As SqlParameter = New SqlParameter("@ViewOnWebForm", SqlDbType.Bit)
        paramViewOnWebForm.Value = 0
        oCmd.Parameters.Add(paramViewOnWebForm)
  
        Dim paramDefaultAccessFlag As SqlParameter = New SqlParameter("@DefaultAccessFlag", SqlDbType.Bit)
        paramDefaultAccessFlag.Value = 0
        oCmd.Parameters.Add(paramDefaultAccessFlag)

        Dim paramRotationProductKey As SqlParameter = New SqlParameter("@RotationProductKey", SqlDbType.Int, 4)
        paramRotationProductKey.Value = System.Data.SqlTypes.SqlInt32.Null
        oCmd.Parameters.Add(paramRotationProductKey)

        Dim paramInactivityAlertDays As SqlParameter = New SqlParameter("@InactivityAlertDays", SqlDbType.Int, 4)
        paramInactivityAlertDays.Value = 0
        oCmd.Parameters.Add(paramInactivityAlertDays)
      
        Dim paramCalendarManaged As SqlParameter = New SqlParameter("@CalendarManaged", SqlDbType.Bit)
        paramCalendarManaged.Value = 0
        oCmd.Parameters.Add(paramCalendarManaged)

        Dim paramOnDemand As SqlParameter = New SqlParameter("@OnDemand", SqlDbType.Int)
        paramOnDemand.Value = 0
        oCmd.Parameters.Add(paramOnDemand)
        
        Dim paramOnDemandPriceList As SqlParameter = New SqlParameter("@OnDemandPriceList", SqlDbType.Int)
        paramOnDemandPriceList.Value = 0
        oCmd.Parameters.Add(paramOnDemandPriceList)
        
        Dim paramCustomLetter As SqlParameter = New SqlParameter("@CustomLetter", SqlDbType.Bit)
        paramCustomLetter.Value = 0
        oCmd.Parameters.Add(paramCustomLetter)
        
        Dim paramProductKey As SqlParameter = New SqlParameter("@ProductKey", SqlDbType.Int, 4)
        paramProductKey.Direction = ParameterDirection.Output
        oCmd.Parameters.Add(paramProductKey)
  
        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
            AddNewProduct = CLng(oCmd.Parameters("@ProductKey").Value)
        Catch ex As SqlException
            If ex.Number = 2627 Then
                SendMail(MESSAGE_TYPE_NHSPIP_WEBFORM_ERROR, EMAIL_ADDRESS_ERRORS, "Could not create NHS PIP product - product already exists", "PRODUCT KEY: " & sLogisticProductKey & "; PCT: " & sPCTAbbreviation, "PRODUCT KEY: " & sLogisticProductKey & "; PCT: " & sPCTAbbreviation)
            Else
                SendMail(MESSAGE_TYPE_NHSPIP_WEBFORM_ERROR, EMAIL_ADDRESS_ERRORS, "Could not create NHS PIP product - see reason", ex.ToString, ex.ToString)
            End If
        Finally
            oConn.Close()
        End Try
    End Function
  
    Protected Function AddNewUser(ByVal sUserId As String, ByVal sPassword As String, ByVal sFirstName As String, ByVal sLastName As String, ByVal sEmailAddr As String) As Integer
        AddNewUser = 0
        Dim bError As Boolean
        Dim oPassword As SprintInternational.Password = New SprintInternational.Password()
        Dim oConn As New SqlConnection(gsConn)
            
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_UserProfile_Add5", oConn)
        Dim oTrans As SqlTransaction
        oCmd.CommandType = CommandType.StoredProcedure

        Dim paramUserKey As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int, 4)
        paramUserKey.Value = 0
        oCmd.Parameters.Add(paramUserKey)
            
        Dim paramUserId As SqlParameter = New SqlParameter("@UserId", SqlDbType.NVarChar, 100)
        paramUserId.Value = sUserId
        oCmd.Parameters.Add(paramUserId)
            
        Dim paramPassword As SqlParameter = New SqlParameter("@Password", SqlDbType.NVarChar, 24)
        paramPassword.Value = oPassword.Encrypt(sPassword)
        oCmd.Parameters.Add(paramPassword)
            
        Dim paramFirstName As SqlParameter = New SqlParameter("@FirstName", SqlDbType.NVarChar, 50)
        paramFirstName.Value = sFirstName
        oCmd.Parameters.Add(paramFirstName)
            
        Dim paramLastName As SqlParameter = New SqlParameter("@LastName", SqlDbType.NVarChar, 50)
        paramLastName.Value = sLastName
        oCmd.Parameters.Add(paramLastName)
            
        Dim paramTitle As SqlParameter = New SqlParameter("@Title", SqlDbType.NVarChar, 50)
        paramTitle.Value = Nothing
        oCmd.Parameters.Add(paramTitle)
            
        Dim paramDepartment As SqlParameter = New SqlParameter("@Department", SqlDbType.NVarChar, 20)
        paramDepartment.Value = ddlPCTs.SelectedValue
        oCmd.Parameters.Add(paramDepartment)
            
        Dim paramUserGroup As SqlParameter = New SqlParameter("@UserGroup", SqlDbType.Int)
        paramUserGroup.Value = 0
        oCmd.Parameters.Add(paramUserGroup)

        Dim paramStatus As SqlParameter = New SqlParameter("@Status", SqlDbType.NVarChar, 20)
        paramStatus.Value = "Active"
        oCmd.Parameters.Add(paramStatus)
        
        Dim paramType As SqlParameter = New SqlParameter("@Type", SqlDbType.NVarChar, 20)
        paramType.Value = "User"
        oCmd.Parameters.Add(paramType)
        
        Dim paramCustomer As SqlParameter = New SqlParameter("@Customer", SqlDbType.Bit)
        paramCustomer.Value = 1
        oCmd.Parameters.Add(paramCustomer)
        
        Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
        paramCustomerKey.Value = gnCustomerKey
        oCmd.Parameters.Add(paramCustomerKey)
        
        Dim paramEmailAddr As SqlParameter = New SqlParameter("@EmailAddr", SqlDbType.NVarChar, 100)
        paramEmailAddr.Value = sEmailAddr
        oCmd.Parameters.Add(paramEmailAddr)
            
        Dim paramTelephone As SqlParameter = New SqlParameter("@Telephone", SqlDbType.NVarChar, 20)
        paramTelephone.Value = String.Empty
        oCmd.Parameters.Add(paramTelephone)
            
        Dim paramCollectionPoint As SqlParameter = New SqlParameter("@CollectionPoint", SqlDbType.NVarChar, 50)
        paramCollectionPoint.Value = String.Empty
        oCmd.Parameters.Add(paramCollectionPoint)
        
        Dim paramURL As SqlParameter = New SqlParameter("@URL", SqlDbType.NVarChar, 100)
        paramURL.Value = "default"
        oCmd.Parameters.Add(paramURL)

        Dim paramAbleToViewStock As SqlParameter = New SqlParameter("@AbleToViewStock", SqlDbType.Bit)
        paramAbleToViewStock.Value = 1
        oCmd.Parameters.Add(paramAbleToViewStock)

        Dim paramAbleToCreateStockBooking As SqlParameter = New SqlParameter("@AbleToCreateStockBooking", SqlDbType.Bit)
        paramAbleToCreateStockBooking.Value = 1
        oCmd.Parameters.Add(paramAbleToCreateStockBooking)

        Dim paramAbleToCreateCollectionRequest As SqlParameter = New SqlParameter("@AbleToCreateCollectionRequest", SqlDbType.Bit)
        paramAbleToCreateCollectionRequest.Value = 0
        oCmd.Parameters.Add(paramAbleToCreateCollectionRequest)
        
        Dim paramAbleToViewGlobalAddressBook As SqlParameter = New SqlParameter("@AbleToViewGlobalAddressBook", SqlDbType.Bit)
        paramAbleToViewGlobalAddressBook.Value = 0
        oCmd.Parameters.Add(paramAbleToViewGlobalAddressBook)
        
        Dim paramAbleToEditGlobalAddressBook As SqlParameter = New SqlParameter("@AbleToEditGlobalAddressBook", SqlDbType.Bit)
        paramAbleToEditGlobalAddressBook.Value = 0
        oCmd.Parameters.Add(paramAbleToEditGlobalAddressBook)
        
        Dim paramRunningHeader As SqlParameter = New SqlParameter("@RunningHeaderImage", SqlDbType.NVarChar, 100)
        paramRunningHeader.Value = "default"
        oCmd.Parameters.Add(paramRunningHeader)
        
        Dim paramStockBookingAlert As SqlParameter = New SqlParameter("@StockBookingAlert", SqlDbType.Bit)
        paramStockBookingAlert.Value = 1
        oCmd.Parameters.Add(paramStockBookingAlert)
        
        Dim paramStockBookingAlertAll As SqlParameter = New SqlParameter("@StockBookingAlertAll", SqlDbType.Bit)
        paramStockBookingAlertAll.Value = 0
        oCmd.Parameters.Add(paramStockBookingAlertAll)
        
        Dim paramStockArrivalAlert As SqlParameter = New SqlParameter("@StockArrivalAlert", SqlDbType.Bit)
        paramStockArrivalAlert.Value = 0
        oCmd.Parameters.Add(paramStockArrivalAlert)
        
        Dim paramLowStockAlert As SqlParameter = New SqlParameter("@LowStockAlert", SqlDbType.Bit)
        paramLowStockAlert.Value = 0
        oCmd.Parameters.Add(paramLowStockAlert)
        
        Dim paramCourierBookingAlert As SqlParameter = New SqlParameter("@ConsignmentBookingAlert", SqlDbType.Bit)
        paramCourierBookingAlert.Value = 0
        oCmd.Parameters.Add(paramCourierBookingAlert)
        
        Dim paramCourierBookingAlertAll As SqlParameter = New SqlParameter("@ConsignmentBookingAlertAll", SqlDbType.Bit)
        paramCourierBookingAlertAll.Value = 0
        oCmd.Parameters.Add(paramCourierBookingAlertAll)
        
        Dim paramCourierDespatchAlert As SqlParameter = New SqlParameter("@ConsignmentDespatchAlert", SqlDbType.Bit)
        paramCourierDespatchAlert.Value = 0
        oCmd.Parameters.Add(paramCourierDespatchAlert)
        
        Dim paramCourierDeliveryAlert As SqlParameter = New SqlParameter("@ConsignmentDeliveryAlert", SqlDbType.Bit)
        paramCourierDeliveryAlert.Value = 0
        oCmd.Parameters.Add(paramCourierDeliveryAlert)
        
        Dim paramKey As SqlParameter = New SqlParameter("@UserProfileKey", SqlDbType.Int, 4)
        paramKey.Direction = ParameterDirection.Output
        oCmd.Parameters.Add(paramKey)
        Try
            oConn.Open()
            oTrans = oConn.BeginTransaction(IsolationLevel.ReadCommitted, "AddRecord")
            oCmd.Connection = oConn
            oCmd.Transaction = oTrans
            oCmd.ExecuteNonQuery()
            oTrans.Commit()
            AddNewUser = CLng(oCmd.Parameters("@UserProfileKey").Value)
        Catch ex As SqlException
            bError = True
            oTrans.Rollback("AddRecord")
            If ex.Number = 2627 Then
                SendMail(MESSAGE_TYPE_NHSPIP_WEBFORM_ERROR, EMAIL_ADDRESS_ERRORS, "Could not create NHS PIP user account", "User ID " & sUserId & " already exists", "User ID " & sUserId & " already exists")
                Exit Function
            Else
                SendMail(MESSAGE_TYPE_NHSPIP_WEBFORM_ERROR, EMAIL_ADDRESS_ERRORS, "Error creating NHS PIP user account", ex.ToString, ex.ToString)
            End If
        Finally
            oConn.Close()
        End Try
    End Function

    Protected Sub btnGoBack_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call HideAllPanels()
        pnlDataEntry.Visible = True
    End Sub
    
    Protected Function GetPCTKeyFromAbbreviation(ByVal sPCTAbbreviation As String) As Integer
        GetPCTKeyFromAbbreviation = ExecuteQueryToDataTable("SELECT [id] FROM NHSPCTs WHERE PCTAbbreviation = '" & sPCTAbbreviation.Replace("'", "''") & "'").Rows(0).Item(0)
    End Function
    
    Protected Function GetPCTKeyFromName(ByVal sPCTName As String) As Integer
        If sPCTName.ToLower.Contains("connecting for health") Then
            sPCTName = "connecting for health"
        End If
        Try
            GetPCTKeyFromName = ExecuteQueryToDataTable("SELECT [id] FROM NHSPCTs WHERE PCTName = '" & sPCTName.Replace("'", "''") & "'").Rows(0).Item(0)
        Catch
            SendMail(MESSAGE_TYPE_NHSPIP_WEBFORM_ERROR, EMAIL_ADDRESS_ERRORS, "NHSPIP Webform failed in GetPCTKeyFromName, sPCTName = " & sPCTName, "NHSPIP Webform failed in GetPCTKeyFromName, sPCTName = " & sPCTName, "NHSPIP Webform failed in GetPCTKeyFromName, sPCTName = " & sPCTName)
        End Try
    End Function
    
    Protected Function CreateOrderRecord() As Integer
        Dim sbSQL As New StringBuilder
        sbSQL.Append("INSERT INTO NHSPIPWebformSubmission (FirstName, LastName, EmailAddr, NHSPCTKey, AccountName, Information, CreatedOn) VALUES (")
        
        sbSQL.Append("'")
        sbSQL.Append(tbFirstName.Text.Replace("'", "''"))
        sbSQL.Append("'")
        sbSQL.Append(",")
        
        sbSQL.Append("'")
        sbSQL.Append(tbLastName.Text.Replace("'", "''"))
        sbSQL.Append("'")
        sbSQL.Append(",")
        
        sbSQL.Append("'")
        sbSQL.Append(tbEmail.Text.Replace("'", "''"))
        sbSQL.Append("'")
        sbSQL.Append(",")
        
        sbSQL.Append(GetPCTKeyFromName(ddlPCTs.SelectedItem.Text))
        sbSQL.Append(",")
        
        Dim sAccountName As String = String.Empty
        If tbUserName.Text.Trim <> String.Empty Then
            sAccountName = tbUserName.Text.Trim
        End If
        If tbNewUserName.Text.Trim <> String.Empty Then
            sAccountName = tbNewUserName.Text.Trim
        End If
        sbSQL.Append("'")
        sbSQL.Append(sAccountName.Replace("'", "''"))
        sbSQL.Append("'")
        sbSQL.Append(",")
        
        sbSQL.Append("'")
        sbSQL.Append(tbInformation.Text.Replace("'", "''"))
        sbSQL.Append("'")
        sbSQL.Append(",")
        
        sbSQL.Append("GETDATE()")
        sbSQL.Append(")")
        sbSQL.Append(" ")
        sbSQL.Append("SELECT SCOPE_IDENTITY()")
        CreateOrderRecord = ExecuteQueryToDataTable(sbSQL.ToString).Rows(0).Item(0)
    End Function
    
    Protected Function CreateOrderDetailRecord(ByVal nOrderKey As Integer, ByVal nLogisticProductKey As Integer, ByVal nQty As Integer) As Integer
        Dim sbSQL As New StringBuilder
        sbSQL.Append("INSERT INTO NHSPIPWebformSubmissionDetail (OrderKey, LogisticProductKey, QtyRequested, QtyAvailable, ConsignmentNo) VALUES (")
        sbSQL.Append(nOrderKey)
        sbSQL.Append(",")
        sbSQL.Append(nLogisticProductKey)
        sbSQL.Append(",")
        sbSQL.Append(nQty)
        sbSQL.Append(",")
        sbSQL.Append("0")
        sbSQL.Append(",")
        sbSQL.Append("0")
        sbSQL.Append(")")
        sbSQL.Append(" ")
        sbSQL.Append("SELECT SCOPE_IDENTITY()")
        CreateOrderDetailRecord = ExecuteQueryToDataTable(sbSQL.ToString).Rows(0).Item(0)
    End Function
    
    Protected Function PermissionProduct(ByVal nLogisticProductKey As Integer, ByVal sPCTAbbreviation As String) As Integer
        Dim sSQL As String = "UPDATE UserProductProfile SET AbleToPick = 1 WHERE ProductKey = " & nLogisticProductKey & " AND UserKey IN (SELECT [key] FROM UserProfile WHERE CustomerKey = " & gnCustomerKey & " AND Department = '" & sPCTAbbreviation.Replace("'", "''") & "') SELECT @@Rowcount"
        PermissionProduct = ExecuteQueryToDataTable(sSQL).Rows(0).Item(0)
    End Function
    
    Protected Function GetQtyAvailableToPick(ByVal sProductCode As String, ByVal nLogisticProductKey As Integer, ByVal nQtyRequested As Integer, ByVal nOrderKey As Integer, ByVal nOrderDetailKey As Integer) As Integer  ' returns 0 if enough to pick, otherwise amount available; if necessary could return negation of amount picked
        Dim sbSQL1 As New StringBuilder
        sbSQL1.Append("SELECT Quantity = CASE ISNUMERIC((SELECT SUM(LogisticProductQuantity) FROM LogisticProductLocation AS lpl WHERE lpl.LogisticProductKey = lp.LogisticProductKey)) ")
        sbSQL1.Append("WHEN 0 THEN 0 ELSE (SELECT SUM(LogisticProductQuantity) FROM LogisticProductLocation AS lpl WHERE lpl.LogisticProductKey = lp.LogisticProductKey) ")
        sbSQL1.Append("END ")
        sbSQL1.Append("FROM LogisticProduct AS lp ")
        sbSQL1.Append("LEFT OUTER JOIN LogisticProductLocation AS lpl ")
        sbSQL1.Append("ON lp.LogisticProductKey = lpl.LogisticProductKey ")
        sbSQL1.Append("WHERE lp.LogisticProductKey = ")
        sbSQL1.Append(nLogisticProductKey)
        Dim nAvailableQty As Integer = ExecuteQueryToDataTable(sbSQL1.ToString).Rows(0).Item(0)
        Dim nMinimumStockLevel As Integer = ExecuteQueryToDataTable("SELECT MinimumStockLevel from LogisticProduct WHERE LogisticProductKey = " & nLogisticProductKey).Rows(0).Item(0)
        Dim nMaxTake As Integer = nAvailableQty - nMinimumStockLevel
        Call Log(LOG_ENTRY_TYPE_PRODUCT, nLogisticProductKey, nOrderKey, ddlPCTs.SelectedValue, "Available qty = " & nAvailableQty.ToString & "; Min stock Level = " & nMinimumStockLevel & "; Max Take = " & nMaxTake & "; product ~ " & sProductCode)
        If nMaxTake > 0 Then
            GetQtyAvailableToPick = nMaxTake
        Else
            GetQtyAvailableToPick = 0
        End If
        ExecuteNonQuery("UPDATE NHSPIPWebformSubmissionDetail SET QtyRequested = " & nQtyRequested & ", QtyAvailable = " & GetQtyAvailableToPick & " WHERE [id] = " & nOrderDetailKey)
    End Function
    
    Protected Function GeneratePick(ByVal nLogisticProductKey As Integer, ByVal nQty As Integer, ByVal sDestinationProduct As String) As String   ' returns consignment key (numeric, = success) or message (error, failure)
        GeneratePick = String.Empty
        Dim sSQL As String
        Dim oDataTable As DataTable
        sSQL = "SELECT ISNULL(CustomerName,''), ISNULL(CustomerAddr1,''), ISNULL(CustomerAddr2,''), ISNULL(CustomerAddr3,''), ISNULL(CustomerTown,''), ISNULL(CustomerCounty,''), ISNULL(CustomerPostCode,''), ISNULL(CustomerCountryKey,0) FROM Customer WHERE CustomerKey = 5"
        oDataTable = ExecuteQueryToDataTable(sSQL)
        Dim oDataRow As DataRow = oDataTable.Rows(0)
        
        Dim lBookingKey As Long
        Dim lConsignmentKey As Long
        Dim BookingFailed As Boolean
        Dim oConn As New SqlConnection(gsConn)
        Dim oTrans As SqlTransaction
        Dim oCmdAddBooking As SqlCommand = New SqlCommand("spASPNET_StockBooking_Add3", oConn)
        oCmdAddBooking.CommandType = CommandType.StoredProcedure

        Dim param1 As SqlParameter = New SqlParameter("@UserProfileKey", SqlDbType.Int, 4)
        param1.Value = GetWebformUserKey()
        oCmdAddBooking.Parameters.Add(param1)
        
        Dim param2 As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
        param2.Value = gnCustomerKey
        oCmdAddBooking.Parameters.Add(param2)

        Dim param2a As SqlParameter = New SqlParameter("@BookingOrigin", SqlDbType.NVarChar, 20)
        param2a.Value = "WEB_BOOKING"
        oCmdAddBooking.Parameters.Add(param2a)

        Dim param3 As SqlParameter = New SqlParameter("@BookingReference1", SqlDbType.NVarChar, 25)
        param3.Value = String.Empty
        oCmdAddBooking.Parameters.Add(param3)

        Dim param4 As SqlParameter = New SqlParameter("@BookingReference2", SqlDbType.NVarChar, 25)
        param4.Value = String.Empty
        oCmdAddBooking.Parameters.Add(param4)

        Dim param5 As SqlParameter = New SqlParameter("@BookingReference3", SqlDbType.NVarChar, 50)
        param5.Value = String.Empty
        oCmdAddBooking.Parameters.Add(param5)

        Dim param6 As SqlParameter = New SqlParameter("@BookingReference4", SqlDbType.NVarChar, 50)
        param6.Value = String.Empty
        oCmdAddBooking.Parameters.Add(param6)
            
        Dim param6a As SqlParameter = New SqlParameter("@ExternalReference", SqlDbType.NVarChar, 50)
        param6a.Value = Nothing
        oCmdAddBooking.Parameters.Add(param6a)
            
        Dim param7 As SqlParameter = New SqlParameter("@SpecialInstructions", SqlDbType.NVarChar, 1000)
        param7.Value = "AUTOMATIC PICK: Please transfer picked items to NHS PIP product " & sDestinationProduct
        oCmdAddBooking.Parameters.Add(param7)

        Dim param8 As SqlParameter = New SqlParameter("@PackingNoteInfo", SqlDbType.NVarChar, 1000)
        param8.Value = String.Empty
        oCmdAddBooking.Parameters.Add(param8)

        Dim param9 As SqlParameter = New SqlParameter("@ConsignmentType", SqlDbType.NVarChar, 20)
        param9.Value = "STOCK ITEM"
        oCmdAddBooking.Parameters.Add(param9)

        Dim param10 As SqlParameter = New SqlParameter("@ServiceLevelKey", SqlDbType.Int, 4)
        param10.Value = -1
        oCmdAddBooking.Parameters.Add(param10)

        Dim param11 As SqlParameter = New SqlParameter("@Description", SqlDbType.NVarChar, 250)
        param11.Value = "INTERNAL TRANSFER"
        oCmdAddBooking.Parameters.Add(param11)

        Dim param13 As SqlParameter = New SqlParameter("@CnorName", SqlDbType.NVarChar, 50)
        param13.Value = oDataRow(0)
        oCmdAddBooking.Parameters.Add(param13)
        
        Dim param14 As SqlParameter = New SqlParameter("@CnorAddr1", SqlDbType.NVarChar, 50)
        param14.Value = oDataRow(1)
        oCmdAddBooking.Parameters.Add(param14)
        
        Dim param15 As SqlParameter = New SqlParameter("@CnorAddr2", SqlDbType.NVarChar, 50)
        param15.Value = oDataRow(2)
        
        oCmdAddBooking.Parameters.Add(param15)
        Dim param16 As SqlParameter = New SqlParameter("@CnorAddr3", SqlDbType.NVarChar, 50)
        param16.Value = oDataRow(3)
        
        oCmdAddBooking.Parameters.Add(param16)
        Dim param17 As SqlParameter = New SqlParameter("@CnorTown", SqlDbType.NVarChar, 50)
        param17.Value = oDataRow(4)
        
        oCmdAddBooking.Parameters.Add(param17)
        Dim param18 As SqlParameter = New SqlParameter("@CnorState", SqlDbType.NVarChar, 50)
        param18.Value = oDataRow(5)
        
        oCmdAddBooking.Parameters.Add(param18)
        Dim param19 As SqlParameter = New SqlParameter("@CnorPostCode", SqlDbType.NVarChar, 50)
        param19.Value = oDataRow(6)
        
        oCmdAddBooking.Parameters.Add(param19)
        Dim param20 As SqlParameter = New SqlParameter("@CnorCountryKey", SqlDbType.Int, 4)
        param20.Value = oDataRow(7)
        
        oCmdAddBooking.Parameters.Add(param20)
        Dim param21 As SqlParameter = New SqlParameter("@CnorCtcName", SqlDbType.NVarChar, 50)
        param21.Value = "Marilyn Quinn X506"
        oCmdAddBooking.Parameters.Add(param21)
        
        Dim param22 As SqlParameter = New SqlParameter("@CnorTel", SqlDbType.NVarChar, 50)
        param22.Value = "020 8751 1111"
        oCmdAddBooking.Parameters.Add(param22)

        Dim param23 As SqlParameter = New SqlParameter("@CnorEmail", SqlDbType.NVarChar, 50)
        param23.Value = "m.quinn@sprintexpress.co.uk"
        oCmdAddBooking.Parameters.Add(param23)
        
        Dim param24 As SqlParameter = New SqlParameter("@CnorPreAlertFlag", SqlDbType.Bit)
        param24.Value = 0
        oCmdAddBooking.Parameters.Add(param24)
        
        Dim param25 As SqlParameter = New SqlParameter("@CneeName", SqlDbType.NVarChar, 50)
        param25.Value = oDataRow(0)
        oCmdAddBooking.Parameters.Add(param25)
        
        Dim param26 As SqlParameter = New SqlParameter("@CneeAddr1", SqlDbType.NVarChar, 50)
        param26.Value = oDataRow(1)
        oCmdAddBooking.Parameters.Add(param26)
        
        Dim param27 As SqlParameter = New SqlParameter("@CneeAddr2", SqlDbType.NVarChar, 50)
        param27.Value = oDataRow(2)
        oCmdAddBooking.Parameters.Add(param27)
        
        Dim param28 As SqlParameter = New SqlParameter("@CneeAddr3", SqlDbType.NVarChar, 50)
        param28.Value = oDataRow(3)
        oCmdAddBooking.Parameters.Add(param28)
        
        Dim param29 As SqlParameter = New SqlParameter("@CneeTown", SqlDbType.NVarChar, 50)
        param29.Value = oDataRow(4)
        oCmdAddBooking.Parameters.Add(param29)
        
        Dim param30 As SqlParameter = New SqlParameter("@CneeState", SqlDbType.NVarChar, 50)
        param30.Value = oDataRow(5)
        oCmdAddBooking.Parameters.Add(param30)
        
        Dim param31 As SqlParameter = New SqlParameter("@CneePostCode", SqlDbType.NVarChar, 50)
        param31.Value = oDataRow(6)
        oCmdAddBooking.Parameters.Add(param31)
        
        Dim param32 As SqlParameter = New SqlParameter("@CneeCountryKey", SqlDbType.Int, 4)
        param32.Value = oDataRow(7)
        oCmdAddBooking.Parameters.Add(param32)
        
        Dim param33 As SqlParameter = New SqlParameter("@CneeCtcName", SqlDbType.NVarChar, 50)
        param33.Value = "Marilyn Quinn X506"
        oCmdAddBooking.Parameters.Add(param33)
        
        Dim param34 As SqlParameter = New SqlParameter("@CneeTel", SqlDbType.NVarChar, 50)
        param34.Value = String.Empty
        oCmdAddBooking.Parameters.Add(param34)
        
        Dim param35 As SqlParameter = New SqlParameter("@CneeEmail", SqlDbType.NVarChar, 50)
        param35.Value = String.Empty
        oCmdAddBooking.Parameters.Add(param35)
        
        Dim param36 As SqlParameter = New SqlParameter("@CneePreAlertFlag", SqlDbType.Bit)
        param36.Value = 0
        oCmdAddBooking.Parameters.Add(param36)
        
        Dim param37 As SqlParameter = New SqlParameter("@LogisticBookingKey", SqlDbType.Int, 4)
        param37.Direction = ParameterDirection.Output
        oCmdAddBooking.Parameters.Add(param37)
        
        Dim param38 As SqlParameter = New SqlParameter("@ConsignmentKey", SqlDbType.Int, 4)
        param38.Direction = ParameterDirection.Output
        oCmdAddBooking.Parameters.Add(param38)
        
        Try
            BookingFailed = False
            oConn.Open()
            oTrans = oConn.BeginTransaction(IsolationLevel.ReadCommitted, "AddBooking")
            oCmdAddBooking.Connection = oConn
            oCmdAddBooking.Transaction = oTrans
            oCmdAddBooking.ExecuteNonQuery()
            lBookingKey = CLng(oCmdAddBooking.Parameters("@LogisticBookingKey").Value.ToString)
            lConsignmentKey = CLng(oCmdAddBooking.Parameters("@ConsignmentKey").Value.ToString)
            If lBookingKey > 0 Then
                Dim oCmdAddStockItem As SqlCommand = New SqlCommand("spASPNET_LogisticMovement_Add", oConn)
                oCmdAddStockItem.CommandType = CommandType.StoredProcedure
                
                Dim param51 As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int, 4)
                param51.Value = CLng(GetWebformUserKey())
                oCmdAddStockItem.Parameters.Add(param51)
                
                Dim param52 As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
                param52.Value = CLng(gnCustomerKey)
                oCmdAddStockItem.Parameters.Add(param52)
                
                Dim param53 As SqlParameter = New SqlParameter("@LogisticBookingKey", SqlDbType.Int, 4)
                param53.Value = lBookingKey
                oCmdAddStockItem.Parameters.Add(param53)
                
                Dim param54 As SqlParameter = New SqlParameter("@LogisticProductKey", SqlDbType.Int, 4)
                param54.Value = nLogisticProductKey
                oCmdAddStockItem.Parameters.Add(param54)
                
                Dim param55 As SqlParameter = New SqlParameter("@LogisticMovementStateId", SqlDbType.NVarChar, 20)
                param55.Value = "PENDING"
                oCmdAddStockItem.Parameters.Add(param55)
                
                Dim param56 As SqlParameter = New SqlParameter("@ItemsOut", SqlDbType.Int, 4)
                param56.Value = nQty
                oCmdAddStockItem.Parameters.Add(param56)
                
                Dim param57 As SqlParameter = New SqlParameter("@ConsignmentKey", SqlDbType.Int, 8)
                param57.Value = lConsignmentKey
                oCmdAddStockItem.Parameters.Add(param57)
                
                oCmdAddStockItem.Connection = oConn
                oCmdAddStockItem.Transaction = oTrans
                oCmdAddStockItem.ExecuteNonQuery()

                Dim oCmdCompleteBooking As SqlCommand = New SqlCommand("spASPNET_LogisticBooking_Complete", oConn)
                oCmdCompleteBooking.CommandType = CommandType.StoredProcedure
                
                Dim param71 As SqlParameter = New SqlParameter("@LogisticBookingKey", SqlDbType.Int, 4)
                param71.Value = lBookingKey
                oCmdCompleteBooking.Parameters.Add(param71)
                
                oCmdCompleteBooking.Connection = oConn
                oCmdCompleteBooking.Transaction = oTrans
                oCmdCompleteBooking.ExecuteNonQuery()
            Else
                BookingFailed = True
                GeneratePick = "Zero booking key returned"
            End If
            If Not BookingFailed Then
                oTrans.Commit()
                GeneratePick = lConsignmentKey.ToString
            Else
                oTrans.Rollback("AddBooking")
            End If
        Catch ex As SqlException
            GeneratePick = "-> " & ex.ToString
            oTrans.Rollback("AddBooking")
        Finally
            oConn.Close()
        End Try
    End Function

    Protected Function ProcessProducts() As Integer
        Dim nOrderKey As Integer = CreateOrderRecord()
        ProcessProducts = nOrderKey
        For Each gvr As GridViewRow In gvProductInfo.Rows
            If gvr.RowType = DataControlRowType.DataRow Then
                Dim tcProductCode As TableCell = gvr.Cells(0)
                Dim tcMaxGrab As TableCell = gvr.Cells(2)
                Dim tb As TextBox = gvr.FindControl("tbQty")
                If tb.Text.Trim <> String.Empty Then
                    Dim nQtyRequested As Integer = CInt(tb.Text)
                    If nQtyRequested > 0 Then
                        Dim hidBaseProductKey As HiddenField = gvr.FindControl("hidLogisticProductKey")
                        Dim nOrderDetailKey = CreateOrderDetailRecord(nOrderKey, hidBaseProductKey.Value, nQtyRequested)
                        If gbIsCreatingProductsAndAccounts Then
                            If Not ProductExists(tcProductCode.Text, ddlPCTs.SelectedValue) Then
                                Call Log(LOG_ENTRY_TYPE_PRODUCT, hidBaseProductKey.Value, nOrderKey, ddlPCTs.SelectedValue, "Cloning product ~ " & tcProductCode.Text & " for " & ddlPCTs.SelectedValue)
                                Dim nLinkedProductKey As Integer = AddNewProduct(hidBaseProductKey.Value, ddlPCTs.SelectedValue)
                                Call Log(LOG_ENTRY_TYPE_PRODUCT, nLinkedProductKey, nOrderKey, ddlPCTs.SelectedValue, "Created product ~ " & tcProductCode.Text & " | " & ddlPCTs.SelectedValue)
                                Call Log(LOG_ENTRY_TYPE_PRODUCT, nLinkedProductKey, nOrderKey, ddlPCTs.SelectedValue, "Permissioned " & PermissionProduct(nLinkedProductKey, ddlPCTs.SelectedValue) & " user(s) for product ~ " & tcProductCode.Text & " | " & ddlPCTs.SelectedValue)
                                Dim nQtyAvailable As Integer = GetQtyAvailableToPick(tcProductCode.Text, hidBaseProductKey.Value, nQtyRequested, nOrderKey, nOrderDetailKey)
                                Dim nBacklogQty As Integer = 0
                                If nQtyAvailable > 0 Then
                                    Dim nQtyToPick As Integer = nQtyRequested
                                    If nQtyAvailable < nQtyRequested Then
                                        nQtyToPick = nQtyAvailable
                                        nBacklogQty = nQtyRequested - nQtyAvailable
                                    Else
                                        nQtyToPick = nQtyRequested
                                    End If
                                    Dim sConsignmentNo As String = GeneratePick(hidBaseProductKey.Value, nQtyToPick, "(" & nLinkedProductKey.ToString & ") Product code: " & tcProductCode.Text & " Product Date: " & ddlPCTs.SelectedValue)
                                    If IsNumeric(sConsignmentNo) Then
                                        ExecuteNonQuery("UPDATE NHSPIPWebformSubmissionDetail SET ConsignmentNo = " & sConsignmentNo & " WHERE [id] = " & nOrderDetailKey)
                                        Call Log(LOG_ENTRY_TYPE_PRODUCT, hidBaseProductKey.Value, nOrderKey, ddlPCTs.SelectedValue, "Pick " & sConsignmentNo & " generated for " & nQtyRequested.ToString & " of ~ " & tcProductCode.Text & " | " & ddlPCTs.SelectedValue)
                                        If nQtyRequested > nQtyAvailable Then
                                            Dim sMessage As String = "Insufficient pick quantity available (" & nQtyAvailable & ") for new ring-fenced product (requested " & nQtyRequested & ") of ~ " & tcProductCode.Text & " | " & ddlPCTs.SelectedValue & " Order: " & nOrderKey & "/" & nOrderDetailKey.ToString
                                            Call Log(LOG_ENTRY_TYPE_PRODUCT, hidBaseProductKey.Value, nOrderKey, ddlPCTs.SelectedValue, sMessage)
                                            Call SendMail(MESSAGE_TYPE_NHSPIP_ALERT, EMAIL_ADDRESS_ALERTS, "NHS PIP Alert - insufficient pick quantity available - please see log", sMessage, sMessage)
                                        End If
                                    Else
                                        Dim sMessage As String = "Attempt to pick failed for ~ " & tcProductCode.Text & " | " & ddlPCTs.SelectedValue & " Consignment value returned: " & sConsignmentNo & " Order: " & nOrderKey & "/" & nOrderDetailKey.ToString
                                        Call Log(LOG_ENTRY_TYPE_PRODUCT, hidBaseProductKey.Value, nOrderKey, ddlPCTs.SelectedValue, sMessage)
                                        SendMail(MESSAGE_TYPE_NHSPIP_WEBFORM_ERROR, EMAIL_ADDRESS_SUPPORT, "NHS PIP Alert - attempt to pick failed - please see log", sMessage, sMessage)
                                    End If
                                Else
                                    nBacklogQty = nQtyRequested
                                    Dim sMessage As String = "No pick quantity available for new ring-fenced product (requested " & nQtyRequested & ") of ~ " & tcProductCode.Text & " | " & ddlPCTs.SelectedValue & " Order: " & nOrderKey & "/" & nOrderDetailKey.ToString
                                    Call Log(LOG_ENTRY_TYPE_PRODUCT, hidBaseProductKey.Value, nOrderKey, ddlPCTs.SelectedValue, sMessage)
                                    Call SendMail(MESSAGE_TYPE_NHSPIP_ALERT, EMAIL_ADDRESS_ALERTS, "NHS PIP Alert - no pick quantity available - please see log", sMessage, sMessage)
                                End If
                                Call ExecuteQueryToDataTable("INSERT INTO NHSPIPLinkedProducts (GenericProductKey, RingFencedProductKey, BacklogQty) VALUES (" & hidBaseProductKey.Value & "," & nLinkedProductKey & ", " & nBacklogQty & ")")
                            Else
                                Dim sMessage As String = "Ring-fenced product exists ~ " & tcProductCode.Text & " | " & ddlPCTs.SelectedValue & " (no pick attempted. Order: " & nOrderKey & "/" & nOrderDetailKey.ToString
                                Call Log(LOG_ENTRY_TYPE_PRODUCT, hidBaseProductKey.Value, nOrderKey, ddlPCTs.SelectedValue, sMessage)
                                Call SendMail(MESSAGE_TYPE_NHSPIP_ALERT, EMAIL_ADDRESS_ALERTS, "NHS PIP Alert - ring-fenced product exists - please see log", sMessage, sMessage)
                            End If
                        End If
                    End If
                End If
            End If
        Next
    End Function
    
    Protected Function ProductExists(ByVal sProductCode As String, ByVal sProductDate As String) As Boolean
        Dim sSQL As String = "SELECT ProductCode FROM LogisticProduct WHERE CustomerKey = " & gnCustomerKey & " AND ProductCode = '" & sProductCode.Replace("'", "''") & "' AND ProductDate = '" & sProductDate.Replace("'", "''") & "'"
        If ExecuteQueryToDataTable(sSQL).Rows.Count = 0 Then
            ProductExists = False
        Else
            ProductExists = True
        End If
    End Function
    
    Protected Sub ProcessOrder()
        Call Log(LOG_ENTRY_TYPE_ORDER, -1, -1, ddlPCTs.SelectedValue, "Received submission from " & tbFirstName.Text & " " & tbLastName.Text & " (" & tbEmail.Text & ") - " & ddlPCTs.SelectedValue)
        If rbRequireAccount.Checked AndAlso gbIsCreatingProductsAndAccounts Then
            If AddNewUser(tbNewUserName.Text, tbNewPassword.Text, tbFirstName.Text, tbLastName.Text, tbEmail.Text) > 0 Then
                SendMail(MESSAGE_TYPE_NHSPIP_NEW_ACCOUNT, tbEmail.Text & "," & EMAIL_ADDRESS_ALERTS, "NHS PIP Account Created", BuildAccountCreatedConfirmation("Plain"), BuildAccountCreatedConfirmation("HTML"))
                Call Log(LOG_ENTRY_TYPE_ACCOUNT, -1, -1, ddlPCTs.SelectedValue, "Created account " & tbNewUserName.Text)
            Else
                SendMail(MESSAGE_TYPE_NHSPIP_WEBFORM_ERROR, EMAIL_ADDRESS_ERRORS, "Failed to create new NHS PIP account " & tbNewUserName.Text, "Failed to create new NHS PIP account " & tbNewUserName.Text, "Failed to create new NHS PIP account " & tbNewUserName.Text)
                Call Log(LOG_ENTRY_TYPE_ACCOUNT, -1, -1, ddlPCTs.SelectedValue, "Failed to create account " & tbNewUserName.Text)
            End If
        End If
        Dim nOrderNo As Integer = ProcessProducts()
        SendMail(MESSAGE_TYPE_NHSPIP_CONFIRM_ORDER, tbEmail.Text & "," & EMAIL_ADDRESS_ALERTS, "SCR materials order confirmation " & Date.Now.ToString("dd-MMM-yyyy hh:mm") & " REF: #" & nOrderNo.ToString, BuildOrderEmailConfirmation("Plain", nOrderNo), BuildOrderEmailConfirmation("HTML", nOrderNo))
    End Sub
    
    Protected Sub btnConfirm_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ProcessOrder()
        Call ShowOrderComplete()
    End Sub
    
    Protected Sub Log(ByVal sType As String, ByVal nLogisticProductKey As Integer, ByVal nOrderNo As Integer, ByVal sPCTAbbreviation As String, ByVal sLogEntry As String)
        Dim sbSQL As New StringBuilder
        sbSQL.Append("INSERT INTO NHSPIPActivityLog (Type, LogisticProductKey, OrderNo, PCTAbbreviation, LogEntry, CreatedOn) VALUES (")
        
        sbSQL.Append("'")
        sbSQL.Append(sType)
        sbSQL.Append("'")
        sbSQL.Append(",")
        
        sbSQL.Append(nLogisticProductKey)
        sbSQL.Append(",")

        sbSQL.Append(nOrderNo)
        sbSQL.Append(",")

        sbSQL.Append("'")
        sbSQL.Append(sPCTAbbreviation.Replace("'", "''"))
        sbSQL.Append("'")
        sbSQL.Append(",")
        
        sbSQL.Append("'")
        sbSQL.Append(sLogEntry.Replace("''", "'"))
        sbSQL.Append("'")
        sbSQL.Append(",")
        
        sbSQL.Append("GETDATE()")
        sbSQL.Append(")")
        Call ExecuteNonQuery(sbSQL.ToString)
    End Sub
    
    Protected Sub lnkbtnPasswordHelp_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        WebMsgBox.Show("You must set a strong password - minimum 6 characters, maximum 12 characters, at least one uppercase character, at least one digit, no repeated characters. Eg Qwert9")
    End Sub
    
    Protected Sub gvProductInfo_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs)
        Dim gvr As GridViewRow = e.Row
        If gvr.RowType = DataControlRowType.DataRow Then
            Dim tc As TableCell = gvr.Cells(0)
            If tc.Text.Contains("/") Then
                Dim sPrunedText = tc.Text.Trim
                Dim nTruncationPoint As Integer = sPrunedText.LastIndexOf(" ")
                If nTruncationPoint > 2 Then
                    sPrunedText = sPrunedText.Substring(0, nTruncationPoint)
                    tc.Text = sPrunedText
                End If
            End If
        End If
    End Sub
    
    Protected Sub ddlOrganisationType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        Call InitPCTList(ddl.SelectedValue)
        If ddl.Items(0).Value = "0" Then
            ddl.Items.RemoveAt(0)
        End If
        divOrganisation.Visible = True
        ddlPCTs.Focus()
    End Sub
    
    Protected Sub ddlPCTs_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        If ddl.Items(0).Value = "0" Then
            ddl.Items.RemoveAt(0)
        End If
    End Sub

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
            WebMsgBox.Show("Error in ExecuteQueryToListItemCollection: " & ex.Message)
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
   
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <link href="NHSPIP.css" rel="stylesheet" type="text/css" />
    <title>NHS Connecting for Health Summary Care Records Public Information Programme -
        Order Materials - v0.90</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:Panel ID="pnlIntro" runat="server" Width="100%">
            <table style="width: 100%">
                <tr>
                    <td style="width: 5%">
                        &nbsp;
                    </td>
                    <td style="width: 90%" colspan="2">
                        &nbsp;
                    </td>
                    <td style="width: 5%">
                        &nbsp;
                    </td>
                </tr>
                <tr>
                    <td />
                    <td>
                        <p>
                            <asp:Label ID="Label16" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="Medium"
                                Text="NHS Connecting for Health Summary Care Records Public Information Programme" />
                        </p>
                        <p>
                            <b><font face="Calibri" size="3">Welcome to the Summary Care Records materials ordering
                                portal</font></b></p>
                    </td>
                    <td align="right">
                        <img alt="" src="images/NHS-RGB.jpg" />
                    </td>
                    <td />
                </tr>
                <tr>
                    <td>
                    </td>
                    <td colspan="2">
                        <p>
                            &nbsp;</p>
                        <p>
                            <font face="Calibri" size="3">This portal will enable you to order materials to support
                                your Summary Care Records Public Information Programme and engagement events.</font></p>
                        <p>
                            <font face="Calibri" size="3">When you have completed the order form and submitted it
                                you will receive an email confirmation of your order. All orders placed are subject
                                to a four week lead time. </font>
                        </p>
                        <p>
                            <font face="Calibri" size="3">Once orders are placed the stock will be allocated and
                                made available to be drawn down for delivery when you need the material.</font></p>
                        <p>
                            <font face="Calibri" size="3">You will require a user ID and log in to access the system
                                to drawn down the materials you have ordered. If you do not have a user ID and log
                                in, complete the section Access to Your Stock Online. A user ID and log in will
                                then be emailed to you.</font></p>
                        <p>
                            <font face="Calibri" size="3">There are maximum order limits on all items. If you require
                                more materials you will have to re-order when you have drawn down the majority of
                                your materials.</font></p>
                        <p>
                            <font face="Calibri" size="3">If you have any queries regarding the materials ordering
                                process please email </font><a href="mailto:scr.comms@nhs.net" target="_blank"><font
                                    color="#0000ff" face="Calibri" size="3"><u>scr.comms@nhs.net</u></font></a>&nbsp;</p>
                        <p>
                            &nbsp;</p>
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td />
                    <td colspan="2">
                    </td>
                    <td />
                </tr>
                <tr>
                    <td />
                    <td colspan="2">
                        <asp:Button ID="btnMaterialsReservation" runat="server" Text="PROCEED TO MATERIALS ORDERING"
                            Width="308px" OnClick="btnMaterialsReservation_Click" />
                    </td>
                    <td />
                </tr>
            </table>
            <br />
            <br />
        </asp:Panel>
        <asp:Panel ID="pnlDataEntry" runat="server" Width="100%" Visible="false">
            <table style="width: 100%">
                <tr>
                    <td style="width: 5%">
                        &nbsp;
                    </td>
                    <td style="width: 90%" colspan="2">
                        &nbsp;
                    </td>
                    <td style="width: 5%">
                        &nbsp;
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td>
                        <asp:Label ID="Label14" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="Medium"
                            Text="NHS Connecting for Health Summary Care Records Public Information Programme" />
                        <br />
                        <br />
                        <b><font face="Calibri" size="3">Order Materials<br />
                        </font></b>
                    </td>
                    <td align="right">
                        <img alt="" src="images/NHS-RGB.jpg" />
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td colspan="2">
                        <span class="toptext">1.&nbsp; Enter your contact details and account choice in the
                            highlighted fields below.<br />
                        </span><span class="toptext">2.&nbsp; For each product you want to order, entry the
                            quantity you require.<br />
                        </span><span class="toptext">3.&nbsp; Check your order by clicking the&nbsp;<asp:Button
                            ID="btnFinalCheckTop" runat="server" OnClick="btnFinalCheck_Click" Text="CHECK ORDER"
                            Width="113px" />
                            &nbsp;button.<br />
                        </span>
                        <table style="width: 95%">
                            <tr>
                                <td style="width: 30%">
                                    &nbsp;
                                </td>
                                <td style="width: 70%">
                                    &nbsp;
                                </td>
                            </tr>
                            <tr>
                                <td colspan="2">
                                    <fieldset>
                                        <legend>&nbsp;Your contact details&nbsp;</legend>
                                        <div class="fieldset_interior">
                                            <div class="field_interior">
                                                <asp:Label ID="Label13" CssClass="fieldlabel_contactdetails" runat="server" Text="Your first name:"
                                                    ForeColor="Red" />
                                                <asp:TextBox ID="tbFirstName" CssClass="field_contactdetails" runat="server" MaxLength="100"
                                                    BackColor="#FFFFee" />
                                            </div>
                                            <div class="field_interior">
                                                <asp:Label ID="Label12" CssClass="fieldlabel_contactdetails" runat="server" Text="Your surname:"
                                                    ForeColor="Red" />
                                                <asp:TextBox ID="tbLastName" CssClass="field_contactdetails" runat="server" MaxLength="100"
                                                    BackColor="#FFFFee" />
                                            </div>
                                            <div class="field_interior">
                                                <asp:Label ID="Label4" CssClass="fieldlabel_contactdetails" runat="server" Text="Your email address:"
                                                    ForeColor="Red" />
                                                <asp:TextBox ID="tbEmail" CssClass="field_contactdetails" runat="server" MaxLength="100"
                                                    BackColor="#FFFFee" />
                                            </div>
                                            <div class="field_interior">
                                                <asp:Label ID="Label5" CssClass="fieldlabel_contactdetails" runat="server" Text="Verify your email address:"
                                                    ForeColor="Red" />
                                                <asp:TextBox ID="tbEmailConfirm" CssClass="field_contactdetails" runat="server" MaxLength="100"
                                                    BackColor="#FFFFee" />
                                            </div>
                                            <div class="field_interior">
                                                <asp:Label ID="Label6a" CssClass="fieldlabel_contactdetails" runat="server" Text="Your organisation type:"
                                                    ForeColor="Red" />
                                                <asp:DropDownList ID="ddlOrganisationType" CssClass="field_contactdetails" 
                                                    runat="server" BackColor="#FFFFee" AutoPostBack="True" OnSelectedIndexChanged="ddlOrganisationType_SelectedIndexChanged" >
                                                    <asp:ListItem Value="0">Select your organisation type</asp:ListItem>
                                                    <asp:ListItem Value="CARE">NHS Care Trust</asp:ListItem>
                                                    <asp:ListItem Value="PCT">Primary Care Trust</asp:ListItem>
                                                    <asp:ListItem Value="RO">Regional Office</asp:ListItem>
                                                    <asp:ListItem Value="SHA">Strategic Health Authority</asp:ListItem>
                                                    <asp:ListItem Value="TRUST">Trust</asp:ListItem>
                                                </asp:DropDownList>
                                            </div>
                                            <div id="divOrganisation" class="field_interior" runat="server" visible="false">
                                                <asp:Label ID="lblLegendOrganisation" CssClass="fieldlabel_contactdetails" runat="server" Text="Your organisation:" ForeColor="Red" />
                                                <asp:DropDownList ID="ddlPCTs" runat="server" BackColor="#FFFFee" AutoPostBack="True" OnSelectedIndexChanged="ddlPCTs_SelectedIndexChanged" />
                                            </div>
                                        </div>
                                    </fieldset>
                                </td>
                            </tr>
                            <tr>
                                <td colspan="2">
                                    <br />
                                </td>
                            </tr>
                            <tr>
                                <td colspan="2">
                                    <fieldset>
                                        <legend>&nbsp;Your draw down account&nbsp;</legend>
                                        <div class="fieldset_interior">
                                            To access your order once it has been processed and allocated, you will require
                                            a user ID and password so that you can draw down materials as and when required.
                                            Please select one of the three options below. If you are not responsible for drawing
                                            down materials, select the <b>No account action required at present</b> option.
                                            If you request an account we will send an email confirming your details.&nbsp; Request
                                            additional accounts by email to <a href="mailto:scr.comms@nhs.net">scr.comms@nhs.net</a>.
                                            <br />
                                            <br />
                                            <div style="text-align: center">
                                            &nbsp;&nbsp;<asp:RadioButton ID="rbHaveAccount" runat="server" AutoPostBack="True"
                                                GroupName="Account" OnCheckedChanged="rbHaveAccount_CheckedChanged" Text="I already have an account"
                                                ForeColor="Red" />
                                            &nbsp;&nbsp;
                                            <asp:RadioButton ID="rbRequireAccount" runat="server" Text="I require an account"
                                                AutoPostBack="True" OnCheckedChanged="rbRequireAccount_CheckedChanged" GroupName="Account"
                                                ForeColor="Red" />
                                            &nbsp;&nbsp;
                                            <asp:RadioButton ID="rbAccountNoActionRequired" runat="server" Text="No account action required at present"
                                                AutoPostBack="True" OnCheckedChanged="rbAccountNoActionRequired_CheckedChanged"
                                                GroupName="Account" ForeColor="Red" />
                                            </div>
                                        </div>
                                    </fieldset>
                                </td>
                            </tr>
                            <tr>
                                <td />
                                <td />
                            </tr>
                            <tr id="trVerifyAccount" runat="server" visible="false">
                                <td colspan="2">
                                    <fieldset>
                                        <legend>&nbsp;Verify your account&nbsp;</legend>
                                        <div class="fieldset_interior">
                                            <asp:Label ID="Label8" CssClass="fieldlabel_contactdetails" runat="server" Text="My user name is:" />
                                            <asp:TextBox ID="tbUserName" CssClass="field_contactdetails" runat="server" MaxLength="100" />
                                            &nbsp;
                                            <asp:LinkButton ID="lnkbtnCheckAccountExists" runat="server" OnClick="lnkbtnCheckAccountExists_Click">verify your account</asp:LinkButton>
                                        </div>
                                    </fieldset>
                                </td>
                            </tr>
                            <tr id="trCreateAccount" runat="server" visible="false">
                                <td colspan="2">
                                    <fieldset>
                                        <legend>&nbsp;Create a new materials ordering account&nbsp;</legend>
                                        <div class="fieldset_interior">
                                            <div class="field_interior">
                                                <asp:RadioButton ID="rbUseEmailAddress" runat="server" CssClass="fieldlabel_contactdetails_nofloat"
                                                    Checked="True" Text="Use my email address as my user name (recommended)" AutoPostBack="True"
                                                    OnCheckedChanged="rbUseEmailAddress_CheckedChanged" GroupName="UserName" />
                                            </div>
                                            <div class="field_interior">
                                                <asp:RadioButton ID="rbSetUserName" runat="server" CssClass="fieldlabel_contactdetails"
                                                    Text="Set my user name to:" AutoPostBack="True" OnCheckedChanged="rbSetUserName_CheckedChanged"
                                                    GroupName="UserName" />
                                            </div>
                                            <asp:TextBox ID="tbNewUserName" runat="server" CssClass="field_contactdetails" Enabled="False"
                                                MaxLength="100" />
                                            &nbsp;
                                            <asp:LinkButton ID="lnkbtnCheckUserNameAvailable" runat="server" OnClick="lnkbtnCheckUserNameAvailable_Click">check this user name is available</asp:LinkButton>
                                            <br />
                                            <br />
                                            <div class="field_interior">
                                                <asp:Label ID="Label11" runat="server" CssClass="fieldlabel_contactdetails" Text="Set my password to:" />
                                                <asp:TextBox ID="tbNewPassword" runat="server" CssClass="field_contactdetails" MaxLength="12" />
                                                &nbsp;
                                                <asp:LinkButton ID="lnkbtnPasswordHelp" runat="server" OnClick="lnkbtnPasswordHelp_Click">password help</asp:LinkButton>
                                            </div>
                                        </div>
                                    </fieldset>
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    &nbsp;
                                </td>
                                <td>
                                    &nbsp;
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    <asp:Label ID="Label7" runat="server" Font-Bold="True" Text="Products available" />
                                </td>
                                <td>
                                </td>
                            </tr>
                        </table>
                        <br />
                        <asp:GridView ID="gvProductInfo" runat="server" AutoGenerateColumns="False" CellPadding="3"
                            Font-Names="Verdana" Font-Size="XX-Small" Width="95%" OnRowDataBound="gvProductInfo_RowDataBound">
                            <Columns>
                                <asp:BoundField DataField="ProductCode" HeaderText="Product" ReadOnly="True" SortExpression="ProductName" />
                                <asp:BoundField DataField="ProductDescription" HeaderText="Description" ReadOnly="True"
                                    SortExpression="ProductDescription" />
                                <asp:BoundField DataField="MaxGrabQty" HeaderText="Max Allowed" SortExpression="MaxGrabQty" />
                                <asp:TemplateField HeaderText="Qty Required">
                                    <ItemTemplate>
                                        <asp:HiddenField ID="hidLogisticProductKey" runat="server" Value='<%# Container.DataItem("LogisticProductKey")%>' />
                                        <asp:TextBox ID="tbQty" runat="server" Font-Names="Verdana" 
                                            Font-Size="XX-Small" Width="40px" BackColor="#FFFF99" />
                                        &nbsp;<asp:RangeValidator ID="rvQtyRequired" runat="server" ControlToValidate="tbQty" ErrorMessage="<b>!!!!</b>" MaximumValue="100000" MinimumValue="0" Type="Integer" />
                                    </ItemTemplate>
                                    <ItemStyle HorizontalAlign="Center" />
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Spare" Visible="false">
                                    <ItemTemplate>
                                    </ItemTemplate>
                                </asp:TemplateField>
                            </Columns>
                            <AlternatingRowStyle BackColor="#FFFFdd" />
                        </asp:GridView>
                        <br />
                        <asp:Label ID="Label3" runat="server" Font-Bold="False" Text="Enter any additional information you want to tell us:"
                            Visible="False" />
                        <asp:TextBox ID="tbInformation" runat="server" Rows="4" TextMode="MultiLine" Width="95%"
                            MaxLength="500" Visible="False"></asp:TextBox>
                        <asp:Button ID="btnFinalCheckBottom" runat="server" Text="CHECK ORDER" Width="228px"
                            OnClick="btnFinalCheck_Click" />
                    </td>
                    <td>
                    </td>
                </tr>
            </table>
        </asp:Panel>
        <asp:Panel ID="pnlFinalCheck" runat="server" Width="100%" Visible="false">
            <table style="width: 100%">
                <tr>
                    <td style="width: 5%" />
                    <td style="width: 90%" colspan="2" />
                    <td style="width: 5%" />
                </tr>
                <tr>
                    <td />
                    <td>
                        <asp:Label ID="Label17" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="Medium"
                            Text="NHS Connecting for Health Summary Care Records Public Information Programme" />
                        <br />
                        <br />
                        <b><font face="Calibri" size="3">Check Order</font></b>
                    </td>
                    <td align="right">
                        <img alt="" src="images/NHS-RGB.jpg" />
                    </td>
                    <td />
                </tr>
                <tr>
                    <td />
                    <td align="right" colspan="2">
                        &nbsp;
                    </td>
                    <td />
                </tr>
                <tr>
                    <td />
                    <td colspan="2">
                        <asp:Label ID="lblOrderSummary" runat="server" Font-Bold="False" Font-Size="Small" />
                    </td>
                    <td />
                </tr>
                <tr>
                    <td />
                    <td colspan="2">
                    </td>
                    <td />
                </tr>
                <tr>
                    <td />
                    <td colspan="2">
                        <asp:Button ID="btnConfirm" runat="server" Text="SUBMIT ORDER" Width="405px" OnClick="btnConfirm_Click" />
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        <asp:Button ID="btnGoBack" runat="server" OnClick="btnGoBack_Click" Text="GO BACK TO ORDER PAGE"
                            Width="201px" />
                    </td>
                    <td />
                </tr>
                <tr>
                    <td />
                    <td colspan="2">
                    </td>
                    <td />
                </tr>
            </table>
        </asp:Panel>
        <asp:Panel ID="pnlComplete" runat="server" Width="100%" Visible="false">
            <table style="width: 100%">
                <tr>
                    <td style="width: 5%" />
                    <td style="width: 90%" />
                    <td style="width: 5%" />
                </tr>
                <tr>
                    <td />
                    <td>
                        <asp:Label ID="Label18" runat="server" Font-Bold="True" Text="NHS Connecting for Health Summary Care Records Public Information Programme"
                            Font-Names="Verdana" Font-Size="Medium" />
                        <br />
                        <br />
                        <b><font face="Calibri" size="3">Order Complete</font></b>
                    </td>
                    <td />
                </tr>
                <tr>
                    <td />
                    <td>
                        <br />
                        <br />
                        <br />
                        <br />
                        <br />
                        <br />
                        Your order is complete. A confirmation has been sent to your email address.<br />
                        <br />
                        <br />
                        <asp:Button ID="btnCloseWindow" runat="server" OnClientClick="javascript: self.close ()"
                            Text="CLOSE WINDOW" />
                        <br />
                        <br />
                    </td>
                    <td />
                </tr>
            </table>
        </asp:Panel>
    </div>
    </form>
</body>
</html>
