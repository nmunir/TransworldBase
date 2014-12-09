<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ import Namespace="System.IO" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient " %>
<%@ import Namespace="System.Collections.Generic" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" " http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

    Const BASE_STORAGE_DIRECTORY As String = "uploads\"
    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    Dim gsUploadTimestamp As String
    Dim gsOrigFileName As String
    Dim gsStoredFilename As String
    Dim gsStorageDirectory As String
    
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsNumeric(Session("CustomerKey")) Then
            Server.Transfer("session_expired.aspx")
            Exit Sub
        End If
        If IsPostBack Then
            Call HideInstructionsPanels()
        End If
        Call SetTitle()
        gsStorageDirectory = BASE_STORAGE_DIRECTORY & Session("CustomerName").ToString.Replace(" ", "_") & "\"
    End Sub

    Protected Sub SetTitle()
        Dim sTitle As String = Session("SiteTitle")
        If sTitle <> String.Empty Then
            sTitle += " - "
        End If
        Page.Header.Title = sTitle & "File Upload"
    End Sub
    
    Protected Sub HideInstructionsPanels()
        pnlInstructions.Visible = False
    End Sub
    
    Protected Sub btnUpload_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call DoUpload()
    End Sub

    Protected Sub DoUpload()
        Dim sFileSuffix As String = ".upl"
        If FileUpload1.HasFile Then
            Call CheckStorageDirectoryExists()
            gsOrigFileName = FileUpload1.FileName
            If Path.HasExtension(gsOrigFileName) Then
                sFileSuffix = Path.GetExtension(gsOrigFileName)
            End If
            gsStoredFilename = Format(Now(), "yyyymmddhhmmssff")
            gsStoredFilename += "(" & gsOrigFileName & ")"
            gsStoredFilename = Server.MapPath("") & "\" & gsStorageDirectory & gsStoredFilename & sFileSuffix
            FileUpload1.SaveAs(gsStoredFilename)

            Dim oConn As New SqlConnection(gsConn)
            Dim oCmd As SqlCommand
            gsUploadTimestamp = Now.ToLongDateString() & " " & Format(Now, "hh:mm:ss")
            tbUploadMessage.Text = tbUploadMessage.Text.Trim
            Dim sSQL As String = "INSERT INTO FileUploadLog (UploadDateTime, OriginalFileName, StoredFileName, CustomerKey, UserId, AppId, UploadMessage) VALUES ('"
            sSQL += gsUploadTimestamp & "', '" & gsOrigFileName.Replace("'", "''") & "', '" & gsStoredFilename.Replace("'", "''") & "', " & Session("CustomerKey") & ", " & Session("UserKey") & ", '', '" & tbUploadMessage.Text.Replace("'", "''") & "')"
            Try
                oConn.Open()
                oCmd = New SqlCommand(sSQL, oConn)
                oCmd.ExecuteNonQuery()
            Catch ex As Exception
                WebMsgBox.Show("Error in DoUpload: " & ex.Message)
            Finally
                oConn.Close()
            End Try
            If pnlUploadHistory.Visible = True Then
                Call DisplayUploadHistory()
            End If
            Dim oDataTable As DataTable = GetNotificationRecipients()
            Dim nNotificationRecipientCount As Integer = oDataTable.Rows.Count
            Dim sConfirmationMessage As String = "File uploaded successfully. "
            If nNotificationRecipientCount > 0 Then
                If nNotificationRecipientCount = 1 Then
                    sConfirmationMessage += "Notification sent to the following recipient: "
                Else
                    sConfirmationMessage += "Notification sent to the following " & nNotificationRecipientCount.ToString & " recipients: "
                End If
                Dim bAppendComma As Boolean = False
                For Each dr As DataRow In oDataTable.Rows
                    SendHTMLEmail(dr("EmailAddr"))
                    If Not bAppendComma Then
                        bAppendComma = True
                    Else
                        sConfirmationMessage += ", "
                    End If
                    sConfirmationMessage += dr("EmailAddr")
                Next
            Else
                sConfirmationMessage += "No upload notifications have been sent as no recipients are currently defined (contact your Account Handler to add or change the list of upload notification recipients)."
            End If
            WebMsgBox.Show(sConfirmationMessage)
            tbUploadMessage.Text = String.Empty
        Else
            Dim sFilename As String = FileUpload1.FileName.Trim
            If sFilename = String.Empty Then
                WebMsgBox.Show("No filename specified")
            Else
                WebMsgBox.Show("Could not upload file " & sFilename & " - please check filename and file availability")
            End If
        End If
    End Sub
    
    Protected Function GetNotificationRecipients() As DataTable
        GetNotificationRecipients = Nothing
        Dim sSQL As String = "SELECT EmailAddr FROM FileUploadNotification WHERE CustomerKey = " & Session("CustomerKey")
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter(sSQL, oConn)
        Dim oDataTable As New DataTable
        Try
            oAdapter.Fill(oDataTable)
            GetNotificationRecipients = oDataTable
        Catch ex As Exception
            WebMsgBox.Show("Error in GetNotificationRecipients: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Function
    
    Protected Sub SendHTMLEmail(ByVal sRecipient As String)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Email_AddToQueue", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
    
        Try
            oCmd.Parameters.Add(New SqlParameter("@MessageId", SqlDbType.NVarChar, 20))
            oCmd.Parameters("@MessageId").Value = "FILE UPLOAD ALERT"
    
            oCmd.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int, 4))
            oCmd.Parameters("@CustomerKey").Value = Session("CustomerKey")
    
            oCmd.Parameters.Add(New SqlParameter("@StockBookingKey", SqlDbType.Int, 4))
            oCmd.Parameters("@StockBookingKey").Value = 0
    
            oCmd.Parameters.Add(New SqlParameter("@ConsignmentKey", SqlDbType.Int, 4))
            oCmd.Parameters("@ConsignmentKey").Value = 0
    
            oCmd.Parameters.Add(New SqlParameter("@ProductKey", SqlDbType.Int, 4))
            oCmd.Parameters("@ProductKey").Value = 0
    
            oCmd.Parameters.Add(New SqlParameter("@To", SqlDbType.NVarChar, 100))
            oCmd.Parameters("@To").Value = sRecipient
    
            oCmd.Parameters.Add(New SqlParameter("@Subject", SqlDbType.NVarChar, 60))
            oCmd.Parameters("@Subject").Value = "File Upload Alert"
    
            Dim sbPlainText As New StringBuilder
            sbplainText.Append("New file uploaded")
            sbplainText.Append(Environment.NewLine)
            sbplainText.Append("Upload received: " & gsUploadTimestamp)
            sbplainText.Append(Environment.NewLine)
            sbplainText.Append("Original filename: " & gsOrigFileName)
            sbplainText.Append(Environment.NewLine)
            sbplainText.Append("Local (server) filename: " & gsStoredFilename)
            sbplainText.Append(Environment.NewLine)
            If tbUploadMessage.Text <> String.Empty Then
                sbPlainText.Append(Environment.NewLine)
                sbplainText.Append("Message:")
                sbPlainText.Append(Environment.NewLine)
                sbplainText.Append(tbUploadMessage.Text)
                sbplainText.Append(Environment.NewLine)
                sbplainText.Append("[end]")
            End If
            oCmd.Parameters.Add(New SqlParameter("@BodyText", SqlDbType.NText))
            oCmd.Parameters("@BodyText").Value = sbPlainText.ToString
    
            Dim sbHTMLText As New StringBuilder
            sbHTMLText.Append("New file uploaded")
            sbHTMLText.Append("<br />" & Environment.NewLine)
            sbHTMLText.Append("Upload received: " & gsUploadTimestamp)
            sbHTMLText.Append("<br />" & Environment.NewLine)
            sbHTMLText.Append("Original filename: " & gsOrigFileName)
            sbHTMLText.Append("<br />" & Environment.NewLine)
            sbHTMLText.Append("Local (server) filename: " & gsStoredFilename)
            sbHTMLText.Append("<br />" & Environment.NewLine)
            If tbUploadMessage.Text <> String.Empty Then
                sbHTMLText.Append("<br />" & Environment.NewLine)
                sbHTMLText.Append("Message:")
                sbHTMLText.Append("<br />" & Environment.NewLine)
                sbHTMLText.Append(tbUploadMessage.Text)
                sbHTMLText.Append("<br />" & Environment.NewLine)
                sbHTMLText.Append("[end]")
            End If
            oCmd.Parameters.Add(New SqlParameter("@BodyHTML", SqlDbType.NText))
            oCmd.Parameters("@BodyHTML").Value = sbHTMLText.ToString
    
            oCmd.Parameters.Add(New SqlParameter("@QueuedBy", SqlDbType.Int, 4))
            oCmd.Parameters("@QueuedBy").Value = Session("UserKey")
    
            oConn.Open()
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            WebMsgBox.Show("Error in SendHTMLEmail: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub CheckStorageDirectoryExists()
        Dim sStoragePath As String = Server.MapPath("") & "\" & gsStorageDirectory
        Dim diDir As New DirectoryInfo(sStoragePath)
        If Not diDir.Exists Then
            Directory.CreateDirectory(sStoragePath)
        End If
    End Sub
    
    Protected Sub lnkbtnUploadHistory_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If lnkbtnUploadHistory.Text.ToLower.Contains("show") Then
            lnkbtnUploadHistory.Text = "hide upload history"
            pnlUploadHistory.Visible = True
            Call DisplayUploadHistory()
        Else
            lnkbtnUploadHistory.Text = "show upload history"
            pnlUploadHistory.Visible = False
        End If
        
    End Sub
    
    Protected Sub DisplayUploadHistory()
        Dim sSQL As String = "SELECT UploadDateTime 'Uploaded At', OriginalFilename 'File Name', FirstName + ' ' + LastName + ' (' + up.UserId + ')' 'Account' FROM FileUploadLog ful INNER JOIN UserProfile up ON ful.UserId = up.[key] WHERE ful.CustomerKey = " & Session("CustomerKey") & " ORDER BY [id] DESC"
            
        Dim sConn As String = ConfigurationManager.AppSettings("AIMSRootConnectionString")
        Dim oConn As New SqlConnection(sConn)
        Dim oAdapter As New SqlDataAdapter(sSQL, oConn)
        Dim oDataTable As New DataTable
        Try
            oAdapter.Fill(oDataTable)
            gvUploadHistory.DataSource = oDataTable
            gvUploadHistory.DataBind()
        Catch ex As Exception
            WebMsgBox.Show("Error in DisplayUploadHistory: " & ex.Message)
        Finally
            oConn.Close()
        End Try

    End Sub
    
</script>

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>File Upload</title>
    </head>
<body>
    <form id="form1" runat="server">
      <main:Header id="ctlHeader" runat="server"></main:Header>
        <table width="100%" cellpadding="0" cellspacing="0">
            <tr class="bar_accounthandler">
                <td style="width:50%; white-space:nowrap">
                </td>
                <td style="width:50%; white-space:nowrap" align="right">
                </td>
            </tr>
        </table>
    <div>
        <table style="width: 100%; font-size: xx-small; font-family: Verdana;" border="0">
            <tr>
                <td align="left" style="width: 20%">
        &nbsp;
        <asp:Label ID="lblTitle" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="Small"
            Text="File Upload"></asp:Label></td>
                <td style="width: 60%">
                    <asp:LinkButton ID="lnkbtnUploadHistory" runat="server" OnClick="lnkbtnUploadHistory_Click">show upload history</asp:LinkButton></td>
                <td style="width: 20%">
                </td>
            </tr>
        </table>
        <br />
        <asp:Panel ID="pnlUpload" runat="server" Width="100%" Font-Names="Verdana">
        <table style="width: 100%; font-size: xx-small; font-family: Verdana;" border="0">
            <tr>
                <td style="width: 20%">
                </td>
                <td style="width: 60%">
                </td>
                <td style="width: 20%">
                </td>
            </tr>
            <tr>
                <td style="width: 20%" align="right">
                    Filename:</td>
                <td style="width: 60%">
                    <asp:FileUpload ID="FileUpload1" runat="server" Font-Names="Verdana" 
                        Font-Size="XX-Small" Width="400px" />
                </td>
                <td style="width: 20%">
                </td>
            </tr>
            <tr>
                <td align="right">&nbsp;
                </td>
                <td>
                <td>
                </td>
            </tr>
            <tr>
                <td align="right">
                    Message to recipient:<br />(optional)</td>
                <td>
        <asp:TextBox ID="tbUploadMessage" runat="server" Rows="6" TextMode="MultiLine" Width="100%" 
                        Font-Names="Verdana" Font-Size="XX-Small" MaxLength="1900"></asp:TextBox></td>
                <td>
                </td>
            </tr>
            <tr>
                <td align="right">
                </td>
                <td>
        <asp:Button ID="btnUpload" runat="server" OnClick="btnUpload_Click" Text="upload" /></td>
                <td>
                </td>
            </tr>
        </table>
        </asp:Panel>

        <asp:Panel ID="pnlUploadHistory" runat="server" Visible="false" Width="100%">
            <strong>
                &nbsp;<asp:Label ID="lblUploadHistory" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="Small" Text="Upload History"></asp:Label>
            <br />
            </strong>
            <table style="width: 100%">
                <tr>
                    <td align="center" style="width: 100%">
                        <asp:GridView ID="gvUploadHistory" runat="server" CellPadding="2" Font-Names="Verdana" Font-Size="XX-Small" Width="95%" EmptyDataText="no entries found">
                            <EmptyDataTemplate>
                                <asp:Label ID="Label1" runat="server" Font-Names="Verdana" Font-Size="XX-Small" 
                                    Text="no upload history available"></asp:Label>
                            </EmptyDataTemplate>    
                        </asp:GridView>
                    </td>
                </tr>
            </table>
        </asp:Panel>
        <asp:Panel runat="server" ID="pnlInstructions" Visible="true" Width="100%">
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
                        <asp:Label ID="Label3" runat="server" Font-Names="Verdana" Font-Size="XX-Small" 
                            Text="2. If required, enter a message describing the file contents into the message box. This message will be included in the upload alert alert emailed to the file recipient. The message is optional."></asp:Label>
                        <br />
                        <br />
                        <asp:Label ID="Label4" runat="server" Font-Names="Verdana" Font-Size="XX-Small" 
                            
                            Text="3. Click the &lt;b&gt;upload&lt;/b&gt; button. On completion of the upload you will see a message 'File uploaded successfully' and a list of the email recipients to whom an upload alert has been sent, if any are defined."></asp:Label>
                        <br />
                        <br />
                        <asp:Label ID="Label6" runat="server" Font-Names="Verdana" Font-Size="XX-Small" 
                            Text="4. You can view the upload history by clicking the &lt;b&gt;show upload history&lt;/b&gt; link at the top of this window."></asp:Label>
                        <br />
                    </td>
                    <td>
                        &nbsp;</td>
                </tr>
            </table>
        </asp:Panel>
        </div>
    </form>
</body>
</html>
