<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ import Namespace="System.IO" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient " %>
<%@ import Namespace="System.Collections.Generic" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" " http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    Dim STORAGE_DIRECTORY As String = "uploads\"

    Protected Sub btnUpload_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim sOrigFileName As String
        Dim sFileSuffix As String = ".upl"
        Session("CustomerKey") = 0
        Session("UserKey") = 0
        If FileUpload1.HasFile Then
            Call CheckStorageDirectoryExists()
            sOrigFileName = FileUpload1.FileName
            If Path.HasExtension(sOrigFileName) Then
                sFileSuffix = Path.GetExtension(sOrigFileName)
            End If
            Dim sStoredFilename = Format(Now(), "yyyymmddhhmmssff")
            sStoredFilename += "(" & sOrigFileName & ")"
            sStoredFilename = Server.MapPath("") & "\" & STORAGE_DIRECTORY & sStoredFilename & sFileSuffix
            FileUpload1.SaveAs(sStoredFilename)

            Dim oConn As New SqlConnection(gsConn)
            Dim oCmd As SqlCommand
            Dim sSQL As String = "INSERT INTO FileUploadLog (UploadDateTime, OriginalFileName, StoredFileName, CustomerKey, UserId, AppId, UploadMessage) VALUES ('"
            sSQL += Now.ToLongDateString() & " " & Format(Now, "hh:mm:ss") & "', '" & sOrigFileName.Replace("'", "''") & "', '" & sStoredFilename.Replace("'", "''") & "', " & Session("CustomerKey") & ", " & Session("UserKey") & ", 'BASKIND', '" & tbUploadMessage.Text.Replace("'", "''") & "')"
            oConn.Open()
            oCmd = New SqlCommand(sSQL, oConn)
            oCmd.ExecuteNonQuery()
            oConn.Close()
            tbUploadMessage.Text = String.Empty
            WebMsgBox.Show("File uploaded successfully")
        Else
            Dim sFilename As String = FileUpload1.FileName.Trim
            If sFilename = String.Empty Then
                WebMsgBox.Show("No filename specified")
            Else
                WebMsgBox.Show("Could not upload file " & sFilename & " - please check filename and file availability")
            End If
        End If

    End Sub
    
    Protected Function NotifyRecipients() As Integer
        Session("CustomerKey") = 0
        Dim sSQL As String = "SELECT EmailAddr FROM FileUploadNotification WHERE CustomerKey = " & Session("CustomerKey")
        Dim sConn As String = ConfigurationManager.AppSettings("AIMSRootConnectionString")
        Dim oConn As New SqlConnection(sConn)
        Dim oAdapter As New SqlDataAdapter(sSQL, oConn)
        Dim oDataTable As New DataTable
        Try
            oAdapter.Fill(oDataTable)
        Catch
        Finally
            oConn.Close()
        End Try
    End Function
    
    Protected Sub CheckStorageDirectoryExists()
        Dim sStoragePath As String = Server.MapPath("") & "\" & STORAGE_DIRECTORY
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
        Session("CustomerKey") = 0
        Dim sSQL As String = "SELECT UploadDateTime 'Uploaded At', OriginalFilename 'File Name' FROM FileUploadLog WHERE CustomerKey = " & Session("CustomerKey") & " ORDER BY [id] DESC"
            
        Dim sConn As String = ConfigurationManager.AppSettings("AIMSRootConnectionString")
        Dim oConn As New SqlConnection(sConn)
        Dim oAdapter As New SqlDataAdapter(sSQL, oConn)
        Dim oDataTable As New DataTable
        Try
            oAdapter.Fill(oDataTable)
            gvUploadHistory.DataSource = oDataTable
            gvUploadHistory.DataBind()
        Catch ex As Exception
        Finally
            oConn.Close()
        End Try

    End Sub
    
</script>

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>Secure Upload</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <table style="width: 100%; font-size: xx-small; font-family: Verdana;" border="0">
            <tr>
                <td align="left" style="width: 20%">
        <asp:Label ID="lblTitle" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="Small"
            Text="Transworld Secure File Upload"></asp:Label></td>
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
        <asp:FileUpload ID="FileUpload1" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
            Width="100%" /></td>
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
                    Message to Recipient:<br />(optional)</td>
                <td>
        <asp:TextBox ID="tbUploadMessage" runat="server" Rows="6" TextMode="MultiLine" Width="100%" Font-Names="Verdana" Font-Size="XX-Small"></asp:TextBox></td>
                <td>
                </td>
            </tr>
            <tr>
                <td align="right">
                </td>
                <td>
        <asp:Button ID="btnUpload" runat="server" OnClick="btnUpload_Click" Text="Upload" /></td>
                <td>
                </td>
            </tr>
        </table>
        </asp:Panel>

        <asp:Panel ID="pnlUploadHistory" runat="server" Visible="false" Width="100%">
            <strong>
                <asp:Label ID="lblUploadHistory" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="Small" Text="Upload History"></asp:Label><br />
                <br />
            </strong>
            <table style="width: 100%">
                <tr>
                    <td align="center" style="width: 100%">
                        <asp:GridView ID="gvUploadHistory" runat="server" CellPadding="2" Font-Names="Verdana" Font-Size="XX-Small" Width="95%" EmptyDataText="no entries found">
                        </asp:GridView>
                    </td>
                </tr>
            </table>
        </asp:Panel>
        </div>
    </form>
</body>
</html>
