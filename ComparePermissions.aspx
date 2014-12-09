<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data.SqlTypes" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Collections.Generic" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

    Const LOG_FILENAME As String = "CompareUserPermissions.txt"
    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    
    Dim dictReferencePermissions As Dictionary(Of Integer, Boolean)
    Dim dictReferenceMaxGrabs As Dictionary(Of Integer, Integer)
    Dim swStreamWriter As StreamWriter

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        Server.ScriptTimeout = 3600
        'Call SetTitle()
        If Not IsPostBack Then
            Call InitUaerGroups()
        End If
    End Sub

    Protected Sub SetTitle()
        Dim sTitle As String = Session("SiteTitle")
        If sTitle <> String.Empty Then
            sTitle += " - "
        End If
        Page.Header.Title = sTitle & "Compare User Permissions"
    End Sub
    
    Protected Sub InitUaerGroups()
        Dim sSQL As String = "SELECT CustomerAccountCode + ' - ' + GroupName 'Entry', [id] FROM UP_UserPermissionGroups upg INNER JOIN Customer c ON upg.CustomerKey = c.CustomerKey ORDER BY Entry"
        lbUserGroups.Items.Clear()
        Dim oListItemCollection As ListItemCollection = ExecuteQueryToListItemCollection(sSQL, "Entry", "id")
        For Each li In oListItemCollection
            lbUserGroups.Items.Add(li)
        Next
    End Sub
    
    Protected Sub InitUaerInGroup()
        
    End Sub
    
    Protected Sub rbAllUsers_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        trCompareUsers.Visible = False
        tbCompareUsers.Text = String.Empty

    End Sub

    Protected Sub rbUsersInSameUserGroup_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        trCompareUsers.Visible = False
        tbCompareUsers.Text = String.Empty

    End Sub

    Protected Sub rbUsersListed_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        trCompareUsers.Visible = True
    End Sub

    Protected Sub btnCompare_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        tbReferenceUser.Text = tbReferenceUser.Text.Trim
        tbResult.Text = String.Empty
        lblUserInfo.Text = String.Empty
        If tbReferenceUser.Text <> String.Empty Then
            Dim oDataTable As DataTable = ExecuteQueryToDataTable("SELECT * FROM UserProfile WHERE UserId = '" & tbReferenceUser.Text.Trim.Replace("''", "'") & "'")
            If oDataTable.Rows.Count > 0 Then
                Dim dr As DataRow = oDataTable.Rows(0)
                lblUserInfo.Text = "USER: " & dr("FirstName") & " " & dr("LastName") & "      CUSTOMER: " & ExecuteQueryToDataTable("SELECT CustomerAccountCode FROM Customer WHERE CustomerKey = " & dr("CustomerKey")).Rows(0).Item(0)
                If dr("Type") = "SuperUser" Then
                    WebMsgBox.Show("User " & tbReferenceUser.Text & " is a Super User. Super Users don't have permissions. Select a non SuperUser to to compare instead.")
                    tbReferenceUser.Text = String.Empty
                    tbReferenceUser.Focus()
                    Exit Sub
                End If
                If rbUsersInSameUserGroup.Checked Then
                    'If IsDBNull(dr("UserGroup")) OrElse dr("UserGroup") >= 0 Then
                    If IsDBNull(dr("UserGroup")) OrElse dr("UserGroup") = 0 Then
                        WebMsgBox.Show("You have asked for comparison with users in the same user group, but this user is not in a user group.")
                        Exit Sub
                    End If
                End If
            Else
                WebMsgBox.Show("Could not locate a user with this User Id.")
            End If
            
            Call Compare()
        Else
            WebMsgBox.Show("Please enter reference user.")
        End If
    End Sub
    
    Protected Sub Compare()
        Dim oDataTable As DataTable = Nothing
        Dim bUserHasDifferences As Boolean = False
        Dim nUsersWithDifferencesCount As Integer = 0
        Dim bUserRecordIsDifferent As Boolean = False
        Dim nPerUserDifferencesCount As Integer = 0
        If cbWriteLogFile.Checked Then
            Call InitLogFile()
        End If
        
        oDataTable = ExecuteQueryToDataTable("SELECT [key], CustomerKey, UserGroup FROM UserProfile WHERE UserId LIKE '" & tbReferenceUser.Text.Replace("'", "''") & "'")     ' get details of reference user
        If oDataTable.Rows.Count > 0 Then
            pnReferenceUserKey = oDataTable.Rows(0).Item(0)
            pnCustomerKey = oDataTable.Rows(0).Item(1)
            If Not IsDBNull(oDataTable.Rows(0).Item(2)) Then
                pnReferenceUserGroup = oDataTable.Rows(0).Item(2)
            Else
                pnReferenceUserGroup = 0
            End If
        Else
            WebMsgBox.Show("Reference user not found.")
        End If
        Dim oListItemCollection As ListItemCollection = GetUsersToCompare()
        If oListItemCollection.Count > 0 Then
            Call InitPermissionsForReferenceUser()
            Call InitMaxGrabsForReferenceUser()
            DisplayReferenceUserSettings()
            For Each li As ListItem In oListItemCollection  ' oListItemCollection contains a li for each user to compare
                bUserHasDifferences = False
                nPerUserDifferencesCount = 0
                oDataTable = ExecuteQueryToDataTable("SELECT ProductKey, AbleToPick, MaxGrabQty FROM UserProductProfile upp INNER JOIN LogisticProduct lp ON upp.ProductKey = lp.LogisticProductKey WHERE lp.DeletedFlag = 'N' AND UserKey = " & li.Value)     ' get product profile of user being compared
                For Each dr As DataRow In oDataTable.Rows
                    Dim nProductKey As Integer = dr("ProductKey")
                    Dim nAbleToPick As Boolean = dr("AbleToPick")
                    Dim nMaxGrabQty As Integer = dr("MaxGrabQty")
                    bUserRecordIsDifferent = False
                    Try
                        If dictReferencePermissions.Item(nProductKey) <> nAbleToPick Then
                            bUserHasDifferences = True
                            bUserRecordIsDifferent = True
                            If cbDetailedListing.Checked Then
                                WriteLog("User: " & li.Text & " (" & li.Value & "); permissions for product " & sGetProductFromKey(nProductKey) & " differs (" & dictReferencePermissions.Item(nProductKey) & ", " & nAbleToPick & ")")
                            Else
                                WriteLog("User: " & li.Text & " (" & li.Value & "); permissions for product " & sGetProductFromKey(nProductKey) & " differs (" & dictReferencePermissions.Item(nProductKey) & ", " & nAbleToPick & ")")
                                Exit For
                            End If
                        End If
                    Catch ex As Exception
                        WriteLog("Product " & nProductKey & " not present at all in this profile !!!")
                        bUserRecordIsDifferent = True
                    End Try
                    Try
                        If dictReferenceMaxGrabs.Item(nProductKey) <> nMaxGrabQty Then
                            bUserHasDifferences = True
                            bUserRecordIsDifferent = True
                            If cbDetailedListing.Checked Then
                                WriteLog("User: " & li.Text & " (" & li.Value & "); max grab for product " & sGetProductFromKey(nProductKey) & " differs (" & dictReferenceMaxGrabs.Item(nProductKey) & ", " & nMaxGrabQty & ")")
                            Else
                                WriteLog("User: " & li.Text & " (" & li.Value & "); max grab for product " & sGetProductFromKey(nProductKey) & " differs (" & dictReferenceMaxGrabs.Item(nProductKey) & ", " & nMaxGrabQty & ")")
                                Exit For
                            End If
                        End If
                    Catch ex As Exception
                        WriteLog("Product " & nProductKey & " not present at all in this profile !!!")
                        bUserRecordIsDifferent = True
                    End Try
                    If bUserRecordIsDifferent = True Then
                        nPerUserDifferencesCount += 1
                    End If
                Next
                If nPerUserDifferencesCount > 0 Then
                    WriteLog("This user had " & nPerUserDifferencesCount & " different record(s)")
                    WriteLog("")
                End If
                If bUserHasDifferences Then
                    nUsersWithDifferencesCount += 1
                End If
            Next
            WriteLog("FINISHED !")
            WriteLog("Total users compared: " & oListItemCollection.Count)
            WriteLog("Users with one or more differences: " & nUsersWithDifferencesCount)
        Else
            WebMsgBox.Show("No user to compare!")
        End If
        If cbWriteLogFile.Checked Then
            Call CloseLogFile()
        End If

    End Sub
    
    Protected Sub DisplayReferenceUserSettings()
        'WriteLog("Product profile for reference account " & tbReferenceUser.Text & " (" & pnReferenceUserKey & ")")
        tbReferencePermissions.Text = String.Empty
        For Each kv As KeyValuePair(Of Integer, Boolean) In dictReferencePermissions
            tbReferencePermissions.Text += sGetProductFromKey(kv.Key) & ": CanPick is " & kv.Value & "; maxgrab is " & dictReferenceMaxGrabs.Item(kv.Key) & Environment.NewLine
        Next
        tbReferencePermissions.Text += dictReferencePermissions.Count & " products in set" & Environment.NewLine
    End Sub
    
    Protected Sub InitLogFile()
        Dim sMappedFilename As String = Server.MapPath(LOG_FILENAME)
        My.Computer.FileSystem.DeleteFile(sMappedFilename)
        swStreamWriter = New StreamWriter(Server.MapPath(LOG_FILENAME))
    End Sub
    
    Protected Sub WriteLog(ByVal sLine As String)
        tbResult.Text += sLine & Environment.NewLine
        If cbWriteLogFile.Checked Then
            swStreamWriter.WriteLine(sLine)
        End If
    End Sub
    
    Protected Sub CloseLogFile()
        swStreamWriter.Close()
    End Sub
    
    Protected Function sGetProductFromKey(ByVal nProductKey As Integer) As String
        Dim oDataTable As DataTable = ExecuteQueryToDataTable("SELECT ISNULL(ProductCode,'') 'ProductCode', ISNULL(ProductDate,'') 'ProductDate', ISNULL(ProductDescription,'') 'ProductDescription' FROM LogisticProduct WHERE LogisticProductKey = " & nProductKey)
        If oDataTable.Rows.Count > 0 Then
            sGetProductFromKey = oDataTable.Rows(0)("ProductCode") & " " & oDataTable.Rows(0)("ProductDate") & " " & oDataTable.Rows(0)("ProductDescription") & " (" & nProductKey & ")"
        Else
            sGetProductFromKey = "Could not retrieve product details!!!!!"
        End If
    End Function
        
    Protected Sub InitPermissionsForReferenceUser()
        dictReferencePermissions = New Dictionary(Of Integer, Boolean)
        Dim oDataTable As DataTable = Nothing
        oDataTable = ExecuteQueryToDataTable("SELECT ProductKey, AbleToPick FROM UserProductProfile WHERE UserKey = " & pnReferenceUserKey)
        For Each dr As DataRow In oDataTable.Rows
            dictReferencePermissions.Add(dr("ProductKey"), dr("AbleToPick"))
        Next
    End Sub
    
    Protected Sub InitMaxGrabsForReferenceUser()
        dictReferenceMaxGrabs = New Dictionary(Of Integer, Integer)
        Dim oDataTable As DataTable = Nothing
        oDataTable = ExecuteQueryToDataTable("SELECT ProductKey, MaxGrabQty FROM UserProductProfile WHERE UserKey = " & pnReferenceUserKey)
        For Each dr As DataRow In oDataTable.Rows
            dictReferenceMaxGrabs.Add(dr("ProductKey"), dr("MaxGrabQty"))
        Next
    End Sub
    
    Protected Function GetUsersToCompare() As ListItemCollection
        Dim oListItemCollection As ListItemCollection = Nothing
        Dim sCompareUsersType As String = String.Empty
        If rbAllUsers.Checked Then
            sCompareUsersType = "A"
        ElseIf rbUsersInSameUserGroup.Checked Then
            sCompareUsersType = "G"
        ElseIf rbUsersListed.Checked Then
            sCompareUsersType = "L"
        End If
        
        Select Case sCompareUsersType
            Case "A"
                oListItemCollection = ExecuteQueryToListItemCollection("SELECT UserId, [key] FROM UserProfile WHERE CustomerKey = " & pnCustomerKey & " AND Type = 'User' AND [key] <> " & pnReferenceUserKey, "UserId", "key")
            Case "G"
                If pnReferenceUserGroup > 0 Then
                    oListItemCollection = ExecuteQueryToListItemCollection("SELECT UserId, [key] FROM UserProfile WHERE UserGroup = " & pnReferenceUserGroup & " AND Type = 'User' AND [key] <> " & pnReferenceUserKey, "UserId", "key")
                Else
                    WebMsgBox.Show("Cannot compare with other users in group since reference user does not have a user group defined.")
                End If
            Case "L"
                oListItemCollection = New ListItemCollection
                Dim sTemp As String = tbCompareUsers.Text.Replace(Environment.NewLine, " ").Trim
                Dim sListedUsers() As String = sTemp.Split(" ")
                For Each s As String In sListedUsers
                    If s.Trim <> String.Empty Then
                        Dim li As New ListItem
                        li.Text = s.Trim
                        li.Value = nGetUserKeyFromUserId(s.Trim)
                        oListItemCollection.Add(li)
                    End If
                Next
        End Select
        GetUsersToCompare = oListItemCollection
    End Function
    
    Protected Function nGetUserKeyFromUserId(ByVal sUserId As String) As Integer
        nGetUserKeyFromUserId = 0
        Dim oDataTable As DataTable = ExecuteQueryToDataTable("SELECT [key], CustomerKey FROM UserProfile WHERE UserId LIKE '" & sUserId.Replace("'", "''") & "'")
        If oDataTable.Rows.Count > 0 Then
            nGetUserKeyFromUserId = oDataTable.Rows(0).Item(0)
            If oDataTable.Rows(0).Item(1) <> pnCustomerKey Then
                nGetUserKeyFromUserId = 0
                WebMsgBox.Show("UserId " & sUserId & " is not the same customer as the reference user!")
            End If
        End If
    End Function

    Protected Function nGetUserInfoFromUserId(ByVal sUserId As String) As Integer
        Dim oDataTable As DataTable = ExecuteQueryToDataTable("SELECT [key], CustomerKey FROM UserProfile WHERE UserId LIKE '" & tbReferenceUser.Text.Replace("'", "''") & "'")
        If oDataTable.Rows.Count > 0 Then
            nGetUserInfoFromUserId = oDataTable.Rows(0).Item(0)
        Else
            nGetUserInfoFromUserId = 0
        End If
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
            WebMsgBox.Show("Error in ExecuteQueryToListItemCollection executing: " & sQuery & " : " & ex.Message)
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

    Property pnReferenceUserKey() As Integer
        Get
            Dim o As Object = ViewState("CUP_ReferenceUserKey")
            If o Is Nothing Then
                Return 0
            End If
            Return CLng(o)
        End Get
        Set(ByVal Value As Integer)
            ViewState("CUP_ReferenceUserKey") = Value
        End Set
    End Property

    Property pnCustomerKey() As Integer
        Get
            Dim o As Object = ViewState("CUP_CustomerKey")
            If o Is Nothing Then
                Return 0
            End If
            Return CLng(o)
        End Get
        Set(ByVal Value As Integer)
            ViewState("CUP_CustomerKey") = Value
        End Set
    End Property
    
    Property pnReferenceUserGroup() As Integer
        Get
            Dim o As Object = ViewState("CUP_ReferenceUserGroup")
            If o Is Nothing Then
                Return 0
            End If
            Return CLng(o)
        End Get
        Set(ByVal Value As Integer)
            ViewState("CUP_ReferenceUserGroup") = Value
        End Set
    End Property

    ' was below first div
    ' <%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
    '            <main:Header ID="ctlHeader" runat="server"></main:Header>

    Protected Sub lbUserGroups_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim lb As ListBox = sender
        Dim sSQL As String = "SELECT UserId + ' (' + FirstName + ' ' + LastName + ')' 'UserEntry', [key] FROM UserProfile WHERE UserGroup = " & lbUserGroups.SelectedItem.Value & " ORDER BY UserId"
        Dim oListItemCollection As ListItemCollection = ExecuteQueryToListItemCollection(sSQL, "UserEntry", "key")
        lbUsersInGroup.Items.Clear()
        For Each li As ListItem In oListItemCollection
            lbUsersInGroup.Items.Add(li)
        Next
        tbPermissionsForSelectedUser.Text = String.Empty
    End Sub
    
    Protected Sub lbUsersInGroup_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim lb As ListBox = sender
        Dim oDataTable As DataTable = ExecuteQueryToDataTable("SELECT ProductCode + ' - ' + ISNULL(ProductDate,''), CAST(ApplyMaxGrab AS varchar(2)) + ' (' + CAST(MaxGrabQty AS varchar(6)) + ')' FROM UserProductProfile upp INNER JOIN LogisticProduct lp on upp.ProductKey = lp.LogisticProductKey WHERE UserKey = " & lb.SelectedItem.Value & " AND lp.DeletedFlag = 'N' ORDER BY ProductCode")
        For Each dr As DataRow In oDataTable.Rows
            tbPermissionsForSelectedUser.Text += dr(0) & ", " & dr(1) & Environment.NewLine
        Next
    End Sub
    
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="Form1" runat="Server">
        <div style="font-size: xx-small; font-family: Verdana">
          <main:Header ID="ctlHeader" runat="server"></main:Header>
            <table width="95%">
                <tr>
                    <td colspan="2" style="height: 14px">
                        <strong>Compare Permissions</strong></td>
                </tr>
                <tr>
                    <td align="right" style="width: 10%" valign="top">
                        Reference user:
                    </td>
                    <td style="width: 90%">
                        <asp:TextBox ID="tbReferenceUser" runat="server" Font-Names="Verdana" 
                            Font-Size="XX-Small" Width="50%" />
                    </td>
                </tr>
                <tr>
                    <td align="right" style="width: 10%" valign="top">
                        Compare with:
                    </td>
                    <td style="width: 90%">
                        <asp:RadioButton ID="rbAllUsers" runat="server" GroupName="CompareUsers" Text="all other users for this customer" OnCheckedChanged="rbAllUsers_CheckedChanged" AutoPostBack="True" />
                        <asp:RadioButton ID="rbUsersInSameUserGroup" runat="server" Checked="True" GroupName="CompareUsers" Text="users in same user group" OnCheckedChanged="rbUsersInSameUserGroup_CheckedChanged" AutoPostBack="True" />
                        <asp:RadioButton ID="rbUsersListed" runat="server" GroupName="CompareUsers" Text="users listed" OnCheckedChanged="rbUsersListed_CheckedChanged" AutoPostBack="True" /></td>
                </tr>
                <tr runat="server" id="trCompareUsers" visible="false">
                    <td align="right" style="width: 10%" valign="top">
                        Users:<br />
                        <br />
                        (enter a space-<br />
                        separated &nbsp;list<br />
                        of user Ids)
                    </td>
                    <td style="width: 90%">
                        <asp:TextBox ID="tbCompareUsers" runat="server" Font-Names="Verdana" 
                            Font-Size="XX-Small" Rows="6" TextMode="MultiLine" Width="95%"/>
                    </td>
                </tr>
                <tr>
                    <td align="right" style="width: 10%" valign="top">
                    </td>
                    <td style="width: 90%">
                        <asp:Button ID="btnCompare" runat="server" Text="compare" OnClick="btnCompare_Click" Width="200px" />&nbsp;
                        <asp:Label ID="lblUserInfo" runat="server" Font-Bold="True" Font-Names="Verdana"
                            Font-Size="XX-Small" ForeColor="Maroon"></asp:Label></td>
                </tr>
                <tr>
                    <td align="right" style="width: 10%" valign="top">
                        Reference user's permissions:
                    </td>
                    <td>
                        <asp:TextBox ID="tbReferencePermissions" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Rows="6" TextMode="MultiLine" Width="95%"/>
                    </td>
                </tr>
                <tr>
                    <td align="right" valign="top">
                        Non matches with other users:
                    </td>
                    <td>
                        <asp:TextBox ID="tbResult" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Rows="20" TextMode="MultiLine" Width="95%"/>
                    </td>
                </tr>
                <tr>
                    <td align="right" valign="top">
                    </td>
                    <td>
                        &nbsp;<asp:CheckBox ID="cbDetailedListing" runat="server" Text="detailed log" />
                        &nbsp; &nbsp; &nbsp;&nbsp;
                        <asp:CheckBox ID="cbWriteLogFile" runat="server" Text="write log file" Visible="False" /><br />
                        <br />
                        <strong>detailed log</strong> check box <strong><span style="color: red">un</span></strong>checked:
                        just report first difference<br />
                        <strong>detailed log</strong> check box checked: report <strong><span style="color: red">
                            all</span></strong> differences.<br />
                        <br />
                    </td>
                </tr>
                <tr>
                    <td align="right" valign="top">
                        User groups:
                    </td>
                    <td>
                        <asp:ListBox ID="lbUserGroups" runat="server" Width="95%" Rows="6" 
                            AutoPostBack="True" onselectedindexchanged="lbUserGroups_SelectedIndexChanged" 
                            Font-Names="Verdana" Font-Size="XX-Small"/>
                    </td>
                </tr>
                <tr>
                    <td align="right" valign="top">
                        Users in group:
                    </td>
                    <td>
                        <asp:ListBox ID="lbUsersInGroup" runat="server" Width="95%" Rows="10" 
                            AutoPostBack="True" 
                            onselectedindexchanged="lbUsersInGroup_SelectedIndexChanged" 
                            Font-Names="Verdana" Font-Size="XX-Small"/>
                    </td>
                </tr>
                <tr>
                    <td align="right" valign="top">
                        Permissions for selected user:
                    </td>
                    <td>
                        <asp:TextBox ID="tbPermissionsForSelectedUser" runat="server" Width="95%" 
                            Font-Names="Verdana" Font-Size="XX-Small" Rows="8" TextMode="MultiLine"/>
                    </td>
                </tr>
            </table>
        </div>
    </form>
</body>
</html>
