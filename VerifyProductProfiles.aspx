<%@ Page Language="VB" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<script runat="server">
    
    ' TO DO ON THIS MODULE
    ' show counts
    ' add checkbox to ignore clone source when cloning
    ' warn when ignore clone on
    

    ' TO DO ON WU
    ' check all in supplied list are marked Active
    ' check all users in supplied WU list are on database
    ' list any users on database that are not in supplied WU list
    
    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    Dim goConn As New SqlConnection(gsConn)
    Dim oCmd As SqlCommand
    Dim sbDetail As New StringBuilder, sbSummary As New StringBuilder, sbSQL As New StringBuilder
    Dim nTotalUsers As Integer = 0, nUsersWithMissingProducts As Integer = 0

    Dim gnTimeout As Int32
    
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        lblError.Text = ""
        If Not IsPostBack Then
            'Call BindCustomerList()
            Call BindCustomerList2()
            Call PopulateModelUserDropdown()
        End If
    End Sub
    
    Protected Sub BindCustomerList()
        Dim sSQL As String = "SELECT CustomerAccountCode, CustomerKey FROM Customer WHERE CustomerStatusId = 'ACTIVE' ORDER BY CustomerAccountCode"
        Dim oAdapter As New SqlDataAdapter(sSQL, goConn)
        Dim oDataTable As New DataTable
        oAdapter.Fill(oDataTable)
        ddlCustomer.DataTextField = "CustomerAccountCode"
        ddlCustomer.DataValueField = "CustomerKey"
        ddlCustomer.DataSource = oDataTable
        ddlCustomer.DataBind()
    End Sub

    Protected Sub BindCustomerList2()
        Dim sSQL As String = "SELECT CustomerAccountCode, CustomerKey FROM Customer WHERE CustomerStatusId = 'ACTIVE' ORDER BY CustomerAccountCode"
        Dim oListItemCollection As ListItemCollection = ExecuteQueryToListItemCollection(sSQL, "CustomerAccountCode", "CustomerKey")
        ddlCustomer.Items.Clear()
        ddlCustomer.Items.Add(New ListItem("- all customers -", 0))
        For Each li As ListItem In oListItemCollection
            ddlCustomer.Items.Add(li)
        Next
    End Sub

    Protected Sub PopulateModelUserDropdown()
        ddlModelUser.Items.Clear()
        ddlModelUser.Items.Add(New ListItem("- please select -", 0))
        ddlModelUser.Items.Add(New ListItem("ALBEMARLE AND BOND - 5KEN - 13508", 13508))
        ddlModelUser.Items.Add(New ListItem("CASH CONVERTERS - A06V - 6313", 6313))
        ddlModelUser.Items.Add(New ListItem("CASH EXPRESS - A1KZ - 6324", 6324))
        ddlModelUser.Items.Add(New ListItem("CASH GENERATOR - 2PW4 - 15791", 15791))
        ddlModelUser.Items.Add(New ListItem("THE CO-OPERATIVE TRAVEL - 9DB7 - 14710", 14710))
        ddlModelUser.Items.Add(New ListItem("EUROCHANGE - JH2G - 9894", 9894))
        ddlModelUser.Items.Add(New ListItem("EXCHANGE INTERNATIONAL - C371 - 6813", 6813))
        ddlModelUser.Items.Add(New ListItem("HARVEY & THOMPSON - A12G - 17063", 17063))
        ddlModelUser.Items.Add(New ListItem("HERBERT BROWN - 5NTJ - 13529", 13529))
        ddlModelUser.Items.Add(New ListItem("ICE PLC (HUMBERSIDE AIRPORT) - M0FL - 7431", 7431))
        ddlModelUser.Items.Add(New ListItem("KANOO - OR9L - 7536", 7536))
        ddlModelUser.Items.Add(New ListItem("NCA - U6GI - 7927", 7927))
        ddlModelUser.Items.Add(New ListItem("RAMSDENS - STWN - 8144", 8144))
        ddlModelUser.Items.Add(New ListItem("THE MONEY SHOP - 1EC3 - 15863", 15863))
        ddlModelUser.Items.Add(New ListItem("THE TRADE EXCHANGE - DGTW - 8747", 8747))
        ddlModelUser.Items.Add(New ListItem("CHEQUE CASHING - DCG9 - 6459", 6459))
        ddlModelUser.Items.Add(New ListItem("THE TRAVEL HOUSE - MGXD - 8750", 8750))
        ddlModelUser.Items.Add(New ListItem("TRAVELCARE - EMN4 - 9336", 9336))
        ddlModelUser.Items.Add(New ListItem("FIRST CHOICE TRAVEL - A774 - 7003", 7003))
        ddlModelUser.Items.Add(New ListItem("CHEQUE CENTRE - JGKH - 11364", 11364))
        ddlModelUser.Items.Add(New ListItem("PREPAID (NEWSMARK PECKHAM) - AWYA - 7972", 7972))
        ddlModelUser.Items.Add(New ListItem("INDEPENDENT (ABI FOOD) - A4NG - 17238", 17238))
        ddlModelUser.Items.Add(New ListItem("WUIRE LOW AGENTS (MALLOW CREDIT UNION) - 7TKA - 12555", 12555))
        ddlModelUser.Items.Add(New ListItem("WUIRE HIGH AGENTS (IRELANDS EDUCATION) - 2OQ7 - 15282", 15282))
        ddlModelUser.Items.Add(New ListItem("WUIRE MEDIUM AGENTS (CASH CREATORS) -  - 15810", 15810))
        ddlModelUser.Items.Add(New ListItem("WUIRE HIGH TRANSACTOR (CALL @ NET STONEY BATTER) - LV2Q - 12778", 12778))
        ddlModelUser.Items.Add(New ListItem("WUIRE TOP AGENT (ANTECH MUNSTER) - G20C - 12671", 12671))
        ddlModelUser.Items.Add(New ListItem("WUIRE STAFF - aidan.kennerk - 12833", 12833))
    End Sub
    
    Protected Sub btnVerify_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call Verify()
    End Sub
        
    Protected Function ConvertUserIDs() As String
        Dim sUserIDs As String = tbUserIDs.Text.Trim
        sUserIDs = sUserIDs.Replace(Environment.NewLine, " ")
        sUserIDs = sUserIDs.Replace(",", " ")
        ConvertUserIDs = sUserIDs
    End Function
    
    Protected Sub Verify(Optional ByVal bVerifyUserInDropdown As Boolean = False)
        If Not IsNumeric(tbQuitLevel.Text) Then
            lblError.Text = "Value must be numeric"
            Exit Sub
        End If
        
        Dim sSQL As String

        If cbOutputDebugInformation.Checked Then
            sSQL = "DELETE FROM AAA_Debug"
            Call ExecuteQueryToDataTable(sSQL)
        End If

        sSQL = "SELECT * FROM UserProfile WHERE Type = 'User'"
        If ddlCustomer.SelectedIndex > 0 Then
            sSQL += " AND CustomerKey = " & ddlCustomer.SelectedValue
        End If
        If Not cbIncludeSuspendedUsers.Checked Then
            sSQL += " AND Status = 'Active'"
        End If
        If ddlUser.SelectedIndex > 0 Then
            sSQL += " AND [key] = " & ddlUser.SelectedValue
        End If
        
        Dim sUserIDs() As String
        tbUserIDs.Text = tbUserIDs.Text.Trim
        If ConvertUserIDs() <> String.Empty Then
            sUserIDs = ConvertUserIDs.Split(" ")
            Dim sUserList As String = String.Empty
            For Each s As String In sUserIDs
                If s.Length > 0 Then
                    Dim sType As String = String.Empty
                    Try
                        sType = ExecuteQueryToDataTable("SELECT Type FROM UserProfile WHERE UserID = '" & s & "'").Rows(0).Item(0) & String.Empty
                    Catch
                    End Try
                    If sType = String.Empty Then
                        WebMsgBox.Show("Could not find User ID " & s)
                        Exit Sub
                    Else
                        If sType.ToLower <> "user" Then
                            WebMsgBox.Show("Account " & s & " is a " & sType & " account, not a standard User account. Please remove from list.")
                            Exit Sub
                        End If
                        sUserList += "'" & s & "', "
                    End If
                End If
            Next
            sUserList += "''"
            sSQL = "SELECT * FROM UserProfile WHERE Type = 'User' AND UserID IN (" & sUserList & ")"
        End If
        If bVerifyUserInDropdown AndAlso ddlUser.SelectedIndex > 0 Then
            sSQL = "SELECT * FROM UserProfile WHERE [key] = " & ddlUser.SelectedValue
        End If
        Dim oAdapter1 As New SqlDataAdapter(sSQL, goConn)
        Dim tblUsers As New DataTable
        Dim nRecordCount As Integer = 0
        oAdapter1.Fill(tblUsers)
        
        For Each drUser As DataRow In tblUsers.Rows
            Dim tblMissingProducts As DataTable
            nTotalUsers += 1
            
            Dim nUserKey As Integer = drUser.Item("Key")
            Dim nCustomerKey As Integer = drUser.Item("CustomerKey")
            Dim oTable As DataTable = GetProductCountForCustomer(nCustomerKey)
            Dim nExpectedProductCount As Integer = oTable.Rows(0).Item(0)
            If cbOutputDebugInformation.Checked Then
                Call ExecuteQueryToDataTable("INSERT INTO AAA_Debug (Result) VALUES ('" & nTotalUsers.ToString & " UserID: " & drUser("UserID") & "')")
            End If
            Dim nMissingProductCount As Integer
            tblMissingProducts = GetMissingProductsForUser(nUserKey, nCustomerKey)
            nMissingProductCount = tblMissingProducts.Rows.Count

            If nMissingProductCount > 0 And Not (cbDontReportWhenEntireProfileMissing.Checked And nMissingProductCount = nExpectedProductCount) Then
                If cbOutputDebugInformation.Checked Then
                    Call ExecuteQueryToDataTable("INSERT INTO AAA_Debug (Result) VALUES ('Products missing')")
                End If
                nUsersWithMissingProducts += 1
                If cbGenerateSQLToFixMissingEntries.Checked Then
                    sbSQL.Append("PRINT 'Restoring values for user " & drUser.Item("UserId") & " (" & drUser.Item("Key") & ")'" & vbNewLine & vbNewLine)
                End If
                For Each drMissingProduct As DataRow In tblMissingProducts.Rows
                    sbDetail.Append("UserID: " & drUser.Item("UserId") & "(" & drUser.Item("Key") & ") missing product " & drMissingProduct.Item("ProductCode") & "(" & drMissingProduct.Item("LogisticProductKey") & ")" & vbNewLine)
                    
                    If cbGenerateSQLToFixMissingEntries.Checked Then
                        sbSQL.Append("INSERT INTO UserProductProfile (")
                        sbSQL.Append("UserKey, ProductKey, AbleToView, AbleToPick, AbleToEdit, AbleToArchive, AbleToDelete, ApplyMaxGrab, MaxGrabQty) ")
                        sbSQL.Append(" VALUES (")
                        sbSQL.Append(drUser.Item("Key") & ", " & drMissingProduct.Item("LogisticProductKey") & ", 1, 1, 1, 1, 1, 0, 0)" & vbNewLine)
                    End If
                Next
                If cbGenerateSQLToMakeActive.Checked Then
                    sbSQL.Append("UPDATE UserProfile SET Status = 'Active' WHERE Type = 'User' AND UserID = '" & drUser.Item("UserId") & "'" & vbNewLine)
                End If
                sbSummary.Append(tblMissingProducts.Rows.Count.ToString & " item(s) out of " & nExpectedProductCount & " missing for " & GetCustomerName(nCustomerKey.ToString) & " user " & drUser.Item("UserId") & "(" & drUser.Item("Key") & ")" & vbNewLine)
                If cbGenerateSQLToFixMissingEntries.Checked Then
                    sbSQL.Append(vbNewLine)
                End If
            End If
            
            If nUsersWithMissingProducts >= CInt(tbQuitLevel.Text) Then
                Exit For
            End If
            
        Next
        sbSummary.Append("Total users checked so far: " & nTotalUsers.ToString & vbNewLine)
        sbSummary.Append("Total users found with missing products: " & nUsersWithMissingProducts.ToString & vbNewLine)
        tbSummary.Text = sbSummary.ToString
        tbDetail.Text = sbDetail.ToString
        tbSQL.Text = sbSQL.ToString
        
        If bVerifyUserInDropdown AndAlso ddlUser.SelectedIndex > 0 Then
            sSQL = "SELECT * FROM UserProductProfile WHERE UserKey = " & ddlUser.SelectedValue
            Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
            gvProductProfile.Visible = True
            lnkbtnHideProductProfile.Visible = True
            gvProductProfile.DataSource = dt
            gvProductProfile.DataBind()
        End If

    End Sub
    
    Protected Function GetCustomerName(ByVal sCustomerKey As String) As String
        Dim sSQL As String
        sSQL = "SELECT CustomerAccountCode FROM Customer WHERE CustomerKey = " & sCustomerKey
        Dim oAdapter As New SqlDataAdapter(sSQL, goConn)
        Dim tblCustomerName As New DataTable
        oAdapter.Fill(tblCustomerName)
        GetCustomerName = tblCustomerName.Rows(0).Item(0)
        oAdapter.Dispose()
        tblCustomerName.Dispose()
    End Function
    
    Protected Function GetMissingProductsForUser(ByVal nUserKey As Integer, ByVal nCustomerKey As Integer) As DataTable
        Dim sSQL As String
        sSQL = "SELECT * FROM LogisticProduct WHERE CustomerKey = " & nCustomerKey.ToString & " "
        sSQL += "AND NOT LogisticProduct.LogisticProductKey IN "
        sSQL += "(SELECT ProductKey FROM UserProductProfile WHERE UserKey = " & nUserKey.ToString & ")"
        Dim oAdapter As New SqlDataAdapter(sSQL, goConn)
        Dim tblMissingProducts As New DataTable
        oAdapter.Fill(tblMissingProducts)
        Return tblMissingProducts
    End Function
    
    Protected Function GetProductCountForCustomer(ByVal nCustomerKey As Integer) As DataTable
        Dim sSQL As String = "SELECT COUNT(*) FROM LogisticProduct WHERE CustomerKey = " & nCustomerKey.ToString
        Dim oAdapter As New SqlDataAdapter(sSQL, goConn)
        Dim tblProductCount As New DataTable
        oAdapter.Fill(tblProductCount)
        Return tblProductCount
    End Function
    
    Protected Function GetProductProfileForUser(ByVal nUserKey As Integer) As DataTable
        Dim sSQL As String = "SELECT * FROM UserProductProfile WHERE UserKey = " & nUserKey.ToString
        Dim oAdapter As New SqlDataAdapter(sSQL, goConn)
        Dim tblUserProductProfile As New DataTable
        oAdapter.Fill(tblUserProductProfile)
        Return tblUserProductProfile
    End Function
    
    Protected Sub btnClear_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        tbSummary.Text = ""
        tbDetail.Text = ""
        tbSQL.Text = ""
        gvProductProfile.Visible = False
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
            tbSQL.Text += "Error in ExecuteQueryToDataTable executing: " & sQuery & " : " & ex.Message
        Finally
            oConn.Close()
        End Try
        ExecuteQueryToDataTable = oDataTable
    End Function
    
    Protected Sub ddlCustomer_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call PopulateUserDropdown()
    End Sub
    
    Protected Sub PopulateUserDropdown()
        ddlModelUser.SelectedIndex = 0
        ddlUser.Items.Clear()
        lnkbtnVerifySelectedUser.Enabled = False
        Dim sSQL As String
        If ddlCustomer.SelectedIndex > 0 Then
            sSQL = "SELECT UserID + ' (' + FirstName + ' ' + LastName + ') - ' + ISNULL(GroupName,'') UserName, [key]  FROM UserProfile up LEFT OUTER JOIN UP_UserPermissionGroups upg on up.UserGroup = upg.[id] WHERE up.CustomerKey = " & ddlCustomer.SelectedValue
            If Not cbIncludeSuspendedUsers.Checked Then
                sSQL += " AND Status = 'Active'"
            End If
            sSQL += " ORDER BY UserID"
            Dim oListItemCollection As ListItemCollection = ExecuteQueryToListItemCollection(sSQL, "UserName", "Key")
            ddlUser.Items.Clear()
            ddlUser.Items.Add(New ListItem("- all users -", 0))
            For Each li As ListItem In oListItemCollection
                ddlUser.Items.Add(li)
            Next
            lnkbtnVerifySelectedUser.Enabled = True
        End If
    End Sub
    
    Protected Sub btnCloneProfiles_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call CloneProfiles()
    End Sub
    
    Protected Function GetUserIDFromUserKey(ByVal nUserKey As Int32) As String
        GetUserIDFromUserKey = String.Empty
        Dim sSQL As String = "SELECT UserID FROM UserProfile WHERE [key] = " & nUserKey
        Dim dtUserProfile As DataTable = ExecuteQueryToDataTable(sSQL)
        If dtUserProfile.Rows.Count = 1 Then
            GetUserIDFromUserKey = dtUserProfile.Rows(0).Item(0)
        Else
            WebMsgBox.Show("GetUserIDFromUserKey: could not find record!")
        End If
    End Function
    
    Protected Sub CloneProfiles()
        Dim sSQL As String = String.Empty
        Dim sUserIDs() As String

        If ddlUser.SelectedIndex = -1 Then
            WebMsgBox.Show("Please select source customer.")
            Exit Sub
        End If
        If Not ddlUser.SelectedValue > 0 Then
            WebMsgBox.Show("Please select source account.")
            Exit Sub
        End If
        Dim sSourceUserID As String = GetUserIDFromUserKey(ddlUser.SelectedValue)
        ' sSQL = "SELECT ISNULL(UserGroup, 0) FROM UserProfile WHERE [key] = " & ddlUser.SelectedValue
        Dim nSourceUserGroup = ExecuteQueryToDataTable("SELECT ISNULL(UserGroup, 0) FROM UserProfile WHERE [key] = " & ddlUser.SelectedValue).Rows(0).Item(0)
        tbUserIDs.Text = tbUserIDs.Text.Trim
        If ConvertUserIDs.Trim <> String.Empty Then
            sUserIDs = ConvertUserIDs.Split(" ")
            Dim sUserList As String = String.Empty
            For Each s As String In sUserIDs
                If s.Length > 0 Then
                    Dim sType As String = String.Empty
                    Dim dtUser As DataTable = ExecuteQueryToDataTable("SELECT Type, CustomerKey, ISNULL(UserGroup, 0) 'UserGroup' FROM UserProfile WHERE UserID = '" & s & "'")
                    If dtUser.Rows.Count <> 1 Then
                        WebMsgBox.Show("Could not find User ID " & s)
                        Exit Sub
                    ElseIf s = sSourceUserID Then
                        WebMsgBox.Show("Destination UserID list contains source UserID! Please remove.")
                        Exit Sub
                    ElseIf Not cbOverrideCloningGroupSanityCheck.Checked Then
                        If nSourceUserGroup <> dtUser.Rows(0).Item("UserGroup") Then
                            WebMsgBox.Show("Account " & s & " has a different user group (" & dtUser.Rows(0).Item("UserGroup") & ") to that of source (" & nSourceUserGroup.ToString & ")")
                            Exit Sub
                        End If
                    End If
                    sType = dtUser.Rows(0).Item("Type")
                    If dtUser.Rows(0).Item("Type").ToString.ToLower <> "user" Then
                        WebMsgBox.Show("Account " & s & " is a " & sType & " account, not a standard User account. Please remove from list.")
                        Exit Sub
                    End If
                    If dtUser.Rows(0).Item("CustomerKey") <> ddlCustomer.SelectedValue Then
                        WebMsgBox.Show("Source account and destination account are for different customers.")
                        Exit Sub
                    End If
                    sUserList += "'" & s & "', "
                End If
            Next
            sUserList += "''"
            
            sSQL = "SELECT * FROM UserProfile WHERE Type = 'User' AND UserID IN (" & sUserList & ")"
            Dim dtDestinationUsers As DataTable = ExecuteQueryToDataTable(sSQL)
            For Each drDestinationUser As DataRow In dtDestinationUsers.Rows
                Dim nDestinationUserKey As Integer = drDestinationUser.Item("Key")
                Call CloneSingleProfile(ddlUser.SelectedValue, nDestinationUserKey)
                tbDetail.Text += "Cloned " & drDestinationUser("UserID") & " from " & ddlUser.SelectedItem.Text & vbCrLf
            Next
        End If
    End Sub

    Protected Sub CloneSingleProfile(ByVal nSourceUser As Int32, ByVal nDestinationUser As Int32)
        Dim sbSQL As New StringBuilder
        sbSQL.Append("DELETE FROM UserProductProfile WHERE UserKey = ")
        sbSQL.Append(nDestinationUser.ToString)
        sbSQL.Append(" ")
        sbSQL.Append("INSERT INTO UserProductProfile (UserKey,ProductKey,AbleToView,AbleToPick,AbleToEdit,AbleToArchive,AbleToDelete,ApplyMaxGrab,MaxGrabQty)")
        sbSQL.Append(" ")
        sbSQL.Append("SELECT ")
        sbSQL.Append(nDestinationUser.ToString)
        sbSQL.Append(", ProductKey, AbleToView, AbleToPick, AbleToEdit, AbleToArchive, AbleToDelete, ApplyMaxGrab, MaxGrabQty FROM UserProductProfile WHERE UserKey = ")
        sbSQL.Append(nSourceUser.ToString)
        Call ExecuteQueryToDataTable(sbSQL.ToString)
    End Sub

    Protected Sub lnkbtnBuildSQLString_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call BuildSQLString()
    End Sub
    
    Protected Sub BuildSQLString()
        Dim sbElementList As New StringBuilder
        Dim sUserIDs() As String

        tbUserIDs.Text = tbUserIDs.Text.Trim
        sUserIDs = ConvertUserIDs.Split(" ")
        sbElementList.Append("(")
        For Each s As String In sUserIDs
            sbElementList.Append("'" & s & "', ")
        Next
        sbElementList.Append(")")
        tbDetail.Text = sbElementList.ToString.Replace(", )", ")")
    End Sub
    
    Protected Sub ddlModelUser_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        If ddl.SelectedIndex > 0 Then
            If ddlUser.SelectedIndex >= 0 Then
                For i As Int32 = 0 To ddlUser.Items.Count - 1
                    If ddlUser.Items(i).Value = ddl.SelectedValue Then
                        ddlUser.SelectedIndex = i
                        Exit Sub
                    End If
                Next
                ddl.SelectedIndex = 0
                WebMsgBox.Show("Not found - make sure you have selected the right customer (WURS or WUIRE).")
            Else
                ddl.SelectedIndex = 0
                WebMsgBox.Show("Please select customer WURS or WUIRE")
            End If
        End If
    End Sub
    
    Protected Sub lnkbtnWURS_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        For i As Int32 = 0 To ddlCustomer.Items.Count - 1
            If ddlCustomer.Items(i).Text = "WURS" Then
                ddlCustomer.SelectedIndex = i
                Exit For
            End If
        Next
        Call PopulateUserDropdown()
    End Sub

    Protected Sub lnkbtnWUIRE_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        For i As Int32 = 0 To ddlCustomer.Items.Count - 1
            If ddlCustomer.Items(i).Text = "WUIRE" Then
                ddlCustomer.SelectedIndex = i
                Exit For
            End If
        Next
        Call PopulateUserDropdown()
    End Sub
    
    Protected Sub lnkbtnToggleHelp_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        pnlHelp.Visible = Not pnlHelp.Visible
    End Sub
    
    Protected Sub lnkbtnVerifySelectedUser_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call Verify(True)
    End Sub
    
    Protected Sub lnkbtnClearUserIDListbox_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        tbUserIDs.Text = String.Empty
    End Sub
    
    Protected Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs)
        gnTimeout = Server.ScriptTimeout
        Server.ScriptTimeout = 3600
    End Sub

    Protected Sub Page_Unload(ByVal sender As Object, ByVal e As System.EventArgs)
        Server.ScriptTimeout = gnTimeout
    End Sub
    
    Protected Sub cbOverrideCloningGroupSanityCheck_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim cb As CheckBox = sender
        If cb.Checked Then
            cb.ForeColor = Drawing.Color.Red
            cb.Font.Bold = True
        Else
            cb.ForeColor = Drawing.Color.Empty
            cb.Font.Bold = false
        End If
    End Sub
    
    Protected Sub lnkbtnHideProductProfile_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        lnkbtnHideProductProfile.Visible = False
        gvProductProfile.Visible = False
    End Sub
</script>
<html xmlns=" http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Verify Product Profiles</title>
</head>
<body style="font-family: Verdana">
    <form id="form1" runat="server">
    <strong>VERIFY / CLONE PRODUCT PROFILES build 22JUL12H</strong><br />
    <table style="font-size: x-small">
        <tr>
            <td style="width: 442px; height: 60px">
                <br />
                <br />
                &nbsp;Customer:
                <asp:DropDownList ID="ddlCustomer" runat="server" Font-Size="X-Small" AutoPostBack="True" OnSelectedIndexChanged="ddlCustomer_SelectedIndexChanged">
                </asp:DropDownList>
                &nbsp;<asp:LinkButton ID="lnkbtnWURS" runat="server" Font-Names="Arial" Font-Size="XX-Small" OnClick="lnkbtnWURS_Click">WURS</asp:LinkButton>
                &nbsp;<asp:LinkButton ID="lnkbtnWUIRE" runat="server" Font-Names="Arial" Font-Size="XX-Small" OnClick="lnkbtnWUIRE_Click">WUIRE</asp:LinkButton>
                <br />
                <br />
                &nbsp;User:&nbsp;
                <asp:LinkButton ID="lnkbtnVerifySelectedUser" runat="server" onclick="lnkbtnVerifySelectedUser_Click" Enabled="False">verify selected user</asp:LinkButton>
                <br />
                <asp:DropDownList ID="ddlUser" runat="server" Font-Size="X-Small" >
                    <asp:ListItem Value="-1">- first select a customer -</asp:ListItem>
                </asp:DropDownList>
                <br />
                &nbsp;<br />
                UserID list (comma, space or newline separated):
                <asp:LinkButton ID="lnkbtnClearUserIDListbox" runat="server" Font-Names="Arial" Font-Size="XX-Small" onclick="lnkbtnClearUserIDListbox_Click">clear</asp:LinkButton>
                <br />
                &nbsp;<asp:TextBox ID="tbUserIDs" runat="server" Font-Names="Arial" Font-Size="XX-Small" Width="95%" Rows="4" TextMode="MultiLine"></asp:TextBox>
                <br />
                <br />
                <asp:Button ID="btnVerify" runat="server" OnClick="btnVerify_Click" Text="Verify" Width="100px" />
                &nbsp;
                <asp:Button ID="btnCloneProfiles" runat="server" OnClick="btnCloneProfiles_Click" Text="Clone profile" Width="100px" />
                &nbsp;<asp:LinkButton ID="lnkbtnBuildSQLString" runat="server" Font-Names="Arial" Font-Size="XX-Small" OnClick="lnkbtnBuildSQLString_Click">build SQL string</asp:LinkButton>
                <br />
                <br />
                (clones from selected user to UserID list)<br />
                <asp:Label ID="lblError" runat="server" ForeColor="Red"></asp:Label>
                <br />
                <br />
            </td>
            <td colspan="2" style="width: 391px; height: 60px">
                &nbsp;Quit after reporting
                <asp:TextBox ID="tbQuitLevel" runat="server" Width="42px" Font-Names="Arial" Font-Size="XX-Small">10</asp:TextBox>
                affected users<br />
                <br />
                WU model user:
                <asp:DropDownList ID="ddlModelUser" runat="server" AutoPostBack="True" Font-Names="Arial" Font-Size="XX-Small" OnSelectedIndexChanged="ddlModelUser_SelectedIndexChanged"/>
                <br />
                <br />
                <asp:CheckBox ID="cbDontReportWhenEntireProfileMissing" runat="server" Text="Don't report when entire profile missing" /><br />
                <br />
                <asp:CheckBox ID="cbGenerateSQLToFixMissingEntries" runat="server" Text="Generate SQL to fix missing entries" Checked="True" />
                <br />
                <br />
                <asp:CheckBox ID="cbGenerateSQLToMakeActive" runat="server" Text="Generate SQL to make Active" Checked="True" />
                <br />
                <br />
                <asp:CheckBox ID="cbOverrideCloningGroupSanityCheck" runat="server" Text="Override cloning group sanity check" AutoPostBack="True" oncheckedchanged="cbOverrideCloningGroupSanityCheck_CheckedChanged" />
                <br />
                <br />
                <asp:CheckBox ID="cbIncludeSuspendedUsers" runat="server" Text="Include suspended users " AutoPostBack="True" />
                <br />
                <br />
                <asp:CheckBox ID="cbOutputDebugInformation" runat="server" Text="Output debug information" />
            </td>
        </tr>
    </table>
    Summary:<br />
    <asp:TextBox ID="tbSummary" runat="server" Height="224px" TextMode="MultiLine" Width="100%" Font-Names="Arial" Font-Size="XX-Small" /><br />
    Detail:<br />
    <asp:TextBox ID="tbDetail" runat="server" Height="224px" TextMode="MultiLine" Width="100%" Font-Names="Arial" Font-Size="XX-Small" /><br />
    SQL:<br />
    <asp:TextBox ID="tbSQL" runat="server" Height="224px" TextMode="MultiLine" Width="100%" Font-Names="Arial" Font-Size="XX-Small" /><br />
    <asp:LinkButton ID="lnkbtnHideProductProfile" runat="server" Font-Names="Arial" Font-Size="XX-Small" onclick="lnkbtnHideProductProfile_Click">hide product profile</asp:LinkButton>
    <br />
    <asp:GridView ID="gvProductProfile" runat="server" CellPadding="2" Font-Names="Arial" Font-Size="XX-Small" Width="100%">
    </asp:GridView>
    <br />
    &nbsp;<asp:Button ID="btnClear" runat="server" OnClick="btnClear_Click" Text="Clear" />
    &nbsp;<asp:LinkButton ID="lnkbtnToggleHelp" runat="server" Font-Names="Arial" Font-Size="XX-Small" onclick="lnkbtnToggleHelp_Click">toggle help</asp:LinkButton>
    <asp:Panel ID="pnlHelp" runat="server" Width="100%" Font-Names="Verdana" Font-Size="X-Small" Visible="false">
        <br />
        <strong>HELP</strong>
        <br />
        <br />
        Useful Queries<br />
        <br />
        &nbsp; SELECT GroupName, up.* FROM UserProfile up LEFT OUTER JOIN UP_UserPermissionGroups upg ON up.UserGroup = upg.[id] WHERE UserID IN (<em>user list</em>) ORDER BY GroupName, UserID<br />
        <br />
        Verify
        <br />
        <br />
        Checks the validity of user product profiles.
        <br />
        <br />
        Only accounts of type User are checked. If a customer is selected from the Customer dropdown, only that customers users are checked. Only Active users are checked unless Include Suspended Users is checked. If a user is selected from the User dropdown, only that user is checked.
        <br />
        <br />
        If there is a list of one or more user IDs in the UserID box, this takes precedence, and these users are checked instead. The module reports if any of the user IDs cannot be found, or if an account that is not a User account is included in the list,
        <br />
        <br />
        Validity is checked by comparing the customers total product count with the number of Product Profile records for this customer, then listing the customer&#39;s products that are not in the user&#39;s Product Profile. The number of users found so far with missing products is tallied and the module stops if the tally reaches the limit set in &#39;Quit after reporting n affected users&#39;.
        <br />
        <br />
        If &#39;Generate SQL to fix missing entries&#39; is checked, SQL statements are generated to recreate the missing entries, and initialised to default values.
        <br />
        <br />
        If &#39;Generate SQL to make Active&#39; is checked, SQL statements are generated to set the user accounts to Active.
        <br />
        <br />
        The module lists the number of missing items for each user, and the overall totals processed.
        <br />
        <br />
        Clone Profile
        <br />
        <br />
        Copies a &#39;model&#39; profile to one or more other users with UserIDs listed in the UserID box. The source account is selected from the Customer / User drop down boxes. An alert is shown if the source user is in the destination list.
        <br />
        <br />
        If both the source user account and the destination user account belong to a User Group, the module checks they are in the same user group, unless &#39;Override cloning group sanity check&#39; is checked. Destination UserIDs that cannot be found are reported, as are UserIDs that are not standard User accounts. Source and destination accounts must be for the same customer.
        <br />
        <br />
        Build SQL String
        <br />
        <br />
        WU Model Users
        <br />
        <br />
        Lists the WURS &amp; WUIRE accounts to be used as &#39;models&#39; from which other users in the same user group are cloned. This list is hard-wired into the tool.
        <br />
        <br />
        Don&#39;t report when entire profile missing</asp:Panel>
    </form>
</body>
</html>
