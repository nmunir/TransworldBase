<%@ Page Language="VB" Theme="AIMSDefault" ValidateRequest="false" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>
<%@ import Namespace="System.Collections.Generic" %>

<script runat="server">

    ' NOTE: Currently because of a design oversight you can only associate a single product group with a user group

    Const RANGE_START_GROUPS As Integer = 1000000
   
    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString
  
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsNumeric(Session("CustomerKey")) Then
            Server.Transfer("session_expired.aspx")
            Exit Sub
        End If
        If Not IsPostBack Then
            Call InitProductGroupListBox()
            Call InitUsersPanel()
            tbNewProductGroupName.Attributes.Add("onkeypress", "return clickButton(event,'" + btnCreateNewProductGroup.ClientID + "')")
            tbRenameProductGroupNewName.Attributes.Add("onkeypress", "return clickButton(event,'" + btnDoRenameProductGroup.ClientID + "')")
            tbNewUserGroupName.Attributes.Add("onkeypress", "return clickButton(event,'" + btnCreateNewUserGroup.ClientID + "')")
            tbRenameUserGroupNewName.Attributes.Add("onkeypress", "return clickButton(event,'" + btnDoRenameUserGroup.ClientID + "')")
        End If
        Call SetTitle()
    End Sub
  
    Protected Sub SetTitle()
        Dim sTitle As String = Session("SiteTitle")
        If sTitle <> String.Empty Then
            sTitle += " - "
        End If
        Page.Header.Title = sTitle & "User Permissions"
    End Sub
  
    Protected Sub HideAllPanels()
        pnlProductGroups.Visible = False
        pnlUserGroups.Visible = False
        pnlNewUserGroup.Visible = False
        pnlDefaultGroupAssociation.Visible = False
        pnlTweakUsers.Visible = False
        pnlNewProductGroup.Visible = False
        pnlRenameProductGroup.Visible = False
        pnlHelp.Visible = False
        pnlMaxGrabs.Visible = False
    End Sub
   
    Protected Sub btnProductGroups_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call HideAllPanels()
        pnlProductGroups.Visible = True
        rblOrderBy.SelectedIndex = 0
        lbProductGroups.SelectedIndex = -1
        lbAvailableProductsForProductGroup.Items.Clear()
        lbProductsInProductGroup.Items.Clear()
    End Sub
   
    Protected Sub btnUsers_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call HideAllPanels()
        pnlTweakUsers.Visible = True
    End Sub
   
    Protected Sub btnNewProductGroup_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call HideAllPanels()
        tbNewProductGroupName.Text = String.Empty
        rbOnCreateNoAction.Checked = True
        rbOnCreateAddGroupAllUsers.Checked = False
        rbOnCreateAddGroupUserGroup.Checked = False
        pnlNewProductGroup.Visible = True
        tbNewProductGroupName.Focus()
        lbProductGroups.SelectedIndex = -1
        SetEnableRenameAndRemoveProductGroup(False)
        ddlOnCreateUserGroups.Visible = False
    End Sub
    
    Protected Sub SetEnableRenameAndRemoveProductGroup(ByVal bStatus As Boolean)
        btnRenameProductGroup.Enabled = bStatus
        btnRemoveProductGroup.Enabled = bStatus
    End Sub
   
    Protected Sub SetEnableRenameAndRemoveUserGroup(ByVal bStatus As Boolean)
        btnRenameUserGroup.Enabled = bStatus
        btnRemoveUserGroup.Enabled = bStatus
    End Sub
   
    Protected Sub btnCreateNewProductGroup_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        tbNewProductGroupName.Text = tbNewProductGroupName.Text.Trim
        If tbNewProductGroupName.Text <> String.Empty Then
            Call CreateNewProductGroup()
        Else
            WebMsgBox.Show("Please specify a name.")
        End If
    End Sub
   
    Protected Sub CreateNewProductGroup()
        If rbOnCreateAddGroupUserGroup.Checked AndAlso ddlOnCreateUserGroups.SelectedIndex = 0 Then
            WebMsgBox.Show("Please select a user group.")
            Exit Sub
        End If
       
        tbNewProductGroupName.Text = tbNewProductGroupName.Text.Trim
        If GetProductGroupKeyFromName(tbNewProductGroupName.Text) > 0 Then
            WebMsgBox.Show("This name is already in use! Please choose another name.")
        Else
            Dim oConn As New SqlConnection(gsConn)
            Dim oCmd As SqlCommand
            Dim sSQL As String = "INSERT INTO UP_ProductPermissionGroups (CustomerKey, ProductGroup, LastModifiedDateTime, LastUpdateBy) VALUES ("
            sSQL += Session("CustomerKey") & ", '" & tbNewProductGroupName.Text.Replace("'", "''") & "', GETDATE(), " & Session("UserKey") & ")"
            Try
                oConn.Open()
                oCmd = New SqlCommand(sSQL, oConn)
                oCmd.ExecuteNonQuery()

                If rbOnCreateAddGroupAllUsers.Checked Then
                    Call ApplyGroupToAllUsers(GetProductGroupKeyFromName(tbNewProductGroupName.Text.Replace("'", "''")))
                End If
               
                If rbOnCreateAddGroupUserGroup.Checked Then
                    If ddlOnCreateUserGroups.SelectedIndex > 0 Then
                        Call ApplyGroupToUsersInUserGroup(GetProductGroupKeyFromName(tbNewProductGroupName.Text.Replace("'", "''")), ddlOnCreateUserGroups.SelectedValue)
                    End If
                End If
           
            Catch ex As Exception
                WebMsgBox.Show("Error in CreateNewProductGroup: " & ex.Message)
            Finally
                oConn.Close()
            End Try
           
            Call HideAllPanels()
            Call InitProductGroupListBox()
            Call InitUserPermissioningListBoxes()
            pnlProductGroups.Visible = True
        End If
    End Sub
   
    Protected Sub ApplyGroupToAllUsers(ByVal nProductGroupKey As Integer)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand
        Dim sSQL As String = "INSERT INTO UP_UserPermissions (UserKey, GroupOrProductKey, LastModifiedDateTime, LastUpdateBy) SELECT [Key], " & nProductGroupKey & ", GETDATE(), " & Session("UserKey") & " FROM UserProfile WHERE Type = 'User' AND CustomerKey = " & Session("CustomerKey")
        Try
            oConn.Open()
            oCmd = New SqlCommand(sSQL, oConn)
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            WebMsgBox.Show("Error in ApplyGroupToAllUsers: " & ex.Message)
        Finally
            oConn.Close()
        End Try
        Call RefreshPermissions(0, Session("CustomerKey"), 0)
    End Sub
   
    Protected Sub ApplyGroupToUsersInUserGroup(ByVal nProductGroupKey As Integer, ByVal sUserGroup As Integer)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand
        'Dim sSQL As String = "INSERT INTO UP_UserPermissions (UserKey, GroupOrProductKey, LastModifiedDateTime, LastUpdateBy) SELECT [Key], " & nProductGroupKey & ", GETDATE(), " & Session("UserKey") & " FROM UserProfile WHERE Type = 'User' AND CustomerKey = " & Session("CustomerKey") & " AND DXepartment = '" & sUserGroup.Replace("'", "''") & "'"
        Dim sSQL As String = "INSERT INTO UP_UserPermissions (UserKey, GroupOrProductKey, LastModifiedDateTime, LastUpdateBy) SELECT [Key], " & nProductGroupKey & ", GETDATE(), " & Session("UserKey") & " FROM UserProfile WHERE Type = 'User' AND CustomerKey = " & Session("CustomerKey") & " AND UserGroup = " & sUserGroup
        Try
            oConn.Open()
            oCmd = New SqlCommand(sSQL, oConn)
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            WebMsgBox.Show("Error in ApplyGroupToUsersInUserGroup: " & ex.Message)
        Finally
            oConn.Close()
        End Try
        Call RefreshPermissions(0, 0, nProductGroupKey)
    End Sub
   
    Protected Function GetProductGroupKeyFromName(ByVal sGroupName As String) As Integer
        GetProductGroupKeyFromName = 0
        Dim oDataReader As SqlDataReader = Nothing
        Dim oConn As New SqlConnection(gsConn)
        Dim sSQL As String = "SELECT id FROM UP_ProductPermissionGroups WHERE CustomerKey = " & Session("CustomerKey") & " AND ProductGroup = '" & sGroupName.Replace("'", "''") & "'"
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            If oDataReader.HasRows Then
                oDataReader.Read()
                GetProductGroupKeyFromName = oDataReader("id")
            End If
        Catch ex As Exception
            WebMsgBox.Show("Error in GetProductGroupKeyFromName: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Function
   
    Protected Function GetUserGroupKeyFromName(ByVal sGroupName As String) As Integer
        GetUserGroupKeyFromName = 0
        Dim oDataReader As SqlDataReader = Nothing
        Dim oConn As New SqlConnection(gsConn)
        Dim sSQL As String = "SELECT id FROM UP_UserPermissionGroups WHERE CustomerKey = " & Session("CustomerKey") & " AND GroupName = '" & sGroupName.Replace("'", "''") & "'"
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            If oDataReader.HasRows Then
                oDataReader.Read()
                GetUserGroupKeyFromName = oDataReader("id")
            End If
        Catch ex As Exception
            WebMsgBox.Show("Error in GetUserGroupKeyFromName: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Function
   
    Protected Sub btnCancelNewProductGroup_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call HideAllPanels()
        pnlProductGroups.Visible = True
    End Sub
   
    Protected Sub InitProductGroupListBox()
        Dim oListItemCollection As ListItemCollection = ExecuteQuery("SELECT ProductGroup, [id] FROM UP_ProductPermissionGroups WHERE CustomerKey = " & Session("CustomerKey") & " ORDER BY ProductGroup", "ProductGroup", "id")
        lbProductGroups.Items.Clear()
        For Each li As ListItem In oListItemCollection
            lbProductGroups.Items.Add(li)
        Next
    End Sub
   
    Protected Sub lbProductGroups_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call PopulateProductsListBoxes()
        SetEnableRenameAndRemoveProductGroup(True)
    End Sub
   
    Protected Sub PopulateProductsListBoxes()
        Dim sGroupName As String = lbProductGroups.SelectedItem.Text
        Call PopulateAvailableProductsListBox(sGroupName)
        Call PopulateProductsInGroupListBox(sGroupName)
        Call InitCopyProductsDropdown()
    End Sub
   
    Protected Sub PopulateAvailableProductsListBox(ByVal sGroupName As String)
        Dim sSQL As String = "SELECT LogisticProductKey, ProductCode + ' - ' + ISNULL(ProductDate,'') + ': ' + ProductDescription 'Product', ArchiveFlag FROM LogisticProduct WHERE CustomerKey = " & Session("CustomerKey") & " AND NOT LogisticProductKey IN (SELECT pgm.LogisticProductKey FROM UP_ProductGroupsMapping pgm INNER JOIN UP_ProductPermissionGroups pg ON pg.[id] = pgm.ProductGroupKey INNER JOIN LogisticProduct lp ON pgm.LogisticProductKey = lp.LogisticProductKey WHERE pg.CustomerKey = " & Session("CustomerKey") & " AND pg.ProductGroup = '" & sGroupName.Replace("'", "''") & "') "
        If rblOrderBy.SelectedItem.ToString.ToLower = "product code" Then
            sSQL += "ORDER BY ProductCode"
        ElseIf rblOrderBy.SelectedItem.ToString.ToLower = "category" Then
            sSQL += "ORDER BY ProductCategory, ProductCode"
        Else
            sSQL += "ORDER BY ArchiveFlag, ProductCode"
        End If
        Dim oListItemCollection As ListItemCollection = ExecuteQueryFlagArchivedProducts(sSQL, "Product", "LogisticProductKey")
        lbAvailableProductsForProductGroup.Items.Clear()
        For Each li As ListItem In oListItemCollection
            lbAvailableProductsForProductGroup.Items.Add(li)
        Next
    End Sub
   
    Protected Sub rblOrderBy_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim rbl As RadioButtonList = sender
        If lbProductGroups.SelectedIndex >= 0 Then
            Call PopulateAvailableProductsListBox(lbProductGroups.SelectedItem.Text)
        End If
    End Sub
   
    Protected Sub PopulateProductsInGroupListBox(ByVal sGroupName As String)
        'Dim sSQL As String = "SELECT pgm.LogisticProductKey, ProductCode + ' - ' + ISNULL(ProductDate,'') + ': ' + ProductDescription 'Product' FROM ProductGroupsMapping pgm INNER JOIN ProductPermissionGroups pg ON pg.[id] = pgm.ProductGroupKey INNER JOIN LogisticProduct lp ON pgm.LogisticProductKey = lp.LogisticProductKey WHERE pg.CustomerKey = " & Session("CustomerKey") & " AND pg.ProductGroup = '" & sGroupName.Replace("'", "''") & "' ORDER BY ArchiveFlag, ProductCode"
        Dim oListItemCollection As ListItemCollection = ExecuteQueryFlagArchivedProducts("SELECT pgm.LogisticProductKey, ProductCode + ' - ' + ISNULL(ProductDate,'') + ': ' + ProductDescription 'Product', ArchiveFlag FROM UP_ProductGroupsMapping pgm INNER JOIN UP_ProductPermissionGroups pg ON pg.[id] = pgm.ProductGroupKey INNER JOIN LogisticProduct lp ON pgm.LogisticProductKey = lp.LogisticProductKey WHERE pg.CustomerKey = " & Session("CustomerKey") & " AND pg.ProductGroup = '" & sGroupName.Replace("'", "''") & "' ORDER BY ArchiveFlag, ProductCode", "Product", "LogisticProductKey")
        lbProductsInProductGroup.Items.Clear()
        For Each li As ListItem In oListItemCollection
            lbProductsInProductGroup.Items.Add(li)
        Next
    End Sub
   
    Protected Function ExecuteQuery(ByVal sQuery As String, ByVal sTextFieldName As String, ByVal sValueFieldName As String) As ListItemCollection
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
            WebMsgBox.Show("Error in ExecuteQuery: " & ex.Message)
        Finally
            oConn.Close()
        End Try
        ExecuteQuery = oListItemCollection
    End Function

    Protected Function ExecuteQueryFlagArchivedProducts(ByVal sQuery As String, ByVal sTextFieldName As String, ByVal sValueFieldName As String) As ListItemCollection
        Dim oListItemCollection As New ListItemCollection
        Dim oDataReader As SqlDataReader = Nothing
        Dim oConn As New SqlConnection(gsConn)
        Dim sTextField As String
        Dim sValueField As String
        Dim oCmd As SqlCommand = New SqlCommand(sQuery, oConn)
        Dim bFirstArchivedProductFound As Boolean = False
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            If oDataReader.HasRows Then
                While oDataReader.Read
                    If sQuery.Contains("ORDER BY ArchiveFlag, ProductCode") Then
                        If Not bFirstArchivedProductFound Then
                            If Not IsDBNull(oDataReader("ArchiveFlag")) Then
                                If oDataReader("ArchiveFlag").ToString.ToUpper = "Y" Then
                                    oListItemCollection.Add(New ListItem("", 0))
                                    oListItemCollection.Add(New ListItem("- ARCHIVED PRODUCTS -", 0))
                                    oListItemCollection.Add(New ListItem("", 0))
                                    bFirstArchivedProductFound = True
                                End If
                            End If
                        End If
                    End If
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
            WebMsgBox.Show("Error in ExecuteQuery: " & ex.Message)
        Finally
            oConn.Close()
        End Try
        ExecuteQueryFlagArchivedProducts = oListItemCollection
    End Function

    Protected Sub lnkbtnRemoveProductFromGroup_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call RemoveProductFromGroup()
    End Sub

    Protected Sub lnkbtnAddProductToGroup_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If lbProductGroups.SelectedIndex >= 0 Then
            Call AddProductToGroup()
        End If
    End Sub
   
    Protected Sub RemoveProductFromGroup()
        For Each item As ListItem In lbProductsInProductGroup.Items
            If item.Selected AndAlso item.Value > 0 Then
                Call RemoveProductFromGroup(item.Value)
            End If
        Next
        Call UpdateUsersOfSelectedGroup(lbProductGroups.SelectedValue)
        Call PopulateProductsListBoxes()
    End Sub

    Protected Sub RemoveProductFromGroup(ByVal nLogisticProductKey As Integer)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand
        Dim sSQL As String = "DELETE FROM UP_ProductGroupsMapping WHERE ProductGroupKey = " & lbProductGroups.SelectedValue & " AND LogisticProductKey = " & nLogisticProductKey
        Try
            oConn.Open()
            oCmd = New SqlCommand(sSQL, oConn)
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            WebMsgBox.Show("Error in RemoveProductFromGroup: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub

    Protected Sub UpdateUsersOfSelectedGroup(ByVal nGroupKey As Integer)
        Call RefreshPermissions(0, 0, nGroupKey)
    End Sub
   
    Protected Sub AddProductToGroup()
        For Each item As ListItem In lbAvailableProductsForProductGroup.Items
            If item.Selected AndAlso item.Value > 0 Then
                Call AddProductToGroup(item.Value)
            End If
        Next
        Call UpdateUsersOfSelectedGroup(lbProductGroups.SelectedValue)
        Call PopulateProductsListBoxes()
    End Sub

    Protected Function IsProductInGroup(ByVal nLogisticProductKey As Integer, ByVal nGroupKey As Integer) As Boolean
        IsProductInGroup = False
        Dim sSQL As String = "SELECT * FROM UP_ProductGroupsMapping WHERE LogisticProductKey = " & nLogisticProductKey & " AND ProductGroupKey = " & nGroupKey
        Dim oDataReader As SqlDataReader = Nothing
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            If oDataReader.HasRows Then
                IsProductInGroup = True
            End If
        Catch ex As Exception
            WebMsgBox.Show("Error in IsProductInGroup: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Function
   
    Protected Sub AddProductToGroup(ByVal LogisticProductKey As Integer)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand
        Dim sSQL As String = "INSERT INTO UP_ProductGroupsMapping (ProductGroupKey, LogisticProductKey, LastModifiedDateTime, LastUpdateBy) VALUES ("
        sSQL += lbProductGroups.SelectedValue & ", " & LogisticProductKey & ", GETDATE(), " & Session("UserKey") & ")"
        Try
            oConn.Open()
            oCmd = New SqlCommand(sSQL, oConn)
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            WebMsgBox.Show("Error in AddProductToGroup: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub

    Protected Sub btnRemoveProductGroup_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If lbProductGroups.SelectedIndex >= 0 Then
            Call RemoveProductGroup()
            Call UpdateUsersOfSelectedGroup(lbProductGroups.SelectedValue)
            Call InitProductGroupListBox()
            lbAvailableProductsForProductGroup.Items.Clear()
            lbProductsInProductGroup.Items.Clear()
            Call InitUserPermissioningListBoxes()
            lbProductGroups.SelectedIndex = -1
            SetEnableRenameAndRemoveProductGroup(False)
            Call InitCopyProductsDropdown()
        Else
            WebMsgBox.Show("No group selected!")
        End If
    End Sub
   
    Protected Sub RemoveProductGroup()
        Dim sGroupKey As Integer = lbProductGroups.SelectedValue
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand
        Dim sSQL As String
        Try
            oConn.Open()
            sSQL = "DELETE FROM UP_UserPermissions WHERE GroupOrProductKey = " & sGroupKey
            oCmd = New SqlCommand(sSQL, oConn)
            oCmd.ExecuteNonQuery()
            sSQL = "DELETE FROM UP_ProductGroupsMapping WHERE ProductGroupKey = " & sGroupKey
            oCmd = New SqlCommand(sSQL, oConn)
            oCmd.ExecuteNonQuery()
            sSQL = "DELETE FROM UP_ProductPermissionGroups WHERE id = " & sGroupKey
            oCmd = New SqlCommand(sSQL, oConn)
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            WebMsgBox.Show("Error in RemoveGroup: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub

    Protected Sub InitUsersPanel()
        Call PopulateUsersListBox()
    End Sub

    Protected Sub PopulateUsersListBox()
        Dim sSQL As String = "SELECT [Key] UserKey, FirstName + ' ' + LastName + ' (' + UserId + ')' UserName FROM UserProfile WHERE Type = 'User' AND CustomerKey = " & Session("CustomerKey")
        Dim oListItemCollection As ListItemCollection = ExecuteQuery(sSQL, "UserName", "UserKey")
        lbUsers.Items.Clear()
        For Each li As ListItem In oListItemCollection
            lbUsers.Items.Add(li)
        Next
    End Sub

    Protected Sub lbUsers_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call InitUserPermissioningDisplay()
    End Sub
   
    Protected Sub InitUserPermissioningDisplay()
        lblLegendUserNamePermissioning.Text = lbUsers.SelectedItem.Text & "'s groups & products:"
        pnUserKey = lbUsers.SelectedValue
        Call InitUserPermissioningListBoxes()
    End Sub
   
    Protected Sub InitUserPermissioningListBoxes()
        If lbUsers.SelectedIndex >= 0 Then
            Call InitUserPermissioningListBox()
            Call InitAvailableGroupsForUserListBox()
            Call InitAvailableProductsForUserListBox()
        End If
    End Sub
   
    Protected Sub InitUserPermissioningListBox()
        Dim sSQL As String = "SELECT LogisticProductKey 'GroupOrProductKey', ProductCode + ' - ' + ISNULL(ProductDate,'') + ': ' + ProductDescription 'Description' FROM LogisticProduct WHERE LogisticProductKey IN (SELECT GroupOrProductKey FROM UP_UserPermissions WHERE UserKey = " & pnUserKey & ") UNION ALL SELECT GroupOrProductKey, ProductGroup 'Description' FROM UP_UserPermissions INNER JOIN UP_ProductPermissionGroups pg ON GroupOrProductKey = pg.id WHERE UserKey = " & pnUserKey
        Dim oListItemCollection As ListItemCollection = ExecuteQuery(sSQL, "Description", "GroupOrProductKey")
        lbPermissionedGroupsAndProducts.Items.Clear()
        For Each li As ListItem In oListItemCollection
            lbPermissionedGroupsAndProducts.Items.Add(li)
        Next
    End Sub
   
    Protected Sub InitAvailableGroupsForUserListBox()
        Dim sSQL As String = "SELECT id, ProductGroup FROM UP_ProductPermissionGroups WHERE CustomerKey = " & Session("CustomerKey") & " AND NOT id IN (SELECT GroupOrProductKey FROM UP_UserPermissions WHERE UserKey = " & pnUserKey & ")"
        Dim oListItemCollection As ListItemCollection = ExecuteQuery(sSQL, "ProductGroup", "id")
        lbAvailableGroupsForUser.Items.Clear()
        For Each li As ListItem In oListItemCollection
            lbAvailableGroupsForUser.Items.Add(li)
        Next
    End Sub

    Protected Sub InitAvailableProductsForUserListBox()
        Dim sSQL As String = "SELECT LogisticProductKey 'GroupOrProductKey', ProductCode + ' - ' + ISNULL(ProductDate,'') + ': ' + ProductDescription 'Description' FROM LogisticProduct WHERE CustomerKey = " & Session("CustomerKey") & " AND NOT LogisticProductKey IN (SELECT GroupOrProductKey FROM UP_UserPermissions WHERE UserKey = " & pnUserKey & ")"
        Dim oListItemCollection As ListItemCollection = ExecuteQuery(sSQL, "Description", "GroupOrProductKey")
        lbAvailableProductsForUser.Items.Clear()
        For Each li As ListItem In oListItemCollection
            lbAvailableProductsForUser.Items.Add(li)
        Next
    End Sub

    Property pnUserKey() As Integer
        Get
            Dim o As Object = ViewState("UP_UserKey")
            If o Is Nothing Then
                Return 0
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Integer)
            ViewState("UP_UserKey") = Value
        End Set
    End Property
 
    Protected Sub lnkbtnAddGroupToUser_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If lbUsers.SelectedIndex >= 0 Then
            Call AddGroupsToUser(lbUsers.SelectedValue)
        End If
    End Sub

    Protected Sub AddGroupsToUser(ByVal sUserKey As String)
        For Each item As ListItem In lbAvailableGroupsForUser.Items
            If item.Selected Then
                Call AddGroupToUser(item.Value, sUserKey)
            End If
        Next
        Call UpdateSelectedUser(sUserKey)
        Call InitUserPermissioningListBoxes()
    End Sub
 
    Protected Sub AddGroupToUser(ByVal sGroupKey As String, ByVal sUserKey As String)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand
        Dim sSQL As String = "INSERT INTO UP_UserPermissions (UserKey, GroupOrProductKey, LastModifiedDateTime, LastUpdateBy) VALUES ("
        sSQL += sUserKey & ", " & sGroupKey & ", GETDATE(), " & Session("UserKey") & ")"
        Try
            oConn.Open()
            oCmd = New SqlCommand(sSQL, oConn)
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            WebMsgBox.Show("Error in AddGroupToUser: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
 
    Protected Sub lnkbtnRemoveGroupFromUser_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If lbUsers.SelectedIndex >= 0 Then
            Call RemoveGroupsOrProductsFromUser(lbUsers.SelectedValue)
        End If
    End Sub
   
    Protected Sub RemoveGroupsOrProductsFromUser(ByVal sUserKey As String)
        For Each item As ListItem In lbPermissionedGroupsAndProducts.Items
            If item.Selected Then
                Call RemoveGroupOrProductFromUser(item.Value, sUserKey)
            End If
        Next
        Call UpdateSelectedUser(sUserKey)
        Call InitUserPermissioningListBoxes()
    End Sub
   
    Protected Sub RemoveGroupOrProductFromUser(ByVal nGroupOrProductKey As Integer, ByVal sUserKey As String)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand
        Dim sSQL As String
        Try
            oConn.Open()
            sSQL = "DELETE FROM UP_UserPermissions WHERE GroupOrProductKey = " & nGroupOrProductKey & " AND UserKey = " & lbUsers.SelectedValue
            oCmd = New SqlCommand(sSQL, oConn)
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            WebMsgBox.Show("Error in RemoveGroupOrProductFromUser: " & ex.Message)
        Finally
            oConn.Close()
        End Try

    End Sub
   
    Protected Sub lnkbtnAddProductToUser_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If lbUsers.SelectedIndex >= 0 Then
            Call AddProductsToUser(lbUsers.SelectedValue)
        End If
    End Sub
   
    Protected Sub AddProductsToUser(ByVal sUserKey As String)
        For Each item As ListItem In lbAvailableProductsForUser.Items
            If item.Selected Then
                Call AddProductToUser(item.Value)
            End If
        Next
        Call UpdateSelectedUser(sUserKey)
        Call InitUserPermissioningListBoxes()
    End Sub
   
    Protected Sub AddProductToUser(ByVal nLogisticProductKey As Integer)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand
        Dim sSQL As String = "INSERT INTO UP_UserPermissions (UserKey, GroupOrProductKey, LastModifiedDateTime, LastUpdateBy) VALUES ("
        sSQL += lbUsers.SelectedValue & ", " & nLogisticProductKey & ", GETDATE(), " & Session("UserKey") & ")"
        Try
            oConn.Open()
            oCmd = New SqlCommand(sSQL, oConn)
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            WebMsgBox.Show("Error in AddProductToUser: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
   
    Protected Sub lnkbtnRemoveProductFromUser_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If lbUsers.SelectedIndex >= 0 Then
            Call RemoveGroupsOrProductsFromUser(lbUsers.SelectedValue)
        End If
    End Sub

    Protected Sub UpdateAllUsers()
        Call RefreshPermissions(0, Session("CustomerKey"), 0)
    End Sub
   
    Protected Sub UpdateSelectedUser(ByVal sUserKey As String)
        Call RefreshPermissions(sUserKey, 0, 0)
    End Sub
   
    Protected Sub RefreshPermissions(ByVal sUserKey As String, ByVal nCustomerKey As Integer, ByVal nProductGroupKey As Integer)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_UserPermissions_Apply2", oConn)
        oCmd.CommandTimeout = 300
        oCmd.CommandType = CommandType.StoredProcedure
        Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int)
        paramCustomerKey.Value = nCustomerKey
        oCmd.Parameters.Add(paramCustomerKey)
        Dim paramUserKey As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int)
        paramUserKey.Value = sUserKey
        oCmd.Parameters.Add(paramUserKey)
        Dim paramProductGroupKey As SqlParameter = New SqlParameter("@ProductGroupKey", SqlDbType.Int)
        paramProductGroupKey.Value = nProductGroupKey
        oCmd.Parameters.Add(paramProductGroupKey)
        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As SqlException
            WebMsgBox.Show("Error in RefreshPermissions: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
   
    Protected Sub btnRefreshAllUserPermissions_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call UpdateAllUsers()
    End Sub
   
    Protected Sub ClearUserList()
        lbUsers.Items.Clear()
        lblLegendUsers.Text = ""
    End Sub
   
    Protected Sub rbUsersAll_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim rb As RadioButton = sender
        If rb.Checked Then
            Call HideUserControls()
            Call ClearUserList()
        End If
    End Sub

    Protected Sub rbUsersWithoutPermissions_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim rb As RadioButton = sender
        If rb.Checked Then
            Call HideUserControls()
            Call ClearUserList()
        End If
    End Sub

    Protected Sub rbUsersMatching_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim rb As RadioButton = sender
        If rb.Checked Then
            Call HideUserControls()
            tbUserSearch.Text = String.Empty
            Call ShowUserTextBox()
            Call ClearUserList()
        End If
    End Sub

    Protected Sub rbUsersInGroup_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim rb As RadioButton = sender
        If rb.Checked Then
            Call HideUserControls()
            Call InitUserGroupDropdown()
            ddlUserGroup.SelectedIndex = 0
            Call ShowGroupDropdown()
            Call ClearUserList()
        End If
    End Sub
   
    Protected Sub HideUserControls()
        tbUserSearch.Visible = False
        ddlUserGroup.Visible = False
    End Sub
   
    Protected Sub ShowUserTextBox()
        tbUserSearch.Visible = True
        tbUserSearch.Focus()
    End Sub
   
    Protected Sub ShowGroupDropdown()
        ddlUserGroup.Visible = True
        ddlUserGroup.Focus()
    End Sub
   
    Protected Sub InitUserGroupDropdown()           ' SWAP FOR LIST in db
        If ddlUserGroup.Items.Count = 0 Then
            'Dim oListItemCollection As ListItemCollection = ExecuteQuery("SELECT DISTINCT DXepartment FROM UserProfile WHERE CustomerKey = " & Session("CustomerKey") & " AND DXepartment <> '' ORDER BY DXepartment", "DXepartment", "DXepartment")
            Dim oListItemCollection As ListItemCollection = ExecuteQuery("SELECT GroupName, [id] FROM UP_UserPermissionGroups WHERE CustomerKey = " & Session("CustomerKey") & " ORDER BY GroupName", "GroupName", "id")
            ddlUserGroup.Items.Clear()
            ddlUserGroup.Items.Add(New ListItem("- please select -", 0))
            For Each li As ListItem In oListItemCollection
                ddlUserGroup.Items.Add(li)
            Next
        End If
    End Sub
   
    Protected Sub btnUsersReport_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim sbText As New StringBuilder
        Call AddHTMLPreamble(sbText, "Users Report")
        sbText.Append(Bold("USERS REPORT"))
        Call NewLine(sbText)
        sbText.Append("report generated " & DateTime.Now)
        Call NewLine(sbText)
        Call NewLine(sbText)
        sbText.Append("This report is divided into 6 sections. <b>Section 1</b> shows the defined user groups and the associated product group, if any. <b>Section 2</b> lists the users who are not in any user group. <b>Section 3</b> shows, for each defined user group, the users in that group. <b>Section 4</b> lists Super Users. <b>Section 5</b> shows, for each User, the permission group(s) and product(s) assigned to that user. <b>Section 6</b> is a summary list of users who are not permissioned at all.")
        Call NewLine(sbText)
        Call NewLine(sbText)
        sbText.Append("<i>NOTE: The permissions section of this report covers standard Users only. Super Users are automatically permissioned for all products. A list of Super Users is shown in section 4.</i>")
        Call NewLine(sbText)
        Call NewLine(sbText)
        sbText.Append("<hr />")
        'Dim oListItemCollection As ListItemCollection = ExecuteQuery("SELECT GroupName, ProductGroup, * FROM UP_UserPermissionGroups upg LEFT OUTER JOIN UP_ProductPermissionGroups ppg ON upg.[id] = ppg.DefaultUserGroup WHERE upg.CustomerKey = " & Session("CustomerKey") & " ORDER BY GroupName", "GroupName", "ProductGroup")
        Dim oListItemCollection As ListItemCollection = ExecuteQuery("SELECT GroupName, ProductGroup FROM UP_UserPermissionGroups upg LEFT OUTER JOIN UP_ProductPermissionGroups ppg ON upg.[id] = ppg.DefaultUserGroup WHERE upg.CustomerKey = " & Session("CustomerKey") & " ORDER BY GroupName", "GroupName", "ProductGroup")
        sbText.Append(Bold("Defined user groups (" & oListItemCollection.Count & ") are:"))
        Call NewLine(sbText)
        Call NewLine(sbText)
        For Each liGroupName As ListItem In oListItemCollection
            sbText.Append(liGroupName.Text)
            If liGroupName.Value <> String.Empty Then
                sbText.Append(" <i>- assigned product group is</i> " & Bold(liGroupName.Value))
            Else
                sbText.Append(" <i>- no product group assigned</i>")
            End If
            Call NewLine(sbText)
        Next
        Call NewLine(sbText)
        'Dim oListItemCollection5 As ListItemCollection = ExecuteQuery("SELECT [Key] UserKey, FirstName + ' ' + LastName + ' (' + UserId + ')' UserName FROM UserProfile WHERE Type = 'User' AND CustomerKey = " & Session("CustomerKey") & " AND NOT DXepartment IN (SELECT GroupName FROM UP_UserPermissionGroups WHERE CustomerKey = " & Session("CustomerKey") & ") ORDER BY FirstName", "UserKey", "UserName")
        Dim oListItemCollection5 As ListItemCollection = ExecuteQuery("SELECT [Key] UserKey, FirstName + ' ' + LastName + ' (' + UserId + ')' UserName FROM UserProfile WHERE Type = 'User' AND CustomerKey = " & Session("CustomerKey") & " AND NOT ISNULL(UserGroup,0) IN (SELECT [id] FROM UP_UserPermissionGroups WHERE CustomerKey = " & Session("CustomerKey") & ") ORDER BY FirstName", "UserKey", "UserName")
        sbText.Append("<hr />")
        sbText.Append(Bold("Users not in a defined group (" & oListItemCollection5.Count & ")"))
        Call NewLine(sbText)
        For Each liUserName As ListItem In oListItemCollection5
            sbText.Append(liUserName.Value)
            Call NewLine(sbText)
        Next
        Call NewLine(sbText)
        Dim oListItemCollectionA As ListItemCollection = ExecuteQuery("SELECT GroupName, [id] FROM UP_UserPermissionGroups WHERE CustomerKey = " & Session("CustomerKey") & " ORDER BY GroupName", "GroupName", "id")
        For Each liGroupName As ListItem In oListItemCollectionA
            'Dim oListItemCollection2 As ListItemCollection = ExecuteQuery("SELECT [Key] UserKey, FirstName + ' ' + LastName + ' (' + UserId + ')' UserName FROM UserProfile WHERE CustomerKey = " & Session("CustomerKey") & " AND DXepartment = '" & liGroupName.Text.Replace("'", "''") & "' ORDER BY FirstName", "UserKey", "UserName")
            Dim oListItemCollection2 As ListItemCollection = ExecuteQuery("SELECT [Key] UserKey, FirstName + ' ' + LastName + ' (' + UserId + ')' UserName FROM UserProfile WHERE CustomerKey = " & Session("CustomerKey") & " AND UserGroup = " & liGroupName.Value & " ORDER BY FirstName", "UserKey", "UserName")
            sbText.Append("<hr />")
            sbText.Append(Bold("Users in group " & liGroupName.Text & " (" & oListItemCollection2.Count & ")"))
            Call NewLine(sbText)
            For Each liUserName As ListItem In oListItemCollection2
                sbText.Append(liUserName.Value)
                Call NewLine(sbText)
            Next
            Call NewLine(sbText)
        Next
        Call NewLine(sbText)
        Call NewLine(sbText)
        sbText.Append("<hr />")
        Dim oListItemCollection6 As ListItemCollection = ExecuteQuery("SELECT [Key] UserKey, FirstName + ' ' + LastName + ' (' + UserId + ')' UserName FROM UserProfile WHERE Type = 'SuperUser' AND CustomerKey = " & Session("CustomerKey") & " ORDER BY FirstName", "UserKey", "UserName")
        sbText.Append(Bold("Super Users (" & oListItemCollection6.Count & ")"))
        Call NewLine(sbText)
        For Each liUserName As ListItem In oListItemCollection6
            sbText.Append(liUserName.Value)
            Call NewLine(sbText)
        Next
        Call NewLine(sbText)
        Call NewLine(sbText)
        sbText.Append("<hr />")
        Dim oListItemCollection3 As ListItemCollection = ExecuteQuery("SELECT [Key] UserKey, FirstName + ' ' + LastName + ' (' + UserId + ')' UserName FROM UserProfile WHERE CustomerKey = " & Session("CustomerKey") & " AND Status = 'Active' AND Type = 'User' AND DeletedFlag = 0 ORDER BY FirstName", "UserName", "UserKey")
        sbText.Append(Bold("Permissioning by user (" & oListItemCollection3.Count & "):"))
        Call NewLine(sbText)
        Dim oListItemCollection7 As New ListItemCollection
        For Each liUser As ListItem In oListItemCollection3
            sbText.Append(liUser.Text)
            Call NewLine(sbText)
            Dim oListItemCollection4 As ListItemCollection = ExecuteQuery("SELECT GroupOrProductKey, 'Product group: ' + ProductGroup 'GroupOrProductName' FROM UP_UserPermissions fup INNER JOIN UP_ProductPermissionGroups fpg ON fup.GroupOrProductKey = fpg.[id] WHERE UserKey = " & liUser.Value & " UNION SELECT GroupOrProductKey, 'Product: ' + ProductCode + ' - ' + ISNULL(ProductDate,'') + ': ' + ProductDescription 'GroupOrProductName' FROM UP_UserPermissions fup INNER JOIN LogisticProduct lp ON fup.GroupOrProductKey = lp.LogisticProductKey WHERE UserKey = " & liUser.Value & " ORDER BY GroupOrProductKey DESC", "GroupOrProductKey", "GroupOrProductName")
            If oListItemCollection4.Count = 0 Then
                oListItemCollection7.Add(New ListItem(liUser.Text, liUser.Text))
                sbText.Append("<i><b><font color='red'>no permissions found for " & liUser.Text & "</font></b></i>")
                Call NewLine(sbText)
            Else
                For Each liPermissionGroup As ListItem In oListItemCollection4
                    sbText.Append(liPermissionGroup.Value)
                    Call NewLine(sbText)
                Next
            End If
        Next
        If oListItemCollection7.Count > 0 Then
            Call NewLine(sbText)
            sbText.Append("<hr />")
            sbText.Append(Bold("Summary of users with no permissions"))
            For Each liUser As ListItem In oListItemCollection7
                sbText.Append(liUser.Value)
                Call NewLine(sbText)
            Next
        End If
        sbText.Append("[end]")
        Call AddHTMLPostamble(sbText)
        Call ExportData(sbText.ToString, "UsersReport")
    End Sub
   
    Protected Sub btnGroupsReport_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim sbText As New StringBuilder
        Dim sGroupName As String = String.Empty
        Dim nGroupKey As Integer = 0
        Call AddHTMLPreamble(sbText, "Product Groups Report")
        sbText.Append(Bold("PRODUCT GROUPS REPORT"))
        Call NewLine(sbText)
        sbText.Append("report generated " & DateTime.Now)
        Call NewLine(sbText)
        Call NewLine(sbText)
        sbText.Append("This report lists all defined product groups and the user group to which each is assigned, if any. Then, for each defined product group, it lists the products in that group and the users to whom the group is assigned.")
        Call NewLine(sbText)
        Call NewLine(sbText)
        Dim oListItemCollection4 As ListItemCollection = ExecuteQuery("SELECT * FROM UP_ProductPermissionGroups ppg LEFT OUTER JOIN UP_UserPermissionGroups upg ON ppg.DefaultUserGroup = upg.[id] WHERE ppg.CustomerKey = " & Session("CustomerKey") & " ORDER BY ProductGroup", "ProductGroup", "GroupName")
        sbText.Append("<hr />")
        sbText.Append(Bold("Defined product groups - association by product group"))
        Call NewLine(sbText)
        Call NewLine(sbText)
        For Each liGroup As ListItem In oListItemCollection4
            sbText.Append(liGroup.Text & " ")
            If liGroup.Value <> String.Empty Then
                sbText.Append("<i>- assigned to user group</i> " & Bold(liGroup.Value))
            Else
                sbText.Append("<i>- not assigned to any user group</i>")
            End If
            Call NewLine(sbText)
        Next

        Dim oListItemCollection As ListItemCollection = ExecuteQuery("SELECT * FROM UP_ProductPermissionGroups WHERE CustomerKey = " & Session("CustomerKey"), "ProductGroup", "id")
        For Each liGroup As ListItem In oListItemCollection
            sGroupName = liGroup.Text
            nGroupKey = liGroup.Value
            sbText.Append("<hr />")
            sbText.Append(Bold("GROUP: " & sGroupName))
            Call NewLine(sbText)
            Call NewLine(sbText)
            Dim oListItemCollection2 As ListItemCollection = ExecuteQuery("SELECT pgm.LogisticProductKey, ProductCode + ' - ' + ISNULL(ProductDate,'') + ': ' + ProductDescription 'Product' FROM UP_ProductGroupsMapping pgm INNER JOIN UP_ProductPermissionGroups pg ON pg.[id] = pgm.ProductGroupKey INNER JOIN LogisticProduct lp ON pgm.LogisticProductKey = lp.LogisticProductKey WHERE pg.CustomerKey = " & Session("CustomerKey") & " AND pg.ProductGroup = '" & sGroupName.Replace("'", "''") & "' ORDER BY ProductCode", "Product", "LogisticProductKey")
            If oListItemCollection2.Count > 0 Then
                sbText.Append(Bold("Products in this group (" & oListItemCollection2.Count & ")"))
                Call NewLine(sbText)
                For Each liProduct As ListItem In oListItemCollection2
                    sbText.Append(liProduct.Text)
                    Call NewLine(sbText)
                Next
            Else
                sbText.Append("<i>There are no products in this group</i>")
                Call NewLine(sbText)
            End If
            Call NewLine(sbText)
            'Dim oListItemCollection3 As ListItemCollection = ExecuteQuery("SELECT [Key] UserKey, FirstName + ' ' + LastName + ' (' + UserId + ')' + ' (user group: ' + ISNULL(DXepartment,'') + ')' UserName  FROM UserProfile up INNER JOIN UP_UserPermissions fup ON up.[key] = fup.UserKey WHERE fup.GroupOrProductKey = " & nGroupKey, "UserName", "UserKey")
            Dim oListItemCollection3 As ListItemCollection = ExecuteQuery("SELECT [Key] UserKey, FirstName + ' ' + LastName + ' (' + UserId + ')' + ' (user group: ' + ISNULL(upg.GroupName,'') + ')' UserName  FROM UserProfile up INNER JOIN UP_UserPermissions fup ON up.[key] = fup.UserKey LEFT OUTER JOIN UP_UserPermissionGroups upg ON upg.[id] = up.UserGroup WHERE fup.GroupOrProductKey = " & nGroupKey, "UserName", "UserKey")
            If oListItemCollection3.Count > 0 Then
                sbText.Append(Bold("This group is assigned to the following " & oListItemCollection3.Count & " user(s)"))
                Call NewLine(sbText)
                Call NewLine(sbText)
                For Each liUser As ListItem In oListItemCollection3
                    sbText.Append(liUser.Text)
                    Call NewLine(sbText)
                Next
            Else
                sbText.Append("<i>This group is not assigned to any users</i>")
                Call NewLine(sbText)
            End If
            Call NewLine(sbText)
        Next
        sbText.Append("[end]")
        Call AddHTMLPostamble(sbText)
        Call ExportData(sbText.ToString, "ProductGroupsReport")
        'sbText.Append("")
    End Sub
   
    Protected Function Bold(ByVal sString As String) As String
        Bold = "<b>" & sString & "</b>"
    End Function
   
    Protected Sub NewLine(ByRef sbText As StringBuilder)
        sbText.Append("<br />" & Environment.NewLine)
    End Sub
   
    Protected Sub AddHTMLPreamble(ByRef sbText As StringBuilder, ByVal sTitle As String)
        sbText.Append("<html><head><title>")
        sbText.Append(sTitle)
        sbText.Append("</title><style>")
        sbText.Append("body { font-family: Verdana; font-size : xx-small }")
        sbText.Append("</style></head><body>")
    End Sub
   
    Protected Sub AddHTMLPostamble(ByRef sbText As StringBuilder)
        sbText.Append("</body></html>")
    End Sub
   
    Private Sub ExportData(ByVal sData As String, ByVal sFilename As String)
        Response.Clear()
        Response.AddHeader("Content-Disposition", "attachment;filename=" & sFilename & ".htm")
        Response.ContentType = "text/html"
   
        Dim eEncoding As Encoding = Encoding.GetEncoding("Windows-1252")
        Dim eUnicode As Encoding = Encoding.Unicode
        Dim byUnicodeBytes As Byte() = eUnicode.GetBytes(sData)
        Dim byEncodedBytes As Byte() = Encoding.Convert(eUnicode, eEncoding, byUnicodeBytes)
        Response.BinaryWrite(byEncodedBytes)

        Response.Flush()
        Response.End()
    End Sub

    Protected Sub btnGo_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call PopulateUserListBox()
    End Sub
   
    Protected Sub PopulateUserListBox()
        Dim sSQL As String
        If rbUsersAll.Checked Then
            sSQL = "SELECT [Key] UserKey, FirstName + ' ' + LastName + ' (' + UserId + ')' UserName FROM UserProfile WHERE CustomerKey = " & Session("CustomerKey") & " AND Status = 'Active' AND DeletedFlag = 0"
            lblLegendUsers.Text = "All users:"
        ElseIf rbUsersWithoutPermissions.Checked Then
            sSQL = "SELECT [Key] UserKey, FirstName + ' ' + LastName + ' (' + UserId + ')' UserName FROM UserProfile WHERE CustomerKey = " & Session("CustomerKey") & " AND Status = 'Active' AND DeletedFlag = 0 AND NOT [Key] IN (SELECT DISTINCT UserKey FROM UserPermissions)"
            lblLegendUsers.Text = "Users without permissions:"
        ElseIf rbUsersInGroup.Checked Then
            If ddlUserGroup.SelectedIndex > 0 Then
                'sSQL = "SELECT [Key] UserKey, FirstName + ' ' + LastName + ' (' + UserId + ')' UserName FROM UserProfile WHERE CustomerKey = " & Session("CustomerKey") & " AND Status = 'Active' AND DeletedFlag = 0 AND DXepartment LIKE '" & ddlUserGroup.SelectedValue & "'"
                sSQL = "SELECT [Key] UserKey, FirstName + ' ' + LastName + ' (' + UserId + ')' UserName FROM UserProfile WHERE CustomerKey = " & Session("CustomerKey") & " AND Status = 'Active' AND DeletedFlag = 0 AND UserGroup = " & ddlUserGroup.SelectedValue & " ORDER BY UserId"
                'lblLegendUsers.Text = "Users in group " & ddlUserGroup.SelectedValue
                lblLegendUsers.Text = "Users in group " & ddlUserGroup.SelectedItem.Text
            Else
                WebMsgBox.Show("Please select a group")
                Exit Sub
            End If
        ElseIf rbUsersMatching.Checked Then
            tbUserSearch.Text = tbUserSearch.Text.Trim
            If tbUserSearch.Text <> String.Empty Then
                sSQL = "SELECT [Key] UserKey, FirstName + ' ' + LastName + ' (' + UserId + ')' UserName FROM UserProfile WHERE CustomerKey = " & Session("CustomerKey") & " AND Status = 'Active' AND DeletedFlag = 0 AND (FirstName LIKE '%" & tbUserSearch.Text & "%' OR LastName LIKE '%" & tbUserSearch.Text & "%' OR UserId LIKE '%" & tbUserSearch.Text & "%')"
                lblLegendUsers.Text = "Users matching '" & tbUserSearch.Text & "'"
            Else
                WebMsgBox.Show("Please enter a full or partial name to match")
                Exit Sub
            End If
        Else
            WebMsgBox.Show("Error in PopulateUserListBox selection")
            Exit Sub
        End If
        Dim oListItemCollection As ListItemCollection = ExecuteQuery(sSQL, "UserName", "UserKey")
        lbUsers.Items.Clear()
        For Each li As ListItem In oListItemCollection
            lbUsers.Items.Add(li)
        Next
    End Sub
   
    Protected Sub btnDefaultGroupAssociation_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call HideAllPanels()
        Call InitUserGroupsListBox()
        Call InitProductGroupsDropdown()
        lblNoProductGroupsDefinedWarning.Visible = False
        pnlDefaultGroupAssociation.Visible = True
    End Sub
   
    Protected Sub InitUserGroupsListBox()
        'Dim oListItemCollection As ListItemCollection = ExecuteQuery("SELECT DISTINCT DXepartment FROM UserProfile WHERE DXepartment <> '' AND CustomerKey = " & Session("CustomerKey"), "DXepartment", "DXepartment")
        Dim oListItemCollection As ListItemCollection = ExecuteQuery("SELECT GroupName, [id] FROM UP_UserPermissionGroups WHERE CustomerKey = " & Session("CustomerKey"), "GroupName", "id")
        lbUserGroups.Items.Clear()
        lbUserGroups.Items.Add(New ListItem("- please select -", 0))
        For Each li As ListItem In oListItemCollection
            lbUserGroups.Items.Add(li)
        Next
    End Sub
   
    Protected Sub InitProductGroupsDropdown()
        Dim sSQL As String = "SELECT id, ProductGroup FROM UP_ProductPermissionGroups WHERE CustomerKey = " & Session("CustomerKey") & " AND NOT id IN (SELECT GroupOrProductKey FROM UP_UserPermissions WHERE UserKey = " & pnUserKey & ") ORDER BY ProductGroup"
        Dim oListItemCollection As ListItemCollection = ExecuteQuery(sSQL, "ProductGroup", "id")
        ddlProductGroups.Items.Clear()
        ddlProductGroups.Items.Add(New ListItem("- please select -", 0))
        For Each li As ListItem In oListItemCollection
            ddlProductGroups.Items.Add(li)
        Next
    End Sub
   
    Protected Sub lbUserGroups_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim lb As ListBox = sender
        If lb.Items(0).Text.Contains("please select") Then
            lb.Items.RemoveAt(0)
        End If
        If lb.SelectedIndex >= 0 Then
            lblUserGroup.Text = lbUserGroups.SelectedItem.Text
            If ddlProductGroups.Items.Count > 1 Then
                lblNoProductGroupsDefinedWarning.Visible = False
                tdProductGroups.Visible = True
                ddlProductGroups.SelectedIndex = 0
                Dim sProductGroupAssociation As String = GetAssociatedProductGroup(lbUserGroups.SelectedValue)
                If sProductGroupAssociation <> String.Empty Then
                    For i As Integer = 1 To ddlProductGroups.Items.Count - 1
                        If ddlProductGroups.Items(i).Text = sProductGroupAssociation Then
                            ddlProductGroups.SelectedIndex = i
                            Exit For
                        End If
                    Next
                End If
            Else
                lblNoProductGroupsDefinedWarning.Visible = True
                tdProductGroups.Visible = False
            End If
        End If
    End Sub
   
    Protected Sub ddlProductGroups_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call SaveAssociation()
    End Sub

    Protected Sub RemoveAssociation(ByVal sDefaultUserGroup)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand
        Dim sSQL As String = "UPDATE UP_ProductPermissionGroups SET DefaultUserGroup = NULL WHERE CustomerKey = " & Session("CustomerKey") & " AND DefaultUserGroup = " & sDefaultUserGroup
        Try
            oConn.Open()
            oCmd = New SqlCommand(sSQL, oConn)
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            WebMsgBox.Show("Error in RemoveAssociation: " & ex.Message)
        Finally
            oConn.Close()
        End Try
        
    End Sub
    
    Protected Sub SaveAssociation()
        Call RemoveAssociation(lbUserGroups.SelectedValue)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand
        Dim sSQL As String = "UPDATE UP_ProductPermissionGroups SET DefaultUserGroup = " & lbUserGroups.SelectedValue & ", LastModifiedDateTime = GETDATE(), LastUpdateBy = " & Session("UserKey") & " WHERE CustomerKey = " & Session("CustomerKey") & " AND id = " & ddlProductGroups.SelectedValue
        Try
            oConn.Open()
            oCmd = New SqlCommand(sSQL, oConn)
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            WebMsgBox.Show("Error in SaveAssociation: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
   
    Protected Sub InitPermissions()
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand
        Dim oListItemCollection1 As ListItemCollection = ExecuteQuery("SELECT ProductGroup, GroupName 'DefaultUserGroup' FROM UP_ProductPermissionGroups ppg LEFT OUTER JOIN UP_UserPermissionGroups upg ON ppg.DefaultUserGroup = upg.[id]  WHERE ppg.CustomerKey = " & Session("CustomerKey"), "ProductGroup", "DefaultUserGroup")
        For Each li As ListItem In oListItemCollection1
            If li.Value <> String.Empty Then
                Dim sProductGroup As String = li.Text
                Dim nProductGroupKey As Integer = GetProductGroupKeyFromName(sProductGroup)
                Dim sDefaultUserGroup As String = li.Value
                Dim nDefaultUserGroup As Integer = GetUserGroupKeyFromName(sDefaultUserGroup)
                'Dim oListItemCollection2 As ListItemCollection = ExecuteQuery("SELECT * FROM UserProfile WHERE CustomerKey = " & Session("CustomerKey") & " AND DXepartment LIKE '" & sDefaultUserGroup & "'", "key", "UserId")
                Dim oListItemCollection2 As ListItemCollection = ExecuteQuery("SELECT * FROM UserProfile WHERE CustomerKey = " & Session("CustomerKey") & " AND UserGroup = " & nDefaultUserGroup, "key", "UserId")
                Dim sSQL As String
                For Each li2 As ListItem In oListItemCollection2
                    sSQL = "INSERT INTO UP_UserPermissions (UserKey, GroupOrProductKey, LastModifiedDateTime, LastUpdateBy) VALUES (" & li2.Text & ", " & nProductGroupKey & ", GETDATE(), " & Session("UserKey") & ")"
                    Try
                        oConn.Open()
                        oCmd = New SqlCommand(sSQL, oConn)
                        oCmd.ExecuteNonQuery()
                    Catch ex As Exception
                        WebMsgBox.Show("Error in InitPermissions: " & ex.Message)
                    Finally
                        oConn.Close()
                    End Try
                Next
            End If
        Next
    End Sub
   
    Protected Sub ClearAllPermissions()
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand
        Dim sSQL As String = "DELETE FROM UP_UserPermissions WHERE UserKey IN (SELECT [key] FROM UserProfile WHERE CustomerKey = " & Session("CustomerKey") & ")"
        Try
            oConn.Open()
            oCmd = New SqlCommand(sSQL, oConn)
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            WebMsgBox.Show("Error in ClearAllPermissions: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
       
    Protected Function GetAssociatedProductGroup(ByVal sDefaultUserGroup As String) As String
        GetAssociatedProductGroup = String.Empty
        Dim sSQL As String = "SELECT ProductGroup FROM UP_ProductPermissionGroups WHERE DefaultUserGroup = " & sDefaultUserGroup & " AND CustomerKey = " & Session("CustomerKey")
        Dim oDataReader As SqlDataReader = Nothing
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            If oDataReader.HasRows Then
                oDataReader.Read()
                GetAssociatedProductGroup = oDataReader("ProductGroup")
            End If
        Catch ex As Exception
            WebMsgBox.Show("Error in GetAssociatedProductGroup: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Function
   
    Protected Sub ShowOnCreateDropdown()
        ddlOnCreateUserGroups.Visible = True
    End Sub
   
    Protected Sub HideOnCreateDropdown()
        ddlOnCreateUserGroups.Visible = False
    End Sub
   
    Protected Sub rbOnCreateNoAction_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call HideOnCreateDropdown()
    End Sub
   
    Protected Sub rbOnCreateAddGroupAllUsers_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call HideOnCreateDropdown()
    End Sub

    Protected Sub rbOnCreateAddGroupUserGroup_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call InitOnCreateUserGroupsListBox()
        Call ShowOnCreateDropdown()
    End Sub

    Protected Sub InitOnCreateUserGroupsListBox()    ' SWAP FOR LIST in db
        'Dim oListItemCollection As ListItemCollection = ExecuteQuery("SELECT DISTINCT DXepartment FROM UserProfile WHERE CustomerKey = " & Session("CustomerKey") & " ORDER BY DXepartment", "DXepartment", "DXepartment")
        Dim oListItemCollection As ListItemCollection = ExecuteQuery("SELECT GroupName, [id] FROM UP_UserPermissionGroups WHERE CustomerKey = " & Session("CustomerKey") & " ORDER BY GroupName", "GroupName", "id")
        ddlOnCreateUserGroups.Items.Clear()
        ddlOnCreateUserGroups.Items.Add(New ListItem("- please select -", 0))
        For Each li As ListItem In oListItemCollection
            If li.Text <> String.Empty Then
                ddlOnCreateUserGroups.Items.Add(li)
            End If
        Next
    End Sub
   
    Protected Sub lnkbtnInitPermissions_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ClearAllPermissions()
        Call InitPermissions()
    End Sub
   
    Protected Sub btnCopyProducts_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If ddlCopyProducts.SelectedIndex > 0 Then
            Call CopyProducts()
        End If
    End Sub
   
    Protected Sub CopyProducts()
        Dim sSQL As String = "SELECT LogisticProductKey FROM UP_ProductGroupsMapping WHERE ProductGroupKey = " & ddlCopyProducts.SelectedValue
        Dim oListItemCollection As ListItemCollection = ExecuteQuery(sSQL, "LogisticProductKey", "LogisticProductKey")
        For Each li As ListItem In oListItemCollection
            If Not IsProductInGroup(li.Text, lbProductGroups.SelectedValue) Then
                AddProductToGroup(li.Text)
            End If
        Next
        Call PopulateAvailableProductsListBox(lbProductGroups.SelectedItem.Text)
        Call PopulateProductsInGroupListBox(lbProductGroups.SelectedItem.Text)
    End Sub
   
    Protected Sub InitCopyProductsDropdown()
        If lbProductGroups.Items.Count > 1 Then
            lblLegendCopyFromGroup.Visible = True
            ddlCopyProducts.Visible = True
            btnCopyProducts.Visible = True

            Dim oListItemCollection As ListItemCollection = ExecuteQuery("SELECT * FROM UP_ProductPermissionGroups WHERE CustomerKey = " & Session("CustomerKey"), "ProductGroup", "id")
            ddlCopyProducts.Items.Clear()
            ddlCopyProducts.Items.Add(New ListItem("- please select -", 0))
            For Each li As ListItem In oListItemCollection
                If li.Text <> lbProductGroups.SelectedItem.Text Then
                    ddlCopyProducts.Items.Add(li)
                End If
            Next
        Else
            lblLegendCopyFromGroup.Visible = False
            ddlCopyProducts.Visible = False
            btnCopyProducts.Visible = False
        End If
    End Sub
   
    Protected Sub btnRenameProductGroup_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call HideAllPanels()
        pnlRenameProductGroup.Visible = True
        lblRenameCurrentName.Text = lbProductGroups.SelectedItem.Text
        tbRenameProductGroupNewName.Text = String.Empty
        tbRenameProductGroupNewName.Focus()
    End Sub
   
    Protected Sub btnDoRename_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        tbRenameProductGroupNewName.Text = tbRenameProductGroupNewName.Text.Trim
        If tbRenameProductGroupNewName.Text <> String.Empty Then
            Call RenameProductGroup()
        Else
            WebMsgBox.Show("Please specify a name.")
        End If
    End Sub
   
    Protected Sub RenameProductGroup()
        If GetProductGroupKeyFromName(tbRenameProductGroupNewName.Text) > 0 Then
            WebMsgBox.Show("This name is already in use! Please choose another name.")
        Else
            Dim oConn As New SqlConnection(gsConn)
            Dim oCmd As SqlCommand
            Dim sSQL As String = "UPDATE UP_ProductPermissionGroups SET ProductGroup = '" & tbRenameProductGroupNewName.Text & "' WHERE [id] = " & lbProductGroups.SelectedValue
            Try
                oConn.Open()
                oCmd = New SqlCommand(sSQL, oConn)
                oCmd.ExecuteNonQuery()
            Catch ex As Exception
                WebMsgBox.Show("Error in RenameGroup: " & ex.Message)
            Finally
                oConn.Close()
            End Try
           
            Call HideAllPanels()
            Call InitProductGroupListBox()
            Call InitUserPermissioningListBoxes()
            pnlProductGroups.Visible = True
        End If
    End Sub
   
    Protected Sub btnCancelRenameProductGroup_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call HideAllPanels()
        pnlProductGroups.Visible = True
    End Sub
   
    Protected Sub btnUserGroups_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call HideAllPanels()
        pnlUserGroups.Visible = True
        Call InitDefinedUserGroupsListbox()
    End Sub
   
    Protected Sub InitDefinedUserGroupsListbox()
        Dim oListItemCollection As ListItemCollection = ExecuteQuery("SELECT * FROM UP_UserPermissionGroups WHERE CustomerKey = " & Session("CustomerKey") & " ORDER BY GroupName", "GroupName", "id")
        lbDefinedUserGroups.Items.Clear()
        lbDefinedUserGroups.Items.Add(New ListItem("- please select -", 0))
        For Each li As ListItem In oListItemCollection
            lbDefinedUserGroups.Items.Add(li)
        Next
    End Sub
   
    Protected Sub btnNewUserGroup_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call HideAllPanels()
        pnlNewUserGroup.Visible = True
        tbNewUserGroupName.Focus()
        SetEnableRenameAndRemoveUserGroup(False)
    End Sub
   
    Protected Sub btnRenameUserGroup_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call HideAllPanels()
        pnlRenameUserGroup.Visible = True
        tbRenameUserGroupNewName.Focus()
    End Sub
   
    Protected Sub btnRemoveUserGroup_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        SetEnableRenameAndRemoveUserGroup(False)
        Call RemoveUserGroup()
        Call InitDefinedUserGroupsListbox()
    End Sub
    
    Protected Sub RemoveUserGroup()
        Call RemoveUserGroupFromAssociatedGroups(lbDefinedUserGroups.SelectedValue)
        Call RemoveUserGroupFromUsers(lbDefinedUserGroups.SelectedValue)
        Call RemoveUserGroupFromList(lbDefinedUserGroups.SelectedValue)
    End Sub

    Protected Sub RemoveUserGroupFromAssociatedGroups(ByVal nUserGroupKey As Integer)
        Call ExecuteNonQuery("UPDATE UP_ProductPermissionGroups SET DefaultUserGroup = NULL WHERE DefaultUserGroup = " & nUserGroupKey & " AND CustomerKey = " & Session("CustomerKey"))
    End Sub
    
    Protected Sub RemoveUserGroupFromUsers(ByVal nUserGroupKey As Integer)
        Call ExecuteNonQuery("UPDATE UserProfile SET UserGroup = NULL WHERE UserGroup = " & nUserGroupKey & " AND CustomerKey = " & Session("CustomerKey"))
    End Sub
    
    Protected Sub RemoveUserGroupFromList(ByVal nUserGroupKey As Integer)
        Call ExecuteNonQuery("DELETE FROM UP_UserPermissionGroups WHERE [id] = " & nUserGroupKey)
    End Sub

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
            WebMsgBox.Show("Error in ExecuteNonQuery executing " & sQuery & " : " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Function
   
    Protected Sub btnCreateNewUserGroup_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call CreateNewUserGroup()
    End Sub

    Protected Sub CreateNewUserGroup()
        tbNewUserGroupName.Text = tbNewUserGroupName.Text.Trim
        If tbNewUserGroupName.Text = String.Empty Then
            WebMsgBox.Show("Please enter a name.")
        Else
            If UserGroupExists(tbNewUserGroupName.Text) Then
                WebMsgBox.Show("This name is already in use! Please choose another name.")
            Else
                Dim oConn As New SqlConnection(gsConn)
                Dim oCmd As SqlCommand
                Dim sSQL As String = "INSERT INTO UP_UserPermissionGroups (GroupName, CustomerKey, LastModifiedDateTime, LastUpdateBy) VALUES ('"
                sSQL += tbNewUserGroupName.Text.Replace("'", "''") & "', " & Session("CustomerKey") & ", GETDATE(), " & Session("UserKey") & ")"
                Try
                    oConn.Open()
                    oCmd = New SqlCommand(sSQL, oConn)
                    oCmd.ExecuteNonQuery()
                Catch ex As Exception
                    WebMsgBox.Show("Error in CreateNewUserGroup: " & ex.Message)
                Finally
                    oConn.Close()
                End Try
            End If
        End If
    End Sub
   
    Protected Function UserGroupExists(ByVal sUserGroupName As String) As Boolean
        Dim oListItemCollection As ListItemCollection = ExecuteQuery("SELECT * FROM UP_UserPermissionGroups WHERE CustomerKey = " & Session("CustomerKey") & " AND GroupName = " & sUserGroupName.Replace("'", "''"), "GroupName", "GroupName")
        If oListItemCollection.Count = 0 Then
            UserGroupExists = False
        Else
            UserGroupExists = True
        End If
    End Function
    
    Protected Sub btnCancelNewUserGroup_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call HideAllPanels()
        pnlUserGroups.Visible = True
    End Sub

    Protected Sub btnDoRenameUserGroup_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call RenameUserGroup()
        Call InitDefinedUserGroupsListbox()
        lbDefinedUserGroups.SelectedIndex = -1
        Call HideAllPanels()
        pnlUserGroups.Visible = True
    End Sub
   
    Protected Sub RenameUserGroup() ' no longer needed once change made to use id of UP_UserPermissionGroups
        tbRenameUserGroupNewName.Text = tbRenameUserGroupNewName.Text.Trim
        If UserGroupExists(tbRenameUserGroupNewName.Text) Then
            WebMsgBox.Show("This name is already in use! Please choose another name.")
        Else
            Dim oConn As New SqlConnection(gsConn)
            Dim oCmd As SqlCommand
            Dim sSQL As String = "UPDATE UP_UserPermissionGroups SET "
            sSQL += "GroupName = '" & tbRenameUserGroupNewName.Text.Replace("'", "''") & "', LastModifiedDateTime = GETDATE(), LastUpdateBy = " & Session("UserKey")
            sSQL += " WHERE CustomerKey = " & Session("CustomerKey") & " AND GroupName = '" & lbDefinedUserGroups.SelectedItem.Text & "'"
            Try
                oConn.Open()
                oCmd = New SqlCommand(sSQL, oConn)
                oCmd.ExecuteNonQuery()
            Catch ex As Exception
                WebMsgBox.Show("Error in RenameUserGroup: " & ex.Message)
            Finally
                oConn.Close()
            End Try
        End If
    End Sub
   
    Protected Sub btnCancelRenameUserGroup_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call HideAllPanels()
        pnlUserGroups.Visible = True
    End Sub
   
    Protected Sub btnMaxGrabs_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call HideAllPanels()
        Call InitMaxGrabsPanel()
        pnlMaxGrabs.Visible = True
    End Sub
   
    Protected Sub InitMaxGrabsPanel()
        Call InitMaxGrabUserGroups()
        Call SetMaxGrabProductsVisibility(False)
        Call InitMaxGrabProductGroupDropdown()
        Call SetCopyMaxGrabsVisibility(False)
        btnCopyMaxGrabs.Enabled = False
    End Sub
   
    Protected Sub InitMaxGrabUserGroups()
        Dim oListItemCollection As ListItemCollection = ExecuteQuery("SELECT GroupName, [id] FROM UP_UserPermissionGroups WHERE CustomerKey = " & Session("CustomerKey"), "GroupName", "id")
        lbMaxGrabUserGroups.Items.Clear()
        lbMaxGrabUserGroups.Items.Add(New ListItem("- please select -", 0))
        For Each li As ListItem In oListItemCollection
            lbMaxGrabUserGroups.Items.Add(li)
        Next
    End Sub
   
    Protected Sub InitMaxGrabProductGroupDropdown()
        Dim oListItemCollection As ListItemCollection = ExecuteQuery("SELECT ProductGroup, [id] FROM UP_ProductPermissionGroups WHERE CustomerKey = " & Session("CustomerKey"), "ProductGroup", "id")
        ddlMaxGrabProductGroup.Items.Clear()
        ddlMaxGrabProductGroup.Items.Add(New ListItem("- please select -", 0))
        For Each li As ListItem In oListItemCollection
            ddlMaxGrabProductGroup.Items.Add(li)
        Next
    End Sub

    Protected Sub lbMaxGrabUserGroups_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If lbMaxGrabUserGroups.Items(0).Value = 0 Then
            lbMaxGrabUserGroups.Items.RemoveAt(0)
            Exit Sub
        End If
        Call SetMaxGrabProductsVisibility(True)
        Call InitMaxGrabProducts()
        Call InitCopyMaxGrabsDropdown()
        Call SetCopyMaxGrabsVisibility(True)
        lblMaxGrabUserGroup.Text = lbMaxGrabUserGroups.SelectedItem.Text
        lblCopyMaxGrabsSource.Text = lbMaxGrabUserGroups.SelectedItem.Text
    End Sub
   
    Protected Sub InitCopyMaxGrabsDropdown()
        Dim oListItemCollection As ListItemCollection = ExecuteQuery("SELECT GroupName, [id] FROM UP_UserPermissionGroups WHERE CustomerKey = " & Session("CustomerKey"), "GroupName", "id")
        ddlCopyMaxGrabs.Items.Clear()
        ddlCopyMaxGrabs.Items.Add(New ListItem("- please select -", 0))
        For Each li As ListItem In oListItemCollection
            If li.Value <> lbMaxGrabUserGroups.SelectedValue Then
                ddlCopyMaxGrabs.Items.Add(li)
            End If
        Next
    End Sub
    
    Protected Sub InitMaxGrabProducts()
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataTable As New DataTable
        Dim sSQL As String = "SELECT LogisticProductKey, ProductCode, ProductDate , ProductDescription FROM LogisticProduct WHERE CustomerKey = " & Session("CustomerKey")
        If rblMaxGrabProductsInGroup.Checked Then
            sSQL += " AND LogisticProductKey IN (SELECT LogisticProductKey FROM UP_ProductGroupsMapping WHERE ProductGroupKey = " & ddlMaxGrabProductGroup.SelectedValue & ")"
        End If

        Dim oAdapter As New SqlDataAdapter(sSQL, oConn)
        Try
            oAdapter.Fill(oDataTable)
            gvMaxGrabProducts.DataSource = oDataTable
            gvMaxGrabProducts.DataBind()
        Catch ex As Exception
            WebMsgBox.Show("InitMaxGrabProducts: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
   
    Protected Sub SetMaxGrabProductsVisibility(ByVal bVisibility As Boolean)
        'trMaxGrabProducts01.Visible = bVisibility
        trMaxGrabProducts02.Visible = bVisibility
        trMaxGrabProducts03.Visible = bVisibility
    End Sub
   
    Protected Sub rblMaxGrabAllProducts_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        ddlMaxGrabProductGroup.SelectedIndex = 0
        ddlMaxGrabProductGroup.Enabled = False
        If lbMaxGrabUserGroups.SelectedIndex >= 0 AndAlso lbMaxGrabUserGroups.SelectedValue > 0 Then
            Call InitMaxGrabProducts()
        End If
    End Sub

    Protected Sub rblMaxGrabProductsInGroup_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        ddlMaxGrabProductGroup.SelectedIndex = 0
        ddlMaxGrabProductGroup.Enabled = True
        If lbMaxGrabUserGroups.SelectedIndex >= 0 AndAlso lbMaxGrabUserGroups.SelectedValue > 0 Then
            Call InitMaxGrabProducts()
        End If
    End Sub
   
    Protected Sub ddlMaxGrabProductGroup_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        rblMaxGrabProductsInGroup.Checked = True
        If lbMaxGrabUserGroups.SelectedIndex >= 0 AndAlso lbMaxGrabUserGroups.SelectedValue > 0 Then
            Call InitMaxGrabProducts()
        End If
    End Sub
   
    Protected Sub btnSaveMaxGrab_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim gv As GridView = gvMaxGrabProducts
        For Each gvr As GridViewRow In gv.Rows
            If gvr.RowType = DataControlRowType.DataRow Then
                Dim hidLogisticProductKey As HiddenField
                Dim tbMaxGrab As TextBox
                hidLogisticProductKey = gvr.Cells(0).FindControl("hidLogisticProductKey")
                tbMaxGrab = gvr.Cells(0).FindControl("tbMaxGrab")
                Call SetMaxGrabValue(hidLogisticProductKey.Value, lbMaxGrabUserGroups.SelectedValue, tbMaxGrab.Text)
            End If
        Next
        If gv.Rows.Count > 0 Then
            If Not cbDontKeepRemindingMe.Checked Then
                WebMsgBox.Show("After you have made all required changes to max grab values, click the [apply all max grabs!] button to apply the changes to user accounts.\n\nIf you have several changes to make, you can stop this reminder coming up by checking the [don't keep reminding me] check box.")
            End If
        End If
    End Sub
   
    Protected Sub gvMaxGrabProducts_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs)
        Dim gvr As GridViewRow = e.Row
        If gvr.RowType = DataControlRowType.DataRow Then
            Dim hidLogisticProductKey As HiddenField
            Dim tbMaxGrab As TextBox
            hidLogisticProductKey = gvr.Cells(0).FindControl("hidLogisticProductKey")
            tbMaxGrab = gvr.Cells(0).FindControl("tbMaxGrab")
            tbMaxGrab.Text = sGetMaxGrabValue(CInt(hidLogisticProductKey.Value), lbMaxGrabUserGroups.SelectedValue)
        End If
    End Sub
    
    Protected Function sGetMaxGrabValue(ByVal nLogisticProductKey As Integer, ByVal nUserGroupKey As Integer) As String
        sGetMaxGrabValue = ""
        Dim oDataReader As SqlDataReader = Nothing
        Dim oConn As New SqlConnection(gsConn)
        Dim sSQL As String = "SELECT MaxGrab FROM UP_UserPermissionGroupsMaxGrabMatrix WHERE ProductKey = " & nLogisticProductKey & " AND UserGroupKey = " & nUserGroupKey
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            If oDataReader.HasRows Then
                oDataReader.Read()
                sGetMaxGrabValue = oDataReader("MaxGrab")
            End If
        Catch ex As Exception
            WebMsgBox.Show("Error in nGetMaxGrabValue: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Function
    
    Protected Sub SetMaxGrabValue(ByVal nLogisticProductKey As Integer, ByVal nUserGroupKey As Integer, ByVal sMaxGrabValue As String)
        If sMaxGrabValue = String.Empty OrElse sMaxGrabValue = 0 Then
            Call DeleteMaxGrabValue(nLogisticProductKey, nUserGroupKey)
        ElseIf sGetMaxGrabValue(nLogisticProductKey, nUserGroupKey) <> String.Empty Then
            Call UpdateMaxGrabValue(nLogisticProductKey, nUserGroupKey, sMaxGrabValue)
        Else
            Call InsertMaxGrabValue(nLogisticProductKey, nUserGroupKey, sMaxGrabValue)
        End If
    End Sub

    Protected Sub DeleteMaxGrabValue(ByVal nLogisticProductKey As Integer, ByVal nUserGroupKey As Integer)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand
        Dim sSQL As String = "DELETE FROM UP_UserPermissionGroupsMaxGrabMatrix WHERE ProductKey = " & nLogisticProductKey & " AND UserGroupKey = " & nUserGroupKey
        Try
            oConn.Open()
            oCmd = New SqlCommand(sSQL, oConn)
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            WebMsgBox.Show("Error in DeleteMaxGrabValue: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub InsertMaxGrabValue(ByVal nLogisticProductKey As Integer, ByVal nUserGroupKey As Integer, ByVal nMaxGrabValue As Integer)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand
        Dim sSQL As String = "INSERT INTO UP_UserPermissionGroupsMaxGrabMatrix (UserGroupKey, ProductKey, MaxGrab, LastModifiedDateTime, LastUpdateBy) VALUES ("
        sSQL += nUserGroupKey & ", " & nLogisticProductKey & ", " & nMaxGrabValue & ", GETDATE(), " & Session("UserKey") & ")"
        Try
            oConn.Open()
            oCmd = New SqlCommand(sSQL, oConn)
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            WebMsgBox.Show("Error in InsertMaxGrabValue: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub UpdateMaxGrabValue(ByVal nLogisticProductKey As Integer, ByVal nUserGroupKey As Integer, ByVal nMaxGrabValue As Integer)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand
        Dim sSQL As String = "UPDATE UP_UserPermissionGroupsMaxGrabMatrix SET MaxGrab = " & nMaxGrabValue & ", LastModifiedDateTime = GETDATE(), LastUpdateBy = " & Session("UserKey") & " WHERE UserGroupKey = " & nUserGroupKey & " AND ProductKey = " & nLogisticProductKey
        Try
            oConn.Open()
            oCmd = New SqlCommand(sSQL, oConn)
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            WebMsgBox.Show("Error in UpdateMaxGrabValue: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub

    Protected Sub btnApplyAllMaxGrabs_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ApplyAllMaxGrabs()
        WebMsgBox.Show("Max Grabs have been applied")
    End Sub
    
    Protected Sub ApplyAllMaxGrabs()
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_UserPermissions_ApplyMaxGrabs2", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int)
        paramCustomerKey.Value = Session("CustomerKey")
        oCmd.Parameters.Add(paramCustomerKey)
        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As SqlException
            WebMsgBox.Show("Error in ApplyAllMaxGrabs: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub

    Protected Sub btnMaxGrabsReport_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim sbText As New StringBuilder
        Dim sGroupName As String = String.Empty
        Dim nGroupKey As Integer = 0
        Call AddHTMLPreamble(sbText, "User Groups Max Grabs Report")
        sbText.Append(Bold("USER GROUPS MAX GRABS REPORT"))
        Call NewLine(sbText)
        sbText.Append("report generated " & DateTime.Now)
        Call NewLine(sbText)
        Call NewLine(sbText)
        sbText.Append("This report shows the default max grabs defined for each user group, followed by a list of all available products, with archived products segregated.")
        Call NewLine(sbText)
        Call NewLine(sbText)
        Dim oListItemCollection As ListItemCollection = ExecuteQuery("SELECT * FROM UP_UserPermissionGroups WHERE CustomerKey = " & Session("CustomerKey") & " ORDER BY GroupName", "GroupName", "id")
        For Each liUserGroup As ListItem In oListItemCollection
            sbText.Append("<hr />")
            sbText.Append(Bold(liUserGroup.Text))
            Call NewLine(sbText)

            Dim nUserGroupKey As Integer = liUserGroup.Value
            Dim oListItemCollection2 As ListItemCollection = ExecuteQuery("SELECT lp.LogisticProductKey, ProductCode + ' - ' + ISNULL(ProductDate,'') + ': ' + ProductDescription 'Product', MaxGrab FROM UP_UserPermissionGroupsMaxGrabMatrix upgmgm INNER JOIN LogisticProduct lp ON upgmgm.ProductKey = lp.LogisticProductKey WHERE UserGroupKey = " & nUserGroupKey & " ORDER BY ProductCode", "Product", "MaxGrab")
            If oListItemCollection2.Count > 0 Then
                For Each liProduct As ListItem In oListItemCollection2
                    sbText.Append(liProduct.Text)
                    Call NewLine(sbText)
                    sbText.Append("Max Grab: " & Bold(liProduct.Value))
                    Call NewLine(sbText)
                    Call NewLine(sbText)
                Next
            Else
                sbText.Append("<i>No product max grabs defined for this user group</i>")
                Call NewLine(sbText)
            End If
            Call NewLine(sbText)
        Next
        Call NewLine(sbText)
        sbText.Append("<hr />")
        sbText.Append(Bold("LIST OF ALL AVAILABLE PRODUCTS"))
        Call NewLine(sbText)
        Call NewLine(sbText)
        Dim oListItemCollection3 As ListItemCollection = ExecuteQueryFlagArchivedProducts("SELECT LogisticProductKey, ProductCode + ' - ' + ISNULL(ProductDate,'') + ': ' + ProductDescription 'Product', ArchiveFlag FROM LogisticProduct WHERE CustomerKey = " & Session("CustomerKey") & " ORDER BY ArchiveFlag, ProductCode", "Product", "LogisticProductKey")
        If oListItemCollection3.Count > 0 Then
            For Each liProduct As ListItem In oListItemCollection3
                sbText.Append(liProduct.Text)
                Call NewLine(sbText)
            Next
        Else
            sbText.Append("<i>No products defined</i>")
        End If
        Call NewLine(sbText)
        sbText.Append("[end]")
        Call AddHTMLPostamble(sbText)
        Call ExportData(sbText.ToString, "MaxGrabsReport")
    End Sub
    
    Protected Sub btnCopyMaxGrabs_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim oListItemCollection As ListItemCollection = ExecuteQuery("SELECT ProductKey, MaxGrab FROM UP_UserPermissionGroupsMaxGrabMatrix WHERE UserGroupKey = " & lbMaxGrabUserGroups.SelectedValue, "ProductKey", "MaxGrab")
        For Each liProductMaxGrab As ListItem In oListItemCollection
            Call SetMaxGrabValue(liProductMaxGrab.Text, ddlCopyMaxGrabs.SelectedValue, liProductMaxGrab.Value)
        Next
        WebMsgBox.Show("Product Max Grab values copied from user group " & lbMaxGrabUserGroups.SelectedItem.Text & " to user group " & ddlCopyMaxGrabs.SelectedItem.Text)
    End Sub
    
    Protected Sub ddlCopyMaxGrabs_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If ddlCopyMaxGrabs.Items(0).Text.Contains("please select") Then
            ddlCopyMaxGrabs.Items.RemoveAt(0)
        End If
        btnCopyMaxGrabs.Enabled = True
    End Sub
    
    Protected Sub SetCopyMaxGrabsVisibility(ByVal bVisible As Boolean)
        tdCopyMaxGrabs01.Visible = bVisible
        tdCopyMaxGrabs02.Visible = bVisible
    End Sub
    
    Protected Sub lbDefinedUserGroups_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If lbDefinedUserGroups.Items(0).Text.Contains("please select") Then
            lbDefinedUserGroups.Items.RemoveAt(0)
        End If
        If lbDefinedUserGroups.SelectedIndex >= 0 Then
            Call SetEnableRenameAndRemoveUserGroup(True)
        End If
    End Sub
    
    '    CREATE PROCEDURE [dbo].[spASPNET_UserPermissions_Apply2]
    '	  @CustomerKey int
    '	, @UserKey int
    '	, @ProductGroupKey int
    'AS
    '	DECLARE @nUserKey int
    '	      , @nProductKey int

    '	-- IF NOT @CustomerKey = 579

    '	CREATE TABLE #temp (UserKey int, ProductKey int)

    '	IF @ProductGroupKey > 0
    '	BEGIN
    '		UPDATE UserProductProfile
    '		SET AbleToPick = 0
    '		WHERE UserKey IN
    '		(SELECT UserKey
    '		FROM UP_UserPermissions
    '		WHERE GroupOrProductKey = @ProductGroupKey)

    '		INSERT INTO #temp (UserKey, ProductKey)
    '		SELECT UserKey, LogisticProductKey
    '		FROM UP_UserPermissions fup
    '		INNER JOIN UP_ProductGroupsMapping fpgm
    '		ON fup.GroupOrProductKey = fpgm.ProductGroupKey
    '		INNER JOIN UP_ProductPermissionGroups fpg
    '		ON fpgm.ProductGroupKey = fpg.[id]
    '	        WHERE GroupOrProductKey = @ProductGroupKey
    '	END
    '	ELSE
    '	IF @UserKey > 0
    '	BEGIN
    '		UPDATE UserProductProfile
    '		SET AbleToPick = 0
    '		WHERE UserKey = @UserKey

    '		INSERT INTO #temp (UserKey, ProductKey)
    '		SELECT UserKey, LogisticProductKey
    '		FROM UP_UserPermissions fup
    '		INNER JOIN UP_ProductGroupsMapping fpgm
    '		ON fup.GroupOrProductKey = fpgm.ProductGroupKey
    '		INNER JOIN UP_ProductPermissionGroups fpg
    '		ON fpgm.ProductGroupKey = fpg.[id]
    '		WHERE UserKey = @UserKey

    '		UNION

    '		SELECT UserKey, GroupOrProductKey 'LogisticProductKey'
    '		FROM UP_UserPermissions fup
    '		INNER JOIN UserProfile up
    '		ON fup.UserKey = up.[Key]
    '		WHERE UserKey = @UserKey
    '		AND (GroupOrProductKey < 1000000)
    '	END
    '	ELSE
    '	BEGIN
    '		UPDATE UserProductProfile
    '		SET AbleToPick = 0
    '		WHERE UserKey IN
    '		(SELECT [key] FROM UserProfile WHERE CustomerKey = @CustomerKey)

    '		INSERT INTO #temp (UserKey, ProductKey)
    '		SELECT UserKey, LogisticProductKey
    '		FROM UP_UserPermissions fup
    '		INNER JOIN UP_ProductGroupsMapping fpgm
    '		ON fup.GroupOrProductKey = fpgm.ProductGroupKey
    '		INNER JOIN UP_ProductPermissionGroups fpg
    '		ON fpgm.ProductGroupKey = fpg.[id]
    '		WHERE CustomerKey = @CustomerKey

    '		UNION

    '		SELECT UserKey, GroupOrProductKey 'LogisticProductKey'
    '		FROM UP_UserPermissions fup
    '		INNER JOIN UserProfile up
    '		ON fup.UserKey = up.[Key]
    '		WHERE CustomerKey = @CustomerKey
    '		AND (GroupOrProductKey < 1000000)
    '	END

    '	DECLARE c CURSOR FOR
    '	SELECT UserKey, ProductKey FROM #temp
    '	OPEN c
    '	FETCH NEXT FROM c INTO @nUserKey, @nProductKey
    '	WHILE (@@FETCH_STATUS) = 0
    '	BEGIN
    '		UPDATE UserProductProfile
    '		SET AbleToPick = 1
    '		WHERE UserKey = @nUserKey
    '		AND ProductKey = @nProductKey

    '		FETCH NEXT FROM c INTO @nUserKey, @nProductKey
    '	END
    '	CLOSE c
    '	DEALLOCATE c

    '	SELECT * FROM #temp
    
</script>
<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>User Permissions</title>
</head>
<body>
    <form id="frmAdministrator" runat="server">
      <main:Header id="ctlHeader" runat="server"></main:Header>
        <table width="100%" cellpadding="0" cellspacing="0">
            <tr class="bar_siteadministrator">
                <td style="width:50%; white-space:nowrap">
                </td>
                <td style="width:50%; white-space:nowrap" align="right">
                </td>
            </tr>
        </table>
            <strong style="color: navy; font-size:x-small; font-family:Verdana">&nbsp;User Permissions<br />
            </strong>
            <table style="width: 100%">
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td style="width: 20%">
                    </td>
                    <td style="width: 29%">
                    </td>
                    <td style="width: 20%">
                    </td>
                    <td align="right" style="width: 29%">
                    </td>
                    <td style="width: 1%">
                    </td>
                </tr>
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td colspan="2" style="white-space:nowrap">
                        <asp:Button ID="btnProductGroups" runat="server" Text="product groups" Width="110px" onclick="btnProductGroups_Click" CausesValidation="False" />
                        <asp:Button ID="btnUserGroups" runat="server" Text="user groups" OnClick="btnUserGroups_Click" Width="110px" CausesValidation="False" />
                        <asp:Button ID="btnDefaultGroupAssociation" runat="server" OnClick="btnDefaultGroupAssociation_Click" Text="default group association" Width="180px" CausesValidation="False" />
                        <asp:Button ID="btnMaxGrabs" runat="server" Text="max grabs" OnClick="btnMaxGrabs_Click" CausesValidation="False" />
                        &nbsp; &nbsp;&nbsp;
                        <asp:Button ID="btnUsers" runat="server" Text="tweak users" Width="150px" onclick="btnUsers_Click" CausesValidation="False" />
                    </td>
                    <td align="right" colspan="2" style="white-space:nowrap"><asp:Label ID="Label32" runat="server" Font-Names="Verdana"
                            Font-Size="XX-Small" Text="Reports:" />
                        <asp:Button ID="btnGroupsReport" runat="server" Text="groups" Width="65px" OnClick="btnGroupsReport_Click" CausesValidation="False" />
                        <asp:Button ID="btnUsersReport" runat="server" Text="users" Width="65px" OnClick="btnUsersReport_Click" CausesValidation="False" />
                        <asp:Button ID="btnMaxGrabsReport" runat="server" OnClick="btnMaxGrabsReport_Click"
                            Text="max grabs" CausesValidation="False" />&nbsp;
                        <asp:Button ID="btnHelp" runat="server" OnClientClick='window.open("help_userpermissions.pdf", "UPHelp","top=10,left=10,width=700,height=450,status=no,toolbar=yes,address=no,menubar=yes,resizable=yes,scrollbars=yes");'
                            Text="help" CausesValidation="False" /></td>
                    <td style="width: 1%">
                    </td>
                </tr>
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td colspan="3" style="white-space:nowrap">
                        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
                    </td>
                    <td style="width: 29%" align="right">
                        </td>
                    <td style="width: 1%">
                    </td>
                </tr>
            </table>
        <asp:Panel ID="pnlProductGroups" runat="server" Width="100%" Visible="False">
            <table style="width: 100%">
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td style="width: 20%">
                    </td>
                    <td style="width: 29%">
                    </td>
                    <td style="width: 20%">
                    </td>
                    <td align="right" style="width: 29%">
                    </td>
                    <td style="width: 1%">
                    </td>
                </tr>
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td style="width: 20%">
                        </td>
                    <td style="width: 29%"><asp:Label ID="Label2" runat="server" Font-Names="Verdana"
                            Font-Size="XX-Small" Text="Product groups:" /></td>
                    <td style="width: 20%">
                    </td>
                    <td style="width: 29%" align="right">
                        </td>
                    <td style="width: 1%">
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="left" valign="top">
                        &nbsp;<asp:Label ID="Label23" runat="server" Font-Names="Verdana"
                            Font-Size="XX-Small" Text="<b>Create, edit, rename & delete product groups here.</b><br /><br />NOTE: Deleting a product group only deletes the group definition, not the products themselves.<br /><br />NOTE: Deleting a product group removes permissions for all products in that group, for users to whom that group is assigned." Width="100%" /></td>
                    <td colspan="2">
                        <asp:ListBox ID="lbProductGroups" runat="server" Rows="10" Width="100%" AutoPostBack="True" OnSelectedIndexChanged="lbProductGroups_SelectedIndexChanged" Font-Names="Verdana" Font-Size="XX-Small"></asp:ListBox>
                    </td>
                    <td align="left" valign="top">
                        <strong style="color: navy; font-size:x-small; font-family:Verdana">&nbsp;&nbsp;&nbsp;&nbsp;<asp:Button
                            ID="btnNewProductGroup" runat="server" Text="new product group" Width="200px" OnClick="btnNewProductGroup_Click" />
                        <br />
                            <br />
                            &nbsp; &nbsp;
                            <asp:Button ID="btnRenameProductGroup" runat="server" Text="rename product group" Width="200px" OnClick="btnRenameProductGroup_Click" Enabled="False" /><br />
                        <br />
                        &nbsp;&nbsp;&nbsp;
                        <asp:Button ID="btnRemoveProductGroup" runat="server" Text="remove product group" Width="200px" OnClientClick='return confirm("Are you sure you want to remove this group?");' OnClick="btnRemoveProductGroup_Click" Enabled="False" />
                        </strong></td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                        &nbsp;</td>
                    <td>
                        &nbsp;</td>
                    <td>
                        <asp:Label ID="Label28" runat="server" Font-Names="Verdana"
                            Font-Size="XX-Small" Text="Available products:" />
                    </td>
                    <td>
                        &nbsp;</td>
                    <td align="left">
                        <asp:Label ID="Label29" runat="server" Font-Names="Verdana"
                            Font-Size="XX-Small" Text="Products in group:" />
                    </td>
                    <td>
                        &nbsp;</td>
                </tr>
                <tr>
                    <td>
                        &nbsp;</td>
                    <td>
                        &nbsp;</td>
                    <td>
                        <asp:ListBox ID="lbAvailableProductsForProductGroup" runat="server" Font-Names="Verdana"
                            Font-Size="XX-Small" Rows="10" SelectionMode="Multiple" Width="100%">
                        </asp:ListBox>
                    </td>
                    <td align="center" valign="middle" style="white-space:nowrap">
                        <asp:LinkButton ID="lnkbtnRemoveProductFromGroup" runat="server" Font-Names="Verdana"
                            Font-Size="XX-Small" OnClick="lnkbtnRemoveProductFromGroup_Click">&lt;&lt;&lt;&lt;&lt;
                        </asp:LinkButton>
                        &nbsp;&nbsp;&nbsp;&nbsp;
                        <asp:LinkButton ID="lnkbtnAddProductToGroup" runat="server" Font-Names="Verdana"
                            Font-Size="XX-Small" OnClick="lnkbtnAddProductToGroup_Click"> &gt;&gt;&gt;&gt;&gt;</asp:LinkButton>
                    </td>
                    <td align="left">
                        <asp:ListBox ID="lbProductsInProductGroup" runat="server" Font-Names="Verdana"
                            Font-Size="XX-Small" Rows="10" SelectionMode="Multiple" Width="100%">
                        </asp:ListBox>
                    </td>
                    <td>
                        &nbsp;</td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right" valign="top">
                        <asp:Label ID="Label18" runat="server" Font-Names="Verdana"
                            Font-Size="XX-Small" Text="Order by" /></td>
                    <td valign="top">
                        <asp:RadioButtonList ID="rblOrderBy" runat="server" AutoPostBack="True" Font-Names="Verdana"
                            Font-Size="XX-Small" OnSelectedIndexChanged="rblOrderBy_SelectedIndexChanged"
                            RepeatDirection="Horizontal" CellPadding="0" CellSpacing="0">
                            <asp:ListItem Selected="True">product code</asp:ListItem>
                            <asp:ListItem>category</asp:ListItem>
                            <asp:ListItem>unarchived/archived</asp:ListItem>
                        </asp:RadioButtonList></td>
                    <td align="center" style="white-space: nowrap" valign="middle">
                    </td>
                    <td align="left">
                        <asp:Label ID="lblLegendCopyFromGroup" runat="server" Font-Names="Verdana"
                            Font-Size="XX-Small" Text="Copy products from group:" Visible="False" /><br />
                        <asp:DropDownList ID="ddlCopyProducts" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Visible="False">
                        </asp:DropDownList>
                        <asp:Button ID="btnCopyProducts" runat="server" OnClick="btnCopyProducts_Click" Text="go" Visible="False" /></td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                        &nbsp;</td>
                    <td>
                        &nbsp;</td>
                    <td>
                        &nbsp;</td>
                    <td>
                        &nbsp;</td>
                    <td align="left">
                        &nbsp;</td>
                    <td>
                        &nbsp;</td>
                </tr>
                <tr>
                    <td>
                        &nbsp;</td>
                    <td>
                        &nbsp;<asp:Label ID="Label4" runat="server" Font-Names="Verdana"
                            Font-Size="XX-Small" Text="Click this button to refresh permissions for all users.<br /><br /><font color='red'> WARNING: this may <b>remove</b> permissions from users if you have not specifically <b>granted</b> permissions using the facilities here</font>" ForeColor="#400000" Width="100%" /></td>
                    <td valign="top">
                        <strong style="color: navy; font-size:x-small; font-family:Verdana">
                        <asp:Button ID="btnRefreshAllUserPermissions" runat="server" Text="refresh all user permissions"  OnClientClick='return confirm("This will overwrite existing permissions for ALL users. Make sure you have defined all the permissions you require!! Are you sure you want to refresh all user permissions?");'
                            Width="250px" OnClick="btnRefreshAllUserPermissions_Click" />
                        </strong>
                    </td>
                    <td>
                        &nbsp;</td>
                    <td align="left" valign="top">
                        <asp:LinkButton ID="lnkbtnInitPermissions" runat="server" Font-Names="Verdana"
                            Font-Size="XX-Small" onclick="lnkbtnInitPermissions_Click" OnClientClick="return confirm(&quot;Propagate default permissions to the permissions cache for EXISTING users? Warning: this overwrites any tweaking done for existing users!! Only do this if you understand what you're doing. If in doubt, contact Chris Newport 020 8751 7524 !!&quot;);">.</asp:LinkButton>
                    </td>
                    <td>
                        &nbsp;</td>
                </tr>
            </table>
        </asp:Panel>
        <asp:Panel ID="pnlDefaultGroupAssociation" runat="server" Width="100%" Visible="False">
            <table style="width: 100%">
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td style="width: 20%">
                    </td>
                    <td style="width: 29%">
                    </td>
                    <td style="width: 20%">
                    </td>
                    <td align="right" style="width: 29%">
                    </td>
                    <td style="width: 1%">
                    </td>
                </tr>
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td style="width: 20%">
                        </td>
                    <td style="width: 29%">
                        <asp:Label ID="Label5" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="User Groups"></asp:Label></td>
                    <td style="width: 20%">
                    </td>
                    <td style="width: 29%" align="right">
                        </td>
                    <td style="width: 1%">
                    </td>
                </tr>
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td style="width: 20%"><asp:Label ID="Label8" runat="server" Font-Names="Verdana"
                            Font-Size="XX-Small" Text="Associated a product group with each user group.<br /><br />New users will then be permissioned automatically for the products in the product group associated with their user group.<br /><br />You can further refine permissions for individual users from the <b>tweak users</b> screen." ForeColor="#400000" Width="100%" /></td>
                    <td colspan="2">
                        <asp:ListBox ID="lbUserGroups" runat="server" Rows="10" Width="100%" AutoPostBack="True" OnSelectedIndexChanged="lbUserGroups_SelectedIndexChanged" Font-Names="Verdana" Font-Size="XX-Small"></asp:ListBox></td>
                    <td align="right" style="width: 29%">
                    </td>
                    <td style="width: 1%">
                    </td>
                </tr>
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td style="width: 20%">
                    </td>
                    <td style="width: 29%">
                        <asp:Label ID="lblNoProductGroupsDefinedWarning" runat="server" Font-Bold="True"
                            Font-Names="Verdana" Font-Size="XX-Small" ForeColor="Red" Text="No product groups defined!"></asp:Label></td>
                    <td style="width: 20%">
                    </td>
                    <td align="right" style="width: 29%">
                    </td>
                    <td style="width: 1%">
                    </td>
                </tr>
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td style="width: 20%">
                    </td>
                    <td id="tdProductGroups" runat="server" visible="false" colspan="2">
                        <asp:Label ID="Label6" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="Associate product group "></asp:Label>
                        <asp:DropDownList ID="ddlProductGroups" runat="server" Font-Names="Verdana" Font-Size="XX-Small" AutoPostBack="True" OnSelectedIndexChanged="ddlProductGroups_SelectedIndexChanged">
                        </asp:DropDownList>
                        <asp:Label ID="Label7" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="with user group "></asp:Label>
                        <asp:Label ID="lblUserGroup" runat="server" Font-Bold="True" Font-Names="Verdana"
                            Font-Size="XX-Small" Text="XXX"></asp:Label></td>
                    <td align="right" style="width: 29%">
                    </td>
                    <td style="width: 1%">
                    </td>
                </tr>
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td style="width: 20%">
                    </td>
                    <td style="width: 29%">
                        </td>
                    <td style="width: 20%">
                    </td>
                    <td align="right" style="width: 29%">
                    </td>
                    <td style="width: 1%">
                    </td>
                </tr>
            </table>
        </asp:Panel>
        <asp:Panel ID="pnlTweakUsers" runat="server" Width="100%" Visible="False">
            <table style="width: 100%">
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td style="width: 20%">
                    </td>
                    <td style="width: 29%">
                    </td>
                    <td style="width: 20%">
                    </td>
                    <td align="right" style="width: 29%">
                    </td>
                    <td style="width: 1%">
                    </td>
                </tr>
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td style="width: 20%">
                        </td>
                    <td style="width: 29%">
                        <asp:Label ID="lblLegendUsers" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="Users"></asp:Label></td>
                    <td style="width: 20%">
                    </td>
                    <td style="width: 29%" align="right">
                        </td>
                    <td style="width: 1%">
                    </td>
                </tr>
                <tr>
                    <td>
                        &nbsp;</td>
                    <td>
                        &nbsp;<asp:Label ID="Label9" runat="server" Font-Names="Verdana"
                            Font-Size="XX-Small" Text="After assigning product groups to user groups, fine tune permissions for individual users" ForeColor="#400000" Width="100%" /></td>
                    <td colspan="2">
                        <asp:ListBox ID="lbUsers" runat="server" Rows="10" Width="100%" Font-Names="Verdana" Font-Size="XX-Small" AutoPostBack="True" OnSelectedIndexChanged="lbUsers_SelectedIndexChanged"></asp:ListBox>
                    </td>
                    <td align="left" valign="top">
                        &nbsp;<asp:RadioButton
                            ID="rbUsersAll" runat="server" Checked="True" Font-Names="Verdana" Font-Size="XX-Small"
                            GroupName="users" Text="all users" OnCheckedChanged="rbUsersAll_CheckedChanged" AutoPostBack="True" /><br />
                            &nbsp;<asp:RadioButton ID="rbUsersWithoutPermissions" runat="server" Font-Names="Verdana"
                                Font-Size="XX-Small" GroupName="users" Text="users without permissions" OnCheckedChanged="rbUsersWithoutPermissions_CheckedChanged" AutoPostBack="True" /><br />
                            &nbsp;<asp:RadioButton ID="rbUsersMatching" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                GroupName="users" Text="users matching..." OnCheckedChanged="rbUsersMatching_CheckedChanged" AutoPostBack="True" />
                            <asp:TextBox ID="tbUserSearch" runat="server" Visible="False" Font-Names="Verdana" Font-Size="XX-Small"></asp:TextBox><br />
                            &nbsp;<asp:RadioButton ID="rbUsersInGroup" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                GroupName="users" Text="users in group..." OnCheckedChanged="rbUsersInGroup_CheckedChanged" AutoPostBack="True" />
                            <asp:DropDownList ID="ddlUserGroup" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Visible="False">
                            </asp:DropDownList><br />
                            &nbsp;&nbsp;<br />
                            &nbsp;<asp:Button ID="btnGo" runat="server" Text="go" Width="150px" OnClick="btnGo_Click" />
                        </td>
                    <td>
                        &nbsp;</td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td>
                    </td>
                    <td>
                        <asp:Label ID="Label25" runat="server" Font-Names="Verdana"
                            Font-Size="XX-Small" Text="Available product groups:" />
                    </td>
                    <td>
                    </td>
                    <td align="left">
                        <asp:Label ID="lblLegendUserNamePermissioning" runat="server" Font-Names="Verdana"
                            Font-Size="XX-Small" Text="Groups & products:" />
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td valign="top">&nbsp;</td>
                    <td>
                        <asp:ListBox ID="lbAvailableGroupsForUser" runat="server" Rows="8"
                            SelectionMode="Multiple" Width="100%" Font-Names="Verdana"
                            Font-Size="XX-Small" >
                        </asp:ListBox></td>
                    <td align="center" valign="middle" style="white-space:nowrap">
                        &nbsp;<asp:LinkButton ID="lnkbtnRemoveGroupFromUser" runat="server"
                            Font-Names="Verdana" Font-Size="XX-Small" OnClick="lnkbtnRemoveGroupFromUser_Click"><<<<< </asp:LinkButton>&nbsp;
                        &nbsp;&nbsp;
                        <asp:LinkButton ID="lnkbtnAddGroupToUser" runat="server" Font-Names="Verdana"
                            Font-Size="XX-Small" OnClick="lnkbtnAddGroupToUser_Click"> >>>>></asp:LinkButton></td>
                    <td align="right" rowspan="3" valign="top">
                        <asp:ListBox ID="lbPermissionedGroupsAndProducts" runat="server" Font-Names="Verdana"
                            Font-Size="XX-Small" Rows="20" SelectionMode="Multiple" Width="100%">
                        </asp:ListBox>
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                        &nbsp;</td>
                    <td valign="top">
                        &nbsp;</td>
                    <td>
                        <asp:Label ID="Label30" runat="server" Font-Names="Verdana"
                            Font-Size="XX-Small" Text="Available products:" />
                    </td>
                    <td align="center" style="white-space:nowrap" valign="middle">
                        &nbsp;</td>
                    &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; &nbsp;&nbsp;
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                    &nbsp;&nbsp; &nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; &nbsp;&nbsp;
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                    &nbsp;&nbsp; &nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; &nbsp;&nbsp;
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                    &nbsp;&nbsp; &nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; &nbsp;&nbsp;
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                    &nbsp;&nbsp; &nbsp;<td>
                        &nbsp;</td>
                </tr>
                <tr>
                    <td>
                        &nbsp;</td>
                    <td valign="top">
                        &nbsp;</td>
                    <td>
                        <asp:ListBox ID="lbAvailableProductsForUser" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Rows="10" SelectionMode="Multiple" Width="100%">
                        </asp:ListBox>
                    </td>
                    <td align="center" style="white-space:nowrap" valign="middle">
                        <asp:LinkButton ID="lnkbtnRemoveProductFromUser" runat="server" Font-Names="Verdana" Font-Size="XX-Small" OnClick="lnkbtnRemoveProductFromUser_Click">&lt;&lt;&lt;&lt;&lt;
                        </asp:LinkButton>
                        &nbsp; &nbsp;&nbsp;&nbsp;<asp:LinkButton ID="lnkbtnAddProductToUser" runat="server" Font-Names="Verdana" Font-Size="XX-Small" OnClick="lnkbtnAddProductToUser_Click">
                        &gt;&gt;&gt;&gt;&gt;</asp:LinkButton>
                    </td>
                    <td>
                        &nbsp;</td>
                </tr>
                <tr>
                    <td>
                        &nbsp;</td>
                    <td valign="top">
                        &nbsp;</td>
                    <td>
                        &nbsp;</td>
                    <td align="center" style="white-space:nowrap" valign="middle">
                        &nbsp;</td>
                    <td align="right">
                        &nbsp;</td>
                    <td>
                        &nbsp;</td>
                </tr>
                <tr>
                    <td>
                        &nbsp;</td>
                    <td valign="top">
                        &nbsp;</td>
                    <td>
                        <strong style="color: navy; font-size:x-small; font-family:Verdana">&nbsp;</strong></td>
                    <td align="center" style="white-space:nowrap" valign="middle">
                        &nbsp;</td>
                    <td align="right">
                        &nbsp;</td>
                    <td>
                        &nbsp;</td>
                </tr>
            </table>
        </asp:Panel>
        <asp:Panel ID="pnlNewProductGroup" runat="server" Width="100%" Visible="False">
            <table style="width: 100%">
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td style="width: 20%">
                        <asp:Label ID="Label3" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="New product group"></asp:Label></td>
                    <td style="width: 29%">
                    </td>
                    <td style="width: 20%">
                    </td>
                    <td align="right" style="width: 29%">
                    </td>
                    <td style="width: 1%">
                    </td>
                </tr>
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td style="width: 20%" align="right">
                        <asp:Label ID="Label1" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="Product group name:"></asp:Label></td>
                    <td colspan="3">
                        <asp:TextBox ID="tbNewProductGroupName" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            MaxLength="50" Width="200px"></asp:TextBox>
                        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                        &nbsp;&nbsp;
                    </td>
                    <td style="width: 1%">
                    </td>
                </tr>
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td align="right" style="width: 20%">
                    </td>
                    <td colspan="3">
                        <asp:Label ID="Label16" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small"
                            ForeColor="#400000" Text="On create"></asp:Label>
                        <asp:RadioButton ID="rbOnCreateNoAction" runat="server" Checked="True" Font-Names="Verdana"
                            Font-Size="XX-Small" GroupName="OnCreate" OnCheckedChanged="rbOnCreateNoAction_CheckedChanged"
                            Text="just create group" AutoPostBack="True" />
                        <asp:RadioButton ID="rbOnCreateAddGroupAllUsers" runat="server" Font-Names="Verdana"
                            Font-Size="XX-Small" GroupName="OnCreate" Text="add group to all users" OnCheckedChanged="rbOnCreateAddGroupAllUsers_CheckedChanged" AutoPostBack="True" />
                        <asp:RadioButton ID="rbOnCreateAddGroupUserGroup" runat="server" Font-Names="Verdana"
                            Font-Size="XX-Small" GroupName="OnCreate" Text="add group to users in user group" OnCheckedChanged="rbOnCreateAddGroupUserGroup_CheckedChanged" AutoPostBack="True" />
                        <asp:DropDownList ID="ddlOnCreateUserGroups" runat="server" Font-Names="Verdana"
                            Font-Size="XX-Small" Visible="False">
                        </asp:DropDownList></td>
                    <td style="width: 1%">
                    </td>
                </tr>
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td align="right" style="width: 20%">
                    </td>
                    <td style="width: 29%">
                        <asp:Button ID="btnCreateNewProductGroup" runat="server" OnClick="btnCreateNewProductGroup_Click"
                            Text="create" Width="80px" />
                        &nbsp;&nbsp;
                        <asp:Button ID="btnCancelNewProductGroup" runat="server" Text="cancel" Width="80px" OnClick="btnCancelNewProductGroup_Click" />
                        </td>
                    <td style="width: 20%">
                    </td>
                    <td align="right" style="width: 29%">
                    </td>
                    <td style="width: 1%">
                    </td>
                </tr>
            </table>
        </asp:Panel>
        <asp:Panel ID="pnlRenameProductGroup" runat="server" Width="100%" Visible="False">
            <table style="width: 100%">
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td style="width: 20%">
                        <asp:Label ID="Label3rn" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="Rename product group"></asp:Label>
                        <asp:Label ID="lblRenameCurrentName" runat="server" Font-Bold="True" Font-Names="Verdana"
                            Font-Size="XX-Small"></asp:Label></td>
                    <td style="width: 29%">
                    </td>
                    <td style="width: 20%">
                    </td>
                    <td align="right" style="width: 29%">
                    </td>
                    <td style="width: 1%">
                    </td>
                </tr>
                <tr>
                    <td style="width: 1%; height: 22px;">
                    </td>
                    <td style="width: 20%; height: 22px;" align="right">
                        <asp:Label ID="Label1rn" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="New name:"></asp:Label></td>
                    <td colspan="3" style="height: 22px">
                        <asp:TextBox ID="tbRenameProductGroupNewName" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            MaxLength="50" Width="200px"></asp:TextBox>
                        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                        &nbsp;&nbsp;
                    </td>
                    <td style="width: 1%; height: 22px;">
                    </td>
                </tr>
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td align="right" style="width: 20%">
                    </td>
                    <td colspan="3">
                        &nbsp;</td>
                    <td style="width: 1%">
                    </td>
                </tr>
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td align="right" style="width: 20%">
                    </td>
                    <td style="width: 29%">
                        <asp:Button ID="btnDoRenameProductGroup" runat="server"
                            Text="rename" Width="80px" OnClick="btnDoRename_Click" />
                        &nbsp;&nbsp;
                        <asp:Button ID="btnCancelRenameProductGroup" runat="server" Text="cancel" Width="80px" OnClick="btnCancelRenameProductGroup_Click" />
                        </td>
                    <td style="width: 20%">
                    </td>
                    <td align="right" style="width: 29%">
                    </td>
                    <td style="width: 1%">
                    </td>
                </tr>
            </table>
        </asp:Panel>
        <asp:Panel ID="pnlHelp" runat="server" Width="100%" Visible="True">
            <table style="width: 100%">
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td style="width: 20%">
                    </td>
                    <td style="width: 29%">
                    </td>
                    <td style="width: 20%">
                    </td>
                    <td align="right" style="width: 29%">
                    </td>
                    <td style="width: 1%">
                    </td>
                </tr>
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td colspan="3">
                        <asp:Label ID="Label10" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="HOW TO USE THE USER PERMISSIONS FACILITY" ForeColor="#400000"></asp:Label><br />
                        <br />
                        <asp:Label ID="Label14" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="This facility lets you control which users can view which products." ForeColor="#400000"></asp:Label><br />
                        <br />
                        <asp:Label ID="Label11" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="1. Define product groups and insert the products you require into each group." ForeColor="#400000"></asp:Label><br />
                        <br />
                        <asp:Label ID="Label12" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="2. Ensure the users to whom you want to assign product permissions have been placed in a user group. Alternatively you can assign permissions to individual users using the <b>tweak users</b> screen." ForeColor="#400000"></asp:Label><br />
                        <br />
                        <asp:Label ID="Label13" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="3. Assign a product group to each user group. When new users are added to the system they will automatically inherit the permissions for the group to which they are assigned. You can refresh <b>all</b> user permissions by clicking the <b>refresh all user permissions</b> button." ForeColor="#400000"></asp:Label><br />
                        <br />
                        <asp:Label ID="Label15" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="4. If necessary, fine tune permissions for individual users from the <b>tweak users</b> screen." ForeColor="#400000"></asp:Label><br />
                        <br />
                        <asp:Label ID="Label17" runat="server" Font-Bold="True" Font-Italic="True" Font-Names="Verdana"
                            Font-Size="XX-Small" ForeColor="#400000" Text="For further information click the help button"></asp:Label></td>
                    <td style="width: 29%" align="right">
                        </td>
                    <td>
                    </td>
                </tr>
            </table>
        </asp:Panel>
        <asp:Panel ID="pnlUserGroups" runat="server" Width="100%" Visible="False">
            <table style="width: 100%">
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td style="width: 20%">
                        <asp:Label ID="Label3xx" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="User groups"></asp:Label></td>
                    <td style="width: 29%">
                    </td>
                    <td style="width: 20%">
                    </td>
                    <td align="right" style="width: 29%">
                    </td>
                    <td style="width: 1%">
                    </td>
                </tr>
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td align="right" style="width: 20%" valign="top">
                        <asp:Label ID="Label1xx" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="User groups:"></asp:Label></td>
                    <td colspan="2">
                        <asp:ListBox ID="lbDefinedUserGroups" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Rows="10" Width="100%" OnSelectedIndexChanged="lbDefinedUserGroups_SelectedIndexChanged" AutoPostBack="True"></asp:ListBox></td>
                    <td align="left" style="width: 29%" valign="top">
                        &nbsp; &nbsp;
                        <asp:Button ID="btnNewUserGroup" runat="server" Text="new user group" Width="200px" OnClick="btnNewUserGroup_Click" /><br />
                        <br />
                        &nbsp; &nbsp;
                        <asp:Button ID="btnRenameUserGroup" runat="server" Text="rename user group" Width="200px" OnClick="btnRenameUserGroup_Click" Enabled="False" /><br />
                        <br />
                        &nbsp; &nbsp;
                        <asp:Button ID="btnRemoveUserGroup" runat="server" Text="remove user group" Width="200px" OnClick="btnRemoveUserGroup_Click" OnClientClick='return confirm("Are you sure you want to remove this user group? User accounts will NOT be removed, but all links and associations with the user group WILL be removed.");' Enabled="False" />
                    </td>
                    <td style="width: 1%">
                    </td>
                </tr>
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td style="width: 20%" align="right">
                        </td>
                    <td colspan="3">
                        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                        &nbsp;&nbsp;
                    </td>
                    <td style="width: 1%">
                    </td>
                </tr>
            </table>
        </asp:Panel>
        <asp:Panel ID="pnlNewUserGroup" runat="server" Width="100%" Visible="False">
            <table style="width: 100%">
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td style="width: 20%">
                        <asp:Label ID="Label19" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="New user group"></asp:Label></td>
                    <td style="width: 29%">
                    </td>
                    <td style="width: 20%">
                    </td>
                    <td align="right" style="width: 29%">
                    </td>
                    <td style="width: 1%">
                    </td>
                </tr>
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td style="width: 20%" align="right">
                        <asp:Label ID="Label20" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="User group name:"></asp:Label></td>
                    <td colspan="3">
                        <asp:TextBox ID="tbNewUserGroupName" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            MaxLength="50" Width="200px"/>
                        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                        &nbsp;&nbsp;
                    </td>
                    <td style="width: 1%">
                    </td>
                </tr>
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td align="right" style="width: 20%">
                    </td>
                    <td colspan="3">
                    </td>
                    <td style="width: 1%">
                    </td>
                </tr>
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td align="right" style="width: 20%">
                    </td>
                    <td colspan="3">
                        <asp:Button ID="btnCreateNewUserGroup" runat="server" Text="create" OnClick="btnCreateNewUserGroup_Click" />
                        &nbsp; &nbsp;<asp:Button ID="btnCancelNewUserGroup" runat="server" Text="cancel" OnClick="btnCancelNewUserGroup_Click" /></td>
                    <td style="width: 1%">
                    </td>
                </tr>
            </table>
        </asp:Panel>
        <asp:Panel ID="pnlRenameUserGroup" runat="server" Width="100%" Visible="False">
            <table style="width: 100%">
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td style="width: 20%">
                        <asp:Label ID="Label21" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="Rename user group"></asp:Label></td>
                    <td style="width: 29%">
                    </td>
                    <td style="width: 20%">
                    </td>
                    <td align="right" style="width: 29%">
                    </td>
                    <td style="width: 1%">
                    </td>
                </tr>
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td style="width: 20%" align="right">
                        <asp:Label ID="Label22" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="User group name:"></asp:Label></td>
                    <td colspan="3">
                        <asp:TextBox ID="tbRenameUserGroupNewName" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            MaxLength="50" Width="200px"/>
                        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                        &nbsp;&nbsp;
                    </td>
                    <td style="width: 1%">
                    </td>
                </tr>
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td align="right" style="width: 20%">
                    </td>
                    <td colspan="3">
                    </td>
                    <td style="width: 1%">
                    </td>
                </tr>
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td align="right" style="width: 20%">
                    </td>
                    <td colspan="3">
                        <asp:Button ID="btnDoRenameUserGroup" runat="server" Text="rename" OnClick="btnDoRenameUserGroup_Click" />
                        &nbsp;&nbsp;
                        <asp:Button ID="btnCancelRenameUserGroup" runat="server" Text="cancel" OnClick="btnCancelRenameUserGroup_Click" /></td>
                    <td style="width: 1%">
                    </td>
                </tr>
            </table>
        </asp:Panel>
        <asp:Panel ID="pnlMaxGrabs" runat="server" Width="100%" Visible="False">
            <table style="width: 100%">
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td style="width: 20%">
                        <asp:Label ID="Label24" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="Set default max grabs"></asp:Label></td>
                    <td style="width: 29%">
                    </td>
                    <td style="width: 20%">
                    </td>
                    <td align="right" style="width: 29%">
                    </td>
                    <td style="width: 1%">
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td>
                    </td>
                    <td>
                        <asp:Label ID="Label26" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="User groups:"></asp:Label></td>
                    <td>
                    </td>
                    <td align="right">
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right" valign="top">
                        </td>
                    <td colspan="2">
                        <asp:ListBox ID="lbMaxGrabUserGroups" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Rows="6" Width="100%" AutoPostBack="True" OnSelectedIndexChanged="lbMaxGrabUserGroups_SelectedIndexChanged"></asp:ListBox></td>
                    <td align="center">
                        &nbsp;<asp:Button ID="btnApplyAllMaxGrabs" runat="server" OnClick="btnApplyAllMaxGrabs_Click"
                            OnClientClick='return confirm("This will apply all max grabs you have defined, for ALL user groups. Are you sure you want to apply all max grabs?");'
                            Text="apply all max grabs!" /><br />
                        <br />
                        <asp:CheckBox ID="cbDontKeepRemindingMe" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="don't keep reminding me" /></td>
                    <td>
                    </td>
                </tr>
                <tr ID="trMaxGrabProducts01" runat="server" visible="true">
                    <td>
                        </td>
                    <td align="right" valign="baseline">
                        <asp:Label ID="Label27" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small" Text="Show"/>
                    </td>
                    <td colspan="2">
                        <asp:RadioButton ID="rblMaxGrabAllProducts" runat="server" AutoPostBack="True"
                            Checked="True" Font-Names="Verdana" Font-Size="XX-Small" GroupName="MaxGrab"
                            OnCheckedChanged="rblMaxGrabAllProducts_CheckedChanged" Text="all products" />
                        <asp:RadioButton ID="rblMaxGrabProductsInGroup" runat="server"
                            AutoPostBack="True" Font-Names="Verdana" Font-Size="XX-Small"
                            GroupName="MaxGrab" OnCheckedChanged="rblMaxGrabProductsInGroup_CheckedChanged"
                            Text="products in product group" />
                        <asp:DropDownList ID="ddlMaxGrabProductGroup" runat="server"
                            AutoPostBack="True" Font-Names="Verdana" Font-Size="XX-Small"
                            onselectedindexchanged="ddlMaxGrabProductGroup_SelectedIndexChanged">
                        </asp:DropDownList>
                    </td>
                    <td align="right">
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
                        &nbsp;</td>
                    <td>
                    </td>
                    <td id="tdCopyMaxGrabs01" runat="server" visible="false" align="left">
                        <asp:Label ID="Label37" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="Copy"></asp:Label>
                        <asp:Label ID="lblCopyMaxGrabsSource" runat="server" Font-Bold="True" Font-Names="Verdana"
                            Font-Size="XX-Small"></asp:Label>
                        <asp:Label ID="Label36" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="group's max grabs to"></asp:Label></td>
                    <td>
                    </td>
                </tr>
                <tr id="trMaxGrabProducts02" runat="server" visible="false">
                    <td>
                    </td>
                    <td>
                        </td>
                    <td colspan="2">
                        <asp:Label ID="Label31" runat="server" Font-Bold="False" Font-Names="Verdana"
                            Font-Size="XX-Small" Text="Product max grabs for users in group"></asp:Label>
                        <asp:Label ID="lblMaxGrabUserGroup" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Font-Bold="True"></asp:Label></td>
                    <td align="left" id="tdCopyMaxGrabs02" runat="server" visible="false" >
                        <asp:Label ID="Label34" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="group"></asp:Label>
                        <asp:DropDownList ID="ddlCopyMaxGrabs" runat="server" Font-Names="Verdana" Font-Size="XX-Small" AutoPostBack="True" OnSelectedIndexChanged="ddlCopyMaxGrabs_SelectedIndexChanged">
                        </asp:DropDownList>
                        <asp:Button ID="btnCopyMaxGrabs" runat="server" Text="go" OnClick="btnCopyMaxGrabs_Click" /></td>
                    <td>
                    </td>
                </tr>
                <tr id="trMaxGrabProducts03" runat="server" visible="false">
                    <td>
                    </td>
                    <td align="right" valign="top">
                    </td>
                    <td colspan="3">
                        <asp:GridView ID="gvMaxGrabProducts" runat="server" CellPadding="2"
                            Font-Names="Verdana" Font-Size="XX-Small" Width="100%" OnRowDataBound="gvMaxGrabProducts_RowDataBound" AutoGenerateColumns="False">
                            <Columns>
                                <asp:TemplateField HeaderText="Max Grab">
                                    <ItemTemplate>
                                        <asp:TextBox ID="tbMaxGrab" runat="server" Font-Names="Verdana"
                                            Font-Size="XX-Small" Width="40px"></asp:TextBox>
                                        <asp:RegularExpressionValidator ID="revMaxGrab" runat="server" ControlToValidate="tbMaxGrab"
                                            ErrorMessage="###" ValidationExpression="\d*"></asp:RegularExpressionValidator>
                                        <asp:HiddenField ID="hidLogisticProductKey" Value='<%# Bind("LogisticProductKey") %>' runat="server" />
                                    </ItemTemplate>
                                    <ItemStyle HorizontalAlign="Center" />
                                    <HeaderTemplate>
                                        <asp:Label ID="Label33" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small"
                                            Text="Max Grabs"></asp:Label>
                                        <asp:Button ID="btnSaveMaxGrab" runat="server" onclick="btnSaveMaxGrab_Click"
                                            Text="save changes" Width="100px" />
                                    </HeaderTemplate>
                                </asp:TemplateField>
                                <asp:BoundField DataField="ProductCode" HeaderText="Product Code" />
                                <asp:BoundField DataField="ProductDate" HeaderText="Value / Date" />
                                <asp:BoundField DataField="ProductDescription" HeaderText="Description" />
                            </Columns>
                            <EmptyDataTemplate>
                                <asp:Label ID="Label26" runat="server" Font-Bold="False" Font-Names="Verdana"
                                    Font-Size="XX-Small" Text="no products found"></asp:Label>
                            </EmptyDataTemplate>
                        </asp:GridView>
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
                    <td align="right">
                    </td>
                    <td>
                    </td>
                </tr>
            </table>
        </asp:Panel>
    </form>
</body>
</html>
