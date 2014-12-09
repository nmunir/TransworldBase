<%@ Page Language="VB" Theme="AIMSDefault" %>

<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="Telerik.Web.UI" %>
<%@ Register Assembly="Telerik.Web.UI" Namespace="Telerik.Web.UI" TagPrefix="telerik" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<script runat="server">

    ' TO DO
    ' add paging
    ' need a regular job to set topics to CLOSED (20) - UPDATE MessagingTopics SET TopicStatus = 20 WHERE LastTopicStateChange < GETDATE() - 1 AND TopicStatus IN (7, 17)
    ' write description of functionality
    ' add help
    
    ' do we need a search facility for the controller?
    
    ' MessagingTopics.IsClosed takes values: 0=OPEN; 1=CLOSED BY USER; 4=CLOSED BY USER, AWAITING FEEDBACK; 7=CLOSED BY USER, FEEDBACK RECEIVED; 10=CLOSED BY CONTROLLER; 20=CLOSED
    
    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    Dim sSQL As String = String.Empty
    Dim gbNewMessages As Boolean
    
    Const TOPIC_STATUS_OPEN As Int32 = 0                                      ' topic is open
    Const TOPIC_STATUS_CLOSED_BY_USER As Int32 = 1                            ' agent requested topic to be closed; controller now needs to set to TOPIC_STATUS_CLOSED_BY_USER_AWAITING_FEEDBACK or TOPIC_STATUS_CLOSED_BY_USER_FEEDBACK_RECEIVED
    Const TOPIC_STATUS_CLOSED_BY_USER_AWAITING_FEEDBACK As Int32 = 4          ' agent requested topic to be closed; controller set status to TOPIC_STATUS_CLOSED_BY_USER_AWAITING_FEEDBACK; now awaiting agent feedback
    Const TOPIC_STATUS_CLOSED_BY_USER_FEEDBACK_RECEIVED As Int32 = 7          ' agent requested topic to be closed; requested feedback; feedback now received or is not required
    Const TOPIC_STATUS_CLOSED_BY_CONTROLLER_AWAITING_FEEDBACK As Int32 = 14   ' controller closed topic, requested feedback; now awaiting agent feedback
    Const TOPIC_STATUS_CLOSED_BY_CONTROLLER_FEEDBACK_RECEIVED As Int32 = 17   ' controller closed topic, feedback has either been received or is not required
    Const TOPIC_STATUS_CLOSED As Int32 = 20                                   ' controller closed topic, feedback has either been received or is not required

    'Const TOPIC_STATUS_CLOSED_BY_CONTROLLER As Int32 = 10                     ' NOT USED

    Const TOPICGRID_ICONS As Int32 = 0
    Const TOPICGRID_TOPIC As Int32 = 1
    Const TOPICGRID_CREATEDON As Int32 = 2
    Const TOPICGRID_CONSIGNMENT As Int32 = 3
    Const TOPICGRID_TOPICREF As Int32 = 4
    Const TOPICGRID_STATUS As Int32 = 5
    Const TOPICGRID_RATING As Int32 = 6
    
    Const MESSAGEGRID_DATE As Int32 = 0
    Const MESSAGEGRID_MESSAGE As Int32 = 1
    Const MESSAGEGRID_AUTHOR As Int32 = 2
    
    Const TITLE_AGENT As String = "AGENT"
    Const TITLE_SYSTEM As String = "SYSTEM"
    Const TITLE_CONTROLLER As String = "WESTERN UNION"

    Const USER_PERMISSION_WU_INTERNAL_USER As Integer = &H8000

    Const USER_PERMISSION_WU_IS_TSE As Integer = &H200000

    Const ITEMS_PER_REQUEST As Integer = 30
    Const TOPIC_RECEIPT_CONFIRMATION_MESSAGE As String = "Thank you. Your topic has been received and is being reviewed."

    Public ReadOnly COLOUR_TOPICSTATUS_CLOSING As System.Drawing.Color = Drawing.Color.Orange
    Public ReadOnly COLOUR_NEWMESSAGE As System.Drawing.Color = Drawing.Color.PaleGreen
    Public ReadOnly COLOUR_MESSAGEAUTHOR_AGENT As System.Drawing.Color = Drawing.Color.PaleGreen
    Public ReadOnly COLOUR_MESSAGEAUTHOR_CONTROLLER As System.Drawing.Color = Drawing.Color.LightYellow

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsNumeric(Session("CustomerKey")) Then
            Server.Transfer("session_expired.aspx")
            Exit Sub
        End If
        If Not IsPostBack Then
            Call CheckLastClosed()
            Call SetTitle()
            Call HideAllPanels()
            Call ShowTopics()
            If (CInt(Session("UserPermissions")) And USER_PERMISSION_WU_IS_TSE) > 0 Then
                trOnBehalfOf.Visible = True
            Else
                trRaisedBy.Visible = True
            End If
            If Not (Session("UserPermissions") And USER_PERMISSION_WU_INTERNAL_USER) > 0 Then
                rbAgentTitles.Visible = False
                rbWUInternalTitles.Visible = False
            End If
        End If
    End Sub
    
    Protected Sub CheckLastClosed() ' NOT CURRENTLY CALLED
        Dim sbSQL As New StringBuilder
        sbSQL.Append("DECLARE @LastCloseDate as smalldatetime ")
        sbSQL.Append("SET @LastCloseDate = (SELECT ISNULL(LastCloseDate,'1-Jan-2000') FROM MessagingConfiguration WHERE CustomerKey = 0) ")
        sbSQL.Append("IF @LastCloseDate <> REPLACE(CONVERT(VARCHAR(11), GETDATE(), 106), ' ', '-') + ' 00:00:00' ")
        sbSQL.Append("BEGIN ")
        sbSQL.Append("  UPDATE MessagingTopics SET TopicStatus = 20 WHERE TopicStatus IN (7, 17) AND LastTopicStateChange < REPLACE(CONVERT(VARCHAR(11), GETDATE() - 1, 106), ' ', '-') + ' 23:59:59' ")
        sbSQL.Append("  UPDATE MessagingConfiguration SET LastCloseDate = REPLACE(CONVERT(VARCHAR(11), GETDATE(), 106), ' ', '-') + ' 00:00:00' WHERE CustomerKey = 0 ")
        sbSQL.Append("END ")
        Call ExecuteQueryToDataTable(sbSQL.ToString)
    End Sub
    ' 
    Protected Sub btnNewTopicSend_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Page.Validate("newtopic")
        If Page.IsValid Then
            If IsValidMessage() Then
                Call SendInitialTopicMessage()
                Call ShowMessages(pnTopicID)
            End If
        End If
    End Sub
    
    Protected Sub SendInitialTopicMessage()
        pnTopicID = CreateNewTopic()
        Dim sMessageTimestamp As String = Format(DateTime.Now, "yyyyMMddhhmmss")
        Call SaveMessage(pnTopicID, GetTopicReferenceFromID(pnTopicID) & "." & sMessageTimestamp, tbNewTopicMessage.Text)
        If trOnBehalfOf.Visible Then
            If IsNumeric(rcbUser.SelectedValue) Then
                Call ExecuteQueryToDataTable("UPDATE MessagingTopics SET OnBehalfOf = " & rcbUser.SelectedValue & " WHERE [id] = " & pnTopicID)
            End If
        End If
        Call SendNewTopicEmailAlertToController()
        Call SaveSystemMessage(pnTopicID, GetSystemTopicReferenceFromID(pnTopicID) & "." & sMessageTimestamp, TOPIC_RECEIPT_CONFIRMATION_MESSAGE)
    End Sub
    
    Protected Function CreateNewTopic() As Int32
        CreateNewTopic = 0
        Dim sAWB As String = String.Empty
        If ddlNewTopicConsignmentReference.SelectedIndex > 0 Then
            sAWB = ddlNewTopicConsignmentReference.SelectedValue
        End If
        Dim sTopic As String
        If ddlTopic.SelectedItem.Text.ToLower.Contains("other") Then
            sTopic = tbSubject.Text.Trim
        Else
            sTopic = ddlTopic.SelectedItem.Text
        End If
        Dim nOnBehalfOf As Int32 = 0
        Dim sSQL As String = "INSERT INTO MessagingTopics (UserKey, TopicStatus, Topic, TopicReference, LastTopicStateChange, AWB, NewMessage, Rating, OnBehalfOf, RaisedBy, ClosedBy, CreatedOn, CreatedBy) VALUES ("
        sSQL += Session("UserKey") & ", " & TOPIC_STATUS_OPEN & ", '" & sTopic.Replace("'", "''") & "', '', GETDATE(), '" & sAWB & "', 0, 0, 0, '" & tbRaisedBy.Text.Replace("'", "''") & "', '', GETDATE(), " & Session("UserKey") & ") SELECT SCOPE_IDENTITY()"
        Dim oDataTable As DataTable = ExecuteQueryToDataTable(sSQL)
        If oDataTable.Rows.Count = 1 Then
            CreateNewTopic = CInt(oDataTable.Rows(0).Item(0))
            pnTopicID = CInt(oDataTable.Rows(0).Item(0))
            Dim sTopicReference As String = GetTopicReferenceFromID(pnTopicID)
            sSQL = "UPDATE MessagingTopics SET TopicReference = '" & sTopicReference.Replace("'", "''") & "' WHERE [id] = " & pnTopicID
            Call ExecuteQueryToDataTable(sSQL)
        Else
            WebMsgBox.Show("CreateNewTopic: unexpectedly row count value returned was " & oDataTable.Rows.Count)
        End If
    End Function
    
    Protected Function GetTopicReferenceFromID(ByVal nTopicNumber As Int32) As String
        GetTopicReferenceFromID = ExecuteQueryToDataTable("SELECT UserID FROM UserProfile WHERE [key] = " & Session("UserKey")).Rows(0).Item(0) & "_" & nTopicNumber.ToString.PadLeft(6, "0")
    End Function

    Protected Function GetSystemTopicReferenceFromID(ByVal nTopicNumber As Int32) As String
        GetSystemTopicReferenceFromID = "SYSTEM_" & nTopicNumber.ToString.PadLeft(7, "0")
    End Function

    Protected Sub SaveMessage(ByVal nTopicNumber As Int32, ByVal sMessageReference As String, ByVal sMessage As String, Optional ByVal NoNewMessage As Boolean = False)
        Dim sSQL As String = "INSERT INTO MessagingMessages (CustomerKey, TopicNumber, MessageRef, MessageBody, IsDeleted, IsAdmin, CreatedOn, CreatedBy) VALUES ("
        sSQL += Session("CustomerKey") & ", " & nTopicNumber.ToString & ", '" & sMessageReference.Replace("'", "''") & "', '" & sMessage.Trim.Replace("'", "''") & "', 0, 0, GETDATE(), " & Session("UserKey") & ")"
        If Not NoNewMessage Then
            sSQL += " UPDATE MessagingTopics SET NewAgentMessage = 1 WHERE ID = " & nTopicNumber
        End If
        Call ExecuteQueryToDataTable(sSQL)
    End Sub
    
    Protected Sub SaveSystemMessage(ByVal nTopicNumber As Int32, ByVal sMessageReference As String, ByVal sMessage As String)
        Dim sSQL As String = "INSERT INTO MessagingMessages (CustomerKey, TopicNumber, MessageRef, MessageBody, IsDeleted, IsAdmin, CreatedOn, CreatedBy) VALUES ("
        sSQL += Session("CustomerKey") & ", " & nTopicNumber.ToString & ", '" & sMessageReference.Replace("'", "''") & "', '" & sMessage.Trim.Replace("'", "''") & "', 0, 0, GETDATE(), 0)"
        sSQL += " UPDATE MessagingTopics SET NewAgentMessage = 1 WHERE ID = " & nTopicNumber
        Call ExecuteQueryToDataTable(sSQL)
    End Sub
    
    Protected Function GetTopicTitle(ByVal sTopicID As String) As String
        Dim sSQL As String = "SELECT Topic, ISNULL(OnBehalfOf, 0) 'OnBehalfOf', ISNULL(RaisedBy, '') RaisedBy FROM MessagingTopics WHERE [id] = " & sTopicID
        Dim dr As DataRow = ExecuteQueryToDataTable(sSQL).Rows(0)
        GetTopicTitle = GetTopicReferenceFromID(CInt(sTopicID)) & ": " & dr("Topic")
        'If Not IsDBNull(dr("RaisedBy")) Then
        If dr("RaisedBy").ToString.Trim <> String.Empty Then
            GetTopicTitle += " (topic raised by " & dr("RaisedBy").ToString.Trim & ")"
        End If
        'End If
        'If Not IsDBNull(dr("OnBehalfOf")) Then
        If CInt(dr("OnBehalfOf")) > 0 Then
            GetTopicTitle += " (on behalf of " & GetOnBehalfOfUserIDFromKey(dr("OnBehalfOf")) & ")"
        End If
        'End If
    End Function

    Protected Function GetTopicConsignmentRef(ByVal sTopicID As String) As String
        Dim sSQL As String = "SELECT AWB FROM MessagingTopics WHERE [id] = " & sTopicID
        Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
        GetTopicConsignmentRef = dt.Rows(0).Item(0).ToString.Trim
    End Function
    
    Protected Sub btnAddMessageSend_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        tbAddMessage.Text = tbAddMessage.Text.Trim
        If tbAddMessage.Text = String.Empty Then
            WebMsgBox.Show("Please enter some text for your message.")
            Exit Sub
        End If
        Call AddMessage()
        If ddlAddMessageConsignmentReference.SelectedIndex > 0 Then
            Dim sAWB As String = ddlAddMessageConsignmentReference.SelectedValue
            Dim sSQL As String = "UPDATE MessagingTopics SET AWB = '" & sAWB & "' WHERE [id] = " & pnTopicID
            Call ExecuteQueryToDataTable(sSQL)
        End If
        Call ShowMessages(pnTopicID)
    End Sub

    Protected Sub AddMessage()
        Dim sMessageTimestamp As String = Format(DateTime.Now, "yyyyMMddhhmmss")
        Call SaveMessage(pnTopicID, GetTopicReferenceFromID(pnTopicID) & "." & sMessageTimestamp, tbAddMessage.Text)
        Call SendNewMessageEmailAlert()
        If cbCloseTopic.Checked Then
            Call CloseTopic(pnTopicID)
        End If
    End Sub
    
#Region "Other Methods"

    Protected Sub SetTitle()
        Dim sTitle As String = Session("SiteTitle")
        If sTitle <> String.Empty Then
            sTitle += " - "
        End If
        Page.Header.Title = sTitle & "Messaging"
    End Sub

    Protected Sub HideAllPanels()
        pnlTopics.Visible = False
        pnlNewMessage.Visible = False
        pnlAddMessage.Visible = False
        pnlMessages.Visible = False
        pnlFeedback.Visible = False
    End Sub

    Protected Sub ShowTopics()
        ' MessagingTopics.IsClosed takes values: 0=OPEN; 1=CLOSED BY USER; 4=CLOSED BY USER, AWAITING FEEDBACK; 7=CLOSED BY USER, FEEDBACK RECEIVED; 10=CLOSED BY CONTROLLER
        Dim sSQL As String
        sSQL = "SELECT [id], Topic, TopicStatus, TopicReference, AWB, NewMessage, Rating, ISNULL(OnBehalfOf, 0) 'OnBehalfOf', CAST(REPLACE(CONVERT(VARCHAR(11), CreatedOn, 106), ' ', '-') AS varchar(20)) + ' ' + SUBSTRING((CONVERT(VARCHAR(8), CreatedOn, 108)),1,5) 'CreatedOn' FROM MessagingTopics WHERE UserKey = " & Session("UserKey")
        If Not cbIncludeClosedTopics.Checked Then
            sSQL += " AND NOT TopicStatus IN (" & TOPIC_STATUS_CLOSED_BY_USER_FEEDBACK_RECEIVED & ", " & TOPIC_STATUS_CLOSED_BY_CONTROLLER_FEEDBACK_RECEIVED & ", " & TOPIC_STATUS_CLOSED & ")"
        End If
        sSQL += " ORDER BY [id] DESC"
        Dim dtTopics As DataTable
        dtTopics = ExecuteQueryToDataTable(sSQL)
        For Each dr As DataRow In dtTopics.Rows
            If Not IsDBNull(dr("OnBehalfOf")) Then
                If CInt(dr("OnBehalfOf")) > 0 Then
                    dr("Topic") += " (on behalf of " & GetOnBehalfOfUserIDFromKey(dr("OnBehalfOf")) & ")"
                End If
            End If
        Next
        lblNewMessagesToView.Visible = False
        lblRatingRequired.Visible = False
        gvTopics.DataSource = dtTopics
        gvTopics.DataBind()
        Call HideAllPanels()
        pnlTopics.Visible = True
    End Sub
    
    Protected Function GetOnBehalfOfUserIDFromKey(ByVal nUserKey As Int32) As String
        GetOnBehalfOfUserIDFromKey = ExecuteQueryToDataTable("SELECT UserID FROM UserProfile WHERE [key] = " & nUserKey).Rows(0).Item(0)
    End Function
    
    Protected Sub btnNewTopicCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowTopics()
    End Sub

    Protected Sub btnNewTopic_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowNewTopicPanel()
    End Sub
    
    Protected Sub ShowNewTopicPanel()
        Call PopulateTopicDropdown()
        lblLegendTopicType.Text = "New Conversation"
        trNewTopicConsignmentReference.Visible = False
        trOtherTopic.Visible = False
        tbSubject.Text = String.Empty
        tbRaisedBy.Text = String.Empty
        tbNewTopicMessage.Text = String.Empty
        Call HideAllPanels()
        pnlNewMessage.Visible = True
        ddlTopic.Focus()
    End Sub
    
    Protected Sub lnkbtnNewTopicAddConsignmentReference_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call PopulateConsignmentReferenceDropdowns()
        trNewTopicConsignmentReference.Visible = True
        ddlNewTopicConsignmentReference.Focus()
    End Sub

    Protected Sub lnkbtnAddConsignmentReference_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call PopulateConsignmentReferenceDropdowns()
        trAddMessageConsignmentReference.Visible = True
        ddlAddMessageConsignmentReference.Focus()
    End Sub
       
    Protected Sub PopulateConsignmentReferenceDropdowns()
        Dim nUserKey As Int32 = Session("UserKey")
        If pnOnBehalfOfUserKey > 0 Then
            nUserKey = pnOnBehalfOfUserKey
        End If
        Dim oListItemCollection As ListItemCollection = ExecuteQueryToListItemCollection("SELECT TOP 10 AWB,  AWB + ' - ' + CAST(REPLACE(CONVERT(VARCHAR(11), CreatedOn, 106), ' ', '-') AS varchar(20)) 'Consignment' FROM Consignment WHERE CustomerKey = " & Session("CustomerKey") & " AND UserKey = " & nUserKey & " ORDER BY [key] DESC", "Consignment", "AWB")
        ddlNewTopicConsignmentReference.Items.Clear()
        ddlAddMessageConsignmentReference.Items.Clear()
        ddlNewTopicConsignmentReference.Items.Add(New ListItem("- please select -", 0))
        ddlAddMessageConsignmentReference.Items.Add(New ListItem("- please select -", 0))
        For Each li As ListItem In oListItemCollection
            ddlNewTopicConsignmentReference.Items.Add(li)
            ddlAddMessageConsignmentReference.Items.Add(li)
        Next
    End Sub
    
    Protected Sub ddlTopic_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        If ddl.Items(0).Value = "0" Then
            ddl.Items.RemoveAt(0)
        End If
        If ddl.SelectedItem.Text.ToLower.Contains("other") Then
            trOtherTopic.Visible = True
            tbSubject.Focus()
        Else
            trOtherTopic.Visible = False
            tbNewTopicMessage.Focus()
        End If
    End Sub
    
    Protected Sub btnAddMessageCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowTopics()
    End Sub
    
    Protected Sub PopulateTopicDropdown()
        Dim sSQL As String = "SELECT [id], CategoryName, CategoryOrder FROM MessagingTopicCategories WHERE CustomerKey = " & Session("CustomerKey") & " AND IsDeleted = 0 "
        If rbAgentTitles.Checked Then
            sSQL += " AND ISNULL(IsAgentVisible, 0) = 1"
        Else
            sSQL += " AND ISNULL(IsAgentVisible, 0) = 0"
        End If
        sSQL += " ORDER BY CategoryOrder"
        Dim oListItemCollection As ListItemCollection = ExecuteQueryToListItemCollection(sSQL, "CategoryName", "CategoryName")
        ddlTopic.Items.Clear()
        ddlTopic.Items.Add(New ListItem("- please select -", "- please select -"))
        For Each li As ListItem In oListItemCollection
            ddlTopic.Items.Add(li)
        Next
    End Sub
    
    Protected Function IsValidMessage() As Boolean
        IsValidMessage = True
        If trOtherTopic.Visible Then
            If tbSubject.Text.Trim = String.Empty Then
                IsValidMessage = False
                tbSubject.Focus()
                WebMsgBox.Show("Please enter the subject of your new conversation.")
            End If
        End If
        If tbNewTopicMessage.Text.Trim = String.Empty Then
            IsValidMessage = False
            tbNewTopicMessage.Focus()
            WebMsgBox.Show("Please enter your message.")
        End If
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
            'Err.Raise(1001, "ExecuteQueryToDataTable", ex.Message)
            WebMsgBox.Show("Error in ExecuteQueryToDataTable executing: " & sQuery & " : " & ex.Message)
        Finally
            oConn.Close()
        End Try
        ExecuteQueryToDataTable = oDataTable
    End Function

    Protected Sub btnBackToTopicsFromMessages_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowTopics()
    End Sub
    
#End Region
    
    Protected Sub cbIncludeClosedTopics_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowTopics()
    End Sub
    
    Protected Function sGetAuthor(ByVal DataItem As Object) As String
        Dim nCreatedBy As Int32 = DataBinder.Eval(DataItem, "CreatedBy")
        If nCreatedBy = Session("UserKey") Then
            sGetAuthor = TITLE_AGENT
        ElseIf nCreatedBy = 0 Then
            sGetAuthor = TITLE_SYSTEM
        Else
            sGetAuthor = TITLE_CONTROLLER
        End If
    End Function

    Protected Function sGetRatingVisibility(ByVal DataItem As Object) As Boolean
        Dim nTopicStatus As Int32 = CInt(DataBinder.Eval(DataItem, "TopicStatus"))
        Select Case nTopicStatus
            Case TOPIC_STATUS_OPEN
                sGetRatingVisibility = False
            Case TOPIC_STATUS_CLOSED_BY_USER
                sGetRatingVisibility = True
            Case TOPIC_STATUS_CLOSED_BY_USER_AWAITING_FEEDBACK
                sGetRatingVisibility = True
            Case TOPIC_STATUS_CLOSED_BY_USER_FEEDBACK_RECEIVED
                sGetRatingVisibility = True
                'Case TOPIC_STATUS_CLOSED_BY_CONTROLLER
                '   sGetRatingVisibility = True
            Case TOPIC_STATUS_CLOSED_BY_CONTROLLER_AWAITING_FEEDBACK
                sGetRatingVisibility = True
            Case TOPIC_STATUS_CLOSED_BY_CONTROLLER_FEEDBACK_RECEIVED
                sGetRatingVisibility = True
            Case Else
                sGetRatingVisibility = False
        End Select
    End Function

    ' MessagingTopics.IsClosed takes values: 0=OPEN; 1=CLOSED BY USER; 4=CLOSED BY USER, AWAITING FEEDBACK; 7=CLOSED BY USER, FEEDBACK RECEIVED; 10=CLOSED BY CONTROLLER
    Protected Function sGetTopicStatus(ByVal DataItem As Object) As String
        Dim nTopicStatus As Int32 = CInt(DataBinder.Eval(DataItem, "TopicStatus"))
        Select Case nTopicStatus
            Case TOPIC_STATUS_OPEN
                sGetTopicStatus = "OPEN"
            Case TOPIC_STATUS_CLOSED_BY_USER
                sGetTopicStatus = "AGENT REQUESTED CLOSE"
            Case TOPIC_STATUS_CLOSED_BY_USER_AWAITING_FEEDBACK
                sGetTopicStatus = "AGENT REQUESTED CLOSE, AWAITING FEEDBACK"
            Case TOPIC_STATUS_CLOSED_BY_USER_FEEDBACK_RECEIVED
                sGetTopicStatus = "CLOSED BY AGENT"
                'Case TOPIC_STATUS_CLOSED_BY_CONTROLLER
                '   sGetTopicStatus = "CONTROLLER IS CLOSING"
            Case TOPIC_STATUS_CLOSED_BY_CONTROLLER_AWAITING_FEEDBACK
                sGetTopicStatus = "CONTROLLER IS CLOSING, AWAITING AGENT FEEDBACK"
            Case TOPIC_STATUS_CLOSED_BY_CONTROLLER_FEEDBACK_RECEIVED
                sGetTopicStatus = "CLOSED BY CONTROLLER"
            Case Else
                sGetTopicStatus = "INDETERMINATE STATUS - PLEASE CONTACT CONTROLLER"
        End Select
    End Function
    
    Protected Sub btnShowMessagesFromTopicList_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim btn As Button = sender
        pnTopicID = btn.CommandArgument
        Call ShowMessages(pnTopicID)
    End Sub

    Protected Sub ShowMessages(ByVal sTopicReference As String)
        lblMessagesForTopicInfo.Text = GetTopicTitle(pnTopicID)
        Dim sSQL As String = "UPDATE MessagingTopics SET NewMessage = 0 WHERE [id] = " & sTopicReference & " SELECT CAST(REPLACE(CONVERT(VARCHAR(11), CreatedOn, 106), ' ', '-') AS varchar(20)) + ' ' + SUBSTRING((CONVERT(VARCHAR(8), CreatedOn, 108)),1,5) 'CreatedOn', MessageRef, MessageBody, CreatedBy FROM MessagingMessages WHERE ISNULL(IsAdmin, 0) = 0 AND TopicNumber = " & sTopicReference & " ORDER BY [id]"
        Dim oDT As DataTable = ExecuteQueryToDataTable(sSQL)
        gvMessages.DataSource = oDT
        gvMessages.DataBind()
        sSQL = "SELECT AWB FROM MessagingTopics WHERE [id] = " & pnTopicID
        oDT = ExecuteQueryToDataTable(sSQL)
        lblMessagesConsignmentNumber.Text = oDT.Rows(0).Item(0).ToString.Trim
        If lblMessagesConsignmentNumber.Text <> String.Empty Then
            lblMessagesConsignmentNumber.Text = "(Consignment: " & lblMessagesConsignmentNumber.Text & ")"
        End If
        Call HideAllPanels()
        pnlMessages.Visible = True
    End Sub
    
    Protected Sub imgbtnShowMessages_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        ' MessagingTopics.IsClosed takes values: 0=OPEN; 1=CLOSED BY USER; 4=CLOSED BY USER, AWAITING FEEDBACK; 7=CLOSED BY USER, FEEDBACK RECEIVED; 10=CLOSED BY CONTROLLER; 20=CLOSED
        Dim imgbtn As ImageButton = sender
        pnTopicID = imgbtn.CommandArgument
        btnAddMessageFromMessages.Visible = bShowAddMessage()
        lblMessagesForTopicInfo.Text = GetTopicTitle(pnTopicID)
        Call ShowMessages(pnTopicID)
    End Sub

    Protected Function bShowAddMessage() As Boolean
        bShowAddMessage = True
        Dim sSQL As String
        sSQL = "SELECT TopicStatus FROM MessagingTopics WHERE [id] = " & pnTopicID
        Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
        Dim nTopicStatus As Int32 = dt.Rows(0).Item(0)
        If nTopicStatus >= TOPIC_STATUS_CLOSED_BY_USER_FEEDBACK_RECEIVED Then
            bShowAddMessage = False
        End If
    End Function
    
    Protected Sub btnAddMessage_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim btn As Button = sender
        If btn.CommandArgument <> String.Empty Then
            pnTopicID = btn.CommandArgument
        End If
        lblTopicInfo.Text = GetTopicTitle(pnTopicID)
        lblAddMessageConsignmentNumber.Text = GetTopicConsignmentRef(pnTopicID)
        trAddMessageConsignmentReference.Visible = False
        If lblAddMessageConsignmentNumber.Text = String.Empty Then
            lnkbtnAddMessageAddConsignmentReference.Visible = True
        Else
            lnkbtnAddMessageAddConsignmentReference.Visible = False
            lblAddMessageConsignmentNumber.Text = "Consignment Ref: " & lblAddMessageConsignmentNumber.Text
        End If
        Call HideAllPanels()
        tbAddMessage.Text = String.Empty
        pnlAddMessage.Visible = True
        tbAddMessage.Focus()
    End Sub

    Protected Sub imgbtnAddMessage_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Dim imgbtn As ImageButton = sender
        If imgbtn.CommandArgument <> String.Empty Then
            pnTopicID = imgbtn.CommandArgument
        End If
        lblTopicInfo.Text = GetTopicTitle(pnTopicID)
        lblAddMessageConsignmentNumber.Text = GetTopicConsignmentRef(pnTopicID)
        trAddMessageConsignmentReference.Visible = False
        If lblAddMessageConsignmentNumber.Text = String.Empty Then
            lnkbtnAddMessageAddConsignmentReference.Visible = True
        Else
            lnkbtnAddMessageAddConsignmentReference.Visible = False
            lblAddMessageConsignmentNumber.Text = "Consignment Ref: " & lblAddMessageConsignmentNumber.Text
        End If
        Call HideAllPanels()
        tbAddMessage.Text = String.Empty
        pnlAddMessage.Visible = True
        tbAddMessage.Focus()
    End Sub

    Protected Sub btnCloseTopic_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim btn As Button = sender
        Dim nTopicID As String = btn.CommandArgument
        Call CloseTopic(nTopicID)
        Call ShowTopics()
    End Sub
    
    Protected Sub imgbtnClose_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Dim imgbtn As ImageButton = sender
        Dim nTopicID As String = imgbtn.CommandArgument
        Call CloseTopic(nTopicID)
        Call ShowTopics()
    End Sub
    
    Protected Sub gvTopics_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs)
        Dim gvr As GridViewRow = e.Row
        If gvr.RowType = DataControlRowType.DataRow Then
            Dim hidID As HiddenField = gvr.Cells(0).FindControl("hidID")
            Dim lblTopicStatus As Label = gvr.Cells(TOPICGRID_STATUS).FindControl("lblTopicStatus")
            Dim rblRating As RadioButtonList = gvr.Cells(TOPICGRID_RATING).FindControl("rblRating")
            If lblTopicStatus.Text.ToLower.Contains("closed") Then
                Dim imgbtn As ImageButton
                imgbtn = gvr.Cells(TOPICGRID_ICONS).FindControl("imgbtnAddMessage")
                imgbtn.Visible = False
                imgbtn = gvr.Cells(TOPICGRID_ICONS).FindControl("imgbtnClose")
                imgbtn.Visible = False
                rblRating.Enabled = False
            ElseIf lblTopicStatus.Text.ToLower.Contains("open") Then
                rblRating.Visible = False
            ElseIf lblTopicStatus.Text.ToLower.Contains("requested") Then
                gvr.Cells(TOPICGRID_RATING).BackColor = COLOUR_TOPICSTATUS_CLOSING
                lblRatingRequired.Visible = True
            End If

            Dim sSQL As String = "SELECT Rating, NewMessage FROM MessagingTopics WHERE ID = " & hidID.Value
            Dim nRating As Int32 = 0
            Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
            If Not IsDBNull(dt.Rows(0).Item("Rating")) Then
                nRating = dt.Rows(0).Item("Rating")
            End If
            If nRating > 0 And nRating < 2 Then
                rblRating.SelectedValue = nRating
            End If
            Dim bNewMessage As Boolean = dt.Rows(0).Item("NewMessage")
            If bNewMessage Then
                gvr.Cells(TOPICGRID_TOPIC).BackColor = COLOUR_NEWMESSAGE
                lblNewMessagesToView.Visible = True
            End If
        End If
    End Sub
    
    Protected Sub btnFeedbackSend_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Page.Validate()
        If Page.IsValid Then
            Dim sMessageTimestamp As String = Format(DateTime.Now, "yyyyMMddhhmmss")
            tbFeedback.Text = tbFeedback.Text.Trim
            If tbFeedback.Text <> String.Empty Then
                Call SaveMessage(pnTopicID, GetTopicReferenceFromID(pnTopicID) & "." & sMessageTimestamp, "USER FEEDBACK: " & tbFeedback.Text, NoNewMessage:=True)
            End If
            Call FeedbackReceivedEmailAlert()
            Dim sSQL As String
            sSQL = "SELECT TopicStatus FROM MessagingTopics WHERE [id] = " & pnTopicID
            Dim nTopicStatus As Int32 = ExecuteQueryToDataTable(sSQL).Rows(0).Item(0)
            If nTopicStatus = TOPIC_STATUS_CLOSED_BY_CONTROLLER_AWAITING_FEEDBACK Then
                sSQL = "UPDATE MessagingTopics SET TopicStatus = " & TOPIC_STATUS_CLOSED_BY_CONTROLLER_FEEDBACK_RECEIVED & ", ClosedBy = '" & tbClosedBy.Text.Replace("'", "''") & "' WHERE [id] = " & pnTopicID
            ElseIf nTopicStatus = TOPIC_STATUS_CLOSED_BY_USER Or nTopicStatus = TOPIC_STATUS_CLOSED_BY_USER_AWAITING_FEEDBACK Then
                sSQL = "UPDATE MessagingTopics SET TopicStatus = " & TOPIC_STATUS_CLOSED_BY_USER_FEEDBACK_RECEIVED & ", ClosedBy = '" & tbClosedBy.Text.Replace("'", "''") & "' WHERE [id] = " & pnTopicID
            End If
            Call ExecuteQueryToDataTable(sSQL)
            Call ShowTopics()
        End If
    End Sub

    Protected Sub CloseTopic(ByVal nTopicID As Int32)
        Dim sSQL As String = "UPDATE MessagingTopics SET TopicStatus = " & TOPIC_STATUS_CLOSED_BY_USER & " WHERE [id] = " & nTopicID
        Call ExecuteQueryToDataTable(sSQL)
    End Sub

    Protected Sub gvMessages_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs)
        Dim gvr As GridViewRow = e.Row
        If gvr.RowType = DataControlRowType.DataRow Then
            Dim lblAuthor As Label = gvr.Cells(MESSAGEGRID_AUTHOR).Controls(1).FindControl("lblAuthor")
            If lblAuthor.Text = TITLE_AGENT Then
                gvr.BackColor = COLOUR_MESSAGEAUTHOR_AGENT
            Else
                gvr.BackColor = COLOUR_MESSAGEAUTHOR_CONTROLLER
            End If
        End If
    End Sub

    Protected Sub rblRating_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim rbl As RadioButtonList = sender
        Dim hidID As HiddenField = rbl.NamingContainer.FindControl("hidID")
        Dim nID As Integer = Convert.ToInt32(hidID.Value)
        pnTopicID = nID
        Dim nRating As Int32 = 1
        If rbl.SelectedValue.ToLower = "satisfactory" Then
            nRating = 2
        End If
        Dim sSQL As String = "UPDATE MessagingTopics SET Rating = " & nRating & " WHERE ID = " & nID
        Call ExecuteQueryToDataTable(sSQL)
        Call ShowFeedbackPanel()
        tbClosedBy.Focus()
    End Sub
    
    Protected Sub ShowFeedbackPanel()
        Call HideAllPanels()
        pnlFeedback.Visible = True
        tbClosedBy.Text = String.Empty
        tbFeedback.Text = String.Empty
    End Sub

    Protected Sub SendNewTopicEmailAlertToController()
        Dim sSQL As String = "SELECT * FROM MessagingConfiguration WHERE CustomerKey = " & Session("CustomerKey")
        Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
        Try
            If dt.Rows(0).Item("NewTopicEmail") <> 0 Then
                Dim sAddressees() As String = dt.Rows(0).Item("EmailAddressListCommaSeparated").Split(",")
                For Each sAddressee As String In sAddressees
                    If sAddressee.Trim <> String.Empty Then
                        Call SendMail("MSGING_NEWTOPICALERT", sAddressee, "Messaging: New Topic Alert", "New topic posted by user " & Session(""), "New topic posted by user " & Session("UserName"))
                    End If
                Next
            End If
        Catch
        End Try
    End Sub
    
    Protected Sub SendNewMessageEmailAlert()
        Dim sSQL As String = "SELECT * FROM MessagingConfiguration WHERE CustomerKey = " & Session("CustomerKey")
        Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
        Try
            If dt.Rows(0).Item("NewMessageEmail") <> 0 Then
                Dim sAddressees() As String = dt.Rows(0).Item("EmailAddressListCommaSeparated").Split(",")
                For Each sAddressee As String In sAddressees
                    If sAddressee.Trim <> String.Empty Then
                        Call SendMail("MSGING_NEWMSGALERT", sAddressee, "Messaging: New Message Alert", "New message posted by user " & Session(""), "New message posted by user " & Session("UserName"))
                    End If
                Next
            End If
        Catch
        End Try
    End Sub

    Protected Sub FeedbackReceivedEmailAlert()
        Dim sSQL As String = "SELECT * FROM MessagingConfiguration WHERE CustomerKey = " & Session("CustomerKey")
        Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
        Try
            If dt.Rows(0).Item("FeedbackReceivedEmail") <> 0 Then
                Dim sAddressees() As String = dt.Rows(0).Item("EmailAddressListCommaSeparated").Split(",")
                For Each sAddressee As String In sAddressees
                    If sAddressee.Trim <> String.Empty Then
                        Call SendMail("MSGING_FEEDBACKALERT", sAddressee, "Messaging: Feedback Received Alert", "Feedback received from user " & Session(""), "Feedback received from user " & Session("UserName"))
                    End If
                Next
            End If
        Catch
        End Try
    End Sub

    Protected Sub SendMail(ByVal sType As String, ByVal sRecipient As String, ByVal sSubject As String, ByVal sBodyText As String, ByVal sBodyHTML As String)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Email_AddToQueue", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Try
            oCmd.Parameters.Add(New SqlParameter("@MessageId", SqlDbType.NVarChar, 20))
            oCmd.Parameters("@MessageId").Value = sType
    
            oCmd.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
            oCmd.Parameters("@CustomerKey").Value = Session("CustomerKey")
    
            oCmd.Parameters.Add(New SqlParameter("@StockBookingKey", SqlDbType.Int))
            oCmd.Parameters("@StockBookingKey").Value = 0
    
            oCmd.Parameters.Add(New SqlParameter("@ConsignmentKey", SqlDbType.Int))
            oCmd.Parameters("@ConsignmentKey").Value = 0
    
            oCmd.Parameters.Add(New SqlParameter("@ProductKey", SqlDbType.Int))
            oCmd.Parameters("@ProductKey").Value = 0
    
            oCmd.Parameters.Add(New SqlParameter("@To", SqlDbType.NVarChar, 100))
            oCmd.Parameters("@To").Value = sRecipient
    
            oCmd.Parameters.Add(New SqlParameter("@Subject", SqlDbType.NVarChar, 60))
            oCmd.Parameters("@Subject").Value = sSubject
    
            oCmd.Parameters.Add(New SqlParameter("@BodyText", SqlDbType.NText))
            oCmd.Parameters("@BodyText").Value = sBodyText
    
            oCmd.Parameters.Add(New SqlParameter("@BodyHTML", SqlDbType.NText))
            oCmd.Parameters("@BodyHTML").Value = sBodyHTML
    
            oCmd.Parameters.Add(New SqlParameter("@QueuedBy", SqlDbType.Int))
            oCmd.Parameters("@QueuedBy").Value = Session("UserKey")
    
            oConn.Open()
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            WebMsgBox.Show("Error in SendMail: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub

    Property pnTopicID() As Integer
        Get
            Dim o As Object = ViewState("ME_TopicID")
            If o Is Nothing Then
                Return 0
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Integer)
            ViewState("ME_TopicID") = Value
        End Set
    End Property

    Property pnOnBehalfOfUserKey() As Integer
        Get
            Dim o As Object = ViewState("ME_OnBehalfOfUserKey")
            If o Is Nothing Then
                Return 0
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Integer)
            ViewState("ME_OnBehalfOfUserKey") = Value
        End Set
    End Property

    Protected Sub rcbUser_SelectedIndexChanged(ByVal o As Object, ByVal e As Telerik.Web.UI.RadComboBoxSelectedIndexChangedEventArgs)
        Dim rcb As RadComboBox = o
        If IsNumeric(rcb.SelectedValue) Then
            pnOnBehalfOfUserKey = rcb.SelectedValue
        Else
            pnOnBehalfOfUserKey = 0
        End If
        Call PopulateConsignmentReferenceDropdowns()
        Call ExecuteQueryToDataTable("UPDATE MessagingTopics SET AWB = '' WHERE [id] = " & pnTopicID)
    End Sub

    Protected Sub rcbUser_ItemsRequested(ByVal o As Object, ByVal e As Telerik.Web.UI.RadComboBoxItemsRequestedEventArgs)
        Dim data As DataTable = GetUsers(e.Text)
        Dim itemOffset As Integer = e.NumberOfItems
        Dim endOffset As Integer = Math.Min(itemOffset + ITEMS_PER_REQUEST, data.Rows.Count)
        e.EndOfItems = endOffset = data.Rows.Count
        rcbUser.DataTextField = "UserName"
        rcbUser.DataValueField = "UserKey"
        For i As Int32 = itemOffset To endOffset - 1
            Dim rcbi As New RadComboBoxItem
            rcbi.Text = data.Rows(i)("UserName").ToString()
            rcbi.Value = data.Rows(i)("UserKey").ToString()
            rcbUser.Items.Add(rcbi)
        Next
    End Sub

    Protected Function GetUsers(Optional ByVal sFilter As String = "") As DataTable
        GetUsers = Nothing
        Dim sSQL As String = "SELECT [key] 'UserKey',  UserId + ' - ' + FirstName + ' ' + LastName 'UserName' FROM UserProfile WHERE [Status] = 'Active' AND Type = 'User' AND CustomerKey = " & Session("CustomerKey")
        If sFilter <> String.Empty Then
            sFilter = sFilter.Replace("'", "''")
            sSQL += " AND UserId LIKE '%" & sFilter & "%'"
        End If
        sSQL += " ORDER BY UserId"
        GetUsers = ExecuteQueryToDataTable(sSQL)
    End Function
    
    'Protected Function GetProductsByCustomer(ByVal sCustomerKey As String, Optional ByVal sFilter As String = "") As DataTable            ' XXXX
    '    Dim sSQL As String
    '    sSQL = "SELECT ProductCode + ' ' + ISNULL(ProductDate,'') + ' ' + ProductDescription 'Product', LogisticProductKey, ThumbnailImage FROM LogisticProduct WHERE ArchiveFlag = 'N' AND DeletedFlag = 'N' AND CustomerKey = " & sCustomerKey
    '    If sFilter <> String.Empty Then
    '        sFilter = sFilter.Replace("'", "''")
    '        sSQL += " AND (ProductCode LIKE '%" & sFilter & "%' OR ProductDescription LIKE '%" & sFilter & "%')"
    '    End If
    '    ' tbSearch.Text = tbSearch.Text.Trim
    '    'If psSearchString <> String.Empty Then
    '    '    sSQL += " AND (ProductCode LIKE '%" & psSearchString & "%' OR ProductDescription LIKE '%" & psSearchString & "%')"
    '    'End If
    '    If rbFavouriteProducts.Checked Then
    '        sSQL += " AND (LogisticProductKey IN (SELECT ProductKey FROM UserProductFavouritesDefaults WHERE CustomerKey = " & pnImpersonateCustomer & ")"
    '        sSQL += " OR LogisticProductKey IN (SELECT ProductKey FROM UserProductFavourites WHERE UserKey = " & pnImpersonateBookedByUser & "))"
    '    End If
    '    sSQL += " ORDER BY ProductCode"
    '    Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
    '    GetProductsByCustomer = dt
    'End Function
    
    Protected Sub lnkbtnClearOnBehalfOf_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        rcbUser.Text = String.Empty
    End Sub
    
    Protected Sub rbAgentTitles_CheckedChanged(sender As Object, e As System.EventArgs)
        trOtherTopic.Visible = False
        tbSubject.Text = String.Empty
        Call PopulateTopicDropdown()
    End Sub

    Protected Sub rbWUInternalTitles_CheckedChanged(sender As Object, e As System.EventArgs)
        trOtherTopic.Visible = False
        tbSubject.Text = String.Empty
        Call PopulateTopicDropdown()
    End Sub
    
</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title></title>
    <style type="text/css">
        .art-blockcontent-body .SubHead, .art-blockcontent-body .Normal
        {
            color: #2E3D4C;
            font-family: Arial, Helvetica, Sans-Serif;
            font-size: 12px;
        }
        
        .Normal, .normal, #LoginInfo, #QuickLinks, #LoginInfo p.LoginNotes, .art-postcontent .SubHead, .art-postcontent .Normal, .SubHead, .WizardText, .SkinObject
        {
            font-size: 1em;
            color: #0F1419;
        }
        
        .Normal, .normal, #LoginInfo, #QuickLinks, #LoginInfo p.LoginNotes, .SubHead, .WizardText, .SkinObject
        {
            font-family: Arial, Helvetica, Sans-Serif;
            font-style: normal;
            font-weight: normal;
            font-size: 13px;
        }
        
        .Normal, .NormalDisabled, .NormalDeleted
        {
            font-size: 11px;
            font-weight: normal;
        }
    </style>
</head>
<body style="font-size: 08pt; font-family: Verdana">
    <form id="Form1" runat="Server">
    <main:Header ID="ctlHeader" runat="server" />
    <table style="width: 100%" cellpadding="0" cellspacing="0">
        <tr class="bar_reports">
            <td style="width: 50%; white-space: nowrap">
            </td>
            <td style="width: 50%; white-space: nowrap" align="right">
            </td>
        </tr>
    </table>
    <asp:Panel ID="pnlFAQ" runat="server" Width="100%">
        <table style="width: 100%">
            <tr>
                <td style="width: 1%">
                </td>
                <td style="width: 11%">
                    &nbsp;
                </td>
                <td style="width: 32%">
                    &nbsp;
                </td>
                <td style="width: 23%">
                </td>
                <td style="width: 32%">
                </td>
                <td style="width: 1%">
                </td>
            </tr>
            <tr>
                <td style="width: 1%">
                    &nbsp;
                </td>
                <td colspan="4">
                    <strong>
                    <asp:Label ID="lblLegendTopicType0" runat="server" Font-Bold="True" 
                        Font-Names="Verdana" Font-Size="X-Small" ForeColor="Navy" 
                        Text="Frequently Asked Questions" />
                    </strong>&nbsp;<table style="width: 100%">
                        <tr>
                            <td style="width: 30%">
                            </td>
                            <td style="width: 70%">
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <asp:Label ID="Label12" runat="server" Font-Bold="True" Font-Italic="True" 
                                    ForeColor="#0070C0" Style="font-size: 08pt" 
                                    Text="· How Do I Change My Address?" />
                            </td>
                            <td>
                                <asp:Label ID="Label11" runat="server" Style="font-size: 08pt" 
                                    Text="Use the Agent Support tab to start a new Conversation. Select category “Other” and enter “Change Address” as your title. In the message field enter in your new address details. Send the message and we will respond" />
                            </td>
                        </tr>
                        <tr>
                            <td/>
                            <td/>
                        </tr>
                        <tr>
                            <td>
                                <asp:Label ID="Label13" runat="server" Font-Bold="True" Font-Italic="True" 
                                    ForeColor="#0070C0" Style="font-size: 08pt" 
                                    Text="· Can I send my package to another address? " />
                            </td>
                            <td>
                                <asp:Label ID="Label18" runat="server" Style="font-size: 08pt" 
                                    Text="Unfortunately due to security restrictions you can only receive items at your registered address." />
                            </td>
                        </tr>
                        <tr>
                            <td/>
                            <td/>
                        </tr>
                        <tr>
                            <td>
                                <asp:Label ID="Label26" runat="server" Font-Bold="True" Font-Italic="True" 
                                    ForeColor="#0070C0" Style="font-size: 08pt" 
                                    Text="· How do I track my consignments? " />
                            </td>
                            <td>
                                <asp:Label ID="Label22" runat="server" Style="font-size: 08pt" 
                                    Text="You can track all your consignments under the “Track &amp; Trace” tab. Simply enter the consignment number in the search bar and click the go button." />
                            </td>
                        </tr>
                        <tr>
                            <td/>
                            <td/>
                        </tr>
                        <tr>
                            <td>
                                <asp:Label ID="Label23" runat="server" Font-Bold="True" Font-Italic="True" 
                                    ForeColor="#0070C0" Style="font-size: 08pt" 
                                    Text="· I haven’t received the correct consignment." />
                            </td>
                            <td>
                                <asp:Label ID="Label19" runat="server" Style="font-size: 08pt" 
                                    Text="Use the Agent Support tab to start a new Conversation. Select category “Wrong product(s) or quantity sent”. Add the associated consignment reference and describe the problem in the Message box. Send the message and we will respond." />
                            </td>
                        </tr>
                        <tr>
                            <td/>
                            <td/>
                        </tr>
                        <tr>
                            <td>
                                <asp:Label ID="Label24" runat="server" Font-Bold="True" Font-Italic="True" 
                                    ForeColor="#0070C0" Style="font-size: 08pt" 
                                    Text="· How do I change my password?" />
                            </td>
                            <td>
                                <asp:Label ID="Label20" runat="server" Style="font-size: 08pt" 
                                    Text="On the top right hand side of the web site click the “Chng pwd” link. The next time you log in the system will ask you to change your password." />
                            </td>
                        </tr>
                        <tr>
                            <td/>
                            <td/>
                        </tr>
                        <tr>
                            <td>
                                <asp:Label ID="Label25" runat="server" Font-Bold="True" Font-Italic="True" 
                                    ForeColor="#0070C0" Style="font-size: 08pt" 
                                    Text="· I want to order an item that doesn't seem to be available" />
                            </td>
                            <td>
                                <asp:Label ID="Label21" runat="server" Style="font-size: 08pt" 
                                    Text="Use the Agent Support tab to start a new Conversation. Select category “Problem placing order”. In the Message box describe your problem. &nbsp;Send the message and we will respond." />
                            </td>
                        </tr>
                    </table>
                    <br />
                </td>
                <td style="width: 1%">
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td style="width: 1%">
                    &nbsp;
                </td>
                <td colspan="4">
                    To contact us for any other issue, please click on the <b>contact us</b> button.
                </td>
                <td style="width: 1%">
                    &nbsp;
                </td>
            </tr>
        </table>
        <br />
    </asp:Panel>
    <asp:Panel ID="pnlNewMessage" runat="server" Width="100%">
        <table style="width: 100%">
            <tr>
                <td style="width: 1%">
                </td>
                <td style="width: 11%">
                    <asp:Label ID="lblLegendTopicType" runat="server" Font-Bold="True" Font-Names="Verdana"
                        Font-Size="X-Small" ForeColor="Navy" Text="New Conversation" />
                </td>
                <td style="width: 32%">
                    <asp:RadioButton ID="rbAgentTitles" runat="server" Style="font-size: 08pt" GroupName="AgentOrInternal"
                        Checked="True" Text="Agent titles" AutoPostBack="True" OnCheckedChanged="rbAgentTitles_CheckedChanged" />
                    <asp:RadioButton ID="rbWUInternalTitles" runat="server" Style="font-size: 08pt" GroupName="AgentOrInternal"
                        Text="WU internal titles" AutoPostBack="True" OnCheckedChanged="rbWUInternalTitles_CheckedChanged" />
                </td>
                <td style="width: 23%">
                </td>
                <td style="width: 32%">
                </td>
                <td style="width: 1%">
                </td>
            </tr>
            <tr>
                <td>
                    &nbsp;
                </td>
                <td align="right">
                    <asp:Label ID="Label1" runat="server" Style="font-size: 08pt" Text="Title:" />
                </td>
                <td colspan="3">
                    <asp:DropDownList ID="ddlTopic" runat="server" Style="font-size: 08pt" ValidationGroup="newtopic"
                        OnSelectedIndexChanged="ddlTopic_SelectedIndexChanged" AutoPostBack="true" />
                    &nbsp;<asp:RequiredFieldValidator ID="rfvTopic" runat="server" ErrorMessage="## please select a title"
                        Display="Dynamic" ForeColor="Red" InitialValue="- please select -" ControlToValidate="ddlTopic"
                        Font-Bold="True" Style="font-size: 08pt" SetFocusOnError="True" ValidationGroup="newtopic" />
                    &nbsp;<asp:LinkButton ID="lnkbtnNewTopicAddConsignmentReference" runat="server" Font-Names="Arial"
                        Font-Size="XX-Small" OnClick="lnkbtnNewTopicAddConsignmentReference_Click">add consignment reference</asp:LinkButton>
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr id="trOtherTopic" runat="server" visible="false">
                <td>
                    &nbsp;
                </td>
                <td align="right">
                    <asp:Label ID="Label2" runat="server" Style="font-size: 08pt" Text="Other title:" />
                </td>
                <td colspan="3">
                    <asp:TextBox ID="tbSubject" runat="server" Style="font-size: 08pt" Width="50%" MaxLength="50" />
                    &nbsp;<asp:RequiredFieldValidator ID="rfvSubject" runat="server" ControlToValidate="tbSubject"
                        ErrorMessage="## please enter a title for your conversation" Font-Bold="True"
                        Style="font-size: 08pt" ForeColor="Red" InitialValue="" SetFocusOnError="True"
                        ValidationGroup="newtopic" />
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr id="trOnBehalfOf" runat="server" visible="false">
                <td>
                    &nbsp;
                </td>
                <td align="right">
                    <asp:Label ID="Label7" runat="server" Style="font-size: 08pt" Text="On behalf of:" />
                </td>
                <td colspan="3">
                    <telerik:RadComboBox ID="rcbUser" runat="server" AutoPostBack="True" CausesValidation="False"
                        EnableLoadOnDemand="True" EnableVirtualScrolling="True" Filter="Contains" Style="font-size: 08pt"
                        HighlightTemplatedItems="true" OnItemsRequested="rcbUser_ItemsRequested" OnSelectedIndexChanged="rcbUser_SelectedIndexChanged"
                        Width="300px" ShowMoreResultsBox="True" ToolTip="Shows all users when no search text is specified. Search for users by typing an agent code or name." />
                    &nbsp;<asp:LinkButton ID="lnkbtnClearOnBehalfOf" runat="server" Style="font-size: 08pt"
                        OnClick="lnkbtnClearOnBehalfOf_Click">clear filter</asp:LinkButton>
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr id="trRaisedBy" runat="server" visible="false">
                <td>
                    &nbsp;
                </td>
                <td align="right">
                    <asp:Label ID="Label10" runat="server" Style="font-size: 08pt" Text="Raised by:" />
                </td>
                <td colspan="3">
                    <asp:TextBox ID="tbRaisedBy" runat="server" Style="font-size: 08pt" MaxLength="100"
                        Width="50%" />
                    &nbsp;<asp:RequiredFieldValidator ID="rfvRaisedBy" runat="server" ControlToValidate="tbRaisedBy"
                        ErrorMessage="## please indicate who is starting this conversation eg 'John Smith'"
                        Font-Bold="True" Style="font-size: 08pt" ForeColor="Red" InitialValue="" SetFocusOnError="True"
                        ValidationGroup="newtopic" />
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr id="trNewTopicConsignmentReference" runat="server" visible="false">
                <td>
                    &nbsp;
                </td>
                <td align="right">
                    <asp:Label ID="Label4" runat="server" Style="font-size: 08pt" Text="Consignment ref:" />
                </td>
                <td colspan="3">
                    <asp:DropDownList ID="ddlNewTopicConsignmentReference" runat="server" Style="font-size: 08pt" />
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td>
                    &nbsp;
                </td>
                <td align="right">
                    <asp:Label ID="Label3" runat="server" Style="font-size: 08pt" Text="Message:" />
                </td>
                <td colspan="3">
                    <asp:TextBox ID="tbNewTopicMessage" runat="server" Style="font-size: 08pt" Rows="6"
                        TextMode="MultiLine" Width="100%" MaxLength="3000" />
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td>
                    &nbsp;
                </td>
                <td align="right">
                    &nbsp;
                </td>
                <td colspan="3">
                    <asp:Button ID="btnNewTopicSend" runat="server" Text="send" Width="120px" CausesValidation="true"
                        ValidationGroup="newtopic" OnClick="btnNewTopicSend_Click" />
                    &nbsp;<asp:Button ID="btnNewTopicCancel" runat="server" Text="cancel" OnClick="btnNewTopicCancel_Click" />
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
        </table>
    </asp:Panel>
    <asp:Panel ID="pnlAddMessage" runat="server" Width="100%">
        <table style="width: 100%">
            <tr>
                <td style="width: 1%">
                </td>
                <td style="width: 11%">
                    <strong style="color: navy; font-size: x-small; font-family: Verdana">
                        <asp:Label ID="lblLegendAddMessage" runat="server" Font-Bold="True" Font-Names="Verdana"
                            Font-Size="X-Small" ForeColor="Navy" Text="Add Message" />
                    </strong>
                </td>
                <td style="width: 32%">
                </td>
                <td style="width: 23%">
                </td>
                <td style="width: 32%" align="right">
                    <asp:CheckBox ID="cbCloseTopic" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="Close this conversation" />
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
                    <asp:Label ID="lblTopicInfo" runat="server" Font-Bold="True" Style="font-size: 08pt" />
                </td>
                <td colspan="2">
                    <asp:LinkButton ID="lnkbtnAddMessageAddConsignmentReference" runat="server" Font-Names="Arial"
                        Font-Size="XX-Small" OnClick="lnkbtnAddConsignmentReference_Click">add consignment reference</asp:LinkButton>
                    &nbsp;<asp:Label ID="lblAddMessageConsignmentNumber" runat="server" Font-Names="Verdana"
                        Font-Size="XX-Small" Text="" />
                </td>
                <td>
                </td>
            </tr>
            <tr id="trAddMessageConsignmentReference" runat="server" visible="false">
                <td>
                    &nbsp;
                </td>
                <td align="right">
                    <asp:Label ID="Label4a" runat="server" Style="font-size: 08pt" Text="Consignment ref:" />
                </td>
                <td>
                    <asp:DropDownList ID="ddlAddMessageConsignmentReference" runat="server" Style="font-size: 08pt" />
                </td>
                <td>
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td>
                    &nbsp;
                </td>
                <td align="right">
                    <asp:Label ID="Label6" runat="server" Style="font-size: 08pt" Text="Message:" />
                </td>
                <td colspan="3">
                    <asp:TextBox ID="tbAddMessage" runat="server" Style="font-size: 08pt" Rows="6" TextMode="MultiLine"
                        Width="100%" MaxLength="3000"></asp:TextBox>
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td>
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
                <td>
                    <asp:Button ID="btnAddMessageSend" runat="server" OnClick="btnAddMessageSend_Click"
                        Text="send" Width="120px" />
                    &nbsp;<asp:Button ID="btnAddMessageCancel" runat="server" OnClick="btnAddMessageCancel_Click"
                        Text="cancel" />
                </td>
                <td>
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
        </table>
    </asp:Panel>
    <asp:Panel ID="pnlFeedback" runat="server" Width="100%">
        <table style="width: 100%">
            <tr>
                <td style="width: 1%">
                </td>
                <td style="width: 11%">
                    <asp:Label ID="lblLegendFeedback" runat="server" Font-Bold="True" Font-Names="Verdana"
                        Font-Size="X-Small" ForeColor="Navy" Text="Feedback" />
                </td>
                <td style="width: 32%">
                </td>
                <td style="width: 23%">
                </td>
                <td style="width: 32%">
                </td>
                <td style="width: 1%">
                </td>
            </tr>
            <tr>
                <td>
                    &nbsp;
                </td>
                <td align="right">
                    <asp:Label ID="Label8" runat="server" Style="font-size: 08pt" Text="Closed by:" />
                </td>
                <td colspan="3">
                    <asp:TextBox ID="tbClosedBy" runat="server" Style="font-size: 08pt" MaxLength="100"
                        Width="50%" />
                    &nbsp;<asp:RequiredFieldValidator ID="rfvClosedBy" runat="server" ControlToValidate="tbClosedBy"
                        ErrorMessage="## please indicate who is closing this conversation, eg 'John Smith'"
                        Font-Bold="True" Style="font-size: 08pt" ForeColor="Red" InitialValue="" SetFocusOnError="True"
                        ValidationGroup="feedback" />
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td>
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
                <td colspan="3">
                    <asp:Label ID="Label4a0" runat="server" Font-Names="Verdana" Font-Size="X-Small"
                        Text="Please enter any comments you would like to make on the status of your enquiry and the response you received. Thank you." />
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td align="right">
                    <asp:Label ID="Label9" runat="server" Style="font-size: 08pt" Text="Comments:" />
                </td>
                <td colspan="3">
                    <asp:TextBox ID="tbFeedback" runat="server" Style="font-size: 08pt" MaxLength="3000"
                        Rows="6" TextMode="MultiLine" Width="100%" />
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
                <td colspan="3">
                    <asp:Button ID="btnFeedbackSend" runat="server" OnClick="btnFeedbackSend_Click" Text="finish"
                        Width="120px" />
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
        </table>
    </asp:Panel>
    <asp:Panel ID="pnlTopics" runat="server" Width="100%">
        <table style="width: 100%">
            <tr>
                <td style="width: 1%">
                </td>
                <td style="width: 11%">
                </td>
                <td style="width: 32%">
                </td>
                <td style="width: 23%">
                </td>
                <td style="width: 32%">
                </td>
                <td style="width: 1%">
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td>
                    &nbsp;
                </td>
                <td>
                    <asp:Button ID="btnNewTopic" runat="server" Text="contact us" Width="120px" OnClick="btnNewTopic_Click" />
                    <br />
                </td>
                <td>
                </td>
                <td align="right">
                    <asp:CheckBox ID="cbIncludeClosedTopics" runat="server" AutoPostBack="True" OnCheckedChanged="cbIncludeClosedTopics_CheckedChanged"
                        Text="Include closed conversations" Style="font-size: 08pt" />
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                    &nbsp;
                </td>
                <td valign="top">
                    <asp:Label ID="lblLegendTopicList" runat="server" Font-Bold="True" Font-Names="Verdana"
                        Font-Size="X-Small" ForeColor="Navy" Text="Conversation History" />
                </td>
                <td colspan="3">
                    <asp:GridView ID="gvTopics" runat="server" Style="font-size: 08pt" Width="100%" CellPadding="2"
                        AutoGenerateColumns="False" OnRowDataBound="gvTopics_RowDataBound">
                        <Columns>
                            <asp:TemplateField>
                                <ItemTemplate>
                                    &nbsp;<asp:ImageButton ID="imgbtnShowMessages" runat="server" ImageUrl="~/images/icon_list.jpg"
                                        CommandArgument='<%# Container.DataItem("id")%>' OnClick="imgbtnShowMessages_Click"
                                        ToolTip="Show messages" />
                                    <asp:ImageButton ID="imgbtnAddMessage" runat="server" ImageUrl="~/images/icon_plus.jpg"
                                        CommandArgument='<%# Container.DataItem("id")%>' OnClick="imgbtnAddMessage_Click"
                                        ToolTip="Add message" />
                                    <asp:ImageButton ID="imgbtnClose" runat="server" ImageUrl="~/images/icon_handshake.jpg"
                                        CommandArgument='<%# Container.DataItem("id")%>' OnClick="imgbtnClose_Click"
                                        ToolTip="Close topic" />
                                    <asp:HiddenField ID="hidID" runat="server" Value='<%# Bind("ID") %>' />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Center" Wrap="False" />
                            </asp:TemplateField>
                            <asp:BoundField DataField="Topic" HeaderText="Topic" ReadOnly="True" SortExpression="Topic" />
                            <asp:BoundField DataField="CreatedOn" HeaderText="Created On" ReadOnly="True" SortExpression="CreatedOn">
                                <ItemStyle HorizontalAlign="Center" />
                            </asp:BoundField>
                            <asp:BoundField DataField="AWB" HeaderText="Consignment" ReadOnly="True" SortExpression="AWB">
                                <ItemStyle HorizontalAlign="Center" />
                            </asp:BoundField>
                            <asp:BoundField DataField="TopicReference" HeaderText="Conversation Ref" ReadOnly="True"
                                SortExpression="TopicReference">
                                <ItemStyle HorizontalAlign="Center" />
                            </asp:BoundField>
                            <asp:TemplateField HeaderText="Status">
                                <ItemTemplate>
                                    <asp:Label ID="lblTopicStatus" runat="server" Text="<%# sGetTopicStatus(Container.DataItem) %>" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Center" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Rating">
                                <ItemTemplate>
                                    <%--                                    <telerik:RadRating ID="rrFeedback" runat="server" ItemCount="5" Value='<%# Bind("Rating") %>'
                                        AutoPostBack="true" SelectionMode="Continuous" OnRate="rrFeedback_Rate" EnableEmbeddedBaseStylesheet="true"
                                        Precision="Item" Visible="false">
                                    </telerik:RadRating>
                                    --%>
                                    <asp:RadioButtonList ID="rblRating" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                        RepeatDirection="Horizontal" AutoPostBack="true" OnSelectedIndexChanged="rblRating_SelectedIndexChanged">
                                        <asp:ListItem Value="1">OK</asp:ListItem>
                                        <asp:ListItem Value="2">Unsatisfactory</asp:ListItem>
                                    </asp:RadioButtonList>
                                </ItemTemplate>
                            </asp:TemplateField>
                        </Columns>
                        <EmptyDataTemplate>
                            <asp:Label ID="Label5" runat="server" Text="no conversations found"></asp:Label>
                        </EmptyDataTemplate>
                    </asp:GridView>
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td>
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
                <td>
                    <asp:Label ID="lblNewMessagesToView" runat="server" BackColor="#66FF66" Font-Bold="True"
                        Font-Names="Verdana" Font-Size="XX-Small" Text="N E W&amp;nbsp;&amp;nbsp;&amp;nbsp;    M E S S A G E S"></asp:Label>
                </td>
                <td>
                    &nbsp;
                </td>
                <td align="right">
                    <asp:Label ID="lblRatingRequired" runat="server" BackColor="Orange" Font-Bold="True"
                        Font-Names="Verdana" Font-Size="XX-Small" Text="R A T I N G&amp;nbsp;&amp;nbsp;&amp;nbsp; R E Q U I R E D"></asp:Label>
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
        </table>
    </asp:Panel>
    <asp:Panel ID="pnlMessages" runat="server" Width="100%">
        <table style="width: 100%">
            <tr>
                <td style="width: 1%">
                </td>
                <td style="width: 11%">
                    <asp:Label ID="lblLegendTopicList0" runat="server" Font-Bold="True" Font-Names="Verdana"
                        Font-Size="X-Small" ForeColor="Navy" Text="Messages" />
                </td>
                <td style="width: 32%">
                </td>
                <td style="width: 23%">
                </td>
                <td style="width: 32%">
                </td>
                <td style="width: 1%">
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td>
                </td>
                <td colspan="3">
                    <asp:Label ID="lblMessagesForTopicInfo" runat="server" Font-Names="Verdana" Font-Size="X-Small"
                        Font-Bold="True" />
                    &nbsp;&nbsp;&nbsp;
                    <asp:Label ID="lblMessagesConsignmentNumber" runat="server" Font-Names="Verdana"
                        Font-Size="XX-Small" />
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td>
                </td>
                <td colspan="3">
                    <asp:GridView ID="gvMessages" runat="server" CellPadding="3" Font-Names="Verdana"
                        Font-Size="Small" Width="100%" AutoGenerateColumns="False" OnRowDataBound="gvMessages_RowDataBound">
                        <Columns>
                            <asp:BoundField DataField="CreatedOn" HeaderText="Date" ReadOnly="True" SortExpression="CreatedOn">
                                <ItemStyle Width="100px" Wrap="False" HorizontalAlign="Center" />
                            </asp:BoundField>
                            <asp:BoundField DataField="MessageBody" HeaderText="Message" ReadOnly="True" SortExpression="MessageBody" />
                            <asp:TemplateField HeaderText="Author" SortExpression="CreatedBy">
                                <ItemTemplate>
                                    <asp:Label ID="lblAuthor" runat="server" Text="<%# sGetAuthor(Container.DataItem) %>" />
                                </ItemTemplate>
                                <ItemStyle Width="1px" Wrap="False" HorizontalAlign="Center" />
                            </asp:TemplateField>
                        </Columns>
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
                    <asp:Button ID="btnAddMessageFromMessages" runat="server" Text="add message" OnClick="btnAddMessage_Click" />
                </td>
                <td>
                </td>
                <td align="right">
                    <asp:Button ID="btnBackToTopicsFromMessages" runat="server" Text="back to conversations"
                        OnClick="btnBackToTopicsFromMessages_Click" />
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
                <td align="right">
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
        </table>
    </asp:Panel>
    <table style="width: 100%">
        <tr>
            <td style="width: 1%">
            </td>
            <td style="width: 11%">
                &nbsp;
            </td>
            <td style="width: 32%">
            </td>
            <td style="width: 23%">
            </td>
            <td style="width: 32%">
            </td>
            <td style="width: 1%">
            </td>
        </tr>
        <tr>
            <td>
                &nbsp;
            </td>
            <td>
                &nbsp;
            </td>
            <td colspan="3" style="margin-left: 40px">
                <%--<iframe src="http://westernunion.transworld.eu.com/AgentSupportFAQ.aspx" width="100%" height="600"/>--%>
                &nbsp;<asp:Label ID="Label4a1" runat="server" Font-Names="Verdana" Font-Size="X-Small"
                    Text="For help using the Agent Support facility, click " 
                    Font-Bold="True" />
                <asp:HyperLink ID="hlinkClickHere" runat="server" Font-Bold="True" Font-Names="Verdana"
                    Font-Size="X-Small" NavigateUrl="http://westernunion.transworld.eu.com/Help/WesternUnionAgentSupportHELP.aspx"
                    Target="_blank">here</asp:HyperLink>
                <asp:Label ID="Label4a2" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="X-Small"
                    Text="." />
                <div id="dnn_ctr476_ModuleContent" class="DNNModuleContent ModDNNHTMLC">
                    <div id="dnn_ctr476_HtmlModule_lblContent" class="Normal">
                    </div>
                </div>
                <!-- End_Module_476 -->
            </td>
            <td>
                &nbsp;
            </td>
        </tr>
    </table>
    </form>
</body>
</html>
