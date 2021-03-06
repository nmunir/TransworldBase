<%@ Page Language="VB" ValidateRequest="false" Theme="AIMSDefault" %>

<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Web.UI" %>
<%@ Register Assembly="Telerik.Web.UI" Namespace="Telerik.Web.UI" TagPrefix="telerik" %>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Wester Union Mobile</title>
    <link rel="stylesheet" href="~/WUMessage/MobileTheme/mobile-theme.min.css" />
    <link rel="stylesheet" href="http://code.jquery.com/mobile/1.2.0/jquery.mobile.structure-1.2.0.min.css" />
    <script type="text/javascript" src="http://code.jquery.com/jquery-1.7.2.min.js"></script>
    <script type="text/javascript" src="http://code.jquery.com/mobile/1.2.0/jquery.mobile-1.2.0.min.js"></script>
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <style type="text/css">
        .Error
        {
            color: Red;
        }
        input.ui-focus, textarea.ui-focus
        {
            outline: none;
            -webkit-box-shadow: none;
        }
        em
        {
            color: red;
            font-weight: bold;
            padding-right: .25em;
        }
    </style>
    <script type="text/javascript">

        $(function () {

            $("#lblUserID").focus();

        });



        $(function () {
            $("#btn_SendMessage").click(function () {

                var lblUserID = $("#lblUserID");
                var userID = lblUserID.text();

                var txtTseUserID = $("#txtTseUserID");
                var tseUserID = txtTseUserID.val();
                var message = $("#message").val();

                if (tseUserID == '') {

                    $("#TseUserIDError").html("please enter a TSE User ID");
                    $("#TseUserIDError").addClass("Error");
                }

                if (message == '') {

                    $("#messageError").html("please enter a message");
                    $("#messageError").addClass("Error");
                }


                if (tseUserID != '' && message != '') {

                    PageMethods.VerifyUserCredentials(tseUserID, OnSuccessVerifyPassword, OnErrorVerifyPassword);
                }

                return false;


            });
        });

        $(function () {
            $("#button_send").click(function () {
                var agentID = $("#text_agentname").val();
                if (agentID == '') {
                    $("#agentIDError").html("please enter an agent ID");
                    $("#agentIDError").addClass("Error");
                }
                else {
                    PageMethods.IsAgentExist(agentID, OnAgentIDSuccess, OnAgentIDError);
                }

                return false;
            });
        });

        function OnAgentIDSuccess(msg) {

            if (msg != null) {

                if (msg.ID > 0) {

                    $("#lblUserID").html(msg.UserID);
                    $("#lblFirstName").html(msg.FirstName);
                    $("#lblLastName").html(msg.LastName);
                    $.mobile.changePage("#pg_Agent");
                }
            }
            else {
                $("#agentIDError").addClass("Error");
                $("#agentIDError").html("please enter a valid agent ID");


            }

        }

        function OnAgentIDError(msg) {
            alert("An error has as occurred in verifying an Agent ID " + msg);
        }

        function OnSuccessMessage(msg) {

            if (msg > 0) {
                $.mobile.changePage("#div_finish");
                $("#lblMessageSent").html("Message sent successfully");
            }

        }

        function OnErrorMessage(msg) {
            alert("An error has occurred in sending a message " + msg);

        }

        function OnSuccessVerifyPassword(msg) {

            if (msg == false) {

                $("#TseUserIDError").html("User ID or Password doesn't match.");
                $("#TseUserIDError").addClass("Error");


            }
            else {

                var txtTseUserID = $("#txtTseUserID");
                var tseUserID = txtTseUserID.val().split(' ')[0];

                var message = $("#message").val();

                PageMethods.SendMessage(tseUserID, message, OnSuccessMessage, OnErrorMessage);
            }

        }

        function OnErrorVerifyPassword(msg) {

            alert("error");

        }
    
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <telerik:RadScriptManager ID="rsm" EnablePageMethods="true" runat="server">
        </telerik:RadScriptManager>
    </div>
    <div data-role="page" id="pg_VerifyAgentID">
        <div data-role="header" data-theme="b">
            <h1>
                Western Union Messaging</h1>
        </div>
        <div data-role="content">
            <div data-role="fieldcontain">
                <em>*</em>Agent ID :
                <input type="text" id="text_agentname" value="" autofocus="autofocus" class="ui-focus"
                    placeholder="please enter an Agent ID" required />
                <span id="agentIDError"></span>
            </div>
            <div>
                <button id="button_send" type="submit" data-role="button" data-theme="b">
                    Send</button>
            </div>
        </div>
        <div data-role="dialog" id="ErrorPopup">
            <p>
                please enter a valid Agent ID<p>
        </div>
    </div>
    <!-- Start of second page -->
    <div data-role="page" id="pg_Agent">
        <div data-role="header" data-theme="b">
            <h1>
                Western Union Agent Info</h1>
        </div>
        <div data-role="content">
            <div data-role="fieldcontain">
                <div>
                    User ID :
                    <label id="lblUserID" data-theme="b">
                    </label>
                </div>
                <div>
                    First Name :
                    <label id="lblFirstName" data-theme="b">
                    </label>
                </div>
                <div>
                    Last Name :
                    <label id="lblLastName" data-theme="b">
                    </label>
                </div>
                <div>
                    <em>*</em>TSE User ID :
                    <input type="text" id="txtTseUserID" value="" autofocus="autofocus" placeholder="please enter a user id and two characters of password separated by a space"
                        required />
                    <span id="TseUserIDError"></span>
                </div>
                <%-- <div>
                    Topic Categories :
                    <ul id="lvTopicCategories" data-role="listview" data-theme="b">
                    </ul>
                </div>--%>
                <div>
                    <em>*</em>Message :
                    <textarea id="message" value="" data-theme="b" placeholder="please enter a message"
                        required></textarea>
                    <span id="messageError"></span>
                </div>
                <div>
                    <button id="btn_SendMessage" type="submit" data-role="button" data-theme="b">
                        Send Message</button>
                </div>
            </div>
        </div>
    </div>
    <!-- Start of third page -->
    <div data-role="page" id="div_finish">
        <div data-role="header" data-theme="b">
            <h1>
                Western Union Messaging</h1>
        </div>
        <div data-role="content">
            <div>
                <label id="lblMessageSent" data-theme="b">
                </label>
            </div>
            <div>
                <a id="Finish_Send" href="#pg_VerifyAgentID" type="submit" data-role="button" data-theme="b">
                    Finish</a>
            </div>
        </div>
    </div>
    </form>
</body>
</html>
<script runat="server">
    
    Private Shared gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    Private Shared CustomerKeys As String = "579,686"
    Const TOPIC_STATUS_OPEN As Int32 = 0                                      ' topic is open
    
    <System.Web.Services.WebMethod()>
    Public Shared Function GetTopicCategories(ByVal sUserID As String) As List(Of MessagingTopicCategories)
        GetTopicCategories = Nothing
        Dim IListOfTopics As New List(Of MessagingTopicCategories)
        Dim sSQL As String = "select ID, CategoryName from MessagingTopicCategories where CustomerKey = (select CustomerKey from UserProfile where UserId = '" & sUserID.ToUpper & "')"
        Dim oDataTable As New DataTable
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter(sSQL, oConn)
        Try
            oConn.Open()
            oAdapter.Fill(oDataTable)
            If Not oDataTable Is Nothing AndAlso oDataTable.Rows.Count > 0 Then
                For Each dr As DataRow In oDataTable.Rows
                    Dim obj_mtc As New MessagingTopicCategories
                    obj_mtc.ID = dr("ID").ToString
                    obj_mtc.CategoryName = dr("CategoryName").ToString
                    IListOfTopics.Add(obj_mtc)
                Next
            End If
            GetTopicCategories = IListOfTopics
        Catch ex As Exception
            WebMsgBox.Show(ex.Message.ToString)
        Finally
            oConn.Close()
        End Try
    End Function
    
    <System.Web.Services.WebMethod()>
    Public Shared Function IsAgentExist(ByVal sUserID As String) As UserProfile
        IsAgentExist = Nothing
        Dim sSQL As String = "select [Key] 'ID', UserID, FirstName, LastName from UserProfile where CustomerKey IN(" & CustomerKeys & ") and [Status] = 'Active' and UserID = '" & sUserID.ToUpper & "'"
        Dim oDataTable As DataTable = ExecuteQueryToDataTable(sSQL)
        If Not oDataTable Is Nothing AndAlso oDataTable.Rows.Count > 0 Then
            Dim dr As DataRow = oDataTable.Rows(0)
            Dim up As New UserProfile
            up.ID = Convert.ToInt32(dr("ID"))
            up.UserID = dr("UserID").ToString
            up.FirstName = dr("FirstName").ToString
            up.LastName = dr("LastName").ToString
            IsAgentExist = up
        End If
     
    End Function
    
    <System.Web.Services.WebMethod()>
    Public Shared Function VerifyUserCredentials(sUserIDAndTwoCharsPassword As String) As Boolean
        
        VerifyUserCredentials = False
        
        Dim sUserName As String
        Dim sPasswordFragment As String
        Dim nPasswordLength As Int32
        
        If sUserIDAndTwoCharsPassword.Contains(" ") Then
            sUserName = sUserIDAndTwoCharsPassword.Split(" ")(0)
            sPasswordFragment = sUserIDAndTwoCharsPassword.Split(" ")(1)
        Else
            VerifyUserCredentials = False
            Exit Function
        End If
        
        sUserName = sUserName.ToUpper
        
        nPasswordLength = Len(sPasswordFragment)
        If nPasswordLength < 2 Then
            Exit Function
        End If
        
        Dim oUserInfo As SprintInternational.UserInfo = New SprintInternational.UserInfo()
        Dim oLogon As SprintInternational.Logon = New SprintInternational.Logon()
        Dim oPassword As SprintInternational.Password = New SprintInternational.Password()
        
        oUserInfo = oLogon.GetUserInfo(sUserName)
        
        If oUserInfo.UserKey = -1 Then
            Exit Function
        Else
            Dim sActualPassword As String = oPassword.Decrypt(oUserInfo.Password)
            If nPasswordLength <= Len(sActualPassword) Then
                If sPasswordFragment = sActualPassword.Substring(0, nPasswordLength) Then
                    VerifyUserCredentials = True
                End If
            End If
        
        End If
        
    End Function
    
    <System.Web.Services.WebMethod()>
    Public Shared Function SendMessage(ByVal sUserID As String, ByVal sMessage As String) As Int32
        
        SendMessage = 0
        Dim sSQL As String = "INSERT INTO MessagingTopics (UserKey, TopicStatus, Topic, TopicReference, LastTopicState, LastTopicStateChange, AWB, NewMessage, CreatedOn, CreatedBy) VALUES ("
        sSQL += "(Select [Key] From UserProfile where UserID = '" & sUserID.Replace("'", "''").ToUpper.ToString & "'), " & TOPIC_STATUS_OPEN & ", '" & "Created_From_Mobile" & "','', 'OPEN', GETDATE(), '" & String.Empty & "', 0, GETDATE(),(Select [Key] From UserProfile where UserID = '" & sUserID.ToUpper.ToString() & "')) SELECT SCOPE_IDENTITY()"
        Dim oDataTable As DataTable = ExecuteQueryToDataTable(sSQL)
        If Not oDataTable Is Nothing AndAlso oDataTable.Rows.Count > 0 Then
            Dim dr As DataRow = oDataTable.Rows(0)
            Dim nTopicID As Integer = oDataTable.Rows(0).Item(0)
            sSQL = String.Empty
            sSQL = "UPDATE MessagingTopics SET TopicReference = '" & sUserID.ToUpper.ToString & "_" & nTopicID.ToString.PadLeft(7, "0") & "' where ID = " & nTopicID
            Call ExecuteQueryToDataTable(sSQL)
            sSQL = String.Empty
            Dim sMessageTimestamp As String = "." & Format(DateTime.Now, "yyyyMMddhhmmss")
            sSQL = "INSERT INTO MessagingMessages (CustomerKey, TopicNumber, MessageRef, MessageBody, IsDeleted, IsAdmin, CreatedOn, CreatedBy) VALUES ("
            sSQL += "(Select CustomerKey from UserProfile where UserId = '" & sUserID & "')," & nTopicID.ToString & ",'" & sUserID.ToUpper.ToString & "_" & nTopicID.ToString & sMessageTimestamp & "','" & sMessage.Replace("'", "''") & "','0','0', GETDATE(), (Select CustomerKey from UserProfile where UserId = '" & sUserID & "')) SELECT SCOPE_IDENTITY()"
            sSQL += "Update MessagingTopics set NewAgentMessage = 1 where ID = " & nTopicID
            oDataTable = ExecuteQueryToDataTable(sSQL)
            SendMessage = oDataTable.Rows(0).Item(0)
        End If
        
    End Function
    
    Protected Shared Function ExecuteQueryToDataTable(ByVal sQuery As String) As DataTable
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
    
    Public Class UserProfile
        
        Private m_ID As Integer
        Private m_UserID As String
        Private m_FirstName As String
        Private m_LastName As String
        
        Public Property ID() As String
            Get
                Return m_ID
            End Get
            Set(value As String)
                m_ID = value
            End Set
        End Property
        
        Public Property UserID() As String
            Get
                Return m_UserID
            End Get
            Set(value As String)
                m_UserID = value
            End Set
        End Property
        
        Public Property FirstName() As String
            Get
                Return m_FirstName
            End Get
            Set(value As String)
                m_FirstName = value
            End Set
        End Property
        Public Property LastName() As String
            Get
                Return m_LastName
            End Get
            Set(value As String)
                m_LastName = value
            End Set
        End Property
    End Class
    
    Public Class MessagingTopicCategories
        Private m_ID As String
        Private m_CategoryName As String
        
        
        Public Property ID() As String
            Get
                Return m_ID
            End Get
            Set(value As String)
                m_ID = value
            End Set
        End Property
        
        Public Property CategoryName() As String
            Get
                Return m_CategoryName
            End Get
            Set(value As String)
                m_CategoryName = value
            End Set
        End Property
    End Class
    
</script>
