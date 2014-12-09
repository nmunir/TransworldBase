<%@ Page Language="VB" ValidateRequest="false" %>

<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Web.UI" %>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Western Union Messaging</title>
    <link rel="stylesheet" href="http://code.jquery.com/mobile/1.2.0/jquery.mobile-1.2.0.min.css" />
    <script type="text/javascript" src="http://code.jquery.com/jquery-1.8.2.min.js"></script>
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
        
        .clear
        {
            clear: both;
        }
        
        
        .hide
        {
            display: none;
        }
        
        .ui-block-b
        {
            text-align: right;
        }
    </style>
    <script type="text/javascript">

        String.prototype.trim = function () {
            return this.replace(/^\s+|\s+$/g, "");
        }


        function getPageName() {

            var foundPageName = false;

            urlStr = window.location.href;

            if (urlStr.indexOf("#") > -1) {
                var param = urlStr.split("#");
                var pageName = param[1];

                if (pageName.toLowerCase() == 'pg_agent') {

                    foundPageName = true;
                    return foundPageName;
                }

                else {
                    return foundPageName;
                }

            }

        }

        $(function () {

            $("#lblUserID").focus();
            $("#divCustomTopic").addClass("hide");

            if (sessionStorage.AgentKey > 0) {

                var foundPageInUrl = getPageName();

                if (foundPageInUrl) {

                    PageMethods.GetTopicCategories(sessionStorage.AgentKey, OnSuccessTopicCategory, OnErrorTopicCategory);
                    PageMethods.AgentInfo(sessionStorage.AgentKey, OnSuccessAgentInfo, OnErrorAgentInfo);

                }

            }



        });



        $(function () {


            $("#selectTopicCategory").on('change', function () {

                var selectTopicCategory = document.getElementById('selectTopicCategory');
                var topicName = selectTopicCategory.options[selectTopicCategory.selectedIndex].text;


                if (topicName.toLowerCase() == "other") {


                    $("#divCustomTopic").removeClass("hide");
                    $("#txtCustomTopic").focus();

                }
                else {

                    $("#divCustomTopic").addClass("hide");

                }



            });



            $("#btnFinish").click(function () {

                $.mobile.changePage('#pg_VerifyAgentID');
                $("#text_agentname").val("");
                $("#message").val("");
                $("#txtTseUserID").val("");
                return false;
            });



        });


        $(function () {
            $("#button_send").click(function () {
                var agentID = $("#text_agentname").val();
                if (agentID == '') {
                    $("#agentIDError").html("enter an agent ID");
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

                    $.mobile.changePage('#pg_VerifyAgentID');
                    sessionStorage.AgentKey = msg.ID;
                    PageMethods.GetTopicCategories(msg.ID, OnSuccessTopicCategory, OnErrorTopicCategory);
                    PageMethods.AgentInfo(sessionStorage.AgentKey, OnSuccessAgentInfo, OnErrorAgentInfo);

                }
            }
            else {
                $("#agentIDError").addClass("Error");
                $("#agentIDError").html(" Agent ID not found");
            }

        }

        function OnSuccessAgentInfo(msg) {

            $("#AgentInfo_header").html(sessionStorage.topicName);
            $("#lblUserID").html(msg.UserID);
            $("#lblFirstName").html(msg.FirstName);
            $("#lblLastName").html(msg.LastName);
            if (localStorage.TseUserID != null) {

                $("#divTseUserID").show();
                $("#lblTseUserID").html(localStorage.TseUserID);
            }
            else {
                $("#divTseUserID").hide();
                $("#lblTseUserID").html(localStorage.TseUserID);
            }

            $.mobile.changePage('#pg_Agent');

        }

        function OnErrorAgentInfo(msg) {

            alert("error occurred in getting agentinfo.");

        }


        function OnSuccessTopicCategory(data) {

            var categories = data;

            $('#selectTopicCategory').empty();

            $('#selectTopicCategory').append('<option>- select topic -</option>');

            $.each(categories, function (index, category) {

                $('#selectTopicCategory').append('<option>' + category.CategoryName + '</option>');

            });

            $('#selectTopicCategory').refresh();

        }


        function OnErrorTopicCategory(msg) {

            alert("error topic category")

        }


        function OnAgentIDError(msg) {
            alert("An error has as occurred in verifying an Agent ID " + msg);
        }

        $(function () {
            $("#btn_SendMessage").click(function () {


                var lblUserID = $("#lblUserID");
                var userID = lblUserID.text();

                var txtTseUserID = $("#txtTseUserID");
                var tseUserID = txtTseUserID.val().trim();

                var message = $("#message").val();

                var selectTopicCategory = document.getElementById('selectTopicCategory');

                var topicName = selectTopicCategory.options[selectTopicCategory.selectedIndex].text;

                if (tseUserID == '') {

                    $("#TseUserIDError").html("enter a TSE User ID and/or password");
                    $("#TseUserIDError").addClass("Error");
                }
                else {

                    $("#TseUserIDError").html("");
                    $("#TseUserIDError").removeClass("Error");

                }

                if (message == '') {

                    $("#messageError").html("enter a message");
                    $("#messageError").addClass("Error");
                }
                else {

                    $("#messageError").html("");
                    $("#messageError").removeClass("Error");

                }

                if (selectTopicCategory.selectedIndex == 0) {

                    $("#selectTopicCategoryError").html("select a conversation title");
                    $("#selectTopicCategoryError").addClass("Error");

                }
                else {

                    $("#selectTopicCategoryError").html("");
                    $("#selectTopicCategoryError").removeClass("Error");

                }

                if (tseUserID != '' && message != '' && selectTopicCategory.selectedIndex > 0) {



                    if (tseUserID.indexOf(' ') != -1) {


                        PageMethods.VerifyUserCredentials(tseUserID, OnSuccessVerifyPassword, OnErrorVerifyPassword);

                    }
                    else {

                        var lblTseUserID = $("#lblTseUserID");
                        var tseUserIDAndPw = lblTseUserID.text() + ' ' + tseUserID;
                        PageMethods.VerifyUserCredentials(tseUserIDAndPw, OnSuccessVerifyPassword, OnErrorVerifyPassword);
                    }
                }

                return false;

            });

        });


        function OnSuccessVerifyPassword(msg) {

            if (msg == false) {

                $("#TseUserIDError").html("User ID or Password doesn't match.");
                $("#TseUserIDError").addClass("Error");


            }
            else {

                var txtTseUserID = $("#txtTseUserID");
                var tseUserID;

                if (txtTseUserID.val().trim().indexOf(' ') != -1) {
                    tseUserID = txtTseUserID.val().trim().split(' ')[0];
                }

                else {
                    tseUserID = $("#lblTseUserID").text();
                }

                var message = $("#message").val();

                var topicName = $("#selectTopicCategory").val();

                if (topicName.toLowerCase() == "other") {

                    topicName = $("#txtCustomTopic").val();

                }

                PageMethods.SendMessage(tseUserID, message, topicName, OnSuccessMessage, OnErrorMessage);
            }

        }

        function OnSuccessMessage(msg) {

            if (msg > 0) {

                var isRemember = $("#sliderRemember").attr('value');

                if (isRemember == 'true') {
                    var txtTseUserID = $("#txtTseUserID");
                    if (txtTseUserID.val().trim().indexOf(' ') != -1) {
                        var tseUserID = txtTseUserID.val().trim().split(' ')[0];
                        localStorage.TseUserID = tseUserID;
                    }
                    else {
                        var tseUserID = $("#lblTseUserID").text();
                        localStorage.TseUserID = tseUserID;
                    }

                }
                else {

                    var tseUserID = $("#lblTseUserID").text();
                    localStorage.removeItem("TseUserID");
                }

                $.mobile.changePage("#div_finish");
                $("#lblMessageSent").html("Message sent successfully");
            }

        }

        function OnErrorMessage(msg) {
            alert("An error has occurred in sending a message " + msg);

        }

        function OnErrorVerifyPassword(msg) {
            alert("error");
        }
    
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:ScriptManager ID="sm" EnablePageMethods="true" runat="server">
        </asp:ScriptManager>
    </div>
    <div data-role="page" id="pg_VerifyAgentID">
        <div data-role="header" data-theme="b">
            <h1>
                Western Union Messaging</h1>
        </div>
        <div data-role="content">
            <div data-role="fieldcontain">
                <div class="ui-grid-a">
                    <div class="ui-block-a">
                        <em>*</em>Agent ID:</div>
                    <div class="ui-block-b">
                        <a style="text-align: right" href="http://westernunion.transworld.eu.com/Help/WesternUnionMobileAppHELP.aspx" target="_blank">
                            help</a>
                    </div>
                </div>
                <div class="clear">
                    &nbsp;
                </div>
                <div>
                    <input type="text" id="text_agentname" value="" autofocus="autofocus" maxlength="10"
                        class="ui-focus" placeholder="enter an Agent Terminal ID" required />
                </div>
                <div>
                    <span id="agentIDError"></span>
                </div>
                <div class="clear">
                    &nbsp;
                </div>
                <div>
                    <button id="button_send" type="submit" data-role="button" data-theme="b">
                        Send</button>
                </div>
            </div>
            <div data-role="dialog" id="ErrorPopup">
                <p>
                    enter a valid Agent ID<p>
            </div>
        </div>
    </div>
    <!-- Start of second page -->
    <!-- Start of third page -->
    <div data-role="page" id="pg_Agent">
        <div data-role="header" data-theme="b">
            <div data-role="header" data-theme="b">
                <h1 id="AgentInfo_header">
                    Western Union Messaging</h1>
            </div>
        </div>
        <div data-role="content">
            <div>
                <label for="selectTopicCategory">
                    Choose conversation title:</label></div>
            <div>
                <select data-mini="true" id="selectTopicCategory" data-theme="b">
                </select>
                <span id="selectTopicCategoryError"></span>
            </div>
            <div id="divCustomTopic">
                <div>
                    Enter conversation title:</div>
                <div>
                    <input id="txtCustomTopic" type="text" data-mini="true" maxlength="50" />
                </div>
            </div>
            <div>
                <b>
                    <label id="lblUserID" data-theme="b">
                    </label>
                    &nbsp;-&nbsp;
                    <label id="lblFirstName" data-theme="b">
                    </label>
                    &nbsp;
                    <label id="lblLastName" data-theme="b">
                    </label>
                </b>
            </div>
            <div class="clear">
                &nbsp;
            </div>
            <div id="divTseUserID">
                <br />
            </div>
            <div class="clear">
                &nbsp;
            </div>
            <div>
                <em>*</em>TSE User ID &amp; password: (<label id="lblTseUserID" runat="server" data-theme="b" />)</div>
            <div>
                <input type="text" id="txtTseUserID" value="" autofocus="autofocus" maxlength="20"
                    data-mini="true" required />
                <span id="TseUserIDError"></span>
            </div>
            <div class="clear">
                &nbsp;
            </div>
            <div>
                <label for="sliderRemember">
                    Remember me?</label>
            </div>
            <div>
                <select id="sliderRemember" data-role="slider" data-mini="true">
                    <option value="false">no</option>
                    <option value="true">yes</option>
                </select>
            </div>
            <div class="clear">
            </div>
            <div>
                <em>*</em>Message:</div>
            <div>
                <textarea id="message" value="" data-theme="b" placeholder="enter your message" data-mini="true"
                    maxlength="200" required></textarea>
                <span id="messageError"></span>
            </div>
            <div class="clear">
            </div>
            <div>
                &nbsp;
            </div>
            <div>
                <button id="btn_SendMessage" type="submit" data-role="button" data-theme="b">
                    Send Message</button>
            </div>
        </div>
    </div>
    <!-- Start of fourth page -->
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
                <button id="btnFinish" type="submit" data-role="button" data-theme="b">
                    Finish</button>
            </div>
        </div>
    </div>
    </form>
</body>
</html>
<script runat="server">
    
    'Private Shared gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    Private Shared gsConn As String = ConfigurationManager.ConnectionStrings("AIMSRootConnectionString").ToString
    Private Shared CustomerKeys As String = "579,686"
    Const TOPIC_STATUS_OPEN As Int32 = 0                                      ' topic is open
    
    <System.Web.Services.WebMethod()>
    Public Shared Function GetTopicCategories(ByVal sUserID As String) As List(Of MessagingTopicCategories)
        GetTopicCategories = Nothing
        Dim IListOfTopics As New List(Of MessagingTopicCategories)
        Dim sSQL As String = "select ID, CategoryName from MessagingTopicCategories where CustomerKey = (select CustomerKey from UserProfile where [Key] = '" & sUserID.ToUpper & "')"
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
            'WebMsgBox.Show(ex.Message.ToString)
        Finally
            oConn.Close()
        End Try
    End Function
    
    <System.Web.Services.WebMethod()>
    Public Shared Function IsAgentExist(ByVal sUserID As String) As UserProfile
        IsAgentExist = Nothing
        Dim sSQL As String = "select [Key] 'ID', UserID, FirstName, LastName from UserProfile where CustomerKey IN(" & CustomerKeys & ") and [Status] = 'Active' and UserID = '" & sUserID & "'"
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
    Public Shared Function AgentInfo(ByVal sAgentKey As String) As UserProfile
        AgentInfo = Nothing
        Dim sSQL As String = "select [Key] 'ID', UserID, FirstName, LastName from UserProfile where CustomerKey IN(" & CustomerKeys & ") and [Status] = 'Active' and [Key] = " & sAgentKey
        Dim oDataTable As DataTable = ExecuteQueryToDataTable(sSQL)
        If Not oDataTable Is Nothing AndAlso oDataTable.Rows.Count > 0 Then
            Dim dr As DataRow = oDataTable.Rows(0)
            Dim up As New UserProfile
            up.ID = Convert.ToInt32(dr("ID"))
            up.UserID = dr("UserID").ToString
            up.FirstName = dr("FirstName").ToString
            up.LastName = dr("LastName").ToString
            AgentInfo = up
        End If
     
    End Function
    
    <System.Web.Services.WebMethod()>
    Public Shared Function VerifyUserCredentials(sUserIDAndTwoCharsPassword As String) As Boolean
        
        VerifyUserCredentials = False
        
        Dim sUserName As String
        Dim sPasswordFragment As String
        Dim nPasswordLength As Int32
        
        If sUserIDAndTwoCharsPassword.Trim.Contains(" ") Then
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
    Public Shared Function SendMessage(ByVal sUserID As String, ByVal sMessage As String, ByVal sTopicName As String) As Int32
        
        SendMessage = 0
        Dim sSQL As String = "INSERT INTO MessagingTopics (UserKey, TopicStatus, Topic, TopicReference, LastTopicState, LastTopicStateChange, AWB, NewMessage, CreatedOn, CreatedBy) VALUES ("
        sSQL += "(Select [Key] From UserProfile where UserID = '" & sUserID.Replace("'", "''").ToUpper.ToString & "'), " & TOPIC_STATUS_OPEN & ", '" & sTopicName & "','', 'OPEN', GETDATE(), '" & String.Empty & "', 0, GETDATE(),(Select [Key] From UserProfile where UserID = '" & sUserID.ToUpper.ToString() & "')) SELECT SCOPE_IDENTITY()"
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
            'WebMsgBox.Show("Error in ExecuteQueryToDataTable executing: " & sQuery & " : " & ex.Message)
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
