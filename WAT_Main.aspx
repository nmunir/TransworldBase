<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ Register Namespace="SampleControls" TagPrefix="sc" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.ComponentModel" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Web.UI" %>
<%@ Import Namespace="Telerik.Web.UI" %>
<%@ Register Assembly="Telerik.Web.UI" Namespace="Telerik.Web.UI" TagPrefix="telerik" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<script runat="server">

    ' TO DO
    ' check logic with Complete / Mark as Reopened button
    ' check audit trail
    
    ' SELECT * INTO ClientData_WAT_AuditTrail2 FROM ClientData_WAT_AuditTrail
    ' SELECT * INTO ClientData_WAT_Notifications2 FROM ClientData_WAT_Notifications
    ' SELECT * INTO ClientData_WAT_Termination2 FROM ClientData_WAT_Termination

    ' DELETE FROM ClientData_WAT_AuditTrail2
    ' DELETE FROM ClientData_WAT_Notifications2
    ' DELETE FROM ClientData_WAT_Termination2

    ' SELECT * FROM ClientData_WAT_AuditTrail2
    ' SELECT * FROM ClientData_WAT_Notifications2
    ' SELECT * FROM ClientData_WAT_Termination2

    ' REPORTS
    ' Number of agents contacted (number of calls & was there a planned collection – also listing of Termination Abandoned and reason)
    ' Numbers of visits per agent
    ' Number of completed collections
    ' Items collected (Forms, general POS, light box, transfer station, street sign, computer)
    ' Cost per collection
    ' Average cost per collection
    ' Failed collections and reason
    ' Total spend to date

    ' CRIB SHEET NOTES
    ' accounts: wucolla (view only), wucollb (create new termination, edit termination details until first agent contact attempt), wucollc (full privilege)
    ' WURS: 11ZP
    
    ' TO DO
    ' Record items to be collected
    ' Improve notification email appearance
    ' check what NEXT ACTION is set with various outputs from Contact Agent panel
    ' check availability of packaging items before placing order
    ' copy latest database then update constant USER_WUCOLL
    ' check action of search with bidirectional sorting
    ' Add ability to edit address while in Contact Agent screen
    ' Add filtering by country
    ' btnEditAgentDetails - disable once attempt has been made to contact agent
    ' Look at spreadsheets in \\public_knowledge\Western Union\Agent Termination to see if they can be imported
    ' sort out margins / padding

    ' access list coming
    ' colour coding required
    ' consolidated report for all countries, with option to narrow to single country
    ' default audit trail to show main events only, with button to add minor events
    ' live on Friday (13DEC ?
    ' MS sent list of reports required
    ' QUESTIONS FOR MARILYN
    
    ' order of sending packaging / arranging collection?
    ' 

    ' QUESTIONS RAISED BY MARK
    ' check cannot enter duplicate entries
    ' contact screen - invert
    ' A4FP missing entries

    ' NOTES

    ' <%@ Import Namespace="System.Web.UI.Controls" %>

    ' change AgentTermID to AgentID
    ' change ClientData_WAT_AuditTrail.LastChangedOn to ClientData_WAT_AuditTrail.AuditEntryDateTime
    ' change ClientData_WAT_AuditTrail.LastChangedBy to ClientData_WAT_AuditTrail.AuditEntryUserKey
    
    ' check for session timeout

    ' ShowMessage("Termination " & sStatus & ".")
    ' Call AuditEntry(ENTRY_TYPE_NEW_TERMINATION, "")
    ' Call DisplayFatalError(sQuery & " : " & ex.Message)
    
    ' use CollectionPoint to control access: -1 = not permissioned; 0 or blank = can view only; 1 = can create new entries, edit address until something else happens; 2 = can do anything

    Const COUNTRY_UK_EXCLUDING_NORTHERN_IRELAND As Int32 = 222
    Const COUNTRY_NORTHERN_IRELAND As Int32 = 260
    Const COUNTRY_IRISH_REPUBLIC As Int32 = 103

    Const PERMISSION_NONE_NEGATIVE_1 As Int32 = -1
    Const PERMISSION_VIEW_ONLY_0 As Int32 = 0
    Const PERMISSION_CREATE_EDIT_ENTRY_1 As Int32 = 1
    Const PERMISSION_ALL_2 As Int32 = 2
    
    Const INTERVAL_SHORT_MESSAGE As Int32 = 3000
    Const INTERVAL_SHORT_CONFIRMATION_MESSAGE As Int32 = 3000
    Const INTERVAL_SHORT_ERROR_MESSAGE As Int32 = 3001
    Const INTERVAL_LONG_MESSAGE As Int32 = 6000
    Const INTERVAL_PERMANENT_MESSAGE As Int32 = -1

    Const ENTRY_TYPE_NEW_TERMINATION As String = "NEW TERMINATION"
    Const ENTRY_TYPE_DETAIL_CHANGE As String = "DETAIL CHANGE"
    Const ENTRY_TYPE_STATUS_CHANGE As String = "STATUS CHANGE"
    Const ENTRY_TYPE_EVENT As String = "EVENT"
    Const ENTRY_TYPE_NOTE As String = "NOTE"
    Const ENTRY_TYPE_TERMINATION_COMPLETION_STATUS As String = "TERMINATION COMPLETION STATUS"
    
    Const CUSTOMER_WUCOLL As Int32 = 840
    'Const USER_WUCOLL As String = "WUCOLLapp"
    Const USER_WUCOLL As String = "marilynwucoll"
    
    Const PRODUCT_KEY_ENVELOPE As Int32 = 86760
    Const PRODUCT_KEY_SMALL_BOX As Int32 = 86761
    Const PRODUCT_KEY_MEDIUM_BOX As Int32 = 86762
    Const PRODUCT_KEY_LARGE_BOX As Int32 = 86802
    Const PRODUCT_KEY_LARGE_BOX_PAVEMENT_SIGN As Int32 = 86763
    Const PRODUCT_KEY_LARGE_BOX_TXFER_STN As Int32 = 86764
    
    Const NEXT_ACTION_CONTACT_AGENT As String = "01CONTACT AGENT"
    Const NEXT_ACTION_COLLECT_FORMS As String = "02COLLECT FORMS"
    'Const NEXT_ACTION_NOTIFY_RECEIVED_AT_TRANSWORLD As String = "03NOTIFY RECEIVED AT TRANSWORLD"
    Const NEXT_ACTION_DELIVER_TO_WESTERN_UNION = "04DELIVER TO WESTERN UNION"
    Const NEXT_ACTION_COMPLETE_TERMINATION = "05COMPLETE TERMINATION"
    Const NEXT_ACTION_COMPLETED = "06COMPLETED"

    Const REPORTS_MENU_REPORT_A As String = "Export All Data"
    Const REPORTS_MENU_REPORT_B As String = "Report (not yet implemented)"
    Const REPORTS_MENU_REPORT_C As String = "Report (not yet implemented)"
    Const REPORTS_MENU_REPORT_D As String = "Report (not yet implemented)"
    
    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    Private gdtBasket As DataTable
    Private gsSystemErrorMessage As String

#Region "Initialisation & Choreography"
    
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)

        If Not IsNumeric(Session("UserKey")) Then
            Response.RedirectLocation = "http:/my.transworld.eu.com/common/session_expired.aspx"
            Server.Transfer("session_expired.aspx")
        End If

        If Not IsPostBack Then
            Call HideAllPanels()
            Call InitTerminationCountryDropDown()
            Call SetPrivilege()
            'Call BuildReportsMenu()
            If pnPermission >= 0 Then
                Call DisplayMainPanel()
            End If
        End If
        Call SetTitle()
    End Sub

    Protected Sub BuildReportsMenu()
        Dim rmi As RadMenuItem = New RadMenuItem
        rmi.Text = "Reports"
        Dim rmi2 As RadMenuItem = New RadMenuItem(REPORTS_MENU_REPORT_A, "ExportData")
        rmi.Items.Add(rmi2)
        'rmi.Items.Add(New RadMenuItem(REPORTS_MENU_REPORT_A))
        rmi.Items.Add(New RadMenuItem(REPORTS_MENU_REPORT_B))
        rmi.Items.Add(New RadMenuItem(REPORTS_MENU_REPORT_C))
        rmi.Items.Add(New RadMenuItem(REPORTS_MENU_REPORT_D))

        radmenuReports.Items.Add(rmi)
        'radmenuReports.Items.Add(New RadMenuItem(REPORTS_MENU_ITEM1))
        'radmenuReports.Items.Add(New RadMenuItem(REPORTS_MENU_ITEM2))
        
    End Sub

    Protected Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim sm As New ScriptManager
        sm.ID = "ScriptMgr"
        Try
            PlaceHolderForScriptManager.Controls.Add(sm)
        Catch ex As Exception
        End Try
    End Sub
    
    Protected Sub SetTitle()
        Dim sTitle As String = Session("SiteTitle")
        If sTitle <> String.Empty Then
            sTitle += " - "
        End If
        Page.Header.Title = sTitle & "WU Agent Terminations"
    End Sub
   
    Protected Sub SetPrivilege()
        'If Session("UserKey") Is Nothing Then
        '    Session("UserKey") = 16747     ' kejaz, for testing
        'End If
        
        Dim sPrivilege As String = ExecuteQueryToDataTable("SELECT ISNULL(CollectionPoint, '') 'CollectionPoint' FROM UserProfile WHERE [key] = " & Session("UserKey")).Rows(0).Item(0)
        
        'Dim sPrivilege As String = "2"
        If IsNumeric(sPrivilege) Then
            pnPermission = CInt(sPrivilege)
        Else
            pnPermission = -1
        End If
        Select Case pnPermission
            Case PERMISSION_VIEW_ONLY_0
                btnNewTermination.Visible = False
                Call ButtonPanelVisibility(False)
            Case PERMISSION_CREATE_EDIT_ENTRY_1
                Call ButtonPanelVisibility(False)
                btnEditAgentDetails.Visible = True
            Case PERMISSION_ALL_2
                
            Case Else
                pnlNoPermission.Visible = True
        End Select
    End Sub
    
    Protected Sub HideAllPanels()
        pnlMain.Visible = False
        pnlTerminationAddress.Visible = False
        pnlTerminationManagement.Visible = False
        pnlNoPermission.Visible = False
        pnlNotifications.Visible = False
        pnlSetNextAction.Visible = False
        Call HideDialogPanels()
    End Sub
    
    Protected Sub HideDialogPanels()
        pnlContactAgent.Visible = False
        pnlFormsReceived.Visible = False
        pnlFormsSent.Visible = False
        pnlTerminationComplete.Visible = False
        pnlAddNote.Visible = False
    End Sub
    
    Protected Sub DisplayMainPanel()
        Call HideAllPanels()
        Call InitMainPanel()
        pnlMain.Visible = True
    End Sub

#End Region
    
    Protected Sub btnNewTermination_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call HideAllPanels()
        Call InitNewTerminationPanel()
        pnlTerminationAddress.Visible = True
    End Sub
    
    Protected Sub btnFindUniqueID_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        tbTerminationUniqueID.Text = tbTerminationUniqueID.Text.Trim
        If tbTerminationUniqueID.Text.Length = 4 Then
            If Not LookupWURSAgent() Then
                If Not LookupWUIREAgent() Then
                    If Not LookupCOSTAAgent() Then
                        Call ShowMessage("Could not find a WURS, WUIRE or COSTA agent matching this Terminal ID.", INTERVAL_SHORT_ERROR_MESSAGE)
                        psTimerAction = "BAD_UNIQUEID"
                    End If
                End If
            End If
        ElseIf tbTerminationUniqueID.Text.Length = 6 Then
            If Not LookupFININTAgent() Then
                Call ShowMessage("Could not find a FININT agent matching this Account Number.", INTERVAL_SHORT_ERROR_MESSAGE)
                psTimerAction = "BAD_UNIQUEID"
            End If
        Else
            If tbTerminationUniqueID.Text = String.Empty Then
                Call ShowMessage("Please provide a Unique Identifier (typically a 4-digit Terminal ID (WURS, WUIRE, COSTA) or a 6-digit Account Number (FININT).", INTERVAL_SHORT_ERROR_MESSAGE)
            Else
                Call ShowMessage("ERROR - can only match 4-character Terminal IDs (WURS, WUIRE, COSTA) or 6-digit Account Numbers (FININT).", INTERVAL_SHORT_ERROR_MESSAGE)
            End If
            psTimerAction = "BAD_UNIQUEID"
        End If
    End Sub

    Protected Function LookupWURSAgent() As Boolean
        LookupWURSAgent = False
        Dim sSQL As String = "SELECT * FROM ClientData_WU_Agents WHERE TermID = " & QuotedNormalised(tbTerminationUniqueID.Text)
        Dim dtAgentProfile As DataTable = ExecuteQueryToDataTable(sSQL)
        If dtAgentProfile.Rows.Count = 1 Then
            Dim drAgentProfile As DataRow = dtAgentProfile.Rows(0)
            tbTerminationName.Text = drAgentProfile("AgentName")
            tbTerminationContactName.Text = drAgentProfile("Contact")
            tbTerminationAddr1.Text = drAgentProfile("Address1")
            tbTerminationAddr2.Text = drAgentProfile("Address2")
            tbTerminationAddr3.Text = drAgentProfile("Address3")
            tbTerminationTownCity.Text = drAgentProfile("City")
            tbTerminationRegionPostCode.Text = drAgentProfile("Postcode")
            If tbTerminationRegionPostCode.Text.Trim = String.Empty Then
                ddlTerminationCountry.SelectedIndex = 3
            Else
                ddlTerminationCountry.SelectedIndex = 1
            End If
            If tbTerminationRegionPostCode.Text.Length >= 2 Then
                If tbTerminationRegionPostCode.Text.Substring(0, 2) = "BT" Then
                    ddlTerminationCountry.SelectedIndex = 2
                End If
            End If
            tbTerminationPhone.Text = drAgentProfile("PhoneNumber")
            tbTerminationEmail.Text = String.Empty
            LookupWURSAgent = True
        End If
    End Function

    Protected Function LookupWUIREAgent() As Boolean
        LookupWUIREAgent = False
        Dim sSQL As String = "SELECT * FROM ClientData_WUIRE_Agents WHERE TermID = " & QuotedNormalised(tbTerminationUniqueID.Text)
        Dim dtAgentProfile As DataTable = ExecuteQueryToDataTable(sSQL)
        If dtAgentProfile.Rows.Count = 1 Then
            Dim drAgentProfile As DataRow = dtAgentProfile.Rows(0)
            tbTerminationName.Text = drAgentProfile("AgentName")
            tbTerminationContactName.Text = String.Empty
            tbTerminationAddr1.Text = drAgentProfile("Address")
            tbTerminationTownCity.Text = drAgentProfile("City")
            tbTerminationRegionPostCode.Text = drAgentProfile("Region")
            ddlTerminationCountry.SelectedIndex = 3
            tbTerminationPhone.Text = String.Empty
            tbTerminationEmail.Text = String.Empty
            LookupWUIREAgent = True
        End If
    End Function

    Protected Function LookupCOSTAAgent() As Boolean
        LookupCOSTAAgent = False
        Dim sSQL As String = "SELECT * FROM ClientData_WUCOSTA_Agents WHERE TermID = " & QuotedNormalised(tbTerminationUniqueID.Text)
        Dim dtAgentProfile As DataTable = ExecuteQueryToDataTable(sSQL)
        If dtAgentProfile.Rows.Count = 1 Then
            Dim drAgentProfile As DataRow = dtAgentProfile.Rows(0)
            tbTerminationName.Text = drAgentProfile("LocationName")
            tbTerminationContactName.Text = String.Empty
            tbTerminationAddr1.Text = drAgentProfile("Address")
            tbTerminationAddr2.Text = String.Empty
            tbTerminationAddr3.Text = String.Empty
            tbTerminationTownCity.Text = String.Empty
            tbTerminationRegionPostCode.Text = drAgentProfile("PostCode")
            ddlTerminationCountry.SelectedIndex = 1
            tbTerminationPhone.Text = drAgentProfile("AreaCode") & " " & drAgentProfile("Phone")
            tbTerminationEmail.Text = String.Empty
            LookupCOSTAAgent = True
        End If
    End Function

    Protected Function LookupFININTAgent() As Boolean
        LookupFININTAgent = False
        Dim sSQL As String = "SELECT * FROM ClientData_WU_LegacyNetwork WHERE AccountNumber = " & QuotedNormalised(tbTerminationUniqueID.Text)
        Dim dtAgentProfile As DataTable = ExecuteQueryToDataTable(sSQL)
        If dtAgentProfile.Rows.Count = 1 Then
            Dim drAgentProfile As DataRow = dtAgentProfile.Rows(0)
            tbTerminationName.Text = drAgentProfile("LocationName")
            tbTerminationContactName.Text = String.Empty
            tbTerminationAddr1.Text = drAgentProfile("AddressLine1")
            tbTerminationAddr2.Text = drAgentProfile("AddressLine2")
            tbTerminationAddr3.Text = String.Empty
            tbTerminationTownCity.Text = drAgentProfile("CityName")
            tbTerminationRegionPostCode.Text = drAgentProfile("PostalCode")
            ddlTerminationCountry.SelectedIndex = 1
            tbTerminationPhone.Text = String.Empty
            tbTerminationEmail.Text = String.Empty
            LookupFININTAgent = True
        End If
        
    End Function

    Protected Sub btnTerminationCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If pbIsEditing Then
            pbIsEditing = False
            Call RefreshDetailPanel()
            pnlTerminationManagement.Visible = True
            pnlTerminationAddress.Visible = False
        Else
            Call DisplayMainPanel()
        End If
        pbGridDisabled = False
    End Sub

    Protected Sub btnTerminationSave_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        tbTerminationUniqueID.Text = tbTerminationUniqueID.Text.ToUpper
        If pbIsEditing Then
            Call SaveEditedTermination()
            pbIsEditing = False
            Call RefreshDetailPanel()
            pnlTerminationManagement.Visible = True
            pnlTerminationAddress.Visible = False
        Else
            Call SaveNewTermination()
        End If
        pbGridDisabled = False
    End Sub

    Protected Sub SaveNewTermination()
        Page.Validate()
        If Page.IsValid Then
            Dim sMessage As String = IsValidTermination()
            If sMessage = String.Empty Then
                Call SaveTermination()
                Call AddNewProduct()
                Call AuditEntry(ENTRY_TYPE_NEW_TERMINATION, "")
                ShowMessage("Termination created.")
                Call DisplayMainPanel()
            Else
                psTimerAction = "DUPLICATE_UNIQUEID"
                ShowMessage("ERROR - " & sMessage, nInterval:=INTERVAL_LONG_MESSAGE)
                tbTerminationUniqueID.Focus()
            End If
        Else
            ShowMessage("Please correct validation errors.", INTERVAL_SHORT_ERROR_MESSAGE)
            psTimerAction = "TERMINATION_VALIDATION"
            ' should work out first control that needs correction and set focus to it
            tbTerminationUniqueID.Focus()
        End If
    End Sub
    
    Protected Function GetCountryNameFromValue(ByVal nValue As Int32) As String
        GetCountryNameFromValue = "- not found -"
        For i = 1 To ddlTerminationCountry.Items.Count - 1
            If ddlTerminationCountry.Items(i).Value = nValue Then
                GetCountryNameFromValue = ddlTerminationCountry.Items(i).Text
                Exit For
            End If
        Next
    End Function
    
    Protected Sub SaveEditedTermination()
        Page.Validate()
        If Page.IsValid Then
            Dim sMessage As String = String.Empty
            Dim drTerminationDetails As DataRow = GetTerminationDetailsFromRecord()
            Call TrimTerminationFields()
            If tbTerminationUniqueID.Text <> drTerminationDetails("AgentUniqueID") Then
                sMessage = IsValidTermination()
            End If
            If sMessage = String.Empty Then
                
                Dim sOriginalValue As String = String.Empty
                Dim sNewValue As String = String.Empty
                Const BLANK_MESSAGE As String = "BLANK"
                
                If tbTerminationUniqueID.Text <> drTerminationDetails("AgentUniqueID") Then
                    sOriginalValue = drTerminationDetails("AgentUniqueID")
                    sNewValue = tbTerminationUniqueID.Text
                    If sOriginalValue.Trim = String.Empty Then
                        sOriginalValue = BLANK_MESSAGE
                    End If
                    If sNewValue.Trim = String.Empty Then
                        sNewValue = BLANK_MESSAGE
                    End If
                    Call AuditEntry(ENTRY_TYPE_DETAIL_CHANGE, "Agent Unique ID changed from " & sOriginalValue & " to " & sNewValue)
                End If

                If tbTerminationName.Text <> drTerminationDetails("AgentName") Then
                    sOriginalValue = drTerminationDetails("AgentName")
                    sNewValue = tbTerminationName.Text
                    If sOriginalValue.Trim = String.Empty Then
                        sOriginalValue = BLANK_MESSAGE
                    End If
                    If sNewValue.Trim = String.Empty Then
                        sNewValue = BLANK_MESSAGE
                    End If
                    Call AuditEntry(ENTRY_TYPE_DETAIL_CHANGE, "Agent Name changed from " & sOriginalValue & " to " & sNewValue)
                End If

                If tbTerminationContactName.Text <> drTerminationDetails("AgentContactName") Then
                    sOriginalValue = drTerminationDetails("AgentContactName")
                    sNewValue = tbTerminationContactName.Text
                    If sOriginalValue.Trim = String.Empty Then
                        sOriginalValue = BLANK_MESSAGE
                    End If
                    If sNewValue.Trim = String.Empty Then
                        sNewValue = BLANK_MESSAGE
                    End If
                    Call AuditEntry(ENTRY_TYPE_DETAIL_CHANGE, "Agent Contact Name changed from " & sOriginalValue & " to " & sNewValue)
                End If

                If tbTerminationAddr1.Text <> drTerminationDetails("AgentAddress1") Then
                    sOriginalValue = drTerminationDetails("AgentAddress1")
                    sNewValue = tbTerminationAddr1.Text
                    If sOriginalValue.Trim = String.Empty Then
                        sOriginalValue = BLANK_MESSAGE
                    End If
                    If sNewValue.Trim = String.Empty Then
                        sNewValue = BLANK_MESSAGE
                    End If
                    Call AuditEntry(ENTRY_TYPE_DETAIL_CHANGE, "Agent Address 1 changed from " & sOriginalValue & " to " & sNewValue)
                End If

                If tbTerminationAddr2.Text <> drTerminationDetails("AgentAddress2") Then
                    sOriginalValue = drTerminationDetails("AgentAddress2")
                    sNewValue = tbTerminationAddr2.Text
                    If sOriginalValue.Trim = String.Empty Then
                        sOriginalValue = BLANK_MESSAGE
                    End If
                    If sNewValue.Trim = String.Empty Then
                        sNewValue = BLANK_MESSAGE
                    End If
                    Call AuditEntry(ENTRY_TYPE_DETAIL_CHANGE, "Agent Address 2 changed from " & sOriginalValue & " to " & sNewValue)
                End If

                If tbTerminationAddr3.Text <> drTerminationDetails("AgentAddress3") Then
                    sOriginalValue = drTerminationDetails("AgentAddress3")
                    sNewValue = tbTerminationAddr3.Text
                    If sOriginalValue.Trim = String.Empty Then
                        sOriginalValue = BLANK_MESSAGE
                    End If
                    If sNewValue.Trim = String.Empty Then
                        sNewValue = BLANK_MESSAGE
                    End If
                    Call AuditEntry(ENTRY_TYPE_DETAIL_CHANGE, "Agent Address 3 changed from " & sOriginalValue & " to " & sNewValue)
                End If

                If tbTerminationTownCity.Text <> drTerminationDetails("AgentTownCity") Then
                    sOriginalValue = drTerminationDetails("AgentTownCity")
                    sNewValue = tbTerminationTownCity.Text
                    If sOriginalValue.Trim = String.Empty Then
                        sOriginalValue = BLANK_MESSAGE
                    End If
                    If sNewValue.Trim = String.Empty Then
                        sNewValue = BLANK_MESSAGE
                    End If
                    Call AuditEntry(ENTRY_TYPE_DETAIL_CHANGE, "Agent Town/City changed from " & sOriginalValue & " to " & sNewValue)
                End If

                If tbTerminationRegionPostCode.Text <> drTerminationDetails("AgentRegionOrPostCode") Then
                    sOriginalValue = drTerminationDetails("AgentRegionOrPostCode")
                    sNewValue = tbTerminationRegionPostCode.Text
                    If sOriginalValue.Trim = String.Empty Then
                        sOriginalValue = BLANK_MESSAGE
                    End If
                    If sNewValue.Trim = String.Empty Then
                        sNewValue = BLANK_MESSAGE
                    End If
                    Call AuditEntry(ENTRY_TYPE_DETAIL_CHANGE, "Agent Region / PostCode changed from " & sOriginalValue & " to " & sNewValue)
                End If

                Dim nAgentCountryKey As Int32 = drTerminationDetails("AgentCountryKey")
                If ddlTerminationCountry.SelectedValue <> nAgentCountryKey Then
                    Call AuditEntry(ENTRY_TYPE_DETAIL_CHANGE, "Agent Country changed from " & GetCountryNameFromValue(nAgentCountryKey) & " to " & ddlTerminationCountry.SelectedItem.Text)
                End If

                If tbTerminationEmail.Text <> drTerminationDetails("AgentEmail") Then
                    sOriginalValue = drTerminationDetails("AgentEmail")
                    sNewValue = tbTerminationEmail.Text
                    If sOriginalValue.Trim = String.Empty Then
                        sOriginalValue = BLANK_MESSAGE
                    End If
                    If sNewValue.Trim = String.Empty Then
                        sNewValue = BLANK_MESSAGE
                    End If
                    Call AuditEntry(ENTRY_TYPE_DETAIL_CHANGE, "Agent Email changed from " & drTerminationDetails("AgentEmail") & " to " & sNewValue)
                End If

                If tbTerminationPhone.Text <> drTerminationDetails("AgentPhone") Then
                    sOriginalValue = drTerminationDetails("AgentPhone")
                    sNewValue = tbTerminationPhone.Text
                    If sOriginalValue.Trim = String.Empty Then
                        sOriginalValue = BLANK_MESSAGE
                    End If
                    If sNewValue.Trim = String.Empty Then
                        sNewValue = BLANK_MESSAGE
                    End If
                    Call AuditEntry(ENTRY_TYPE_DETAIL_CHANGE, "Agent Phone Number changed from " & sOriginalValue & " to " & sNewValue)
                End If

                If tbTerminationComments.Text <> drTerminationDetails("Comments") Then
                    sOriginalValue = drTerminationDetails("Comments")
                    sNewValue = tbTerminationComments.Text
                    If sOriginalValue.Trim = String.Empty Then
                        sOriginalValue = BLANK_MESSAGE
                    End If
                    If sNewValue.Trim = String.Empty Then
                        sNewValue = BLANK_MESSAGE
                    End If
                    Call AuditEntry(ENTRY_TYPE_DETAIL_CHANGE, "Agent Comments changed from " & sOriginalValue & " to " & sNewValue)
                End If

                Call UpdateAgentTerminationDetails()
                ShowMessage("Saved changes.")
                Call RefreshDetailPanel()
            Else
                psTimerAction = "DUPLICATE_UNIQUEID"
                ShowMessage("ERROR - " & sMessage, nInterval:=INTERVAL_LONG_MESSAGE)
                tbTerminationUniqueID.Focus()
            End If
        Else
            ShowMessage("Please correct validation errors.", INTERVAL_SHORT_ERROR_MESSAGE)
            psTimerAction = "TERMINATION_VALIDATION"
            ' should work out first control that needs correction and set focus to it
            tbTerminationUniqueID.Focus()
        End If
    End Sub

    Protected Function IsValidTermination() As String
        IsValidTermination = String.Empty
        If TerminationExists(tbTerminationUniqueID.Text) Then
            IsValidTermination = "A termination record already exists for the Agent with Unique ID " & tbTerminationUniqueID.Text
        End If
    End Function
    
    Protected Function TerminationExists(ByVal sUniqueID As String) As Boolean
        TerminationExists = False
        Dim sSQL As String
        sSQL = "SELECT * FROM ClientData_WAT_Termination2 WHERE AgentUniqueID = '" & Normalised(sUniqueID) & "'"
        Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
        If dt.Rows.Count > 0 Then
            TerminationExists = True
        End If
    End Function

    Protected Function UpdateAgentTerminationDetails() As String
        UpdateAgentTerminationDetails = String.Empty
        tbTerminationUniqueID.Text = tbTerminationUniqueID.Text.ToUpper
        Dim sbSQL As New StringBuilder
        sbSQL.Append("UPDATE ClientData_WAT_Termination2 SET ")
        sbSQL.Append("AgentUniqueID")
        sbSQL.Append(" = ")
        sbSQL.Append(QuotedNormalised(tbTerminationUniqueID.Text))
        sbSQL.Append(",")
        sbSQL.Append("AgentName")
        sbSQL.Append(" = ")
        sbSQL.Append(QuotedNormalised(tbTerminationName.Text))
        sbSQL.Append(",")
        sbSQL.Append("AgentAddress1")
        sbSQL.Append(" = ")
        sbSQL.Append(QuotedNormalised(tbTerminationAddr1.Text))
        sbSQL.Append(",")
        sbSQL.Append("AgentAddress2")
        sbSQL.Append(" = ")
        sbSQL.Append(QuotedNormalised(tbTerminationAddr2.Text))
        sbSQL.Append(",")
        sbSQL.Append("AgentAddress3")
        sbSQL.Append(" = ")
        sbSQL.Append(QuotedNormalised(tbTerminationAddr3.Text))
        sbSQL.Append(",")
        sbSQL.Append("AgentTownCity")
        sbSQL.Append(" = ")
        sbSQL.Append(QuotedNormalised(tbTerminationTownCity.Text))
        sbSQL.Append(",")
        sbSQL.Append("AgentRegionOrPostCode")
        sbSQL.Append(" = ")
        sbSQL.Append(QuotedNormalised(tbTerminationRegionPostCode.Text))
        sbSQL.Append(",")
        sbSQL.Append("AgentCountryKey")
        sbSQL.Append(" = ")
        sbSQL.Append(ddlTerminationCountry.SelectedValue)
        sbSQL.Append(",")
        sbSQL.Append("AgentContactName")
        sbSQL.Append(" = ")
        sbSQL.Append(QuotedNormalised(tbTerminationContactName.Text))
        sbSQL.Append(",")
        sbSQL.Append("AgentPhone")
        sbSQL.Append(" = ")
        sbSQL.Append(QuotedNormalised(tbTerminationPhone.Text))
        sbSQL.Append(",")
        sbSQL.Append("AgentEmail")
        sbSQL.Append(" = ")
        sbSQL.Append(QuotedNormalised(tbTerminationEmail.Text))
        sbSQL.Append(",")
        sbSQL.Append("Comments")
        sbSQL.Append(" = ")
        sbSQL.Append(QuotedNormalised(tbTerminationComments.Text))  ' Comments 
        sbSQL.Append(" WHERE [id] = ")
        sbSQL.Append(pnTerminationKey.ToString)
        Call ExecuteQueryToDataTable(sbSQL.ToString)
    End Function

    Protected Sub TrimTerminationFields()
        tbTerminationUniqueID.Text = tbTerminationUniqueID.Text.Trim
        tbTerminationName.Text = tbTerminationName.Text.Trim
        tbTerminationContactName.Text = tbTerminationContactName.Text.Trim
        tbTerminationAddr1.Text = tbTerminationAddr1.Text.Trim
        tbTerminationAddr2.Text = tbTerminationAddr2.Text.Trim
        tbTerminationAddr3.Text = tbTerminationAddr3.Text.Trim
        tbTerminationTownCity.Text = tbTerminationTownCity.Text.Trim
        tbTerminationRegionPostCode.Text = tbTerminationRegionPostCode.Text.Trim
        tbTerminationEmail.Text = tbTerminationEmail.Text.Trim
        tbTerminationPhone.Text = tbTerminationPhone.Text.Trim
        tbTerminationComments.Text = tbTerminationComments.Text.Trim
    End Sub
    
    Protected Function SaveTermination() As String
        SaveTermination = String.Empty
        Dim nCustomerKey As Int32
        Dim dtTermination As DataTable
        Dim nUserKey As Int32
        Call TrimTerminationFields()
        If Session("CustomerKey") IsNot Nothing Then
            nCustomerKey = Session("CustomerKey")
        End If
        If Session("UserKey") IsNot Nothing Then
            nUserKey = Session("UserKey")
        End If
        tbTerminationUniqueID.Text = tbTerminationUniqueID.Text.ToUpper
        Dim sbSQL As New StringBuilder
        sbSQL.Append("INSERT INTO ClientData_WAT_Termination2 ")
        sbSQL.Append("(")
        sbSQL.Append("AgentUniqueID")
        sbSQL.Append(",")
        sbSQL.Append("AgentName")
        sbSQL.Append(",")
        sbSQL.Append("AgentAddress1")
        sbSQL.Append(",")
        sbSQL.Append("AgentAddress2")
        sbSQL.Append(",")
        sbSQL.Append("AgentAddress3")
        sbSQL.Append(",")
        sbSQL.Append("AgentTownCity")
        sbSQL.Append(",")
        sbSQL.Append("AgentRegionOrPostCode")
        sbSQL.Append(",")
        sbSQL.Append("AgentCountryKey")
        sbSQL.Append(",")
        sbSQL.Append("AgentContactName")
        sbSQL.Append(",")
        sbSQL.Append("AgentPhone")
        sbSQL.Append(",")
        sbSQL.Append("AgentEmail")
        sbSQL.Append(",")
        sbSQL.Append("CustomerKey")
        sbSQL.Append(",")
        sbSQL.Append("UserProfileKey")
        sbSQL.Append(",")
        sbSQL.Append("Comments")
        sbSQL.Append(",")
        sbSQL.Append("Collection01Ref")
        sbSQL.Append(",")
        ' sbSQL.Append("Collection01Date")  ' omitted
        ' sbSQL.Append(",")                 ' omitted
        sbSQL.Append("Collection01Notes")
        sbSQL.Append(",")
        sbSQL.Append("Collection01Cost")
        sbSQL.Append(",")
        sbSQL.Append("Collection01Failed")
        sbSQL.Append(",")
        sbSQL.Append("POSActivitiesCompleted")
        sbSQL.Append(",")
        sbSQL.Append("NextAction")
        sbSQL.Append(",")
        sbSQL.Append("TerminationClosed")
        sbSQL.Append(",")
        sbSQL.Append("CreatedOn")
        sbSQL.Append(",")
        sbSQL.Append("CreatedBy")
        sbSQL.Append(",")
        sbSQL.Append("LastChangedOn")
        sbSQL.Append(",")
        sbSQL.Append("LastChangedBy")
        sbSQL.Append(")")
        sbSQL.Append(" VALUES ")
        sbSQL.Append("(")
        sbSQL.Append(QuotedNormalised(tbTerminationUniqueID.Text))
        sbSQL.Append(",")
        sbSQL.Append(QuotedNormalised(tbTerminationName.Text))
        sbSQL.Append(",")
        sbSQL.Append(QuotedNormalised(tbTerminationAddr1.Text))
        sbSQL.Append(",")
        sbSQL.Append(QuotedNormalised(tbTerminationAddr2.Text))
        sbSQL.Append(",")
        sbSQL.Append(QuotedNormalised(tbTerminationAddr3.Text))
        sbSQL.Append(",")
        sbSQL.Append(QuotedNormalised(tbTerminationTownCity.Text))
        sbSQL.Append(",")
        sbSQL.Append(QuotedNormalised(tbTerminationRegionPostCode.Text))
        sbSQL.Append(",")
        sbSQL.Append(ddlTerminationCountry.SelectedValue)
        sbSQL.Append(",")
        sbSQL.Append(QuotedNormalised(tbTerminationContactName.Text))
        sbSQL.Append(",")
        sbSQL.Append(QuotedNormalised(tbTerminationPhone.Text))
        sbSQL.Append(",")
        sbSQL.Append(QuotedNormalised(tbTerminationEmail.Text))
        sbSQL.Append(",")
        sbSQL.Append(nCustomerKey.ToString)
        sbSQL.Append(",")
        sbSQL.Append(nUserKey.ToString)
        sbSQL.Append(",")
        sbSQL.Append(QuotedNormalised(tbTerminationComments.Text))  ' Comments 
        sbSQL.Append(",")
        sbSQL.Append(Quoted(String.Empty))  ' Collection01Ref
        sbSQL.Append(",")
        sbSQL.Append(Quoted(String.Empty))  ' Collection01Notes
        sbSQL.Append(",")
        sbSQL.Append("0")  ' Collection01Cost
        sbSQL.Append(",")
        sbSQL.Append("0")  'Collection01Failed
        sbSQL.Append(",")
        sbSQL.Append(Quoted(String.Empty))  ' POSActivitiesCompleted
        sbSQL.Append(",")
        sbSQL.Append(Quoted("01CONTACT AGENT"))  ' NextAction
        sbSQL.Append(",")
        sbSQL.Append("0")  ' TerminationClosed
        sbSQL.Append(",")
        sbSQL.Append("GETDATE()")  ' LastChangedOn
        sbSQL.Append(",")
        sbSQL.Append(nUserKey.ToString)  ' LastChangedBy
        sbSQL.Append(",")
        sbSQL.Append("GETDATE()")  ' LastChangedOn
        sbSQL.Append(",")
        sbSQL.Append(nUserKey.ToString)  ' LastChangedBy
        sbSQL.Append(") ")
        sbSQL.Append("SELECT SCOPE_IDENTITY()")
        dtTermination = ExecuteQueryToDataTable(sbSQL.ToString)
        If dtTermination.Rows.Count = 1 Then
            pnTerminationKey = dtTermination.Rows(0).Item(0)
        Else
            Call DisplayFatalError("Expected one row to be returned from new termination creation but " & dtTermination.Rows.Count & " rows were returned.")
            pnTerminationKey = 0
        End If
    End Function
    
    Protected Sub InitMainPanel()
        Call BindTerminationsGrid(String.Empty)
    End Sub
    
    Protected Function Normalise(ByVal s As String) As String
        Normalise = s.Replace("'", "''")
    End Function
    
    Protected Sub BindTerminationsGrid(ByVal sSearchString As String)
        Dim dtTerminations As DataTable = Nothing
        Dim sSQL As String
        Dim sModifier As String = String.Empty
        sSearchString = Normalise(sSearchString.Trim)
        If (sSearchString <> String.Empty) Or cbIgnoreCompletedTerminations.Checked Then
            sModifier = "WHERE "
        End If
        If sSearchString <> String.Empty Then
            sModifier &= "(AgentUniqueID LIKE '%" & sSearchString & "%' OR AgentName LIKE '%" & sSearchString & "%' OR AgentAddress1 LIKE '%" & sSearchString & "%' OR AgentTownCity LIKE '%" & sSearchString & "%' OR AgentRegionOrPostCode LIKE '%" & sSearchString & "%' OR AgentContactName LIKE '%" & sSearchString & "%') "
        End If
        If cbIgnoreCompletedTerminations.Checked Then
            If sModifier <> "WHERE " Then
                sModifier &= " AND "
            End If
            sModifier &= " TerminationClosed = 0 "
        End If
        sSQL = "SELECT [id], CAST(REPLACE(CONVERT(VARCHAR(11),  CreatedOn, 106), ' ', '-') AS varchar(20)) + ' ' + SUBSTRING((CONVERT(VARCHAR(8), CreatedOn, 108)),1,5) 'CreatedOn', SUBSTRING(NextAction, 3, 999) 'NextAction', AgentUniqueID 'UniqueID', AgentName 'Agent', AgentName + ' ' + AgentAddress1 + ' ' + AgentTownCity 'AgentSummary', TerminationClosed FROM ClientData_WAT_Termination2 " & sModifier & " ORDER BY " & psTerminationsSortExpression
        dtTerminations = ExecuteQueryToDataTable(sSQL)
        gvTerminations.DataSource = dtTerminations
        gvTerminations.DataBind()
    End Sub
    
    Protected Sub InitNewTerminationPanel()
        tbTerminationUniqueID.Text = String.Empty
        tbTerminationName.Text = String.Empty
        tbTerminationContactName.Text = String.Empty
        tbTerminationAddr1.Text = String.Empty
        tbTerminationAddr2.Text = String.Empty
        tbTerminationAddr3.Text = String.Empty
        tbTerminationTownCity.Text = String.Empty
        tbTerminationRegionPostCode.Text = String.Empty
        Call InitTerminationCountryDropDown()
        tbTerminationEmail.Text = String.Empty
        tbTerminationPhone.Text = String.Empty
        tbTerminationComments.Text = String.Empty
        lblLegendAgentTermination.Text = "New Agent Termination"
        tbTerminationUniqueID.Focus()
    End Sub

    Protected Sub InitTerminationCountryDropDown()
        ddlTerminationCountry.Items.Clear()
        ddlTerminationCountry.Items.Add(New ListItem("- please select -", 0))
        ddlTerminationCountry.Items.Add(New ListItem("UK (excluding N.IRELAND)", COUNTRY_UK_EXCLUDING_NORTHERN_IRELAND))
        ddlTerminationCountry.Items.Add(New ListItem("UK (N.IRELAND)", COUNTRY_NORTHERN_IRELAND))
        ddlTerminationCountry.Items.Add(New ListItem("IRISH REPUBLIC", COUNTRY_IRISH_REPUBLIC))
    End Sub

    Protected Function Quoted(ByVal s As String) As String
        Quoted = "'" & s & "'"
    End Function

    Protected Function Normalised(ByVal s As String) As String
        Normalised = s.Replace("'", "''")
    End Function
   
    Protected Function QuotedNormalised(ByVal s As String) As String
        QuotedNormalised = "'" & Normalised(s) & "'"
    End Function
   
    Protected Sub lnkbtnNewTerminationUK_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        For i As Int32 = 0 To ddlTerminationCountry.Items.Count - 1
            If ddlTerminationCountry.Items(i).Value = COUNTRY_UK_EXCLUDING_NORTHERN_IRELAND Then
                ddlTerminationCountry.SelectedIndex = i
                Exit For
            End If
        Next
    End Sub

    Protected Sub lnkbtnNewTerminationIreland_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        For i As Int32 = 0 To ddlTerminationCountry.Items.Count - 1
            If ddlTerminationCountry.Items(i).Value = COUNTRY_IRISH_REPUBLIC Then
                ddlTerminationCountry.SelectedIndex = i
                Exit For
            End If
        Next
    End Sub

    Protected Sub lnkbtnNewTerminationNorthernIreland_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        For i As Int32 = 0 To ddlTerminationCountry.Items.Count - 1
            If ddlTerminationCountry.Items(i).Value = COUNTRY_NORTHERN_IRELAND Then
                ddlTerminationCountry.SelectedIndex = i
                Exit For
            End If
        Next
    End Sub
    
    Protected Sub gvTerminations_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs)
        'If e.Row.RowType = DataControlRowType.DataRow Then
        '    e.Row.Attributes.Add("onclick", Page.ClientScript.GetPostBackEventReference(sender, "Select$" & e.Row.RowIndex.ToString))
        'End If
        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim gvr As GridViewRow = e.Row
            Dim hidTerminationClosed As HiddenField = gvr.FindControl("hidTerminationClosed")
            Dim lblCreatedOn As Label = gvr.FindControl("lblCreatedOn")
            If hidTerminationClosed.Value = True Then
                'gvr.BackColor = Drawing.Color.Silver
                lblCreatedOn.Text &= Colour(Bold(" (CLOSED)"), "green")
            End If
        End If
    End Sub

    Protected Sub gvTerminations_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        ' this is currently called by the grid
    End Sub

    Protected Sub myGridView1_RowClicked(ByVal sender As Object, ByVal e As SampleControls.GridViewRowClickedEventArgs)
        If Not pbGridDisabled Then
            Dim gvrcea As GridViewRowClickedEventArgs = e
            Dim gvr As GridViewRow = gvrcea.Row
            Dim hidID As HiddenField
            hidID = gvr.Cells(0).FindControl("hidID")
            pnTerminationKey = hidID.Value
            pnlTerminationAddress.Visible = False
            Dim sAgentDetails As String = gvr.Cells(2).Text & " - " & gvr.Cells(3).Text
            lblTerminationIdentifier.Text = sAgentDetails
            If pnPermission = PERMISSION_ALL_2 Then
                Call ButtonPanelVisibility(True)
            End If
            Call RefreshDetailPanel()
        End If
    End Sub

    Protected Sub gvTerminations_Sorting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewSortEventArgs)
        psTerminationsSortExpression = e.SortExpression
        Select Case e.SortExpression.ToLower
            Case "createdon"
                psTerminationsSortExpression = "CreatedOn" & psCreatedOnSortDirection
                If psCreatedOnSortDirection = " ASC" Then
                    psCreatedOnSortDirection = " DESC"
                    lblSortInfo.Text = "sorted on <b>Created On</b> in <b>ascending</b> order"
                Else
                    psCreatedOnSortDirection = " ASC"
                    lblSortInfo.Text = "sorted on <b>Created On</b> in <b>descending</b> order"
                End If
            Case "nextaction"
                psTerminationsSortExpression = "NextAction" & psNextActionSortDirection
                If psNextActionSortDirection = " ASC" Then
                    psNextActionSortDirection = " DESC"
                    lblSortInfo.Text = "sorted on <b>Next Action</b> in <b>ascending</b> order"
                Else
                    psNextActionSortDirection = " ASC"
                    lblSortInfo.Text = "sorted on <b>Next Action</b> in <b>descending</b> order"
                End If
            Case "uniqueid"
                psTerminationsSortExpression = "AgentUniqueId" & psUniqueIDSortDirection
                If psUniqueIDSortDirection = " ASC" Then
                    psUniqueIDSortDirection = " DESC"
                    lblSortInfo.Text = "sorted on <b>Unique ID</b> in <b>ascending</b> order"
                Else
                    psUniqueIDSortDirection = " ASC"
                    lblSortInfo.Text = "sorted on <b>Unique ID</b> in <b>descending</b> order"
                End If
            Case "agentsummary"
                psTerminationsSortExpression = "AgentName" & psAgentSortDirection
                If psAgentSortDirection = " ASC" Then
                    psAgentSortDirection = " DESC"
                    lblSortInfo.Text = "sorted on <b>Agent</b> in <b>ascending</b> order"
                Else
                    psAgentSortDirection = " ASC"
                    lblSortInfo.Text = "sorted on <b>Agent</b> in <b>descending</b> order"
                End If
        End Select
        lblSortInfo.ToolTip = lblSortInfo.Text.Replace("<b>", "").Replace("</b>", "")
        Call BindTerminationsGrid(psSearchTerm)
    End Sub

    Protected Sub gvTerminations_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs)
        gvTerminations.PageIndex = e.NewPageIndex
        Call BindTerminationsGrid(tbSearchTerminations.Text)
    End Sub

    Protected Function GetTerminationDetailsFromRecord() As DataRow
        Dim sSQL As String
        sSQL = "SELECT CAST(REPLACE(CONVERT(VARCHAR(11),  CreatedOn, 106), ' ', '-') AS varchar(20)) + ' ' + SUBSTRING((CONVERT(VARCHAR(8), CreatedOn, 108)),1,5) 'CreatedOnFormatted', SUBSTRING(NextAction, 3, 999) 'NextActionFormatted', CAST(REPLACE(CONVERT(VARCHAR(11),  Collection01Date, 106), ' ', '-') AS varchar(20)) 'CollectionDateFormatted', * FROM ClientData_WAT_Termination2 WHERE [id] = " & pnTerminationKey
        Dim dtTerminationDetails As DataTable = ExecuteQueryToDataTable(sSQL)
        GetTerminationDetailsFromRecord = dtTerminationDetails.Rows(0)
    End Function
    
    Protected Sub PopulateTerminationDetails()
        Dim drTerminationDetails As DataRow = GetTerminationDetailsFromRecord()
        Dim sbTerminationDetails As New StringBuilder
        
        reTerminationDetails.Content = String.Empty

        sbTerminationDetails.Append("<table style=""width:100%"">")

        sbTerminationDetails.Append("<tr>")
        sbTerminationDetails.Append("<td style=""width:50%"">")
        sbTerminationDetails.Append("</td>")
        sbTerminationDetails.Append("<td style=""width:50%"">")
        sbTerminationDetails.Append("</td>")
        sbTerminationDetails.Append("</tr>")
        
        sbTerminationDetails.Append("<tr>")
        sbTerminationDetails.Append("<td>")
        sbTerminationDetails.Append("Created:")
        sbTerminationDetails.Append("</td>")
        sbTerminationDetails.Append("<td>")
        sbTerminationDetails.Append(Bold(drTerminationDetails("CreatedOnFormatted")))
        sbTerminationDetails.Append("</td>")
        sbTerminationDetails.Append("</tr>")
        
        sbTerminationDetails.Append("<tr>")
        sbTerminationDetails.Append("<td>")
        sbTerminationDetails.Append("Next Action:")
        sbTerminationDetails.Append("</td>")
        sbTerminationDetails.Append("<td>")
        sbTerminationDetails.Append(Bold(drTerminationDetails("NextActionFormatted")))
        sbTerminationDetails.Append("</td>")
        sbTerminationDetails.Append("</tr>")
        
        sbTerminationDetails.Append("<tr>")
        sbTerminationDetails.Append("<td style=""width:50%"">")
        sbTerminationDetails.Append("&nbsp;")
        sbTerminationDetails.Append("</td>")
        sbTerminationDetails.Append("<td style=""width:50%"">")
        sbTerminationDetails.Append("&nbsp;")
        sbTerminationDetails.Append("</td>")
        sbTerminationDetails.Append("</tr>")

        sbTerminationDetails.Append("<tr>")
        sbTerminationDetails.Append("<td>")
        sbTerminationDetails.Append("Unique ID:")
        sbTerminationDetails.Append("</td>")
        sbTerminationDetails.Append("<td>")
        sbTerminationDetails.Append(Bold(drTerminationDetails("AgentUniqueID")))
        sbTerminationDetails.Append("</td>")
        sbTerminationDetails.Append("</tr>")
        
        sbTerminationDetails.Append("<tr>")
        sbTerminationDetails.Append("<td style=""width:50%"">")
        sbTerminationDetails.Append("Agent Name:")
        sbTerminationDetails.Append("</td>")
        sbTerminationDetails.Append("<td style=""width:50%"">")
        sbTerminationDetails.Append(Bold(drTerminationDetails("AgentName")))
        sbTerminationDetails.Append("</td>")
        sbTerminationDetails.Append("</tr>")

        sbTerminationDetails.Append("<tr>")
        sbTerminationDetails.Append("<td style=""width:50%"">")
        sbTerminationDetails.Append("Contact Name:")
        sbTerminationDetails.Append("</td>")
        sbTerminationDetails.Append("<td style=""width:50%"">")
        sbTerminationDetails.Append(Bold(drTerminationDetails("AgentContactName")))
        sbTerminationDetails.Append("</td>")
        sbTerminationDetails.Append("</tr>")

        sbTerminationDetails.Append("<tr>")
        sbTerminationDetails.Append("<td style=""width:50%"">")
        sbTerminationDetails.Append("Addr 1:")
        sbTerminationDetails.Append("</td>")
        sbTerminationDetails.Append("<td style=""width:50%"">")
        sbTerminationDetails.Append(Bold(drTerminationDetails("AgentAddress1")))
        sbTerminationDetails.Append("</td>")
        sbTerminationDetails.Append("</tr>")

        If drTerminationDetails("AgentAddress2").ToString.Trim <> String.Empty Then
            sbTerminationDetails.Append("<tr>")
            sbTerminationDetails.Append("<td style=""width:50%"">")
            sbTerminationDetails.Append("Addr 2:")
            sbTerminationDetails.Append("</td>")
            sbTerminationDetails.Append("<td style=""width:50%"">")
            sbTerminationDetails.Append(Bold(drTerminationDetails("AgentAddress2")))
            sbTerminationDetails.Append("</td>")
            sbTerminationDetails.Append("</tr>")
        End If

        If drTerminationDetails("AgentAddress3").ToString.Trim <> String.Empty Then
            sbTerminationDetails.Append("<tr>")
            sbTerminationDetails.Append("<td style=""width:50%"">")
            sbTerminationDetails.Append("Addr 3:")
            sbTerminationDetails.Append("</td>")
            sbTerminationDetails.Append("<td style=""width:50%"">")
            sbTerminationDetails.Append(Bold(drTerminationDetails("AgentAddress3")))
            sbTerminationDetails.Append("</td>")
            sbTerminationDetails.Append("</tr>")
        End If

        sbTerminationDetails.Append("<tr>")
        sbTerminationDetails.Append("<td style=""width:50%"">")
        sbTerminationDetails.Append("Town/City:")
        sbTerminationDetails.Append("</td>")
        sbTerminationDetails.Append("<td style=""width:50%"">")
        sbTerminationDetails.Append(Bold(drTerminationDetails("AgentTownCity")))
        sbTerminationDetails.Append("</td>")
        sbTerminationDetails.Append("</tr>")

        sbTerminationDetails.Append("<tr>")
        sbTerminationDetails.Append("<td style=""width:50%"">")
        sbTerminationDetails.Append("Region/Postcode:")
        sbTerminationDetails.Append("</td>")
        sbTerminationDetails.Append("<td style=""width:50%"">")
        sbTerminationDetails.Append(Bold(drTerminationDetails("AgentRegionOrPostCode")))
        sbTerminationDetails.Append("</td>")
        sbTerminationDetails.Append("</tr>")

        sbTerminationDetails.Append("<tr>")
        sbTerminationDetails.Append("<td style=""width:50%"">")
        sbTerminationDetails.Append("Country:")
        sbTerminationDetails.Append("</td>")
        sbTerminationDetails.Append("<td style=""width:50%"">")
        sbTerminationDetails.Append(Bold(GetCountryNameFromCountryCode(drTerminationDetails("AgentCountryKey"))))
        sbTerminationDetails.Append("</td>")
        sbTerminationDetails.Append("</tr>")

        sbTerminationDetails.Append("<tr>")
        sbTerminationDetails.Append("<td style=""width:50%"">")
        sbTerminationDetails.Append("Telephone:")
        sbTerminationDetails.Append("</td>")
        sbTerminationDetails.Append("<td style=""width:50%"">")
        sbTerminationDetails.Append(Bold(drTerminationDetails("AgentPhone")))
        sbTerminationDetails.Append("</td>")
        sbTerminationDetails.Append("</tr>")

        sbTerminationDetails.Append("<tr>")
        sbTerminationDetails.Append("<td style=""width:50%"">")
        sbTerminationDetails.Append("Email:")
        sbTerminationDetails.Append("</td>")
        sbTerminationDetails.Append("<td style=""width:50%"">")
        sbTerminationDetails.Append(Bold(drTerminationDetails("AgentEmail")))
        sbTerminationDetails.Append("</td>")
        sbTerminationDetails.Append("</tr>")

        sbTerminationDetails.Append("<tr>")
        sbTerminationDetails.Append("<td style=""width:50%"">")
        sbTerminationDetails.Append("&nbsp;")
        sbTerminationDetails.Append("</td>")
        sbTerminationDetails.Append("<td style=""width:50%"">")
        sbTerminationDetails.Append("&nbsp;")
        sbTerminationDetails.Append("</td>")
        sbTerminationDetails.Append("</tr>")

        sbTerminationDetails.Append("<tr>")
        sbTerminationDetails.Append("<td style=""width:50%"">")
        sbTerminationDetails.Append("Comments:")
        sbTerminationDetails.Append("</td>")
        sbTerminationDetails.Append("<td style=""width:50%"">")
        sbTerminationDetails.Append(Bold(drTerminationDetails("Comments")))
        sbTerminationDetails.Append("</td>")
        sbTerminationDetails.Append("</tr>")

        sbTerminationDetails.Append("<tr>")
        sbTerminationDetails.Append("<td style=""width:50%"">")
        sbTerminationDetails.Append("&nbsp;")
        sbTerminationDetails.Append("</td>")
        sbTerminationDetails.Append("<td style=""width:50%"">")
        sbTerminationDetails.Append("&nbsp;")
        sbTerminationDetails.Append("</td>")
        sbTerminationDetails.Append("</tr>")

        sbTerminationDetails.Append("<tr>")
        sbTerminationDetails.Append("<td style=""width:50%"">")
        sbTerminationDetails.Append("Collection Date:")
        sbTerminationDetails.Append("</td>")
        sbTerminationDetails.Append("<td style=""width:50%"">")
        If drTerminationDetails("CollectionDateFormatted").ToString <> String.Empty Then
            sbTerminationDetails.Append(Bold(drTerminationDetails("CollectionDateFormatted")))
        Else
            sbTerminationDetails.Append(Bold("NONE"))
        End If
        sbTerminationDetails.Append("</td>")
        sbTerminationDetails.Append("</tr>")
        
        sbTerminationDetails.Append("</table>")

        reTerminationDetails.Content = sbTerminationDetails.ToString
        
        If drTerminationDetails("TerminationClosed") = True Then
            btnComplete.Text = "Mark as Re-opened"
        Else
            btnComplete.Text = "Mark as Complete"
        End If
        
    End Sub

    Protected Function GetCountryNameFromCountryCode(ByVal nCountryCode As Int32) As String
        GetCountryNameFromCountryCode = String.Empty
        Select Case nCountryCode
            Case COUNTRY_UK_EXCLUDING_NORTHERN_IRELAND
                GetCountryNameFromCountryCode = "UK (EXCL N.IRELAND)"
            Case COUNTRY_NORTHERN_IRELAND
                GetCountryNameFromCountryCode = "N.IRELAND"
            Case COUNTRY_IRISH_REPUBLIC
                GetCountryNameFromCountryCode = "IRISH REPUBLIC"
            Case Else
                GetCountryNameFromCountryCode = "UNKNOWN COUNTRY - PLEASE INFORM DEVELOPMENT"
        End Select

    End Function
    
    Protected Function LineBreak() As String
        LineBreak = "<br />" & Environment.NewLine
    End Function
    
    Protected Function Bold(ByVal s As String) As String
        Bold = "<b>" & s & "</b>"
    End Function

    Protected Function Size(ByVal s As String, ByVal nFontSize As Int32) As String
        Size = "<span style=""font-size:" & nFontSize.ToString & """>" & s & "</span>"
    End Function
    
    Protected Function Italic(ByVal s As String) As String
        Italic = "<span style=""font-style:italic""" & s & "</span>"
    End Function

    Protected Function Underline(ByVal s As String) As String
        Underline = "<span style=""text-decoration:underline""" & s & "</span>"
    End Function

    Protected Function Colour(ByVal s As String, ByVal sColourValue As String) As String
        Colour = "<span style=""color:" & sColourValue & """>" & s & "</span>"
    End Function

    Protected Sub AuditEntry(ByVal sEntryType As String, ByVal sEntryText As String)
        Dim sbSQL As New StringBuilder
        Dim nUserKey As Int32
        If Session("UserKey") IsNot Nothing Then
            nUserKey = Session("UserKey")
        End If

        sbSQL.Append("INSERT INTO ClientData_WAT_AuditTrail2 ")
        sbSQL.Append("(")
        sbSQL.Append("TerminationKey")
        sbSQL.Append(",")
        sbSQL.Append("EntryType")
        sbSQL.Append(",")
        sbSQL.Append("EntryText")
        sbSQL.Append(",")
        sbSQL.Append("LastChangedBy")
        sbSQL.Append(",")
        sbSQL.Append("LastChangedOn")
        sbSQL.Append(") ")
        sbSQL.Append("VALUES")
        sbSQL.Append(" (")

        sbSQL.Append(pnTerminationKey.ToString)
        sbSQL.Append(",")
        
        sbSQL.Append(QuotedNormalised(sEntryType))
        sbSQL.Append(",")
        sbSQL.Append(QuotedNormalised(sEntryText))
        sbSQL.Append(",")
        sbSQL.Append(nUserKey.ToString)
        sbSQL.Append(",")
        sbSQL.Append("GETDATE()")
        sbSQL.Append(")")
        Call ExecuteQueryToDataTable(sbSQL.ToString)
        
        Dim sSQL As String
        Dim sNotificationType As String = String.Empty
        Dim sTitle As String = String.Empty
        Dim sBodyText As String = String.Empty
        Dim drTerminationDetails As DataRow = GetTerminationDetailsFromRecord()

        Select Case sEntryType
            Case ENTRY_TYPE_NEW_TERMINATION
                sNotificationType = "Termination_CREATE"
                sTitle = "AGENT TERMINATION notification (new termination) - agent " & drTerminationDetails("AgentUniqueID")
                sBodyText = "Termination commenced for agent with unique ID " & drTerminationDetails("AgentUniqueID") & "."
            Case ENTRY_TYPE_DETAIL_CHANGE
                sNotificationType = "Termination_NOTEDETAILCHANGE"
                sTitle = "AGENT TERMINATION notification (note or detail change) - agent " & drTerminationDetails("AgentUniqueID")
                sBodyText = sEntryText
            Case ENTRY_TYPE_EVENT, ENTRY_TYPE_NOTE
                sNotificationType = "Termination_EVENT"
                sTitle = "AGENT TERMINATION notification (event or note) - agent " & drTerminationDetails("AgentUniqueID")
                sBodyText = sEntryText
            Case ENTRY_TYPE_TERMINATION_COMPLETION_STATUS
                sNotificationType = "Termination_COMPLETE"
                sTitle = "AGENT TERMINATION notification (completion) - agent " & drTerminationDetails("AgentUniqueID")
                sBodyText = "Termination COMPLETED for agent with unique ID " & drTerminationDetails("AgentUniqueID") & "."
            Case Else
                Call DisplayFatalError("Unrecognised audit entry type " & sEntryType)
        End Select
        sSQL = "SELECT * FROM ClientData_WAT_Notifications2 wn INNER JOIN UserProfile up ON wn.UserKey = up.[key] WHERE " & sNotificationType & " = 1"
        Dim dtNotifications As DataTable = ExecuteQueryToDataTable(sSQL)
        For Each drNotification As DataRow In dtNotifications.Rows
            Call SendMail("WUTERMINATION_EVENT", drNotification("EmailAddr"), sTitle, sBodyText, sBodyText)
        Next
    End Sub

    Protected Sub PopulateAuditTrail()
        Dim sSQL As String
        'sSQL = "SELECT * FROM ClientData_WAT_AuditTrail WHERE TerminationKey = " & sTerminationID & " ORDER BY [id]"
        Dim sRestrictionClause As String = String.Empty
        If rbAuditShowMajorEvents.Checked Then
            sRestrictionClause = " AND ((EntryType = 'NEW TERMINATION' OR EntryType = 'EVENT' OR EntryType LIKE '%TERMINATION%') OR EntryType = 'NOTE') "
        End If
        sSQL = "SELECT CAST(REPLACE(CONVERT(VARCHAR(11),  LastChangedOn, 106), ' ', '-') AS varchar(20)) + ' ' + SUBSTRING((CONVERT(VARCHAR(8), LastChangedOn, 108)),1,5) 'LastChangedOnFormatted', wat.*, ISNULL(up.UserID, 'not recorded') 'UserID' FROM ClientData_WAT_AuditTrail2 wat LEFT OUTER JOIN UserProfile up ON wat.LastChangedBy = up.[key] WHERE TerminationKey = " & pnTerminationKey & sRestrictionClause & " ORDER BY [id]"
        Dim dtAuditTrail As DataTable = ExecuteQueryToDataTable(sSQL)
        reAuditTrail.Content = String.Empty
        
        Dim sbAuditTrail As New StringBuilder
        For Each drAuditEntry As DataRow In dtAuditTrail.Rows
            sbAuditTrail.Append(drAuditEntry("LastChangedOnFormatted"))
            sbAuditTrail.Append("&nbsp;&nbsp;&nbsp;")
            sbAuditTrail.Append(Colour(Bold(drAuditEntry("EntryType")), "blue"))
            sbAuditTrail.Append("&nbsp;&nbsp;&nbsp;")
            sbAuditTrail.Append(drAuditEntry("EntryText"))
            sbAuditTrail.Append("&nbsp;&nbsp;&nbsp;")
            sbAuditTrail.Append("user: ")
            sbAuditTrail.Append(drAuditEntry("UserID"))
            sbAuditTrail.Append("<br />")
            sbAuditTrail.Append(Environment.NewLine)
        Next
        reAuditTrail.Content = sbAuditTrail.ToString
    End Sub

    Protected Sub SendMail(ByVal sType As String, ByVal sRecipient As String, ByVal sSubject As String, ByVal sBodyText As String, ByVal sBodyHTML As String)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Email_AddToQueue", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Try
            oCmd.Parameters.Add(New SqlParameter("@MessageId", SqlDbType.NVarChar, 20))
            oCmd.Parameters("@MessageId").Value = sType
    
            oCmd.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
            oCmd.Parameters("@CustomerKey").Value = CUSTOMER_WUCOLL
    
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
            Call WebMsgBox.Show("Error in SendMail: " & ex.Message)
            Call ShowMessage("Error in SendMail: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub

    Protected Sub btnSearchTerminations_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call BindTerminationsGrid(tbSearchTerminations.Text)
    End Sub
    
    Protected Function sGetEmptyDataMessage() As String
        tbSearchTerminations.Text = tbSearchTerminations.Text.Trim
        If tbSearchTerminations.Text = String.Empty Then
            sGetEmptyDataMessage = "No terminations found."
        Else
            sGetEmptyDataMessage = "No terminations matching the search term '" & tbSearchTerminations.Text & "' found."
        End If
    End Function
    
    Protected Sub lnkbtnClearSearchTerminationsBox_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        tbSearchTerminations.Text = String.Empty
        Call BindTerminationsGrid(String.Empty)
    End Sub
    
    Protected Sub cbIgnoreCompletedTerminations_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call BindTerminationsGrid(tbSearchTerminations.Text)
    End Sub
    
    Protected Sub btnEditAgentDetails_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        'Call HideAllPanels()
        Call InitTerminationAddressPanelFromTerminationRecord()
        pnlTerminationManagement.Visible = False
        pnlTerminationAddress.Visible = True
        pbIsEditing = True
        pbGridDisabled = True
    End Sub

    Protected Sub InitTerminationAddressPanelFromTerminationRecord()
        Dim sSQL As String
        sSQL = "SELECT * FROM ClientData_WAT_Termination2 WHERE [id] = " & pnTerminationKey
        Dim drTerminationDetails As DataRow = ExecuteQueryToDataTable(sSQL).Rows(0)
        tbTerminationUniqueID.Text = drTerminationDetails("AgentUniqueID")
        tbTerminationName.Text = drTerminationDetails("AgentName")
        tbTerminationContactName.Text = drTerminationDetails("AgentContactName")
        tbTerminationAddr1.Text = drTerminationDetails("AgentAddress1")
        tbTerminationAddr2.Text = drTerminationDetails("AgentAddress2")
        tbTerminationAddr3.Text = drTerminationDetails("AgentAddress3")
        tbTerminationTownCity.Text = drTerminationDetails("AgentTownCity")
        tbTerminationRegionPostCode.Text = drTerminationDetails("AgentRegionOrPostCode")
        Dim nAgentCountryKey As Int32 = drTerminationDetails("AgentCountryKey")
        For i As Int32 = 0 To ddlTerminationCountry.Items.Count - 1
            If ddlTerminationCountry.Items(i).Value = nAgentCountryKey Then
                ddlTerminationCountry.SelectedIndex = i
                Exit For
            End If
        Next
        tbTerminationEmail.Text = drTerminationDetails("AgentEmail")
        tbTerminationPhone.Text = drTerminationDetails("AgentPhone")
        tbTerminationComments.Text = drTerminationDetails("Comments")
        lblLegendAgentTermination.Text = "Edit Agent Termination Details"
        tbTerminationUniqueID.Focus()
    
    End Sub
    
    Protected Sub RefreshDetailPanel()
        pnlTerminationManagement.Visible = True
        Call PopulateTerminationDetails()
        Call PopulateAuditTrail()
        If pnPermission = PERMISSION_CREATE_EDIT_ENTRY_1 Then   ' DON'T make this button available once an attempt has been made to contact Agent
            btnEditAgentDetails.Visible = True
        End If
        Call SetButtonAdvisories()
    End Sub
    
    Protected Sub ddlDisplayCount_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        gvTerminations.PageSize = ddlDisplayCount.SelectedValue
        gvTerminations.PageIndex = 0
        Call BindTerminationsGrid(tbSearchTerminations.Text)
    End Sub
    
    Protected Function CodeStrippedNextAction(sNextAction As String) As String
        If IsNumeric(sNextAction.Substring(0, 2)) Then
            CodeStrippedNextAction = sNextAction.Substring(2, sNextAction.Length - 2)
        Else
            CodeStrippedNextAction = sNextAction
        End If
    End Function
    
    Protected Sub btnContactAgentCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call HideDialogPanels()
        pnlContactAgent.Visible = False
        Call ButtonPanelVisibility(True)
        pbGridDisabled = False
    End Sub
    
    Protected Sub btnGeneralCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        pnlContactAgent.Visible = False
        pnlFormsReceived.Visible = False
        pnlFormsSent.Visible = False
        pnlAddNote.Visible = False
        Call ButtonPanelVisibility(True)
        pbGridDisabled = False
        Call RefreshDetailPanel
    End Sub
    
    Protected Sub btnAddNote_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        tbNote.Text = String.Empty
        pnlAddNote.Visible = True
        tbNote.Focus()
    End Sub
    
    Protected Sub ButtonPanelVisibility(ByVal bVisible As Boolean)
        btnEditAgentDetails.Visible = bVisible
        btnContactAgent.Visible = bVisible
        btnRcvdAtTransworld.Visible = bVisible
        btnSentToWesternUnion.Visible = bVisible
        btnComplete.Visible = bVisible
        btnAddNote.Visible = bVisible
        btnSetNextAction.Visible = bVisible
        
        lblContactAgent.Visible = False
        lblRcvdAtTransworld.Visible = False
        lblSentToWesternUnion.Visible = False
        lblComplete.Visible = False
        
    End Sub

    Protected Sub lnkbtnNotifications_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call HideAllPanels()
        pnlNotifications.Visible = True
        Call InitNotifications()
    End Sub
    
    Protected Sub InitNotifications()
        Dim sSQL As String = "SELECT wn.*, up.FirstName + ' ' + up.LastName + ' (' + up.UserID + ')' 'UserAccount'  FROM ClientData_WAT_Notifications2 wn INNER JOIN UserProfile up ON wn.UserKey = up.[key] ORDER BY up.UserID"
        Dim dtNotifications As DataTable = ExecuteQueryToDataTable(sSQL)
        gvNotifications.DataSource = dtNotifications
        gvNotifications.DataBind()
    End Sub
    
    Protected Sub btnNotificationsEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        'WebMsgBox.Show("Sorry, editing notification settings is not yet supported.")
        Call ShowMessage("Sorry, editing notification settings is not yet supported.")
    End Sub
    
    Protected Sub btnNotificationsClose_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        pnlNotifications.Visible = False
        Call DisplayMainPanel()
    End Sub
    
    Protected Sub btnSaveNote_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call AuditEntry(ENTRY_TYPE_NOTE, tbNote.Text.Replace(Environment.NewLine, " "))
        Call RefreshDetailPanel()
        Call HideDialogPanels()
        'pnlAddNote.Visible = False
        pnlTerminationManagement.Visible = True
    End Sub
    
#Region "Set Next Action"

    Protected Sub btnSetNextAction_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        pnlSetNextAction.Visible = True
        ddlNextAction.Focus()
        Dim drTerminationDetails As DataRow = GetTerminationDetailsFromRecord()
        Dim sNextAction As String = drTerminationDetails("NextAction")
        For i = 0 To ddlNextAction.Items.Count - 1
            If ddlNextAction.Items(i).Value = sNextAction Then
                ddlNextAction.SelectedIndex = i
                Exit For
            End If
        Next
    End Sub
    
    Protected Sub ddlNextAction_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim drTerminationDetails As DataRow = GetTerminationDetailsFromRecord()
        'Call AuditEntry(ENTRY_TYPE_EVENT, "Set Next Action: Action set to " & ddlNextAction.SelectedItem.Text & ".")
        Call ExecuteQueryToDataTable("UPDATE ClientData_WAT_Termination2 SET NextAction = " & QuotedNormalised(ddlNextAction.SelectedValue) & " WHERE [id] = " & pnTerminationKey)
        Call AuditEntry(ENTRY_TYPE_EVENT, "Set Next Action: Next Action changed from " & CodeStrippedNextAction(drTerminationDetails("NextAction")) & " to " & CodeStrippedNextAction(ddlNextAction.SelectedValue) & ".")
        Call BindTerminationsGrid(psSearchTerm)
        Call RefreshDetailPanel()
    End Sub

    Protected Sub btnSetNextActionClose_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        pnlSetNextAction.Visible = False
    End Sub

#End Region

#Region "Notifications"
    
    Protected Sub SetNotificationTimer(ByVal c As Control)
        c.Visible = True
        tmrNotificationTimer.Enabled = True
    End Sub
    
    Protected Sub ShowMessage(ByVal sMessage As String, Optional ByVal nInterval As Int32 = INTERVAL_SHORT_CONFIRMATION_MESSAGE)
        If lblMessage.Text.Contains("SYSTEM ERROR") Then
            Exit Sub
        End If
        lblMessage.Text = "&nbsp;" & sMessage & "&nbsp;"
        lblMessage.Visible = True
        lblMessage.BackColor = Drawing.Color.Red
        If nInterval = INTERVAL_SHORT_CONFIRMATION_MESSAGE Then
            lblMessage.BackColor = Drawing.Color.DarkGreen
        End If
        If nInterval <> INTERVAL_PERMANENT_MESSAGE Then
            tmrNotificationTimer.Interval = nInterval
            tmrNotificationTimer.Enabled = True
        End If
    End Sub

    Protected Sub tmrNotificationTimer_Tick(ByVal sender As Object, ByVal e As System.EventArgs)
        'lblLegendSaved.Visible = False
        lblMessage.Visible = False
        'lblMessage.Text = ""
        tmrNotificationTimer.Enabled = False
        If psTimerAction = "DUPLICATE_UNIQUEID" Or psTimerAction = "BAD_UNIQUEID" Then
            tbTerminationUniqueID.Focus()
        End If
        If psTimerAction = "TERMINATION_VALIDATION" Then
            If tbTerminationName.Text = String.Empty Then
                tbTerminationName.Focus()
            ElseIf tbTerminationAddr1.Text = String.Empty Then
                tbTerminationAddr1.Focus()
            ElseIf tbTerminationTownCity.Text = String.Empty Then
                tbTerminationTownCity.Focus()
            ElseIf tbTerminationRegionPostCode.Text = String.Empty Then
                tbTerminationRegionPostCode.Focus()
            ElseIf ddlTerminationCountry.SelectedValue = 0 Then
                ddlTerminationCountry.Focus()
            ElseIf tbTerminationPhone.Text = String.Empty Then
                tbTerminationPhone.Focus()
            End If
        End If
        If psTimerAction = "ADD_CONTACT_NOTE" Then
            tbContactNotes.Focus()
        End If
        psTimerAction = String.Empty
    End Sub
    
    Protected Sub DisplayFatalError(ByVal sMessage As String)
        gsSystemErrorMessage &= " " & sMessage
        lblMessage.Text = "SYSTEM ERROR - " & gsSystemErrorMessage
        lblMessage.BackColor = Drawing.Color.Red
        lblMessage.Visible = True
    End Sub
    
#End Region

#Region "Helper Functions"

    Protected Function ExecuteQueryToDataTable(ByVal sQuery As String) As DataTable
        Dim oDataTable As New DataTable
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter(sQuery, oConn)
        Dim oCmd As SqlCommand = New SqlCommand(sQuery, oConn)
        Try
            oAdapter.Fill(oDataTable)
            oConn.Open()
        Catch ex As Exception
            Call DisplayFatalError(sQuery & " : " & ex.Message)
            ExecuteQueryToDataTable = Nothing
        Finally
            oConn.Close()
            ExecuteQueryToDataTable = oDataTable
        End Try
    End Function
    
#End Region

    Protected Sub rbContactAgentAbandoned_CheckedChanged(sender As Object, e As System.EventArgs)
        tbContactNotes.Text = tbContactNotes.Text.Trim
        If tbContactNotes.Text = String.Empty Then
            Call ShowMessage("Please enter a reason in the Contact Notes")
            tbContactNotes.Focus()
            psTimerAction = "ADD_CONTACT_NOTE"
        End If
    End Sub
    
    Protected Sub btnContactAgent_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ContactAgent()
    End Sub

    Protected Sub ContactAgent()
        Call HideDialogPanels()
        pnlContactAgent.Visible = True
        Call ButtonPanelVisibility(False)
        pbGridDisabled = True
        Dim drTerminationDetails As DataRow = GetTerminationDetailsFromRecord()
        tbContactAgentAddressContactName.Text = drTerminationDetails("AgentContactName")
        tbContactAgentAddressAddr1.Text = drTerminationDetails("AgentAddress1")
        tbContactAgentAddressAddr2.Text = drTerminationDetails("AgentAddress2")
        tbContactAgentAddressAddr3.Text = drTerminationDetails("AgentAddress3")
        tbContactAgentAddressTownCity.Text = drTerminationDetails("AgentTownCity")
        tbContactAgentAddressRegionPostcode.Text = drTerminationDetails("AgentRegionOrPostCode")
        
        lblTelephoneNumber.Text = drTerminationDetails("AgentPhone")
        
        If lblTelephoneNumber.Text = String.Empty Then
            lblTelephoneNumber.Text = "(No number recorded)"
        End If
        rbContactAgentSuccess.Checked = False
        rbContactAgentFail.Checked = False
        rbContactAgentAbandoned.Checked = False
        tbContactNotes.Text = String.Empty
        rdtpCollectionDateTime.SelectedDate = Nothing
        If Not IsDBNull(drTerminationDetails("Collection01Date")) AndAlso CDate(drTerminationDetails("Collection01Date")) > "01-Jan-1900 00:00:00" Then
            rdtpCollectionDateTime.SelectedDate = drTerminationDetails("Collection01Date")
        End If
    End Sub
    
    Protected Sub btnContactAgentFinish_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim drTerminationDetails As DataRow = GetTerminationDetailsFromRecord()
        If rbContactAgentFail.Checked = False And rbContactAgentSuccess.Checked = False And rbContactAgentAbandoned.Checked = False Then
            Call ShowMessage("Please select 'Success', 'Fail' or 'Abandoned' to indicate the status of Agent Contact.")
            Exit Sub
        End If
        Dim nWorkingDays = 3     ' Wed -> following Mon, Thu -> following Tue, Fri - following Wed
        If (Date.Today.DayOfWeek = DayOfWeek.Wednesday) Or (Date.Today.DayOfWeek = DayOfWeek.Thursday) Or (Date.Today.DayOfWeek = DayOfWeek.Friday) Then
            nWorkingDays += 2
        End If
        If rbContactAgentSuccess.Checked Then
            If rdtpCollectionDateTime.SelectedDate Is Nothing Then
                Call ShowMessage("Please enter the Collection Date and Time.")
                Exit Sub
            ElseIf Date.Parse(rdtpCollectionDateTime.SelectedDate) < DateAdd(DateInterval.Day, nWorkingDays, Date.Now) Then
                Call ShowMessage("Collection date must be at least 3 days working from today.")
                Exit Sub
            Else
                
                Call AuditEntry(ENTRY_TYPE_EVENT, "Contact Agent: attempt marked as SUCCESS.")
                
                Call AuditEntry(ENTRY_TYPE_EVENT, "Contact Agent: Collection date/time recorded: " & rdtpCollectionDateTime.SelectedDate)
                Call ExecuteQueryToDataTable("UPDATE ClientData_WAT_Termination2 SET Collection01Date = '" & Format(rdtpCollectionDateTime.SelectedDate, "dd-MMM-yyyy hh:mm:ss") & "' WHERE [id] = " & pnTerminationKey)
                
                Call ExecuteQueryToDataTable("UPDATE ClientData_WAT_Termination2 SET NextAction = '" & NEXT_ACTION_COLLECT_FORMS & "' WHERE [id] = " & pnTerminationKey)
                Call AuditEntry(ENTRY_TYPE_EVENT, "Next Action changed from " & CodeStrippedNextAction(drTerminationDetails("NextAction")) & " to " & CodeStrippedNextAction(NEXT_ACTION_COLLECT_FORMS) & ".")
            End If
        ElseIf rbContactAgentFail.Checked Then
            Call AuditEntry(ENTRY_TYPE_EVENT, "Contact Agent: attempt marked as FAIL.")
        ElseIf rbContactAgentAbandoned.Checked Then
            tbContactNotes.Text = tbContactNotes.Text.Trim
            If tbContactNotes.Text = String.Empty Then
                Call ShowMessage("Please enter a reason for abandoning the termination.")
                Exit Sub
            Else
                Call ExecuteQueryToDataTable("UPDATE ClientData_WAT_Termination2 SET NextAction = '" & NEXT_ACTION_COMPLETE_TERMINATION & "' WHERE [id] = " & pnTerminationKey)
                Call AuditEntry(ENTRY_TYPE_EVENT, "Next Action changed from " & CodeStrippedNextAction(drTerminationDetails("NextAction")) & " to " & CodeStrippedNextAction(NEXT_ACTION_COLLECT_FORMS) & ".")
                Call AuditEntry(ENTRY_TYPE_EVENT, "Contact Agent: attempt marked as ABANDONED.")
                Call ExecuteQueryToDataTable("UPDATE ClientData_WAT_Termination2 SET TerminationClosed = ~TerminationClosed WHERE [id] = " & pnTerminationKey)
                Call AuditEntry(ENTRY_TYPE_EVENT, "Contact Agent: termination marked as CLOSED because contact attempt marked as ABANDONED.")
            End If
            Call RefreshDetailPanel()
            Call BindTerminationsGrid(psSearchTerm)
            pnlContactAgent.Visible = False
            pbGridDisabled = False
            Exit Sub
        Else
            Call ShowMessage("ERROR - bad option")
        End If
        
        If tbContactNotes.Text.Trim <> String.Empty Then
            Call AuditEntry(ENTRY_TYPE_DETAIL_CHANGE, "Contact Agent: recorded note " & tbContactNotes.Text.Trim)
        End If
        
        Dim sOriginalValue As String = String.Empty
        Dim sNewValue As String = String.Empty
        Const BLANK_MESSAGE As String = "BLANK"

        If trContactAgentAddressContactName.Visible Then

            If tbContactAgentAddressContactName.Text <> drTerminationDetails("AgentContactName") Then
                sOriginalValue = drTerminationDetails("AgentContactName")
                sNewValue = tbContactAgentAddressContactName.Text
                If sOriginalValue.Trim = String.Empty Then
                    sOriginalValue = BLANK_MESSAGE
                End If
                If sNewValue.Trim = String.Empty Then
                    sNewValue = BLANK_MESSAGE
                End If
                Call AuditEntry(ENTRY_TYPE_DETAIL_CHANGE, "Agent Contact Name changed from " & sOriginalValue & " to " & sNewValue)
            End If

            If tbContactAgentAddressAddr1.Text <> drTerminationDetails("AgentAddress1") Then
                sOriginalValue = drTerminationDetails("AgentAddress1")
                sNewValue = tbContactAgentAddressAddr1.Text
                If sOriginalValue.Trim = String.Empty Then
                    sOriginalValue = BLANK_MESSAGE
                End If
                If sNewValue.Trim = String.Empty Then
                    sNewValue = BLANK_MESSAGE
                End If
                Call AuditEntry(ENTRY_TYPE_DETAIL_CHANGE, "Agent Address 1 changed from " & sOriginalValue & " to " & sNewValue)
            End If

            If tbContactAgentAddressAddr2.Text <> drTerminationDetails("AgentAddress2") Then
                sOriginalValue = drTerminationDetails("AgentAddress2")
                sNewValue = tbContactAgentAddressAddr2.Text
                If sOriginalValue.Trim = String.Empty Then
                    sOriginalValue = BLANK_MESSAGE
                End If
                If sNewValue.Trim = String.Empty Then
                    sNewValue = BLANK_MESSAGE
                End If
                Call AuditEntry(ENTRY_TYPE_DETAIL_CHANGE, "Agent Address 2 changed from " & sOriginalValue & " to " & sNewValue)
            End If

            If tbContactAgentAddressAddr3.Text <> drTerminationDetails("AgentAddress3") Then
                sOriginalValue = drTerminationDetails("AgentAddress3")
                sNewValue = tbContactAgentAddressAddr3.Text
                If sOriginalValue.Trim = String.Empty Then
                    sOriginalValue = BLANK_MESSAGE
                End If
                If sNewValue.Trim = String.Empty Then
                    sNewValue = BLANK_MESSAGE
                End If
                Call AuditEntry(ENTRY_TYPE_DETAIL_CHANGE, "Agent Address 3 changed from " & sOriginalValue & " to " & sNewValue)
            End If

            If tbContactAgentAddressTownCity.Text <> drTerminationDetails("AgentTownCity") Then
                sOriginalValue = drTerminationDetails("AgentTownCity")
                sNewValue = tbContactAgentAddressTownCity.Text
                If sOriginalValue.Trim = String.Empty Then
                    sOriginalValue = BLANK_MESSAGE
                End If
                If sNewValue.Trim = String.Empty Then
                    sNewValue = BLANK_MESSAGE
                End If
                Call AuditEntry(ENTRY_TYPE_DETAIL_CHANGE, "Agent Town/City changed from " & sOriginalValue & " to " & sNewValue)
            End If

            If tbContactAgentAddressRegionPostcode.Text <> drTerminationDetails("AgentRegionOrPostCode") Then
                sOriginalValue = drTerminationDetails("AgentRegionOrPostCode")
                sNewValue = tbContactAgentAddressRegionPostcode.Text
                If sOriginalValue.Trim = String.Empty Then
                    sOriginalValue = BLANK_MESSAGE
                End If
                If sNewValue.Trim = String.Empty Then
                    sNewValue = BLANK_MESSAGE
                End If
                Call AuditEntry(ENTRY_TYPE_DETAIL_CHANGE, "Agent Region / PostCode changed from " & sOriginalValue & " to " & sNewValue)
            End If

            Call ExecuteQueryToDataTable("UPDATE ClientData_WAT_Termination2 SET AgentContactName = " & QuotedNormalised(tbContactAgentAddressContactName.Text) & ", AgentAddress1 = " & QuotedNormalised(tbContactAgentAddressAddr1.Text) & ", AgentAddress2 = " & QuotedNormalised(tbContactAgentAddressAddr2.Text) & ", AgentAddress3 = " & QuotedNormalised(tbContactAgentAddressAddr3.Text) & ", AgentTownCity = " & QuotedNormalised(tbContactAgentAddressTownCity.Text) & ", AgentRegionOrPostCode = " & QuotedNormalised(tbContactAgentAddressRegionPostcode.Text) & " WHERE [id] = " & pnTerminationKey)
            'Call SetButtonAdvisories()
        End If
            
        pnlContactAgent.Visible = False
        Call ButtonPanelVisibility(True)
        Call BindTerminationsGrid(psSearchTerm)
        Call RefreshDetailPanel()
        pbGridDisabled = False
    End Sub
    
    Protected Sub rbAuditShowCheckedChanged(sender As Object, e As System.EventArgs)
        Call PopulateAuditTrail()
    End Sub

    Protected Sub lnkbtnContactAgentEditAddress_Click(sender As Object, e As System.EventArgs)
        Dim lnkbtn As LinkButton = sender
        If lnkbtn.Text.Contains("edit") Then
            trContactAgentAddressContactName.Visible = True
            trContactAgentAddressAddr1.Visible = True
            trContactAgentAddressAddr2.Visible = True
            trContactAgentAddressAddr3.Visible = True
            trContactAgentAddressTownCity.Visible = True
            trContactAgentAddressRegionPostcode.Visible = True
        Else
            trContactAgentAddressContactName.Visible = False
            trContactAgentAddressAddr1.Visible = False
            trContactAgentAddressAddr2.Visible = False
            trContactAgentAddressAddr3.Visible = False
            trContactAgentAddressTownCity.Visible = False
            trContactAgentAddressRegionPostcode.Visible = False
            lnkbtn.Text = "edit address"
        End If
    End Sub
    
    Protected Sub radmenuReports_ItemClick(sender As Object, e As Telerik.Web.UI.RadMenuEventArgs)
        Dim rm As RadMenu = sender
        Select Case rm.SelectedValue
            Case REPORTS_MENU_REPORT_A
                Call ShowMessage(REPORTS_MENU_REPORT_A & " is not yet available")
            Case REPORTS_MENU_REPORT_B
                Call ShowMessage(REPORTS_MENU_REPORT_B & " is not yet available")
            Case REPORTS_MENU_REPORT_C
                Call ShowMessage(REPORTS_MENU_REPORT_C & " is not yet available")
            Case REPORTS_MENU_REPORT_D
                Call ShowMessage(REPORTS_MENU_REPORT_D & " is not yet available")
            Case Else
                Call ShowMessage("Triggered undefined menu item")
        End Select
    End Sub

#Region "Properties"

    Property psTerminationsSortExpression() As String
        Get
            Dim o As Object = ViewState("WAT_TerminationsSortExpression")
            If o Is Nothing Then
                Return " [id]"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("WAT_TerminationsSortExpression") = Value
        End Set
    End Property

    Property psCreatedOnSortDirection() As String
        Get
            Dim o As Object = ViewState("WAT_CreatedOnSortDirection")
            If o Is Nothing Then
                Return " ASC"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("WAT_CreatedOnSortDirection") = Value
        End Set
    End Property

    Property psNextActionSortDirection() As String
        Get
            Dim o As Object = ViewState("WAT_NextActionSortDirection")
            If o Is Nothing Then
                Return " ASC"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("WAT_NextActionSortDirection") = Value
        End Set
    End Property

    Property psUniqueIDSortDirection() As String
        Get
            Dim o As Object = ViewState("WAT_UniqueIDSortDirection")
            If o Is Nothing Then
                Return " ASC"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("WAT_UniqueIDSortDirection") = Value
        End Set
    End Property

    Property psAgentSortDirection() As String
        Get
            Dim o As Object = ViewState("WAT_AgentSortDirection")
            If o Is Nothing Then
                Return " ASC"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("WAT_AgentSortDirection") = Value
        End Set
    End Property

    Property pnTerminationKey() As Integer
        Get
            Dim o As Object = ViewState("WAT_TerminationKey")
            If o Is Nothing Then
                Return 0
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Integer)
            ViewState("WAT_TerminationKey") = Value
        End Set
    End Property
    
    Property psTimerAction() As String
        Get
            Dim o As Object = ViewState("WAT_TimerAction")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("WAT_TimerAction") = Value
        End Set
    End Property
  
    Property pnPermission() As Int32
        Get
            Dim o As Object = ViewState("WAT_Permission")
            If o Is Nothing Then
                Return 0
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Int32)
            ViewState("WAT_Permission") = Value
        End Set
    End Property

    Property pnTerminationPage() As Int32
        Get
            Dim o As Object = ViewState("WAT_TerminationPage")
            If o Is Nothing Then
                Return 0
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Int32)
            ViewState("WAT_TerminationPage") = Value
        End Set
    End Property

    Property psSearchTerm() As String
        Get
            Dim o As Object = ViewState("WAT_SearchTerm")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("WAT_SearchTerm") = Value
        End Set
    End Property
    
    Property pbIsEditing() As Boolean
        Get
            Dim o As Object = ViewState("WAT_IsEditing")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("WAT_IsEditing") = Value
        End Set
    End Property

    Property pbGridDisabled() As Boolean
        Get
            Dim o As Object = ViewState("WAT_GridDisabled")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("WAT_GridDisabled") = Value
            If Value Then
                gvTerminations.RowStyle.CssClass = Nothing
                gvTerminations.AlternatingRowStyle.CssClass = Nothing
                tbSearchTerminations.Enabled = False
                btnSearchTerminations.Enabled = False
                lnkbtnClearSearchTerminationsBox.Enabled = False
                cbIgnoreCompletedTerminations.Enabled = False
                btnNewTermination.Visible = False
            Else
                gvTerminations.RowStyle.CssClass = "DataRow"
                gvTerminations.AlternatingRowStyle.CssClass = "DataRow DataRowAlt"
                tbSearchTerminations.Enabled = True
                btnSearchTerminations.Enabled = True
                lnkbtnClearSearchTerminationsBox.Enabled = True
                cbIgnoreCompletedTerminations.Enabled = True
                btnNewTermination.Visible = True
            End If
        End Set
    End Property

#End Region

    Protected Sub btnRcvdAtTransworld_Click(sender As Object, e As System.EventArgs)
        Call HideDialogPanels()
        pnlFormsReceived.Visible = True
        tbNotesFormsReceived.Text = String.Empty
        tbNotesFormsReceived.Focus()
    End Sub

    Protected Sub btnSentToWesternUnion_Click(sender As Object, e As System.EventArgs)
        Call HideDialogPanels()
        pnlFormsSent.Visible = True
        tbNotesFormsSent.Text = String.Empty
        tbNotesFormsSent.Focus()
    End Sub

    Protected Sub btnComplete_Click(sender As Object, e As System.EventArgs)
        Call HideDialogPanels()
        pnlTerminationComplete.Visible = True
        tbNotesTerminationComplete.Text = String.Empty
        tbNotesTerminationComplete.Focus()
    End Sub

    Protected Sub btnConfirmFormsReceived_Click(sender As Object, e As System.EventArgs)
        Dim drTerminationDetails As DataRow = GetTerminationDetailsFromRecord()
        Call ExecuteQueryToDataTable("UPDATE ClientData_WAT_Termination2 SET NextAction = '" & NEXT_ACTION_DELIVER_TO_WESTERN_UNION & "' WHERE [id] = " & pnTerminationKey)
        Call AuditEntry(ENTRY_TYPE_EVENT, "Next Action changed from " & CodeStrippedNextAction(drTerminationDetails("NextAction")) & " to " & CodeStrippedNextAction(NEXT_ACTION_DELIVER_TO_WESTERN_UNION) & ".")
        Call BindTerminationsGrid(psSearchTerm)
        Call AuditEntry(ENTRY_TYPE_NOTE, tbNotesFormsReceived.Text.Replace(Environment.NewLine, " "))
        Call RefreshDetailPanel()
        pnlTerminationManagement.Visible = True
        Call HideDialogPanels()
    End Sub

    Protected Sub btnConfirmSent_Click(sender As Object, e As System.EventArgs)
        Dim drTerminationDetails As DataRow = GetTerminationDetailsFromRecord()
        Call ExecuteQueryToDataTable("UPDATE ClientData_WAT_Termination2 SET NextAction = '" & NEXT_ACTION_COMPLETE_TERMINATION & "' WHERE [id] = " & pnTerminationKey)
        Call AuditEntry(ENTRY_TYPE_EVENT, "Next Action changed from " & CodeStrippedNextAction(drTerminationDetails("NextAction")) & " to " & CodeStrippedNextAction(NEXT_ACTION_COMPLETE_TERMINATION) & ".")
        Call BindTerminationsGrid(psSearchTerm)
        Call AuditEntry(ENTRY_TYPE_NOTE, tbNotesFormsSent.Text.Replace(Environment.NewLine, " "))
        Call RefreshDetailPanel()
        pnlTerminationManagement.Visible = True
        Call HideDialogPanels()
    End Sub

    Protected Sub btnConfirmTerminationComplete_Click(sender As Object, e As System.EventArgs)
        Dim drTerminationDetails As DataRow = GetTerminationDetailsFromRecord()
        Dim sStatus As String = String.Empty
        Dim sSQL As String
        sSQL = "UPDATE ClientData_WAT_Termination2 SET TerminationClosed = ~TerminationClosed WHERE [id] = " & pnTerminationKey & " SELECT TerminationClosed FROM ClientData_WAT_Termination2 WHERE [id] = " & pnTerminationKey
        Dim bTerminationClosed As Boolean = ExecuteQueryToDataTable(sSQL).Rows(0).Item(0)
        If bTerminationClosed Then
            sStatus = "CLOSED"
            btnComplete.Text = "Mark as Re-opened"
            Call ExecuteQueryToDataTable("UPDATE ClientData_WAT_Termination2 SET NextAction = '" & NEXT_ACTION_COMPLETE_TERMINATION & "' WHERE [id] = " & pnTerminationKey)
            Call AuditEntry(ENTRY_TYPE_EVENT, "Next Action changed from " & CodeStrippedNextAction(drTerminationDetails("NextAction")) & " to " & CodeStrippedNextAction(NEXT_ACTION_COMPLETE_TERMINATION) & ".")
        Else
            sStatus = "RE-OPENED"
            btnComplete.Text = "Mark as Complete"
            Call ExecuteQueryToDataTable("UPDATE ClientData_WAT_Termination2 SET NextAction = '" & NEXT_ACTION_COMPLETED & "' WHERE [id] = " & pnTerminationKey)
            Call AuditEntry(ENTRY_TYPE_EVENT, "Next Action changed from " & CodeStrippedNextAction(drTerminationDetails("NextAction")) & " to " & CodeStrippedNextAction(NEXT_ACTION_COMPLETED) & ".")
        End If
        
        Call AuditEntry(ENTRY_TYPE_TERMINATION_COMPLETION_STATUS, sStatus)
        ShowMessage("Termination " & sStatus & ".")
        Call AuditEntry(ENTRY_TYPE_NOTE, tbNotesTerminationComplete.Text.Replace(Environment.NewLine, " "))
        Call BindTerminationsGrid(tbSearchTerminations.Text)
        Call RefreshDetailPanel()
        Call HideDialogPanels()
    End Sub
    
    Protected Sub SetButtonAdvisories()

        Dim sCurrentState As String = GetCurrentState()
        Dim sWarning As String = "return confirm('The action you have selected differs from the normal sequence of termination events.\n\nAre you sure this is the action you want to take? CURRENTACTION = " & sCurrentState & "');"
        'Dim sWarning As String = "return confirm('The action you have selected differs from the normal sequence of termination events.\n\nAre you sure this is the action you want to take?');"

        btnConfirmTerminationComplete.OnClientClick = String.Empty
        btnComplete.OnClientClick = String.Empty
        btnConfirmFormsReceived.OnClientClick = String.Empty
        btnConfirmSent.OnClientClick = String.Empty
        btnContactAgent.OnClientClick = String.Empty

        lblContactAgent.Visible = False
        lblRcvdAtTransworld.Visible = False
        lblSentToWesternUnion.Visible = False
        lblComplete.Visible = False

        Select Case GetCurrentState()
            Case NEXT_ACTION_CONTACT_AGENT
                'btnComplete.OnClientClick = sWarning
                'btnRcvdAtTransworld.OnClientClick = sWarning
                'btnConfirmSent.OnClientClick = sWarning
                ''''''''''''''btnContactAgent.OnClientClick = sWarning
                lblContactAgent.Visible = True
            Case NEXT_ACTION_COLLECT_FORMS
                'btnComplete.OnClientClick = sWarning
                ''''''''''''''btnRcvdAtTransworld.OnClientClick = sWarning
                lblRcvdAtTransworld.Visible = True
                'btnConfirmSent.OnClientClick = sWarning
                'btnContactAgent.OnClientClick = sWarning
            Case NEXT_ACTION_DELIVER_TO_WESTERN_UNION
                'btnComplete.OnClientClick = sWarning
                'btnRcvdAtTransworld.OnClientClick = sWarning
                ''''''''''''''btnConfirmSent.OnClientClick = sWarning
                lblSentToWesternUnion.Visible = True
                'btnContactAgent.OnClientClick = sWarning
            Case NEXT_ACTION_COMPLETE_TERMINATION
                'btnComplete.'''''''''''''OnClientClick = sWarning
                lblComplete.Visible = True
                'btnRcvdAtTransworld.OnClientClick = sWarning
                'btnConfirmSent.OnClientClick = sWarning
                'btnContactAgent.OnClientClick = sWarning
        End Select
    End Sub
    
    Protected Sub AddNewProduct()
        
        Exit Sub
        
        
        Dim drTerminationDetails As DataRow = GetTerminationDetailsFromRecord()
        Dim sAgentDetails As String = drTerminationDetails("AgentName") & " " & drTerminationDetails("AgentAddress1") & " " & drTerminationDetails("AgentTownCity") & " " & drTerminationDetails("AgentRegionOrPostCode") & " " & drTerminationDetails("AgentContactName")
        Dim nIndex As Integer
        Dim oConn As New SqlConnection(gsConn)
        Dim oTrans As SqlTransaction
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Product_AddWithAccessControl8", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
 
        Dim paramUserKey As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int)
        paramUserKey.Value = 0
        oCmd.Parameters.Add(paramUserKey)

        Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int)
        paramCustomerKey.Value = CUSTOMER_WUCOLL
        oCmd.Parameters.Add(paramCustomerKey)
 
        Dim paramProductCode As SqlParameter = New SqlParameter("@ProductCode", SqlDbType.NVarChar, 25)
        paramProductCode.Value = drTerminationDetails("AgentUniqueID")
        oCmd.Parameters.Add(paramProductCode)
      
        Dim paramProductDate As SqlParameter = New SqlParameter("@ProductDate", SqlDbType.NVarChar, 10)
        paramProductDate.Value = String.Empty
        oCmd.Parameters.Add(paramProductDate)
 
        Dim paramMinimumStockLevel As SqlParameter = New SqlParameter("@MinimumStockLevel", SqlDbType.Int, 4)
        paramMinimumStockLevel.Value = 0
        oCmd.Parameters.Add(paramMinimumStockLevel)
      
        Dim paramDescription As SqlParameter = New SqlParameter("@ProductDescription", SqlDbType.NVarChar, 300)
        paramDescription.Value = sAgentDetails
        oCmd.Parameters.Add(paramDescription)
      
        Dim paramItemsPerBox As SqlParameter = New SqlParameter("@ItemsPerBox", SqlDbType.Int, 4)
        paramItemsPerBox.Value = 0
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
        paramUnitValue.Value = 0
        oCmd.Parameters.Add(paramUnitValue)
      
        Dim paramUnitValue2 As SqlParameter = New SqlParameter("@UnitValue2", SqlDbType.Money, 8)
        paramUnitValue2.Value = 0
        oCmd.Parameters.Add(paramUnitValue2)

        Dim paramLanguage As SqlParameter = New SqlParameter("@LanguageId", SqlDbType.NVarChar, 20)
        paramLanguage.Value = String.Empty
        oCmd.Parameters.Add(paramLanguage)

        Dim paramDepartment As SqlParameter = New SqlParameter("@ProductDepartmentId", SqlDbType.NVarChar, 20)
        paramDepartment.Value = String.Empty
        oCmd.Parameters.Add(paramDepartment)
      
        Dim paramWeight As SqlParameter = New SqlParameter("@UnitWeightGrams", SqlDbType.Int, 4)
        paramWeight.Value = 0
        oCmd.Parameters.Add(paramWeight)
      
        Dim paramStockOwnedByKey As SqlParameter = New SqlParameter("@StockOwnedByKey", SqlDbType.Int, 4)
        paramStockOwnedByKey.Value = 0
        oCmd.Parameters.Add(paramStockOwnedByKey)
      
        Dim paramMisc1 As SqlParameter = New SqlParameter("@Misc1", SqlDbType.NVarChar, 50)
        paramMisc1.Value = String.Empty
        oCmd.Parameters.Add(paramMisc1)
      
        Dim paramMisc2 As SqlParameter = New SqlParameter("@Misc2", SqlDbType.NVarChar, 50)
        paramMisc2.Value = String.Empty
        oCmd.Parameters.Add(paramMisc2)
      
        Dim paramArchive As SqlParameter = New SqlParameter("@ArchiveFlag", SqlDbType.NVarChar, 1)
        paramArchive.Value = "N"
        oCmd.Parameters.Add(paramArchive)
    
        Dim paramStatus As SqlParameter = New SqlParameter("@Status", SqlDbType.TinyInt)
        paramStatus.Value = 0
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
        paramNotes.Value = String.Empty
        oCmd.Parameters.Add(paramNotes)

        Dim paramViewOnWebForm As SqlParameter = New SqlParameter("@ViewOnWebForm", SqlDbType.Bit)
        paramViewOnWebForm.Value = 0
        oCmd.Parameters.Add(paramViewOnWebForm)
 
        Dim paramDefaultAccessFlag As SqlParameter = New SqlParameter("@DefaultAccessFlag", SqlDbType.Bit)
        paramDefaultAccessFlag.Value = Not 0
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
            Dim lProductKey As Long = CLng(oCmd.Parameters("@ProductKey").Value)

        Catch ex As SqlException
            If ex.Number = 2627 Then
                WebMsgBox.Show("ERROR: A product record already exists with this product code.")
            Else
                WebMsgBox.Show(ex.ToString)
            End If
        Finally
            oConn.Close()
        End Try
    End Sub

    Protected Function GetCurrentState() As String
        Dim sSQL As String
        sSQL = "SELECT NextAction FROM ClientData_WAT_Termination2 WHERE [id] = " & pnTerminationKey
        GetCurrentState = ExecuteQueryToDataTable(sSQL).Rows(0).Item(0)
    End Function
    
    Protected Sub ExportData()
        Dim sSQL As String = "SELECT AgentUniqueID 'Agent Unique ID', AgentName 'Agent Name', AgentAddress1 'Addr 1', AgentAddress2 'Addr 2', AgentAddress3 'Addr 3', AgentTownCity 'Town', AgentRegionOrPostCode 'Region / Postcode', AgentContactName 'Contact Name', AgentPhone 'Phone', AgentEmail 'Email', Comments 'Comments', ISNULL(CAST(REPLACE(CONVERT(VARCHAR(11), Collection01Date, 106), ' ', '-') AS varchar(20)),'(none)') 'Collection Date' , CAST(REPLACE(CONVERT(VARCHAR(11),  CreatedOn, 106), ' ', '-') AS varchar(20)) + ' ' + SUBSTRING((CONVERT(VARCHAR(8), CreatedOn, 108)),1,5) 'Created On', dbo.AgentTerminationGetAuditEntries2(wt.[id]) 'Audit Trail' FROM ClientData_WAT_Termination2 wt ORDER BY [id]"
        'Dim dtOrders As DataTable = ExecuteQueryToDataTable("SELECT * FROM ClientData_WAT_Termination2 ORDER BY [id]")
        Dim dtData As DataTable = ExecuteQueryToDataTable(sSQL)
        If dtData.Rows.Count > 0 Then
            Response.Clear()
            Response.ContentType = "text/csv"
            Response.AddHeader("Content-Disposition", "attachment; filename=agent_terminations_" & DateTime.Now.ToString("dd-MMM-yyyyhhmmss") & ".csv")
    
            Dim oDataColumn As DataColumn
            Dim sItem As String
    
            For Each oDataColumn In dtData.Columns  ' write column header
                Response.Write(oDataColumn.ColumnName)
                Response.Write(",")
            Next
            Response.Write(vbCrLf)
    
            For Each dr As DataRow In dtData.Rows
                For Each oDataColumn In dtData.Columns
                    sItem = (dr(oDataColumn.ColumnName).ToString)
                    sItem = sItem.Replace(ControlChars.Quote, ControlChars.Quote & ControlChars.Quote)
                    sItem = ControlChars.Quote & sItem & ControlChars.Quote
                    Response.Write(sItem)
                    Response.Write(",")
                Next
                Response.Write(vbCrLf)
            Next
            Response.End()
        Else
            Call ShowMessage("No data to export.")
        End If
    End Sub

    'Protected Sub rdtpCollectionDateTime_SelectedDateChanged(sender As Object, e As Telerik.Web.UI.Calendar.SelectedDateChangedEventArgs)
    '    If rbContactAgentSuccess.Checked = False And rbContactAgentFail.Checked = False And rbContactAgentAbandoned.Checked = False Then
    '        rbContactAgentSuccess.Checked = True
    '    End If
    'End Sub
    
    Protected Sub lnkbtnExportData_Click(sender As Object, e As System.EventArgs)
        Call ExportData()
    End Sub
    
</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <style type="text/css">
        .style_padded_xx_verdana
        {
            font-family: Verdana;
            font-size: xx-small;
            margin-left: 10px;
            margin-right: 10px;
        }
        .style_textbox_base
        {
            font-family: Verdana;
            font-size: xx-small;
        }
        .style_textbox_short
        {
            width: 80px;
        }
        .style_textbox_medium
        {
            width: 200px;
        }
        .style_textbox_long
        {
            width: 300px;
        }
        .style_textbox_full
        {
            width: 100%;
        }
        .style_button_action
        {
            width: 150px;
        }
        .style_button_cancel
        {
            width: 80px;
        }
        .style_panel_name
        {
            padding: 15px;
        }
        .style_rfv
        {
            color: Red;
            font-size: medium;
            font-weight: bold;
        }
        .style_label_required_field
        {
            color: Red;
        }
        .DataRow:hover, .DataRowAlt:hover
        {
            background-color: yellow;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
    <main:Header ID="ctlHeader" runat="server" />
    <div style="margin-left: 0px; margin-right: 10px;">
        <asp:PlaceHolder ID="PlaceHolderForScriptManager" runat="server" />
        <asp:Timer ID="tmrNotificationTimer" runat="server" OnTick="tmrNotificationTimer_Tick"
            Interval="2000" Enabled="False" />
        <div style="text-align: center; height: 10px;">
            <asp:Label ID="lblMessage" runat="server" BackColor="Red" Font-Bold="True" Font-Names="Verdana"
                Font-Size="Small" ForeColor="White" Text="MESSAGE" Visible="False" />
        </div>
        <asp:Panel ID="pnlMain" Width="100%" runat="server" CssClass="style_padded_xx_verdana">
            <asp:Label ID="lblLegendWUAgentTerminations" runat="server" Font-Names="Verdana"
                Font-Size="Small" Font-Bold="True" Text="WU Agent Terminations (18MAR14)" CssClass="style_panel_name" />
            <table style="width: 99%">
                <tr>
                    <td style="width: 10%">
                    </td>
                    <td style="width: 70%">
                    </td>
                    <td style="width: 20%" align="right">
                        <asp:LinkButton ID="lnkbtnExportData" runat="server" 
                            onclick="lnkbtnExportData_Click">export</asp:LinkButton><telerik:RadMenu ID="radmenuReports" runat="server" OnItemClick="radmenuReports_ItemClick"
                            EnableRoundedCorners="True" EnableShadows="True" Width="150px" />
                    </td>
                </tr>
                <tr>
                    <td>
                        <asp:Label ID="lblLegendTerminations" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="Terminations:" />
                    </td>
                    <td>
                        <asp:Label ID="lblLegendSearch" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            ForeColor="Gray">search:</asp:Label>
                        &nbsp;<asp:TextBox ID="tbSearchTerminations" runat="server" Font-Size="XX-Small"
                            Width="100px" />
                        <asp:Button ID="btnSearchTerminations" runat="server" Text="go" ToolTip="search for a termination record"
                            OnClick="btnSearchTerminations_Click" />
                        &nbsp;
                        <asp:LinkButton ID="lnkbtnClearSearchTerminationsBox" runat="server" OnClick="lnkbtnClearSearchTerminationsBox_Click">clear</asp:LinkButton>
                        &nbsp;&nbsp;&nbsp;
                        <asp:CheckBox ID="cbIgnoreCompletedTerminations" runat="server" Checked="True" Font-Names="Verdana"
                            Font-Size="XX-Small" Text="ignore completed terminations" AutoPostBack="True"
                            CausesValidation="True" OnCheckedChanged="cbIgnoreCompletedTerminations_CheckedChanged" />
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                    </td>
                    <td align="right" style="width: 150px">
                        <asp:Button ID="btnNewTermination" runat="server" OnClick="btnNewTermination_Click"
                            Text="New Termination" />
                    </td>
                </tr>
                <tr>
                    <td colspan="3">
                        <sc:MyGridView ID="gvTerminations" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Width="100%" CellPadding="2" OnRowDataBound="gvTerminations_RowDataBound" OnSelectedIndexChanged="gvTerminations_SelectedIndexChanged"
                            EnableRowClick="True" OnRowClicked="myGridView1_RowClicked" AllowPaging="True"
                            AllowSorting="True" OnSorting="gvTerminations_Sorting" OnPageIndexChanging="gvTerminations_PageIndexChanging"
                            PageSize="5" AutoGenerateColumns="False">
                            <AlternatingRowStyle CssClass="DataRow DataRowAlt" />
                            <Columns>
                                <asp:TemplateField HeaderText="Created On" SortExpression="CreatedOn">
                                    <ItemTemplate>
                                        <asp:Label ID="lblCreatedOn" runat="server" Text='<%# Bind("CreatedOn") %>' /><asp:HiddenField
                                            ID="hidID" Value='<%# Bind("ID") %>' runat="server" />
                                        <asp:HiddenField ID="hidTerminationClosed" Value='<%# Bind("TerminationClosed") %>'
                                            runat="server" />
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:BoundField DataField="NextAction" HeaderText="Next Action" ReadOnly="True" SortExpression="NextAction" />
                                <asp:BoundField DataField="UniqueID" HeaderText="Unique ID" ReadOnly="True" SortExpression="UniqueID" />
                                <asp:BoundField DataField="AgentSummary" HeaderText="Agent" ReadOnly="True" SortExpression="AgentSummary" />
                            </Columns>
                            <EmptyDataTemplate>
                                <asp:Label ID="lblEmptyDataMessage" runat="server" Text='<%# sGetEmptyDataMessage %>' /></EmptyDataTemplate>
                            <RowStyle CssClass="DataRow" />
                        </sc:MyGridView>
                    </td>
                </tr>
                <tr>
                    <td>
                        <asp:Label ID="lblLegendItemCount" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="Items:" />
                        &nbsp;<asp:DropDownList ID="ddlDisplayCount" runat="server" AutoPostBack="True" Font-Size="X-Small"
                            OnSelectedIndexChanged="ddlDisplayCount_SelectedIndexChanged">
                            <asp:ListItem>5</asp:ListItem>
                            <asp:ListItem>20</asp:ListItem>
                            <asp:ListItem>50</asp:ListItem>
                        </asp:DropDownList>
                    </td>
                    <td>
                        &nbsp;
                        <asp:Label ID="lblSortInfo" runat="server" ForeColor="#666666" Text="sorted on Created On in ascending order"></asp:Label>
                    </td>
                    <td align="right">
                        <asp:LinkButton ID="lnkbtnNotifications" runat="server" OnClick="lnkbtnNotifications_Click">notifications</asp:LinkButton>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td>
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
                </tr>
                <tr>
                    <td>
                    </td>
                    <td>
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
                </tr>
            </table>
        </asp:Panel>
        <asp:Panel ID="pnlTerminationAddress" Width="99%" runat="server" CssClass="style_padded_xx_verdana"
            BackColor="#99CCFF">
            <table style="width: 99%">
                <tr>
                    <td style="width: 10%">
                    </td>
                    <td style="width: 25%">
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td colspan="2">
                        <asp:Label ID="lblLegendAgentTermination" runat="server" CssClass="style_panel_name"
                            Font-Bold="True" Font-Names="Verdana" Font-Size="Small" Text="New Agent Termination" />
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
                    <td>
                        &nbsp;
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right">
                        <asp:Label ID="lblLegendTerminationUniqueID" runat="server" Text="Unique ID:" CssClass="style_label_required_field" />
                    </td>
                    <td>
                        <asp:TextBox ID="tbTerminationUniqueID" runat="server" CssClass="style_textbox_base style_textbox_short"
                            MaxLength="50" Style="text-transform: uppercase" />
                        &nbsp;<asp:RequiredFieldValidator ID="rfvTerminationUniqueID" runat="server" ErrorMessage="*****"
                            CssClass="style_rfv" ControlToValidate="tbTerminationUniqueID" ValidationGroup="AgentTermination" />
                        &nbsp;
                        <asp:Button ID="btnFindUniqueID" runat="server" CausesValidation="False" CssClass="style_button_cancel"
                            OnClick="btnFindUniqueID_Click" Text="Find" />
                    </td>
                </tr>
                <tr>
                    <td>
                        &nbsp;
                    </td>
                    <td align="right">
                        <asp:Label ID="lblLegendTerminationAgentName" runat="server" Text="Agent Name:" CssClass="style_label_required_field" />
                    </td>
                    <td>
                        <asp:TextBox ID="tbTerminationName" runat="server" CssClass="style_textbox_base style_textbox_medium"
                            MaxLength="50" />
                        &nbsp;<asp:RequiredFieldValidator ID="rfvTerminationName" runat="server" ControlToValidate="tbTerminationName"
                            CssClass="style_rfv" ErrorMessage="*****" ValidationGroup="AgentTermination" />
                    </td>
                </tr>
                <tr>
                    <td>
                        &nbsp;
                    </td>
                    <td align="right">
                        <asp:Label ID="lblLegendTerminationAgentName0" runat="server" Text="Contact Name:" />
                    </td>
                    <td>
                        <asp:TextBox ID="tbTerminationContactName" runat="server" CssClass="style_textbox_base style_textbox_medium"
                            MaxLength="50" />
                    </td>
                </tr>
                <tr>
                    <td>
                        &nbsp;
                    </td>
                    <td align="right">
                        <asp:Label ID="lblLegendTerminationAddr1" runat="server" Text="Addr 1:" CssClass="style_label_required_field" />
                    </td>
                    <td>
                        <asp:TextBox ID="tbTerminationAddr1" runat="server" CssClass="style_textbox_base style_textbox_medium"
                            MaxLength="50" />
                        &nbsp;<asp:RequiredFieldValidator ID="rfvTerminationAddr1" runat="server" ControlToValidate="tbTerminationAddr1"
                            CssClass="style_rfv" ErrorMessage="*****" ValidationGroup="AgentTermination" />
                    </td>
                </tr>
                <tr>
                    <td>
                        &nbsp;
                    </td>
                    <td align="right">
                        <asp:Label ID="lblLegendTerminationAddr2" runat="server" Text="Addr 2:" />
                    </td>
                    <td>
                        <asp:TextBox ID="tbTerminationAddr2" runat="server" CssClass="style_textbox_base style_textbox_medium"
                            MaxLength="50" />
                    </td>
                </tr>
                <tr>
                    <td>
                        &nbsp;
                    </td>
                    <td align="right">
                        <asp:Label ID="lblLegendTerminationAddr3" runat="server" Text="Addr 3:" />
                    </td>
                    <td>
                        <asp:TextBox ID="tbTerminationAddr3" runat="server" CssClass="style_textbox_base style_textbox_medium"
                            MaxLength="50" />
                    </td>
                </tr>
                <tr>
                    <td>
                        &nbsp;
                    </td>
                    <td align="right">
                        <asp:Label ID="lblLegendTerminationTownCity" runat="server" Text="Town / City:" CssClass="style_label_required_field" />
                    </td>
                    <td>
                        <asp:TextBox ID="tbTerminationTownCity" runat="server" CssClass="style_textbox_base style_textbox_medium"
                            MaxLength="50" />
                        &nbsp;<asp:RequiredFieldValidator ID="rfvTerminationTownCity" runat="server" ControlToValidate="tbTerminationTownCity"
                            CssClass="style_rfv" ErrorMessage="*****" ValidationGroup="AgentTermination" />
                    </td>
                </tr>
                <tr>
                    <td>
                        &nbsp;
                    </td>
                    <td align="right">
                        <asp:Label ID="lblLegendTerminationRegionPostCode" runat="server" Text="Region / Postcode:"
                            CssClass="style_label_required_field" />
                    </td>
                    <td>
                        <asp:TextBox ID="tbTerminationRegionPostCode" runat="server" CssClass="style_textbox_base style_textbox_medium"
                            MaxLength="50" />
                        &nbsp;<asp:RequiredFieldValidator ID="rfvTerminationRegionPostCode" runat="server"
                            ControlToValidate="tbTerminationRegionPostCode" CssClass="style_rfv" ErrorMessage="*****"
                            ValidationGroup="AgentTermination" />
                    </td>
                </tr>
                <tr>
                    <td>
                        &nbsp;
                    </td>
                    <td align="right">
                        <asp:Label ID="lblLegendTerminationCountry" runat="server" CssClass="style_label_required_field"
                            Text="Country:" />
                    </td>
                    <td>
                        <asp:DropDownList ID="ddlTerminationCountry" runat="server" AutoPostBack="True" Font-Size="X-Small" />
                        &nbsp;<asp:LinkButton ID="lnkbtnTerminationUK" runat="server" CausesValidation="False"
                            OnClick="lnkbtnNewTerminationUK_Click">UK</asp:LinkButton>
                        &nbsp;<asp:LinkButton ID="lnkbtnTerminationNorthernIreland" runat="server" CausesValidation="False"
                            OnClick="lnkbtnNewTerminationNorthernIreland_Click">N.IRE</asp:LinkButton>
                        &nbsp;<asp:LinkButton ID="lnkbtnTerminationIreland" runat="server" CausesValidation="False"
                            OnClick="lnkbtnNewTerminationIreland_Click">IRE</asp:LinkButton>
                        &nbsp;<asp:RequiredFieldValidator ID="rfvTerminationCountry" runat="server" ControlToValidate="ddlTerminationCountry"
                            CssClass="style_rfv" ErrorMessage="*****" InitialValue="0" ValidationGroup="AgentTermination" />
                    </td>
                </tr>
                <tr>
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
                <tr>
                    <td>
                        &nbsp;
                    </td>
                    <td align="right">
                        <asp:Label ID="lblLegendTerminationPhone" runat="server" Text="Phone:" CssClass="style_label_required_field" />
                    </td>
                    <td>
                        <asp:TextBox ID="tbTerminationPhone" runat="server" CssClass="style_textbox_base style_textbox_medium"
                            MaxLength="50" />
                        &nbsp;<asp:RequiredFieldValidator ID="rfvTerminationPhone" runat="server" ControlToValidate="tbTerminationPhone"
                            CssClass="style_rfv" ErrorMessage="*****" ValidationGroup="AgentTermination" />
                    </td>
                </tr>
                <tr>
                    <td>
                        &nbsp;
                    </td>
                    <td align="right">
                        <asp:Label ID="lblLegendTerminationEmail" runat="server" Text="Email:" />
                    </td>
                    <td>
                        <asp:TextBox ID="tbTerminationEmail" runat="server" CssClass="style_textbox_base style_textbox_medium"
                            MaxLength="150" />
                    </td>
                </tr>
                <tr>
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
                <tr>
                    <td>
                        &nbsp;
                    </td>
                    <td align="right">
                        <asp:Label ID="lblLegendTerminationComments" runat="server" Text="Comments:" />
                    </td>
                    <td>
                        <asp:TextBox ID="tbTerminationComments" runat="server" CssClass="style_textbox_long"
                            Font-Names="Verdana" Font-Size="XX-Small" MaxLength="1000" Rows="4" TextMode="MultiLine" />
                    </td>
                </tr>
                <tr>
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
                <tr>
                    <td>
                        &nbsp;
                    </td>
                    <td>
                        &nbsp;
                    </td>
                    <td>
                        <asp:Button ID="btnTerminationSave" runat="server" CssClass="style_button_action"
                            OnClick="btnTerminationSave_Click" Text="Save" />
                        &nbsp;<asp:Button ID="btnTerminationCancel" runat="server" CausesValidation="False"
                            CssClass="style_button_cancel" OnClick="btnTerminationCancel_Click" Text="Cancel" />
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td>
                    </td>
                    <td>
                    </td>
                </tr>
            </table>
        </asp:Panel>
        <asp:Panel ID="pnlTerminationManagement" Width="99%" runat="server" CssClass="style_padded_xx_verdana"
            BackColor="#EEEEEE">
            <table style="width: 99%">
                <tr>
                    <td style="width: 10%">
                    </td>
                    <td style="width: 25%">
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td colspan="3">
                        <asp:Label ID="lblLegendManageTermination" runat="server" CssClass="style_panel_name"
                            Font-Bold="True" Font-Names="Verdana" Font-Size="Small" Text="Manage Termination" />
                        <asp:Label ID="lblTerminationIdentifier" runat="server" Text="Termination Identifier" />
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
                </tr>
                <tr>
                    <td>
                    </td>
                    <td>
                    </td>
                    <td align="right">
                        <asp:RadioButton ID="rbAuditShowMajorEvents" runat="server" Text="major events" AutoPostBack="True"
                            Checked="True" GroupName="AuditShow" OnCheckedChanged="rbAuditShowCheckedChanged" />
                        <asp:RadioButton ID="rbAuditShowAllEvents" runat="server" Text="all events" AutoPostBack="True"
                            GroupName="AuditShow" OnCheckedChanged="rbAuditShowCheckedChanged" />
                        &nbsp;&nbsp;&nbsp;&nbsp;
                    </td>
                </tr>
                <tr>
                    <td style="white-space:nowrap">
                        <asp:Button ID="btnEditAgentDetails" runat="server" CssClass="style_button_action"
                            Text="Edit Agent Details..." OnClick="btnEditAgentDetails_Click" />
                        <br />
                        <br />
                        <asp:Button ID="btnContactAgent" runat="server" CssClass="style_button_action" OnClick="btnContactAgent_Click"
                            Text="Contact Agent..." />
                        <asp:Label ID="lblContactAgent" runat="server" Font-Bold="True" 
                            ForeColor="Red" Text="&amp;nbsp;&lt;&lt;" Visible="False"/>
                        <br />
                        <br />
                        <asp:Button ID="btnRcvdAtTransworld" runat="server" CssClass="style_button_action"
                            Text="Forms Received (TW)" onclick="btnRcvdAtTransworld_Click" />
                        <asp:Label ID="lblRcvdAtTransworld" runat="server" Font-Bold="True" 
                            ForeColor="Red" Text="&amp;nbsp;&lt;&lt;" Visible="False"/>
                        <br />
                        <br />
                        <asp:Button ID="btnSentToWesternUnion" runat="server" CssClass="style_button_action"
                            Text="Forms Sent to WU" onclick="btnSentToWesternUnion_Click"  />
                        <asp:Label ID="lblSentToWesternUnion" runat="server" Font-Bold="True" 
                            ForeColor="Red" Text="&amp;nbsp;&lt;&lt;" Visible="False"/>
                        <br />
                        <br />
                        <asp:Button ID="btnComplete" runat="server" CssClass="style_button_action"
                            Text="Complete" onclick="btnComplete_Click" />
                        <asp:Label ID="lblComplete" runat="server" Font-Bold="True" 
                            ForeColor="Red" Text="&amp;nbsp;&lt;&lt;" Visible="False"/>
                        <br />
                        <br />
                        <br />
                        <br />
                        <asp:Button ID="btnAddNote" runat="server" CssClass="style_button_action" Text="Add Note..."
                            OnClick="btnAddNote_Click" />
                        <br />
                        <br />
                        <asp:Button ID="btnSetNextAction" runat="server" CssClass="style_button_action" Text="Set Next Action..."
                            OnClick="btnSetNextAction_Click" />
                        <br />
                        <br />
                    </td>
                    <td rowspan="6" valign="top" style="padding-left: 5px; padding-right: 5px; border-top-style: dashed;
                        border-left-style: dashed; border-right-style: dashed; border-width: 1px;">
                        <telerik:RadEditor ID="reTerminationDetails" runat="server" Font-Names="Verdana"
                            Font-Size="XX-Small" Width="100%" EditModes="Design" ToolbarMode="ShowOnFocus"
                            Enabled="False" BorderColor="#333333" Height="300px">
                            <CssFiles>
                                <telerik:EditorCssFile Value="~/WAT_RADEditorDefaultStyleSheet.css" />
                            </CssFiles>
                            <Content>
                            </Content>
                            <TrackChangesSettings CanAcceptTrackChanges="False" />
                        </telerik:RadEditor>
                    </td>
                    <td rowspan="6" valign="top" style="padding-left: 5px; padding-right: 5px; overflow: auto;
                        border-top-style: dashed; border-right-style: dashed; border-width: 1px; margin-right: 5px;">
                        <telerik:RadEditor ID="reAuditTrail" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Width="100%" EditModes="Design" ToolbarMode="ShowOnFocus" Enabled="False" BorderColor="#333333"
                            Height="300px">
                            <Content>
                            </Content>
                            <TrackChangesSettings CanAcceptTrackChanges="False" />
                        </telerik:RadEditor>
                    </td>
                </tr>
                <tr>
                    <td align="center">
                        &nbsp;
                    </td>
                </tr>
                <tr>
                    <td align="center">
                        &nbsp;
                    </td>
                </tr>
                <tr valign="top">
                    <td align="center">
                        &nbsp;
                    </td>
                </tr>
                <tr>
                    <td align="center">
                        &nbsp;
                    </td>
                </tr>
                <tr>
                    <td align="center">
                    </td>
                </tr>
                <tr>
                    <td align="center">
                        &nbsp;
                    </td>
                    <td style="border-top-style: dashed; border-width: 1px;">
                        &nbsp;
                    </td>
                    <td style="border-top-style: dashed; border-width: 1px;">
                        &nbsp;
                    </td>
                </tr>
            </table>
        </asp:Panel>
        <asp:Panel ID="pnlContactAgent" Width="99%" runat="server" BackColor="#FFCC66" CssClass="style_padded_xx_verdana">
            <table style="width: 99%">
                <tr>
                    <td style="width: 10%" align="right">
                    </td>
                    <td style="width: 25%" colspan="2">
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td colspan="3">
                        <asp:Label ID="lblLegenContactAgent" runat="server" Text="Contact Agent" CssClass="style_panel_name"
                            Font-Bold="True" Font-Names="Verdana" Font-Size="Small" />
                    </td>
                    <td>
                        &nbsp;
                    </td>
                </tr>
                <tr>
                    <td>
                        &nbsp;
                    </td>
                    <td colspan="2">
                        &nbsp;
                    </td>
                    <td>
                        &nbsp;
                    </td>
                </tr>
                <tr>
                    <td align="right">
                        <asp:Label ID="lblLegendContactAgentTelephoneNumber" runat="server" Text="Telephone Number:" />
                    </td>
                    <td colspan="2">
                        <asp:Label ID="lblTelephoneNumber" runat="server" Font-Names="Verdana" Font-Size="Medium" />
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td align="right">
                        <asp:Label ID="lblLegendContactAgentResult" runat="server" Text="Contact Result:" />
                    </td>
                    <td colspan="3">
                        <asp:RadioButton ID="rbContactAgentSuccess" runat="server" Text="Success" GroupName="ContactAgentResult" />
                        &nbsp;
                        <asp:RadioButton ID="rbContactAgentFail" runat="server" Text="Fail - further attempt required"
                            GroupName="ContactAgentResult" />
                        &nbsp;<asp:RadioButton ID="rbContactAgentAbandoned" runat="server" GroupName="ContactAgentResult"
                            Text="Termination ABANDONED" AutoPostBack="True" OnCheckedChanged="rbContactAgentAbandoned_CheckedChanged" />
                    </td>
                </tr>
                <tr>
                    <td align="right">
                        <asp:Label ID="lblLegendContactNotes" runat="server" Text="Contact Notes:" />
                    </td>
                    <td colspan="2">
                        <asp:TextBox ID="tbContactNotes" runat="server" Rows="3" TextMode="MultiLine" Width="100%"
                            Font-Names="Verdana" Font-Size="XX-Small"></asp:TextBox>
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td align="right">
                        &nbsp;
                    </td>
                    <td colspan="2">
                        <asp:LinkButton ID="lnkbtnContactAgentEditAddress" runat="server" OnClick="lnkbtnContactAgentEditAddress_Click">edit address</asp:LinkButton>
                    </td>
                    <td>
                        &nbsp;</td>
                </tr>
                <tr id="trContactAgentAddressContactName" runat="server" visible="false">
                    <td align="right">
                        <asp:Label ID="lblLegendContactAgentAddressContactName" runat="server" Text="Contact Name:" />
                    </td>
                    <td colspan="2">
                        <asp:TextBox ID="tbContactAgentAddressContactName" runat="server" Font-Names="Verdana"
                            Font-Size="XX-Small" Width="100%" MaxLength="50" />
                    </td>
                    <td>
                        &nbsp;
                    </td>
                </tr>
                <tr id="trContactAgentAddressAddr1" runat="server" visible="false">
                    <td align="right">
                        <asp:Label ID="lblLegendContactAgentAddressAddr1" runat="server" Text="Addr 1:" />
                    </td>
                    <td colspan="2">
                        <asp:TextBox ID="tbContactAgentAddressAddr1" runat="server" Font-Names="Verdana"
                            Font-Size="XX-Small" Width="100%" MaxLength="50" />
                    </td>
                    <td>
                        &nbsp;
                    </td>
                </tr>
                <tr id="trContactAgentAddressAddr2" runat="server" visible="false">
                    <td align="right">
                        <asp:Label ID="lblLegendContactAgentAddressAddr2" runat="server" Text="Addr 2:" />
                    </td>
                    <td colspan="2">
                        <asp:TextBox ID="tbContactAgentAddressAddr2" runat="server" Font-Names="Verdana"
                            Font-Size="XX-Small" Width="100%" MaxLength="50" />
                    </td>
                    <td>
                        &nbsp;
                    </td>
                </tr>
                <tr id="trContactAgentAddressAddr3" runat="server" visible="false">
                    <td align="right">
                        <asp:Label ID="lblLegendContactAgentAddressAddr3" runat="server" Text="Addr 3:" />
                    </td>
                    <td colspan="2">
                        <asp:TextBox ID="tbContactAgentAddressAddr3" runat="server" Font-Names="Verdana"
                            Font-Size="XX-Small" Width="100%" MaxLength="50" />
                    </td>
                    <td>
                        &nbsp;
                    </td>
                </tr>
                <tr id="trContactAgentAddressTownCity" runat="server" visible="false">
                    <td align="right">
                        <asp:Label ID="lblLegendContactAgentAddressTownCity" runat="server" Text="Town / City:" />
                    </td>
                    <td colspan="2">
                        <asp:TextBox ID="tbContactAgentAddressTownCity" runat="server" Font-Names="Verdana"
                            Font-Size="XX-Small" Width="100%" MaxLength="50" />
                    </td>
                    <td>
                        &nbsp;
                    </td>
                </tr>
                <tr id="trContactAgentAddressRegionPostcode" runat="server" visible="false">
                    <td align="right">
                        <asp:Label ID="lblLegendContactAgentAddressRegionPostcode" runat="server" Text="Region / Postcode:" />
                    </td>
                    <td colspan="2">
                        <asp:TextBox ID="tbContactAgentAddressRegionPostcode" runat="server" Font-Names="Verdana"
                            Font-Size="XX-Small" Width="100%" MaxLength="50" />
                    </td>
                    <td>
                        &nbsp;
                    </td>
                </tr>
                <tr>
                    <td align="right">
                        &nbsp;
                    </td>
                    <td align="center" colspan="2">
                        &nbsp;
                    </td>
                    <td>
                        &nbsp;
                    </td>
                </tr>
                <tr>
                    <td align="right">
                        &nbsp;
                    </td>
                    <td align="center">
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
                    <td align="right">
                        <asp:Label ID="Label1" runat="server" Text="Collection Date/Time:" />
                    </td>
                    <td colspan="2">
                        <telerik:RadDatePicker ID="rdtpCollectionDateTime" runat="server" >
                        </telerik:RadDatePicker>
                    </td>
                    <td>
                        &nbsp;
                    </td>
                </tr>
                <tr>
                    <td align="right">
                        &nbsp;
                    </td>
                    <td align="center">
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
                    </td>
                    <td colspan="2">
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td colspan="2">
                        <asp:Button ID="btnContactAgentFinish" runat="server" CssClass="style_button_action"
                            OnClick="btnContactAgentFinish_Click" Text="Finish" />
                        &nbsp;<asp:Button ID="btnContactAgentCancel" runat="server" OnClick="btnGeneralCancel_Click"
                            Text="Cancel" />
                    </td>
                    <td>
                    </td>
                </tr>
            </table>
        </asp:Panel>
        <asp:Panel ID="pnlFormsReceived" Width="99%" runat="server" CssClass="style_padded_xx_verdana">
            <table style="width: 100%">
                <tr>
                    <td style="width: 10%">
                    </td>
                    <td style="width: 25%">
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td colspan="2">
                        <asp:Label ID="Label6" runat="server" CssClass="style_panel_name" Font-Bold="True"
                            Font-Names="Verdana" Font-Size="Small" Text="Forms Received" />
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td align="right">
                        <asp:Label ID="Label8" runat="server" Text="Notes:" />
                    </td>
                    <td colspan="2">
                        <asp:TextBox ID="tbNotesFormsReceived" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Rows="2" TextMode="MultiLine" Width="100%"/>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td>
                        &nbsp;</td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td>
                        <asp:Button ID="btnConfirmFormsReceived" runat="server" 
                            Text="Confirm FORMS RECEIVED" Width="228px" 
                            onclick="btnConfirmFormsReceived_Click" />
                        &nbsp;<asp:Button ID="btnConfirmFormsReceivedCancel" runat="server" 
                            OnClick="btnGeneralCancel_Click" Text="Cancel" />
                    </td>
                    <td align="right">
                        &nbsp;</td>
                </tr>
            </table>
        </asp:Panel>
        <asp:Panel ID="pnlFormsSent" Width="99%" runat="server" CssClass="style_padded_xx_verdana">
            <table style="width: 100%">
                <tr>
                    <td style="width: 10%">
                    </td>
                    <td style="width: 25%">
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td colspan="2">
                        <asp:Label ID="Label9" runat="server" CssClass="style_panel_name" Font-Bold="True"
                            Font-Names="Verdana" Font-Size="Small" Text="Forms Sent to WU" />
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td align="right">
                        <asp:Label ID="Label10" runat="server" Text="Notes:" />
                    </td>
                    <td colspan="2">
                        <asp:TextBox ID="tbNotesFormsSent" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Rows="2" TextMode="MultiLine" Width="100%"/>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td>
                        &nbsp;</td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td>
                        <asp:Button ID="btnConfirmSent" runat="server" 
                            Text="Confirm FORMS SENT" Width="228px" onclick="btnConfirmSent_Click" />
                        &nbsp;<asp:Button ID="btnConfirmSentCancel" runat="server" 
                            OnClick="btnGeneralCancel_Click" Text="Cancel" />
                    </td>
                    <td align="right">
                        &nbsp;</td>
                </tr>
            </table>
        </asp:Panel>
        <asp:Panel ID="pnlTerminationComplete" Width="99%" runat="server" CssClass="style_padded_xx_verdana">
            <table style="width: 100%">
                <tr>
                    <td style="width: 10%">
                    </td>
                    <td style="width: 25%">
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td colspan="2">
                        <asp:Label ID="Label11" runat="server" CssClass="style_panel_name" Font-Bold="True"
                            Font-Names="Verdana" Font-Size="Small" Text="Termination Complete" />
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td align="right">
                        <asp:Label ID="Label12" runat="server" Text="Notes:" />
                    </td>
                    <td colspan="2">
                        <asp:TextBox ID="tbNotesTerminationComplete" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Rows="2" TextMode="MultiLine" Width="100%"/>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td>
                        &nbsp;</td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td>
                        <asp:Button ID="btnConfirmTerminationComplete" runat="server" 
                            Text="Confirm" Width="228px" 
                            onclick="btnConfirmTerminationComplete_Click" />
                        &nbsp;<asp:Button ID="btnConfirmTerminationCompleteCancel" runat="server" 
                            OnClick="btnGeneralCancel_Click" Text="Cancel" />
                    </td>
                    <td align="right">
                        &nbsp;</td>
                </tr>
            </table>
        </asp:Panel>
        <asp:Panel ID="pnlAddNote" Width="100%" runat="server" CssClass="style_padded_xx_verdana">
            <table style="width: 99%">
                <tr>
                    <td style="width: 10%">
                    </td>
                    <td style="width: 25%">
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td colspan="2">
                        <asp:Label ID="lblLegendAddNote" runat="server" CssClass="style_panel_name" Font-Bold="True"
                            Font-Names="Verdana" Font-Size="Small" Text="Add Note" />
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td align="right">
                        <asp:Label ID="lblLegendNote" runat="server" Text="Note:" />
                    </td>
                    <td colspan="2">
                        <asp:TextBox ID="tbNote" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Rows="2" TextMode="MultiLine" Width="100%"/>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td>
                        &nbsp;</td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td>
                        <asp:Button ID="btnSaveNote" runat="server" OnClick="btnSaveNote_Click" 
                            Text="Save Note" Width="150px" />
                        &nbsp;<asp:Button ID="btnSaveNoteCancel" runat="server" 
                            OnClick="btnGeneralCancel_Click" Text="Cancel" />
                    </td>
                    <td align="right">
                        &nbsp;</td>
                </tr>
            </table>
        </asp:Panel>
        <asp:Panel ID="pnlNoPermission" Width="100%" runat="server" CssClass="style_padded_xx_verdana">
            <table width="100%">
                <tr>
                    <td style="width: 10%">
                    </td>
                    <td style="width: 25%">
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td colspan="2">
                        <asp:Label ID="Label15" runat="server" CssClass="style_panel_name" Font-Bold="True"
                            Font-Names="Verdana" Font-Size="Small" Text="NO PERMISSION!!" />
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
                        &nbsp;
                    </td>
                </tr>
                <tr>
                    <td colspan="3">
                        <p align="center">
                            You are not permissioned to access this application.<br />
                            <br />
                            Please contact your system administrator to request the necessary permission.</p>
                    </td>
                </tr>
            </table>
        </asp:Panel>
        <asp:Panel ID="pnlNotifications" Width="100%" runat="server" CssClass="style_padded_xx_verdana">
            <table width="99%">
                <tr>
                    <td style="width: 10%">
                    </td>
                    <td style="width: 25%">
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td colspan="2">
                        <asp:Label ID="Label3" runat="server" CssClass="style_panel_name" Font-Bold="True"
                            Font-Names="Verdana" Font-Size="Small" Text="Notifications" />
                    </td>
                    <td>
                        &nbsp;
                    </td>
                </tr>
                <tr>
                    <td>
                        &nbsp;
                    </td>
                    <td colspan="2">
                        <asp:GridView ID="gvNotifications" runat="server" CellPadding="2" Font-Names="Verdana"
                            Font-Size="XX-Small" Width="100%" AutoGenerateColumns="False">
                            <Columns>
                                <asp:TemplateField>
                                    <ItemTemplate>
                                        <asp:Button ID="btnNotificationsEdit" runat="server" OnClick="btnNotificationsEdit_Click"
                                            Text="edit" />
                                    </ItemTemplate>
                                    <ItemStyle HorizontalAlign="Center" Width="50px" />
                                </asp:TemplateField>
                                <asp:BoundField DataField="UserAccount" HeaderText="Account" ReadOnly="True" SortExpression="UserAccount" />
                                <asp:CheckBoxField DataField="Termination_CREATE" HeaderText="CREATE  termination"
                                    ReadOnly="True" SortExpression="Termination_CREATE" />
                                <asp:CheckBoxField DataField="Termination_COMPLETE" HeaderText="COMPLETE termination"
                                    ReadOnly="True" SortExpression="Termination_COMPLETE" />
                                <asp:CheckBoxField DataField="Termination_NOTEDETAILCHANGE" HeaderText="Detail Changes"
                                    ReadOnly="True" SortExpression="Termination_NOTEDETAILCHANGE" />
                                <asp:CheckBoxField DataField="Termination_EVENT" HeaderText="Events & Notes" ReadOnly="True"
                                    SortExpression="Termination_EVENT" />
                            </Columns>
                            <EmptyDataTemplate>
                                <asp:Label ID="lblNotificationsEmptyDataMessage" runat="server" Text="No notifications defined"></asp:Label>
                            </EmptyDataTemplate>
                        </asp:GridView>
                    </td>
                </tr>
                <tr>
                    <td>
                        &nbsp;
                    </td>
                    <td>
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                        &nbsp;
                    </td>
                    <td colspan="2">
                        To modify notification settings please contact development
                        <a href="mailto:(chris.newport@transworld.eu.com">
                        (chris.newport@transworld.eu.com</a>, 0208 751 7536&nbsp;)</td>
                </tr>
                <tr>
                    <td>
                        &nbsp;</td>
                    <td colspan="2">
                        &nbsp;</td>
                </tr>
                <tr>
                    <td>
                        &nbsp;
                    </td>
                    <td>
                        <asp:Button ID="btnNotificationsClose" runat="server" 
                            OnClick="btnNotificationsClose_Click" Style="margin-left: 0px" Text="Close" 
                            Width="100px" />
                    </td>
                    <td align="right">
                        &nbsp;</td>
                </tr>
            </table>
        </asp:Panel>
        <asp:Panel ID="pnlSetNextAction" Width="99%" runat="server" CssClass="style_padded_xx_verdana">
            <table width="100%">
                <tr>
                    <td style="width: 10%">
                    </td>
                    <td style="width: 25%">
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td colspan="2">
                        <asp:Label ID="Label5" runat="server" CssClass="style_panel_name" Font-Bold="True"
                            Font-Names="Verdana" Font-Size="Small" Text="Set Next Action" />
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
                        &nbsp;
                    </td>
                </tr>
                <tr>
                    <td>
                        &nbsp;
                    </td>
                    <td>
                        <asp:DropDownList ID="ddlNextAction" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            AutoPostBack="True" OnSelectedIndexChanged="ddlNextAction_SelectedIndexChanged">
                            <asp:ListItem Value="01CONTACT AGENT">CONTACT AGENT</asp:ListItem>
                            <asp:ListItem Value="02COLLECT FORMS">COLLECT FORMS</asp:ListItem>
                            <asp:ListItem Value="03NOTIFY RECEIVED AT TRANSWORLD">NOTIFY RECEIPT</asp:ListItem>
                            <asp:ListItem Value="04DELIVER TO WESTERN UNION">COMPLETE TERMINATION</asp:ListItem>
                            <asp:ListItem Value="05COMPLETE TERMINATION">COMPLETE TERMINATION</asp:ListItem>
                        </asp:DropDownList>
                    </td>
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
                </tr>
                <tr>
                    <td>
                        &nbsp;
                    </td>
                    <td>
                        <asp:Button ID="btnSetNextActionClose" runat="server" 
                            OnClick="btnSetNextActionClose_Click" Style="margin-left: 0px" Text="Close" 
                            Width="100px" />
                    </td>
                    <td align="right">
                        &nbsp;</td>
                </tr>
            </table>
        </asp:Panel>
        <asp:Panel ID="pnlHelp" Width="100%" runat="server" CssClass="style_padded_xx_verdana">
            <table width="100%">
                <tr>
                    <td style="width: 10%">
                    </td>
                    <td style="width: 25%">
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td colspan="2">
                        <asp:LinkButton ID="lnkbtnWATHelp" runat="server" Font-Names="Verdana" 
                            Font-Size="XX-Small" 
                            OnClientClick="window.open(&quot;WATHelpGuide.pdf&quot;, &quot;WATHelp&quot;,&quot;top=10,left=10,width=700,height=450,status=no,toolbar=yes,address=no,menubar=yes,resizable=yes,scrollbars=yes&quot;);">help</asp:LinkButton>
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
                        &nbsp;
                    </td>
                </tr>
                <tr>
                    <td>
                        &nbsp;
                        </td>
                    <td>
                    </td>
                    <td>
                    </td>
                </tr>
            </table>
        </asp:Panel>
    </div>
    </form>
</body>
</html>
