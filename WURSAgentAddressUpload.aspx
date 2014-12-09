<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="FileHelpers" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<script runat="server">

    ' https://www.simple-talk.com/sql/learn-sql-server/the-merge-statement-in-sql-server-2008/
    ' http://msdn.microsoft.com/en-us/library/bb522522(v=sql.105).aspx
    
    Const INTERLOCK_TERMID As String = "----"
    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsNumeric(Session("UserKey")) Then
            Server.Transfer("session_expired.aspx")
        End If
        If Not IsPostBack Then
            Call SetTitle()
            Call CreateWURSAddressesFolder()
            Dim guidInterlockGUID As Guid = Guid.NewGuid
            psInterlockGUID = guidInterlockGUID.ToString
        End If
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
        Page.Header.Title = sTitle & "Western Union - Agent Order Profile"
    End Sub

    Protected Sub btnReadAddresses_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ReadAndCheckAddresses()
    End Sub
    
    Protected Sub ReadAndCheckAddresses()
        Const MAX_ERROR_COUNT As Int32 = 10
        Dim nLineCount As Integer
        lblError.Text = ""
        lblMessage.Text = String.Empty
        Dim sUploadFileName As String = FileUpload1.FileName
        If Path.GetExtension(sUploadFileName).ToLower <> ".csv" Then
            WebMsgBox.Show("This is not a .CSV file.\n\nDid you forget to convert it perhaps?")
            Exit Sub
        End If

        If FileUpload1.HasFile Then
            
            Call BuildAgentIDTable()
            
            Dim sUniqueFilePrefix As String = Format(Now(), "yyyymmddhhmmssff")
            psFilename = Server.MapPath("~/WURSAddresses/") & sUniqueFilePrefix & sUploadFileName
            FileUpload1.SaveAs(psFilename)
            nLineCount = nCSVLineCount(psFilename) ' return number of lines in CSV address file
            lblMessage.Text = "Read " & nLineCount.ToString & " address lines from file " & sUploadFileName & "."
            Dim WURSAddressLines As WURSAddressLine()
            Dim engine As New FileHelperEngine(GetType(WURSAddressLine))
            Try
                WURSAddressLines = DirectCast(engine.ReadFile(psFilename), WURSAddressLine())
            Catch ex As Exception
                lblMessage.Text = sUploadFileName & ": COULD NOT READ FILE. Error information: " & ex.Message
                Exit Sub
            End Try
        
            Dim nErrorCount As Int32 = 0
            Dim bLineError As Boolean
            Dim sbMessage As New StringBuilder
            Dim nBlankLineCount As Int32 = 0
            nLineCount = 1

            For Each o As WURSAddressLine In WURSAddressLines
                bLineError = False
                If o.sTermId.Length = 0 And o.sStatus.Length = 0 And o.sAgentName.Length = 0 And o.sAddress1.Length = 0 And o.sAddress2.Length = 0 And o.sAddress3.Length = 0 And o.sTownCity.Length = 0 And o.sRegion.Length = 0 And o.sPostCode.Length = 0 Then
                    nBlankLineCount += 1
                    bLineError = True
                Else
                    
                    If o.sTermId.Length <> 4 Then
                        sbMessage.Append("Line " & nLineCount & ": expected 4 character Terminal ID, found " & o.sTermId.Length & " Terminal ID.<br />")
                        bLineError = True
                        nErrorCount += 1
                    End If
                    For Each m As Match In Regex.Matches(o.sTermId, "[^A-Z0-9]")   ' http://stackoverflow.com/questions/386495/how-do-you-determine-if-a-char-is-a-letter-from-a-z
                        sbMessage.Append("Line " & nLineCount & ": found Terminal ID containing one or more non alphanumeric characters.<br />")
                        bLineError = True
                        nErrorCount += 1
                    Next
                    If Not (o.sStatus.ToLower = "active" Or o.sStatus.ToLower = "suspended" Or o.sStatus.ToLower = "agent approval") Then
                        sbMessage.Append("Line " & nLineCount & ": missing or unrecognised status code.<br />")
                        bLineError = True
                        nErrorCount += 1
                    End If
                    If o.sAddress1.Trim.Length = 0 Then
                        sbMessage.Append("Line " & nLineCount & ": missing Address 1 field.<br />")
                        bLineError = True
                        nErrorCount += 1
                    End If
                    If o.sAddress1.Length > 50 Then
                        sbMessage.Append("Line " & nLineCount & ": overlength Address 1 field.<br />")
                        bLineError = True
                        nErrorCount += 1
                    End If
                    If o.sAddress2.Length > 50 Then
                        sbMessage.Append("Line " & nLineCount & ": overlength Address 2 field.<br />")
                        bLineError = True
                        nErrorCount += 1
                    End If
                    If o.sAddress3.Length > 50 Then
                        sbMessage.Append("Line " & nLineCount & ": found overlength Address 3 field.<br />")
                        bLineError = True
                        nErrorCount += 1
                    End If
                    If o.sTownCity.Trim.Length = 0 Then
                        sbMessage.Append("Line " & nLineCount & ": missing Town/City field.<br />")
                        bLineError = True
                        nErrorCount += 1
                    End If
                    If o.sTownCity.Length > 50 Then
                        sbMessage.Append("Line " & nLineCount & ": overlength Town/City field.<br />")
                        bLineError = True
                        nErrorCount += 1
                    End If
                    If o.sRegion.Length > 50 Then
                        sbMessage.Append("Line " & nLineCount & ": overlength Region field.<br />")
                        bLineError = True
                        nErrorCount += 1
                    End If
                    If o.sPostCode.Length > 10 Then
                        sbMessage.Append("Line " & nLineCount & ": overlength Post Code field.<br />")
                        bLineError = True
                        nErrorCount += 1
                    End If
                    If o.sPostCode.Trim.Length = 0 And Not (o.sTownCity.ToLower.Contains("gibraltar")) Then
                        sbMessage.Append("Line " & nLineCount & ": missing Post Code field.<br />")
                        bLineError = True
                        nErrorCount += 1
                    End If
                    If Not bLineError AndAlso o.sStatus.ToLower = "active" Then
                        Dim sbSQL As New StringBuilder
                        sbSQL.Append("INSERT INTO ClientData_WU_AgentsTEMP (Termid, StatusDesc, AgentName, Address1, Address2, Address3, City, State, Postcode)")
                        sbSQL.Append(" VALUES (")
                        sbSQL.Append("'")
                        sbSQL.Append(o.sTermId)
                        sbSQL.Append("',")
                        sbSQL.Append("'")
                        sbSQL.Append(o.sStatus)
                        sbSQL.Append("',")
                        sbSQL.Append("'")
                        sbSQL.Append(o.sAgentName)
                        sbSQL.Append("',")
                        sbSQL.Append("'")
                        sbSQL.Append(o.sAddress1)
                        sbSQL.Append("',")
                        sbSQL.Append("'")
                        sbSQL.Append(o.sAddress2)
                        sbSQL.Append("',")
                        sbSQL.Append("'")
                        sbSQL.Append(o.sAddress3)
                        sbSQL.Append("',")
                        sbSQL.Append("'")
                        sbSQL.Append(o.sTownCity)
                        sbSQL.Append("',")
                        sbSQL.Append("'")
                        sbSQL.Append(o.sRegion)
                        sbSQL.Append("',")
                        sbSQL.Append("'")
                        sbSQL.Append(o.sPostCode)
                        sbSQL.Append("')")
                        Call ExecuteQueryToDataTable(sbSQL.ToString)
                    End If
                    If nErrorCount > MAX_ERROR_COUNT Then
                        WebMsgBox.Show("Maximum error count (" & MAX_ERROR_COUNT & ") exceeded - aborting.")
                    End If
                End If
                nLineCount += 1
            Next

            lblError.Text = sbMessage.ToString
            If lblError.Text = String.Empty Or lblError.Text.StartsWith("Ignored") Then
                btnComparisonReport.Enabled = True
                btnAddAddresses.Enabled = True
            Else
                btnComparisonReport.Enabled = False
                btnAddAddresses.Enabled = False
            End If
        Else
            WebMsgBox.Show("Specified file could not be found or file could not be processed")
            FileUpload1.Focus()
            Exit Sub
        End If

        If My.Computer.FileSystem.FileExists(sUploadFileName) Then
            My.Computer.FileSystem.DeleteFile(sUploadFileName)
        End If
    End Sub

    Private Sub CreateWURSAddressesFolder()
        Dim sPath As String = Server.MapPath("~/")
        If Not Directory.Exists(sPath & "\WURSAddresses") Then
            Directory.CreateDirectory(sPath & "\WURSAddresses")
        End If
    End Sub
        
    <DelimitedRecord(",")> _
<IgnoreFirst(1)> _
    Public Class WURSAddressLine
        <FieldQuoted("""", QuoteMode.OptionalForBoth)> _
        Public sTermId As String
        
        <FieldQuoted("""", QuoteMode.OptionalForBoth)> _
        Public sStatus As String
        
        <FieldQuoted("""", QuoteMode.OptionalForBoth)> _
        Public sAgentName As String
        
        <FieldQuoted("""", QuoteMode.OptionalForBoth)> _
        Public sAddress1 As String
        
        <FieldQuoted("""", QuoteMode.OptionalForBoth)> _
        Public sAddress2 As String
        
        <FieldQuoted("""", QuoteMode.OptionalForBoth)> _
        Public sAddress3 As String

        <FieldQuoted("""", QuoteMode.OptionalForBoth)> _
        Public sTownCity As String
        
        <FieldQuoted("""", QuoteMode.OptionalForBoth)> _
        Public sRegion As String
        
        <FieldQuoted("""", QuoteMode.OptionalForBoth)> _
        Public sPostCode As String

    End Class

    Protected Function GetAgentCurrentAddress(ByVal sTermID As String) As DataRow
        GetAgentCurrentAddress = Nothing
        Dim sSQL As String = "SELECT * FROM ClientData_WU_Agents WHERE TermID = '" & sTermID & "'"
        Dim dtAgent As DataTable = ExecuteQueryToDataTable(sSQL)
        If dtAgent.Rows.Count > 0 Then
            If dtAgent.Rows.Count = 1 Then
                GetAgentCurrentAddress = dtAgent.Rows(0)
            Else
                GetAgentCurrentAddress = dtAgent.Rows(0)
                WebMsgBox.Show("More than one Agent record retrieved for TermID " & sTermID & "!!")
            End If
        End If
    End Function

    Protected Sub BuildAgentIDTable()
        Dim sSQL As String = "IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].ClientData_WU_AgentsTEMP') AND type in (N'U')) DROP TABLE [dbo].ClientData_WU_AgentsTEMP CREATE TABLE [dbo].[ClientData_WU_AgentsTEMP]([Termid] [varchar](4) NOT NULL, [StatusDesc] [varchar](50) NOT NULL, [AgentName] [varchar](50) NOT NULL, [Address1] [varchar](50) NOT NULL, [Address2] [varchar](50) NOT NULL, [Address3] [varchar](50) NOT NULL, [City] [varchar](50) NOT NULL, [State] [varchar](50) NOT NULL, [Postcode] [varchar](50) NOT NULL) ON [PRIMARY]"
        Call ExecuteQueryToDataTable(sSQL)
        sSQL = "INSERT INTO ClientData_WU_AgentsTEMP (TermID, StatusDesc, AgentName, Address1, Address2, Address3, City, State, Postcode) VALUES ('" & INTERLOCK_TERMID & "', '" & psInterlockGUID & "', '', '', '', '', '', '', '')"
        Call ExecuteQueryToDataTable(sSQL)
    End Sub
    
    Protected Function nCSVLineCount(ByVal sFileName As String) As Integer ' return number of lines in CSV order file
        Dim sr As New StreamReader(sFileName)
        Dim nLineCount As Integer = 0
        Do While sr.Peek >= 0
            sr.ReadLine()
            nLineCount += 1
        Loop
        sr.Close()
        nCSVLineCount = nLineCount
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

    Public Function ExecuteStoredProcedureToDataTable(ByVal sp_name As String, Optional ByVal IListPrams As List(Of SqlParameter) = Nothing) As DataTable
        Dim oDataTable As New DataTable
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter(sp_name, oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        If Not IListPrams Is Nothing AndAlso IListPrams.Count > 0 Then
            oAdapter.SelectCommand.Parameters.AddRange(IListPrams.ToArray)
        End If
        Try
            oAdapter.Fill(oDataTable)
        Catch ex As Exception
            WebMsgBox.Show(ex.Message.ToString())
        End Try
        ExecuteStoredProcedureToDataTable = oDataTable
    End Function

    Property psFilename() As String
        Get
            Dim o As Object = ViewState("WAAL_Filename")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("WAAL_Filename") = Value
        End Set
    End Property

    Property psInterlockGUID() As String
        Get
            Dim o As Object = ViewState("WAAL_InterlockGUID")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("WAAL_InterlockGUID") = Value
        End Set
    End Property

    Protected Sub btnComparisonReport_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call GenerateComparisonReport()
    End Sub

    Protected Sub GenerateComparisonReport()
        Dim sbText As New StringBuilder
        Call AddHTMLPreamble(sbText, "WURS Addresses Comparison Report")
        sbText.Append(Bold("WURS ADDRESSES COMPARISON REPORT"))
        Call NewLine(sbText)
        sbText.Append("report generated " & DateTime.Now)
        Call NewLine(sbText)
        Call NewLine(sbText)
        sbText.Append("This report compares the list of active users in the incoming WURS Agent address spreadsheet with the Agents defined in the Transworld database.")
        Call NewLine(sbText)
        Call NewLine(sbText)
        sbText.Append("<hr />")
        Call NewLine(sbText)
        Const TABLE_HEADER As String = "<table width=""100%"">"
        Const TABLE_TROPEN As String = "<tr>"
        Const TABLE_TRCLOSE As String = "</tr>"
        Const TABLE_TDOPEN As String = "<td>"
        Const TABLE_TDCLOSE As String = "</td>"
        Const TABLE_FOOTER As String = "</table>"
        Dim dtUsers As DataTable
        dtUsers = ExecuteQueryToDataTable("SELECT UserID, FirstName, LastName, ISNULL(CAST(REPLACE(CONVERT(VARCHAR(11), LastLogon, 106), ' ', '-') AS varchar(20)),'(never)') 'LastLogon', ISNULL(CAST(REPLACE(CONVERT(VARCHAR(11), LastUpdatedOn, 106), ' ', '-') AS varchar(20)),'(never)') 'LastUpdatedOn', CASE WHEN UserPermissions & 32768 <> 0 THEN 'INTERNAL' ELSE 'AGENT' END 'UserType' FROM UserProfile WHERE UserId IN (SELECT UserID FROM UserProfile WHERE CustomerKey = 579 AND Status = 'Active' EXCEPT SELECT TermID FROM ClientData_WU_AgentsTEMP WHERE TermID <> '" & INTERLOCK_TERMID & "') ORDER BY UserType, LastLogon")
        sbText.Append(Bold("1. Users active in Transworld database but not in incoming address list (" & dtUsers.Rows.Count & "):"))
        Call NewLine(sbText)
        If dtUsers.Rows.Count > 0 Then
            sbText.Append(TABLE_HEADER)
            sbText.Append(TABLE_TROPEN)
            sbText.Append(TABLE_TDOPEN)
            sbText.Append(Bold("User ID"))
            sbText.Append(TABLE_TDCLOSE)
            sbText.Append(TABLE_TDOPEN)
            sbText.Append(Bold("First Name"))
            sbText.Append(TABLE_TDCLOSE)
            sbText.Append(TABLE_TDOPEN)
            sbText.Append(Bold("Last Name"))
            sbText.Append(TABLE_TDCLOSE)
            sbText.Append(TABLE_TDOPEN)
            sbText.Append(Bold("Last Logon"))
            sbText.Append(TABLE_TDCLOSE)
            sbText.Append(TABLE_TDOPEN)
            sbText.Append(Bold("Last Updated On"))
            sbText.Append(TABLE_TDCLOSE)
            sbText.Append(TABLE_TDOPEN)
            sbText.Append(Bold("User Type"))
            sbText.Append(TABLE_TDCLOSE)
            sbText.Append(TABLE_TRCLOSE)
            For Each drUser As DataRow In dtUsers.Rows
                sbText.Append(TABLE_TROPEN)
                sbText.Append(TABLE_TDOPEN)
                sbText.Append(drUser("UserID"))
                sbText.Append(TABLE_TDCLOSE)
                sbText.Append(TABLE_TDOPEN)
                sbText.Append(drUser("FirstName"))
                sbText.Append(TABLE_TDCLOSE)
                sbText.Append(TABLE_TDOPEN)
                sbText.Append(drUser("LastName"))
                sbText.Append(TABLE_TDCLOSE)
                sbText.Append(TABLE_TDOPEN)
                sbText.Append(drUser("LastLogon"))
                sbText.Append(TABLE_TDCLOSE)
                sbText.Append(TABLE_TDOPEN)
                sbText.Append(drUser("LastUpdatedOn"))
                sbText.Append(TABLE_TDCLOSE)
                sbText.Append(TABLE_TDOPEN)
                sbText.Append(drUser("UserType"))
                sbText.Append(TABLE_TDCLOSE)
                sbText.Append(TABLE_TRCLOSE)
            Next
            sbText.Append(TABLE_FOOTER)
        Else
            Call NewLine(sbText)
            sbText.Append("NONE")
            Call NewLine(sbText)
        End If
        Call NewLine(sbText)
        Call NewLine(sbText)
        Call NewLine(sbText)
        dtUsers = ExecuteQueryToDataTable("SELECT TermID, AgentName FROM ClientData_WU_AgentsTEMP WHERE TermID IN (SELECT TermID FROM ClientData_WU_AgentsTEMP EXCEPT SELECT UserID FROM UserProfile WHERE CustomerKey = 579 AND Status = 'Active') AND NOT Termid = '" & INTERLOCK_TERMID & "' ORDER BY AgentName")
        sbText.Append(Bold("2. Agents present in incoming address list but either missing from Transworld database or marked SUSPENDED (" & dtUsers.Rows.Count & "):"))
        Call NewLine(sbText)
        If dtUsers.Rows.Count > 0 Then
            sbText.Append(TABLE_HEADER)
            sbText.Append(TABLE_TROPEN)
            sbText.Append(TABLE_TDOPEN)
            sbText.Append(Bold("Terminal ID"))
            sbText.Append(TABLE_TDCLOSE)
            sbText.Append(TABLE_TDOPEN)
            sbText.Append(Bold("Agent Name"))
            sbText.Append(TABLE_TDCLOSE)
            sbText.Append(TABLE_TRCLOSE)
            For Each drUser As DataRow In dtUsers.Rows
                sbText.Append(TABLE_TROPEN)
                sbText.Append(TABLE_TDOPEN)
                sbText.Append(drUser("TermID"))
                sbText.Append(TABLE_TDCLOSE)
                sbText.Append(TABLE_TDOPEN)
                sbText.Append(drUser("AgentName"))
                sbText.Append(TABLE_TDCLOSE)
                sbText.Append(TABLE_TRCLOSE)
            Next
            sbText.Append(TABLE_FOOTER)
            Call NewLine(sbText)
        Else
            Call NewLine(sbText)
            sbText.Append("NONE")
        End If
        sbText.Append("[end]")
        Call AddHTMLPostamble(sbText)
        Call ExportData(sbText.ToString, "WURS_Addresses_Comparison_Report")
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
        sbText.Append("body { font-family: Verdana; font-size : xx-small } ")
        sbText.Append("table { font-family: Verdana; font-size : xx-small } ")
        sbText.Append("</style></head><body>")
    End Sub
   
    Protected Sub AddHTMLPostamble(ByRef sbText As StringBuilder)
        sbText.Append("</body></html>")
    End Sub

    Protected Sub btnAddAddresses_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call AddAddresses()
    End Sub
    
    Protected Sub AddAddresses()
        Dim sSQL As String
        sSQL = "SELECT StatusDesc FROM ClientData_WU_AgentsTEMP WHERE TermID = '" & INTERLOCK_TERMID & "'"
        Dim dtInterlock As DataTable = ExecuteQueryToDataTable(sSQL)
        If dtInterlock.Rows.Count <> 1 OrElse dtInterlock.Rows(0).Item(0) <> psInterlockGUID Then
            Call BuildAgentIDTable()
            WebMsgBox.Show("There was a problem processing the addresses.\n\nEither the addresses have not been read correctly, or the addresses are left over from a previous invocation of the utility.\n\nPlease restart the utility and read the addresses again.")
            lblError.Text = "There was a problem processing the addresses.<br /><br />Either the addresses have not been read correctly, or the addresses are left over from a previous invocation of the utility.<br /><br />Please restart the utility and read the addresses again."
            lblMessage.Text = String.Empty
            Exit Sub
        End If
        sSQL = "DELETE FROM ClientData_WU_AgentsTEMP WHERE TermID = '" & INTERLOCK_TERMID & "'"
        Call ExecuteQueryToDataTable(sSQL)
        '
        ' IF/WHEN MODIFYING THE STATEMENT BELOW (EG TO INCLUDE LOGIC TO MARK UNMATCHED ENTRIES AS 'SUSPENDED'
        ' DON'T FORGET TO MOVE THE SEMI-COLON TO THE END OF THE STATEMENT (SEMI-COLON *IS* REQUIRED HERE !!)
        '
        sSQL = "MERGE ClientData_WU_Agents "
        sSQL &= "USING ClientData_WU_AgentsTEMP "
        sSQL &= "ON ClientData_WU_Agents.TermID = ClientData_WU_AgentsTEMP.TermID "
        sSQL &= "WHEN NOT MATCHED BY TARGET "
        sSQL &= "THEN INSERT (TermID, AgentName, Address1, Address2, Address3, City, State, Postcode, LastChangedOn, StatusDesc, PCEquipped, NetworkAgent, PendingDeletionFlag, WUAccountNumber, CompanyName, Contact, PhoneNumber, IsDeleted, LastChangedBy) VALUES (ClientData_WU_AgentsTEMP.TermID, ClientData_WU_AgentsTEMP.AgentName, ClientData_WU_AgentsTEMP.Address1, ClientData_WU_AgentsTEMP.Address2, ClientData_WU_AgentsTEMP.Address3, ClientData_WU_AgentsTEMP.City, ClientData_WU_AgentsTEMP.State, ClientData_WU_AgentsTEMP.Postcode, GETDATE(), '', '', '', '', '', '', '', '', '0', '0') "
        sSQL &= "WHEN MATCHED "
        sSQL &= "THEN UPDATE SET AgentName = ClientData_WU_AgentsTEMP.AgentName, Address1 = ClientData_WU_AgentsTEMP.Address1, Address2 = ClientData_WU_AgentsTEMP.Address2, Address3 = ClientData_WU_AgentsTEMP.Address3, City = ClientData_WU_AgentsTEMP.City, State = ClientData_WU_AgentsTEMP.State, Postcode = ClientData_WU_AgentsTEMP.Postcode, LastChangedOn = GETDATE() "
        sSQL &= "; "   ' !!!!!!!!!!!! THE SEMI-COLON (see above) !!!!!!!!!!!
        sSQL &= "--WHEN NOT MATCHED BY SOURCE "
        sSQL &= "--THEN "
        Call ExecuteQueryToDataTable(sSQL)

        sSQL = "DELETE FROM ClientData_WU_AgentsTEMP"
        Call ExecuteQueryToDataTable(sSQL)
        WebMsgBox.Show("The WURS Agent address list has been updated.")
        lblError.Text = String.Empty
        lblMessage.Text = String.Empty
        btnReadAddresses.Enabled = False
        btnComparisonReport.Enabled = False
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

</script>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head runat="server">
        <title></title>
    </head>
    <body>
        <form id="form1" runat="server">
        <asp:PlaceHolder ID="PlaceHolderForScriptManager" runat="server" />
        <main:header ID="ctlHeader" runat="server" />
            <div>
                <table style="width: 100%">
                    <tr>
                        <td colspan="3">
                            &nbsp;<asp:Label ID="lblLegendTitle" runat="server" Font-Bold="True" Font-Names="Verdana"
                                       Font-Size="Small">
                                WURS Agent Address Loader
                            </asp:Label>
                            &nbsp;
                            &nbsp;
                        </td>
                    </tr>
                    <tr>
                        <td align="right">
                            &nbsp;
                        </td>
                        <td valign="middle">
                            &nbsp;
                        </td>
                        <td>
                            &nbsp;
                        </td>
                    </tr>
                    <tr>
                        <td>
                            &nbsp;
                            <asp:Label ID="Label1" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Select Address CSV file:"></asp:Label>
                            &nbsp;
                            <asp:FileUpload ID="FileUpload1" runat="server" Font-Names="Verdana" Font-Size="X-Small" Font-Bold="true"
                                            Width="300px" />
                        </td>
                        <td valign="middle">
                            &nbsp;
                        </td>
                        <td>
                            &nbsp;
                        </td>
                    </tr>
                    <tr>
                        <td>
                            &nbsp;<asp:Button ID="btnReadAddresses" runat="server" Text="Read Addresses"
                                        Width="200px" onclick="btnReadAddresses_Click" />
                        </td>
                        <td valign="middle">
                            &nbsp;
                        </td>
                        <td>
                            &nbsp;
                        </td>
                    </tr>
                    <tr>
                        <td>
                            &nbsp;<asp:Label ID="lblMessage" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="Small" />
                        </td>
                        <td valign="middle">
                            &nbsp;</td>
                        <td>
                            &nbsp;</td>
                    </tr>
                    <tr>
                        <td>
                            &nbsp;<asp:Button ID="btnComparisonReport" runat="server" Text="Comparison Report..." Width="200px" Enabled="False" onclick="btnComparisonReport_Click"/>
                        </td>
                        <td valign="middle">
                            &nbsp;</td>
                        <td>
                            &nbsp;</td>
                    </tr>
                    <tr>
                        <td>
                            &nbsp;<asp:Button ID="btnAddAddresses" runat="server" Text="Add Addresses" Width="200px" Enabled="False" onclick="btnAddAddresses_Click"/>
                        </td>
                        <td valign="middle">
                            &nbsp;
                        </td>
                        <td>
                            &nbsp;
                        </td>
                    </tr>
                    <tr>
                        <td colspan="3">
                            &nbsp;
                            <asp:Label ID="lblError" runat="server" ForeColor="Red" Font-Bold="True" Font-Names="Verdana" Font-Size="Small" />
                        </td>
                    </tr>
                    <tr>
                        <td colspan="3">
                            &nbsp;
                            <br />
                            &nbsp;
                        </td>
                    </tr>
                    <tr>
                        <td colspan="3">
                            <hr />
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
                    <tr id="trMaintenance01" runat="server" visible="false">
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
                    <tr id="trMaintenance02" runat="server" visible="false">
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
                            <label>
                                &nbsp;
                            </label>
                        </td>
                        <td>
                            &nbsp;
                        </td>
                        <td>
                            &nbsp;
                        </td>
                    </tr>
                    <tr id="trMaintenance03" runat="server" visible="false">
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
            </div>
        </form>
    </body>
</html>
