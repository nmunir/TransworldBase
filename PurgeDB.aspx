<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Collections.Generic" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">
    ' NOTES
    
    ' Data goes back to October 2000
    ' Version 2 LogisticBooking starts on 3FEB2006
    
    ' REPORTS
    ' Count of Consignments by (live) customer
    ' Customers by last transaction date
    ' Global address book usage by customer
    
    ' TO DO
    ' log all table changes to database so we can check deletions, ie before counts, after counts

    ' purge LogisticProductSerialNumber by LogisticBookingKey
    ' if purging before 3FEB2006 then purge CustomerDestinationAddress
    
    ' Also purge...
    ' AGENT through BillingID
    ' CM
    ' ConsignmentEventAlert
    ' ConsignmentTempStatus    

    ' EMPTY TABLES AT 6NOV09
    ' AgentNote
    ' Airline
    ' AirlineNote
    ' ArchiveTable
    
    ' LogisticAncillaryCharge
    ' LogisticDispatch
    ' LogisticGoodsIn
    ' LogisticMovementTrackingStage
    
    ' TABLES THAT LOOK AS THOUGH THEY ARE NOT USED, BUT NOT CHECKED
    ' AutoMailerStatus
    ' AutoPickerStatus
    ' CustomLetter - NEED TO REMOVE CUSTOM LETTER CODE
    ' CustomLetterTemplate - NEED TO REMOVE CUSTOM LETTER CODE
    ' LANDGControl
    ' Projects
    ' ProjectDocs
    
    ' TABLES TO PROCESS THAT ARE PROBABLY NOT YET BEING PROCESSED
    ' AdhocFulfilmentRequest
    ' ConsignmentExport
    ' ConsignmentTempStatus
    
    ' CalendarManagedEventNote
    ' CalendarManagedItemDays
    ' CalendarManagedItemEvent
    
    ' OTHER NOTES
    ' UP_UserPermissionGroups - PROBABLY THE ONLY TABLE USED IN THE UP_ SERIES
    ' UP_UserPermissions
    ' UP_ProductPermissionGroups ?????????

    ' Change User Profile so that if deactivated is reactivated, UPP records are recreated (and set to no permission).
    ' Delete all UPP records where user is inactive
    
    Const DB_TIMEOUT As Integer = 1000
    
    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    Private dtDeletionList As DataTable
    Dim gnCustomerKey As Integer
    Dim gdtFinish As DateTime
    Dim alDeletionList As New ArrayList
    Dim gnRecordCount As Integer
    Dim gnPageTimeout As Integer
    Dim gnCustomerCount As Integer
    Private gnCustomersProcessedThisRun As Integer = 0

    Protected Sub Page_Load(ByVal Source As Object, ByVal E As EventArgs)
        If Not Page.IsPostBack Then
            tbLog.Text = String.Empty
        End If
    End Sub
    
    Protected Sub btnPurgeDatabase_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call PurgeDatabase()
    End Sub
    
    Protected Sub Log(ByVal sLogEntry As String)
        tbLog.Text += sLogEntry & Environment.NewLine
        If sLogEntry.Contains("exceeded") Then
            Call Log("This run processed " & gnCustomersProcessedThisRun & " customer(s).")
        End If
    End Sub
    
    Protected Sub PurgeDatabase()
        Call CalculateDuration()
        Call GetDeletionList()
        Call ProcessDeletionList()
    End Sub
    
    Protected Sub CalculateDuration()
        Dim dtStart As DateTime = Now
        Dim tsDuration As TimeSpan = New System.TimeSpan(0, 0, CInt(ddlDuration.SelectedValue))
        gdtFinish = dtStart.Add(tsDuration)
    End Sub
    
    Protected Sub GetDeletionList()
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataReader As SqlDataReader = Nothing
        Dim sSQL As String = "SELECT CustomerKey FROM Customer WHERE (CustomerStatusId <> 'ACTIVE') OR DeletedFlag = 'Y' OR CustomerAccountCode LIKE 'DEMO2005%' OR CustomerAccountCode LIKE 'DEMO2008A' OR CustomerAccountCode = 'DEMO1' OR CustomerAccountCode = 'DEMO2' OR CustomerAccountCode = 'DEMO3'"
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            If oDataReader.HasRows Then
                alDeletionList = New ArrayList
                While oDataReader.Read
                    alDeletionList.Add(oDataReader(0))
                End While
            Else
            End If
        Catch ex As Exception
            WebMsgBox.Show("Error in GetDeletionList: " & ex.Message)
        Finally
            oConn.Close()
            Call Log(alDeletionList.Count & " customers to process")
        End Try
    End Sub
    
    Protected Sub ProcessDeletionList()
        Call Log("Starting run at " & DateTime.Now.ToLongTimeString)
        For Each nCustomerKey As Integer In alDeletionList
            gnCustomerKey = nCustomerKey
            gnCustomerCount += 1
            Call Log("")
            Call Log("PROCESSING CUSTOMER " & gnCustomerKey & " (" & gnCustomerCount & " of " & alDeletionList.Count & ")")

            Call ProcessTable("AddressDistributionLists")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If

            Call ProcessLinkedTableByCustomerKey("Consignment", "[key]", "ConsignmentChange", "ConsignmentKey")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If

            Call ProcessLinkedTableByCustomerKey("Consignment", "[key]", "ConsignmentCost", "ConsignmentKey")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If

            Call ProcessLinkedTableByCustomerKey("Consignment", "[key]", "ConsignmentNote", "ConsignmentKey")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If

            Call ProcessLinkedTableByCustomerKey("Consignment", "[key]", "ConsignmentRoute", "ConsignmentKey")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If

            Call ProcessLinkedTableByCustomerKey("Consignment", "[key]", "ConsignmentSplit", "MasterConsignmentKey")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If

            Call ProcessLinkedTableByCustomerKey("Consignment", "[key]", "ConsignmentTrackingStage", "ConsignmentKey")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If

            Call ProcessTable("Consignment")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If

            Call ProcessTable("ConsignmentExport")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If

            Call ProcessTable("ConsignmentExportParams")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If

            Call ProcessTable("Contact")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If

            Call ProcessTable("CostCentreLookup")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If

            Call ProcessLinkedTableByCustomerKey("CourierBooking", "[key]", "CourierBookingComment", "CourierBookingKey")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If

            Call ProcessLinkedTableByCustomerKey("CourierBooking", "[key]", "CourierBookingTrackingStage", "CourierBookingKey")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If

            Call ProcessTable("CourierBooking")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If

            Call ProcessTable("Customer30DayQuotation")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If

            Call ProcessTable("CustomerAddressBookProfile")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If

            Call ProcessTable("CustomerBilling")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If

            Call ProcessTable("CustomerCollectionAddress")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If

            Call ProcessTable("CustomerComment")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If

            Call ProcessLinkedTableByCustomerKey("CustomerContact", "[key]", "CustomerContactName", "CustomerContactNameKey")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If

            Call ProcessTable("CustomerContact")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If

            Call ProcessTable("CustomerDataFeed")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If

            Call ProcessTable("CustomerDestinationAddress")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If
            
            ' DELETE CustomerPageContent

            Call ProcessTable("CustomerRegularCollection")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If
            
            ' DELETE CustomerStorage & CustomerStorageIndex
            
            Call ProcessTable("Fulfilment")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If
            
            Call ProcessLinkedTableByCustomerKey("FulfilmentJob", "[key]", "FulfilmentJobComment", "FulfilmentKey")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If

            Call ProcessLinkedTableByCustomerKey("FulfilmentJob", "[key]", "FulfilmentJobNote", "FulfilmentKey")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If

            Call ProcessLinkedTableByCustomerKey("FulfilmentJob", "[key]", "FulfilmentJobStage", "FulfilmentKey")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If

            Call ProcessLinkedTableByCustomerKey("FulfilmentJob", "[key]", "FulfilmentJobTrackingStage", "FulfilmentKey")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If

            Call ProcessTable("FulfilmentJob")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If
            
            Call ProcessLinkedTableByCustomerKey("GlobalAddressBook", "[key]", "UserAddressBookProfile", "GlobalAddressKey")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If
            
            Call ProcessTable("GlobalAddressBook")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If
            
            Call ProcessLinkedTableByCustomerKey("LogisticBooking", "LogisticBookingKey", "LogisticBookingTracking", "StockBookingKey")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If

            Call ProcessTable("LogisticBooking")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If
            
            Call ProcessLinkedTableByCustomerKey("LogisticProduct", "LogisticProductKey", "LogisticAssociatedProduct", "LogisticProductKey")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If

            ' need to check here if any product has a non-zero quantity

            Call ProcessLinkedTableByCustomerKey("LogisticProduct", "LogisticProductKey", "LogisticProductAuthorisable", "LogisticProductAuthorisable.LogisticProductKey")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If

            Call ProcessLinkedTableByCustomerKey("LogisticProduct", "LogisticProductKey", "LogisticProductLocation", "LogisticProductLocation.LogisticProductKey")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If

            Call ProcessLinkedTableByCustomerKey("LogisticProduct", "LogisticProductKey", "LogisticProductSerialNumber", "LogisticProductSerialNumber.LogisticProductKey")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If

            Call ProcessLinkedTableByCustomerKey("LogisticProduct", "LogisticProductKey", "LogisticProductTracking", "ProductKey")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If

            Call ProcessLinkedTableByCustomerKey("LogisticProduct", "LogisticProductKey", "StorageItem", "ProductKey")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If
            
            Call ProcessTable("LogisticProduct")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If

            Call ProcessTable("StorageCharge")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If
            
            Call ProcessTable("LogisticProductAuthorisation")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If
             
            ' DELETE LogisticCneeEmailAlerts, LogisticCnorEmailAlerts
            ' DELETE LogisticDispatch, LogisticGoodsIn
            
            Call ProcessTable("LogisticMovement")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If
             
            ' DELETE LogisticMovementTrackingStage DELETE ALL ENTRIES MANUALLY - THIS TABLE NO LONGER USED
            ' DELETE LogisticPicking - EMPTY TABLE - BELIEVE NO LONGER USED
            ' DELETE MANConsignmentsProcessed - EMPTY TABLE - BELIEVE NO LONGER USED
            
            Call ProcessTable("RegularCollection")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If
             
            ' DELETE UserEmailHistory, UserPickHistory, UserPickProfile
 
            Call ProcessLinkedTableByCustomerKey("UserProfile", "[key]", "UserProductProfile", "UserKey")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If

            Call ProcessLinkedTableByCustomerKey("UserProfile", "[key]", "UserStockAlertProfile", "UserKey")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If

            Call ProcessLinkedTableByCustomerKey("UserProfile", "[key]", "UserConsignmentAlertProfile", "UserKey")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If

            Call ProcessLinkedTableByCustomerKey("UserProfile", "[key]", "UserAddressBookProfile", "UserKey")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If

            Call ProcessTable("UserProfile")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If
            
            ' DELETE TerminalAgent, TimedDelivery, TrackingID, SystemUserProfile, SystemUserNote
            
            If Not ExecuteNonQuery("DELETE FROM Customer WHERE CustomerKey = " & gnCustomerKey) Then
                WebMsgBox.Show("Error deleting customer record")
                Exit Sub
            Else
                Call Log("DELETED customer " & gnCustomerKey)
                gnCustomersProcessedThisRun += 1
            End If
            
        Next
        
        For i As Integer = 1 To 1
            
            Call PurgeSystemUserAuditHistory()
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If

            ' ExpiredPassword: DELETE FROM ExpiredPassword WHERE UserKey IN (SELECT UserKey FROM Users WHERE DeletedFlag = 'Y')
            If Not ExecuteNonQuery("DELETE FROM ExpiredPassword WHERE UserKey IN (SELECT UserKey FROM Users WHERE DeletedFlag = 'Y')") Then
                WebMsgBox.Show("Error deleting ExpiredPassword record")
                Exit Sub
            End If
            
            Call ProcessTableWhereDeletedFlagY("Users")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If
            
            Call ProcessTableWhereDeletedFlagY("WarehouseSection")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If
            
            Call ProcessTableWhereDeletedFlagY("WarehouseRack")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If
            
            Call ProcessTableWhereDeletedFlagY("WarehouseBay")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If
            
            Call ProcessTableWhereDeletedFlagY("Warehouse")
            If Now > gdtFinish Then
                Call Log("Maximum run time exceeded")
                Exit For
            End If
        Next

        Call Log("Finishing run at " & DateTime.Now.ToLongTimeString)
        Call Log("Total records deleted: " & gnRecordCount)
    End Sub
    
    Protected Sub ProcessLinkedTableByCustomerKeyJoined(ByVal sLinkedTable As String, ByVal sForeignKey As String, ByVal sMainTable As String, ByVal sKey As String)
        Call Log("Processing table " & sLinkedTable & " (joined to table " & sMainTable & ")")
        Call LogRecordCount(sLinkedTable)
        Dim sSQL As String = "DELETE FROM " & sLinkedTable & " WHERE " & sForeignKey & " IN (SELECT " & sKey & " FROM " & sMainTable & " WHERE CustomerKey = " & gnCustomerKey & ")"
        Dim nResult As Integer
        Dim sLogMessage As String = sLinkedTable & ": "
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        Try
            oConn.Open()
            oCmd.Connection = oConn
            nResult = oCmd.ExecuteNonQuery()
            gnRecordCount += nResult
        Catch ex As Exception
            WebMsgBox.Show("Error in ProcessLinkedTableByCustomerKey processing " & sLinkedTable & ": " & ex.Message)
            Throw New Exception("Error in ProcessLinkedTableByCustomerKey processing " & sLinkedTable & ": " & ex.Message)
        Finally
            oConn.Close()
            sLogMessage += "deleted " & nResult & " entries"
            If nResult > 0 Then
                sLogMessage = ">> " & sLogMessage
                Call LogRecordCount(sLinkedTable)
            End If
            Call Log(sLogMessage)
        End Try
    End Sub
    
    Protected Sub ProcessLinkedTableByCustomerKey(ByVal sMainTable As String, ByVal sKey As String, ByVal sLinkedTable As String, ByVal sForeignKey As String)
        Call Log("Processing table " & sLinkedTable & " (linked to table " & sMainTable & ")")
        Call LogRecordCount(sLinkedTable)
        Dim sSQL As String = "DELETE FROM " & sLinkedTable & " WHERE " & sForeignKey & " IN (SELECT " & sKey & " FROM " & sMainTable & " WHERE CustomerKey = " & gnCustomerKey & ")"
        Dim nResult As Integer
        Dim sLogMessage As String = sLinkedTable & ": "
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        Try
            oConn.Open()
            oCmd.Connection = oConn
            nResult = oCmd.ExecuteNonQuery()
            gnRecordCount += nResult
        Catch ex As Exception
            WebMsgBox.Show("Error in ProcessLinkedTableByCustomerKey processing " & sLinkedTable & ": " & ex.Message)
            Throw New Exception("Error in ProcessLinkedTableByCustomerKey processing " & sLinkedTable & ": " & ex.Message)
        Finally
            oConn.Close()
            sLogMessage += "deleted " & nResult & " entries"
            If nResult > 0 Then
                sLogMessage = ">> " & sLogMessage
                Call LogRecordCount(sLinkedTable)
            End If
            Call Log(sLogMessage)
        End Try
    End Sub
    
    Protected Sub ProcessTable(ByVal sTableName As String)
        Call Log("Processing table " & sTableName)
        Call LogRecordCount(sTableName)
        Dim sSQL As String = "DELETE FROM " & sTableName & " WHERE CustomerKey = " & gnCustomerKey
        Dim nResult As Integer
        Dim sLogMessage As String = sTableName & ": "
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        oCmd.CommandTimeout = DB_TIMEOUT
        Try
            oConn.Open()
            oCmd.Connection = oConn
            nResult = oCmd.ExecuteNonQuery()
            gnRecordCount += nResult
        Catch ex As Exception
            WebMsgBox.Show("Error in ProcessTable processing " & sTableName & ": " & ex.Message)
            Throw New Exception("Error in ProcessTable processing " & sTableName & ": " & ex.Message)
            Exit Sub
        Finally
            oConn.Close()
            sLogMessage += "deleted " & nResult & " entries"
            If nResult > 0 Then
                sLogMessage = ">> " & sLogMessage
                Call LogRecordCount(sTableName)
            End If
            Call Log(sLogMessage)
        End Try
    End Sub
    
    Protected Sub ProcessTableWhereDeletedFlagY(ByVal sTableName As String)
        Call Log("Processing table (deleted flag = Y) " & sTableName)
        Call LogRecordCount(sTableName)
        Dim sSQL As String = "DELETE FROM " & sTableName & " WHERE DeletedFlag = 'Y'"
        Dim nResult As Integer
        Dim sLogMessage As String = sTableName & ": "
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        oCmd.CommandTimeout = DB_TIMEOUT
        Try
            oConn.Open()
            oCmd.Connection = oConn
            nResult = oCmd.ExecuteNonQuery()
            gnRecordCount += nResult
        Catch ex As Exception
            WebMsgBox.Show("Error in ProcessTableWhereDeletedFlagY processing " & sTableName & ": " & ex.Message)
            Throw New Exception("Error in ProcessTableWhereDeletedFlagY processing " & sTableName & ": " & ex.Message)
        Finally
            oConn.Close()
            sLogMessage += "deleted " & nResult & " entries"
            If nResult > 0 Then
                sLogMessage = ">> " & sLogMessage
                Call LogRecordCount(sTableName)
            End If
            Call Log(sLogMessage)
        End Try
    End Sub
    
    Protected Sub PurgeSystemUserAuditHistory()
        Dim sSQL As String = "DELETE FROM SystemUserAuditHistory WHERE AddedOn < (GETDATE() - 7)"
        Dim nResult As Integer
        Dim sLogMessage As String = "SystemUserAuditHistory: "
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        oCmd.CommandTimeout = DB_TIMEOUT
        Try
            oConn.Open()
            oCmd.Connection = oConn
            nResult = oCmd.ExecuteNonQuery()
            gnRecordCount += nResult
        Catch ex As Exception
            WebMsgBox.Show("Error in PurgeSystemUserAuditHistory: " & ex.Message)
            Throw New Exception("Error in PurgeSystemUserAuditHistory: " & ex.Message)
        Finally
            oConn.Close()
            sLogMessage += "deleted " & nResult & " entries"
            If nResult > 0 Then
                sLogMessage = ">> " & sLogMessage
                Call LogRecordCount("SystemUserAuditHistory")
            End If
            Call Log(sLogMessage)
        End Try
    End Sub
    
    Protected Sub PurgeEmailQueueHistory()
        Dim sSQL As String = "DELETE FROM EmailQueueHistory WHERE SentOn < (GETDATE() - 90)"
        Dim nResult As Integer
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        oCmd.CommandTimeout = DB_TIMEOUT * 5
        Try
            Call Log("Processing EmailQueueHistory")
            Call LogRecordCount("EmailQueueHistory")
            oConn.Open()
            oCmd.Connection = oConn
            nResult = oCmd.ExecuteNonQuery()
            gnRecordCount += nResult
        Catch ex As Exception
            WebMsgBox.Show("Error in PurgeEmailHistory: " & ex.Message)
            Throw New Exception("Error in PurgeEmailHistory: " & ex.Message)
        Finally
            oConn.Close()
            Call Log("Deleted " & nResult & " entries")
            Call LogRecordCount("EmailQueueHistory")
        End Try
    End Sub
    
    Protected Sub PurgeEmailQueueHistoryAll()
        Dim sSQL As String = "DELETE FROM EmailQueueHistory"
        Dim nResult As Integer
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        oCmd.CommandTimeout = DB_TIMEOUT * 5
        Try
            Call Log("Processing EmailQueueHistory (All)")
            Call LogRecordCount("EmailQueueHistory")
            oConn.Open()
            oCmd.Connection = oConn
            nResult = oCmd.ExecuteNonQuery()
            gnRecordCount += nResult
        Catch ex As Exception
            WebMsgBox.Show("Error in PurgeEmailHistoryAll: " & ex.Message)
            Throw New Exception("Error in PurgeEmailHistoryAll: " & ex.Message)
        Finally
            oConn.Close()
            Call Log("Deleted " & nResult & " entries")
            Call LogRecordCount("EmailQueueHistory")
        End Try
    End Sub
    
    Protected Sub PurgeEmailMessageQueue()
        Dim sSQL As String = "DELETE TOP (10000) FROM EmailMessageQueue"
        Dim nResult As Integer
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        oCmd.CommandTimeout = DB_TIMEOUT * 5
        Try
            Call Log("Processing EmailMessageQueue (next 10000 entries)")
            oConn.Open()
            oCmd.Connection = oConn
            nResult = oCmd.ExecuteNonQuery()
            gnRecordCount += nResult
        Catch ex As Exception
            WebMsgBox.Show("Error in PurgeEmailMessageQueue: " & ex.Message)
            Throw New Exception("Error in PurgeEmailMessageQueue: " & ex.Message)
        Finally
            oConn.Close()
            Call Log("Deleted " & nResult & " entries")
        End Try
    End Sub
    
    Protected Sub PurgeMDSTransactionLog()
        Dim sSQL As String = "DELETE FROM MDSTransactionLog"
        Dim nResult As Integer
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        oCmd.CommandTimeout = DB_TIMEOUT
        Try
            Call Log("Processing MDSTransactionLog")
            oConn.Open()
            oCmd.Connection = oConn
            nResult = oCmd.ExecuteNonQuery()
            gnRecordCount += nResult
        Catch ex As Exception
            WebMsgBox.Show("Error in PurgeMDSTransactionLog: " & ex.Message)
            Throw New Exception("Error in PurgeMDSTransactionLog: " & ex.Message)
        Finally
            oConn.Close()
            Call Log("Deleted " & nResult & " entries")
            Call LogRecordCount("MDSTransactionLog")
        End Try
    End Sub
    
    Protected Sub PurgeLogisticWebHit()
        Dim sSQL As String = "DELETE FROM LogisticWebHit"
        Dim nResult As Integer
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        oCmd.CommandTimeout = 1000
        Try
            Call Log("Processing LogisticWebHit")
            Call LogRecordCount("LogisticWebHit")
            oConn.Open()
            oCmd.Connection = oConn
            nResult = oCmd.ExecuteNonQuery()
            gnRecordCount += nResult
        Catch ex As Exception
            WebMsgBox.Show("Error in PurgeLogisticWebHit: " & ex.Message)
            Throw New Exception("Error in PurgeLogisticWebHit: " & ex.Message)
        Finally
            oConn.Close()
            Call Log("Deleted " & nResult & " entries")
            Call LogRecordCount("LogisticWebHit")
        End Try
    End Sub
    
    Protected Sub PurgePostCodeLookup()
        Dim sSQL As String = "DELETE FROM PostCodeLookup"
        Dim nResult As Integer
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        oCmd.CommandTimeout = 1000
        Try
            Call Log("Processing PostCodeLookup")
            Call LogRecordCount("PostCodeLookup")
            oConn.Open()
            oCmd.Connection = oConn
            nResult = oCmd.ExecuteNonQuery()
            gnRecordCount += nResult
        Catch ex As Exception
            WebMsgBox.Show("Error in PurgePostCodeLookup: " & ex.Message)
            Throw New Exception("Error in PurgePostCodeLookup: " & ex.Message)
        Finally
            oConn.Close()
            Call Log("Deleted " & nResult & " entries")
            Call LogRecordCount("PostCodeLookup")
        End Try
    End Sub
    
    Protected Sub PurgeClientData_FEXCO_ReportData()
        Dim sSQL As String = "DELETE FROM ClientData_FEXCO_ReportData"
        Dim nResult As Integer
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        oCmd.CommandTimeout = 1000
        Try
            Call Log("Processing ClientData_FEXCO_ReportData")
            Call LogRecordCount("ClientData_FEXCO_ReportData")
            oConn.Open()
            oCmd.Connection = oConn
            nResult = oCmd.ExecuteNonQuery()
            gnRecordCount += nResult
        Catch ex As Exception
            WebMsgBox.Show("Error in PurgeClientData_FEXCO_ReportData: " & ex.Message)
            Throw New Exception("Error in PurgeClientData_FEXCO_ReportData: " & ex.Message)
        Finally
            oConn.Close()
            Call Log("Deleted " & nResult & " entries")
            Call LogRecordCount("ClientData_FEXCO_ReportData")
        End Try
    End Sub
    
    Protected Sub Purge_AAA_FEXCO_Debug()
        Dim sSQL As String = "DELETE FROM AAA_FEXCO_Debug"
        Dim nResult As Integer
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        oCmd.CommandTimeout = 1000
        Try
            Call Log("Processing AAA_FEXCO_Debug")
            Call LogRecordCount("AAA_FEXCO_Debug")
            oConn.Open()
            oCmd.Connection = oConn
            nResult = oCmd.ExecuteNonQuery()
            gnRecordCount += nResult
        Catch ex As Exception
            WebMsgBox.Show("Error in Purge_AAA_FEXCO_Debug: " & ex.Message)
            Throw New Exception("Error in Purge_AAA_FEXCO_Debug: " & ex.Message)
        Finally
            oConn.Close()
            Call Log("Deleted " & nResult & " entries")
            Call LogRecordCount("AAA_FEXCO_Debug")
        End Try
    End Sub
    
    Protected Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs)
        gnPageTimeout = Server.ScriptTimeout
        Server.ScriptTimeout = 3600
    End Sub

    Protected Sub Page_Unload(ByVal sender As Object, ByVal e As System.EventArgs)
        Server.ScriptTimeout = gnPageTimeout
    End Sub
    
    Protected Sub btnQuickPurge_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call QuickPurge()
    End Sub
    
    Protected Sub QuickPurge()
        Call Log("Starting quick purge at " & DateTime.Now.ToLongTimeString)
        If cbKeepLast90DaysEmails.Checked Then
            Call PurgeEmailQueueHistory()
        Else
            Call PurgeEmailQueueHistoryAll()
        End If
        Call PurgeMDSTransactionLog()
        Call PurgeLogisticWebHit()
        Call PurgePostCodeLookup()
        Call PurgeEmailMessageQueue()
        Call PurgeClientData_FEXCO_ReportData()
        Call Purge_AAA_FEXCO_Debug()
        Call UserProductProfilePurge()
        Call Log("Finished quick purge at " & DateTime.Now.ToLongTimeString)
    End Sub
    
    Protected Sub SaveTableStats()
        Dim lstTableList As New List(Of String)
        Dim dictTableStats As New Dictionary(Of String, Integer)
        lstTableList.Add("tbl")
        For Each s As String In lstTableList
            dictTableStats.Add(s, GetRecordCount(s))
        Next
        pdictTableStats = dictTableStats
    End Sub

    Protected Sub LogRecordCount(ByVal sTableName As String)
        Call Log("Table " & sTableName & " contains " & GetRecordCount(sTableName) & " record(s)")
    End Sub
    
    Protected Function GetRecordCount(ByVal sTableName As String) As Integer
        GetRecordCount = -1
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataReader As SqlDataReader = Nothing
        Dim sSQL As String = "SELECT COUNT (*) FROM " & sTableName
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            oDataReader.Read()
            GetRecordCount = oDataReader(0)
        Catch ex As Exception
            WebMsgBox.Show("Error in GetRecordCount: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Function
    
    Protected Sub lnkbtnClearLog_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        tbLog.Text = String.Empty
    End Sub
    
    Protected Sub lnkbtnLastJob_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim sSQL As String = "SELECT CustomerAccountCode, LastJobOn FROM Customer ORDER BY LastJobOn, CustomerAccountCode"
        Dim oDataTable As DataTable = ExecuteQueryToDataTable(sSQL)
        For Each dr As DataRow In oDataTable.Rows
            tbLog.Text += dr("CustomerAccountCode") & ": " & dr("LastJobOn") & Environment.NewLine
        Next
    End Sub
    
    Protected Sub btnPurgeBefore_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim sSQLPeriod As String
        Dim sSQL As String
        sSQLPeriod = "(SELECT [key] FROM Consignment WHERE CreatedOn < '1-Jan-" & ddlPurgeBefore.SelectedValue & "')"
        Call Log("PURGING Consignments before " & ddlPurgeBefore.SelectedValue)
        Call LogRecordCount("ConsignmentChange")
        sSQL = "DELETE FROM ConsignmentChange WHERE ConsignmentKey IN " & sSQLPeriod
        If Not ExecuteNonQuery(sSQL) Then
            WebMsgBox.Show("Error executing SQL statement: " & sSQL)
            Exit Sub
        End If
        Call LogRecordCount("ConsignmentChange")
        
        Call Log("")
        Call LogRecordCount("ConsignmentCost")
        sSQL = "DELETE FROM ConsignmentCost WHERE ConsignmentKey IN " & sSQLPeriod
        If Not ExecuteNonQuery(sSQL) Then
            WebMsgBox.Show("Error executing SQL statement: " & sSQL)
            Exit Sub
        End If
        Call LogRecordCount("ConsignmentCost")

        Call Log("")
        Call LogRecordCount("ConsignmentCost")
        sSQL = "DELETE FROM ConsignmentNote WHERE ConsignmentKey IN " & sSQLPeriod
        If Not ExecuteNonQuery(sSQL) Then
            WebMsgBox.Show("Error executing SQL statement: " & sSQL)
            Exit Sub
        End If
        Call LogRecordCount("ConsignmentCost")

        Call Log("")
        Call LogRecordCount("ConsignmentRoute")
        sSQL = "DELETE FROM ConsignmentRoute WHERE ConsignmentKey IN " & sSQLPeriod
        If Not ExecuteNonQuery(sSQL) Then
            WebMsgBox.Show("Error executing SQL statement: " & sSQL)
            Exit Sub
        End If
        Call LogRecordCount("ConsignmentRoute")

        Call Log("")
        Call LogRecordCount("ConsignmentSplit")
        sSQL = "DELETE FROM ConsignmentSplit WHERE MasterConsignmentKey IN " & sSQLPeriod
        If Not ExecuteNonQuery(sSQL) Then
            WebMsgBox.Show("Error executing SQL statement: " & sSQL)
            Exit Sub
        End If
        Call LogRecordCount("ConsignmentSplit")

        Call Log("")
        Call LogRecordCount("ConsignmentTrackingStage")
        sSQL = "DELETE FROM ConsignmentTrackingStage WHERE ConsignmentKey IN " & sSQLPeriod
        If Not ExecuteNonQuery(sSQL) Then
            WebMsgBox.Show("Error executing SQL statement: " & sSQL)
            Exit Sub
        End If
        Call LogRecordCount("ConsignmentTrackingStage")
        
        Call Log("")
        Call LogRecordCount("ConsignmentNote")
        sSQL = "DELETE FROM ConsignmentNote WHERE ConsignmentKey IN " & sSQLPeriod
        If Not ExecuteNonQuery(sSQL) Then
            WebMsgBox.Show("Error executing SQL statement: " & sSQL)
            Exit Sub
        End If
        Call LogRecordCount("ConsignmentNote")
        
        Call Log("")
        Call LogRecordCount("LogisticMovement")
        sSQL = "DELETE FROM LogisticMovement WHERE ConsignmentKey IN " & sSQLPeriod
        If Not ExecuteNonQuery(sSQL) Then
            WebMsgBox.Show("Error executing SQL statement: " & sSQL)
            Exit Sub
        End If
        Call LogRecordCount("LogisticMovement")
        
        Call Log("")
        Call LogRecordCount("Consignment")
        sSQL = "DELETE FROM Consignment WHERE [key] IN " & sSQLPeriod
        If Not ExecuteNonQuery(sSQL) Then
            WebMsgBox.Show("Error executing SQL statement: " & sSQL)
            Exit Sub
        End If
        Call LogRecordCount("Consignment")
        
        sSQLPeriod = "(SELECT LogisticBookingKey FROM LogisticBooking WHERE BookedOn < '1-Jan-" & ddlPurgeBefore.SelectedValue & "')"

        Call Log("")
        Call LogRecordCount("LogisticBookingTracking")
        sSQL = "DELETE FROM LogisticBookingTracking WHERE StockBookingKey IN " & sSQLPeriod
        If Not ExecuteNonQuery(sSQL) Then
            WebMsgBox.Show("Error executing SQL statement: " & sSQL)
            Exit Sub
        End If
        Call LogRecordCount("LogisticBookingTracking")
        
        Call Log("")
        Call LogRecordCount("LogisticMovement")
        sSQL = "DELETE FROM LogisticMovement WHERE LogisticBookingKey IN " & sSQLPeriod
        If Not ExecuteNonQuery(sSQL) Then
            WebMsgBox.Show("Error executing SQL statement: " & sSQL)
            Exit Sub
        End If
        Call LogRecordCount("LogisticMovement")
        
        Call Log("")
        Call LogRecordCount("LogisticBooking")
        sSQL = "DELETE FROM LogisticBooking WHERE LogisticBookingKey IN " & sSQLPeriod
        If Not ExecuteNonQuery(sSQL) Then
            WebMsgBox.Show("Error executing SQL statement: " & sSQL)
            Exit Sub
        End If
        Call LogRecordCount("LogisticBooking")
        Call Log("DONE !!!")
    End Sub
    
    Protected Sub btnCourierPurgeBefore_Click(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub
    
    Protected Sub btnMANHistoryPurgeBefore_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsNumeric(tbMANHistoryPurge.Text) Then
            WebMsgBox.Show("Please specify a period in days")
        Else
            Call MANHistoryPurge()
        End If
    End Sub
    
    Protected Sub MANHistoryPurge()
        Dim sSQL As String
        
        Call Log("")
        Call LogRecordCount("ManProdHistWithProspectusNos")
        sSQL = "SET ROWCOUNT 10000 DELETE FROM ManProdHistWithProspectusNos WHERE SHIP_DATE < (GETDATE() - " & tbMANHistoryPurge.Text & ") SET ROWCOUNT 0"
        Do While ExecuteNonQueryReturnRowsAffected(sSQL) > 0
        Loop
        Call LogRecordCount("ManProdHistWithProspectusNos")
        
        Call LogRecordCount("ManProdHistWithOutProspectusNos")
        sSQL = "SET ROWCOUNT 10000 DELETE FROM ManProdHistWithOutProspectusNos WHERE SHIP_DATE < (GETDATE() - " & tbMANHistoryPurge.Text & ") SET ROWCOUNT 0"
        Do While ExecuteNonQueryReturnRowsAffected(sSQL) > 0
        Loop
        Call LogRecordCount("ManProdHistWithOutProspectusNos")
    End Sub
    
    Protected Sub UserProductProfilePurge()
        Dim sSQL As String
        
        Call Log("")
        Call LogRecordCount("UserProductProfile")
        sSQL = "SET ROWCOUNT 10000 DELETE FROM UserProductProfile WHERE ProductKey IN (SELECT LogisticProductKey FROM LogisticProduct WHERE DeletedFlag = 'Y') SET ROWCOUNT 0"
        Do While ExecuteNonQueryReturnRowsAffected(sSQL) > 0
        Loop
        
        sSQL = "SET ROWCOUNT 10000 DELETE FROM UserProductProfile WHERE UserKey IN (SELECT [key] from UserProfile WHERE Status = 'Suspended' AND ISNULL(LastLogon,'1-Jan-2000') < '1-Jan-2009') SET ROWCOUNT 0"
        Do While ExecuteNonQueryReturnRowsAffected(sSQL) > 0
        Loop
        Call LogRecordCount("UserProductProfile")
    End Sub
    
    Protected Sub btnTrackingBefore_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsNumeric(tbTrackingPurge.Text) Then
            WebMsgBox.Show("Please specify a period in days")
        Else
            Call TrackingPurge()
        End If
    End Sub
    
    Protected Sub TrackingPurge()
        Dim sSQL As String
        Call Log("")

        Call LogRecordCount("LogisticProductTracking")
        sSQL = "SET ROWCOUNT 10000 DELETE FROM LogisticProductTracking WHERE EventTime < (GETDATE() - " & tbTrackingPurge.Text & ") SET ROWCOUNT 0"
        Do While ExecuteNonQueryReturnRowsAffected(sSQL) > 0
        Loop
        Call LogRecordCount("LogisticProductTracking")
        
        Call LogRecordCount("LogisticBookingTracking")
        sSQL = "SET ROWCOUNT 10000 DELETE FROM LogisticBookingTracking WHERE EventTime < (GETDATE() - " & tbTrackingPurge.Text & ") SET ROWCOUNT 0"
        Do While ExecuteNonQueryReturnRowsAffected(sSQL) > 0
        Loop
        Call LogRecordCount("LogisticBookingTracking")

        Call LogRecordCount("CourierBookingComment")
        sSQL = "SET ROWCOUNT 10000 DELETE FROM CourierBookingComment WHERE AddedOn < (GETDATE() - " & tbTrackingPurge.Text & ") SET ROWCOUNT 0"
        Do While ExecuteNonQueryReturnRowsAffected(sSQL) > 0
        Loop
        Call LogRecordCount("CourierBookingComment")
        
        Call LogRecordCount("CourierBookingTrackingStage")
        sSQL = "SET ROWCOUNT 10000 DELETE FROM CourierBookingTrackingStage WHERE TrackedOn < (GETDATE() - " & tbTrackingPurge.Text & ") SET ROWCOUNT 0"
        Do While ExecuteNonQueryReturnRowsAffected(sSQL) > 0
        Loop
        Call LogRecordCount("CourierBookingTrackingStage")
        
        Call LogRecordCount("ConsignmentTrackingStage")
        sSQL = "SET ROWCOUNT 10000 DELETE FROM ConsignmentTrackingStage WHERE TrackedOn < (GETDATE() - " & tbTrackingPurge.Text & ") SET ROWCOUNT 0"
        Do While ExecuteNonQueryReturnRowsAffected(sSQL) > 0
        Loop
        Call LogRecordCount("ConsignmentTrackingStage")
        
        Call LogRecordCount("ConsignmentNote")
        sSQL = "SET ROWCOUNT 10000 DELETE FROM ConsignmentNote WHERE AddedOn < (GETDATE() - " & tbTrackingPurge.Text & ") SET ROWCOUNT 0"
        Do While ExecuteNonQueryReturnRowsAffected(sSQL) > 0
        Loop
        Call LogRecordCount("ConsignmentNote")
        
    End Sub

    Protected Sub lnkbtnGABByCustomer_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim sSQL As String = "SELECT gab.CustomerKey, CustomerAccountCode, COUNT(gab.CustomerKey) 'Entries' FROM GlobalAddressBook gab INNER JOIN Customer c ON gab.CustomerKey = c.CustomerKey GROUP BY gab.CustomerKey, CustomerAccountCode ORDER BY Entries DESC"
        Dim oDataTable As DataTable = ExecuteQueryToDataTable(sSQL)
        For Each dr As DataRow In oDataTable.Rows
            tbLog.Text += dr("CustomerAccountCode") & ": " & dr("Entries") & Environment.NewLine
        Next
    End Sub
    
    Protected Sub lnkbtnTotConsignments_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim sSQL As String = "SELECT con.CustomerKey, CustomerAccountCode, COUNT(con.CustomerKey) 'Entries' FROM Consignment con INNER JOIN Customer c ON con.CustomerKey = c.CustomerKey GROUP BY con.CustomerKey, CustomerAccountCode ORDER BY Entries DESC"
        Dim oDataTable As DataTable = ExecuteQueryToDataTable(sSQL)
        For Each dr As DataRow In oDataTable.Rows
            tbLog.Text += dr("CustomerAccountCode") & ": " & dr("Entries") & Environment.NewLine
        Next
    End Sub
    
    Protected Sub lnkbtnTotalRevenuePerMonth_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim sbSQL As New StringBuilder
        sbSQL.Append("DECLARE @Temp TABLE ( [id] int IDENTITY, Period varchar(50), Amount money)")

        sbSQL.Append("INSERT INTO @Temp (Period, Amount)  (SELECT 'Jan 07', SUM(ISNULL(CashOnDelAmount,0)) FROM Consignment WHERE CreatedOn > '31-dec-2006' AND CreatedOn < '1-feb-2007') ")
        sbSQL.Append("INSERT INTO @Temp (Period, Amount)  (SELECT 'Feb 07', SUM(ISNULL(CashOnDelAmount,0)) FROM Consignment WHERE CreatedOn > '31-jan-2007' AND CreatedOn < '1-mar-2007') ")
        sbSQL.Append("INSERT INTO @Temp (Period, Amount)  (SELECT 'Mar 07', SUM(ISNULL(CashOnDelAmount,0)) FROM Consignment WHERE CreatedOn > '28-feb-2007' AND CreatedOn < '1-apr-2007') ")
        sbSQL.Append("INSERT INTO @Temp (Period, Amount)  (SELECT 'Apr 07', SUM(ISNULL(CashOnDelAmount,0)) FROM Consignment WHERE CreatedOn > '31-mar-2007' AND CreatedOn < '1-may-2007') ")
        sbSQL.Append("INSERT INTO @Temp (Period, Amount)  (SELECT 'May 07', SUM(ISNULL(CashOnDelAmount,0)) FROM Consignment WHERE CreatedOn > '30-apr-2007' AND CreatedOn < '1-jun-2007') ")
        sbSQL.Append("INSERT INTO @Temp (Period, Amount)  (SELECT 'Jun 07', SUM(ISNULL(CashOnDelAmount,0)) FROM Consignment WHERE CreatedOn > '31-may-2007' AND CreatedOn < '1-jul-2007') ")
        sbSQL.Append("INSERT INTO @Temp (Period, Amount)  (SELECT 'Jul 07', SUM(ISNULL(CashOnDelAmount,0)) FROM Consignment WHERE CreatedOn > '30-jun-2007' AND CreatedOn < '1-aug-2007') ")
        sbSQL.Append("INSERT INTO @Temp (Period, Amount)  (SELECT 'Aug 07', SUM(ISNULL(CashOnDelAmount,0)) FROM Consignment WHERE CreatedOn > '31-jul-2007' AND CreatedOn < '1-sep-2007') ")
        sbSQL.Append("INSERT INTO @Temp (Period, Amount)  (SELECT 'Sep 07', SUM(ISNULL(CashOnDelAmount,0)) FROM Consignment WHERE CreatedOn > '31-aug-2007' AND CreatedOn < '1-oct-2007') ")
        sbSQL.Append("INSERT INTO @Temp (Period, Amount)  (SELECT 'Oct 07', SUM(ISNULL(CashOnDelAmount,0)) FROM Consignment WHERE CreatedOn > '30-sep-2007' AND CreatedOn < '1-nov-2007') ")
        sbSQL.Append("INSERT INTO @Temp (Period, Amount)  (SELECT 'Nov 07', SUM(ISNULL(CashOnDelAmount,0)) FROM Consignment WHERE CreatedOn > '31-oct-2007' AND CreatedOn < '1-dec-2007') ")
        sbSQL.Append("INSERT INTO @Temp (Period, Amount)  (SELECT 'Dec 07', SUM(ISNULL(CashOnDelAmount,0)) FROM Consignment WHERE CreatedOn > '30-nov-2007' AND CreatedOn < '1-jan-2008') ")

        sbSQL.Append("INSERT INTO @Temp (Period, Amount)  (SELECT 'Jan 08', SUM(ISNULL(CashOnDelAmount,0)) FROM Consignment WHERE CreatedOn > '31-dec-2007' AND CreatedOn < '1-feb-2008') ")
        sbSQL.Append("INSERT INTO @Temp (Period, Amount)  (SELECT 'Feb 08', SUM(ISNULL(CashOnDelAmount,0)) FROM Consignment WHERE CreatedOn > '31-jan-2008' AND CreatedOn < '1-mar-2008') ")
        sbSQL.Append("INSERT INTO @Temp (Period, Amount)  (SELECT 'Mar 08', SUM(ISNULL(CashOnDelAmount,0)) FROM Consignment WHERE CreatedOn > '28-feb-2008' AND CreatedOn < '1-apr-2008') ")
        sbSQL.Append("INSERT INTO @Temp (Period, Amount)  (SELECT 'Apr 08', SUM(ISNULL(CashOnDelAmount,0)) FROM Consignment WHERE CreatedOn > '31-mar-2008' AND CreatedOn < '1-may-2008') ")
        sbSQL.Append("INSERT INTO @Temp (Period, Amount)  (SELECT 'May 08', SUM(ISNULL(CashOnDelAmount,0)) FROM Consignment WHERE CreatedOn > '30-apr-2008' AND CreatedOn < '1-jun-2008') ")
        sbSQL.Append("INSERT INTO @Temp (Period, Amount)  (SELECT 'Jun 08', SUM(ISNULL(CashOnDelAmount,0)) FROM Consignment WHERE CreatedOn > '31-may-2008' AND CreatedOn < '1-jul-2008') ")
        sbSQL.Append("INSERT INTO @Temp (Period, Amount)  (SELECT 'Jul 08', SUM(ISNULL(CashOnDelAmount,0)) FROM Consignment WHERE CreatedOn > '30-jun-2008' AND CreatedOn < '1-aug-2008') ")
        sbSQL.Append("INSERT INTO @Temp (Period, Amount)  (SELECT 'Aug 08', SUM(ISNULL(CashOnDelAmount,0)) FROM Consignment WHERE CreatedOn > '31-jul-2008' AND CreatedOn < '1-sep-2008') ")
        sbSQL.Append("INSERT INTO @Temp (Period, Amount)  (SELECT 'Sep 08', SUM(ISNULL(CashOnDelAmount,0)) FROM Consignment WHERE CreatedOn > '31-aug-2008' AND CreatedOn < '1-oct-2008') ")
        sbSQL.Append("INSERT INTO @Temp (Period, Amount)  (SELECT 'Oct 08', SUM(ISNULL(CashOnDelAmount,0)) FROM Consignment WHERE CreatedOn > '30-sep-2008' AND CreatedOn < '1-nov-2008') ")
        sbSQL.Append("INSERT INTO @Temp (Period, Amount)  (SELECT 'Nov 08', SUM(ISNULL(CashOnDelAmount,0)) FROM Consignment WHERE CreatedOn > '31-oct-2008' AND CreatedOn < '1-dec-2008') ")
        sbSQL.Append("INSERT INTO @Temp (Period, Amount)  (SELECT 'Dec 08', SUM(ISNULL(CashOnDelAmount,0)) FROM Consignment WHERE CreatedOn > '30-nov-2008' AND CreatedOn < '1-jan-2009') ")

        sbSQL.Append("INSERT INTO @Temp (Period, Amount)  (SELECT 'Jan 09', SUM(ISNULL(CashOnDelAmount,0)) FROM Consignment WHERE CreatedOn > '31-dec-2008' AND CreatedOn < '1-feb-2009') ")
        sbSQL.Append("INSERT INTO @Temp (Period, Amount)  (SELECT 'Feb 09', SUM(ISNULL(CashOnDelAmount,0)) FROM Consignment WHERE CreatedOn > '31-jan-2009' AND CreatedOn < '1-mar-2009') ")
        sbSQL.Append("INSERT INTO @Temp (Period, Amount)  (SELECT 'Mar 09', SUM(ISNULL(CashOnDelAmount,0)) FROM Consignment WHERE CreatedOn > '28-feb-2009' AND CreatedOn < '1-apr-2009') ")
        sbSQL.Append("INSERT INTO @Temp (Period, Amount)  (SELECT 'Apr 09', SUM(ISNULL(CashOnDelAmount,0)) FROM Consignment WHERE CreatedOn > '31-mar-2009' AND CreatedOn < '1-may-2009') ")
        sbSQL.Append("INSERT INTO @Temp (Period, Amount)  (SELECT 'May 09', SUM(ISNULL(CashOnDelAmount,0)) FROM Consignment WHERE CreatedOn > '30-apr-2009' AND CreatedOn < '1-jun-2009') ")
        sbSQL.Append("INSERT INTO @Temp (Period, Amount)  (SELECT 'Jun 09', SUM(ISNULL(CashOnDelAmount,0)) FROM Consignment WHERE CreatedOn > '31-may-2009' AND CreatedOn < '1-jul-2009') ")
        sbSQL.Append("INSERT INTO @Temp (Period, Amount)  (SELECT 'Jul 09', SUM(ISNULL(CashOnDelAmount,0)) FROM Consignment WHERE CreatedOn > '30-jun-2009' AND CreatedOn < '1-aug-2009') ")
        sbSQL.Append("INSERT INTO @Temp (Period, Amount)  (SELECT 'Aug 09', SUM(ISNULL(CashOnDelAmount,0)) FROM Consignment WHERE CreatedOn > '31-jul-2009' AND CreatedOn < '1-sep-2009') ")
        sbSQL.Append("INSERT INTO @Temp (Period, Amount)  (SELECT 'Sep 09', SUM(ISNULL(CashOnDelAmount,0)) FROM Consignment WHERE CreatedOn > '31-aug-2009' AND CreatedOn < '1-oct-2009') ")
        sbSQL.Append("INSERT INTO @Temp (Period, Amount)  (SELECT 'Oct 09', SUM(ISNULL(CashOnDelAmount,0)) FROM Consignment WHERE CreatedOn > '30-sep-2009' AND CreatedOn < '1-nov-2009') ")
        sbSQL.Append("INSERT INTO @Temp (Period, Amount)  (SELECT 'Nov 09', SUM(ISNULL(CashOnDelAmount,0)) FROM Consignment WHERE CreatedOn > '31-oct-2009' AND CreatedOn < '1-dec-2009') ")
        sbSQL.Append("INSERT INTO @Temp (Period, Amount)  (SELECT 'Dec 09', SUM(ISNULL(CashOnDelAmount,0)) FROM Consignment WHERE CreatedOn > '30-nov-2009' AND CreatedOn < '1-jan-2010') ")

        sbSQL.Append("SELECT Period, Amount FROM @Temp ORDER BY [id] ")
        Dim oDataTable As DataTable = ExecuteQueryToDataTable(sbSQL.ToString)
        For Each dr As DataRow In oDataTable.Rows
            tbLog.Text += dr("Period") & ControlChars.Tab & dr("Amount") & Environment.NewLine
        Next
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
            WebMsgBox.Show("Error in ExecuteNonQuery executing: " & sQuery & " : " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Function
   
    Protected Function ExecuteNonQueryReturnRowsAffected(ByVal sQuery As String) As Integer
        ExecuteNonQueryReturnRowsAffected = 0
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand
        Try
            oConn.Open()
            oCmd = New SqlCommand(sQuery, oConn)
            ExecuteNonQueryReturnRowsAffected = oCmd.ExecuteNonQuery()
        Catch ex As Exception
            ExecuteNonQueryReturnRowsAffected = -1
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

    Protected Sub GetMatchingCustomers()
        Dim sSQL As String = "SELECT CustomerName FROM Customer WHERE CustomerName = ~~ OR CustomerAddress1 = ~~"
        Dim oDT As DataTable = ExecuteQueryToDataTable2(sSQL, tbLog.Text)
    End Sub
    
    Protected Function ExecuteQueryToDataTable2(ByVal sQuery As String, Optional ByVal sSearchString As String = "") As DataTable
        Dim oDataTable As New DataTable
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter(sQuery, oConn)
        Dim oCmd As SqlCommand = New SqlCommand(sQuery, oConn)
        If sQuery.Contains("~~") Then
            If sSearchString = String.Empty Then
                sSearchString = "_"
            End If
            sSearchString = sSearchString.Replace("'", "''")
            sQuery = sQuery.Replace("~~", "'" & sSearchString & "'")
        End If
        Try
            oAdapter.Fill(oDataTable)
            oConn.Open()
        Catch ex As Exception
            oDataTable = Nothing
        Finally
            oConn.Close()
        End Try
        ExecuteQueryToDataTable2 = oDataTable
    End Function

    Property pdictTableStats() As Dictionary(Of String, Integer)
        Get
            Dim o As Object = ViewState("PD_TableStats")
            If o Is Nothing Then
                Return Nothing
            End If
            Return CType(o, Dictionary(Of String, Integer))
        End Get
        Set(ByVal Value As Dictionary(Of String, Integer))
            ViewState("PD_TableStats") = Value
        End Set
    End Property
    
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Purge Database</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        &nbsp;
        <asp:Label ID="Label1" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small"
            Text="Purge Database"></asp:Label><br />
        <br />
        &nbsp;<asp:Button ID="btnQuickPurge" runat="server" Height="26px" OnClick="btnQuickPurge_Click"
            Text="quick purge" Width="150px" />
        &nbsp;<asp:Label ID="Label3" runat="server" Font-Bold="False" Font-Names="Verdana"
            Font-Size="XX-Small" Text="(email queue, MDS log, event list, etc.)"></asp:Label>
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        <asp:CheckBox ID="cbKeepLast90DaysEmails" runat="server" Checked="True" Font-Names="Verdana"
            Font-Size="XX-Small" Text="keep last 90 days emails" />
        <br />
        <hr />
        &nbsp;<asp:Button ID="btnPurgeDatabase" runat="server" OnClick="btnPurgeDatabase_Click"
            Text="full purge" Width="150px" />
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        <asp:Label ID="Label2" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small"
            Text="Maximum run time:"></asp:Label><asp:DropDownList ID="ddlDuration" runat="server"
                Font-Names="Verdana" Font-Size="XX-Small">
                <asp:ListItem Value="30">30 seconds</asp:ListItem>
                <asp:ListItem Value="60">1 minute</asp:ListItem>
                <asp:ListItem Value="300">5 minutes</asp:ListItem>
                <asp:ListItem Value="600">10 minutes</asp:ListItem>
                <asp:ListItem Value="1800">30 minutes</asp:ListItem>
                <asp:ListItem Value="3600">1 hour</asp:ListItem>
                <asp:ListItem Value="7200">2 hours</asp:ListItem>
                <asp:ListItem Value="99999999">unlimited</asp:ListItem>
            </asp:DropDownList>
        <hr />
        <asp:Button ID="btnPurgeBefore" runat="server" Text="purge stock accounts before" 
            Width="250px" onclick="btnPurgeBefore_Click" />
        &nbsp;
        <asp:Label ID="Label4" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small" Text="31st Dec"/>
        &nbsp;
        <asp:DropDownList ID="ddlPurgeBefore" runat="server">
            <asp:ListItem>2001</asp:ListItem>
            <asp:ListItem>2002</asp:ListItem>
            <asp:ListItem>2003</asp:ListItem>
            <asp:ListItem>2004</asp:ListItem>
            <asp:ListItem>2005</asp:ListItem>
            <asp:ListItem>2006</asp:ListItem>
            <asp:ListItem>2007</asp:ListItem>
            <asp:ListItem>2008</asp:ListItem>
            <asp:ListItem>2009</asp:ListItem>
        </asp:DropDownList>
        <br />
        <hr />
        <asp:Button ID="btnCourierPurgeBefore" runat="server" Text="purge courier accounts before" 
            Width="250px" onclick="btnCourierPurgeBefore_Click" />
        &nbsp;
        <asp:Label ID="Label5" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small" Text="31st Dec"/>
        &nbsp;
        <asp:DropDownList ID="ddlPurgeCourierBefore" runat="server">
            <asp:ListItem>2001</asp:ListItem>
            <asp:ListItem>2002</asp:ListItem>
            <asp:ListItem>2003</asp:ListItem>
            <asp:ListItem>2004</asp:ListItem>
            <asp:ListItem>2005</asp:ListItem>
            <asp:ListItem>2006</asp:ListItem>
            <asp:ListItem>2007</asp:ListItem>
            <asp:ListItem>2008</asp:ListItem>
            <asp:ListItem>2009</asp:ListItem>
        </asp:DropDownList>
        <br />
        <hr />
        <asp:Button ID="btnTrackingBefore" runat="server" Text="purge tracking before" 
            Width="250px" onclick="btnTrackingBefore_Click" />
        &nbsp;
        <asp:TextBox ID="tbTrackingPurge" runat="server" Font-Names="Verdana" 
            Font-Size="XX-Small" Width="50px">90</asp:TextBox>
&nbsp;<asp:Label ID="Label7" runat="server" Font-Bold="False" Font-Names="Verdana" 
            Font-Size="XX-Small" Text="days"/>
        <br />
        <hr />
        <asp:Button ID="btnMANHistoryPurgeBefore" runat="server" Text="purge MAN history before" 
            Width="250px" onclick="btnMANHistoryPurgeBefore_Click" />
        &nbsp;
        <asp:TextBox ID="tbMANHistoryPurge" runat="server" Font-Names="Verdana" 
            Font-Size="XX-Small" Width="50px">90</asp:TextBox>
&nbsp;<asp:Label ID="Label6" runat="server" Font-Bold="False" Font-Names="Verdana" 
            Font-Size="XX-Small" Text="days"/>
        <br />
        &nbsp;<asp:LinkButton ID="LinkButton1" runat="server" Font-Names="Verdana" Font-Size="XX-Small">LinkButton</asp:LinkButton>
        &nbsp;<asp:LinkButton ID="lnkbtnLastJob" runat="server" Font-Names="Verdana" Font-Size="XX-Small" onclick="lnkbtnLastJob_Click">last job</asp:LinkButton>
        &nbsp;<asp:LinkButton ID="lnkbtnGABByCustomer" runat="server" 
            Font-Names="Verdana" Font-Size="XX-Small" onclick="lnkbtnGABByCustomer_Click">GAB by customer</asp:LinkButton>
        &nbsp;<asp:LinkButton ID="lnkbtnTotConsignments" runat="server" 
            Font-Names="Verdana" Font-Size="XX-Small" onclick="lnkbtnTotConsignments_Click">Tot # Consignments</asp:LinkButton>
        &nbsp;<asp:LinkButton ID="lnkbtnTotalRevenuePerMonth" runat="server" 
            Font-Names="Verdana" Font-Size="XX-Small" 
            onclick="lnkbtnTotalRevenuePerMonth_Click">Tot Revenue/Month</asp:LinkButton>
        <hr />
        &nbsp;<asp:TextBox ID="tbLog" runat="server" Rows="20" TextMode="MultiLine" Width="100%" />
        <br />
        <asp:LinkButton ID="lnkbtnClearLog" runat="server" Font-Names="Verdana" Font-Size="XX-Small" onclick="lnkbtnClearLog_Click">clear log</asp:LinkButton>
    </div>
    </form>
    <p style="font-family: Verdana; font-size: xx-small">
        USE Logistics<br />
        IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id(N&#39;[dbo].[TableSpaceUsed]&#39;)
        AND OBJECTPROPERTY(id, N&#39;IsProcedure&#39;) = 1) DROP PROCEDURE [dbo].[TableSpaceUsed]<br />
        GO<br />
        CREATE PROCEDURE dbo.TableSpaceUsed AS<br />
        -- Create the temporary table...<br />
        CREATE TABLE #tblResults (
        <br />
        [name] nvarchar(50),
        <br />
        [rows] int,<br />
        [reserved] varchar(18),<br />
        [reserved_int] int default(0),<br />
        [data] varchar(18),<br />
        [data_int] int default(0),<br />
        [index_size] varchar(18),<br />
        [index_size_int] int default(0),<br />
        [unused] varchar(18),<br />
        [unused_int] int default(0) )<br />
        -- Populate the temp table...<br />
        EXEC sp_MSforeachtable @command1= &quot;INSERT INTO #tblResults ([name],[rows],[reserved],[data],[index_size],[unused])
        EXEC sp_spaceused &#39;?&#39;&quot;<br />
        -- Strip out the &quot; KB&quot; portion from the fields<br />
        UPDATE #tblResults SET<br />
        [reserved_int] = CAST(SUBSTRING([reserved], 1, CHARINDEX(&#39; &#39;, [reserved]))
        AS int),<br />
        [data_int] = CAST(SUBSTRING([data], 1, CHARINDEX(&#39; &#39;,[data])) AS int),<br />
        [index_size_int] = CAST(SUBSTRING([index_size], 1, CHARINDEX(&#39; &#39;, [index_size]))
        AS int),<br />
        [unused_int] = CAST(SUBSTRING([unused], 1, CHARINDEX(&#39; &#39;, [unused])) AS
        int)<br />
        -- Return the results...<br />
        SELECT * FROM #tblResults ORDER BY reserved_int<br />
        GO<br />
        GRANT EXECUTE ON [dbo].[TableSpaceUsed] TO [LogisticsUserRole]<br />
        GO<br />
        GRANT EXECUTE ON [dbo].[TableSpaceUsed] TO [LogisticsAdminRole]<br />
        GO<br />
        EXEC TableSpaceUsed
    </p>
    <p style="font-family: Verdana; font-size: xx-small">
        &nbsp;</p>
</body>
</html>
