<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.IO" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

    Const CUSTOMER_BOULEVARD As Int32 = 785
    Const COL_PRODUCTCODE = 0
    'Const COL_PRODUCTDESCRIPTION = 1
    Const COL_QTY = 1
    Const COL_PALLET = 2
    
    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString

    Protected Sub Page_Load(ByVal Source As Object, ByVal E As EventArgs)
        Server.ScriptTimeout = 3600
        If Not IsPostBack Then
            If Not IsNumeric(tbWarehouseSectionKey.Text) Then
                Server.Transfer("session_expired.aspx")
            End If
        End If
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
        Finally
            oConn.Close()
        End Try
        ExecuteQueryToDataTable = oDataTable
    End Function

    Protected Function GetProductKeyFromProductCode(ByVal sProductCode As String) As Int32
        Dim sSQL As String
        sSQL = "SELECT LogisticProductKey FROM LogisticProduct WHERE CustomerKey = " & CUSTOMER_BOULEVARD & " AND ProductCode = '" & sProductCode & "'"
        Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
        If dt.Rows.Count = 1 Then
            GetProductKeyFromProductCode = dt.Rows(0).Item(0)
        Else
            GetProductKeyFromProductCode = 0
        End If
    End Function
  
    Protected Sub ReadProductFile()
        Dim sFilename As String = MapPath(tbFilename.Text)
        If Not File.Exists(sFilename) Then
            WebMsgBox.Show("Could not find file " & sFilename)
            Exit Sub
        End If
       
        'Dim sr As New StreamReader(MapPath("QLG_RELOCATION_PICK.csv"))
        Call WriteToLog("Start processing file " & sFilename)
        tbLog.Text += "Start processing file " & sFilename & Environment.NewLine
        Dim sr As New StreamReader(sFilename)
        Dim nLineCount As Integer = 1
        Do While sr.Peek >= 0
            Dim sLine = sr.ReadLine()
            Dim sElements = Split(sLine, ",")   ' split on delimiter
            If sElements.Count = 3 Then
                Dim sProductCode As String = sElements(COL_PRODUCTCODE)
                Call WriteToLog(nLineCount.ToString & ": Processing " & sProductCode)
                Dim nProductKey As Int32 = GetProductKeyFromProductCode(sProductCode)
                If nProductKey > 0 Then
                    Dim sQty As String = sElements(COL_QTY)
                    Call WriteToLog(nLineCount.ToString & ": Qty: " & sQty)
                    If IsNumeric(sQty) AndAlso CInt(sQty) > 0 Then
                        Dim sLocation As String = sElements(COL_PALLET).Trim
                        If sLocation <> String.Empty Then
                            Call WriteToLog(nLineCount.ToString & ": Location: " & sLocation)
                            Dim nBayKey As Int32 = GetBayKey(sLocation)
                            If nBayKey > 0 Then
                                Call AddQuantity(nProductKey, nBayKey, CInt(sQty))
                            Else
                                Call WriteToLog(nLineCount.ToString & ": ERR - could not match bay name " & sLocation & " to key")
                            End If
                        Else
                            Call WriteToLog(nLineCount.ToString & ": ERR - blank location")
                        End If
                    Else
                        Call WriteToLog(nLineCount.ToString & ": ERR - 0 or non-numeric quantity")
                    End If
                Else
                    Call WriteToLog(nLineCount.ToString & ": ERR - Could not match product " & sProductCode)
                End If
            Else
                Call WriteToLog(nLineCount.ToString & ": ERR - line did not contain 3 elements - may be trailing newline")
            End If
            nLineCount += 1
        Loop
        sr.Close()
        Call WriteToLog("Finished Processing file " & sFilename)
        tbLog.Text += "Finished Processing file " & sFilename & Environment.NewLine
    End Sub

    Protected Function GetBayKey(ByVal sBayName As String) As Int32
        Dim sSQL As String
        GetBayKey = 0
        sSQL = "SELECT WarehouseBayKey FROM WarehouseBay WHERE WarehouseSectionKey = " & tbWarehouseSectionKey.Text & " AND WarehouseBayId = '" & sBayName & "'"
        Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
        If dt.Rows.Count = 1 Then
            Dim sBayKey As String = dt.Rows(0).Item(0)
            If IsNumeric(sBayKey) Then
                GetBayKey = CInt(dt.Rows(0).Item(0))
            Else
                Call WriteToLog("ERR: non-numeric bay key retrieved from database")
            End If
        End If
    End Function
   
    Protected Sub WriteToLog(ByVal sMessage As String)
        Dim sSQL As String
        sSQL = "INSERT INTO AAA_Debug (result) VALUES ('" & sMessage.Replace("'", "''") & "')"
        Call ExecuteQueryToDataTable(sSQL)
    End Sub
   
    Protected Sub AddQuantity(nLogisticProductKey As Int32, nWarehouseBayKey As Int32, nLogisticProductQuantity As Int32)
        Dim sSQL As String
        sSQL = "SELECT LogisticProductQuantity FROM LogisticProductLocation WHERE LogisticProductKey = " & nLogisticProductKey & " AND WarehouseBayKey = " & nWarehouseBayKey
        Dim oDT As DataTable = ExecuteQueryToDataTable(sSQL)
        If oDT.Rows.Count > 0 Then
            If oDT.Rows.Count = 1 Then
                Dim nQuantity As Int32 = CInt(oDT.Rows(0).Item(0))
                nQuantity = nQuantity + nLogisticProductQuantity
                If nQuantity >= 0 Then
                    sSQL = "UPDATE LogisticProductLocation SET LogisticProductQuantity = " & nQuantity.ToString & " WHERE WarehouseBayKey = " & nWarehouseBayKey.ToString & " AND LogisticProductKey = " & nLogisticProductKey.ToString
                Else
                    Call WriteToLog("ERR: Quantity adjustment entered would make total quantity in this location negative")
                End If
            Else
                Call WriteToLog("ERR: multiple instances of one product in a single location")
            End If
        Else
            sSQL = "INSERT INTO LogisticProductLocation (LogisticProductKey, WarehouseBayKey, LogisticProductQuantity, DateStored) VALUES ("
            sSQL &= nLogisticProductKey
            sSQL &= ", "
            sSQL &= nWarehouseBayKey
            sSQL &= ", "
            sSQL &= nLogisticProductQuantity.ToString
            sSQL &= ", GETDATE())"
        End If
       
        Call ExecuteQueryToDataTable(sSQL)

        Dim sWarehouseBayName As String = GetBayNameFromBayKey(nWarehouseBayKey)
        Dim nWarehouseSectionKey As Int32 = GetSectionKeyFromBayKey(nWarehouseBayKey)
        Dim sWarehouseSectionName As String = GetSectionNameFromSectionKey(nWarehouseSectionKey)
        Dim nWarehouseRackKey As Int32 = GetRackKeyFromSectionKey(nWarehouseSectionKey)
        Dim sWarehouseRackName As String = GetRackNameFromRackKey(nWarehouseRackKey)
        Dim nWarehouseKey As Int32 = GetWarehouseKeyFromRackKey(nWarehouseRackKey)
        Dim sWarehousename As String = GetWarehouseNameFromWarehouseKey(nWarehouseKey)
        Dim sMsg As String = "Added " & nLogisticProductQuantity.ToString & " of " & GetProductCodeFromProductKey(nLogisticProductKey) & " (" & nLogisticProductKey.ToString & ") to Warehouse " & sWarehousename & ", Rack " & sWarehouseRackName & ", Section " & sWarehouseSectionName & ", Bay " & sWarehouseBayName & Environment.NewLine
        tbLog.Text += sMsg
        Call WriteToLog(sMsg)
    End Sub
     
    Protected Function GetProductCodeFromProductKey(nLogisticProductKey As Int32) As String
        Dim sSQL As String
        sSQL = "SELECT ProductCode FROM LogisticProduct WHERE LogisticProductKey = " & nLogisticProductKey
        GetProductCodeFromProductKey = ExecuteQueryToDataTable(sSQL).Rows(0).Item(0)
    End Function

    Protected Function GetBayNameFromBayKey(nWarehouseBayKey As Int32) As String
        Dim sSQL As String
        sSQL = "SELECT WarehouseBayId FROM WarehouseBay WHERE WarehouseBayKey = " & nWarehouseBayKey
        GetBayNameFromBayKey = ExecuteQueryToDataTable(sSQL).Rows(0).Item(0)
    End Function

    Protected Function GetSectionKeyFromBayKey(nWarehouseBayKey As Int32) As String
        Dim sSQL As String
        sSQL = "SELECT WarehouseSectionKey FROM WarehouseBay WHERE WarehouseBayKey = " & nWarehouseBayKey
        GetSectionKeyFromBayKey = ExecuteQueryToDataTable(sSQL).Rows(0).Item(0)
    End Function

    Protected Function GetSectionNameFromSectionKey(nWarehouseSectionKey As Int32) As String
        Dim sSQL As String
        sSQL = "SELECT WarehouseSectionId FROM WarehouseSection WHERE WarehouseSectionKey = " & nWarehouseSectionKey
        GetSectionNameFromSectionKey = ExecuteQueryToDataTable(sSQL).Rows(0).Item(0)
    End Function

    Protected Function GetRackKeyFromSectionKey(nWarehouseSectionKey As Int32) As String
        Dim sSQL As String
        sSQL = "SELECT WarehouseRackKey FROM WarehouseSection WHERE WarehouseSectionKey = " & nWarehouseSectionKey
        GetRackKeyFromSectionKey = ExecuteQueryToDataTable(sSQL).Rows(0).Item(0)
    End Function

    Protected Function GetRackNameFromRackKey(nWarehouseRackKey As Int32) As String
        Dim sSQL As String
        sSQL = "SELECT WarehouseRackId FROM WarehouseRack WHERE WarehouseRackKey = " & nWarehouseRackKey
        GetRackNameFromRackKey = ExecuteQueryToDataTable(sSQL).Rows(0).Item(0)
    End Function

    Protected Function GetWarehouseKeyFromRackKey(nWarehouseRackKey As Int32) As String
        Dim sSQL As String
        sSQL = "SELECT WarehouseKey FROM WarehouseRack WHERE WarehouseRackKey = " & nWarehouseRackKey
        GetWarehouseKeyFromRackKey = ExecuteQueryToDataTable(sSQL).Rows(0).Item(0)
    End Function

    Protected Function GetWarehouseNameFromWarehouseKey(nWarehouseKey As Int32) As String
        Dim sSQL As String
        sSQL = "SELECT WarehouseId FROM Warehouse WHERE WarehouseKey = " & nWarehouseKey
        GetWarehouseNameFromWarehouseKey = ExecuteQueryToDataTable(sSQL).Rows(0).Item(0)
    End Function

    Protected Sub btngo_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ReadProductFile()
    End Sub
   
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>QL Import Utility</title>
</head>
<body>
    <form id="form1" runat="server">
    Warehouse Section Key
    <asp:TextBox ID="tbWarehouseSectionKey" runat="server">1051</asp:TextBox>
&nbsp;(1004 on VOSTRO)<br />
    Filename:
    <asp:TextBox ID="tbFilename" runat="server" Width="362px"></asp:TextBox>
    <br />
    <asp:Button ID="btngo" runat="server" onclick="btngo_Click" Text="go" Width="127px" />
    <br />
    <br />
    <asp:TextBox ID="tbLog" runat="server" Rows="10" TextMode="MultiLine" Width="100%" Font-Names="Verdana" Font-Size="XX-Small"/>
    </form>
</body>
</html>