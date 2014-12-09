<%@ Page Language="VB" Theme="AIMSDefault" ValidateRequest="false" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="Microsoft.Win32" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<script runat="server">

    ' WHAT MUST BE STORED IN THE PRODUCT TO LINK IT WITH THE VPO??? (a) SPONumber in LanguageID, (b) SupplierID in Misc2, Flag2
    
    ' Q: Will the ItemList always be coordinated with the VPO list?
    
    ' ClientData_CSN_ProductList
    ' [id] [int] IDENTITY(1,1) NOT NULL,
    ' [3PLSUID] [varchar](55) NOT NULL,
    ' [PartNo] [varchar](55) NOT NULL,
    ' [SKUDescription] [varchar](55) NOT NULL,
    ' [OptionSetDescription] [varchar](55) NOT NULL,
    ' [TrueSupplierSUID] [varchar](55) NOT NULL,
    ' [TrueSupplierName] [varchar](55) NOT NULL,
    ' [CSNSKU] [varchar](55) NOT NULL,
    ' [Inactive] [int] NOT NULL,
    ' [Discontinued] [int] NOT NULL,
    ' [CreatedOn] [smalldatetime] NOT NULL,
    ' [CreatedBy] [int] NOT NULL,
    
    ' ClientData_CSN_VPOReceipts
    ' [id] [int] IDENTITY(1,1) NOT NULL,
    ' [DateReceived] [smalldatetime] NOT NULL,
    ' [PONumber] [varchar](30) NOT NULL,
    ' [SupplierID] [varchar](30) NOT NULL,
    ' [SupplierPartNumber] [varchar](30) NOT NULL,
    ' [QuantityReceived] [int] NOT NULL,
    ' [CreatedBy] [int] NOT NULL
    
    ' ClientData_CSN_VendorPurchaseOrders
    ' [id] [int] IDENTITY(1,1) NOT NULL,
    ' [CreatedOn] [smalldatetime] NOT NULL,
    ' [SupplierID] [varchar](55) NOT NULL,
    ' [SupplierName] [varchar](55) NOT NULL,
    ' [SPONumber] [varchar](30) NOT NULL,
    ' [SPOSentDate] [smalldatetime] NOT NULL,
    ' [SPOEstShipDate] [smalldatetime] NOT NULL,
    ' [SupplierPartNumber] [varchar](55) NOT NULL,
    ' [CSNSKU] [varchar](55) NOT NULL,
    ' [SKUDescription] [varchar](55) NOT NULL,
    ' [WSCost] [money] NOT NULL,
    ' [QuantityOrdered] [int] NOT NULL,
    ' [TotalQuantityReceived] [int] NOT NULL,
    ' [RemainingOnOrder] [int] NOT NULL,
    ' [Closed] [char](1) NOT NULL,
    ' [CreatedBy] [int] NOT NULL

    ' ClientData_CSN_AuditTrail
    ' [id] [int] IDENTITY(1,1) NOT NULL,
    ' [CreatedOn] [smalldatetime] NOT NULL,
    ' [Code] [varchar](30) NOT NULL,
    ' [Description] [varchar](1000) NOT NULL,
    ' [CreatedBy] [int] NOT NULL

    Const LOG_CODE_GOODSINERROR As String = "GOODSINERROR"
    Const LOG_CODE_GOODSINSUCCESS As String = "GOODSINSUCCESS"
    
    
    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsNumeric(Session("UserKey")) Then
            Server.Transfer("session_expired.aspx")
        End If
        If Not IsPostBack Then
            Call InitCustomerKey()
            Call InitVPOs()
            Call InitRecentGoodsInGrid()
            ddlProducts.Focus()
            pnlProductCode.Visible = False
            ddlVPO.Focus()
            Call PopulateWarehouseDropdown()
        End If
    End Sub
    
    Protected Sub Log(ByVal sCode As String, ByVal sDescription As String)
        Dim sSQL As String = "INSERT INTO ClientData_CSN_AuditTrail (CreatedOn, Code, Description, CreatedBy) VALUES (GETDATE(), '" & sCode & "', '" & sDescription & "', " & Session("UserKey") & ")"
        Call ExecuteQueryToDataTable(sSQL)
    End Sub
    
    Protected Sub InitCustomerKey()
        'Try
        pnCustomerKey = GetRegistryValue(RegistryHive.LocalMachine, "SOFTWARE\CourierSoftware\CSN", "CSNCustomerKey")
        'Catch ex As Exception
        'pnCustomerKey = 16
        'End Try
    End Sub
    
    Protected Sub InitVPOs()
        Dim oListItemCollection As ListItemCollection = ExecuteQueryToListItemCollection("SELECT DISTINCT SPONumber FROM ClientData_CSN_VendorPurchaseOrders", "SPONumber", "SPONumber")
        ddlVPO.Items.Clear()
        ddlVPO.Items.Add(New ListItem("- please select -", 0))
        For Each li As ListItem In oListItemCollection
            ddlVPO.Items.Add(li)
        Next
    End Sub
    
    Protected Sub InitProducts()
        Dim sSQL As String = "SELECT pl.PartNo + ' ' + pl.SKUDescription + ' ' + pl.OptionSetDescription 'ProductCode', PartNo FROM ClientData_CSN_ProductList pl INNER JOIN ClientData_CSN_VendorPurchaseOrders vpo ON pl.PartNo = vpo.SupplierPartNumber WHERE vpo.SPONumber = '" & ddlVPO.SelectedValue.Replace("'", "''") & "' ORDER BY pl.PartNo"
        Dim oListItemCollection As ListItemCollection = ExecuteQueryToListItemCollection(sSQL, "ProductCode", "PartNo")
        ddlProducts.Items.Clear()
        ddlProducts.Items.Add(New ListItem("- please select -", 0))
        For Each li As ListItem In oListItemCollection
            ddlProducts.Items.Add(li)
        Next
    End Sub

    Protected Sub InitRecentGoodsInGrid()
        Dim sSQL As String = "SELECT  [id], CONVERT(VARCHAR(9), DateReceived, 6) + ' ' + SUBSTRING((CONVERT(VARCHAR(8), DateReceived, 108)),1,5) 'Rcvd', PONumber, SupplierID, SupplierPartNumber, QuantityReceived, up.FirstName + ' ' + up.LastName 'EnteredBy' FROM ClientData_CSN_VPOReceipts vpor INNER JOIN UserProfile up ON vpor.CreatedBy = up.[Key] ORDER BY [id]"
        Dim dtRecentGoodsIn As DataTable = ExecuteQueryToDataTable(sSQL)
        gvRecentGoodsIn.DataSource = dtRecentGoodsIn
        gvRecentGoodsIn.DataBind()
    End Sub

    Protected Sub btnSaveGoodsInRecord_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call SaveGoodsInRecord()
    End Sub

    Protected Sub SaveGoodsInRecord()
        If ddlProducts.SelectedIndex > 0 Then
            If Not (IsNumeric(tbQtyRcvd.Text) AndAlso CInt(tbQtyRcvd.Text) > 0) Then
                WebMsgBox.Show("Please enter a quantity greater than zero.")
                tbQtyRcvd.Focus()
            Else
                Dim sSQL As String
                sSQL = "SELECT * FROM ClientData_CSN_VendorPurchaseOrders WHERE SupplierPartNumber = '" & ddlProducts.SelectedValue & "'"
                sSQL = "SELECT * FROM ClientData_CSN_VendorPurchaseOrders WHERE SupplierPartNumber = '" & ddlProducts.SelectedValue & "' AND SPONumber = '" & ddlVPO.SelectedItem.Text.Replace("'", "''") & "'"

                Dim dtVPO As DataTable = ExecuteQueryToDataTable(sSQL)
                If dtVPO.Rows.Count = 0 Then
                    WebMsgBox.Show("Oops - the product has disappeared from the VPO list since you selected it, possibly due to a conflict with the automatic update process.")      ' product has disappeared from list - maybe the list was replaced by ServiceCSN
                    Call Log(LOG_CODE_GOODSINERROR, "Product disappeared from VPO list")
                ElseIf dtVPO.Rows.Count > 1 Then
                    WebMsgBox.Show("Arghh! - found two or more instances of this product in the VPO list - something is amiss - please contact development.")      ' product has disappeared from list - maybe the list was replaced by ServiceCSN
                    Call Log(LOG_CODE_GOODSINERROR, "Found two or more instances of product in the VPO list (" & ddlProducts.SelectedValue & ", " & ddlVPO.SelectedItem.Text & ")")
                Else
                    Dim drVPOEntry As DataRow = dtVPO.Rows(0)
                    sSQL = "SELECT LogisticProductKey, ProductDate FROM LogisticProduct WHERE ProductCode = '" & ddlProducts.SelectedValue.Replace("'", "''") & "' AND CustomerKey = " & pnCustomerKey & " AND LanguageID = '" & ddlVPO.SelectedItem.Text.Replace("'", "''") & "'"
                    Dim dtExistingProduct As DataTable = ExecuteQueryToDataTable(sSQL)
                    Dim bContinue As Boolean = True
                    lblStorageProductCodeNote.Text = String.Empty
                    pnlProductCode.Visible = True
                    If dtExistingProduct.Rows.Count = 0 Then
                        pnProductKey = AddNewProduct(dtVPO.Rows(0))
                        If pnProductKey = 0 Then
                            bContinue = False
                            lblStorageProductCode.Text = "COULD NOT CREATE PRODUCT!!!"
                            Call Log(LOG_CODE_GOODSINERROR, "Could not create product " & ddlProducts.SelectedValue & ", " & ddlVPO.SelectedItem.Text & ")")
                            If Not Server.MachineName.ToLower.Contains("sprint") Then
                                pnlLocation.Visible = False
                            End If
                        Else
                            lblStorageProductCode.Text = ddlProducts.SelectedValue
                            lblStorageProductDate.Text = Format(DateTime.Today, "dd/MM/yyyy")
                            lblStorageProductCodeNote.Text = "(newly created product)"
                            Call Log(LOG_CODE_GOODSINSUCCESS, "Created product " & lblStorageProductCode.Text & ", " & lblStorageProductCode.Text & ", " & pnProductKey.ToString & ")")
                            If Not Server.MachineName.ToLower.Contains("sprint") Then
                                pnlLocation.Visible = True
                            End If
                        End If
                    ElseIf dtExistingProduct.Rows.Count = 1 Then
                        lblStorageProductCode.Text = ddlProducts.SelectedValue
                        lblStorageProductDate.Text = dtExistingProduct.Rows(0).Item("ProductDate")
                        pnProductKey = dtExistingProduct.Rows(0).Item("LogisticProductKey")
                        lblStorageProductCodeNote.Text = "(product created for an earlier goods in event - internal code " & pnProductKey & ")"
                        Call Log(LOG_CODE_GOODSINSUCCESS, "Using previously created product " & lblStorageProductCode.Text & ", " & lblStorageProductCode.Text & ", " & pnProductKey.ToString & ")")
                        If Not Server.MachineName.ToLower.Contains("sprint") Then
                            pnlLocation.Visible = True
                        End If
                    Else
                        WebMsgBox.Show("Error - found two or more products with the same Stocking Purchase Order number. Please contact development.")
                        Call Log(LOG_CODE_GOODSINERROR, "ound two or more products with the same Stocking Purchase Order number " & ddlProducts.SelectedValue & ", " & ddlVPO.SelectedItem.Text & ")")
                        bContinue = False
                    End If
                    If bContinue Then
                        Dim sbSQL As New StringBuilder
                        sbSQL.Append("INSERT INTO ClientData_CSN_VPOReceipts")
                        sbSQL.Append(" ")
                        sbSQL.Append("(")
                        sbSQL.Append("DateReceived")
                        sbSQL.Append(",")
                        sbSQL.Append("PONumber")
                        sbSQL.Append(",")
                        sbSQL.Append("SupplierID")
                        sbSQL.Append(",")
                        sbSQL.Append("SupplierPartNumber")
                        sbSQL.Append(",")
                        sbSQL.Append("QuantityReceived")
                        sbSQL.Append(",")
                        sbSQL.Append("CreatedBy")
                        sbSQL.Append(")")
                        sbSQL.Append(" ")
                        sbSQL.Append("VALUES")
                        sbSQL.Append(" ")
                        sbSQL.Append("(")
                        sbSQL.Append("GETDATE()")
                        sbSQL.Append(",")
                        sbSQL.Append("'")
                        sbSQL.Append(drVPOEntry("SPONumber"))
                        sbSQL.Append("'")
                        sbSQL.Append(",")
                        sbSQL.Append("'")
                        sbSQL.Append(drVPOEntry("SupplierID").ToString.Replace("'", "''"))
                        sbSQL.Append("'")
                        sbSQL.Append(",")
                        sbSQL.Append("'")
                        sbSQL.Append(ddlProducts.SelectedValue)
                        sbSQL.Append("'")
                        sbSQL.Append(",")
                        sbSQL.Append(tbQtyRcvd.Text)
                        sbSQL.Append(",")
                        sbSQL.Append(Session("UserKey"))
                        sbSQL.Append(")")
                        Call ExecuteQueryToDataTable(sbSQL.ToString)
                        Call InitRecentGoodsInGrid()

                        Call InitVPOs()
                        ddlProducts.Items.Clear()
                        ddlProducts.Enabled = False
                        btnSaveGoodsInRecord.Enabled = False
                        tbQtyToAdd.Text = tbQtyRcvd.Text
                        tbQtyRcvd.Text = String.Empty
                        ddlVPO.Focus()
                    Else
                        WebMsgBox.Show("Error creating product, Goods In NOT recorded. Please inform development.")
                    End If
                End If
                'End If
                'End If
            End If
        Else
        WebMsgBox.Show("Please select a product from the dropdown list.")
        End If
    End Sub
    
    Protected Sub ddlProducts_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        Dim sSQL As String
        If ddl.SelectedIndex > 0 Then
            sSQL = "SELECT * FROM ClientData_CSN_ProductList WHERE PartNo = '" & ddl.SelectedValue & "'"
            Dim dtProduct As DataTable = ExecuteQueryToDataTable(sSQL)
            If dtProduct.Rows.Count = 0 Then
                WebMsgBox.Show("Oops - the product has disappeared from the database since you selected it, possibly due to a conflict with the automatic update process.")      ' product has disappeared from list - maybe the list was replaced by ServiceCSN
            ElseIf dtProduct.Rows.Count > 1 Then
                WebMsgBox.Show("Arghh! - found two or more instances of this product in the database - something is amiss - please contact development.")      ' product has disappeared from list - maybe the list was replaced by ServiceCSN
            End If
            btnSaveGoodsInRecord.Enabled = True

            sSQL = "SELECT * FROM ClientData_CSN_VendorPurchaseOrders WHERE SPONumber = '" & ddlVPO.SelectedValue & "' AND SupplierPartNumber = '" & ddl.SelectedValue & "'"
            Dim dtVPO As DataTable = ExecuteQueryToDataTable(sSQL)
            If dtVPO.Rows.Count = 0 Then
                WebMsgBox.Show("Oops - cannot identify the purchase order for this product. This could be due to a conflict with the automatic update process.")
            ElseIf dtVPO.Rows.Count > 1 Then
                WebMsgBox.Show("Oops - matched more than one purchase order entry - something is amiss - please contact development.")
            Else
                Dim drVPO As DataRow = dtVPO.Rows(0)
                lblTotalExpected.Text = drVPO("QuantityOrdered")
                lblReceivedSoFar.Text = drVPO("TotalQuantityReceived")
            End If
        Else
            btnSaveGoodsInRecord.Enabled = False
            lblReceivedSoFar.Text = ""
            lblTotalExpected.Text = ""
        End If
        tbQtyRcvd.Focus()
        pnlProductCode.Visible = False
        Call InitProductCodePanel()
    End Sub
    
    Protected Sub InitProductCodePanel()
        lblStorageProductCode.Text = String.Empty
        lblStorageProductDate.Text = String.Empty
    End Sub
    
    Protected Function AddNewProduct(ByVal drVPO As DataRow) As Int32
        Dim oConn As New SqlConnection(gsConn)
        Dim oTrans As SqlTransaction
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Product_AddWithAccessControl9", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
  
        Dim paramUserKey As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int, 4)
        paramUserKey.Value = CLng(Session("UserKey"))
        oCmd.Parameters.Add(paramUserKey)
        
        Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
        paramCustomerKey.Value = 741
        oCmd.Parameters.Add(paramCustomerKey)
  
        Dim paramProductCode As SqlParameter = New SqlParameter("@ProductCode", SqlDbType.NVarChar, 25)
        paramProductCode.Value = ddlProducts.SelectedValue
        oCmd.Parameters.Add(paramProductCode)
        
        Dim paramProductDate As SqlParameter = New SqlParameter("@ProductDate", SqlDbType.NVarChar, 10)
        paramProductDate.Value = Format(DateTime.Today, "dd/MM/yyyy")
        oCmd.Parameters.Add(paramProductDate)
  
        Dim paramMinimumStockLevel As SqlParameter = New SqlParameter("@MinimumStockLevel", SqlDbType.Int, 4)
        paramMinimumStockLevel.Value = 0
        oCmd.Parameters.Add(paramMinimumStockLevel)
        
        Dim paramDescription As SqlParameter = New SqlParameter("@ProductDescription", SqlDbType.NVarChar, 300)
        Dim sSQL As String = "SELECT SKUDescription + ' ' + OptionSetDescription FROM ClientData_CSN_ProductList WHERE PartNo = '" & ddlProducts.SelectedValue & "'"
        Try
            paramDescription.Value = ExecuteQueryToDataTable(sSQL).Rows(0).Item(0)
        Catch ex As Exception
            paramDescription.Value = "Description not retrieved"
        End Try
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
        paramUnitValue.Value = drVPO("WSCost")
        oCmd.Parameters.Add(paramUnitValue)
        
        Dim paramUnitValue2 As SqlParameter = New SqlParameter("@UnitValue2", SqlDbType.Money, 8)
        paramUnitValue2.Value = drVPO("WSCost")
        oCmd.Parameters.Add(paramUnitValue2)

        Dim paramLanguage As SqlParameter = New SqlParameter("@LanguageId", SqlDbType.NVarChar, 20)
        paramLanguage.Value = drVPO("SPONumber")
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
        paramMisc2.Value = drVPO("SupplierID")
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
  
        Dim paramFlag1 As SqlParameter = New SqlParameter("@Flag1", SqlDbType.Bit)
        paramFlag1.Value = 0
        oCmd.Parameters.Add(paramFlag1)

        Dim paramFlag2 As SqlParameter = New SqlParameter("@Flag2", SqlDbType.Bit)
        paramFlag2.Value = 1
        oCmd.Parameters.Add(paramFlag2)

        Dim paramDefaultAccessFlag As SqlParameter = New SqlParameter("@DefaultAccessFlag", SqlDbType.Bit)
        paramDefaultAccessFlag.Value = 0 ' CHECK THIS !!!!!!!!!!!!!!!!!!!
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
            AddNewProduct = CLng(oCmd.Parameters("@ProductKey").Value)
        Catch ex As SqlException
            If ex.Number = 2627 Then
                AddNewProduct = -1   ' a record already exists with the same product CODE and DATE combination
            Else
                AddNewProduct = 0
            End If
        Finally
            oConn.Close()
        End Try
    End Function
  

    Protected Sub lnkbtnRemoveItem_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim lnkbtn As LinkButton = sender
        Call ExecuteQueryToDataTable("DELETE FROM ClientData_CSN_VPOReceipts WHERE [id] = " & lnkbtn.CommandArgument)
        Call InitRecentGoodsInGrid()
    End Sub

    Protected Function GetRegistryValue(ByVal Hive As RegistryHive, ByVal Key As String, ByVal ValueName As String) As String
        Dim objParent As RegistryKey = Nothing
        Dim objSubkey As RegistryKey = Nothing
        Dim sAns As String
        Dim ErrInfo As String = String.Empty

        Select Case Hive
            Case RegistryHive.ClassesRoot
                objParent = Registry.ClassesRoot
            Case RegistryHive.CurrentConfig
                objParent = Registry.CurrentConfig
            Case RegistryHive.CurrentUser
                objParent = Registry.CurrentUser
            Case RegistryHive.LocalMachine
                objParent = Registry.LocalMachine
            Case RegistryHive.PerformanceData
                objParent = Registry.PerformanceData
            Case RegistryHive.Users
                objParent = Registry.Users
        End Select

        Try
            objSubkey = objParent.OpenSubKey(Key)
            'if can't be found, object is not initialized
            If Not objSubkey Is Nothing Then
                sAns = (objSubkey.GetValue(ValueName))
            End If
        Catch ex As Exception
            sAns = "Error"
            'ErrInfo = ex.Message
        Finally
            If ErrInfo = "" And sAns = "" Then
                sAns = "No value found for requested registry key"
            End If
        End Try
        GetRegistryValue = sAns
        
    End Function
    
    Protected Sub ddlVPO_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        If ddl.SelectedIndex > 0 Then
            ddlProducts.Enabled = True
            Call InitProducts()
            ddlProducts.Focus()
        Else
            ddlProducts.Enabled = False
            ddlProducts.Items.Clear()
            btnSaveGoodsInRecord.Enabled = False
            lblReceivedSoFar.Text = ""
            lblTotalExpected.Text = ""
        End If
        pnlProductCode.Visible = False
        pnlConfirmation.Visible = False
    End Sub

    Protected Sub PopulateWarehouseDropdown()
        Dim sSQL As String = "SELECT WarehouseId, WarehouseKey FROM Warehouse WHERE DeletedFlag = 'N' ORDER BY WarehouseId"
        Dim oListItemCollection As ListItemCollection = ExecuteQueryToListItemCollection(sSQL, "WarehouseId", "WarehouseKey")
        ddlWarehouse.Items.Add(New ListItem("- please select -", 0))
        For Each li As ListItem In oListItemCollection
            ddlWarehouse.Items.Add(li)
        Next
        ddlRack.SelectedIndex = 0
        ddlSection.SelectedIndex = 0
        ddlBay.SelectedIndex = 0
    End Sub

    Protected Sub ddlWarehouse_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call InitRackDropdown()
        'Call ClearRackDropdown()
        Call ClearSectionDropdown()
        Call ClearBayDropdown()
        'ddlSection.SelectedIndex = 0
        'ddlBay.SelectedIndex = 0
        ddlRack.Focus()
    End Sub

    Protected Sub InitRackDropdown()
        ddlRack.Items.Clear()
        Dim sSQL As String = "SELECT WarehouseRackId, WarehouseRackKey FROM WarehouseRack WHERE DeletedFlag = 'N' AND WarehouseKey = " & ddlWarehouse.SelectedValue & " ORDER BY WarehouseRackId"
        Dim oListItemCollection As ListItemCollection = ExecuteQueryToListItemCollection(sSQL, "WarehouseRackId", "WarehouseRackKey")
        ddlRack.Items.Add(New ListItem("- please select -", 0))
        For Each li As ListItem In oListItemCollection
            ddlRack.Items.Add(li)
        Next
    End Sub
    
    Protected Sub ddlRack_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call InitSectionDropdown()
        ' Call ClearSectionDropdown()
        Call ClearBayDropdown()
        ddlSection.Focus()
    End Sub

    Protected Sub InitSectionDropdown()
        ddlSection.Items.Clear()
        Dim sSQL As String = "SELECT WarehouseSectionId, WarehouseSectionKey FROM WarehouseSection WHERE DeletedFlag = 'N' AND WarehouseRackKey = " & ddlRack.SelectedValue & " ORDER BY WarehouseSectionId"
        Dim oListItemCollection As ListItemCollection = ExecuteQueryToListItemCollection(sSQL, "WarehouseSectionId", "WarehouseSectionKey")
        ddlSection.Items.Add(New ListItem("- please select -", 0))
        For Each li As ListItem In oListItemCollection
            ddlSection.Items.Add(li)
        Next
    End Sub
    
    Protected Sub ddlSection_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call InitBayDropdown()
        ' Call ClearBayDropdown()     ' ??????????????
        ddlSection.Focus()
    End Sub
    
    Protected Sub InitBayDropdown()
        ddlBay.Items.Clear()
        Dim sSQL As String = "SELECT WarehouseBayId, WarehouseBayKey FROM WarehouseBay WHERE DeletedFlag = 'N' AND WarehouseSectionKey = " & ddlSection.SelectedValue & " ORDER BY WarehouseBayId"
        Dim oListItemCollection As ListItemCollection = ExecuteQueryToListItemCollection(sSQL, "WarehouseBayId", "WarehouseBayKey")
        ddlBay.Items.Add(New ListItem("- please select -", 0))
        For Each li As ListItem In oListItemCollection
            ddlBay.Items.Add(li)
        Next
    End Sub

    Protected Sub ClearRackDropdown()
        If ddlRack.Items.Count > 0 Then
            ddlRack.SelectedIndex = 0
        End If
    End Sub
    
    Protected Sub ClearSectionDropdown()
        If ddlSection.Items.Count > 0 Then
            ddlSection.SelectedIndex = 0
        End If
    End Sub
    
    Protected Sub ClearBayDropdown()
        If ddlBay.Items.Count > 0 Then
            ddlBay.SelectedIndex = 0
        End If
    End Sub
    
    Protected Sub btnAddToLocation_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsNumeric(tbQtyToAdd.Text) Then
            WebMsgBox.Show("Enter a numeric quantity to add.")
            Exit Sub
        End If
        Call AddQuantity()
    End Sub

    Protected Sub AddQuantity()
        If ddlBay.Items.Count > 0 AndAlso ddlBay.SelectedValue > 0 Then
            Dim sSQL As String
            sSQL = "SELECT LogisticProductQuantity FROM LogisticProductLocation WHERE LogisticProductKey = " & pnProductKey & " AND WarehouseBayKey = " & ddlBay.SelectedValue
            Dim oDT As DataTable = ExecuteQueryToDataTable(sSQL)
            If oDT.Rows.Count > 0 Then
                If oDT.Rows.Count = 1 Then
                    Dim nQuantity As Int32 = CInt(oDT.Rows(0).Item(0))
                    nQuantity = nQuantity + CInt(tbQtyRcvd.Text)
                    If nQuantity >= 0 Then
                        sSQL = "UPDATE LogisticProductLocation SET LogisticProductQuantity = " & nQuantity & " WHERE WarehouseBayKey = " & ddlBay.SelectedValue & " AND LogisticProductKey = " & pnProductKey
                    Else
                        WebMsgBox.Show("Quantity adjustment entered would make total quantity in this location negative.")
                    End If
                Else
                    WebMsgBox.Show("Error - multiple instances of one product in a single location.")
                End If
            Else
                sSQL = "INSERT INTO LogisticProductLocation (LogisticProductKey, WarehouseBayKey, LogisticProductQuantity, DateStored) VALUES ("
                sSQL &= pnProductKey
                sSQL &= ", "
                sSQL &= ddlBay.SelectedValue
                sSQL &= ", "
                sSQL &= tbQtyToAdd.Text
                sSQL &= ", GETDATE())"
            End If
            Call ExecuteQueryToDataTable(sSQL)
            lblConfirmation.Text = tbQtyToAdd.Text & " units of " & lblStorageProductCode.Text & " " & lblStorageProductDate.Text & " stored in Warehouse " & ddlWarehouse.SelectedItem.Text & ", Rack " & ddlRack.SelectedItem.Text & ", Section " & ddlSection.SelectedItem.Text & ", Bay " & ddlBay.SelectedItem.Text
            WebMsgBox.Show(tbQtyToAdd.Text & " units of " & lblStorageProductCode.Text & " " & lblStorageProductDate.Text & " stored in Warehouse " & ddlWarehouse.SelectedItem.Text & ", Rack " & ddlRack.SelectedItem.Text & ", Section " & ddlSection.SelectedItem.Text & ", Bay " & ddlBay.SelectedItem.Text)
            pnlConfirmation.Visible = True
            pnlProductCode.Visible = False
            If Not Server.MachineName.ToLower.Contains("sprint") Then
                pnlLocation.Visible = False
            End If
            'tbLog.Text += "Added to " & ddlCustomer.SelectedItem.Text & " " & tbQtyToAdd.Text & " of " & ddlProduct.SelectedItem.Text & " (" & lblLogisticProductKey.Text & ") to Warehouse " & ddlWarehouse.SelectedItem.Text & ", Rack " & ddlRack.SelectedItem.Text & ", Section " & ddlSection.SelectedItem.Text & ", Bay " & ddlBay.SelectedItem.Text & Environment.NewLine
            'tbQtyToAdd.Text = String.Empty
            'Call SetLocationGrid()
            Call Log(LOG_CODE_GOODSINSUCCESS, "Added qty " & tbQtyToAdd.Text & " to product " & pnProductKey & " in location " & ddlWarehouse.SelectedItem.Text & ", " & ddlRack.SelectedItem.Text & ", " & ddlSection.SelectedItem.Text & ", " & ddlBay.SelectedItem.Text & ")")

        Else
            WebMsgBox.Show("Please select a location.")
        End If
    End Sub
    
    Property pnCustomerKey() As Int32
        Get
            Dim o As Object = ViewState("CSNIA_CustomerKey")
            If o Is Nothing Then
                Return -1
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Int32)
            ViewState("CSNIA_CustomerKey") = Value
        End Set
    End Property

    Property pnProductKey() As Int32
        Get
            Dim o As Object = ViewState("CSNIA_ProductKey")
            If o Is Nothing Then
                Return -1
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Int32)
            ViewState("CSNIA_ProductKey") = Value
        End Set
    End Property

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
    
    Protected Sub lnkbtnCloseConfirmationPanel_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        pnlConfirmation.Visible = False
    End Sub
    
    Protected Sub lnkbtnDeliveryForecast_Click(sender As Object, e As System.EventArgs)
        Dim lnkbtn As LinkButton = sender
        If lnkbtn.Text.Contains("show") Then
            lnkbtn.Text = "hide delivery forecast"
            pnlDeliveryForecast.Visible = True
            Call BindDeliveryForecast()
        Else
            lnkbtn.Text = "show delivery forecast"
            pnlDeliveryForecast.Visible = False
        End If
    End Sub
    
    Protected Sub BindDeliveryForecast()
        Dim sSQL As String = "SELECT SupplierName 'Supplier', SPONumber 'SPO #', CONVERT(Varchar(11), SPOSentDate, 106) 'SPO Sent',  CONVERT(Varchar(11), SPOEstShipDate, 106) 'Est. Ship Date', SupplierPartNumber 'Product Code', SKUDescription 'Description', QuantityOrdered 'Qty Ordered', TotalQuantityReceived 'Rcvd so far', RemainingOnOrder 'Remaining', Closed FROM ClientData_CSN_VendorPurchaseOrders ORDER BY SPOEstShipDate"
        Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
        gvDeliveryForecast.DataSource = dt
        gvDeliveryForecast.DataBind()
    End Sub
    
</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Wayfair Goods In</title>
    <style type="text/css">
        .style1
        {
            width: 100%;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
    <main:Header ID="ctlHeader" runat="server" />
    <table width="100%" cellpadding="0" cellspacing="0">
        <tr class="bar_accounthandler">
            <td style="width: 50%; white-space: nowrap">
            </td>
            <td style="width: 50%; white-space: nowrap" align="right">
            </td>
        </tr>
    </table>
    <p>
        <asp:Label ID="Label4" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Wayfair Goods In" Font-Bold="True" />
    </p>
    <p>
        <asp:Label ID="Label18" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="INSTRUCTIONS: 1. Select the purchase order 2. Select a product 3. Enter the Quantity Received" Font-Bold="True" Font-Italic="True" />
    &nbsp;<asp:Label ID="Label27" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="4. Click the Save button - you should see the Product Code created for this purchase order" Font-Bold="True" Font-Italic="True" />
        <br />
        <asp:Label ID="Label28" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="5. Now enter the GOODS IN on AIMS, using the product code &amp; value/date shown." Font-Bold="True" Font-Italic="True" />
    </p>
    <p>
        <asp:Label ID="Label22" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="If the purchase order or incoming product is not listed, contact the Operations Manager." Font-Bold="True" Font-Italic="True" />
    </p>
    <br />

    <asp:Panel ID="pnlAdjustment" runat="server" Width="100%" Font-Names="Verdana" Font-Size="XX-Small" GroupingText="Record Goods In">
        <table class="style1">
            <tr>
                <td style="width: 2%">&nbsp;
                </td>
                <td style="width: 20%" align="right">
                    <asp:Label ID="Label24" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small" Text=" Purchase Order #:" />
                </td>
                <td style="width: 2%">
                    &nbsp;
                </td>
                <td style="width: 20%">
                    &nbsp;
                    <asp:DropDownList ID="ddlVPO" runat="server" AutoPostBack="True" Font-Names="Verdana" Font-Size="XX-Small" Width="100%" onselectedindexchanged="ddlVPO_SelectedIndexChanged">
                    </asp:DropDownList>
                </td>
                <td style="width: 2%">
                    &nbsp;
                </td>
                <td align="left" colspan="3">
                    &nbsp;</td>
                <td style="width: 2%">
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td>
                    &nbsp;
                </td>
                <td align="right">
                    <asp:Label ID="Label19" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small" Text="Wayfair Product:" />
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
                <td colspan="3">
                    &nbsp;
                    <asp:DropDownList ID="ddlProducts" runat="server" AutoPostBack="True" Font-Names="Verdana" Font-Size="XX-Small" OnSelectedIndexChanged="ddlProducts_SelectedIndexChanged" Width="100%">
                    </asp:DropDownList>
                    &nbsp; &nbsp;
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
                </td>
                <td align="right">
                    <asp:Label ID="Label13" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small" Text="Quantity Received:" />
                </td>
                <td>
                </td>
                <td colspan="2">
                    <asp:TextBox ID="tbQtyRcvd" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="16%" MaxLength="6"></asp:TextBox>
                    &nbsp;<asp:Label ID="Label31" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small" Text="(tot qty expected *: " />
                    <asp:Label ID="lblTotalExpected" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small" Text="X" />
                    <asp:Label ID="Label32" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small" Text=")" />
                </td>
                <td>
                   <table class="style1">
                        <tr>
                            <td>
                                <asp:Label ID="Label33" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small" Text="(previously rcvd *: " />
                                <asp:Label ID="lblReceivedSoFar" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small" Text="X" />
                                <asp:Label ID="Label34" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small" Text=")" />
                            </td>
                            <td align="right">
                                <asp:Button ID="btnSaveGoodsInRecord" runat="server" Enabled="False" EnableTheming="True" OnClick="btnSaveGoodsInRecord_Click" Text="Save" Width="180px" />
                            </td>
                        </tr>
                    </table>
                </td>
                <td>
                </td>
                <td>
                    &nbsp;</td>
                <td>
                </td>
            </tr>
        </table>
    </asp:Panel>
        <asp:Label ID="Label36" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="* 'total qty expected' &amp; 'previously rcvd' do not include today's bookings." Font-Bold="True" Font-Italic="True" />
    <br />
        <br />
    <asp:Panel ID="pnlProductCode" runat="server" Width="100%" Font-Names="Verdana" Font-Size="XX-Small" Visible="false">
        <asp:Label ID="Label20" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Store under Product Code:" Font-Bold="True" />
        &nbsp;<asp:Label ID="lblStorageProductCode" runat="server" Font-Names="Verdana" Font-Size="Small" Font-Bold="True" ForeColor="Red" />
        &nbsp;<asp:Label ID="Label23" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Product Date:" Font-Bold="True" />
        &nbsp;<asp:Label ID="lblStorageProductDate" runat="server" Font-Names="Verdana" Font-Size="Small" Font-Bold="True" ForeColor="Red" />
        &nbsp;<asp:Label ID="lblStorageProductCodeNote" runat="server" Font-Names="Verdana" Font-Size="XX-Small" />
        <br />
        <br />
        NOTE: You must also <strong>
        <asp:Label ID="Label30" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small" ForeColor="Red" Text="ENTER WAYFAIR GOODS-IN ON AIMS" />
        . Use the product code / value date shown above.</strong><br />
        <br />
    <asp:Panel ID="pnlLocation" runat="server" Width="100%" Font-Names="Verdana" Font-Size="XX-Small" Visible="false">
        <asp:Label ID="Label26" runat="server" Text="Qty:"></asp:Label>
        &nbsp;<asp:TextBox ID="tbQtyToAdd" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="45px"/>
        &nbsp;<asp:Label ID="Label3" runat="server" Text="Warehouse:"></asp:Label>
        <asp:DropDownList ID="ddlWarehouse" runat="server" Font-Names="Verdana" Font-Size="XX-Small"  AutoPostBack="True" onselectedindexchanged="ddlWarehouse_SelectedIndexChanged"/>
        &nbsp;<asp:Label ID="Label25" runat="server" Text="Rack:"></asp:Label>
        &nbsp;<asp:DropDownList ID="ddlRack" runat="server" Font-Names="Verdana" Font-Size="XX-Small"  AutoPostBack="True" onselectedindexchanged="ddlRack_SelectedIndexChanged"/>
        &nbsp;<asp:Label ID="Label5" runat="server" Text="Section:"></asp:Label>
        &nbsp;<asp:DropDownList ID="ddlSection" runat="server" Font-Names="Verdana" Font-Size="XX-Small"  AutoPostBack="True" onselectedindexchanged="ddlSection_SelectedIndexChanged"/>
        &nbsp;<asp:Label ID="Label6" runat="server" Text="Bay:"></asp:Label>
        &nbsp;<asp:DropDownList ID="ddlBay" runat="server" Font-Names="Verdana" Font-Size="XX-Small" />
        &nbsp; <asp:Button ID="btnAddToLocation" runat="server" Text="Add to Location" onclick="btnAddToLocation_Click" Width="180px" />
        &nbsp;</asp:Panel>
    </asp:Panel>
    <asp:Panel ID="pnlConfirmation" runat="server" Width="100%" Font-Names="Verdana" Font-Size="XX-Small" Visible="false">
        <asp:Label ID="lblConfirmation" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="Small" ForeColor="Red" />
        &nbsp;<asp:LinkButton ID="lnkbtnCloseConfirmationPanel" runat="server" Font-Names="Verdana" Font-Size="XX-Small" onclick="lnkbtnCloseConfirmationPanel_Click">close</asp:LinkButton>
    </asp:Panel>
    <br />
    <asp:Label ID="Label21" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Recent Goods In:" />
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        <asp:Label ID="Label29" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="The Recent Goods In list is normally cleared overnight." Font-Bold="True" Font-Italic="True" />
    <asp:GridView ID="gvRecentGoodsIn" runat="server" Width="100%" EnableModelValidation="True" AutoGenerateColumns="False" Font-Names="Verdana" Font-Size="XX-Small">
        <Columns>
            <asp:TemplateField>
                <ItemTemplate>
                    <asp:LinkButton ID="lnkbtnRemoveItem" runat="server" Font-Names="Verdana" Font-Size="XX-Small" CommandArgument='<%# Container.DataItem("id")%>' OnClick="lnkbtnRemoveItem_Click" OnClientClick='return confirm("Are you sure you want to delete this entry?\n\nThe product associated with it will NOT be removed.");'>remove</asp:LinkButton>
                </ItemTemplate>
            </asp:TemplateField>
            <asp:BoundField DataField="Rcvd" HeaderText="Rcvd" ReadOnly="True" SortExpression="Rcvd" />
            <asp:BoundField DataField="PONumber" HeaderText="PO #" ReadOnly="True" SortExpression="PONumber" />
            <asp:BoundField DataField="SupplierID" HeaderText="Supplier" ReadOnly="True" SortExpression="SupplierID" />
            <asp:BoundField DataField="SupplierPartNumber" HeaderText="Product Code" ReadOnly="True" SortExpression="SupplierPartNumber" />
            <asp:BoundField DataField="QuantityReceived" HeaderText="Qty" ReadOnly="True" SortExpression="QuantityReceived" />
            <asp:BoundField DataField="EnteredBy" HeaderText="Recorded By" ReadOnly="True" SortExpression="EnteredBy" />
        </Columns>
        <EmptyDataTemplate>
            <div style="text-align: center">
                <asp:Label ID="Label20" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="No Goods In records found" />
            </div>
        </EmptyDataTemplate>
    </asp:GridView>
    <br />
    <asp:LinkButton ID="lnkbtnDeliveryForecast" runat="server" Font-Names="Verdana" Font-Size="XX-Small" onclick="lnkbtnDeliveryForecast_Click">show delivery forecast</asp:LinkButton>
    <asp:Panel ID="pnlDeliveryForecast" runat="server" Width="100%" Font-Names="Verdana" Font-Size="XX-Small" Visible="false" >
        <asp:GridView ID="gvDeliveryForecast" runat="server" CellPadding="2" 
            Font-Names="Verdana" Font-Size="XX-Small" Width="100%">
        </asp:GridView>
        <br />
    </asp:Panel>
    </form>
</body>
</html>