<%@ Page Language="VB" Theme="AIMSDefault" ValidateRequest="false" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Data.SqlTypes" %>
<%@ import Namespace="System.Drawing.Image" %>
<%@ import Namespace="System.Drawing.Color" %>
<%@ import Namespace="System.Globalization" %>
<%@ import Namespace="System.Threading" %>
<%@ import Namespace="System.Collections.Generic" %>
<%@ Register TagPrefix="FCKeditorV2" Namespace="FredCK.FCKeditorV2" Assembly="FredCK.FCKeditorV2" %>

<script runat="server">

    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    Private gbDataBound As Boolean = False
    Private gsSiteType As String = ConfigLib.GetConfigItem_SiteType
    Private gbSiteTypeDefined As Boolean = gsSiteType.Length > 0

    Private  gbExplicitProductPermissions As Boolean

    Const NEW_CATEGORY As String = "- new category -"
    Const NEW_SUBCATEGORY As String = "- new subcategory -"
  
    Const PER_CUSTOMER_CONFIGURATION_NONE As Integer = 0

    Const PER_USERTYPE_OWNER_GROUP_NONE As Integer = 0

    Const CATEGORY_MODE_0_CATEGORIES As Integer = 0
    Const CATEGORY_MODE_2_CATEGORIES As Integer = 2
    Const CATEGORY_MODE_3_CATEGORIES As Integer = 3
  
    Const DISPLAY_MODE_CATEGORY As String = "category"
    Const DISPLAY_MODE_ALL As String = "all"
    Const DISPLAY_MODE_SEARCH As String = "search"

    Protected Sub Page_Load(ByVal Source As Object, ByVal E As EventArgs)
        If Not IsNumeric(Session("CustomerKey")) Then
            Server.Transfer("session_expired.aspx")
        End If
        If Not IsPostBack Then
            Call GetSiteFeatures()
            psProdImageFolder = ConfigLib.GetConfigItem_prod_image_folder
            psProdThumbFolder = ConfigLib.GetConfigItem_prod_thumb_folder
            psProdPDFFolder = ConfigLib.GetConfigItem_prod_pdf_folder
  
            psVirtualJPGFolder = ConfigLib.GetConfigItem_Virtual_JPG_URL
            psVirtualThumbFolder = ConfigLib.GetConfigItem_Virtual_Thumb_URL
            psVirtualPDFFolder = ConfigLib.GetConfigItem_Virtual_PDF_URL

            Call PerCustomerConfiguration()

            Call GetProductNumbers()
            Call ShowMainPanel()
        End If
      
        SqlDataSourceCategoryList.ConnectionString = ConfigLib.GetConfigItem_ConnectionString

        txtSearchCriteriaAllProducts.Attributes.Add("onkeypress", "return clickButton(event,'" + btn_SearchAllProducts.ClientID + "')")
       
        Thread.CurrentThread.CurrentCulture = New CultureInfo("en-GB", False)
        Call CheckVisibility()
        Response.Buffer = True
        Call SetTitle()
    End Sub
  
    Protected Sub SetTitle()
        Dim sTitle As String = Session("SiteTitle")
        If sTitle <> String.Empty Then
            sTitle += " - "
        End If
        Page.Header.Title = sTitle & "Product Manager"
    End Sub
   
    Protected Sub GetSiteFeatures()
        Dim oDataTable As New DataTable
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter("spASPNET_SiteContent2", oConn)
       
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
       
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@Action", SqlDbType.NVarChar, 50))
        oAdapter.SelectCommand.Parameters("@Action").Value = "GET"
       
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SiteKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@SiteKey").Value = Session("SiteKey")
       
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ContentType", SqlDbType.NVarChar, 50))
        oAdapter.SelectCommand.Parameters("@ContentType").Value = "SiteSettings"

        Try
            oAdapter.Fill(oDataTable)
        Catch ex As Exception
            WebMsgBox.Show("GetSiteFeatures: " & ex.Message)
        Finally
            oConn.Close()
        End Try

        Dim dr As DataRow = oDataTable.Rows(0)
        pnCategoryMode = dr("CategoryCount")
        If pnCategoryMode = 2 Then
            Call Set2CategoryLevels()
        ElseIf pnCategoryMode = 3 Then
            Call Set3CategoryLevels()
        Else
            WebMsgBox.Show("Error retrieving category levels - please report this problem to your Account Handler")
        End If
    End Sub

    Protected Sub PerCustomerConfiguration()
        plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_NONE
        If plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_NONE Then
            btnShowCategories.Visible = False
        End If
    End Sub
  
    Protected Sub Set2CategoryLevels()
        pnCategoryMode = CATEGORY_MODE_2_CATEGORIES
        trSubCategory2.Visible = False
        pnlCategorySelection1.Visible = True
    End Sub

    Protected Sub Set3CategoryLevels()
        pnCategoryMode = CATEGORY_MODE_3_CATEGORIES
        trSubCategory2.Visible = True
        pnlCategorySelection2.Visible = True
    End Sub

    Protected Sub HideAllPanels()
        pnlMain.Visible = False
        pnlProductList.Visible = False
        pnlCategorySelection1.Visible = False
        pnlCategorySelection2.Visible = False
        pnlEditProduct.Visible = False
        pnlProductUserProfile.Visible = False
        pnlAssociatedProducts.Visible = False
        pnlProductInactivityAlertStatus.Visible = False
        pnlConfigureProductInactivityAlert.Visible = False
        lblError.Text = ""
    End Sub

    Protected Sub ShowMainPanel()
        Call HideAllPanels()
        pnlMain.Visible = True
    End Sub
  
    Protected Sub ShowProductList()
        Call HideAllPanels()
        pnlMain.Visible = True
        pnlProductList.Visible = True
    End Sub
  
    Protected Sub btn_ShowAllProducts_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowAllProducts()
    End Sub
  
    Protected Sub ShowAllProducts()
        psDisplayMode = DISPLAY_MODE_ALL
        Call HideAllPanels()
        pnlMain.Visible = True
        pnlProductList.Visible = True
        dg_ProductList.CurrentPageIndex = 0
        txtSearchCriteriaAllProducts.Text = ""
        Call BindProductGridDispatcher()
    End Sub
  
    Protected Sub btn_SearchAllProducts_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call SearchAllProducts()
    End Sub
  
    Protected Sub SearchAllProducts()
        psDisplayMode = DISPLAY_MODE_SEARCH
        Call HideAllPanels()
        pnlMain.Visible = True
        pnlProductList.Visible = True
        dg_ProductList.CurrentPageIndex = 0
        Call BindProductGridDispatcher()
    End Sub
  
    Protected Sub lnkbtnShowProductsByCategory_click(ByVal sender As Object, ByVal e As CommandEventArgs)
        If pnCategoryMode = CATEGORY_MODE_2_CATEGORIES Then
            psSubCategory = CStr(e.CommandArgument)
        Else
            psSubSubCategory = CStr(e.CommandArgument)
        End If
        BindProductGridDispatcher()
        Call ShowProductList()
    End Sub
  
    Protected Sub ShowCategories()
        psDisplayMode = DISPLAY_MODE_CATEGORY
        txtSearchCriteriaAllProducts.Text = ""
        Call HideAllPanels()
        pnlMain.Visible = True
        If pnCategoryMode = CATEGORY_MODE_2_CATEGORIES Then
            pnlCategorySelection1.Visible = True
        Else
            pnlCategorySelection2.Visible = True
        End If
        Repeater2.Visible = False
        Repeater2a.Visible = False
        Repeater3a.Visible = False
        Call GetCategories()
    End Sub
  
    Protected Sub ShowProductDetail()
        Call HideAllPanels()
        pnlEditProduct.Visible = True
        Call CheckVisibility()
    End Sub
  
    Protected Sub ShowProductUserProfile()
        Call HideAllPanels()
        Call BindProductUserProfileGrid(txtProductUserSearch.Text, "UserID")
        If txtProductDate.Text.Trim.Length > 0 Then
            lblUserPermissionsProductCode.Text = txtProductCode.Text & "-" & txtProductDate.Text
        Else
            lblUserPermissionsProductCode.Text = txtProductCode.Text
        End If
        pnlProductUserProfile.Visible = True
    End Sub
  
    Protected Sub ShowNewProduct()
        Call HideAllPanels()
        txtProductCode.Focus()
        pnlEditProduct.Visible = True
    End Sub
  
    Protected Sub ShowAssociatedProductsPanel()
        Call HideAllPanels()
        Call InitAssociatedProductsPanel()
        pnlAssociatedProducts.Visible = True
    End Sub

    Protected Sub btn_ShowAllUsers_click(ByVal sender As Object, ByVal e As System.EventArgs)
        txtProductUserSearch.Text = ""
        Call BindProductUserProfileGrid(txtProductUserSearch.Text, "UserID")
    End Sub
  
    Protected Sub btn_SearchUsers_click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call BindProductUserProfileGrid(txtProductUserSearch.Text, "UserID")
    End Sub
  
    Protected Sub btnUploadImage_click(ByVal sender As Object, ByVal e As System.EventArgs)
        If fuBrowseImageFile.PostedFile.FileName.Trim <> "" Then
            Call SaveImage()
        End If
        Call CheckVisibility()
    End Sub
  
    Protected Sub btnUploadPDF_click(ByVal sender As Object, ByVal e As System.EventArgs)
        If fuBrowsePDFFile.PostedFile.FileName.Trim <> "" Then
            Call SavePDF()
        End If
        Call CheckVisibility()
    End Sub
  
    Protected Sub btnShowCategories_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowCategories()
    End Sub

    Protected Sub repeater1_Item_click(ByVal s As Object, ByVal e As RepeaterCommandEventArgs)
        Dim item As RepeaterItem
        For Each item In s.Items
            Dim x As LinkButton = CType(item.Controls(1), LinkButton)
            x.ForeColor = Navy
        Next
        Dim Link As LinkButton = CType(e.CommandSource, LinkButton)
        Link.ForeColor = Blue
    End Sub
  
    Protected Sub lnkbtnShowSubCategories_click(ByVal sender As Object, ByVal e As CommandEventArgs)
        psCategory = CStr(e.CommandArgument)
        Repeater2.Visible = True
        Repeater2a.Visible = True
        Repeater3a.Visible = False
        Call GetSubCategories()
    End Sub
  
    Protected Sub lnkbtnShowSubSubCategories_click(ByVal sender As Object, ByVal e As CommandEventArgs)
        psSubCategory = CStr(e.CommandArgument)
        Repeater3a.Visible = True
        GetSubSubCategories()
    End Sub
  
    Protected Sub GetCategories()
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataSet As New DataSet()
        Dim oAdapter As New SqlDataAdapter("spASPNET_Product_GetCategories", oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Session("CustomerKey")
        lblError.Text = ""
        Try
            oAdapter.Fill(oDataSet, "Categories")
            Dim iCount As Integer = oDataSet.Tables(0).Rows.Count
            If iCount > 0 Then
                If pnCategoryMode = CATEGORY_MODE_2_CATEGORIES Then
                    Repeater1.Visible = True
                    Repeater1.DataSource = oDataSet
                    Repeater1.DataBind()
                Else
                    Repeater1a.Visible = True
                    Repeater1a.DataSource = oDataSet
                    Repeater1a.DataBind()
                End If
            Else
                Repeater1.Visible = False
                Repeater1a.Visible = False
            End If
        Catch ex As SqlException
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
    End Sub
  
    Protected Sub GetSubCategories()
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataSet As New DataSet()
        Dim oAdapter As New SqlDataAdapter("spASPNET_Product_GetSubCategories", oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@Category", SqlDbType.NVarChar, 50))
        oAdapter.SelectCommand.Parameters("@Category").Value = psCategory
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Session("CustomerKey")
        lblError.Text = ""
        Try
            oAdapter.Fill(oDataSet, "SubCategories")
            Dim iCount As Integer = oDataSet.Tables(0).Rows.Count
            If iCount > 0 Then
                If pnCategoryMode = CATEGORY_MODE_2_CATEGORIES Then
                    Repeater2.Visible = True
                    Repeater2.DataSource = oDataSet
                    Repeater2.DataBind()
                Else
                    Repeater2a.Visible = True
                    Repeater2a.DataSource = oDataSet
                    Repeater2a.DataBind()
                End If
            Else
                Repeater2.Visible = False
                Repeater2a.Visible = False
            End If
        Catch ex As SqlException
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
    End Sub
  
    Protected Sub GetSubSubCategories()
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataSet As New DataSet()
        Dim oAdapter As New SqlDataAdapter("spASPNET_Product_GetSubSubCategories2", oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ProductCategory", SqlDbType.NVarChar, 50))
        oAdapter.SelectCommand.Parameters("@ProductCategory").Value = psCategory
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SubCategory", SqlDbType.NVarChar, 50))
        oAdapter.SelectCommand.Parameters("@SubCategory").Value = psSubCategory
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Session("CustomerKey")
        lblError.Text = ""
        Try
            oAdapter.Fill(oDataSet, "SubSubCategories")
            Dim iCount As Integer = oDataSet.Tables(0).Rows.Count
            If iCount > 0 Then
                Repeater3a.Visible = True
                Repeater3a.DataSource = oDataSet
                Repeater3a.DataBind()
            Else
                Repeater3a.Visible = False
            End If
        Catch ex As SqlException
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
    End Sub
  
    Protected Sub DisplayCategories()
        Dim item As RepeaterItem
        For Each item In Repeater1.Items
            Dim x As LinkButton = CType(item.Controls(1), LinkButton)
            x.ForeColor = Navy
        Next
        Repeater2.Visible = False
        Repeater2a.Visible = False
        Repeater3a.Visible = False
        Call ShowCategories()
    End Sub
  
    Protected Function bDatesValid() As Boolean
        bDatesValid = True
        Dim sExpiryDate As String = tbExpiryDate.Text.Trim
        If sExpiryDate <> String.Empty Then
            Try
                Dim dtExpiryDate As DateTime = DateTime.Parse(sExpiryDate)
                If dtExpiryDate.Year < 2000 Or dtExpiryDate.Year > 2019 Then
                    bDatesValid = False
                End If
            Catch ex As Exception
                bDatesValid = False
            End Try
          
        End If
        Dim sReplenishmentDate As String = tbReplenishmentDate.Text.Trim
        If sReplenishmentDate <> String.Empty Then
            Try
                Dim dtReplenishmentDate As DateTime = DateTime.Parse(sReplenishmentDate)
                If dtReplenishmentDate.Year < 2000 Or dtReplenishmentDate.Year > 2019 Then
                    bDatesValid = False
                End If
            Catch ex As Exception
                bDatesValid = False
            End Try
          
        End If
    End Function
  
    Protected Sub btn_SaveProductChanges_click(ByVal sender As Object, ByVal e As System.EventArgs)
        lblDateError.Visible = False
        If Page.IsValid Then
            If bDatesValid() Then
                Call CheckVisibility()          ' do this in case add or update errors out before completion
                If Not gbDataBound Then         ' this is a kludge because the DataBinder is not always called
                    Call AdjustddlCategory()
                    Call AdjustddlSubCategory()
                    If pnCategoryMode = CATEGORY_MODE_3_CATEGORIES Then
                        Call AdjustddlSubSubCategory()
                    End If
                End If
      
                If Not pbIsAddingCategory Then
                    txtCategory.Text = ddlCategory.SelectedItem.ToString
                End If
                If Not pbIsAddingSubCategory Then
                    txtSubCategory.Text = ddlSubCategory.SelectedItem.ToString
                End If
                If pnCategoryMode = CATEGORY_MODE_3_CATEGORIES Then
                    If Not pbIsAddingSubSubCategory Then
                        tbSubSubCategory.Text = ddlSubSubCategory.SelectedItem.ToString
                    End If
                End If
          
                If pbIsAddingNew Then
                    txtProductCode.Text = txtProductCode.Text.Trim
                    txtProductDate.Text = txtProductDate.Text.Trim
                    If txtProductCode.Text.Length > 0 Then
                        gbExplicitProductPermissions = bSetExplicitProductPermissionsFlag()
                        Call AddNewProduct()
                        If gbExplicitProductPermissions Then
                            Call ShowProductUserProfile()
                        End If
                        Call GetProductNumbers()
                    Else
                        WebMsgBox.Show("Blank product code not allowed.")
                    End If
                Else
                    Call UpdateProduct()
                End If
            Else
                lblDateError.Visible = True
            End If
        End If
    End Sub

    Protected Sub DeleteReservation()
        Dim sSQL As String = "DELETE FROM ProductCodeReservation WHERE CustomerKey = " & Session("CustomerKey") & " AND ProductCode = '" & txtProductCode.Text.Replace("'", "''") & "'"
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand
        Try
            oConn.Open()
            oCmd = New SqlCommand(sSQL, oConn)
            oCmd.ExecuteNonQuery()
        Catch ex As SqlException
            WebMsgBox.Show("Error in DeleteReservation: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub

    Protected Sub btnSetUserProfiles_click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowProductUserProfile()
    End Sub
  
    Protected Sub btn_DeleteProduct_click(ByVal sender As Object, ByVal e As System.EventArgs)
        If CLng(lblProductQuantity.Text) > 0 Then
            WebMsgBox.Show("You cannot delete a product with a positive stock balance. Pick all remaining stock then delete the product.")
        Else
            If System.IO.File.Exists(psProdImageFolder & plProductKey.ToString & ".jpg") Then
                System.IO.File.Delete(psProdImageFolder & plProductKey.ToString & ".jpg")
            End If
            If System.IO.File.Exists(psProdThumbFolder & plProductKey.ToString & ".jpg") Then
                System.IO.File.Delete(psProdThumbFolder & plProductKey.ToString & ".jpg")
            End If
            If System.IO.File.Exists(psProdPDFFolder & plProductKey.ToString & ".pdf") Then
                System.IO.File.Delete(psProdPDFFolder & plProductKey.ToString & ".pdf")
            End If
            Call DeleteProduct()
        End If
    End Sub
  
    Protected Sub btn_GoToProductListPanel_click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call BackToProductListPanel()
    End Sub
  
    Protected Sub BackToProductListPanel()
        Call BindProductGridDispatcher()
        Call ShowMainPanel()
    End Sub
  
    Protected Sub btn_GoBackToProductDetail_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ReturnToProductDetail()
    End Sub

    Protected Sub ReturnToProductDetail()
        Call ShowProductDetail()
    End Sub
  
    Protected Sub dg_ProductList_item_click(ByVal s As Object, ByVal e As DataGridCommandEventArgs)
        If e.CommandSource.CommandName = "Edit" Then
            pbIsAddingNew = False
            pbIsAddingCategory = False
            pbIsAddingSubCategory = False
            pbIsAddingSubSubCategory = False
            lblImageUploadUnavailable.Visible = False
            lblPDFUploadUnavailable.Visible = False
            fuBrowseImageFile.Visible = True
            fuBrowsePDFFile.Visible = True
            btnUploadImage.Visible = True
            btnUploadPDF.Visible = True
            btnAssociatedProducts.Visible = True
            Dim cell_Product As TableCell = e.Item.Cells(0)
            If IsNumeric(cell_Product.Text) Then
                plProductKey = CLng(cell_Product.Text)
            End If
            Call GetProductFromKey()
            btnSetUserProfiles.Visible = True
            btn_DeleteProduct.Visible = True
            lnkbtnNewProductCode.Visible = False
            Call SetHelpStatus()
            Call ShowProductDetail()
        End If
    End Sub
  
    Protected Sub dg_ProductList_Page_Change(ByVal s As Object, ByVal e As DataGridPageChangedEventArgs)
        dg_ProductList.CurrentPageIndex = e.NewPageIndex
        Call BindProductGridDispatcher()
    End Sub
  
    Protected Sub BindProductGridDispatcher()
        If psDisplayMode = DISPLAY_MODE_CATEGORY Then
            Call BindProductGrid(bUseCategories:=True)
        Else
            Call BindProductGrid(bUseCategories:=False)
        End If
    End Sub
      
    Protected Sub BindProductGrid(ByVal bUseCategories As Boolean)
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataTable As New DataTable()
        Dim sProc As String
        If Session("UserType").ToString.ToLower.Contains("owner") Then
            sProc = "spASPNET_Product_GetCustProdsToManageOwned"
        Else
            sProc = "spASPNET_Product_GetCustProdsToManage3"
        End If
        Dim oAdapter As New SqlDataAdapter(sProc, oConn)
        lblProductMessage.Text = ""
        Try
            oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@UserKey", SqlDbType.Int))
          
            oAdapter.SelectCommand.Parameters("@UserKey").Value = Session("UserKey")
          
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
            oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Session("CustomerKey")
          
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SearchCriteria", SqlDbType.NVarChar, 50))
            oAdapter.SelectCommand.Parameters("@SearchCriteria").Value = txtSearchCriteriaAllProducts.Text
          
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@GetByCategory", SqlDbType.Bit))
            oAdapter.SelectCommand.Parameters("@GetByCategory").Value = IIf(bUseCategories, 1, 0)

            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CategoryMode", SqlDbType.Int))
            oAdapter.SelectCommand.Parameters("@CategoryMode").Value = pnCategoryMode
      
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@Category", SqlDbType.NVarChar, 50))
            oAdapter.SelectCommand.Parameters("@Category").Value = psCategory

            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SubCategory", SqlDbType.NVarChar, 50))
            oAdapter.SelectCommand.Parameters("@SubCategory").Value = psSubCategory

            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SubCategory2", SqlDbType.NVarChar, 50))
            oAdapter.SelectCommand.Parameters("@SubCategory2").Value = psSubSubCategory

            oAdapter.Fill(oDataTable)
            If oDataTable.Rows.Count > 0 Then
                dg_ProductList.DataSource = oDataTable                ' when I navigated to last page in product index, then created a new product, error "invalid current page index". Current page index was 1, Page count was 4, start is 0
                dg_ProductList.DataBind()
                dg_ProductList.Visible = True
                'btn_RefreshProductGrid.Visible= True
                If oDataTable.Rows.Count > 8 Then
                    dg_ProductList.PagerStyle.Visible = True
                Else
                    dg_ProductList.PagerStyle.Visible = False
                End If
            Else
                lblProductMessage.Text = "No products found"
                dg_ProductList.Visible = False
                'btn_RefreshProductGrid.Visible= False
            End If
        Catch ex As SqlException
            lblError.Text = ""
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
        If txtSearchCriteriaAllProducts.Text <> String.Empty Then
            txtSearchCriteriaAllProducts.Focus()
        End If
    End Sub
  
    Protected Sub SaveImage()
        Dim fi As New System.IO.FileInfo(fuBrowseImageFile.PostedFile.FileName)
        If fi.Extension.ToLower() = ".jpg" Then
            Try
                Dim sTempPath As String = psProdImageFolder & plProductKey.ToString & ".upload.jpg"
                If System.IO.File.Exists(sTempPath) Then
                    System.IO.File.Delete(sTempPath)
                End If
                fuBrowseImageFile.PostedFile.SaveAs(sTempPath)
                Call SaveResizedImage(fi, 600, 600) 'new
                Call MakeThumbNail(fi, 60, 60)
                Call SetImageAttributes()
                hlnk_DetailThumb.ImageUrl = psVirtualThumbFolder & plProductKey.ToString & ".jpg?" & Now.ToString
                hlnk_DetailThumb.NavigateUrl = psVirtualJPGFolder & plProductKey.ToString & ".jpg"
                imgbtnDeleteImage.Visible = True
                'now delete the original file
                If System.IO.File.Exists(sTempPath) Then
                    System.IO.File.Delete(sTempPath)
                End If
            Catch ex As Exception
                Response.Write(ex.ToString)
                hlnk_DetailThumb.ImageUrl = psVirtualThumbFolder & "blank_thumb.jpg"
                hlnk_DetailThumb.NavigateUrl = psVirtualJPGFolder & "blank_image.jpg"
            End Try
        Else
            WebMsgBox.Show("Only files with a .JPG extension can be uploaded")
        End If
    End Sub
  
    Protected Sub SavePDF()
        Dim fi As New System.IO.FileInfo(fuBrowsePDFFile.PostedFile.FileName)
        If fi.Extension.ToLower() = ".pdf" Then
            Try
                Dim sTempPath As String = psProdPDFFolder & plProductKey.ToString & ".pdf"
                If System.IO.File.Exists(sTempPath) Then
                    System.IO.File.Delete(sTempPath)
                End If
                fuBrowsePDFFile.PostedFile.SaveAs(psProdPDFFolder & plProductKey.ToString & ".pdf")
                hlnk_PDF.ImageUrl = psVirtualPDFFolder & "pdf_logo.gif"
                hlnk_PDF.NavigateUrl = psVirtualPDFFolder & plProductKey.ToString & ".pdf"
                hlnk_PDF.Target = "_blank"
                imgbtnDeletePDF.Visible = True
                Call SetPDFAttribute()
            Catch ex As Exception
                Response.Write(ex.ToString)
            End Try
        Else
            WebMsgBox.Show("Only files with a .PDF extension can be uploaded")
        End If
    End Sub
  
    Protected Sub MakeThumbNail(ByVal fi As System.IO.FileInfo, ByVal MaxWidth As Double, ByVal MaxHeight As Double)
        Dim OriginalImg As System.Drawing.Image = System.Drawing.Image.FromFile(psProdImageFolder & plProductKey.ToString & ".upload.jpg")
        Dim TheSize As New System.Drawing.Size(OriginalImg.Width, OriginalImg.Height)
  
        Dim sizer As Double = 1
  
        If (MaxWidth > -1 And TheSize.Width > MaxWidth) Or (MaxHeight > -1 And TheSize.Height > MaxHeight) Then
            If MaxWidth > -1 And TheSize.Width > MaxWidth Then
                sizer = MaxWidth / TheSize.Width
                TheSize.Width = Convert.ToInt32(TheSize.Width * sizer)
                TheSize.Height = Convert.ToInt32(TheSize.Height * sizer)
            End If
            If MaxHeight > -1 And TheSize.Height > MaxHeight Then
                sizer = MaxHeight / TheSize.Height
                TheSize.Width = Convert.ToInt32(TheSize.Width * sizer)
                TheSize.Height = Convert.ToInt32(TheSize.Height * sizer)
            End If
        Else
            TheSize.Width = OriginalImg.Width  'Don't try and reduce an image that's already smaller than target size
            TheSize.Height = OriginalImg.Height
        End If
  
        Dim SavePath As String = psProdThumbFolder & plProductKey.ToString & ".jpg" '& F.Name
  
        Dim NewImg As New System.Drawing.Bitmap(OriginalImg, TheSize)
        OriginalImg.Dispose()
  
        If System.IO.File.Exists(SavePath) Then
            System.IO.File.Delete(SavePath)
        End If
        NewImg.Save(SavePath, System.Drawing.Imaging.ImageFormat.Jpeg)
        NewImg.Dispose()
    End Sub
  
    Protected Sub SaveResizedImage(ByVal fi As System.IO.FileInfo, ByVal MaxWidth As Double, ByVal MaxHeight As Double)
        Dim OriginalImg As System.Drawing.Image = System.Drawing.Image.FromFile(psProdImageFolder & plProductKey.ToString & ".upload.jpg")
        Dim TheSize As New System.Drawing.Size(OriginalImg.Width, OriginalImg.Height)
  
        Dim sizer As Double = 1
        Dim sSavePath As String = psProdImageFolder & plProductKey.ToString & ".jpg"
        Try

            If (MaxWidth > -1 And TheSize.Width > MaxWidth) Or (MaxHeight > -1 And TheSize.Height > MaxHeight) Then
                If MaxWidth > -1 And TheSize.Width > MaxWidth Then
                    sizer = MaxWidth / TheSize.Width
                    TheSize.Width = Convert.ToInt32(TheSize.Width * sizer)
                    TheSize.Height = Convert.ToInt32(TheSize.Height * sizer)
                End If
                If MaxHeight > -1 And TheSize.Height > MaxHeight Then
                    sizer = MaxHeight / TheSize.Height
                    TheSize.Width = Convert.ToInt32(TheSize.Width * sizer)
                    TheSize.Height = Convert.ToInt32(TheSize.Height * sizer)
                End If
            Else
                'Don't try and reduce an image that's already smaller than our target size
                TheSize.Width = OriginalImg.Width
                TheSize.Height = OriginalImg.Height
            End If
  
            Dim NewImg As New System.Drawing.Bitmap(OriginalImg, TheSize)
            OriginalImg.Dispose()
  
            If System.IO.File.Exists(sSavePath) Then
                System.IO.File.Delete(sSavePath)
            End If
            NewImg.Save(sSavePath, System.Drawing.Imaging.ImageFormat.Jpeg)
      
            NewImg.Dispose()

        Catch ex As Exception
            WebMsgBox.Show("Unable to resize image. The system may be too busy to allocated the required amount of memory.  Try resizing your image to make it smaller (maximum dimension 600 pixels). ")
        End Try
    End Sub
  
    Protected Sub SetImageAttributes()
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Product_SetImageAttributes", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Dim paramUserProfileKey As SqlParameter = New SqlParameter("@UserProfileKey", SqlDbType.Int, 4)
        paramUserProfileKey.Value = Session("UserKey")
        oCmd.Parameters.Add(paramUserProfileKey)
        Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
        paramCustomerKey.Value = Session("CustomerKey")
        oCmd.Parameters.Add(paramCustomerKey)
        Dim paramProductKey As SqlParameter = New SqlParameter("@ProductKey", SqlDbType.Int, 4)
        paramProductKey.Value = plProductKey
        oCmd.Parameters.Add(paramProductKey)
        Dim paramThumbNailImage As SqlParameter = New SqlParameter("@ThumbNailImage", SqlDbType.NVarChar, 20)
        paramThumbNailImage.Value = plProductKey.ToString & ".jpg"
        oCmd.Parameters.Add(paramThumbNailImage)
        Dim paramOriginalImage As SqlParameter = New SqlParameter("@OriginalImage", SqlDbType.NVarChar, 20)
        paramOriginalImage.Value = plProductKey.ToString & ".jpg"
        oCmd.Parameters.Add(paramOriginalImage)
        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As SqlException
            lblError.Text = "Error in SetImageAttributes: " & ex.ToString
        Finally
            oConn.Close()
        End Try
    End Sub
  
    Protected Sub ResetImageAttributes()
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Product_SetImageAttributes", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Dim paramUserProfileKey As SqlParameter = New SqlParameter("@UserProfileKey", SqlDbType.Int, 4)
        paramUserProfileKey.Value = Session("UserKey")
        oCmd.Parameters.Add(paramUserProfileKey)
        Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
        paramCustomerKey.Value = Session("CustomerKey")
        oCmd.Parameters.Add(paramCustomerKey)
        Dim paramProductKey As SqlParameter = New SqlParameter("@ProductKey", SqlDbType.Int, 4)
        paramProductKey.Value = plProductKey
        oCmd.Parameters.Add(paramProductKey)
        Dim paramThumbNailImage As SqlParameter = New SqlParameter("@ThumbNailImage", SqlDbType.NVarChar, 20)
        paramThumbNailImage.Value = "blank_thumb.jpg"
        oCmd.Parameters.Add(paramThumbNailImage)
        Dim paramOriginalImage As SqlParameter = New SqlParameter("@OriginalImage", SqlDbType.NVarChar, 20)
        paramOriginalImage.Value = "blank_image.jpg"
        oCmd.Parameters.Add(paramOriginalImage)
        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As SqlException
            lblError.Text = "Error in ResetImageAttributes: " & ex.Message
        Finally
            oConn.Close()
        End Try
    End Sub
  
    Protected Sub SetPDFAttribute()
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Product_SetPDFAttribute", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Dim paramUserProfileKey As SqlParameter = New SqlParameter("@UserProfileKey", SqlDbType.Int, 4)
        paramUserProfileKey.Value = Session("UserKey")
        oCmd.Parameters.Add(paramUserProfileKey)
        Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
        paramCustomerKey.Value = Session("CustomerKey")
        oCmd.Parameters.Add(paramCustomerKey)
        Dim paramProductKey As SqlParameter = New SqlParameter("@ProductKey", SqlDbType.Int, 4)
        paramProductKey.Value = plProductKey
        oCmd.Parameters.Add(paramProductKey)
        Dim paramThumbNailImage As SqlParameter = New SqlParameter("@PDFFileName", SqlDbType.NVarChar, 60)
        paramThumbNailImage.Value = plProductKey.ToString & ".pdf"
        oCmd.Parameters.Add(paramThumbNailImage)
        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As SqlException
            lblError.Text = "Error in SetPDFAttribute " & ex.Message
        Finally
            oConn.Close()
        End Try
    End Sub

    Protected Sub ResetPDFAttribute()
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Product_SetPDFAttribute", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Dim paramUserProfileKey As SqlParameter = New SqlParameter("@UserProfileKey", SqlDbType.Int, 4)
        paramUserProfileKey.Value = Session("UserKey")
        oCmd.Parameters.Add(paramUserProfileKey)
        Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
        paramCustomerKey.Value = Session("CustomerKey")
        oCmd.Parameters.Add(paramCustomerKey)
        Dim paramProductKey As SqlParameter = New SqlParameter("@ProductKey", SqlDbType.Int, 4)
        paramProductKey.Value = plProductKey
        oCmd.Parameters.Add(paramProductKey)
        Dim paramThumbNailImage As SqlParameter = New SqlParameter("@PDFFileName", SqlDbType.NVarChar, 60)
        paramThumbNailImage.Value = "blank_pdf.jpg"
        oCmd.Parameters.Add(paramThumbNailImage)
        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As SqlException
            lblError.Text = "Error in ResetPDFAttribute: " & ex.Message
        Finally
            oConn.Close()
        End Try
    End Sub
  
  
    Protected Sub GetProductFromKey()
        Dim oDataReader As SqlDataReader
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As New SqlCommand("spASPNET_Product_GetProductFromKey8", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Dim oParam As SqlParameter = oCmd.Parameters.Add("@ProductKey", SqlDbType.Int, 4)
        oParam.Value = plProductKey
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            oDataReader.Read()
            If IsDBNull(oDataReader("ProductCode")) Then
                txtProductCode.Text = ""
            Else
                txtProductCode.Text = oDataReader("ProductCode")
            End If
            If IsDBNull(oDataReader("ProductDate")) Then
                txtProductDate.Text = ""
            Else
                txtProductDate.Text = oDataReader("ProductDate")
            End If
            If IsDBNull(oDataReader("ProductDescription")) Then
                txtDescription.Text = ""
            Else
                txtDescription.Text = oDataReader("ProductDescription")
            End If
            If IsDBNull(oDataReader("LanguageId")) Then
                txtLanguage.Text = ""
            Else
                txtLanguage.Text = oDataReader("LanguageId")
            End If
            If IsDBNull(oDataReader("ProductDepartmentId")) Then
                txtDepartment.Text = ""
            Else
                txtDepartment.Text = oDataReader("ProductDepartmentId")
            End If
            If IsDBNull(oDataReader("ProductCategory")) Then
                hidCategory.Value = ""
            Else
                hidCategory.Value = oDataReader("ProductCategory")
            End If
            If IsDBNull(oDataReader("SubCategory")) Then
                hidSubCategory.Value = ""
            Else
                hidSubCategory.Value = oDataReader("SubCategory")
            End If
            If IsDBNull(oDataReader("SubCategory2")) Then
                hidSubSubCategory.Value = ""
            Else
                hidSubSubCategory.Value = oDataReader("SubCategory2")
            End If
            If IsDBNull(oDataReader("Notes")) Then
                txtNotes.Text = ""
            Else
                txtNotes.Text = oDataReader("Notes")
            End If
            If IsDBNull(oDataReader("ItemsPerBox")) Then
                txtItemsPerBox.Text = ""
            Else
                txtItemsPerBox.Text = oDataReader("ItemsPerBox")
            End If
            If IsDBNull(oDataReader("MinimumStockLevel")) Then
                txtMinStockLevel.Text = ""
            Else
                txtMinStockLevel.Text = oDataReader("MinimumStockLevel")
            End If
            If IsDBNull(oDataReader("UnitValue")) Then
                txtUnitValue.Text = ""
            Else
                txtUnitValue.Text = Format(oDataReader("UnitValue"), "#,##0.00")
            End If
            If IsDBNull(oDataReader("UnitWeightGrams")) Then
                txtUnitWeight.Text = ""
            Else
                txtUnitWeight.Text = Format(oDataReader("UnitWeightGrams"), "#,##0")
            End If
            ' expiry & replenishment date handling note: the extra < 1990 test is to cater for possible 1/1/1900 dates (equivalent to NULL) in the database
            If IsDBNull(oDataReader("ExpiryDate")) Then
                tbExpiryDate.Text = ""
            Else
                If Year(oDataReader("ExpiryDate")) < 1990 Then
                    tbExpiryDate.Text = ""
                Else
                    tbExpiryDate.Text = Format(oDataReader("ExpiryDate"), "dd-MMM-yyyy")
                End If
            End If
            If IsDBNull(oDataReader("ReplenishmentDate")) Then
                tbReplenishmentDate.Text = ""
            Else
                If Year(oDataReader("ReplenishmentDate")) < 1990 Then
                    tbReplenishmentDate.Text = ""
                Else
                    tbReplenishmentDate.Text = Format(oDataReader("ReplenishmentDate"), "dd-MMM-yyyy")
                End If
            End If
            If IsDBNull(oDataReader("Misc1")) Then
                txtMisc1.Text = ""
            Else
                txtMisc1.Text = oDataReader("Misc1")
            End If
            If IsDBNull(oDataReader("Misc2")) Then
                txtMisc2.Text = ""
            Else
                txtMisc2.Text = oDataReader("Misc2")
            End If
          
            If IsDBNull(oDataReader("SerialNumbersFlag")) Then
                chkProspectusNumbers.Checked = False
            ElseIf oDataReader("SerialNumbersFlag") = "Y" Then
                chkProspectusNumbers.Checked = True
            ElseIf oDataReader("SerialNumbersFlag") = "N" Then
                chkProspectusNumbers.Checked = False
            End If
            If IsDBNull(oDataReader("ArchiveFlag")) Then
                chkArchivedFlag.Checked = False
            ElseIf oDataReader("ArchiveFlag") = "Y" Then
                chkArchivedFlag.Checked = True
            ElseIf oDataReader("ArchiveFlag") = "N" Then
                chkArchivedFlag.Checked = False
            End If
            hlnk_DetailThumb.ImageUrl = psVirtualThumbFolder & oDataReader("ThumbNailImage")
            hlnk_DetailThumb.NavigateUrl = psVirtualJPGFolder & oDataReader("OriginalImage")
            If oDataReader("ThumbNailImage") = "blank_thumb.jpg" Then
                imgbtnDeleteImage.Visible = False
            Else
                imgbtnDeleteImage.Visible = True
            End If
          
            If oDataReader("PDFFileName") = "blank_pdf.jpg" Then
                hlnk_PDF.ImageUrl = psVirtualPDFFolder & "blank_pdf_thumb.jpg"
                hlnk_PDF.NavigateUrl = psVirtualPDFFolder & "blank_pdf.jpg"
                imgbtnDeletePDF.Visible = False
            Else
                hlnk_PDF.ImageUrl = psVirtualPDFFolder & "pdf_logo.gif"
                hlnk_PDF.NavigateUrl = psVirtualPDFFolder & oDataReader("PDFFileName")
                imgbtnDeletePDF.Visible = True
            End If
  
            If oDataReader("WebsiteAdRotatorFlag") = True Then
                chkAdRotator.Checked = True
            Else
                chkAdRotator.Checked = False
            End If
            If IsDBNull(oDataReader("AdRotatorText")) Then
                txtAdRotatorText.Text = ""
            Else
                txtAdRotatorText.Text = oDataReader("AdRotatorText")
            End If
            If oDataReader("ViewOnWebForm") = True Then
                chkViewOnWebForm.Checked = True
            Else
                chkViewOnWebForm.Checked = False
            End If
            If IsDBNull(oDataReader("InactivityAlertDays")) Then
                tbInactivityAlertDays.Text = 0
            Else
                tbInactivityAlertDays.Text = oDataReader("InactivityAlertDays")
            End If
            lblLegendLanguage.Text = "Language:"
            lblProductQuantity.Text = Format(oDataReader("Quantity"), "#,##0")
            
            oDataReader.Close()
        Catch ex As SqlException
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
    End Sub
  
    Protected Sub UpdateProduct()
        lblError.Text = ""
        Dim oConn As New SqlConnection(gsConn)
        Dim oTrans As SqlTransaction
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Product_FullUpdate9", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
  
        Dim paramUserKey As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int, 4)
        paramUserKey.Value = CLng(Session("UserKey"))
        oCmd.Parameters.Add(paramUserKey)
        Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
        paramCustomerKey.Value = CLng(Session("CustomerKey"))
        oCmd.Parameters.Add(paramCustomerKey)
        Dim paramProductKey As SqlParameter = New SqlParameter("@LogisticProductKey", SqlDbType.Int, 4)
        paramProductKey.Value = plProductKey
        oCmd.Parameters.Add(paramProductKey)
  
        Dim paramMinimumStockLevel As SqlParameter = New SqlParameter("@MinimumStockLevel", SqlDbType.Int, 4)
        If IsNumeric(txtMinStockLevel.Text) Then
            paramMinimumStockLevel.Value = CLng(txtMinStockLevel.Text)
        Else
            paramMinimumStockLevel.Value = 0
        End If
        oCmd.Parameters.Add(paramMinimumStockLevel)
        Dim paramDescription As SqlParameter = New SqlParameter("@ProductDescription", SqlDbType.NVarChar, 300)
        paramDescription.Value = txtDescription.Text
        oCmd.Parameters.Add(paramDescription)
        Dim paramItemsPerBox As SqlParameter = New SqlParameter("@ItemsPerBox", SqlDbType.Int, 4)
        If IsNumeric(txtItemsPerBox.Text) Then
            paramItemsPerBox.Value = CLng(txtItemsPerBox.Text)
        Else
            paramItemsPerBox.Value = 0
        End If
        oCmd.Parameters.Add(paramItemsPerBox)
        Dim paramCategory As SqlParameter = New SqlParameter("@ProductCategory", SqlDbType.NVarChar, 50)
        paramCategory.Value = txtCategory.Text
        oCmd.Parameters.Add(paramCategory)
        Dim paramSubCategory As SqlParameter = New SqlParameter("@SubCategory", SqlDbType.NVarChar, 50)
        paramSubCategory.Value = txtSubCategory.Text
        oCmd.Parameters.Add(paramSubCategory)
        Dim paramSubCategory2 As SqlParameter = New SqlParameter("@SubCategory2", SqlDbType.NVarChar, 50)
        paramSubCategory2.Value = tbSubSubCategory.Text
        oCmd.Parameters.Add(paramSubCategory2)
        Dim paramUnitValue As SqlParameter = New SqlParameter("@UnitValue", SqlDbType.Money, 8)
        If IsNumeric(txtUnitValue.Text) Then
            If CDec(txtUnitValue.Text) > 0 Then
                paramUnitValue.Value = CDec(txtUnitValue.Text)
            Else
                paramUnitValue.Value = 0
            End If
        Else
            paramUnitValue.Value = 0
        End If
        oCmd.Parameters.Add(paramUnitValue)
        Dim paramUnitValue2 As SqlParameter = New SqlParameter("@UnitValue2", SqlDbType.Money, 8)
        paramUnitValue2.Value = 0
        oCmd.Parameters.Add(paramUnitValue2)
        
        Dim paramLanguage As SqlParameter = New SqlParameter("@LanguageId", SqlDbType.NVarChar, 20)
        paramLanguage.Value = txtLanguage.Text
        oCmd.Parameters.Add(paramLanguage)
        
        Dim paramDepartment As SqlParameter = New SqlParameter("@ProductDepartmentId", SqlDbType.NVarChar, 20)
        paramDepartment.Value = txtDepartment.Text
        oCmd.Parameters.Add(paramDepartment)
        
        Dim paramWeight As SqlParameter = New SqlParameter("@UnitWeightGrams", SqlDbType.Int, 4)
        If IsNumeric(txtUnitWeight.Text) Then
            paramWeight.Value = CLng(txtUnitWeight.Text)
        Else
            paramWeight.Value = 0
        End If
        oCmd.Parameters.Add(paramWeight)
      
        Dim paramStockOwnedByKey As SqlParameter = New SqlParameter("@StockOwnedByKey", SqlDbType.Int, 4)
        paramStockOwnedByKey.Value = 0
        oCmd.Parameters.Add(paramStockOwnedByKey)
        Dim paramMisc1 As SqlParameter = New SqlParameter("@Misc1", SqlDbType.NVarChar, 50)
        paramMisc1.Value = txtMisc1.Text
        oCmd.Parameters.Add(paramMisc1)
        Dim paramMisc2 As SqlParameter = New SqlParameter("@Misc2", SqlDbType.NVarChar, 50)
        paramMisc2.Value = txtMisc2.Text
        oCmd.Parameters.Add(paramMisc2)
        Dim paramArchive As SqlParameter = New SqlParameter("@ArchiveFlag", SqlDbType.NVarChar, 1)
        If chkArchivedFlag.Checked Then
            paramArchive.Value = "Y"
        Else
            paramArchive.Value = "N"
        End If
        oCmd.Parameters.Add(paramArchive)
      
        Dim paramStatus As SqlParameter = New SqlParameter("@Status", SqlDbType.TinyInt)
        If IsNumeric(ddlStatus.SelectedValue) Then
            paramStatus.Value = ddlStatus.SelectedValue
        Else
            paramStatus.Value = 0
        End If
        oCmd.Parameters.Add(paramStatus)

        Dim paramExpiryDate As SqlParameter = New SqlParameter("@ExpiryDate", SqlDbType.SmallDateTime)
        Dim sExpiryDate As String
        sExpiryDate = tbExpiryDate.Text.Trim()
        If sExpiryDate <> "" Then
            Try
                sExpiryDate = DateTime.Parse(sExpiryDate)
            Catch ex As Exception
                lblError.Text = "ERROR: Invalid Expiry Date"
                Exit Sub
            End Try
        End If
        If sExpiryDate = "" Then
            paramExpiryDate.Value = Nothing
        Else
            paramExpiryDate.Value = sExpiryDate
        End If
        oCmd.Parameters.Add(paramExpiryDate)

        Dim paramReplenishmentDate As SqlParameter = New SqlParameter("@ReplenishmentDate", SqlDbType.SmallDateTime)
        Dim sReplenishmentDate As String
        sReplenishmentDate = tbReplenishmentDate.Text.Trim
        If sReplenishmentDate <> "" Then
            Try
                sReplenishmentDate = DateTime.Parse(sReplenishmentDate)
            Catch ex As Exception
                lblError.Text = "ERROR: Invalid Renewal / Review Date"
                Exit Sub
            End Try
        End If
        If sReplenishmentDate = "" Then
            paramReplenishmentDate.Value = Nothing
        Else
            paramReplenishmentDate.Value = sReplenishmentDate
        End If
        oCmd.Parameters.Add(paramReplenishmentDate)

        Dim paramSerialNumbers As SqlParameter = New SqlParameter("@SerialNumbersFlag", SqlDbType.NVarChar, 1)
        If chkProspectusNumbers.Checked Then
            paramSerialNumbers.Value = "Y"
        Else
            paramSerialNumbers.Value = "N"
        End If
        oCmd.Parameters.Add(paramSerialNumbers)
        Dim paramAdRotatorText As SqlParameter = New SqlParameter("@AdRotatorText", SqlDbType.NVarChar, 120)
        paramAdRotatorText.Value = txtAdRotatorText.Text
        oCmd.Parameters.Add(paramAdRotatorText)

        Dim paramWebsiteAdRotatorFlag As SqlParameter = New SqlParameter("@WebsiteAdRotatorFlag", SqlDbType.Bit)
        If chkAdRotator.Checked Then
            paramWebsiteAdRotatorFlag.Value = 1
        Else
            paramWebsiteAdRotatorFlag.Value = 0
        End If
        oCmd.Parameters.Add(paramWebsiteAdRotatorFlag)

        Dim paramNotes As SqlParameter = New SqlParameter("@Notes", SqlDbType.NVarChar, 1000)
        paramNotes.Value = txtNotes.Text
        oCmd.Parameters.Add(paramNotes)

        Dim paramViewOnWeb As SqlParameter = New SqlParameter("@ViewOnWebForm", SqlDbType.Bit)
        If chkViewOnWebForm.Checked Then
            paramViewOnWeb.Value = 1
        Else
            paramViewOnWeb.Value = 0
        End If
        oCmd.Parameters.Add(paramViewOnWeb)
      
        Dim paramRotationProductKey As SqlParameter = New SqlParameter("@RotationProductKey", SqlDbType.Int, 4)
        paramRotationProductKey.Value = System.Data.SqlTypes.SqlInt32.Null
        oCmd.Parameters.Add(paramRotationProductKey)

        Dim paramInactivityAlertDays As SqlParameter = New SqlParameter("@InactivityAlertDays", SqlDbType.Int)
        If IsNumeric(tbInactivityAlertDays.Text) Then
            paramInactivityAlertDays.Value = CLng(tbInactivityAlertDays.Text)
        Else
            paramInactivityAlertDays.Value = 0
        End If
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
        
        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
            Call BindProductGridDispatcher()
            Call ShowMainPanel()
        Catch ex As SqlException
            lblError.Text = "Error in UpdateProduct: " & ex.Message
        Finally
            oConn.Close()
        End Try
    End Sub
  
    Protected Sub AddNewProduct()
        lblError.Text = ""
        Dim oConn As New SqlConnection(gsConn)
        Dim oTrans As SqlTransaction
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Product_AddWithAccessControl8", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
  
        Dim paramUserKey As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int, 4)
        paramUserKey.Value = CLng(Session("UserKey"))
        oCmd.Parameters.Add(paramUserKey)
        Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
        paramCustomerKey.Value = CLng(Session("CustomerKey"))
        oCmd.Parameters.Add(paramCustomerKey)
  
        Dim paramProductCode As SqlParameter = New SqlParameter("@ProductCode", SqlDbType.NVarChar, 25)
        paramProductCode.Value = txtProductCode.Text
        oCmd.Parameters.Add(paramProductCode)
        Dim paramProductDate As SqlParameter = New SqlParameter("@ProductDate", SqlDbType.NVarChar, 10)
        paramProductDate.Value = txtProductDate.Text
        oCmd.Parameters.Add(paramProductDate)
  
        Dim paramMinimumStockLevel As SqlParameter = New SqlParameter("@MinimumStockLevel", SqlDbType.Int, 4)
        If IsNumeric(txtMinStockLevel.Text) Then
            paramMinimumStockLevel.Value = CLng(txtMinStockLevel.Text)
        Else
            paramMinimumStockLevel.Value = 0
        End If
        oCmd.Parameters.Add(paramMinimumStockLevel)
        Dim paramDescription As SqlParameter = New SqlParameter("@ProductDescription", SqlDbType.NVarChar, 300)
        paramDescription.Value = txtDescription.Text
        oCmd.Parameters.Add(paramDescription)
        Dim paramItemsPerBox As SqlParameter = New SqlParameter("@ItemsPerBox", SqlDbType.Int, 4)
        If IsNumeric(txtItemsPerBox.Text) Then
            paramItemsPerBox.Value = CLng(txtItemsPerBox.Text)
        Else
            paramItemsPerBox.Value = 0
        End If
        oCmd.Parameters.Add(paramItemsPerBox)
        Dim paramCategory As SqlParameter = New SqlParameter("@ProductCategory", SqlDbType.NVarChar, 50)
        paramCategory.Value = txtCategory.Text
        oCmd.Parameters.Add(paramCategory)
        Dim paramSubCategory As SqlParameter = New SqlParameter("@SubCategory", SqlDbType.NVarChar, 50)
        paramSubCategory.Value = txtSubCategory.Text
        oCmd.Parameters.Add(paramSubCategory)
        Dim paramSubCategory2 As SqlParameter = New SqlParameter("@SubCategory2", SqlDbType.NVarChar, 50)
        paramSubCategory2.Value = tbSubSubCategory.Text
        oCmd.Parameters.Add(paramSubCategory2)
        Dim paramUnitValue As SqlParameter = New SqlParameter("@UnitValue", SqlDbType.Money, 8)
        If IsNumeric(txtUnitValue.Text) Then
            If CDec(txtUnitValue.Text) > 0 Then
                paramUnitValue.Value = CDec(txtUnitValue.Text)
            Else
                paramUnitValue.Value = 0
            End If
        Else
            paramUnitValue.Value = 0
        End If
        oCmd.Parameters.Add(paramUnitValue)
        Dim paramUnitValue2 As SqlParameter = New SqlParameter("@UnitValue2", SqlDbType.Money, 8)
        paramUnitValue2.Value = 0
        oCmd.Parameters.Add(paramUnitValue2)

        Dim paramLanguage As SqlParameter = New SqlParameter("@LanguageId", SqlDbType.NVarChar, 20)
        paramLanguage.Value = txtLanguage.Text
        oCmd.Parameters.Add(paramLanguage)

        Dim paramDepartment As SqlParameter = New SqlParameter("@ProductDepartmentId", SqlDbType.NVarChar, 20)
        paramDepartment.Value = txtDepartment.Text
        oCmd.Parameters.Add(paramDepartment)
        Dim paramWeight As SqlParameter = New SqlParameter("@UnitWeightGrams", SqlDbType.Int, 4)
        If IsNumeric(txtUnitWeight.Text) Then
            paramWeight.Value = CLng(txtUnitWeight.Text)
        Else
            paramWeight.Value = 0
        End If
        oCmd.Parameters.Add(paramWeight)
        Dim paramStockOwnedByKey As SqlParameter = New SqlParameter("@StockOwnedByKey", SqlDbType.Int, 4)
        paramStockOwnedByKey.Value = 0
        oCmd.Parameters.Add(paramStockOwnedByKey)
        Dim paramMisc1 As SqlParameter = New SqlParameter("@Misc1", SqlDbType.NVarChar, 50)
        paramMisc1.Value = txtMisc1.Text
        oCmd.Parameters.Add(paramMisc1)
        Dim paramMisc2 As SqlParameter = New SqlParameter("@Misc2", SqlDbType.NVarChar, 50)
        paramMisc2.Value = txtMisc2.Text
        oCmd.Parameters.Add(paramMisc2)
        Dim paramArchive As SqlParameter = New SqlParameter("@ArchiveFlag", SqlDbType.NVarChar, 1)
        If chkArchivedFlag.Checked Then
            paramArchive.Value = "Y"
        Else
            paramArchive.Value = "N"
        End If
        oCmd.Parameters.Add(paramArchive)
      
        Dim paramStatus As SqlParameter = New SqlParameter("@Status", SqlDbType.TinyInt)
        If IsNumeric(ddlStatus.SelectedValue) Then
            paramStatus.Value = ddlStatus.SelectedValue
        Else
            paramStatus.Value = 0
        End If
        oCmd.Parameters.Add(paramStatus)

        Dim paramExpiryDate As SqlParameter = New SqlParameter("@ExpiryDate", SqlDbType.SmallDateTime)
        Dim sExpiryDate As String
        sExpiryDate = tbExpiryDate.Text.Trim()
        If sExpiryDate <> "" Then
            Try
                sExpiryDate = DateTime.Parse(sExpiryDate)
            Catch ex As Exception
                lblError.Text = "ERROR: Invalid Expiry Date"
                Exit Sub
            End Try
        End If
        If sExpiryDate = "" Then
            paramExpiryDate.Value = Nothing
        Else
            paramExpiryDate.Value = sExpiryDate
        End If
        oCmd.Parameters.Add(paramExpiryDate)

        Dim paramReplenishmentDate As SqlParameter = New SqlParameter("@ReplenishmentDate", SqlDbType.SmallDateTime)
        Dim sReplenishmentDate As String
        sReplenishmentDate = tbReplenishmentDate.Text.Trim
        If sReplenishmentDate <> "" Then
            Try
                sReplenishmentDate = DateTime.Parse(sReplenishmentDate)
            Catch ex As Exception
                lblError.Text = "ERROR: Invalid Renewal / Review Date"
                Exit Sub
            End Try
        End If
        If sReplenishmentDate = "" Then
            paramReplenishmentDate.Value = Nothing
        Else
            paramReplenishmentDate.Value = sReplenishmentDate
        End If
        oCmd.Parameters.Add(paramReplenishmentDate)
      
        Dim paramSerialNumbers As SqlParameter = New SqlParameter("@SerialNumbersFlag", SqlDbType.NVarChar, 1)
        If chkProspectusNumbers.Checked Then
            paramSerialNumbers.Value = "Y"
        Else
            paramSerialNumbers.Value = "N"
        End If
        oCmd.Parameters.Add(paramSerialNumbers)
        Dim paramAdRotatorText As SqlParameter = New SqlParameter("@AdRotatorText", SqlDbType.NVarChar, 120)
        paramAdRotatorText.Value = txtAdRotatorText.Text
        oCmd.Parameters.Add(paramAdRotatorText)
        Dim paramWebsiteAdRotatorFlag As SqlParameter = New SqlParameter("@WebsiteAdRotatorFlag", SqlDbType.Bit)
        If chkAdRotator.Checked Then
            paramWebsiteAdRotatorFlag.Value = 1
        Else
            paramWebsiteAdRotatorFlag.Value = 0
        End If
        oCmd.Parameters.Add(paramWebsiteAdRotatorFlag)
        Dim paramNotes As SqlParameter = New SqlParameter("@Notes", SqlDbType.NVarChar, 1000)
        paramNotes.Value = txtNotes.Text
        oCmd.Parameters.Add(paramNotes)

        Dim paramViewOnWebForm As SqlParameter = New SqlParameter("@ViewOnWebForm", SqlDbType.Bit)
        If chkViewOnWebForm.Checked Then
            paramViewOnWebForm.Value = 1
        Else
            paramViewOnWebForm.Value = 0
        End If
        oCmd.Parameters.Add(paramViewOnWebForm)
  
        Dim paramDefaultAccessFlag As SqlParameter = New SqlParameter("@DefaultAccessFlag", SqlDbType.Bit)
        paramDefaultAccessFlag.Value = Not gbExplicitProductPermissions
        oCmd.Parameters.Add(paramDefaultAccessFlag)

        Dim paramRotationProductKey As SqlParameter = New SqlParameter("@RotationProductKey", SqlDbType.Int, 4)
        paramRotationProductKey.Value = System.Data.SqlTypes.SqlInt32.Null
        oCmd.Parameters.Add(paramRotationProductKey)

        Dim paramInactivityAlertDays As SqlParameter = New SqlParameter("@InactivityAlertDays", SqlDbType.Int, 4)
        If IsNumeric(tbInactivityAlertDays.Text) Then
            paramInactivityAlertDays.Value = CLng(tbInactivityAlertDays.Text)
        Else
            paramInactivityAlertDays.Value = 0
        End If
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
            plProductKey = CLng(oCmd.Parameters("@ProductKey").Value)
            'txtSearchCriteriaAllProducts.Text = txtProductCode.Text
            Call BindProductGridDispatcher()
            Call ShowMainPanel()
        Catch ex As SqlException
            If ex.Number = 2627 Then
                lblError.Text = "ERROR: A record already exists with the same product CODE and DATE combination"
            Else
                lblError.Text = ex.ToString
            End If
        Finally
            oConn.Close()
        End Try
    End Sub
  
    Protected Function bSetExplicitProductPermissionsFlag() As Boolean
        Dim oDataTable As New DataTable()
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter("spASPNET_Customer_ExplicitProductPermissions_GetFlag", oConn)

        Try
            oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
            oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Session("CustomerKey")
            oAdapter.Fill(oDataTable)
        Catch ex As SqlException
            NotifyException("bSetExplicitProductPermissionsFlag", "Unable to retrieve ExplicitProductPermissions Flag because of an unexpected system error", ex, True, "Please retry - if the problem persists notify your Account Handler")
            bSetExplicitProductPermissionsFlag = False
        Finally
            oConn.Close()
            If IsDBNull(oDataTable.Rows(0).Item(0)) Then
                bSetExplicitProductPermissionsFlag = False
            Else
                bSetExplicitProductPermissionsFlag = oDataTable.Rows(0).Item(0)
            End If
        End Try
    End Function
  
    Protected Sub DeleteProduct()
        lblError.Text = ""
        Dim oConn As New SqlConnection(gsConn)
        Dim oTrans As SqlTransaction
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Product_Delete", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
  
        Dim paramUserKey As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int, 4)
        paramUserKey.Value = CLng(Session("UserKey"))
        oCmd.Parameters.Add(paramUserKey)
        Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
        paramCustomerKey.Value = CLng(Session("CustomerKey"))
        oCmd.Parameters.Add(paramCustomerKey)
        Dim paramProductKey As SqlParameter = New SqlParameter("@ProductKey", SqlDbType.Int, 4)
        paramProductKey.Value = plProductKey
        oCmd.Parameters.Add(paramProductKey)
        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
            Call BindProductGridDispatcher()
            Call GetProductNumbers()
            Call ShowMainPanel()
        Catch ex As SqlException
            lblError.Text = "Error in DeleteProduct: " & ex.Message
        Finally
            oConn.Close()
            ResetForm()
        End Try
    End Sub
  
    Protected Sub GetProductNumbers()
        Dim oConn As New SqlConnection(gsConn)
        Dim sStoredProcedure As String
        If Session("UserType").ToString.ToLower.Contains("owner") Then
            sStoredProcedure = "spASPNET_Product_GetNumbersOwned"
        Else
            sStoredProcedure = "spASPNET_Product_GetNumbers"
        End If
        Dim oCmd As New SqlCommand(sStoredProcedure, oConn)
        oCmd.CommandType = CommandType.StoredProcedure

        Dim ParamCustomerKey As SqlParameter = oCmd.Parameters.Add("@CustomerKey", SqlDbType.Int, 4)
        ParamCustomerKey.Value = CLng(Session("CustomerKey"))

        If Session("UserType").ToString.ToLower.Contains("owner") Then
            Dim ParamUserKey As SqlParameter = oCmd.Parameters.Add("@UserKey", SqlDbType.Int, 4)
            ParamUserKey.Value = Session("UserKey")
        End If

        Dim ParamProductCount As SqlParameter = oCmd.Parameters.Add("@ProductCount", SqlDbType.Int, 4)
        ParamProductCount.Direction = ParameterDirection.Output

        Dim ParamArchivedProductCount As SqlParameter = oCmd.Parameters.Add("@ArchivedProductCount", SqlDbType.Int, 4)
        ParamArchivedProductCount.Direction = ParameterDirection.Output

        Try
            oConn.Open()
            oCmd.ExecuteNonQuery()
            Dim nProductCount As Integer = CInt(ParamProductCount.Value)
            If nProductCount = 1 Then
                lblProductCountText.Text = "live product and"
            Else
                lblProductCountText.Text = "live products and"
            End If
            lblProductCount.Text = ParamProductCount.Value

            Dim nArchivedProductCount As Integer = CInt(ParamArchivedProductCount.Value)
            If nArchivedProductCount = 1 Then
                lblArchivedProductCountText.Text = "archived product"
            Else
                lblArchivedProductCountText.Text = "archived products"
            End If
            lblArchivedProductCount.Text = ParamArchivedProductCount.Value
        Catch ex As SqlException
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
    End Sub
  
    Protected Sub BindProductUserProfileGrid(ByVal sSearchCriteria As String, ByVal sSortOrder As String)
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataTable As New DataTable()
        sSearchCriteria = sSearchCriteria.Trim
        Dim oAdapter As New SqlDataAdapter("", oConn)
        If sSearchCriteria.Length = 0 OrElse sSearchCriteria = "_" Then
            oAdapter.SelectCommand.CommandText = "spASPNET_Product_GetUserProfilesFromKey"
        Else
            oAdapter.SelectCommand.CommandText = "spASPNET_Product_GetUserProfilesMatchingSearch"
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SearchCriteria", SqlDbType.NVarChar, 50))
            oAdapter.SelectCommand.Parameters("@SearchCriteria").Value = sSearchCriteria
        End If
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ProductKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@ProductKey").Value = plProductKey
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Session("CustomerKey")
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        Try
            oAdapter.Fill(oDataTable)

            If oDataTable.Rows.Count > 0 Then
                lblProductProfileMessage.Text = ""
                Dim oDataView As DataView = oDataTable.DefaultView
                oDataView.Sort = sSortOrder
                grid_ProductUsers.DataSource = oDataView
                grid_ProductUsers.DataBind()
                grid_ProductUsers.Visible = True
            Else
                grid_ProductUsers.Visible = False
                lblProductProfileMessage.Text = "No users found"
            End If
        Catch ex As SqlException
            lblProductProfileMessage.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
    End Sub

    Protected Sub SortProductUsersGrid(ByVal Source As Object, ByVal E As DataGridSortCommandEventArgs)
        Call BindProductUserProfileGrid(txtProductUserSearch.Text, E.SortExpression)
    End Sub
  
    Protected Sub btnToggleAllowToPickCheckboxes_Click(ByVal s As Object, ByVal e As EventArgs)
        Dim dgi As DataGridItem
        Dim cb As CheckBox
        Dim b As Button = s
        For Each dgi In grid_ProductUsers.Items
            cb = dgi.FindControl("cbAllowToPick")
            If b.Text = "select all" Then
                cb.Checked = True
            Else
                cb.Checked = False
            End If
        Next
        If b.Text = "select all" Then
            b.Text() = "clear all"
        Else
            b.Text = "select all"
        End If
    End Sub
  
    Protected Sub btnToggleApplyMaxGrabCheckboxes_Click(ByVal s As Object, ByVal e As EventArgs)
        Dim dgi As DataGridItem
        Dim cb As CheckBox
        Dim b As Button = s
        For Each dgi In grid_ProductUsers.Items
            cb = dgi.FindControl("cbApplyMaxGrab")
            If b.Text = "select all" Then
                cb.Checked = True
            Else
                cb.Checked = False
            End If
        Next
        If b.Text = "select all" Then
            b.Text() = "clear all"
        Else
            b.Text = "select all"
        End If
    End Sub

    Protected Sub btn_ApplyMaxGrabQty_Click(ByVal s As Object, ByVal e As EventArgs)
        Dim dgi As DataGridItem
        Dim tb As TextBox
        If IsNumeric(txtDefaultGrabQty.Text) AndAlso (CInt(txtDefaultGrabQty.Text) >= 0 And CInt(txtDefaultGrabQty.Text) <= 99999) Then
            For Each dgi In grid_ProductUsers.Items
                tb = dgi.FindControl("txtMaxGrabQty")
                tb.Text = txtDefaultGrabQty.Text
            Next
        End If
    End Sub

    Protected Sub btnSaveProductUserProfileChanges_click(ByVal s As Object, ByVal e As EventArgs)
        Call SaveProductUserProfileChanges()
    End Sub
  
    Protected Sub SaveProductUserProfileChanges()
        Dim dgi As DataGridItem
        Dim cb As CheckBox
        Dim tb As TextBox
        Dim bMaxGrabQtyAllValid As Boolean = True
        For Each dgi In grid_ProductUsers.Items
            tb = dgi.FindControl("txtMaxGrabQty")
            Dim sMaxGrabQty As String = tb.Text.Trim
            If sMaxGrabQty.Length > 0 AndAlso Not IsNumeric(sMaxGrabQty) Then
                tb.ForeColor = Drawing.Color.Red
                tb.Font.Bold = True
                bMaxGrabQtyAllValid = False
            Else
                tb.ForeColor = Nothing
                tb.Font.Bold = Nothing
            End If
        Next
        If Not bMaxGrabQtyAllValid Then
            WebMsgBox.Show("One or more max grab quantities has a value that is not a valid number - check max grab values are blank or in the range 0 - 99999")
        Else
            Dim oConn As New SqlConnection(gsConn)
            Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Product_SetUserProductProfile", oConn)
            oCmd.CommandType = CommandType.StoredProcedure
            Dim paramUserKey As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int, 4)
            oCmd.Parameters.Add(paramUserKey)
            Dim paramProductKey As SqlParameter = New SqlParameter("@ProductKey", SqlDbType.Int, 4)
            paramProductKey.Value = plProductKey
            oCmd.Parameters.Add(paramProductKey)
            Dim paramAbleToView As SqlParameter = New SqlParameter("@AbleToView", SqlDbType.Bit)
            paramAbleToView.Value = 1
            oCmd.Parameters.Add(paramAbleToView)
            Dim paramAbleToPick As SqlParameter = New SqlParameter("@AbleToPick", SqlDbType.Bit)
            oCmd.Parameters.Add(paramAbleToPick)
            Dim paramAbleToEdit As SqlParameter = New SqlParameter("@AbleToEdit", SqlDbType.Bit)
            paramAbleToEdit.Value = 0
            oCmd.Parameters.Add(paramAbleToEdit)
            Dim paramAbleToArchive As SqlParameter = New SqlParameter("@AbleToArchive", SqlDbType.Bit)
            paramAbleToArchive.Value = 0
            oCmd.Parameters.Add(paramAbleToArchive)
            Dim paramAbleToDelete As SqlParameter = New SqlParameter("@AbleToDelete", SqlDbType.Bit)
            paramAbleToDelete.Value = 0
            oCmd.Parameters.Add(paramAbleToDelete)
            Dim paramApplyMaxGrab As SqlParameter = New SqlParameter("@ApplyMaxGrab", SqlDbType.Bit)
            oCmd.Parameters.Add(paramApplyMaxGrab)
            Dim paramMaxGrabQty As SqlParameter = New SqlParameter("@MaxGrabQty", SqlDbType.Int, 4)
            oCmd.Parameters.Add(paramMaxGrabQty)
            Try
                oConn.Open()
                oCmd.Connection = oConn
                For Each dgi In grid_ProductUsers.Items
                    oCmd.Parameters("@UserKey").Value = dgi.Cells(0).Text
                    cb = dgi.FindControl("cbAllowToPick")
                    oCmd.Parameters("@AbleToPick").Value = cb.Checked
                    cb = dgi.FindControl("cbApplyMaxGrab")
                    oCmd.Parameters("@ApplyMaxGrab").Value = cb.Checked
                    tb = dgi.FindControl("txtMaxGrabQty")
                    Dim sTest = tb.Text.Trim
                    If tb.Text.Trim.Length > 0 Then
                        oCmd.Parameters("@MaxGrabQty").Value = CInt(tb.Text)
                    Else
                        oCmd.Parameters("@MaxGrabQty").Value = 0
                    End If
                    oCmd.ExecuteNonQuery()
                Next
            Catch ex As Exception
                NotifyException("SaveProductUserProfileChanges", "Unable to update user permissions because of an unexpected system error", ex, True, "Please retry - if the problem persists notify your Account Handler")
            Finally
                oConn.Close()
            End Try
            Call ReturnToProductDetail()
        End If
    End Sub
  
    Protected Sub ResetCategoryDropdowns()
        pbIsAddingCategory = False
        pbIsAddingSubCategory = False
        pbIsAddingSubSubCategory = False
        Call CheckVisibility()
    End Sub
  
    Protected Sub ddlCategory_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        If ddl.SelectedItem.Text = NEW_CATEGORY Then
            pbIsAddingCategory = True
        End If
        Call CheckVisibility()
    End Sub

    Protected Sub ddlSubCategory_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        If ddl.SelectedItem.Text = NEW_SUBCATEGORY Then
            pbIsAddingSubCategory = True
        End If
        Call CheckVisibility()
    End Sub
  
    Protected Sub ddlSubSubCategory_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        If ddl.SelectedItem.Text = NEW_SUBCATEGORY Then
            pbIsAddingSubSubCategory = True
        End If
        Call CheckVisibility()
    End Sub
  
    Protected Sub CheckVisibility()
        pbIsAddingCategory = pbIsAddingCategory
        pbIsAddingSubCategory = pbIsAddingSubCategory
        pbIsAddingSubSubCategory = pbIsAddingSubSubCategory
    End Sub

    Protected Sub ddlCategory_DataBound(ByVal sender As Object, ByVal e As System.EventArgs)
        gbDataBound = True
        Call AdjustddlCategory()
    End Sub

    Protected Sub ddlSubCategory_DataBound(ByVal sender As Object, ByVal e As System.EventArgs)
        gbDataBound = True
        Call AdjustddlSubCategory()
    End Sub

    Protected Sub ddlSubSubCategory_DataBound(ByVal sender As Object, ByVal e As System.EventArgs)
        gbDataBound = True
        Call AdjustddlSubSubCategory()
    End Sub

    Protected Sub AdjustddlCategory()
        ddlCategory.Items.Insert(0, NEW_CATEGORY)
        ddlCategory.Items.Insert(0, "")
        If Not pbIsAddingCategory Then
            If pbIsAddingNew Then
                ddlCategory.SelectedIndex = 0
            Else
                Dim i As Integer
                For i = 0 To ddlCategory.Items.Count - 1
                    If ddlCategory.Items(i).Text.Trim.ToLower = hidCategory.Value.ToString.Trim.ToLower Then
                        ddlCategory.SelectedIndex = i
                        Exit For
                    End If
                Next
            End If
        End If
    End Sub

    Protected Function ContainsNewSubCategory(ByVal ddl As DropDownList) As Boolean
        ContainsNewSubCategory = False
        If ddl.Items.Count > 0 Then
            For Each li As ListItem In ddl.Items
                If li.Text = NEW_SUBCATEGORY Then
                    ContainsNewSubCategory = True
                    Exit Function
                End If
            Next
        End If
    End Function
  
    Protected Sub AdjustddlSubCategory()
        If Not ContainsNewSubCategory(ddlSubCategory) Then
            ddlSubCategory.Items.Insert(0, NEW_SUBCATEGORY)
            ddlSubCategory.Items.Insert(0, "")
        End If
        If Not pbIsAddingSubCategory Then
            If pbIsAddingNew Then
                ddlSubCategory.SelectedIndex = 0
            Else
                Dim i As Integer
                For i = 0 To ddlSubCategory.Items.Count - 1
                    If ddlSubCategory.Items(i).Text.Trim.ToLower = hidSubCategory.Value.ToString.Trim.ToLower Then   'CN 3FEB09
                        ddlSubCategory.SelectedIndex = i
                        Exit For
                    End If
                Next
            End If
        End If
    End Sub
  
    Protected Sub AdjustddlSubSubCategory()
        If Not ContainsNewSubCategory(ddlSubSubCategory) Then
            ddlSubSubCategory.Items.Insert(0, NEW_SUBCATEGORY)
            ddlSubSubCategory.Items.Insert(0, "")
        End If
        If Not pbIsAddingSubSubCategory Then
            If pbIsAddingNew Then
                ddlSubSubCategory.SelectedIndex = 0
            Else
                Dim i As Integer
                For i = 0 To ddlSubSubCategory.Items.Count - 1
                    If ddlSubSubCategory.Items(i).Text.Trim.ToLower = hidSubSubCategory.Value.ToString.Trim.ToLower Then   'CN 3FEB09
                        ddlSubSubCategory.SelectedIndex = i
                        Exit For
                    End If
                Next
            End If
        End If
    End Sub
  
    Protected Sub NotifyException(ByVal sLocation As String, ByVal sReason As String, Optional ByVal ex As Exception = Nothing, Optional ByVal bContinue As Boolean = False, Optional ByVal sAdviceString As String = "")
        Dim sbMessage As New StringBuilder
        sbMessage.Append(sReason & " in " & sLocation)
        If ex IsNot Nothing Then
            sbMessage.Append(vbCrLf & vbCrLf & "Exception: ")
            sbMessage.Append(ex.Message & vbCrLf & vbCrLf)
            sbMessage.Append("Stack Trace: ")
            sbMessage.Append(ex.StackTrace & vbCrLf & vbCrLf)
        End If
        If sAdviceString.Length > 0 Then
            sbMessage.Append(sAdviceString)
        End If
        WebMsgBox.Show(sbMessage.ToString.Replace("'", "*").Replace("""", "*").Replace(vbLf, "").Replace(vbCr, "\n"))
    End Sub
  
    Protected Sub ResetForm()
        txtProductCode.Text = String.Empty
        txtProductDate.Text = String.Empty
        txtMinStockLevel.Text = String.Empty
        txtDescription.Text = String.Empty
        txtItemsPerBox.Text = String.Empty
        txtCategory.Text = String.Empty
        txtSubCategory.Text = String.Empty
        tbSubSubCategory.Text = String.Empty
        txtUnitValue.Text = String.Empty
        txtLanguage.Text = String.Empty
        txtUnitWeight.Text = String.Empty
        txtMisc1.Text = String.Empty
        txtMisc2.Text = String.Empty
        chkArchivedFlag.Checked = False
        tbExpiryDate.Text = String.Empty
        tbReplenishmentDate.Text = String.Empty
        pbIsAddingCategory = False
        pbIsAddingSubCategory = False
        chkProspectusNumbers.Checked = False
        txtAdRotatorText.Text = String.Empty
        chkAdRotator.Checked = False
        chkViewOnWebForm.Checked = False
        txtNotes.Text = String.Empty
        hlnk_PDF.ImageUrl = String.Empty
        hlnk_PDF.NavigateUrl = String.Empty
        hlnk_DetailThumb.ImageUrl = String.Empty
        hlnk_DetailThumb.NavigateUrl = String.Empty
        ddlStatus.SelectedIndex = 0
        txtDepartment.Text = String.Empty
    End Sub

    Protected Sub PopulateSuperUserDropdown(ByVal ddl As DropDownList)
        Dim oConn As New SqlConnection(gsConn)
        Dim oDatatable As New DataTable
        Dim oAdapter As New SqlDataAdapter("spASPNET_UserProfile_GetAllSuperUsersForCustomer2", oConn)
      
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Session("CustomerKey")

        oAdapter.Fill(oDatatable)
        ddl.Items.Clear()
        Dim li As New ListItem
        li.Text = "-- select authoriser --"
        li.Value = 0
        ddl.Items.Add(li)

        For Each dr As DataRow In oDatatable.Rows
            Dim li2 As New ListItem
            Dim sFirstName As String = dr("FirstName")
            Dim sLastName As String = dr("LastName")
            li2.Text = Char.ToUpper(sFirstName(0)) & sFirstName.Substring(1) & " " & Char.ToUpper(sLastName(0)) & sLastName.Substring(1) & "  (" & dr("UserId") & ")"
            li2.Value = dr("key")
            ddl.Items.Add(li2)
        Next
    End Sub

    Protected Function sTimeStamp() As String
        sTimeStamp = Now
    End Function
  
    Protected Sub btnBackFromProductGroupsToList_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call BackToProductListPanel()
    End Sub
  
    Protected Function GetArchiveFlagLegend() As String
        GetArchiveFlagLegend = "Archive Flag:"
    End Function
  
    Protected Function GetProductDateLegend() As String
        GetProductDateLegend = "Product Date:"
    End Function
  
    Protected Sub SetHelpStatus()
        If Request.Cookies("HelpStatus") Is Nothing Then
            Call ShowHelp()
        Else
            Dim sState As String = Request.Cookies("HelpStatus")("CreateNewProduct")
            lnkbtnShowHelp.Text = sState
            If lnkbtnShowHelp.Text = "show help" Then
                Call HideHelp()
            Else
                Call ShowHelp()
            End If
        End If
    End Sub
  
    Protected Sub btn_AddProduct_click(ByVal sender As Object, ByVal e As System.EventArgs)
        pbIsAddingNew = True
        pbIsAddingCategory = False
        pbIsAddingSubCategory = False
        pbIsAddingSubSubCategory = False
        imgbtnDeleteImage.Visible = False
        imgbtnDeletePDF.Visible = False
        btnUploadImage.Visible = False
        btnUploadPDF.Visible = False
        fuBrowseImageFile.Visible = False
        fuBrowsePDFFile.Visible = False
        lblImageUploadUnavailable.Visible = True
        lblPDFUploadUnavailable.Visible = True
        btnAssociatedProducts.Visible = False
        dg_ProductList.CurrentPageIndex = 0
        btnSetUserProfiles.Visible = False
        btn_DeleteProduct.Visible = False
        tbInactivityAlertDays.Text = GetDefaultInactivityAlertDays()
        Call ShowNewProduct()
        Call SetHelpStatus()
    End Sub
  
    Protected Sub lnkbtnShowHelp_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ToggleHelp()
    End Sub
    
    Protected Sub StoreHelpStatus()
        Dim c As HttpCookie = New HttpCookie("HelpStatus")
        c.Values.Add("CreateNewProduct", lnkbtnShowHelp.Text)
        c.Expires = DateTime.Now.AddDays(365)
        Response.Cookies.Add(c)
        Response.Flush()
    End Sub

    Protected Sub ToggleHelp()
        If lnkbtnShowHelp.Text = "show help" Then
            lnkbtnShowHelp.Text = "hide help"
            Call ShowHelp()
        Else
            lnkbtnShowHelp.Text = "show help"
            Call HideHelp()
        End If
        Call StoreHelpStatus()
        Call CheckVisibility()
    End Sub
  
    Protected Sub SetVisible(ByVal a As HtmlContainerControl)
        If a.InnerText <> "NULL" Then
            a.Visible = True
        End If
    End Sub
   
    Protected Sub ShowHelp()
        If Not pbIsAddingNew Then
            aHelpDeleteProduct.Visible = True
        End If
        Call SetVisible(aHelpProductCode)
        Call SetVisible(aHelpProductDate)
        Call SetVisible(aHelpMinStockLevel)
        Call SetVisible(aHelpDescription)
        Call SetVisible(aHelpItemsPerBox)
        Call SetVisible(aHelpCategory)
        Call SetVisible(aHelpSubCategory)
        Call SetVisible(aHelpArchived)
        Call SetVisible(aHelpSubSubCategory)
        Call SetVisible(aHelpInactivityAlert)

        Call SetVisible(aHelpLanguage)
        Call SetVisible(aHelpProspectusNumbers)
        Call SetVisible(aHelpViewOnWebForm)
        Call SetVisible(aHelpAdRotatorText)
        Call SetVisible(aHelpAdRotator)
        Call SetVisible(aHelpDepartment)
        Call SetVisible(aHelpUnitWeight)
        Call SetVisible(aHelpMisc1)
        Call SetVisible(aHelpMisc2)
        Call SetVisible(aHelpUnitValue)
        Call SetVisible(aHelpExpiryDate)
        Call SetVisible(aHelpReplenishmentDate)
        Call SetVisible(aHelpNotes)
        Call SetVisible(aHelpUploadImage)
        Call SetVisible(aHelpUploadPDF)
    End Sub
  
    Protected Sub HideHelp()
        aHelpDeleteProduct.Visible = False
        aHelpProductCode.Visible = False
        aHelpProductDate.Visible = False
        aHelpMinStockLevel.Visible = False
        aHelpDescription.Visible = False
        aHelpItemsPerBox.Visible = False
        aHelpCategory.Visible = False
        aHelpSubCategory.Visible = False
        aHelpArchived.Visible = False
        aHelpSubSubCategory.Visible = False
        aHelpInactivityAlert.Visible = False
        aHelpLanguage.Visible = False
        aHelpProspectusNumbers.Visible = False
        aHelpViewOnWebForm.Visible = False
        aHelpAdRotatorText.Visible = False
        aHelpAdRotator.Visible = False
        aHelpDepartment.Visible = False
        aHelpUnitWeight.Visible = False
        aHelpMisc1.Visible = False
        aHelpMisc2.Visible = False
        aHelpUnitValue.Visible = False
        aHelpExpiryDate.Visible = False
        aHelpReplenishmentDate.Visible = False
        aHelpNotes.Visible = False
        aHelpUploadImage.Visible = False
        aHelpUploadPDF.Visible = False
    End Sub
  
    Protected Sub lnkbtnNewProductCode_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        txtProductCode.Text = FormatNextFreeProductCode()
        lnkbtnNewProductCode.Visible = False
    End Sub
  
    Protected Function FormatNextFreeProductCode() As String
        Dim sCode As String
        Do
            sCode = CStr(GetNextFreeProductCode())
            While sCode.Length < 6
                sCode = "0" & sCode
            End While
            sCode = "W" & sCode
        Loop Until IsUniqueProductCode(sCode, Session("CustomerKey"))
        FormatNextFreeProductCode = sCode
    End Function

    Protected Function GetNextFreeProductCode() As Integer
        Dim oDataReader As SqlDataReader = Nothing
        Dim oConn As New SqlConnection(gsConn)
        Dim sSQL As String
        Dim nCode As Integer
        sSQL = "SELECT * FROM LANDGControl"
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        oConn.Open()
        oDataReader = oCmd.ExecuteReader()
        GetNextFreeProductCode = 0
        If oDataReader.HasRows Then
            oDataReader.Read()
            nCode = oDataReader("ProductCodeSeed")
        End If
        oDataReader.Close()
        GetNextFreeProductCode = nCode
        nCode += 1
        sSQL = "UPDATE LANDGControl SET ProductCodeSeed = " & nCode
        Dim oCmd2 As SqlCommand = New SqlCommand(sSQL, oConn)
        oCmd2.ExecuteNonQuery()
        oConn.Close()
    End Function

    Protected Function IsUniqueProductCode(ByVal sProductCode As String, ByVal nCustomerKey As Integer) As Boolean
        Dim oDataReader As SqlDataReader = Nothing
        Dim oConn As New SqlConnection(gsConn)
        Dim sSQL As String
        sSQL = "SELECT * FROM LogisticProduct WHERE ProductCode = '" & sProductCode & "' AND CustomerKey = " & nCustomerKey
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        oConn.Open()
        oDataReader = oCmd.ExecuteReader()
        If oDataReader.HasRows Then
            IsUniqueProductCode = False
        Else
            IsUniqueProductCode = True
        End If
        oConn.Close()
    End Function
  
    Protected Sub ddlStatus_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        If ddl.SelectedValue = 4 Then
            chkArchivedFlag.Checked = True
        End If
    End Sub
  
    Protected Function IsBlankOrPositiveInteger(ByVal sString As String) As Integer
        IsBlankOrPositiveInteger = -1
        sString = sString.Trim
        If sString.Length = 0 Then
            Return 0
        End If
        If Not IsNumeric(sString) Then
            Exit Function
        End If
        For Each c As Char In sString
            If Not Char.IsDigit(c) Then
                Return -1
            End If
        Next
        Return CInt(sString)
    End Function
  
    Protected Function gvAuthOrderDetailsItemForeColor(ByVal DataItem As Object) As System.Drawing.Color
        gvAuthOrderDetailsItemForeColor = Black
        If Not IsDBNull(DataBinder.Eval(DataItem, "Authorised")) AndAlso DataBinder.Eval(DataItem, "Authorised") = "N" Then
            gvAuthOrderDetailsItemForeColor = Red
        End If
    End Function

    Protected Sub dgPreAuthorise_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs)
        Dim dgi As DataGridItem = e.Item
        If dgi.ItemType = ListItemType.Item Or dgi.ItemType = ListItemType.AlternatingItem Then
            Dim sKey As String = dgi.Cells(0).Text

            Dim oDataReader As SqlDataReader
            Dim oConn As New SqlConnection(gsConn)
            Dim oCmd As New SqlCommand("spASPNET_Product_GetAuthorised", oConn)
            oCmd.CommandType = CommandType.StoredProcedure
            Dim oParam1 As SqlParameter = oCmd.Parameters.Add("@LogisticProductKey", SqlDbType.Int, 4)
            oParam1.Value = plProductKey
            Dim oParam2 As SqlParameter = oCmd.Parameters.Add("@UserProfileKey", SqlDbType.Int, 4)
            oParam2.Value = sKey
            Try
                oConn.Open()
                oDataReader = oCmd.ExecuteReader()
                If oDataReader.HasRows Then
                    oDataReader.Read()
                    Dim sAuthorisedQuantity As String = oDataReader("AuthorisedQuantity")
                    Dim sQuantityRemaining As String = oDataReader("QuantityRemaining")
                    Dim l As Label = dgi.Cells(4).Controls(1)
                    l.ForeColor = Red
                    l.Text = "(" & sQuantityRemaining & " unused of " & sAuthorisedQuantity & " auth'd)"
                End If
                oDataReader.Close()
            Catch ex As SqlException
                lblError.Text = ex.ToString
            Finally
                oConn.Close()
            End Try
        End If
    End Sub

    Protected Sub imgbtnDeleteImage_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Call DeleteImage()
    End Sub
  
    Protected Sub DeleteImage()
        Dim SavePath As String = psProdThumbFolder & plProductKey.ToString & ".jpg"
        If System.IO.File.Exists(SavePath) Then
            System.IO.File.Delete(SavePath)
        End If
        Dim sSavePath As String = psProdImageFolder & plProductKey.ToString & ".jpg"
        If System.IO.File.Exists(sSavePath) Then
            System.IO.File.Delete(sSavePath)
        End If
        Call ResetImageAttributes()
        hlnk_DetailThumb.ImageUrl = psVirtualThumbFolder & "blank_thumb.jpg"
        hlnk_DetailThumb.NavigateUrl = psVirtualJPGFolder & "blank_image.jpg"
        imgbtnDeleteImage.Visible = False
    End Sub

    Protected Sub imgbtnDeletePDF_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Call DeletePDF()
    End Sub
  
    Protected Sub DeletePDF()
        hlnk_PDF.ImageUrl = psVirtualPDFFolder & "blank_pdf_thumb.jpg"
        hlnk_PDF.NavigateUrl = psVirtualPDFFolder & "blank_pdf.jpg"
        Call SetPDFAttribute()
        imgbtnDeletePDF.Visible = False
    End Sub
  
    Protected Sub btnAssociatedProducts_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowAssociatedProductsPanel()
    End Sub

    Protected Sub InitAssociatedProductsPanel()
        Call PopulateAssociatedProductsGrid()
        Call PopulateUnassociatedProductsGrid()
        If txtProductDate.Text.Trim.Length > 0 Then
            lblAssociatedProductsProductCode.Text = txtProductCode.Text & "-" & txtProductDate.Text
        Else
            lblAssociatedProductsProductCode.Text = txtProductCode.Text
        End If
    End Sub
   
    Protected Sub PopulateAssociatedProductsGrid()
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataTable As New DataTable
        Dim oAdapter As New SqlDataAdapter("spASPNET_Product_GetAssociatedProducts", oConn)
        Try
            oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@LogisticProductKey", SqlDbType.Int))
            oAdapter.SelectCommand.Parameters("@LogisticProductKey").Value = plProductKey

            oAdapter.Fill(oDataTable)
            gvAssociatedProducts.DataSource = oDataTable
            gvAssociatedProducts.DataBind()
        Catch ex As SqlException
        Finally
            oConn.Close()
        End Try
    End Sub
   
    Protected Sub PopulateUnassociatedProductsGrid()
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataTable As New DataTable
        Dim oAdapter As New SqlDataAdapter("spASPNET_Product_GetUnassociatedProducts", oConn)
        Try
            oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
            oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Session("CustomerKey")
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@LogisticProductKey", SqlDbType.Int))
            oAdapter.SelectCommand.Parameters("@LogisticProductKey").Value = plProductKey

            oAdapter.Fill(oDataTable)
            gvUnassociatedProducts.DataSource = oDataTable
            gvUnassociatedProducts.DataBind()
        Catch ex As SqlException
        Finally
            oConn.Close()
        End Try
    End Sub

    Protected Sub btnAddAssociatedProduct_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim b As Button = sender
        Dim nAssociatedProductKey As Integer = b.CommandArgument
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Product_AddAssociatedProduct", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Dim paramLogisticProductKey As SqlParameter = New SqlParameter("@LogisticProductKey", SqlDbType.Int)
        paramLogisticProductKey.Value = plProductKey
        oCmd.Parameters.Add(paramLogisticProductKey)
        Dim paramAssociatedProductKey As SqlParameter = New SqlParameter("@AssociatedProductKey", SqlDbType.Int)
        paramAssociatedProductKey.Value = nAssociatedProductKey
        oCmd.Parameters.Add(paramAssociatedProductKey)
        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As SqlException
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
        Call InitAssociatedProductsPanel()
    End Sub

    Protected Sub btnRemoveAssociatedProduct_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim b As Button = sender
        Dim nLogisticAssociatedProductKey As Integer = b.CommandArgument
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Product__Product_DeleteAssociatedProduct", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Dim paramLogisticAssociatedProductKey As SqlParameter = New SqlParameter("@LogisticAssociatedProductKey", SqlDbType.Int)
        paramLogisticAssociatedProductKey.Value = nLogisticAssociatedProductKey
        oCmd.Parameters.Add(paramLogisticAssociatedProductKey)
        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As SqlException
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
        Call InitAssociatedProductsPanel()
    End Sub
   
    Property plProductKey() As Long
        Get
            Dim o As Object = ViewState("PM_ProductKey")
            If o Is Nothing Then
                Return -1
            End If
            Return CLng(o)
        End Get
        Set(ByVal Value As Long)
            ViewState("PM_ProductKey") = Value
        End Set
    End Property
  
    Property psProdImageFolder() As String
        Get
            Dim o As Object = ViewState("PM_ProdImageFolder")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("PM_ProdImageFolder") = Value
        End Set
    End Property
  
    Property psVirtualJPGFolder() As String
        Get
            Dim o As Object = ViewState("PM_VirtualJPGFolder")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("PM_VirtualJPGFolder") = Value
        End Set
    End Property
  
    Property psProdThumbFolder() As String
        Get
            Dim o As Object = ViewState("PM_ProdThumbFolder")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("PM_ProdThumbFolder") = Value
        End Set
    End Property
  
    Property psVirtualThumbFolder() As String
        Get
            Dim o As Object = ViewState("PM_VirtualThumbFolder")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("PM_VirtualThumbFolder") = Value
        End Set
    End Property
  
    Property psProdPDFFolder() As String
        Get
            Dim o As Object = ViewState("PM_ProdPDFFolder")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("PM_ProdPDFFolder") = Value
        End Set
    End Property
  
    Property psVirtualPDFFolder() As String
        Get
            Dim o As Object = ViewState("PM_VirtualPDFFolder")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("PM_VirtualPDFFolder") = Value
        End Set
    End Property
  
    Property pbIsAddingNew() As Boolean
        Get
            Dim o As Object = ViewState("PM_IsAddingNew")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("PM_IsAddingNew") = Value
            If Value = True Then
                ResetForm()
                lblProductQuantity.Text = "0"
                txtProductCode.Enabled = True
                txtProductDate.Enabled = True
            Else
                txtProductCode.Enabled = False
                txtProductDate.Enabled = False
            End If
        End Set
    End Property

    Property pbIsAddingCategory() As Boolean
        Get
            Dim o As Object = ViewState("PM_IsAddingCategory")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("PM_IsAddingCategory") = Value
            If Value = True Then
                txtCategory.Visible = True
                ddlCategory.Visible = False
                txtCategory.Focus()
            Else
                txtCategory.Visible = False
                ddlCategory.Visible = True
            End If
        End Set
    End Property

    Property pbIsAddingSubCategory() As Boolean
        Get
            Dim o As Object = ViewState("PM_IsAddingSubCategory")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("PM_IsAddingSubCategory") = Value
            If Value = True Then
                txtSubCategory.Visible = True
                ddlSubCategory.Visible = False
                txtSubCategory.Focus()
            Else
                txtSubCategory.Visible = False
                ddlSubCategory.Visible = True
            End If
        End Set
    End Property

    Property pbIsAddingSubSubCategory() As Boolean
        Get
            Dim o As Object = ViewState("PM_IsAddingSubSubCategory")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("PM_IsAddingSubSubCategory") = Value
            If Value = True Then
                tbSubSubCategory.Visible = True
                ddlSubSubCategory.Visible = False
                tbSubSubCategory.Focus()
            Else
                tbSubSubCategory.Visible = False
                ddlSubSubCategory.Visible = True
            End If
        End Set
    End Property

    Property plPerCustomerConfiguration() As Long
        Get
            Dim o As Object = ViewState("PM_PerCustomerConfiguration")
            If o Is Nothing Then
                Return PER_CUSTOMER_CONFIGURATION_NONE
            End If
            Return CLng(o)
        End Get
        Set(ByVal Value As Long)
            ViewState("PM_PerCustomerConfiguration") = Value
        End Set
    End Property
  
    Property pnCategoryMode() As Integer
        Get
            Dim o As Object = ViewState("PM_CategoryMode")
            If o Is Nothing Then
                Return 2
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Integer)
            ViewState("PM_CategoryMode") = Value
        End Set
    End Property
  
    Property pbUsesCategories() As Boolean
        Get
            Dim o As Object = ViewState("PM_UsesCategories")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("PM_UsesCategories") = Value
        End Set
    End Property
  
    Property psCategory() As String
        Get
            Dim o As Object = ViewState("PM_Category")
            If o Is Nothing Then
                Return "_"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("PM_Category") = Value
        End Set
    End Property
  
    Property psSubCategory() As String
        Get
            Dim o As Object = ViewState("PM_SubCategory")
            If o Is Nothing Then
                Return "_"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("PM_SubCategory") = Value
        End Set
    End Property
  
    Property psSubSubCategory() As String
        Get
            Dim o As Object = ViewState("PM_SubSubCategory")
            If o Is Nothing Then
                Return "_"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("PM_SubSubCategory") = Value
        End Set
    End Property
  
    Property psDisplayMode() As String
        Get
            Dim o As Object = ViewState("PM_DisplayMode")
            If o Is Nothing Then
                Return "_"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("PM_DisplayMode") = Value
        End Set
    End Property
  
    Protected Sub lnkbtnConfigureProductInactivityAlerts_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call HideAllPanels()
        Dim nDefaultInactivityAlertDays As Integer = GetDefaultInactivityAlertDays()
        tbProductInactivityAlertPeriodExistingProducts.Text = nDefaultInactivityAlertDays
        tbProductInactivityAlertPeriodNewProducts.Text = nDefaultInactivityAlertDays
        If Session("UserType").ToString.ToLower.Contains("owner") Then
            btnSetProductInactivityAlertNewProducts.Enabled = False
            tbProductInactivityAlertPeriodNewProducts.Enabled = False
            lblAvailableToSuperUsersOnly.Visible = True
        End If
        pnlConfigureProductInactivityAlert.Visible = True
    End Sub

    Protected Sub btnViewProductsUsingInactivityAlert_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call HideAllPanels()
        Call ViewProductsUsingInactivityAlert()
        pnlProductInactivityAlertStatus.Visible = True
    End Sub
    
    Protected Sub ViewProductsUsingInactivityAlert()
        Dim oDataReader As SqlDataReader = Nothing
        Dim oConn As New SqlConnection(gsConn)
        Dim sSelectFields As String = "SELECT ProductCode 'Product Code', ProductDate 'Value/Date', ProductDescription 'Description', InactivityAlertDays 'Inactivity Alert Days' FROM LogisticProduct lp "
        Dim sCondition As String = " AND ISNULL(InactivityAlertDays,0) > 0 ORDER BY LogisticProductKey"
        Dim sSQL As String
        If Session("UserType").ToString.ToLower.Contains("owner") Then
            sSQL = sSelectFields & "INNER JOIN ProductGroup pg ON lp.StockOwnedByKey = pg.ProductGroupKey WHERE pg.ProductOwner1 = " & Session("UserKey") & " OR pg.ProductOwner2 = " & Session("UserKey")
        Else
            sSQL = sSelectFields & "WHERE CustomerKey = " & Session("CustomerKey") & sCondition
        End If
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            gvProductInactivityAlert.DataSource = oDataReader
            gvProductInactivityAlert.DataBind()
        Catch ex As Exception
            WebMsgBox.Show("Error in ViewProductsUsingInactivityAlert: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Function GetDefaultInactivityAlertDays() As Integer
        GetDefaultInactivityAlertDays = 0
        Dim oDataReader As SqlDataReader = Nothing
        Dim oConn As New SqlConnection(gsConn)
        Dim sSQL As String = "SELECT DefaultInactivityAlertDays FROM Customer WHERE CustomerKey = " & Session("CustomerKey")
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            oDataReader.Read()
            If Not IsDBNull(oDataReader(0)) Then
                GetDefaultInactivityAlertDays = oDataReader(0)
            End If
        Catch ex As Exception
            WebMsgBox.Show("Error in GetDefaultInactivityAlertDays: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Function
    
    Protected Sub btnBackFromConfigureProductInactivityAlert_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowProductDetail()
    End Sub
    
    Protected Sub btnBackFromProductInactivityAlertStatus_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowProductDetail()
    End Sub
    
    Protected Sub btnSetProductInactivityAlertAllExistingProducts_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call SetProductInactivityAlertAllExistingProducts()
    End Sub

    Protected Sub SetProductInactivityAlertAllExistingProducts()
        Dim oConn As New SqlConnection(gsConn)
        Dim sSelectField As String = "SELECT LogisticProductKey FROM LogisticProduct "
        Dim sSQL As String
        Dim nResult As Integer
        If Session("UserType").ToString.ToLower.Contains("owner") Then
            sSQL = sSelectField & "INNER JOIN ProductGroup pg ON lp.StockOwnedByKey = pg.ProductGroupKey WHERE pg.ProductOwner1 = " & Session("UserKey") & " OR pg.ProductOwner2 = " & Session("UserKey")
        Else
            sSQL = sSelectField & "WHERE CustomerKey = " & Session("CustomerKey")
        End If
        sSQL = "UPDATE LogisticProduct SET InactivityAlertDays = " & tbProductInactivityAlertPeriodExistingProducts.Text & " WHERE LogisticProductKey IN (" & sSQL & ")"
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        Try
            oConn.Open()
            nResult = oCmd.ExecuteNonQuery()
        Catch ex As Exception
            WebMsgBox.Show("Error in SetProductInactivityAlertAllExistingProducts: " & ex.Message)
        Finally
            oConn.Close()
        End Try
        Dim sProducts As String = " product"
        If nResult <> 1 Then
            sProducts += "s"
        End If
        Dim sDays As String = " day"
        If CInt(tbProductInactivityAlertPeriodExistingProducts.Text) > 1 Then
            sDays += "s"
            sDays += "."
        End If
        If CInt(tbProductInactivityAlertPeriodExistingProducts.Text) > 0 Then
            WebMsgBox.Show(nResult.ToString & sProducts & " assigned an Inactivity Alert value of " & tbProductInactivityAlertPeriodExistingProducts.Text & " day(s)")
        Else
            WebMsgBox.Show(nResult.ToString & sProducts & " will no longer receive Inactivity Alerts")
        End If
        If Not pbIsAddingNew Then
            tbInactivityAlertDays.Text = tbProductInactivityAlertPeriodExistingProducts.Text
        End If
    End Sub
    
    Protected Sub btnSetProductInactivityAlertNewProducts_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call SetProductInactivityAlertNewProducts()
    End Sub
    
    Protected Sub SetProductInactivityAlertNewProducts()
        Dim oConn As New SqlConnection(gsConn)
        Dim sSelectField As String = "SELECT LogisticProductKey FROM LogisticProduct "
        Dim sSQL As String
        sSQL = "UPDATE Customer SET DefaultInactivityAlertDays = " & tbProductInactivityAlertPeriodNewProducts.Text & " WHERE CustomerKey = " & Session("CustomerKey")
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        Try
            oConn.Open()
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            WebMsgBox.Show("Error in SetProductInactivityAlertAllExistingProducts: " & ex.Message)
        Finally
            oConn.Close()
        End Try
        Dim sDays As String = " day"
        If CInt(tbProductInactivityAlertPeriodNewProducts.Text) > 1 Then
            sDays += "s"
        End If
        sDays += "."
        If CInt(tbProductInactivityAlertPeriodNewProducts.Text) > 0 Then
            WebMsgBox.Show("Newly created products will be assigned an Inactivity Alert value of " & tbProductInactivityAlertPeriodNewProducts.Text & sDays)
        Else
            WebMsgBox.Show("Newly created products will not receive Inactivity Alerts.")
        End If
        If pbIsAddingNew Then
            tbInactivityAlertDays.Text = tbProductInactivityAlertPeriodNewProducts.Text
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

    Protected Function GetTariffEntries(ByVal nTariffId As Integer) As ListItemCollection
        GetTariffEntries = ExecuteQueryToListItemCollection("SELECT Quantity, Price FROM OnDemandTariff WHERE TariffId = " & nTariffId & " ORDER BY Quantity", "Quantity", "Price")
    End Function
    
    Protected Sub ddlItemsPerPage_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        dg_ProductList.PageSize = ddl.SelectedValue
        dg_ProductList.CurrentPageIndex = 0
        Call BindProductGridDispatcher()
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
    
</script>


<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Product Manager</title>
</head>
<body>
    <form id="frmProductManager" method="post" enctype="multipart/form-data" runat="server">
    <main:Header ID="ctlHeader" runat="server"></main:Header>
    <table width="100%" cellpadding="0" cellspacing="0">
        <tr class="bar_productmanager">
            <td style="width: 50%; white-space: nowrap">
            </td>
            <td style="width: 50%; white-space: nowrap" align="right">
                <asp:Label ID="Label3" runat="server" Font-Names="Verdana" Font-Size="XX-Small" ForeColor="white"
                    Text="You have" />
                &nbsp;<asp:Label runat="server" ID="lblProductCount" ForeColor="#F9D938" Font-Names="Verdana"
                    Font-Size="XX-Small" Font-Bold="true"></asp:Label>
                &nbsp;<asp:Label ID="lblProductCountText" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                    ForeColor="white" Text="live products and" />
                &nbsp;<asp:Label runat="server" ID="lblArchivedProductCount" ForeColor="#F9D938"
                    Font-Names="Verdana" Font-Size="XX-Small" Font-Bold="true"></asp:Label>
                &nbsp;<asp:Label ID="lblArchivedProductCountText" runat="server" Font-Names="Verdana"
                    Font-Size="XX-Small" ForeColor="white" Text="archived products" />&nbsp;&nbsp;
            </td>
        </tr>
    </table>
    <asp:Panel ID="pnlMain" runat="server" Width="100%" Visible="true">
        <table width="100%" style="font-family: Verdana; font-size: x-small">
            <tr valign="middle">
                <td align="left" valign="middle" style="white-space: nowrap">
                    <asp:Button ID="btn_ShowAllProducts" runat="server" OnClick="btn_ShowAllProducts_Click"
                        Text="show all products" ToolTip="get full product list" />
                    &nbsp;&nbsp;<asp:Button ID="btnShowCategories" runat="server" Text="show categories"
                        OnClick="btnShowCategories_Click" />
                    &nbsp;&nbsp;<asp:Label ID="Label19" runat="server" ForeColor="Gray" Font-Size="XX-Small"
                        Font-Names="Verdana">search:</asp:Label>
                    &nbsp;<asp:TextBox runat="server" Width="80px" Font-Size="XX-Small" ID="txtSearchCriteriaAllProducts"></asp:TextBox>
                    &nbsp;<asp:Button ID="btn_SearchAllProducts" OnClick="btn_SearchAllProducts_Click"
                        runat="server" Text="go" ToolTip="search across all products" />
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:Button ID="btnAddProduct" OnClick="btn_AddProduct_click"
                        Text="add product" runat="server" />
                </td>
                <td align="right" valign="middle" style="white-space: nowrap">
                </td>
            </tr>
        </table>
        <br />
    </asp:Panel>
    <asp:Panel ID="pnlCategorySelection1" runat="server" Visible="True" Width="100%">
        <table id="tblCategorySelection" runat="server" width="100%" style="font-family: Verdana;
            font-size: small" cellpadding="2" cellspacing="1">
            <tr>
                <td style="width: 2%">
                </td>
                <td valign="top" style="white-space: nowrap; width: 48%; background-color: #DFE5E2">
                    &nbsp;&nbsp;<asp:Label ID="Label111" runat="server" ForeColor="Navy" Font-Bold="True"
                        Font-Size="X-Small">Product Categories</asp:Label>
                    <br />
                    <br />
                    <asp:Repeater runat="server" ID="Repeater1" OnItemCommand="repeater1_Item_click">
                        <ItemTemplate>
                            &nbsp;&nbsp;&nbsp;<asp:LinkButton ID="LinkButton5" runat="server" Style="text-decoration: none"
                                OnCommand="lnkbtnShowSubCategories_click" CommandArgument='<%# Container.DataItem("Category")%>'
                                Text='<%# Container.DataItem("Category")%>' ForeColor="Navy" Font-Size="X-Small"
                                EnableTheming="false" />
                            <br />
                        </ItemTemplate>
                    </asp:Repeater>
                    <br />
                </td>
                <td valign="top" style="white-space: nowrap; width: 48%; background-color: #DFE5E2">
                    &nbsp;&nbsp;<asp:Label runat="server" ID="lblSubCategoryHeading" ForeColor="Navy"
                        Font-Bold="True" Font-Size="X-Small">Sub Categories</asp:Label>
                    <br />
                    <br />
                    <asp:Repeater runat="server" Visible="False" ID="Repeater2">
                        <ItemTemplate>
                            &nbsp;&nbsp;&nbsp;<asp:LinkButton ID="LinkButton6" runat="server" Style="text-decoration: none"
                                OnCommand="lnkbtnShowProductsByCategory_click" CommandArgument='<%# Container.DataItem("SubCategory")%>'
                                Text='<%# Container.DataItem("SubCategory")%>' ForeColor="Navy" Font-Size="X-Small"
                                EnableTheming="false" />
                            <br />
                        </ItemTemplate>
                    </asp:Repeater>
                    <br />
                </td>
                <td style="width: 2%">
                </td>
            </tr>
        </table>
    </asp:Panel>
    <asp:Panel ID="pnlCategorySelection2" runat="server" Visible="True" Width="100%">
        <table id="tblCategorySelection2" runat="server" width="100%" style="font-family: Verdana;
            font-size: small" cellpadding="2" cellspacing="1">
            <tr>
                <td style="width: 2%">
                </td>
                <td valign="top" style="white-space: nowrap; width: 32%; background-color: #DFE5E2">
                    &nbsp;&nbsp;<asp:Label ID="Label93" runat="server" ForeColor="Navy" Font-Bold="True"
                        Font-Size="X-Small">Product Category</asp:Label>
                    <br />
                    <br />
                    <asp:Repeater runat="server" ID="Repeater1a" OnItemCommand="repeater1_Item_click">
                        <ItemTemplate>
                            &nbsp;&nbsp;&nbsp;<asp:LinkButton ID="LinkButton1" runat="server" Style="text-decoration: none"
                                OnCommand="lnkbtnShowSubCategories_click" CommandArgument='<%# Container.DataItem("Category")%>'
                                Text='<%# Container.DataItem("Category")%>' ForeColor="Navy" Font-Size="X-Small"
                                EnableTheming="false" />
                            <br />
                        </ItemTemplate>
                    </asp:Repeater>
                    <br />
                </td>
                <td valign="top" style="white-space: nowrap; width: 32%; background-color: #DFE5E2">
                    &nbsp;&nbsp;<asp:Label runat="server" ID="lblSubCategoryHeadingA" ForeColor="Navy"
                        Font-Bold="True" Font-Size="X-Small">Sub Category 1</asp:Label>
                    <br />
                    <br />
                    <asp:Repeater runat="server" Visible="False" ID="Repeater2a" OnItemCommand="repeater1_Item_click">
                        <ItemTemplate>
                            &nbsp;&nbsp;&nbsp;<asp:LinkButton ID="LinkButton3" runat="server" Style="text-decoration: none"
                                OnCommand="lnkbtnShowSubSubCategories_click" CommandArgument='<%# Container.DataItem("SubCategory")%>'
                                Text='<%# Container.DataItem("SubCategory")%>' ForeColor="Navy" Font-Size="X-Small"
                                EnableTheming="false" />
                            <br />
                        </ItemTemplate>
                    </asp:Repeater>
                    <br />
                </td>
                <td valign="top" style="white-space: nowrap; width: 32%; background-color: #DFE5E2">
                    &nbsp;&nbsp;<asp:Label runat="server" ID="lblSubCategoryHeadingB" ForeColor="Navy"
                        Font-Bold="True" Font-Size="X-Small">Sub Category 2</asp:Label>
                    <br />
                    <br />
                    <asp:Repeater runat="server" Visible="False" ID="Repeater3a">
                        <ItemTemplate>
                            &nbsp;&nbsp;&nbsp;<asp:LinkButton ID="LinkButton4" runat="server" Style="text-decoration: none"
                                OnCommand="lnkbtnShowProductsByCategory_click" CommandArgument='<%# Container.DataItem("SubCategory2")%>'
                                Text='<%# Container.DataItem("SubCategory2")%>' ForeColor="Navy" Font-Size="X-Small"
                                EnableTheming="false" />
                            <br />
                        </ItemTemplate>
                    </asp:Repeater>
                    <br />
                </td>
                <td style="width: 2%">
                </td>
            </tr>
        </table>
    </asp:Panel>
    <asp:Panel ID="pnlProductList" runat="server" Visible="False" Width="100%">
        <asp:DataGrid ID="dg_ProductList" runat="server" Width="100%" Font-Names="Verdana"
            Font-Size="XX-Small" PageSize="6" OnPageIndexChanged="dg_ProductList_Page_Change"
            AllowPaging="True" Visible="False" AutoGenerateColumns="False" GridLines="None"
            ShowFooter="True" OnItemCommand="dg_ProductList_item_click">
            <FooterStyle Wrap="False"></FooterStyle>
            <HeaderStyle Font-Names="Verdana" Wrap="False"></HeaderStyle>
            <PagerStyle NextPageText="Next Page  " Font-Size="X-Small" Font-Names="Verdana" Font-Bold="True"
                PrevPageText="Previous Page" HorizontalAlign="Center" ForeColor="Blue" Position="Top"
                BackColor="Silver" Wrap="False" Mode="NumericPages"></PagerStyle>
            <Columns>
                <asp:BoundColumn Visible="False" DataField="LogisticProductKey">
                    <ItemStyle Wrap="False"></ItemStyle>
                </asp:BoundColumn>
                <asp:TemplateColumn>
                    <ItemStyle HorizontalAlign="Left"></ItemStyle>
                    <ItemTemplate>
                        <asp:Table ID="tabProductList" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Width="100%" ForeColor="Gray">
                            <asp:TableRow>
                                <asp:TableCell RowSpan="4" Width="7%" VerticalAlign="Top">
                                    <asp:HyperLink ID="hlnk_ThumbNail" runat="server" ToolTip="click here to see larger image"
                                        NavigateUrl='<%# "Javascript:ShowImage(""" & DataBinder.Eval(Container.DataItem,"OriginalImage") & """)" %> '
                                        ImageUrl='<%# psVirtualThumbFolder & DataBinder.Eval(Container.DataItem,"ThumbNailImage") & "?" & Now.ToString %>'></asp:HyperLink>
                                </asp:TableCell>
                                <asp:TableCell Width="12%" VerticalAlign="Top" Wrap="False">
                                    <asp:Label ID="Label5" runat="server" ForeColor="Gray">Product Code:</asp:Label>
                                </asp:TableCell>
                                <asp:TableCell Width="15%" VerticalAlign="Top" Wrap="False">
                                    <asp:Label ID="Label4" runat="server" ForeColor="Red"><%# DataBinder.Eval(Container.DataItem,"ProductCode") %></asp:Label>
                                </asp:TableCell>
                                <asp:TableCell Width="12%" VerticalAlign="Top" Wrap="False">
                                    <asp:Label ID="Label6" runat="server" ForeColor="Gray" Text="<%# GetProductDateLegend() %>" />
                                </asp:TableCell><asp:TableCell Width="15%" VerticalAlign="Top">
                                    <asp:Label ID="Label7" runat="server" ForeColor="Red"><%# DataBinder.Eval(Container.DataItem,"ProductDate") %></asp:Label>
                                </asp:TableCell><asp:TableCell Width="27%"></asp:TableCell><asp:TableCell Width="12%"
                                    VerticalAlign="Top" HorizontalAlign="Right" Wrap="False">
                                    <asp:Label ID="Label8" runat="server" ForeColor="Gray">Quantity:</asp:Label>
                                    &nbsp;<asp:Label ID="Label9" runat="server" ForeColor="Navy"><%# Format(DataBinder.Eval(Container.DataItem,"Quantity"),"#,##0") %></asp:Label>
                                </asp:TableCell></asp:TableRow><asp:TableRow>
                                <asp:TableCell VerticalAlign="Top">
                                    <asp:Label ID="Label10" runat="server" ForeColor="Gray">Category:</asp:Label>
                                </asp:TableCell><asp:TableCell VerticalAlign="Top">
                                    <asp:Label ID="Label11" runat="server" ForeColor="Navy"><%# DataBinder.Eval(Container.DataItem,"ProductCategory") %></asp:Label>
                                </asp:TableCell><asp:TableCell VerticalAlign="Top" Wrap="False">
                                    <asp:Label ID="Label12" runat="server" ForeColor="Gray">Sub Category:</asp:Label>
                                </asp:TableCell><asp:TableCell VerticalAlign="Top">
                                    <asp:Label ID="Label13" runat="server" ForeColor="Navy"><%# DataBinder.Eval(Container.DataItem,"SubCategory") %></asp:Label>
                                </asp:TableCell><asp:TableCell VerticalAlign="Top" HorizontalAlign="Right">
                                    <asp:Label ID="Label16" runat="server" ForeColor="Gray" Text="<%# GetArchiveFlagLegend() %>" />
                                </asp:TableCell><asp:TableCell VerticalAlign="Top">
                                    <asp:Label ID="Label17" runat="server" ForeColor="Navy"><%# DataBinder.Eval(Container.DataItem,"ArchiveFlag") %></asp:Label>
                                </asp:TableCell></asp:TableRow><asp:TableRow>
                                <asp:TableCell VerticalAlign="Top" Wrap="False">
                                    <asp:Label ID="Label14" runat="server" ForeColor="Gray">Description:</asp:Label>
                                </asp:TableCell><asp:TableCell VerticalAlign="Top" ColumnSpan="4" RowSpan="2">
                                    <asp:Label ID="Label15" runat="server" ForeColor="Navy"><%# DataBinder.Eval(Container.DataItem,"ProductDescription") %></asp:Label>
                                </asp:TableCell><asp:TableCell VerticalAlign="Bottom" HorizontalAlign="Right" RowSpan="2">
                                    <asp:Button ID="EditProduct" runat="server" CommandName="Edit" Text="edit this product"
                                        ToolTip="edit this product" />
                                </asp:TableCell></asp:TableRow><asp:TableRow>
                                <asp:TableCell VerticalAlign="Top">                                      
                                </asp:TableCell></asp:TableRow><asp:TableRow>
                                <asp:TableCell ColumnSpan="8" VerticalAlign="Top">
                                        <hr />
                                </asp:TableCell></asp:TableRow></asp:Table></ItemTemplate></asp:TemplateColumn></Columns></asp:DataGrid><asp:Label ID="Label72" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
            ForeColor="Gray">Items per page:</asp:Label>&nbsp;<asp:DropDownList ID="ddlItemsPerPage"
                runat="server" AutoPostBack="True" Font-Names="Verdana" Font-Size="XX-Small"
                OnSelectedIndexChanged="ddlItemsPerPage_SelectedIndexChanged">
                <asp:ListItem Selected="True">6</asp:ListItem><asp:ListItem>20</asp:ListItem><asp:ListItem>50</asp:ListItem></asp:DropDownList><asp:Label ID="lblProductMessage" runat="server" Font-Names="Verdana" Font-Size="X-Small"
            ForeColor="Gray"></asp:Label></asp:Panel><asp:Panel ID="pnlEditProduct" runat="server" Visible="False" Width="100%">
        <table id="table1xx" width="100%" style="font-family: Verdana; font-size: x-small">
            <tr valign="middle">
                <td style="white-space: nowrap; width: 40%; height: 26px;">
                    <asp:Label ID="lblLegendProductDetail" runat="server" Font-Size="X-Small" Font-Names="Verdana"
                        Font-Bold="True" ForeColor="Gray">Product Detail: </asp:Label>&nbsp;<asp:Label runat="server"
                            ID="lblProductQuantity" Font-Size="X-Small" Font-Names="Verdana" ForeColor="Red">
                        </asp:Label><asp:Label ID="lblLegendItemsInStock" runat="server" Font-Size="X-Small"
                            Font-Names="Verdana" ForeColor="Gray"> items in stock.</asp:Label></td><td align="right" style="white-space: nowrap; width: 60%; height: 26px;">
                    <asp:LinkButton ID="lnkbtnShowHelp" runat="server" OnClick="lnkbtnShowHelp_Click"
                        CausesValidation="False">hide help</asp:LinkButton>&nbsp; &nbsp;<asp:Button ID="btnAssociatedProducts"
                            runat="server" OnClick="btnAssociatedProducts_Click" Text="associated products..." />
                    &nbsp;&nbsp; <asp:Button ID="btnSetUserProfiles" runat="server" OnClick="btnSetUserProfiles_click"
                        Text="set max order levels..." ToolTip="set max order levels" />
                    &nbsp;&nbsp;&nbsp;&nbsp;<asp:Button ID="btn_DeleteProduct" runat="server" OnClick="btn_DeleteProduct_click"
                        Text="delete product" OnClientClick='return confirm("Are you sure you want to delete this product?");'
                        ToolTip="delete this product" /><a runat="server" id="aHelpDeleteProduct" visible="false"
                            onmouseover="return escape('Click this button to remove the product completely. As a precaution against accidental deletion you must click the OK button when asked \'Are you sure?\'. To be deleted the product must have a stock level of 0 (zero). You can restore a deleted product using the UnDelete facility (not currently enabled).')"
                            style="color: gray; cursor: help; font-size: xx-small">&nbsp;?&nbsp;</a> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:Button ID="btn_GoToProductListPanel" runat="server"
                        OnClick="btn_GoToProductListPanel_click" CausesValidation="false" Text="return to list"
                        ToolTip="go back to product list" />
                </td>
            </tr>
        </table>
        <br />
        <table style="width: 100%">
            <tr>
                <td style="width: 12%; white-space: nowrap">
                </td>
                <td style="width: 25%; white-space: nowrap">
                </td>
                <td style="width: 12%; white-space: nowrap">
                </td>
                <td style="width: 24%; white-space: nowrap">
                </td>
                <td style="width: 15%; white-space: nowrap">
                </td>
                <td style="width: 12%; white-space: nowrap">
                </td>
            </tr>
            <tr>
                <td colspan="6" align="right">
                    <asp:Label runat="server" ForeColor="#00C000" ID="lblEditDateError" Font-Size="X-Small"
                        Font-Names="Verdana"></asp:Label></td></tr><tr>
                <td align="right" style="white-space: nowrap">
                    <asp:RequiredFieldValidator ID="rfdProductCode" runat="server" ControlToValidate="txtProductCode"
                        Font-Size="XX-Small" Text="#" />
                    &nbsp; <asp:Label ID="lblLegendProdCode" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        Text="Prod Code:" />
                </td>
                <td style="white-space: nowrap">
                    <asp:TextBox ID="txtProductCode" MaxLength="25" runat="server" ForeColor="Navy" Width="150"
                        TabIndex="1" Font-Size="XX-Small" Font-Names="Verdana"></asp:TextBox><a runat="server"
                            id="aHelpProductCode" visible="false" onmouseover="return escape('<b>Product Code</b> (maximum length 25 chars) when combined with <b>Product Date</b> (sometimes called <b>Version Date</b>) uniquely identifies this product.')"
                            style="color: gray; cursor: help">&nbsp;?&nbsp;</a> <asp:LinkButton ID="lnkbtnNewProductCode" Visible="false" runat="server" CausesValidation="False"
                        Font-Names="Verdana" Font-Size="XX-Small" OnClick="lnkbtnNewProductCode_Click">new code</asp:LinkButton><a
                            runat="server" id="aHelpNewCode" visible="false" onmouseover="return escape('Click <b>new code</b> to assign a new, unique product code to this product.<br /><br />If you require a new, unique product code but you do not want to create the product at this time (eg because the Version Date is not yet available) use the <b>reserve W-code</b> facility.<br /><br />When you subsequently create the product, enter the product code you reserved. That code will then be removed from your reservation list.')"
                            style="color: gray; cursor: help">&nbsp;?&nbsp;</a> </td><td align="right" style="white-space: nowrap">
                    <asp:RequiredFieldValidator ID="rfdProductDate" runat="server" ControlToValidate="txtProductDate"
                        Enabled="false" Font-Size="XX-Small" Text="#" />
                    &nbsp; <asp:Label ID="lblLegendProductDate" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        Text="Product Date:" />
                </td>
                <td>
                    <asp:TextBox ID="txtProductDate" MaxLength="10" runat="server" ForeColor="Navy" Width="100"
                        TabIndex="2" Font-Size="XX-Small" Font-Names="Verdana"></asp:TextBox><a runat="server"
                            id="aHelpProductDate" visible="false" onmouseover="return escape('<b>Product Date</b> (sometimes called <b>Version Date</b>) when combined with <b>Product Code</b> uniquely identifies this product. Use this field to identify a specific version or variant of a product, the versions of which share the same <b>Product Code</b>. Maximum length 10 chars.')"
                            style="color: gray; cursor: help">&nbsp;?&nbsp;</a> </td><td align="right" style="white-space: nowrap">
                    <asp:RegularExpressionValidator ID="revMinStockLevel" runat="server" ControlToValidate="txtMinStockLevel"
                        Enabled="False" Font-Size="XX-Small" ValidationExpression="[123456789]\d*">#</asp:RegularExpressionValidator><asp:RequiredFieldValidator
                            ID="rfdMinStockLevel" runat="server" ControlToValidate="txtMinStockLevel" Font-Size="XX-Small"
                            Text="#" Enabled="False" />
                    &nbsp; <asp:Label ID="lblLegendMinStockLevel" runat="server" Font-Size="XX-Small" Font-Names="Verdana">Min Stock Level:</asp:Label></td><td>
                    <asp:TextBox ID="txtMinStockLevel" runat="server" ForeColor="Navy" Width="50" TabIndex="3"
                        Font-Size="XX-Small" Font-Names="Verdana" MaxLength="6"></asp:TextBox><a runat="server"
                            id="aHelpMinStockLevel" visible="false" onmouseover="return escape('The system sends an email alert when the available stock quantity falls to (or below) this level. In some installations this field is mandatory. If this field is not mandatory, you can set the value to 0 to disable <b>Low Stock</b> email alerts.')"
                            style="color: gray; cursor: help">&nbsp;?&nbsp;</a> </td></tr><tr>
                <td align="right" style="white-space: nowrap">
                    <asp:RequiredFieldValidator ID="rfdDescription" runat="server" ControlToValidate="txtDescription"
                        Font-Size="XX-Small" Text="#" Enabled="False" />
                    &nbsp; <asp:Label ID="lblLegendDescription" runat="server" Font-Size="XX-Small" Font-Names="Verdana">Description:</asp:Label></td><td colspan="3">
                    <asp:TextBox ID="txtDescription" MaxLength="300" runat="server" ForeColor="Navy"
                        Width="470px" TabIndex="4" Font-Size="XX-Small" Font-Names="Verdana"></asp:TextBox><a
                            runat="server" id="aHelpDescription" visible="false" onmouseover="return escape('Description of the product. Maximum length 300 characters.')"
                            style="color: gray; cursor: help">&nbsp;?&nbsp;</a> </td><td align="right" style="white-space: nowrap">
                    <asp:Label ID="lblLegendItemsPerBox" runat="server" Font-Size="XX-Small" Font-Names="Verdana">Items Per Box:</asp:Label></td><td>
                    <asp:TextBox ID="txtItemsPerBox" runat="server" ForeColor="Navy" Width="50" TabIndex="5"
                        Font-Size="XX-Small" Font-Names="Verdana" MaxLength="6"></asp:TextBox><a runat="server"
                            id="aHelpItemsPerBox" visible="false" onmouseover="return escape('The number of individual pieces of this product in a box (or other container). Typically used to give guidance on preferred order quantities. Unless explicitly stated, the available stock quantity refers to the number of individual pieces of a product, not the number of boxes.')"
                            style="color: gray; cursor: help">&nbsp;?&nbsp;</a> </td></tr><tr>
                <td align="right" style="white-space: nowrap; height: 46px;">
                    <asp:Label ID="lblLegendCategory" runat="server" Font-Size="XX-Small" Font-Names="Verdana">Category:</asp:Label></td><td style="height: 46px">
                    <asp:HiddenField ID="hidCategory" runat="server"></asp:HiddenField>
                    <asp:TextBox ID="txtCategory" Visible="False" MaxLength="50" runat="server" ForeColor="Navy"
                        TabIndex="6" Width="150" Font-Size="XX-Small" Font-Names="Verdana" />
                    <asp:DropDownList ID="ddlCategory" runat="server" Visible="False" DataSourceID="SqlDataSourceCategoryList"
                        DataTextField="Category" DataValueField="Category" AutoPostBack="True" EnableViewState="false"
                        Font-Size="XX-Small" Font-Names="Verdana" Width="150" OnSelectedIndexChanged="ddlCategory_SelectedIndexChanged"
                        OnDataBound="ddlCategory_DataBound">
                    </asp:DropDownList>
                    <a runat="server" id="aHelpCategory" visible="false" onmouseover="return escape('The top level category for this product. Add a new category by clicking <b>- new category -</b> then entering the new category name. If you change the top level category of a product, you must then set the sub category.')"
                        style="color: gray; cursor: help">&nbsp;?&nbsp;</a> </td><td align="right" style="white-space: nowrap; height: 46px;">
                    <asp:Label ID="lblLegendProductUsersMessage" runat="server" Font-Size="XX-Small"
                        Font-Names="Verdana">Sub Category:</asp:Label></td><td style="height: 46px">
                    <asp:HiddenField ID="hidSubCategory" runat="server"></asp:HiddenField>
                    <asp:TextBox ID="txtSubCategory" Visible="False" MaxLength="50" runat="server" ForeColor="Navy"
                        TabIndex="7" Width="150" Font-Size="XX-Small" Font-Names="Verdana" />
                    <asp:DropDownList ID="ddlSubCategory" runat="server" Visible="False" DataSourceID="SqlDataSourceSubCategoryList"
                        DataTextField="SubCategory" DataValueField="SubCategory" AutoPostBack="True"
                        EnableViewState="false" Font-Size="XX-Small" Font-Names="Verdana" Width="150"
                        OnSelectedIndexChanged="ddlSubCategory_SelectedIndexChanged" OnDataBound="ddlSubCategory_DataBound">
                    </asp:DropDownList>
                    <a runat="server" id="aHelpSubCategory" visible="false" onmouseover="return escape('The 2nd level (sub) category for this product.  Add a new sub category by clicking <b>- new subcategory -</b> then entering the new sub category name'). If you are using a further sub category level, you must then set the further (final) sub category."
                        style="color: gray; cursor: help">&nbsp;?&nbsp;</a> </td><td align="right" style="white-space: nowrap; height: 46px;">
                    <asp:Label ID="lblLegendArchiveFlag" runat="server" Font-Size="XX-Small" Font-Names="Verdana">Archive Flag:</asp:Label></td><td style="height: 46px">
                    <asp:CheckBox runat="server" ID="chkArchivedFlag" TabIndex="14" Font-Names="Verdana"
                        Font-Size="XX-Small"></asp:CheckBox>
                    <a runat="server" id="aHelpArchived" visible="false" onmouseover="return escape('Controls whether this product is shown on the Orders page')"
                        style="color: gray; cursor: help">&nbsp;?&nbsp;</a> </td></tr><tr id="trSubCategory2" runat="server" visible="false">
                <td align="right">
                </td>
                <td>
                </td>
                <td align="right">
                    <asp:Label ID="lblLegendSubCategory2" runat="server" Font-Size="XX-Small" Font-Names="Verdana">Sub Category 2:</asp:Label></td><td>
                    <asp:HiddenField ID="hidSubSubCategory" runat="server"></asp:HiddenField>
                    <asp:TextBox ID="tbSubSubCategory" Visible="False" MaxLength="50" runat="server"
                        ForeColor="Navy" TabIndex="7" Width="150" Font-Size="XX-Small" Font-Names="Verdana" />
                    <asp:DropDownList ID="ddlSubSubCategory" runat="server" Visible="True" DataSourceID="SqlDataSourceSubSubCategoryList"
                        DataTextField="SubCategory2" DataValueField="SubCategory2" AutoPostBack="True"
                        EnableViewState="false" Font-Size="XX-Small" Font-Names="Verdana" Width="150"
                        OnSelectedIndexChanged="ddlSubSubCategory_SelectedIndexChanged" OnDataBound="ddlSubSubCategory_DataBound">
                    </asp:DropDownList>
                    <a runat="server" id="aHelpSubSubCategory" visible="false" onmouseover="return escape('The 3rd level (sub) category for this product.  Add a new sub category by clicking <b>- new subcategory -</b> then entering the new sub category name').')"
                        style="color: gray; cursor: help">&nbsp;?&nbsp;</a> </td><td align="right">
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td align="right" style="white-space: nowrap; height: 22px;">
                    <asp:Label ID="lblLegendLanguage" runat="server" Font-Size="XX-Small" Font-Names="Verdana">Language:</asp:Label><asp:Label
                        ID="lblLegendStatus" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Visible="False">Status:</asp:Label></td><td style="height: 22px">
                    <asp:TextBox ID="txtLanguage" MaxLength="50" runat="server" ForeColor="Navy" TabIndex="9"
                        Width="150" Font-Size="XX-Small" Font-Names="Verdana"></asp:TextBox><a runat="server"
                            id="aHelpLanguage" visible="false" onmouseover="return escape('The language of this product')"
                            style="color: gray; cursor: help">&nbsp;?&nbsp;</a> <asp:DropDownList ID="ddlStatus" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Visible="False" OnSelectedIndexChanged="ddlStatus_SelectedIndexChanged" AutoPostBack="True">
                        <asp:ListItem Selected="True" Value="0">Active</asp:ListItem><asp:ListItem Value="1">On Hold</asp:ListItem><asp:ListItem Value="2">Obsolete</asp:ListItem><asp:ListItem Value="3">Replaced</asp:ListItem><asp:ListItem Value="4">Not Held At Sprint</asp:ListItem></asp:DropDownList><a runat="server" id="aHelpStatus" visible="false" onmouseover="return escape('Indicates the status of this product and controls whether it can be ordered, as follows:<br /><br /><b></b><b>ACTIVE:</b> Product available to be ordered<br /><b>ON HOLD:</b> Product appears on Orders page but cannot be ordered<br /><b>OBSOLETE:</b> Product appears on Orders page for information only; cannot be ordered<br /><b>REPLACED:</b> Product appears on Orders page for information only; cannot be ordered<br /><b>NOT HELD AT SPRINT:</b> Product appears on Product Manager page only. Selecting this value sets the <b>Hide From View</b> check box so that product does not appear on the Orders page. If the <b>Hide From View</b> attribute is subsequently removed, the product will appear on the Orders page but cannot be ordered<br /><br />To remove a product from the Orders page, use the <b>Hide From View</b> check box')"
                        style="color: gray; cursor: help">&nbsp;?&nbsp;</a> </td><td align="right" style="white-space: nowrap; height: 22px;">
                    <asp:RequiredFieldValidator ID="rfdCostCentre" runat="server" ControlToValidate="txtDepartment"
                        Font-Size="XX-Small" Text="#" Enabled="False" />
                    &nbsp; <asp:Label ID="lblLegendCostCentre" runat="server" Font-Size="XX-Small" Font-Names="Verdana">Cost Centre:</asp:Label></td><td style="height: 22px">
                    <asp:TextBox ID="txtDepartment" MaxLength="50" runat="server" ForeColor="Navy" TabIndex="10"
                        Width="150" Font-Size="XX-Small" Font-Names="Verdana"></asp:TextBox><a runat="server"
                            id="aHelpDepartment" visible="false" onmouseover="return escape('The cost centre associated with this product')"
                            style="color: gray; cursor: help">&nbsp;?&nbsp;</a> </td><td align="right" style="white-space: nowrap; height: 22px;">
                    <asp:Label ID="lblLegendUnitWeight" runat="server" Font-Size="XX-Small" Font-Names="Verdana">Unit Weight (gm):</asp:Label></td><td style="height: 22px">
                    <asp:TextBox ID="txtUnitWeight" runat="server" ForeColor="Navy" TabIndex="11" Width="50"
                        Font-Size="XX-Small" Font-Names="Verdana" MaxLength="6"></asp:TextBox><a id="aHelpUnitWeight"
                            runat="server" onmouseover="return escape('The weight in grams of a single unit of this product')"
                            style="color: gray; cursor: help" visible="false">&nbsp;?&nbsp;</a> </td></tr><tr>
                <td align="right" style="white-space: nowrap">
                    <asp:Label ID="lblLegendMisc1" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        Text="Misc 1:" />
                </td>
                <td>
                    <asp:TextBox ID="txtMisc1" MaxLength="50" runat="server" ForeColor="Navy" TabIndex="12"
                        Width="150" Font-Size="XX-Small" Font-Names="Verdana"></asp:TextBox><a runat="server"
                            id="aHelpMisc1" visible="false" onmouseover="return escape('This field is available to store additional data of your choice. Data in this field will be checked when using the <b>search</b> facility. Maximum length 50 characters.')"
                            style="color: gray; cursor: help">&nbsp;?&nbsp;</a> </td><td align="right" style="white-space: nowrap">
                    <asp:RegularExpressionValidator ID="revFEXCOCriticalProduct" runat="server" ControlToValidate="txtMisc2"
                        Enabled="False" Font-Size="XX-Small" ValidationExpression="[YyNn]">#</asp:RegularExpressionValidator>&nbsp; <asp:Label ID="lblLegendMisc2" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        Text="Misc 2:" />
                </td>
                <td>
                    <asp:TextBox ID="txtMisc2" MaxLength="50" runat="server" ForeColor="Navy" TabIndex="13"
                        Width="150" Font-Size="XX-Small" Font-Names="Verdana"></asp:TextBox><a runat="server"
                            id="aHelpMisc2" visible="false" onmouseover="return escape('This field is available to store additional data of your choice. Data in this field will be checked when using the <b>search</b> facility. Maximum length 50 characters.')"
                            style="color: gray; cursor: help">&nbsp;?&nbsp;</a> </td><td align="right" style="white-space: nowrap">
                    <asp:Label ID="lblLegendUnitValue" runat="server" Font-Names="Verdana" Font-Size="XX-Small">Unit Value (&pound;):</asp:Label>&nbsp; </td><td>
                    <asp:TextBox ID="txtUnitValue" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        ForeColor="Navy" TabIndex="8" Width="50" MaxLength="6"></asp:TextBox><a runat="server"
                            id="aHelpUnitValue" visible="false" onmouseover="return escape('The value, or cost price, of a single item or unit of this product')"
                            style="color: gray; cursor: help">&nbsp;?&nbsp;</a> </td></tr>
                        <tr>
                <td align="right" style="white-space: nowrap; height: 37px;">
                    <asp:Label ID="lblLegendExpiryDate" runat="server" Font-Size="XX-Small" Font-Names="Verdana">Expiry Date:</asp:Label></td><td style="height: 37px">
                    <asp:TextBox ID="tbExpiryDate" Font-Names="Verdana" Font-Size="XX-Small" Width="90"
                        runat="server" />
                    <a href="javascript:;" onclick="window.open('PopupCalendar4.aspx?textbox=tbExpiryDate','cal','width=300,height=305,left=270,top=180')">
                        <img id="Img1" alt="" src="~/images/SmallCalendar.gif" runat="server" border="0"
                            ie:visible="true" visible="false" /></a> <a runat="server" id="aHelpExpiryDate" visible="false"
                                onmouseover="return escape('When the Expiry Date is reached or passed, the system periodically sends an email alert. To activate this function click the calendar icon (available in Internet Explorer only) and follow the instructions to select a date , or type the date directly in the format dd-mmm-yyyy, eg 29-Jan-2006 ')"
                                style="color: gray; cursor: help">&nbsp;?&nbsp;</a> </td><td align="right" style="white-space: nowrap; height: 37px;">
                    <asp:Label ID="lblLegendRenewalDate" runat="server" Font-Size="XX-Small" Font-Names="Verdana">Renewal Date:</asp:Label></td><td style="height: 37px">
                    <asp:TextBox ID="tbReplenishmentDate" Font-Names="Verdana" Font-Size="XX-Small" Width="90"
                        runat="server" />
                    <a href="javascript:;" onclick="window.open('PopupCalendar4.aspx?textbox=tbReplenishmentDate','cal','width=300,height=305,left=270,top=180')">
                        <img id="SmallCalendar" alt="" src="~/images/SmallCalendar.gif" runat="server" border="0"
                            ie:visible="true" visible="false" /></a> <a runat="server" id="aHelpReplenishmentDate"
                                visible="false" onmouseover="return escape('When the Renewal (sometimes called Review or Replenishment) Date is reached or passed, the system periodically sends an email alert. To activate this function click the calendar icon (available in Internet Explorer only) and follow the instructions to select a date, or type the date directly in the format dd-mmm-yyyy, eg 29-Jan-2006 ')"
                                style="color: gray; cursor: help">&nbsp;?&nbsp;</a> </td><td align="right" style="white-space: nowrap; height: 37px;">
                    <asp:Label ID="lblLegendSerialNumbers" runat="server" Font-Size="XX-Small" Font-Names="Verdana">Serial Numbers:</asp:Label></td><td id="CHECK_THIS_FIELD" style="height: 37px">
                    <asp:CheckBox runat="server" ID="chkProspectusNumbers" TabIndex="21" Font-Names="Verdana"
                        Font-Size="XX-Small"></asp:CheckBox>
                    <a runat="server" id="aHelpProspectusNumbers" visible="false" onmouseover="return escape('Controls whether serial numbers are used for this product. Serial numbers are typically used to track products for auditing.')"
                        style="color: gray; cursor: help">&nbsp;?&nbsp;</a> </td></tr><tr>
                <td align="right" style="white-space: nowrap">
                </td>
                <td>
                </td>
                <td align="right" style="white-space: nowrap">
                </td>
                <td>
                    
                </td>
                <td align="right" style="white-space: nowrap">
                    <asp:Label ID="lblLegendInactivityAlert" runat="server" Font-Size="XX-Small" Font-Names="Verdana">Product Inactivity Alert:</asp:Label></td><td>
                    <asp:TextBox ID="tbInactivityAlertDays" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        ForeColor="Navy" TabIndex="3" Width="50" MaxLength="3" />
                    <asp:LinkButton ID="lnkbtnConfigureProductInactivityAlerts" runat="server" Font-Names="Verdana"
                        Font-Size="XX-Small" OnClick="lnkbtnConfigureProductInactivityAlerts_Click" CausesValidation="False">config</asp:LinkButton><a
                            runat="server" id="aHelpInactivityAlert" visible="false" onmouseover="return escape('Sets the period, in days, after which an email alert will be sent if no orders have been received for this product. A value of zero (0) disables this feature.<br/><br/>Click the <b>config</b> button to configure Inactivity Alerts for all products.')"
                            style="color: gray; cursor: help">&nbsp;?&nbsp;</a> </td></tr><tr>
                <td align="right" style="white-space: nowrap">
                </td>
                <td>
                </td>
                <td align="right" style="white-space: nowrap">
                </td>
                <td>
                </td>
                <td align="right" style="white-space: nowrap">
                </td>
                <td>
                </td>
            </tr>
            <tr runat="server" id="trAdRotatorControls" visible="true">
                <td align="right" style="white-space: nowrap">
                    <asp:Label ID="lblLegendAdRotatorText" runat="server" Font-Size="XX-Small" Font-Names="Verdana">AdRotator Txt:</asp:Label></td><td valign="top" colspan="3">
                    <asp:TextBox ID="txtAdRotatorText" MaxLength="120" runat="server" ForeColor="Navy"
                        TabIndex="22" Width="420px" Font-Size="XX-Small" Font-Names="Verdana"></asp:TextBox><a
                            runat="server" id="aHelpAdRotatorText" visible="false" onmouseover="return escape('The text to appear in the \'advertisement\' above the tabs in the browser window')"
                            style="color: gray; cursor: help">&nbsp;?&nbsp;</a> </td><td align="right" style="white-space: nowrap">
                    <asp:Label ID="lblLegendAdRotator" runat="server" Font-Size="XX-Small" Font-Names="Verdana">Ad Rotator:</asp:Label></td><td>
                    <asp:CheckBox runat="server" ID="chkAdRotator" TabIndex="23" Font-Names="Verdana"
                        Font-Size="XX-Small"></asp:CheckBox>
                    <a runat="server" id="aHelpAdRotator" visible="false" onmouseover="return escape('Controls whether \'advertisements\' are displayed above the tabs in the browser window')"
                        style="color: gray; cursor: help">&nbsp;?&nbsp;</a> </td></tr><tr runat="server" id="trLandGControls" visible="false">
                <td align="right" style="white-space: nowrap; height: 22px;">
                    <asp:RequiredFieldValidator ID="rfdSupplier" runat="server" ControlToValidate="tbSupplier"
                        Font-Size="XX-Small" Text="#" Enabled="False" />&nbsp; <asp:Label ID="lblLegendSupplier" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Red">Supplier:</asp:Label></td><td style="height: 22px">
                    <asp:TextBox ID="tbSupplier" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        ForeColor="Navy" TabIndex="3" Width="150px" MaxLength="40"></asp:TextBox><a runat="server"
                            id="aHelpSupplier" visible="false" onmouseover="return escape('The supplier name for this product')"
                            style="color: gray; cursor: help">&nbsp;?&nbsp;</a> </td><td align="right" style="white-space: nowrap; height: 22px;">
                    <asp:Label ID="lblLegendBPCanView" runat="server" Font-Size="XX-Small" Font-Names="Verdana">BP Can View:</asp:Label></td><td valign="top" style="height: 22px">
                    <asp:CheckBox runat="server" ID="cbBPCanView" Font-Names="Verdana" Font-Size="XX-Small">
                    </asp:CheckBox>
                    <a runat="server" id="aHelpBPCanView" visible="false" onmouseover="return escape('Controls whether BPs can see this product on their Order page. Note that the <b>Hide From View</b> tick box overrides this setting.  That is, if <b>Hide From View</b> is ticked, this product cannot be ordered by anyone.')"
                        style="color: gray; cursor: help">&nbsp;?&nbsp;</a> </td><td align="right" style="white-space: nowrap; height: 22px;">
                    <asp:Label ID="lblLegendIFACanView" runat="server" Font-Size="XX-Small" Font-Names="Verdana">IFA Can View:</asp:Label>&nbsp; </td><td style="height: 22px">
                    <asp:CheckBox runat="server" ID="cbIFACanView" Font-Names="Verdana" Font-Size="XX-Small">
                    </asp:CheckBox>
                    <a runat="server" id="aHelpIFACanView" visible="false" onmouseover="return escape('Controls whether IFAs can see this product on their Order page. Note that the <b>Hide From View</b> tick box overrides this setting.  That is, if <b>Hide From View</b> is ticked, this product cannot be ordered by anyone.')"
                        style="color: gray; cursor: help">&nbsp;?&nbsp;</a> </td></tr><tr>
                <td align="right" valign="top" style="white-space: nowrap; height: 22px;">
                    <asp:Label ID="lblLegendComments" runat="server" Font-Size="XX-Small" Font-Names="Verdana">Comments:</asp:Label></td><td colspan="3" rowspan="2">
                    <asp:TextBox ID="txtNotes" MaxLength="1000" TextMode="MultiLine" Rows="3" Width="420"
                        runat="server" ForeColor="Navy" TabIndex="24" Font-Size="XX-Small" Font-Names="Verdana"></asp:TextBox><a
                            runat="server" id="aHelpNotes" visible="false" onmouseover="return escape('Additional notes on this product. Depending on your installation options, these notes may appear on the Order page to convey additional information about the product to orderers.')"
                            style="color: gray; cursor: help">&nbsp;?&nbsp;</a> </td><td align="right" style="white-space: nowrap; height: 22px;">
                    <asp:Label ID="lblLegendViewOnWebForm" runat="server" Font-Size="XX-Small" Font-Names="Verdana">View on Web Form:</asp:Label>&nbsp; </td><td style="height: 22px">
                    <asp:CheckBox runat="server" ID="chkViewOnWebForm" TabIndex="25" Font-Names="Verdana"
                        Font-Size="XX-Small"></asp:CheckBox>
                    <a runat="server" id="aHelpViewOnWebForm" visible="false" onmouseover="return escape('Controls whether this product is displayed on additional web forms. Web Forms are an installation option.')"
                        style="color: gray; cursor: help">&nbsp;?&nbsp;</a> </td></tr><tr>
                <td>
                </td>
                <td align="right" style="white-space: nowrap">
                    <asp:Label ID="lblLegendRequiresAuth" runat="server" Font-Size="XX-Small" Font-Names="Verdana">Requires Auth:</asp:Label>
                 </td>
                <td>
                </td>
                </tr><tr valign="top">
                <td align="right" rowspan="2">
                    <asp:HyperLink ID="hlnk_DetailThumb" runat="server" Target="_blank" ToolTip="click here to see larger image"
                        Font-Names="Verdana" Font-Size="XX-Small" />
                </td>
                <td>
                    <input id="fuBrowseImageFile" style="width: 240px; font-family: Verdana; font-size: xx-small"
                        type="file" runat="server" /><asp:Label ID="lblImageUploadUnavailable" runat="server"
                            Font-Names="Verdana" Font-Size="XX-Small" Text="(image upload unavailable until product created)"></asp:Label></td><td align="right" rowspan="2">
                    <asp:HyperLink ID="hlnk_PDF" runat="server" Target="_blank" ToolTip="click here to view PDF file"
                        Font-Names="Verdana" Font-Size="XX-Small" />
                </td>
                <td>
                    <input id="fuBrowsePDFFile" style="width: 240px; font-family: Verdana; font-size: xx-small"
                        type="file" runat="server" /><asp:Label ID="lblPDFUploadUnavailable" runat="server"
                            Font-Names="Verdana" Font-Size="XX-Small" Text="(PDF upload unavailable until product created)"></asp:Label></td><td align="right">
                </td>
                <td>
                </td>
            </tr>
            <tr valign="top">
                <td>
                    <asp:Button ID="btnUploadImage" runat="server" OnClick="btnUploadImage_click" Text="upload jpg"
                        ToolTip="upload the selected jpg file to the server" CausesValidation="False" />&nbsp; <asp:ImageButton ID="imgbtnDeleteImage" runat="server" ImageUrl="~/images/delete.gif"
                        OnClientClick='return confirm("Are you sure you want to delete this image?");'
                        ToolTip="delete this image" OnClick="imgbtnDeleteImage_Click" />
                    <a runat="server" id="aHelpUploadImage" visible="false" onmouseover="return escape('Allows you to upload a picture of this product. The picture must be in standard JPG format.  Pictures are automatically resized on upload if necessary.')"
                        style="color: gray; cursor: help">&nbsp;?&nbsp;</a> </td><td>
                    <asp:Button ID="btnUploadPDF" runat="server" OnClick="btnUploadPDF_click" Text="upload pdf"
                        ToolTip="upload the selected pdf file to the server" CausesValidation="False" />
                    <asp:ImageButton ID="imgbtnDeletePDF" runat="server" ImageUrl="~/images/delete.gif"
                        OnClientClick='return confirm("Are you sure you want to delete this PDF?");'
                        ToolTip="delete this PDF" OnClick="imgbtnDeletePDF_Click" />
                    <a runat="server" id="aHelpUploadPDF" visible="false" onmouseover="return escape('Allows you to upload an Adobe PDF file which can be downloaded by orderers eg to provide further information about a product')"
                        style="color: gray; cursor: help">&nbsp;?&nbsp;</a> </td><td>
                </td>
                <td align="right">
                    <asp:Button ID="btn_SaveProductChanges" runat="server" OnClick="btn_SaveProductChanges_click"
                        Text="save changes" ToolTip="save changes to product record" />&nbsp;&nbsp; </td></tr></table><asp:RegularExpressionValidator ID="revExpiryDate" runat="server" ErrorMessage="Invalid format for expiry date - use dd-mmm-yyyy"
            ControlToValidate="tbExpiryDate" ValidationExpression="^\d\d-(jan|Jan|feb|Feb|mar|Mar|apr|Apr|may|May|jun|Jun|jul|Jul|aug|Aug|sep|Sep|oct|Oct|nov|Nov|dec|Dec)-\d\d\d\d"
            Font-Names="Verdana" Font-Size="X-Small" EnableClientScript="false"></asp:RegularExpressionValidator><br />
        <asp:RangeValidator ID="rvExpiryDate" runat="server" ErrorMessage="Expiry year before 2000, after 2020, or not a valid date!"
            EnableClientScript="false" Enabled="false" ControlToValidate="tbExpiryDate" MaximumValue="2019/1/1"
            MinimumValue="2000/1/1" CultureInvariantValues="True" Font-Names="Verdana" Font-Size="X-Small">
        </asp:RangeValidator><br />
        <asp:RegularExpressionValidator ID="revReplenishmentDate" runat="server" ErrorMessage="Invalid format for renewal date - use dd-mmm-yyyy"
            ControlToValidate="tbReplenishmentDate" ValidationExpression="^\d\d-(jan|Jan|feb|Feb|mar|Mar|apr|Apr|may|May|jun|Jun|jul|Jul|aug|Aug|sep|Sep|oct|Oct|nov|Nov|dec|Dec)-\d\d\d\d"
            Font-Names="Verdana" Font-Size="X-Small" EnableClientScript="false"></asp:RegularExpressionValidator><br />
        <asp:RangeValidator ID="rvReplenishmentDate" runat="server" ErrorMessage="Renewal year before 2000, after 2020, or not a valid date!"
            EnableClientScript="false" Enabled="false" ControlToValidate="tbReplenishmentDate"
            MaximumValue="2019/1/1" MinimumValue="2000/1/1" CultureInvariantValues="True"
            Font-Names="Verdana" Font-Size="X-Small">
        </asp:RangeValidator><asp:Label ID="lblDateError" runat="server" Font-Names="Verdana"
            Font-Size="X-Small" Text="Please check Expiry Date and Replenishment Date. If present, these must be valid dates and the year must be between 2000 and 2020"
            ForeColor="Red" Visible="false" />
    </asp:Panel>
    <asp:SqlDataSource ID="SqlDataSourceCategoryList" runat="server" ConnectionString="<%$ ConnectionStrings:AIMSRootConnectionString %>"
        SelectCommand="spASPNET_Product_GetCategoriesIncludeArchivedProds" SelectCommandType="StoredProcedure">
        <SelectParameters>
            <asp:SessionParameter Name="CustomerKey" SessionField="CustomerKey" Type="Int32" />
        </SelectParameters>
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="SqlDataSourceSubCategoryList" runat="server" ConnectionString="<%$ ConnectionStrings:AIMSRootConnectionString %>"
        SelectCommand="spASPNET_Product_GetSubCategoriesIncludeArchivedProds" SelectCommandType="StoredProcedure">
        <SelectParameters>
            <asp:SessionParameter Name="CustomerKey" SessionField="CustomerKey" Type="Int32" />
            <asp:ControlParameter ControlID="ddlCategory" Name="Category" PropertyName="SelectedValue"
                Type="String" />
        </SelectParameters>
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="SqlDataSourceSubSubCategoryList" runat="server" ConnectionString="<%$ ConnectionStrings:AIMSRootConnectionString %>"
        SelectCommand="spASPNET_Product_GetSubSubCategoriesIncludeArchivedProds" SelectCommandType="StoredProcedure">
        <SelectParameters>
            <asp:SessionParameter Name="CustomerKey" SessionField="CustomerKey" Type="Int32" />
            <asp:ControlParameter ControlID="ddlSubCategory" Name="SubCategory" PropertyName="SelectedValue"
                Type="String" />
        </SelectParameters>
    </asp:SqlDataSource>
    <asp:Panel ID="pnlProductUserProfile" runat="server" Visible="False" Width="100%">
        <table style="font-family: Verdana; font-size: x-small; width: 100%">
            <tr>
                <td style="white-space: nowrap">
                    <asp:Label ID="Label20" runat="server" Visible="True" Font-Bold="true" Text="Set User Permissions for Product " />
                    <asp:Label ID="lblUserPermissionsProductCode" runat="server" ForeColor="Red" Visible="True"
                        Font-Bold="true" Text="" />
                </td>
                <td style="white-space: nowrap" align="right">
                </td>
            </tr>
            <tr>
                <td style="white-space: nowrap">
                    <asp:Button ID="btnShowAllUsers" OnClick="btn_ShowAllUsers_click" runat="server"
                        Text="show all users" />
                    &nbsp;&nbsp;&nbsp;&nbsp;search for user:&nbsp<asp:TextBox runat="server" Width="100px"
                        Font-Size="XX-Small" ID="txtProductUserSearch" />
                    &nbsp;<asp:Button ID="btnSearchUsers" OnClick="btn_SearchUsers_click" runat="server"
                        Text="go" />
                </td>
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="lblDefaultMaxGrabQty" runat="server" Visible="True" Text="Default max grab quantity:" />
                    &nbsp;<asp:TextBox ID="txtDefaultGrabQty" runat="server" Visible="True" Font-Names="Verdana"
                        Font-Size="XX-Small" Width="50px" Text="0" />
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:Button ID="btnSaveProductUserProfileChanges1"
                        OnClick="btnSaveProductUserProfileChanges_click" runat="server" Text="save changes" />
                    &nbsp;&nbsp;&nbsp;&nbsp;<asp:Button ID="btnCancel1" OnClick="btn_GoBackToProductDetail_Click"
                        runat="server" Text="cancel" />
                </td>
            </tr>
        </table>
        <asp:DataGrid ID="grid_ProductUsers" runat="server" Width="100%" Font-Names="Verdana"
            Font-Size="XX-Small" Visible="False" AutoGenerateColumns="False" GridLines="None"
            ShowFooter="True" AllowSorting="True" OnSortCommand="SortProductUsersGrid">
            <HeaderStyle Font-Size="10pt" Font-Names="Arial" Wrap="False" BorderColor="Gray">
            </HeaderStyle>
            <AlternatingItemStyle BackColor="White"></AlternatingItemStyle>
            <ItemStyle Font-Size="XX-Small" Font-Names="Arial" BackColor="WhiteSmoke"></ItemStyle>
            <Columns>
                <asp:BoundColumn Visible="False" DataField="Key" HeaderText="Key"></asp:BoundColumn>
                <asp:BoundColumn DataField="UserID" SortExpression="UserID" HeaderText="User ID">
                    <HeaderStyle Font-Size="XX-Small" Font-Names="Verdana" Font-Bold="True" ForeColor="Blue"
                        VerticalAlign="Bottom"></HeaderStyle>
                    <ItemStyle Wrap="False"></ItemStyle>
                </asp:BoundColumn>
                <asp:BoundColumn DataField="FirstName" SortExpression="FirstName" HeaderText="First Name">
                    <HeaderStyle Font-Size="XX-Small" Font-Names="Verdana" Font-Bold="True" Wrap="False"
                        ForeColor="Blue" VerticalAlign="Bottom"></HeaderStyle>
                    <ItemStyle Wrap="False"></ItemStyle>
                </asp:BoundColumn>
                <asp:BoundColumn DataField="LastName" SortExpression="LastName" HeaderText="Last Name">
                    <HeaderStyle Font-Size="XX-Small" Font-Names="Verdana" Font-Bold="True" Wrap="False"
                        ForeColor="Blue" VerticalAlign="Bottom"></HeaderStyle>
                    <ItemStyle Wrap="False"></ItemStyle>
                </asp:BoundColumn>
                <asp:TemplateColumn HeaderText="Allow Pick">
                    <HeaderStyle Font-Size="XX-Small" Font-Names="Verdana" Wrap="False" HorizontalAlign="Center"
                        ForeColor="Gray" Width="6%" VerticalAlign="Bottom"></HeaderStyle>
                    <ItemStyle Wrap="False" HorizontalAlign="Center"></ItemStyle>
                    <HeaderTemplate>
                        <asp:Button ID="btnToggleAllowToPickCheckboxes" OnClick="btnToggleAllowToPickCheckboxes_Click"
                            runat="server" Text="select all" />
                        <br />
                        <asp:Label ID="Label21" runat="server" Font-Size="XX-Small" Font-Names="Verdana">Allow to pick</asp:Label>&nbsp;</HeaderTemplate><ItemTemplate>
                        <asp:CheckBox ID="cbAllowToPick" Checked='<%# DataBinder.Eval(Container, "DataItem.AbleToPick") %>'
                            runat="server"></asp:CheckBox>
                    </ItemTemplate>
                </asp:TemplateColumn>
                <asp:TemplateColumn HeaderText="Apply max grab">
                    <HeaderStyle Font-Size="XX-Small" Font-Names="Verdana" Wrap="False" HorizontalAlign="Center"
                        ForeColor="Gray" VerticalAlign="Bottom"></HeaderStyle>
                    <ItemStyle Wrap="False" HorizontalAlign="Center"></ItemStyle>
                    <HeaderTemplate>
                        <asp:Button ID="btnToggleApplyMaxGrabCheckboxes" OnClick="btnToggleApplyMaxGrabCheckboxes_Click"
                            runat="server" Text="select all" />
                        <br />
                        <asp:Label ID="Label22" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                            Text="Apply max grab" />&nbsp;</HeaderTemplate><ItemTemplate>
                        <asp:CheckBox ID="cbApplyMaxGrab" Checked='<%# DataBinder.Eval(Container, "DataItem.ApplyMaxGrab") %>'
                            runat="server"></asp:CheckBox>
                    </ItemTemplate>
                </asp:TemplateColumn>
                <asp:TemplateColumn HeaderText="Qty">
                    <HeaderStyle Font-Size="XX-Small" Font-Names="Verdana" Wrap="False" HorizontalAlign="Center"
                        ForeColor="Gray" VerticalAlign="Bottom"></HeaderStyle>
                    <ItemStyle Wrap="False" HorizontalAlign="Center"></ItemStyle>
                    <HeaderTemplate>
                        <asp:Button ID="btnApplyMaxGrabQty" OnClick="btn_ApplyMaxGrabQty_Click" runat="server"
                            Text="apply default max grab qty" />
                        <a onmouseover="return escape('Copies the value you place in <b>Default max grab quantity</b> to the <b>Max grab quantity</b> for each user - the values are not saved until you click <b>save changes</b>')"
                            style="color: gray; cursor: help">&nbsp;?&nbsp;</a> <br />
                        <asp:Label ID="Label001" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                            Text="Max grab qty" />
                    </HeaderTemplate>
                    <ItemTemplate>
                        <asp:TextBox ID="txtMaxGrabQty" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Width="50px" Text='<%# DataBinder.Eval(Container, "DataItem.MaxGrabQty") %>'></asp:TextBox></ItemTemplate></asp:TemplateColumn></Columns></asp:DataGrid><asp:Label ID="lblProductProfileMessage" runat="server" ForeColor="Gray"
            Font-Size="X-Small" Font-Names="Verdana"></asp:Label><br />
        <asp:Table ID="tblSaveCancelProductProfile" runat="server" Width="100%" Visible="True"
            Font-Names="Verdana" Font-Size="X-Small">
            <asp:TableRow VerticalAlign="Middle">
                <asp:TableCell HorizontalAlign="Left" VerticalAlign="Middle"></asp:TableCell><asp:TableCell
                    HorizontalAlign="Right" VerticalAlign="Middle">
                    <asp:Button ID="btnSaveProductUserProfileChanges2" OnClick="btnSaveProductUserProfileChanges_click"
                        runat="server" Text="save changes" />
                    &nbsp;&nbsp;<asp:Button ID="btnCancel2" OnClick="btn_GoBackToProductDetail_Click"
                        runat="server" Text="cancel" />
                </asp:TableCell></asp:TableRow></asp:Table>&nbsp;</asp:Panel>
    <asp:Panel ID="pnlAssociatedProducts" runat="server" Width="100%">
        <table style="width: 100%">
            <tr>
                <td style="width: 50%; height: 26px;">
                    <strong>
                        <asp:Label ID="Label56" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="X-Small"
                            Text="Associated Products for Product "></asp:Label></strong><asp:Label ID="lblAssociatedProductsProductCode"
                                runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small"></asp:Label></td><td align="right" style="width: 50%; height: 26px;">
                    <asp:Button ID="btnAssociatedProductsGoBack" runat="server" Text="go back" OnClick="btn_GoBackToProductDetail_Click" />
                </td>
            </tr>
        </table>
        <br />
        <asp:GridView ID="gvAssociatedProducts" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
            AutoGenerateColumns="false" Width="100%" CellPadding="2">
            <Columns>
                <asp:TemplateField ItemStyle-Width="10%">
                    <ItemTemplate>
                        <asp:Button ID="btnRemoveAssociatedProduct" runat="server" CommandArgument='<%# Container.DataItem("LogisticAssociatedProductKey")%>'
                            Text="remove" OnClick="btnRemoveAssociatedProduct_Click" Style="width: 80px" />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:BoundField DataField="ProductCode" HeaderText="Product Code" SortExpression="ProductCode"
                    ControlStyle-Width="10%" />
                <asp:BoundField DataField="ProductDate" HeaderText="Value Date" SortExpression="ProductDate"
                    ControlStyle-Width="10%" />
                <asp:BoundField DataField="ProductDescription" HeaderText="Description" SortExpression="ProductDescription"
                    ControlStyle-Width="60%" />
                <asp:BoundField DataField="LanguageId" HeaderText="Language" SortExpression="LanguageId"
                    ControlStyle-Width="10%" />
            </Columns>
            <RowStyle BackColor="WhiteSmoke" />
            <AlternatingRowStyle BackColor="White" />
            <EmptyDataTemplate>
                this product has no associated products</EmptyDataTemplate></asp:GridView><br />
        <strong>
            <asp:Label ID="Label57" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="X-Small"
                Text="Add Associated Product"></asp:Label></strong><br />
        <br />
        <asp:GridView ID="gvUnassociatedProducts" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
            AutoGenerateColumns="false" Width="100%" CellPadding="2">
            <Columns>
                <asp:TemplateField ItemStyle-Width="10%">
                    <ItemTemplate>
                        <asp:Button ID="btnAddAssociatedProduct" runat="server" CommandArgument='<%# Container.DataItem("LogisticProductKey")%>'
                            Text="add" OnClick="btnAddAssociatedProduct_Click" />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:BoundField DataField="ProductCode" HeaderText="Product Code" SortExpression="ProductCode"
                    ControlStyle-Width="10%" />
                <asp:BoundField DataField="ProductDate" HeaderText="Value Date" SortExpression="ProductDate"
                    ControlStyle-Width="10%" />
                <asp:BoundField DataField="ProductDescription" HeaderText="Description" SortExpression="ProductDescription"
                    ControlStyle-Width="60%" />
                <asp:BoundField DataField="LanguageId" HeaderText="Language" SortExpression="LanguageId"
                    ControlStyle-Width="10%" />
            </Columns>
            <RowStyle BackColor="WhiteSmoke" />
            <AlternatingRowStyle BackColor="White" />
            <EmptyDataTemplate>
                no products found</EmptyDataTemplate></asp:GridView></asp:Panel><asp:Panel ID="pnlProductInactivityAlertStatus" runat="server" Font-Names="Verdana"
        Font-Size="X-Small" Width="100%" Visible="false">
        <table style="width: 100%">
            <tr>
                <td style="width: 50%; height: 26px;">
                    <asp:Label ID="Label48pias" runat="server" Font-Bold="True" Font-Names="Verdana"
                        Font-Size="X-Small" Text="Products Using Inactivity Alert" />
                </td>
                <td style="width: 50%; height: 26px;" align="right">
                    <asp:Button ID="btnBackFromProductInactivityAlertStatus1" runat="server" Text="back to product"
                        OnClick="btnBackFromProductInactivityAlertStatus_Click" />
                    &nbsp; </td></tr></table><table style="width: 100%">
            <tr>
                <td style="width: 5%">
                </td>
                <td>
                    &nbsp; </td></tr><tr>
                <td style="width: 5%">
                </td>
                <td>
                    <asp:GridView ID="gvProductInactivityAlert" runat="server" CellPadding="2" Font-Names="Verdana"
                        Font-Size="XX-Small" Width="100%">
                        <EmptyDataTemplate>
                            no products are using the inactivity alert feature</EmptyDataTemplate></asp:GridView></td></tr><tr>
                <td style="height: 26px">
                </td>
                <td style="height: 26px">
                    <asp:Button ID="btnBackFromProductInactivityAlertStatus2" runat="server" Text="back to product"
                        OnClick="btnBackFromProductInactivityAlertStatus_Click" />
                </td>
            </tr>
        </table>
    </asp:Panel>
    <asp:Panel ID="pnlConfigureProductInactivityAlert" runat="server" Font-Names="Verdana"
        Font-Size="X-Small" Width="100%" Visible="false">
        <table style="width: 100%">
            <tr>
                <td style="width: 50%; height: 26px;">
                    <asp:Label ID="Label48cpia" runat="server" Font-Bold="True" Font-Names="Verdana"
                        Font-Size="X-Small" Text="Configure Product Inactivity Alerts" />
                </td>
                <td style="width: 50%; height: 26px;" align="right">
                    <asp:Button ID="btnBackFromConfigureProductInactivityAlert1" runat="server" Text="back to product"
                        CausesValidation="false" OnClick="btnBackFromConfigureProductInactivityAlert_Click" />
                    &nbsp; </td></tr></table><table style="width: 100%">
            <tr>
                <td style="width: 5%">
                </td>
                <td>
                    <asp:Label ID="Label64" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small"
                        Text='NOTE: These buttons view or set Product Inactivity Alerts on <b><font color="red">ALL</font></b> products. A value of 0 disables Product Inactivity Alerts.' />
                </td>
            </tr>
            <tr>
                <td style="width: 5%">
                </td>
                <td>
                    <asp:Button ID="btnViewProductsUsingInactivityAlert" runat="server" Text="view products using inactivity alert"
                        Width="250px" OnClick="btnViewProductsUsingInactivityAlert_Click" />
                </td>
            </tr>
            <tr>
                <td style="width: 5%">
                </td>
                <td>
                    <asp:Button ID="btnSetProductInactivityAlertAllExistingProducts" runat="server" Text="set all existing products"
                        Width="250px" OnClick="btnSetProductInactivityAlertAllExistingProducts_Click" />
                    <asp:Label ID="Label53cpia" runat="server" Font-Bold="False" Font-Names="Verdana"
                        Font-Size="XX-Small" Text="to:" />
                    <asp:TextBox ID="tbProductInactivityAlertPeriodExistingProducts" runat="server" Font-Names="Verdana"
                        Font-Size="XX-Small" MaxLength="3" Width="50px">0</asp:TextBox><asp:Label ID="Label61"
                            runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small" Text="days" />
                    <asp:RequiredFieldValidator ID="rfvProductInactivityAlertPeriodExistingProducts"
                        runat="server" ControlToValidate="tbProductInactivityAlertPeriodExistingProducts"
                        ErrorMessage="required!" Font-Names="Verdana" Font-Size="XX-Small"></asp:RequiredFieldValidator><asp:RangeValidator
                            ID="rvProductInactivityAlertPeriodExistingProducts" runat="server" ControlToValidate="tbProductInactivityAlertPeriodExistingProducts"
                            ErrorMessage="must be a number between 0 and 999" Font-Names="Verdana" Font-Size="XX-Small"
                            MaximumValue="1000" MinimumValue="0" Type="Integer"></asp:RangeValidator></td></tr><tr>
                <td>
                </td>
                <td>
                    <asp:Button ID="btnSetProductInactivityAlertNewProducts" runat="server" Text="set new products"
                        Width="250px" OnClick="btnSetProductInactivityAlertNewProducts_Click" />
                    <asp:Label ID="Label62" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="to:" />
                    <asp:TextBox ID="tbProductInactivityAlertPeriodNewProducts" runat="server" Font-Names="Verdana"
                        Font-Size="XX-Small" MaxLength="3" Width="50px">0</asp:TextBox><asp:Label ID="Label63"
                            runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small" Text="days" />
                    <asp:RequiredFieldValidator ID="rfvProductInactivityAlertPeriodNewProducts" runat="server"
                        ControlToValidate="tbProductInactivityAlertPeriodNewProducts" ErrorMessage="required!"
                        Font-Names="Verdana" Font-Size="XX-Small"></asp:RequiredFieldValidator><asp:RangeValidator
                            ID="rvProductInactivityAlertPeriodNewProducts" runat="server" ControlToValidate="tbProductInactivityAlertPeriodNewProducts"
                            ErrorMessage="must be a number between 0 and 999" Font-Names="Verdana" Font-Size="XX-Small"
                            MaximumValue="1000" MinimumValue="0" Type="Integer"></asp:RangeValidator><asp:Label
                                ID="lblAvailableToSuperUsersOnly" runat="server" Text="(available to Super Users only)"
                                Visible="False" Font-Names="Verdana" Font-Size="XX-Small" ForeColor="Red"></asp:Label></td></tr><tr>
                <td>
                </td>
                <td>
                    &nbsp; &nbsp; </td></tr><tr>
                <td>
                </td>
                <td>
                    <asp:Button ID="btnBackFromConfigureProductInactivityAlert2" runat="server" Text="back to product"
                        CausesValidation="false" OnClick="btnBackFromConfigureProductInactivityAlert_Click" />
                </td>
            </tr>
        </table>
    </asp:Panel>
    <br />
    <asp:Label ID="lblError" runat="server" Font-Names="Verdana" Font-Size="X-Small"
        ForeColor="Red"></asp:Label><script type="text/javascript">

                                        function OpenHelpWindow(value) {
                                            window.open(value, "Help", "top=10,left=10,width=500,height=400,status=no,toolbar=no,address=no,menubar=no,resizable=no,scrollbars=yes");
                                        }
                                        function ShowImage(value) {
                                            window.open("show_image.aspx?Image=" + value, "ProductImage", "top=10,left=10,width=700,height=700,status=no,toolbar=no,address=no,menubar=no,resizable=yes,scrollbars=yes");
                                        }
        </script></form><script language="JavaScript" type="text/javascript" src="wz_tooltip.js"></script><script language="JavaScript" type="text/javascript" src="library_functions.js"></script></body></html>