Imports Microsoft.VisualBasic
Imports System.Collections
Imports System.Configuration.ConfigurationManager

' TO DO
' use constant location prefix PHYSICAL_WEB_ROOT and take from master WEB.CONFIG, then Path.Combine with default

Public Class ConfigLib
    Public Shared Function GetConfigItem_LegalAndGeneralCustomerKey() As Integer
        'GetConfigItem_LegalAndGeneralCustomerKey = 476
        GetConfigItem_LegalAndGeneralCustomerKey = 16
    End Function

    Public Shared Function GetConfigItem_PrintOnDemand() As Boolean
        Const DEFAULT_PRINT_ON_DEMAND As Boolean = False
        Try
            GetConfigItem_PrintOnDemand = CBool(AppSettings.Item("PrintOnDemand"))
        Catch e As Exception
            GetConfigItem_PrintOnDemand = DEFAULT_PRINT_ON_DEMAND
        End Try
    End Function

    Public Shared Function GetConfigItem_ShowNotes() As Boolean
        Const DEFAULT_SHOW_NOTES As Boolean = True
        Try
            GetConfigItem_ShowNotes = CBool(AppSettings.Item("ShowNotes"))
        Catch e As Exception
            GetConfigItem_ShowNotes = DEFAULT_SHOW_NOTES
        End Try
    End Function

    Public Shared Function GetConfigItem_ApplyMaxGrabs() As Boolean
        Const DEFAULT_APPLY_MAX_GRABS As Boolean = True
        Try
            GetConfigItem_ApplyMaxGrabs = CBool(AppSettings.Item("ApplyMaxGrabs"))
        Catch e As Exception
            GetConfigItem_ApplyMaxGrabs = DEFAULT_APPLY_MAX_GRABS
        End Try
    End Function

    Public Shared Function GetConfigItem_AppTitle() As String
        Dim o As Object = AppSettings.Item("AppTitle")
        If IsNothing(o) Then
            GetConfigItem_AppTitle = ""
        Else
            GetConfigItem_AppTitle = CStr(o)
        End If
    End Function

    Public Shared Function GetConfigItem_OrderAuthorisationAdvisory() As String
        Const DEFAULT_ORDER_AUTHORISATION_ADVISORY As String = " (THIS ORDER WILL BE SENT FOR AUTHORISATION)"
        Dim o As Object = AppSettings.Item("OrderAuthorisationAdvisory")
        If IsNothing(o) Then
            GetConfigItem_OrderAuthorisationAdvisory = DEFAULT_ORDER_AUTHORISATION_ADVISORY
        Else
            GetConfigItem_OrderAuthorisationAdvisory = CStr(o)
        End If
    End Function

    Public Shared Function GetConfigItem_ConnectionString() As String
        GetConfigItem_ConnectionString = ""
        Try
            GetConfigItem_ConnectionString = ConnectionStrings("TempConnectionString").ConnectionString
        Catch e As System.NullReferenceException
            Try
                GetConfigItem_ConnectionString = ConnectionStrings("AIMSRootConnectionString").ConnectionString
            Catch e2 As System.NullReferenceException

            End Try
        End Try
    End Function

    Public Shared Function GetConfigItem_EnableAuthorisation() As Boolean
        Const DEFAULT_ENABLE_AUTHORISATION As Boolean = False
        Try
            GetConfigItem_EnableAuthorisation = CBool(AppSettings.Item("EnableAuthorisation"))
        Catch e As Exception
            GetConfigItem_EnableAuthorisation = DEFAULT_ENABLE_AUTHORISATION
        End Try
    End Function

    Public Shared Function GetConfigItem_EnableCalendarManagement() As Boolean
        Const DEFAULT_ENABLE_CALENDAR_MANAGEMENT As Boolean = False
        Try
            GetConfigItem_EnableCalendarManagement = CBool(AppSettings.Item("EnableCalendarManagement"))
        Catch e As Exception
            GetConfigItem_EnableCalendarManagement = DEFAULT_ENABLE_CALENDAR_MANAGEMENT
        End Try
    End Function

    Public Shared Function GetConfigItem_AuthorisationGranularity() As String
        Const DEFAULT_AUTHORISATION_GRANULARITY As String = "Product"
        Dim o As Object = AppSettings.Item("AuthorisationGranularity")
        If IsNothing(o) Then
            GetConfigItem_AuthorisationGranularity = DEFAULT_AUTHORISATION_GRANULARITY
        Else
            GetConfigItem_AuthorisationGranularity = CStr(o)
        End If
    End Function

    Public Shared Function GetConfigItem_EnablePostCodeLookup() As Boolean
        Const DEFAULT_ENABLE_POSTCODE_LOOKUP As Boolean = False
        Try
            GetConfigItem_EnablePostCodeLookup = CBool(AppSettings.Item("EnablePostCodeLookup"))
        Catch e As Exception
            GetConfigItem_EnablePostCodeLookup = DEFAULT_ENABLE_POSTCODE_LOOKUP
        End Try
    End Function

    Public Shared Function GetConfigItem_EnableRotation() As Boolean
        Const DEFAULT_ENABLE_ROTATION As Boolean = False
        Try
            GetConfigItem_EnableRotation = CBool(AppSettings.Item("EnableRotation"))
        Catch e As Exception
            GetConfigItem_EnableRotation = DEFAULT_ENABLE_ROTATION
        End Try
    End Function

    Public Shared Function GetConfigItem_EstimatePackaging() As String
        '            <add key="EstimatePackaging" value="Y" />
        Const DEFAULT_ESTIMATE_PACKAGING As String = "Y"
        Dim o As Object = AppSettings.Item("EstimatePackaging")
        If IsNothing(o) Then
            GetConfigItem_EstimatePackaging = DEFAULT_ESTIMATE_PACKAGING
        Else
            GetConfigItem_EstimatePackaging = CStr(o)
        End If
        'Try
        ' GetConfigItem_EstimatePackaging = CBool(AppSettings.Item("EstimatePackaging"))
        ' Catch e As Exception
        ' GetConfigItem_EstimatePackaging = DEFAULT_ESTIMATE_PACKAGING
        ' End Try
    End Function

    Public Shared Function GetConfigItem_prod_image_folder() As String
        Const DEFAULT_PROD_IMAGE_FOLDER As String = "D:\Couriersoftware\www\images\jpgs\"
        Try
            GetConfigItem_prod_image_folder = AppSettings.Item("prod_image_folder")
        Catch e As System.NullReferenceException
            GetConfigItem_prod_image_folder = DEFAULT_PROD_IMAGE_FOLDER
        End Try
        If GetConfigItem_prod_image_folder Is Nothing Then
            GetConfigItem_prod_image_folder = DEFAULT_PROD_IMAGE_FOLDER
        End If
    End Function

    Public Shared Function GetConfigItem_SiteType() As String
        Try
            GetConfigItem_SiteType = AppSettings.Item("SiteType")
        Catch e As System.NullReferenceException
            GetConfigItem_SiteType = String.Empty
        End Try
        If GetConfigItem_SiteType Is Nothing Then
            GetConfigItem_SiteType = String.Empty
        End If
    End Function

    Public Shared Function GetConfigItem_CategoryCount() As Integer
        Try
            Dim o As Object = AppSettings.Item("CategoryCount")
            If IsNothing(o) Then
                GetConfigItem_CategoryCount = 2
            Else
                GetConfigItem_CategoryCount = CInt(o)
            End If
            'GetConfigItem_CategoryCount = CInt(AppSettings.Item("CategoryCount"))
        Catch e As System.NullReferenceException
            GetConfigItem_CategoryCount = 2
        End Try
    End Function

    Public Shared Function GetConfigItem_prod_pdf_folder() As String
        Const DEFAULT_PROD_PDF_FOLDER As String = "D:\Couriersoftware\www\images\pdfs\"
        Try
            GetConfigItem_prod_pdf_folder = AppSettings.Item("prod_pdf_folder")
        Catch e As System.NullReferenceException
            GetConfigItem_prod_pdf_folder = DEFAULT_PROD_PDF_FOLDER
        End Try
        If GetConfigItem_prod_pdf_folder Is Nothing Then
            GetConfigItem_prod_pdf_folder = DEFAULT_PROD_PDF_FOLDER
        End If
    End Function

    Public Shared Function GetConfigItem_prod_thumb_folder() As String
        Const DEFAULT_PROD_THUMB_FOLDER As String = "D:\Couriersoftware\www\images\thumbs\"
        Try
            GetConfigItem_prod_thumb_folder = AppSettings.Item("prod_thumb_folder")
        Catch e As System.NullReferenceException
            GetConfigItem_prod_thumb_folder = DEFAULT_PROD_THUMB_FOLDER
        End Try
        If GetConfigItem_prod_thumb_folder Is Nothing Then
            GetConfigItem_prod_thumb_folder = DEFAULT_PROD_THUMB_FOLDER
        End If
    End Function

    Public Shared Function GetConfigItem_ProjectsFolder() As String
        Const DEFAULT_PROJECTS_FOLDER As String = "ProjectDocs"
        Try
            GetConfigItem_ProjectsFolder = AppSettings.Item("ProjectsFolder")
        Catch e As System.NullReferenceException
            GetConfigItem_ProjectsFolder = DEFAULT_PROJECTS_FOLDER
        End Try
        If GetConfigItem_ProjectsFolder Is Nothing Then
            GetConfigItem_ProjectsFolder = DEFAULT_PROJECTS_FOLDER
        End If
    End Function

    Public Shared Function GetConfigItem_ProjectsSourceDoc() As String
        Const DEFAULT_SOURCE_DOC As String = "source.doc"
        Try
            GetConfigItem_ProjectsSourceDoc = AppSettings.Item("ProjectsSourceDoc")
        Catch e As System.NullReferenceException
            GetConfigItem_ProjectsSourceDoc = DEFAULT_SOURCE_DOC
        End Try
        If GetConfigItem_ProjectsSourceDoc Is Nothing Then
            GetConfigItem_ProjectsSourceDoc = DEFAULT_SOURCE_DOC
        End If
    End Function

    Public Shared Function GetConfigItem_MultiAddressOrders() As Boolean
        Const DEFAULT_MULTI_ADDRESS_ORDERS As Boolean = False
        Try
            GetConfigItem_MultiAddressOrders = CBool(AppSettings.Item("MultiAddressOrders"))
        Catch e As Exception
            GetConfigItem_MultiAddressOrders = DEFAULT_MULTI_ADDRESS_ORDERS
        End Try
    End Function

    Public Shared Function GetConfigItem_ShowZeroStockBalances() As Boolean
        Const DEFAULT_SHOW_ZERO_STOCK_BALANCES As Boolean = False
        Try
            GetConfigItem_ShowZeroStockBalances = CBool(AppSettings.Item("ShowZeroStockBalances"))
        Catch e As Exception
            GetConfigItem_ShowZeroStockBalances = DEFAULT_SHOW_ZERO_STOCK_BALANCES
        End Try
    End Function

    Public Shared Function GetConfigItem_Virtual_JPG_URL() As String
        Const DEFAULT_VIRTUAL_JPG_URL As String = "./prod_images/jpgs/"
        Try
            GetConfigItem_Virtual_JPG_URL = AppSettings.Item("Virtual_JPG_URL")
        Catch e As System.NullReferenceException
            GetConfigItem_Virtual_JPG_URL = DEFAULT_VIRTUAL_JPG_URL
        End Try
        If GetConfigItem_Virtual_JPG_URL Is Nothing Then
            GetConfigItem_Virtual_JPG_URL = DEFAULT_VIRTUAL_JPG_URL
        End If
    End Function

    Public Shared Function GetConfigItem_Virtual_PDF_URL() As String
        Const DEFAULT_VIRTUAL_PDF_URL As String = "./prod_images/pdfs/"
        Try
            GetConfigItem_Virtual_PDF_URL = AppSettings.Item("Virtual_PDF_URL")
        Catch e As System.NullReferenceException
            GetConfigItem_Virtual_PDF_URL = DEFAULT_VIRTUAL_PDF_URL
        End Try
        If GetConfigItem_Virtual_PDF_URL Is Nothing Then
            GetConfigItem_Virtual_PDF_URL = DEFAULT_VIRTUAL_PDF_URL
        End If
    End Function

    Public Shared Function GetConfigItem_Virtual_Thumb_URL() As String
        Const DEFAULT_VIRTUAL_THUMB_URL As String = "./prod_images/thumbs/"
        Try
            GetConfigItem_Virtual_Thumb_URL = AppSettings.Item("Virtual_Thumb_URL")
        Catch e As System.NullReferenceException
            GetConfigItem_Virtual_Thumb_URL = DEFAULT_VIRTUAL_THUMB_URL
        End Try
        If GetConfigItem_Virtual_Thumb_URL Is Nothing Then
            GetConfigItem_Virtual_Thumb_URL = DEFAULT_VIRTUAL_THUMB_URL
        End If
    End Function

    Public Shared Function GetConfigItem_SA_Home_Page() As String
        Const DEFAULT_HOME_PAGE As String = "NoticeBoard.aspx"
        Dim sPage As String
        Try
            sPage = ConfigurationManager.AppSettings("sa_home_page")
        Catch
            sPage = DEFAULT_HOME_PAGE
        End Try
        If sPage Is Nothing Then
            sPage = DEFAULT_HOME_PAGE
        End If
        GetConfigItem_SA_Home_Page = sPage
    End Function

    Public Shared Function GetConfigItem_Admin_Home_Page() As String
        Const DEFAULT_HOME_PAGE As String = "NoticeBoard.aspx"
        Dim sPage As String
        Try
            sPage = ConfigurationManager.AppSettings("admin_home_page")
        Catch
            sPage = DEFAULT_HOME_PAGE
        End Try
        If sPage Is Nothing Then
            sPage = DEFAULT_HOME_PAGE
        End If
        GetConfigItem_Admin_Home_Page = sPage
    End Function

    Public Shared Function GetConfigItem_AccountHandler_Home_Page() As String
        Const DEFAULT_HOME_PAGE As String = "NoticeBoard.aspx"
        Dim sPage As String
        Try
            sPage = ConfigurationManager.AppSettings("accounthandler_home_page")
        Catch
            sPage = DEFAULT_HOME_PAGE
        End Try
        If sPage Is Nothing Then
            sPage = DEFAULT_HOME_PAGE
        End If
        GetConfigItem_AccountHandler_Home_Page = sPage
    End Function

    Public Shared Function GetConfigItem_SuperUser_Page() As String
        Const DEFAULT_HOME_PAGE As String = "NoticeBoard.aspx"
        Dim sPage As String
        Try
            sPage = ConfigurationManager.AppSettings("superuser_home_page")
        Catch e As System.NullReferenceException
            sPage = DEFAULT_HOME_PAGE
        End Try
        If sPage Is Nothing Then
            sPage = DEFAULT_HOME_PAGE
        End If
        GetConfigItem_SuperUser_Page = sPage
    End Function

    Public Shared Function GetConfigItem_ProductOwner_Home_Page() As String
        Const DEFAULT_HOME_PAGE As String = "NoticeBoard.aspx"
        Dim sPage As String
        Try
            sPage = ConfigurationManager.AppSettings("product_owner_home_page")
        Catch
            sPage = DEFAULT_HOME_PAGE
        End Try
        If sPage Is Nothing Then
            sPage = DEFAULT_HOME_PAGE
        End If
        GetConfigItem_ProductOwner_Home_Page = sPage
    End Function

    Public Shared Function GetConfigItem_User_Home_Page() As String
        Const DEFAULT_HOME_PAGE As String = "NoticeBoard.aspx"
        Dim sPage As String
        Try
            sPage = ConfigurationManager.AppSettings("user_home_page")
        Catch e As System.NullReferenceException
            sPage = DEFAULT_HOME_PAGE
        End Try
        If sPage Is Nothing Then
            sPage = DEFAULT_HOME_PAGE
        End If
        GetConfigItem_User_Home_Page = sPage
    End Function

    Public Shared Function GetConfigItem_Supplier_Home_Page() As String
        Const DEFAULT_HOME_PAGE As String = "NoticeBoard.aspx"
        Dim sPage As String
        Try
            sPage = ConfigurationManager.AppSettings("supplier_home_page")
        Catch e As System.NullReferenceException
            sPage = DEFAULT_HOME_PAGE
        End Try
        If sPage Is Nothing Then
            sPage = DEFAULT_HOME_PAGE
        End If
        GetConfigItem_Supplier_Home_Page = sPage
    End Function

    Public Shared Function GetStartPage(ByVal sUserType As String) As String
        Const DEFAULT_HOME_PAGE As String = "NoticeBoard.aspx"
        GetStartPage = DEFAULT_HOME_PAGE
        Select Case sUserType.Trim.ToLower
            Case "sa"
                GetStartPage = GetConfigItem_SA_Home_Page()
            Case "admin"
                GetStartPage = GetConfigItem_Admin_Home_Page()
            Case "accounthandler"
                GetStartPage = GetConfigItem_AccountHandler_Home_Page()
            Case "superuser"
                GetStartPage = GetConfigItem_SuperUser_Page()
            Case "product owner"
                GetStartPage = GetConfigItem_ProductOwner_Home_Page()
            Case "user"
                GetStartPage = GetConfigItem_User_Home_Page()
            Case "supplier"
                GetStartPage = GetConfigItem_Supplier_Home_Page()
        End Select
    End Function

    Public Shared Function GetConfigItem_Default_Logon_Page() As String
        Const DEFAULT_LOGON_PAGE As String = "default.aspx"
        Dim sPage As String
        Try
            sPage = ConfigurationManager.AppSettings("Default_Logon_Page")
        Catch e As System.NullReferenceException
            sPage = DEFAULT_LOGON_PAGE
        End Try
        If sPage Is Nothing Then
            sPage = DEFAULT_LOGON_PAGE
        End If
        GetConfigItem_Default_Logon_Page = sPage
    End Function

    Public Shared Function GetConfigItem_Default_Running_Header_Image() As String
        Const DEFAULT_RUNNING_HEADER_IMAGE As String = "http://www.sprintexpress.co.uk/images/sprint_logo.png"
        Try
            GetConfigItem_Default_Running_Header_Image = AppSettings.Item("default_running_header_image")
        Catch e As System.NullReferenceException
            GetConfigItem_Default_Running_Header_Image = DEFAULT_RUNNING_HEADER_IMAGE
        End Try
        If GetConfigItem_Virtual_Thumb_URL Is Nothing Then
            GetConfigItem_Default_Running_Header_Image = DEFAULT_RUNNING_HEADER_IMAGE
        End If
    End Function

    Public Shared Function GetConfigItem_UseLabelPrinter() As Boolean
        Const DEFAULT_USE_LABEL_PRINTER As Boolean = False
        Try
            GetConfigItem_UseLabelPrinter = CBool(AppSettings.Item("UseLabelPrinter"))
        Catch e As Exception
            GetConfigItem_UseLabelPrinter = DEFAULT_USE_LABEL_PRINTER
        End Try
    End Function

    Public Shared Function GetConfigItem_SearchCompanyNameOnly() As Boolean
        Const DEFAULT_SEARCH_COMPANY_NAME_ONLY As Boolean = False
        Try
            GetConfigItem_SearchCompanyNameOnly = CBool(AppSettings.Item("SearchCompanyNameOnly"))
        Catch e As Exception
            GetConfigItem_SearchCompanyNameOnly = DEFAULT_SEARCH_COMPANY_NAME_ONLY
        End Try
    End Function

    Public Shared Function GetConfigItem_DefaultDescription() As String
        Const DEFAULT_DEFAULTDESCRIPTION As String = "Documents"
        Try
            GetConfigItem_DefaultDescription = AppSettings.Item("DefaultDescription")
        Catch e As System.NullReferenceException
            GetConfigItem_DefaultDescription = DEFAULT_DEFAULTDESCRIPTION
        End Try
    End Function

    Public Shared Function GetConfigItem_MakeRef1Mandatory() As Boolean
        Const DEFAULT_MAKEREF1MANDATORY As Boolean = True
        Try
            GetConfigItem_MakeRef1Mandatory = CBool(AppSettings.Item("MakeRef1Mandatory"))
        Catch e As Exception
            GetConfigItem_MakeRef1Mandatory = DEFAULT_MAKEREF1MANDATORY
        End Try
    End Function

    Public Shared Function GetConfigItem_Ref1Label() As String
        Const DEFAULT_REF1LABEL As String = "Company Cost Code"
        Try
            GetConfigItem_Ref1Label = AppSettings.Item("Ref1Label")
        Catch e As System.NullReferenceException
            GetConfigItem_Ref1Label = DEFAULT_REF1LABEL
        End Try
    End Function

    Public Shared Function GetConfigItem_MakeRef2Mandatory() As Boolean
        Const DEFAULT_MAKEREF2MANDATORY As Boolean = False
        Try
            GetConfigItem_MakeRef2Mandatory = CBool(AppSettings.Item("MakeRef2Mandatory"))
        Catch e As Exception
            GetConfigItem_MakeRef2Mandatory = DEFAULT_MAKEREF2MANDATORY
        End Try
    End Function

    Public Shared Function GetConfigItem_Ref2Label() As String
        Const DEFAULT_REF2LABEL As String = "Job Number"
        Try
            GetConfigItem_Ref2Label = AppSettings.Item("Ref2Label")
        Catch e As System.NullReferenceException
            GetConfigItem_Ref2Label = DEFAULT_REF2LABEL
        End Try
    End Function

    Public Shared Function GetConfigItem_MakeRef3Mandatory() As Boolean
        Const DEFAULT_MAKEREF3MANDATORY As Boolean = False
        Try
            GetConfigItem_MakeRef3Mandatory = CBool(AppSettings.Item("MakeRef3Mandatory"))
        Catch e As Exception
            GetConfigItem_MakeRef3Mandatory = DEFAULT_MAKEREF3MANDATORY
        End Try
    End Function

    Public Shared Function GetConfigItem_Ref3Label() As String
        Const DEFAULT_REF3LABEL As String = "Customer Ref 3:"
        Try
            GetConfigItem_Ref3Label = AppSettings.Item("Ref3Label")
        Catch e As System.NullReferenceException
            GetConfigItem_Ref3Label = DEFAULT_REF3LABEL
        End Try
    End Function

    Public Shared Function GetConfigItem_MakeRef4Mandatory() As Boolean
        Const DEFAULT_MAKEREF4MANDATORY As Boolean = False
        Try
            GetConfigItem_MakeRef4Mandatory = CBool(AppSettings.Item("MakeRef4Mandatory"))
        Catch e As Exception
            GetConfigItem_MakeRef4Mandatory = DEFAULT_MAKEREF4MANDATORY
        End Try
    End Function

    Public Shared Function GetConfigItem_Ref4Label() As String
        Const DEFAULT_REF4LABEL As String = "Customer Ref 4:"
        Try
            GetConfigItem_Ref4Label = AppSettings.Item("Ref4Label")
        Catch e As System.NullReferenceException
            GetConfigItem_Ref4Label = DEFAULT_REF4LABEL
        End Try
    End Function

    Public Shared Function GetConfigItem_ThirdPartyCollectionKey() As String
        Const DEFAULT_THIRDPARTYCOLLECTIONKEY As String = "-1"
        Try
            GetConfigItem_ThirdPartyCollectionKey = AppSettings.Item("ThirdPartyCollectionKey")
        Catch e As System.NullReferenceException
            GetConfigItem_ThirdPartyCollectionKey = DEFAULT_THIRDPARTYCOLLECTIONKEY
        End Try
    End Function

    Public Shared Function GetConfigItem_HideCollectionButton() As Boolean
        Const DEFAULT_HIDECOLLECTIONBUTTON As Boolean = False
        Try
            GetConfigItem_HideCollectionButton = CBool(AppSettings.Item("MakeRef4Mandatory"))
        Catch e As Exception
            GetConfigItem_HideCollectionButton = DEFAULT_HIDECOLLECTIONBUTTON
        End Try
    End Function
End Class