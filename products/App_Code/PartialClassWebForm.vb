Imports Microsoft.VisualBasic
Imports System.Web.UI.Page

'Public Class PartialClassWebForm : Inherits System.Web.UI.MasterPage

Public Class PartialClassWebForm : Inherits System.Web.UI.Page

    ' Property declarations - properties all held in VIEWSTATE

    Property lCustomerKey() As Long
        Get
            Dim o As Object = ViewState("CustomerKey")
            If o Is Nothing Then
                Return 0
            End If
            Return CLng(o)
        End Get
        Set(ByVal Value As Long)
            ViewState("CustomerKey") = Value
        End Set
    End Property

    Property lGenericUserKey() As Long
        Get
            Dim o As Object = ViewState("GenericUserKey")
            If o Is Nothing Then
                Return 0
            End If
            Return CLng(o)
        End Get
        Set(ByVal Value As Long)
            ViewState("GenericUserKey") = Value
        End Set
    End Property

    Property sCategory() As String
        Get
            Dim o As Object = ViewState("Category")
            If o Is Nothing Then
                Return "_"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("Category") = Value
        End Set
    End Property

    Property sSubCategory() As String
        Get
            Dim o As Object = ViewState("SubCategory")
            If o Is Nothing Then
                Return "_"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("SubCategory") = Value
        End Set
    End Property

    '    Property sProdImageFolder() As String
    '        Get
    '    Dim o As Object = ViewState("ProdImageFolder")
    '            If o Is Nothing Then
    '                Return ""
    '            End If
    '            Return CStr(o)
    '        End Get
    '        Set(ByVal Value As String)
    '            ViewState("ProdImageFolder") = Value
    '        End Set
    '    End Property

    '   Property sVirtualJPGFolder() As String
    '       Get
    '   Dim o As Object = ViewState("VirtualJPGFolder")
    '           If o Is Nothing Then
    '               Return ""
    '           End If
    '           Return CStr(o)
    '       End Get
    '       Set(ByVal Value As String)
    '           ViewState("VirtualJPGFolder") = Value
    '       End Set
    '   End Property'

    '    Property sProdThumbFolder() As String
    '        Get
    '    Dim o As Object = ViewState("ProdThumbFolder")
    '            If o Is Nothing Then
    '                Return ""
    '            End If
    '            Return CStr(o)
    '        End Get
    '        Set(ByVal Value As String)
    '            ViewState("ProdThumbFolder") = Value
    '        End Set
    '    End Property

    '   Property sVirtualThumbFolder() As String
    '       Get
    '   Dim o As Object = ViewState("VirtualThumbFolder")
    '           If o Is Nothing Then
    '               Return ""
    '           End If
    '           Return CStr(o)
    '       End Get
    '       Set(ByVal Value As String)
    '           ViewState("VirtualThumbFolder") = Value
    '       End Set
    '   End Property

    Property bCategoryProductsFound() As Boolean
        Get
            Dim o As Object = ViewState("CategoryProductsFound")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("CategoryProductsFound") = Value
        End Set
    End Property

End Class
