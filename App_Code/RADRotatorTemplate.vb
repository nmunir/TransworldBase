Imports Microsoft.VisualBasic
Imports Telerik.Web.UI
Imports System.Data
Imports System.Data.SqlClient


Public Class RadRotatorTemplate
    Implements ITemplate

    Private bTitle As Boolean
    Private bContent As Boolean
    Private bImage As Boolean
    Private nImageHeight As Integer
    Private nImageWidth As Integer

    Public Property ImageHeight() As Integer
        Get
            Return nImageHeight
        End Get
        Set(value As Integer)
            nImageHeight = value
        End Set
    End Property
    Public Property ImageWidth() As Integer
        Get
            Return nImageWidth
        End Get
        Set(value As Integer)
            nImageWidth = value
        End Set
    End Property

    Public Sub New(ByVal args() As String)

        For Each arg As String In args

            If arg.ToLower.Contains("title") Then
                bTitle = True
            End If

            If arg.ToLower.Contains("content") Then
                bContent = True
            End If
            If arg.ToLower.Contains("image") Then
                bImage = True
            End If

        Next

    End Sub

    Public Sub New()


    End Sub
  
    Public Sub InstantiateIn(ByVal container As System.Web.UI.Control) Implements System.Web.UI.ITemplate.InstantiateIn
        'Dim lc1 As New lbleralControl("<div>")
        'container.Controls.Add(lc1)

        If bImage Then
            Dim img As New Image()
            img.ID = "RotatorImage"
            'img.Width = Unit.Pixel(nImageWidth)
            'img.Height = Unit.Pixel(nImageHeight)
            img.CssClass = "image"
            AddHandler img.DataBinding, AddressOf img_DataBinding
            container.Controls.Add(img)
        End If

        If bTitle Then
            Dim lblTitle As New Label
            'lblTitle.CssClass = "label"
            AddHandler lblTitle.DataBinding, AddressOf lblTitle_DataBinding
            container.Controls.Add(lblTitle)
        End If

        If bContent Then
            Dim lblContent As New Label
            AddHandler lblContent.DataBinding, AddressOf lblContent_DataBinding
            container.Controls.Add(lblContent)
        End If




        'Dim lc2 As New lbleralControl("</div>")
        'container.Controls.Add(lc2)


        'Dim lc3 As New lbleralControl
        'Dim lc4 As New lbleralControl

        'lblDate.ID = "lblDate"
        'lblDate.Text = "<%# Eval('Title') %>"
        'lblDate.Text = "Muhammad"

        'AddHandler lc3.DataBinding, AddressOf lblDate_DataBinding

        'lblContent.ID = "lblContent"
        'lblContent.Text = "<%# Eval('Content') %>"
        'lblContent.Text = "Kazim"

        'AddHandler lc4.DataBinding, AddressOf lblContent_DataBinding

        'container.Controls.Add(lc3)
        'container.Controls.Add(lc4)

        'Dim img As New Image()
        'img.AlternateText = "RadRotator in SharePoint"
        'img.Width = Unit.Pixel(350)
        'AddHandler img.DataBinding, AddressOf img_DataBinding
        'container.Controls.Add(img)


        'container.Controls.Add(lblDate)
        'container.Controls.Add(lblContent)




    End Sub

    Sub img_DataBinding(ByVal sender As Object, ByVal e As EventArgs)

        Dim img As Image = TryCast(sender, Image)
        Dim item As RadRotatorItem = TryCast(img.NamingContainer, RadRotatorItem)
        Dim sImageTag As String = TryCast(item.DataItem, System.Data.DataRowView)("ImageTag").ToString
        Dim sql As String = "select [FileName] from CustomerImages where ImageTag = '" & sImageTag & "'"
        Dim oConn As New SqlConnection(ConfigLib.GetConfigItem_ConnectionString)
        Dim oCmd As New SqlCommand(sql, oConn)
        Dim sImageName As String = String.Empty

        Try
            oConn.Open()
            If Not IsDBNull(oCmd.ExecuteScalar) Then
                sImageName = oCmd.ExecuteScalar()
            End If
        Catch ex As Exception
            WebMsgBox.Show(ex.Message.ToString())
        Finally
            oConn.Close()
        End Try

        img.ImageUrl = "~/Images/" & sImageName

        Dim sUrl As String = TryCast(item.DataItem, System.Data.DataRowView)("Url").ToString

        If Not String.IsNullOrEmpty(sUrl) Then
            img.Attributes.Add("onclick", "window.open('" & sUrl & "')")
        End If

        'img.Attributes.Add("onclick", "window.open('" & TryCast(item.DataItem, System.Data.DataRowView)("Url").ToString & "')")



    End Sub

    Protected Sub lblContent_DataBinding(ByVal sender As Object, ByVal e As EventArgs)

        Dim lblContent As Label = TryCast(sender, Label)
        Dim item As RadRotatorItem = TryCast(lblContent.NamingContainer, RadRotatorItem)
        'lblContent.Text = "&nbsp;" & TryCast(item.DataItem, System.Data.DataRowView)("Content").ToString & "<br>"
        lblContent.Text = "&nbsp;" & TryCast(item.DataItem, System.Data.DataRowView)("Content").ToString

    End Sub

    Protected Sub lblTitle_DataBinding(ByVal sender As Object, ByVal e As EventArgs)

        Dim lblTitle As Label = TryCast(sender, Label)
        Dim item As RadRotatorItem = TryCast(lblTitle.NamingContainer, RadRotatorItem)
        lblTitle.Text = "&nbsp;" & TryCast(item.DataItem, System.Data.DataRowView)("Title").ToString & "<br>"

    End Sub

End Class