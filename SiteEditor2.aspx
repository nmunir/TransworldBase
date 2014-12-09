<%@ Page Language="VB" ValidateRequest="false" Theme="AIMSDefault" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ Register Assembly="Telerik.Web.UI" Namespace="Telerik.Web.UI" TagPrefix="telerik" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.XML" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.ServiceModel.Syndication" %>
<%@ Import Namespace="System.Net" %>
<%@ Import Namespace="System.Collections.Generic" %>
<script runat="server">
    
    ' TO DO
    ' Resized JPGs are not valid JPGs
    ' GIFs?
    ' Skin="Vista"

    Const TABLENAME_HEADER_ROTATOR As String = "HeaderRotator"
    Const TABLENAME_LHSBODY As String = "LHSBody"
    Const TABLENAME_PAGETITLE As String = "PageTitle"
    
    Const STYLESHEET_FILENAME_WORKING As String = "sprint.css"
    Const STYLESHEET_FILENAME_DEFAULT As String = "sprint_default.css"
    Const DEFAULT_STYLESHEET_PATH As String = "~\css\sprint.css"
    Const UPLOADED_IMAGES_PATH As String = "~\images\UploadedImages\"

    Const COLOUR_HIGHLIGHT As String = "#FFFF60"
    
    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    Dim gsXMLRotatorConfigFilePath As String
    Dim gsXMLNewsContentFilePath As String
    Dim gds As New DataTable
    Dim gdt As DataTable
    Dim gcol As DataColumn
    
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)
        If Not IsNumeric(Session("SiteKey")) Then
            Server.Transfer("session_expired.aspx")
            Exit Sub
        End If
        If Not IsPostBack Then
            'Dim sImagePath() As String = {MapPath(UPLOADED_IMAGES_PATH)}
            Dim sImagePath() As String = {"~/Images/Uploadedimages"}
            'Dim sImagePath() As String = {UPLOADED_IMAGES_PATH}
            'Dim sImagePath() As String = {"~\images\UploadedImages"}
            RadEditor1.ImageManager.ViewPaths = sImagePath
            Call SetTitle()
            Call CheckUploadImageDirectoryDefined()
            Call InitAreaSelector()
            psLastFieldEdited = ddlAreaSelector.SelectedValue
            Call BindDataSourceTags()
            'Call LoadDataset()
            'Call BindNewsEditor()
            Call PopulateCSSEditor()
            tblDataSourceEditor.Visible = False
            Call HideAllPanels()
            'pnlLoginPageEditor.Visible = True
        End If
    End Sub
  
    Protected Sub CheckUploadImageDirectoryDefined()
        Dim sTest As String = MapPath(UPLOADED_IMAGES_PATH)
        If Not My.Computer.FileSystem.DirectoryExists(sTest) Then
            WebMsgBox.Show("Image upload directory not defined (" & UPLOADED_IMAGES_PATH & ")")
        End If
    End Sub
    
    Protected Sub SetTitle()
        Dim sTitle As String = Session("SiteTitle")
        If sTitle <> String.Empty Then
            sTitle += " - "
        End If
        Page.Header.Title = sTitle & "Site Editor"
    End Sub
    
    Protected Sub SetStyleSheet()
        Dim hlCSSLink As New HtmlLink
        hlCSSLink.Href = Session("StyleSheetPath")
        hlCSSLink.Attributes.Add("rel", "stylesheet")
        hlCSSLink.Attributes.Add("type", "text/css")
        Page.Header.Controls.Add(hlCSSLink)
    End Sub

    Protected Sub InitAreaSelector()
        ddlAreaSelector.Items.Clear()
        If cbShowAdvancedFeatures.Checked Then
            'ddlAreaSelector.Items.Add(New ListItem("notice board - main text area", "NB1left+"))
            ddlAreaSelector.Items.Add(New ListItem("login page - main text area", "LPleft+"))
            ddlAreaSelector.Items.Add(New ListItem("", ""))
            'ddlAreaSelector.Items.Add(New ListItem("notice board - top text area", "NB1top"))
            'ddlAreaSelector.Items.Add(New ListItem("notice board - bottom text area", "NB1bottom"))
            ' ddlAreaSelector.Items.Add(New ListItem("", ""))
            ddlAreaSelector.Items.Add(New ListItem("login page - top text area", "LPtop"))
            ddlAreaSelector.Items.Add(New ListItem("login page - bottom text area", "LPbottom"))
            ddlAreaSelector.Items.Add(New ListItem("login page - right text area", "LPright"))
            ddlAreaSelector.Items.Add(New ListItem("", ""))
            'ddlAreaSelector.Items.Add(New ListItem("notice board - column layout", "NB1layout"))
            ddlAreaSelector.Items.Add(New ListItem("login page - column layout", "LPlayout"))
            ddlAreaSelector.Items.Add(New ListItem("site logo URL", "SiteLogo"))
        Else
            'ddlAreaSelector.Items.Add(New ListItem("notice board - main text area", "NB1left"))
            ddlAreaSelector.Items.Add(New ListItem("login page - main text area", "LPleft"))
        End If
        ddlAreaSelector.SelectedIndex = 0
        Call InitArea()
    End Sub
    
    Protected Sub InitArea()
        If ddlAreaSelector.SelectedValue = String.Empty Then
            Exit Sub
        End If
        Dim oDataTable As New DataTable
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter("spASPNET_SiteContent", oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@Action", SqlDbType.NVarChar, 50))
        oAdapter.SelectCommand.Parameters("@Action").Value = "GET"
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SiteKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@SiteKey").Value = Session("SiteKey")
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ContentType", SqlDbType.NVarChar, 50))
        oAdapter.SelectCommand.Parameters("@ContentType").Value = "FullRecord"
        Try
            oAdapter.Fill(oDataTable)
        Catch ex As Exception
            WebMsgBox.Show("InitArea: " & ex.Message)
        Finally
            oConn.Close()
        End Try
        Dim dr As DataRow = oDataTable.Rows(0)
        Select Case ddlAreaSelector.SelectedValue
            Case "LPleft", "LPleft+"
                RadEditor1.Visible = True
                RadEditor1.Content = (dr("LP1Content") & String.Empty).ToString.TrimEnd
                trLoginPageAdvancedControls.Visible = False
                trSiteLogoURL.Visible = False
                trNoticeBoard1AdvancedControls.Visible = False
                tbCSSAttributes.Text = dr("LP1Attr") & String.Empty
                If cbShowAdvancedFeatures.Checked Then
                    tbCSSAttributes.Visible = True
                Else
                    tbCSSAttributes.Visible = False
                End If
            Case "LPtop"
                RadEditor1.Visible = True
                RadEditor1.Content = (dr("LPTopContent") & String.Empty).ToString.TrimEnd
                trLoginPageAdvancedControls.Visible = False
                trNoticeBoard1AdvancedControls.Visible = False
                trSiteLogoURL.Visible = False
                tbCSSAttributes.Text = dr("LPTopAttr") & String.Empty
                If cbShowAdvancedFeatures.Checked Then
                    tbCSSAttributes.Visible = True
                Else
                    tbCSSAttributes.Visible = False
                End If
            Case "LPbottom"
                RadEditor1.Visible = True
                RadEditor1.Content = (dr("LPBottomContent") & String.Empty).ToString.TrimEnd
                trLoginPageAdvancedControls.Visible = False
                trNoticeBoard1AdvancedControls.Visible = False
                trSiteLogoURL.Visible = False
                tbCSSAttributes.Text = dr("LPBottomAttr") & String.Empty
                If cbShowAdvancedFeatures.Checked Then
                    tbCSSAttributes.Visible = True
                Else
                    tbCSSAttributes.Visible = False
                End If
            Case "LPright"
                RadEditor1.Visible = True
                RadEditor1.Content = (dr("LP4Content") & String.Empty).ToString.TrimEnd
                trLoginPageAdvancedControls.Visible = False
                trNoticeBoard1AdvancedControls.Visible = False
                trSiteLogoURL.Visible = False
                tbCSSAttributes.Text = dr("LP4Attr") & String.Empty
                If cbShowAdvancedFeatures.Checked Then
                    tbCSSAttributes.Visible = True
                Else
                    tbCSSAttributes.Visible = False
                End If
                'Case "NB1left", "NB1left+"
                '    RadEditor1.Visible = True
                '    RadEditor1.Content = (dr("NB1_BodyContent") & String.Empty).ToString.TrimEnd
                '    trLoginPageAdvancedControls.Visible = False
                '    trNoticeBoard1AdvancedControls.Visible = False
                '    trSiteLogoURL.Visible = False
                '    tbCSSAttributes.Text = dr("NB1_BodyAttr") & String.Empty
                '    If cbShowAdvancedFeatures.Checked Then
                '        tbCSSAttributes.Visible = True
                '    Else
                '        tbCSSAttributes.Visible = False
                '    End If
                'Case "NB1top"
                '    RadEditor1.Visible = True
                '    RadEditor1.Content = (dr("NB1_TopContent") & String.Empty).ToString.TrimEnd
                '    trLoginPageAdvancedControls.Visible = False
                '    trNoticeBoard1AdvancedControls.Visible = False
                '    trSiteLogoURL.Visible = False
                '    tbCSSAttributes.Text = dr("NB1_TopAttr") & String.Empty
                '    If cbShowAdvancedFeatures.Checked Then
                '        tbCSSAttributes.Visible = True
                '    Else
                '        tbCSSAttributes.Visible = False
                '    End If
                'Case "NB1bottom"
                '    RadEditor1.Visible = True
                '    RadEditor1.Content = (dr("NB1_BottomContent") & String.Empty).ToString.TrimEnd
                '    trLoginPageAdvancedControls.Visible = False
                '    trNoticeBoard1AdvancedControls.Visible = False
                '    trSiteLogoURL.Visible = False
                '    tbCSSAttributes.Text = dr("NB1_BottomAttr") & String.Empty
                '    If cbShowAdvancedFeatures.Checked Then
                '        tbCSSAttributes.Visible = True
                '    Else
                '        tbCSSAttributes.Visible = False
                '    End If
                'Case "NB1layout"
                '    RadEditor1.Visible = False
                '    trNoticeBoard1AdvancedControls.Visible = True
                '    trLoginPageAdvancedControls.Visible = False
                '    trSiteLogoURL.Visible = False
                '    tbNB1LeftAttributes.Text = dr("NB1LeftAttr") & String.Empty
                '    tbNB1CentreAttributes.Text = dr("NB1CentreAttr") & String.Empty
                '    tbNB1RightAttributes.Text = dr("NB1RightAttr") & String.Empty
                '    tbNB1LeftAttributes.Focus()
                '    tbCSSAttributes.Visible = False
            Case "LPlayout"
                RadEditor1.Visible = False
                trLoginPageAdvancedControls.Visible = True
                trNoticeBoard1AdvancedControls.Visible = False
                trSiteLogoURL.Visible = False
                tbLPLeftAttributes.Text = dr("LPLeftAttr") & String.Empty
                tbLPRightAttributes.Text = dr("LPRightAttr") & String.Empty
                tbLPLeftAttributes.Focus()
                tbCSSAttributes.Visible = False
                Dim nLoginPanelPane As Integer = CInt(dr("LgnPnlPane"))
                ddlLoginBoxPosition.SelectedIndex = nLoginPanelPane
            Case "SiteLogo"
                RadEditor1.Visible = False
                trSiteLogoURL.Visible = True
                trLoginPageAdvancedControls.Visible = False
                trNoticeBoard1AdvancedControls.Visible = False
                tbSiteLogoURL.Text = dr("DefaultRunningHeaderImage") & String.Empty
                tbSiteLogoURL.Focus()
                tbCSSAttributes.Visible = False
        End Select
    End Sub
    
    Protected Sub LoadPagesPanel()
        Dim oDataTable As New DataTable
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter("spASPNET_RotatorGetPagePanelControlsFromCustomerKey", oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Convert.ToInt32(Session("CustomerKey"))
        Try
            reLeft.Content = String.Empty
            reRight.Content = String.Empty
            reCentre.Content = String.Empty
            tbHeaderPanel.Text = String.Empty
            oAdapter.Fill(oDataTable)
            If oDataTable.Rows.Count = 1 Then
                Dim dr As DataRow = oDataTable.Rows(0)
                reLeft.Content = dr("LeftHTML")
                reRight.Content = dr("RightHTML")
                reCentre.Content = dr("CentreHTML")
                tbHeaderPanel.Text = dr("Header")
            Else
                If oDataTable.Rows.Count > 1 Then
                    WebMsgBox.Show("LoadPagesPanel: multiple rows detected - please inform development")
                End If
            End If
        Catch ex As Exception
            WebMsgBox.Show("LoadPagesPanel: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub SaveUploadedImage()
        radImageEditor.SaveEditableImage(Path.GetFileNameWithoutExtension(radImageEditor.ImageUrl), True)
    End Sub
    
    Protected Sub SaveNoticeBoardPanel()
        Dim sLeft As String = reLeft.Content
        Dim sRight As String = reRight.Content
        Dim sCentre As String = reCentre.Content
        Dim sHeader As String = tbHeaderPanel.Text.Trim
        
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_RotatorInsertUpdateTokenContent", oConn)
        oCmd.CommandType = CommandType.StoredProcedure

        Dim paramCustomerKey As New SqlParameter("@CustomerKey", SqlDbType.VarChar)
        paramCustomerKey.Value = Session("CustomerKey")
        oCmd.Parameters.Add(paramCustomerKey)

        Dim paramLeftHTML As New SqlParameter("@LeftHTML", SqlDbType.VarChar)
        paramLeftHTML.Value = sLeft
        oCmd.Parameters.Add(paramLeftHTML)

        Dim paramRightHTML As New SqlParameter("@RightHTML", SqlDbType.VarChar)
        paramRightHTML.Value = sRight
        oCmd.Parameters.Add(paramRightHTML)

        Dim paramCentreHTML As New SqlParameter("@CentreHTML", SqlDbType.VarChar)
        paramCentreHTML.Value = sCentre
        oCmd.Parameters.Add(paramCentreHTML)
        
        Dim paramHeader As New SqlParameter("@Header", SqlDbType.VarChar)
        paramHeader.Value = sHeader
        oCmd.Parameters.Add(paramHeader)
        
        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As SqlException
            WebMsgBox.Show("SaveDynamicTextChanges: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub lnkDeleteDataSource_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkDeleteDataSource.Click
        
        Dim sSQL As String = "Delete from RotatorContent where DataSourceTag = '" & ddlDataSource.SelectedValue & "'"
        Call ExecuteQueryToDataTable(sSQL)
        Call BindDataSourceTags()
        rgDataSourceEditor.Rebind()
        
    End Sub
    
    Protected Sub ddlDataSource_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ddlDataSource.SelectedIndexChanged
        If ddlDataSource.SelectedValue = "+new data source" Then
            tblAddDataSource.Visible = True
            tblDataSourceEditor.Visible = False
            tblRssEditor.Visible = False
            tbDataSource.Focus()
        ElseIf ddlDataSource.SelectedValue = "- select data source -" Then
            tblDataSourceEditor.Visible = False
            tblAddDataSource.Visible = False
            tblRssEditor.Visible = False
        ElseIf ddlDataSource.SelectedValue = "Rss Feed" Then
            tblRssEditor.Visible = True
            tblAddDataSource.Visible = False
            tblDataSourceEditor.Visible = False
            Call LoadRssDatatable()
        Else
            lnkDeleteDataSource.Text = "Delete <b>" & ddlDataSource.SelectedItem.Text & "</b> Data Source"
            tblAddDataSource.Visible = False
            tblRssEditor.Visible = False
            tblDataSourceEditor.Visible = True
            rgDataSourceEditor.Rebind()
        End If
    End Sub
    
    Protected Sub BindDataSourceTags()
        
        ddlDataSource.Items.Clear()
        ddlDataSourceTag.Items.Clear()
        
        Dim oListItemCollection As ListItemCollection = ExecuteQueryToListItemCollection("SELECT DISTINCT DataSourceTag FROM RotatorContent", "DataSourceTag", "DataSourceTag")
        'ddlDataSource.DataSource = LoadImageTags()
        'ddlDataSource.DataTextField = "DataSourceTag"
        'ddlDataSource.DataValueField = "DataSourceTag"
        'ddlDataSource.DataBind()
        ddlDataSource.Items.Add(New ListItem("- select data source -", "- select data source -"))
        'ddlDataSource.Items.Add(New ListItem("Products", "Products"))
        ddlDataSource.Items.Add(New ListItem("Rss Feed", "Rss Feed"))
        
        'ddlDataSourceTag.Items.Add(New ListItem("Products", "Products"))
        ddlDataSourceTag.Items.Add(New ListItem("Rss Feed", "Rss Feed"))
        
        For Each li As ListItem In oListItemCollection
            ddlDataSource.Items.Add(li)
            ddlDataSourceTag.Items.Add(li)
        Next
        ddlDataSource.Items.Insert(ddlDataSource.Items.Count, New ListItem("+new data source", "+new data source"))
        'ddlDataSourceTag.DataSource = LoadImageTags()
        'ddlDataSourceTag.DataTextField = "DataSourceTag"
        'ddlDataSourceTag.DataValueField = "DataSourceTag"
        'ddlDataSourceTag.DataBind()
    End Sub
    
    Protected Function LoadImageTags() As DataTable
        Dim sql As String = "SELECT DISTINCT dataSourceTag FROM RotatorContent"
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand(sql, oConn)
        Dim oDA As SqlDataAdapter = New SqlDataAdapter(sql, gsConn)
        Dim oDataTable As New DataTable
        Try
            oConn.Open()
            oDA.Fill(oDataTable)
        Catch ex As Exception
            WebMsgBox.Show(ex.Message.ToString())
        Finally
            oConn.Close()
        End Try
        Return oDataTable
    End Function
    
    Protected Function LoadDataSourceEditorSource() As DataTable
        
        LoadDataSourceEditorSource = Nothing
        
        Dim oDataTable As New DataTable
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter("spASPNET_RotatorGetContentFromCustomerKey", oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Convert.ToInt32(Session("CustomerKey"))
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@DataSourceTag", SqlDbType.VarChar))
        oAdapter.SelectCommand.Parameters("@DataSourceTag").Value = ddlDataSource.SelectedValue
        Try
            oAdapter.Fill(oDataTable)
            gdt = oDataTable
            LoadDataSourceEditorSource = oDataTable
        Catch ex As Exception
            WebMsgBox.Show(ex.Message.ToString())
        Finally
            oConn.Close()
        End Try

        
    End Function
    
    Protected Sub rgDataSourceEditor_NeedDataSource(ByVal source As Object, ByVal e As Telerik.Web.UI.GridNeedDataSourceEventArgs) Handles rgDataSourceEditor.NeedDataSource
      
        rgDataSourceEditor.DataSource = LoadDataSourceEditorSource()

    End Sub
    
    Protected Sub rgDataSourceEditor_ItemCommand(ByVal source As Object, ByVal e As Telerik.Web.UI.GridCommandEventArgs) Handles rgDataSourceEditor.ItemCommand

        If e.CommandName = "PerformInsert" Then
            
            Insert(e)

        ElseIf e.CommandName = "Update" Then
            Update(e)
            
        ElseIf e.CommandName = "MoveUp" Then
            MoveUp(e)
            
        ElseIf e.CommandName = "MoveDown" Then
            MoveDown(e)
        End If

    End Sub
  
    
    Protected Sub MoveDown(ByVal e As Telerik.Web.UI.GridCommandEventArgs)
        
        LoadDataSourceEditorSource()
    
        Dim nBelowRowPosition As Integer
        Dim nBelowRowID As Integer
        Dim hidPosition As HiddenField = e.Item.FindControl("hidPosition")
        Dim hidID As HiddenField = e.Item.FindControl("hidID")
        Dim nCurrentRowPosition As Integer = Convert.ToInt32(hidPosition.Value)
        If gdt.Rows.Count > 0 AndAlso e.Item.ItemIndex < gdt.Rows.Count - 1 Then
            nBelowRowPosition = gdt.Rows(e.Item.ItemIndex + 1).Item("Position")
            nBelowRowID = gdt.Rows(e.Item.ItemIndex + 1).Item("ID")
            Dim nID As Integer = Convert.ToInt32(hidID.Value)
            nCurrentRowPosition = nCurrentRowPosition + 1
            nBelowRowPosition = nBelowRowPosition - 1
            Dim sb As New StringBuilder
            sb.Append("UPDATE RotatorContent SET Position = " & nCurrentRowPosition & " WHERE ID = " & nID & " ")
            sb.Append("UPDATE RotatorContent SET Position = " & nBelowRowPosition & " WHERE ID = " & nBelowRowID)
            Dim sql As String = sb.ToString
            Dim oConn As New SqlConnection(gsConn)
            Dim oCmd As SqlCommand = New SqlCommand(sql, oConn)
            Try
                oConn.Open()
                oCmd.ExecuteNonQuery()
            Catch ex As Exception
                WebMsgBox.Show(ex.Message.ToString())
            End Try
        End If
        
        rgDataSourceEditor.Rebind()
        
    End Sub
    
    Protected Sub MoveUp(ByVal e As Telerik.Web.UI.GridCommandEventArgs)
        
        LoadDataSourceEditorSource()
    
        Dim nAboveRowPosition As Integer
        Dim nAoveRowID As Integer
        Dim hidPosition As HiddenField = e.Item.FindControl("hidPosition")
        Dim hidID As HiddenField = e.Item.FindControl("hidID")
        Dim nCurrentRowPosition As Integer = Convert.ToInt32(hidPosition.Value)
        If e.Item.ItemIndex > 0 Then
            nAboveRowPosition = gdt.Rows(e.Item.ItemIndex - 1)("Position")
            nAoveRowID = gdt.Rows(e.Item.ItemIndex - 1).Item("ID")
            Dim nID As Integer = Convert.ToInt32(hidID.Value)
            nCurrentRowPosition = nCurrentRowPosition - 1
            nAboveRowPosition = nAboveRowPosition + 1
            Dim sb As New StringBuilder
            sb.Append("UPDATE RotatorContent SET Position = " & nCurrentRowPosition & " WHERE ID = " & nID & " ")
            sb.Append("UPDATE RotatorContent SET Position = " & nAboveRowPosition & " WHERE ID = " & nAoveRowID)
            Dim sql As String = sb.ToString
            Dim oConn As New SqlConnection(gsConn)
            Dim oCmd As SqlCommand = New SqlCommand(sql, oConn)
            Try
                oConn.Open()
                oCmd.ExecuteNonQuery()
            Catch ex As Exception
                WebMsgBox.Show(ex.Message.ToString())
            End Try
            
        End If
        
        rgDataSourceEditor.Rebind()
        
    End Sub
    
    Protected Sub lnkDelete_Click(ByVal sender As Object, ByVal e As EventArgs)

        Dim lnkDelete As LinkButton = sender
        Dim sID As String = lnkDelete.CommandArgument
        Dim sb As New StringBuilder
        sb.Append("Delete from RotatorContent where ID = " & sID)
        Dim sSql As String = sb.ToString
        ExecuteQueryToDataTable(sSql)
        rgDataSourceEditor.Rebind()

    End Sub
    
    Protected Sub Update(ByVal e As Telerik.Web.UI.GridCommandEventArgs)

        If TypeOf e.Item Is GridEditableItem AndAlso e.Item.IsInEditMode Then

            Dim reTitle As RadEditor = e.Item.FindControl("reTitle")
            Dim reContent As RadEditor = e.Item.FindControl("reContent")
            Dim txtImageTag As TextBox = e.Item.FindControl("txtImageTag")
            
            Dim hidID As HiddenField = e.Item.FindControl("hidID")
            Dim nID As Integer = Convert.ToInt32(hidID.Value)
            
            Dim oConn As New SqlConnection(gsConn)
            Dim oCmd As SqlCommand = New SqlCommand("spASPNET_RotatorInsertUpdateRotatorContent", oConn)
            oCmd.CommandType = CommandType.StoredProcedure
            Dim paramID As New SqlParameter("@ID", SqlDbType.BigInt)
            paramID.Value = nID
            oCmd.Parameters.Add(paramID)
            Dim paramCustKey As New SqlParameter("@CustomerKey", SqlDbType.BigInt)
            paramCustKey.Value = Session("CustomerKey")
            oCmd.Parameters.Add(paramCustKey)
                
            Dim paramDataSource As New SqlParameter("@DataSourceTag", SqlDbType.VarChar)
            paramDataSource.Value = ddlDataSource.SelectedValue
            oCmd.Parameters.Add(paramDataSource)
                
            Dim paramTitle As New SqlParameter("@Title", SqlDbType.VarChar)
            paramTitle.Value = reTitle.Content
            oCmd.Parameters.Add(paramTitle)
                
            Dim paramContent As New SqlParameter("@Content", SqlDbType.VarChar)
            paramContent.Value = reContent.Content
            oCmd.Parameters.Add(paramContent)
                
            Dim paramImageTag As New SqlParameter("@ImageTag", SqlDbType.VarChar)
            paramImageTag.Value = txtImageTag.Text.Trim
            oCmd.Parameters.Add(paramImageTag)
                
            Dim paramUrl As New SqlParameter("@Url", SqlDbType.VarChar)
            paramUrl.Value = tbURL.Text
            oCmd.Parameters.Add(paramUrl)
                
            Dim paramPosition As New SqlParameter("@Position", SqlDbType.Int)
            paramPosition.Value = rgDataSourceEditor.Items.Count
            oCmd.Parameters.Add(paramPosition)

            Try
                oConn.Open()
                oCmd.Connection = oConn
                oCmd.ExecuteNonQuery()
            Catch ex As SqlException
                WebMsgBox.Show("rgDataSourceEditor Update: " & ex.Message)
            Finally
                oConn.Close()
            End Try

        End If

    End Sub
    
    Protected Sub Insert(ByVal e As Telerik.Web.UI.GridCommandEventArgs)

        If (TypeOf e.Item Is Telerik.Web.UI.GridEditFormInsertItem) AndAlso e.Item.IsInEditMode Then

            Dim txtImageTag As TextBox = e.Item.FindControl("txtImageTag")
            Dim reTitle As RadEditor = e.Item.FindControl("reTitle")
            Dim reContent As RadEditor = e.Item.FindControl("reContent")
            
            'Dim hidID As HiddenField = e.Item.FindControl("hidID")
            'Dim nID As Integer = Convert.ToInt32(hidID.Value)
            
            Dim oConn As New SqlConnection(gsConn)
            
            Dim oCmd As SqlCommand = New SqlCommand("spASPNET_RotatorInsertUpdateRotatorContent", oConn)
            oCmd.CommandType = CommandType.StoredProcedure
            Dim paramID As New SqlParameter("@ID", SqlDbType.BigInt)
            paramID.Value = -1
            oCmd.Parameters.Add(paramID)
            Dim paramCustKey As New SqlParameter("@CustomerKey", SqlDbType.BigInt)
            paramCustKey.Value = Session("CustomerKey")
            oCmd.Parameters.Add(paramCustKey)
                
            Dim paramDataSource As New SqlParameter("@DataSourceTag", SqlDbType.VarChar)
            paramDataSource.Value = ddlDataSource.SelectedValue
            oCmd.Parameters.Add(paramDataSource)
                
            Dim paramTitle As New SqlParameter("@Title", SqlDbType.VarChar)
            paramTitle.Value = reTitle.Content
            oCmd.Parameters.Add(paramTitle)
                
            Dim paramContent As New SqlParameter("@Content", SqlDbType.VarChar)
            paramContent.Value = reContent.Content
            oCmd.Parameters.Add(paramContent)
                
            Dim paramImageTag As New SqlParameter("@ImageTag", SqlDbType.VarChar)
            paramImageTag.Value = txtImageTag.Text.Trim
            oCmd.Parameters.Add(paramImageTag)
                
            Dim paramUrl As New SqlParameter("@Url", SqlDbType.VarChar)
            paramUrl.Value = tbURL.Text
            oCmd.Parameters.Add(paramUrl)
                
            Dim paramPosition As New SqlParameter("@Position", SqlDbType.Int)
            paramPosition.Value = rgDataSourceEditor.Items.Count + 1
            oCmd.Parameters.Add(paramPosition)

            Try
                oConn.Open()
                oCmd.Connection = oConn
                oCmd.ExecuteNonQuery()
            Catch ex As SqlException
                WebMsgBox.Show("rgDataSourceEditor Insert: " & ex.Message)
            Finally
                oConn.Close()
            End Try

        End If

    End Sub

    
    'Protected Sub Item_Button(ByVal sender As Object, ByVal e As DataGridCommandEventArgs)
    '    Dim tb As TextBox
    '    Dim tbUrl As TextBox
    '    Dim reContent As RadEditor = Nothing
    '    Dim reTitle As RadEditor = Nothing
    '    Select Case e.CommandName
    '        Case "MoveUp"
    '            Call LoadDataset()
    '            If e.Item.ItemIndex > 0 Then
    '                Dim nAboveRowPosition As Integer
    '                Dim nAoveRowID As Integer
    '                Dim hidPosition As HiddenField = e.Item.FindControl("hidPosition")
    '                Dim hidID As HiddenField = e.Item.FindControl("hidID")
    '                Dim nCurrentRowPosition As Integer = Convert.ToInt32(hidPosition.Value)
    '                nAboveRowPosition = gds.Tables(0).Rows(e.Item.ItemIndex - 1).Item("Position")
    '                nAoveRowID = gds.Tables(0).Rows(e.Item.ItemIndex - 1).Item("ID")
    '                Dim nID As Integer = Convert.ToInt32(hidID.Value)
    '                nCurrentRowPosition = nCurrentRowPosition - 1
    '                nAboveRowPosition = nAboveRowPosition + 1
    '                Dim sb As New StringBuilder
    '                sb.Append("UPDATE RotatorContent SET Position = " & nCurrentRowPosition & " WHERE ID = " & nID & " ")
    '                sb.Append("UPDATE RotatorContent SET Position = " & nAboveRowPosition & " WHERE ID = " & nAoveRowID)
    '                Dim sql As String = sb.ToString
    '                Dim oConn As New SqlConnection(gsConn)
    '                Dim oCmd As SqlCommand = New SqlCommand(sql, oConn)
    '                Try
    '                    oConn.Open()
    '                    oCmd.ExecuteNonQuery()
    '                Catch ex As Exception
    '                    WebMsgBox.Show(ex.Message.ToString())
    '                End Try
    '                LoadDataset()
    '                BindNewsEditor()
    '            End If
    '        Case "MoveDown"
    '            Call LoadDataset()
    '            Dim nBelowRowPosition As Integer
    '            Dim nBelowRowID As Integer
    '            Dim hidPosition As HiddenField = e.Item.FindControl("hidPosition")
    '            Dim hidID As HiddenField = e.Item.FindControl("hidID")
    '            Dim nCurrentRowPosition As Integer = Convert.ToInt32(hidPosition.Value)
    '            nBelowRowPosition = gds.Tables(0).Rows(e.Item.ItemIndex + 1).Item("Position")
    '            nBelowRowID = gds.Tables(0).Rows(e.Item.ItemIndex + 1).Item("ID")
    '            Dim nID As Integer = Convert.ToInt32(hidID.Value)
    '            nCurrentRowPosition = nCurrentRowPosition + 1
    '            nBelowRowPosition = nBelowRowPosition - 1
    '            Dim sb As New StringBuilder
    '            sb.Append("UPDATE RotatorContent SET Position = " & nCurrentRowPosition & " WHERE ID = " & nID & " ")
    '            sb.Append("UPDATE RotatorContent SET Position = " & nBelowRowPosition & " WHERE ID = " & nBelowRowID)
    '            Dim sql As String = sb.ToString
    '            Dim oConn As New SqlConnection(gsConn)
    '            Dim oCmd As SqlCommand = New SqlCommand(sql, oConn)
    '            Try
    '                oConn.Open()
    '                oCmd.ExecuteNonQuery()
    '            Catch ex As Exception
    '                WebMsgBox.Show(ex.Message.ToString())
    '            End Try
    '            LoadDataset()
    '            BindNewsEditor()
    '        Case "Insert"
    '            Call LoadDataset()
    '            Dim dr As DataRow
    '            dr = NewDataRow()
    '            gdt.Rows.InsertAt(dr, e.Item.ItemIndex)
    '            dgNews.EditItemIndex = e.Item.ItemIndex
    '            BindNewsEditor()
    '        Case "Edit"
    '            LoadDataset()
    '            'dgNews.Columns(2).Visible = False
    '            If gds.Tables(0).Rows.Count = 0 Then
    '                Dim dr As DataRow
    '                dr = NewDataRow()
    '                gdt.Rows.InsertAt(dr, e.Item.ItemIndex)
    '            End If
    '            dgNews.EditItemIndex = e.Item.ItemIndex
    '            'BindNewsEditor()
    '        Case "Update"
    '            tb = e.Item.Cells(3).Controls(0)
    '            tbUrl = e.Item.FindControl("txtUrl")
    '            reTitle = e.Item.Cells(4).FindControl("reTitle")
    '            reContent = e.Item.Cells(5).FindControl("reContent")
    '            dgNews.EditItemIndex = -1
    '            Dim hidID As HiddenField = e.Item.FindControl("hidID")
    '            Dim nID As Integer = Convert.ToInt32(hidID.Value)
    '            Dim oConn As New SqlConnection(gsConn)
    '            Dim oCmd As SqlCommand = New SqlCommand("spASPNET_RotatorInsertUpdateRotatorContent", oConn)
    '            oCmd.CommandType = CommandType.StoredProcedure
    '            Dim paramID As New SqlParameter("@ID", SqlDbType.BigInt)
    '            paramID.Value = nID
    '            oCmd.Parameters.Add(paramID)
    '            Dim paramCustKey As New SqlParameter("@CustomerKey", SqlDbType.BigInt)
    '            paramCustKey.Value = Session("CustomerKey")
    '            oCmd.Parameters.Add(paramCustKey)
                
    '            Dim paramDataSource As New SqlParameter("@DataSourceTag", SqlDbType.VarChar)
    '            paramDataSource.Value = ddlDataSource.SelectedValue
    '            oCmd.Parameters.Add(paramDataSource)
                
    '            Dim paramTitle As New SqlParameter("@Title", SqlDbType.VarChar)
    '            paramTitle.Value = reTitle.Content
    '            oCmd.Parameters.Add(paramTitle)
                
    '            Dim paramContent As New SqlParameter("@Content", SqlDbType.VarChar)
    '            paramContent.Value = reContent.Content
    '            oCmd.Parameters.Add(paramContent)
                
    '            Dim paramImageTag As New SqlParameter("@ImageTag", SqlDbType.VarChar)
    '            paramImageTag.Value = tb.Text
    '            oCmd.Parameters.Add(paramImageTag)
                
    '            Dim paramUrl As New SqlParameter("@Url", SqlDbType.VarChar)
    '            paramUrl.Value = tbUrl.Text
    '            oCmd.Parameters.Add(paramUrl)
                
    '            Dim paramPosition As New SqlParameter("@Position", SqlDbType.Int)
    '            paramPosition.Value = dgNews.Items.Count
    '            oCmd.Parameters.Add(paramPosition)

    '            Try
    '                oConn.Open()
    '                oCmd.Connection = oConn
    '                oCmd.ExecuteNonQuery()
    '            Catch ex As SqlException
    '                WebMsgBox.Show("Item_Button: " & ex.Message)
    '            Finally
    '                oConn.Close()
    '            End Try
    '            LoadDataset()
    '            BindNewsEditor()
    '        Case "Cancel"
    '            dgNews.EditItemIndex = -1
    '            Call LoadDataset()
    '            Call BindNewsEditor()
    '        Case "Delete"
    '            LoadDataset()
    '            BindNewsEditor()
    '            If gds.Tables(0).Rows.Count > 1 Then
    '                Dim nItemIndex As Integer = e.Item.ItemIndex
    '                Dim dr As DataRow = gds.Tables(0).Rows(nItemIndex)
    '                Dim nID As Integer = Convert.ToInt16(dr("ID"))
    '                Dim sb As New StringBuilder
    '                sb.Append("Delete from RotatorContent where ID = " & nID)
    '                Dim sql As String = sb.ToString
    '                Dim oConn As New SqlConnection(gsConn)
    '                Dim oCmd As SqlCommand = New SqlCommand(sql, oConn)
    '                Try
    '                    oConn.Open()
    '                    oCmd.ExecuteNonQuery()
    '                Catch ex As Exception
    '                    WebMsgBox.Show(ex.Message.ToString())
    '                End Try
    '            End If
    '            gds.Tables(0).Rows.RemoveAt(e.Item.ItemIndex)
    '        Case Else
    '    End Select
    'End Sub
    
    'Protected Function NewDataRow() As DataRow
    '    gdt = gds.Tables(0)
    '    Dim row As DataRow = gdt.NewRow
    '    Dim nMaxID As Integer
    '    row(0) = -1
    '    row(1) = Convert.ToInt32(Session("CustomerKey"))
    '    row(2) = "+++ Image Tag +++"
    '    row(3) = ddlDataSource.SelectedValue
    '    row(4) = DateTime.Now.ToString
    '    row(5) = "+++ Title +++"
    '    row(6) = "+++ Content +++"
    '    row(7) = "+++ Url +++"
    '    If gdt.Rows.Count > 0 Then
    '        nMaxID = Convert.ToInt16(gdt.DefaultView(0)("Position")) + 1
    '        row(8) = nMaxID
    '    Else
    '        nMaxID = 1
    '        row(8) = nMaxID
    '    End If
    '    gdt.DefaultView.Sort = "Position DESC"
    '    Return row
    'End Function
    
    Protected Sub AddPredefinedDataSource()
        Dim oDataTable As New DataTable
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter("spASPNET_Customer_GetRotatorAds", oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Convert.ToInt32(Session("CustomerKey"))
        Try
            oAdapter.Fill(oDataTable)
        Catch ex As Exception
            WebMsgBox.Show(ex.Message.ToString())
        Finally
            oConn.Close()
        End Try
    End Sub
    
    'Protected Sub LoadDataSourceDataset()
        
        
    'End Sub

    'Protected Sub PageIndexChanged(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs)
    '    dgNews.CurrentPageIndex = e.NewPageIndex
    '    LoadDataset()
    '    BindNewsEditor()
    'End Sub
    
    'Protected Sub BindNewsEditor()
    '    dgNews.DataSource = gds
    '    dgNews.DataBind()
    '    If dgNews.Items.Count > 0 Then
    '        If ddlDataSource.SelectedValue = "Products" Then
    '            dgNews.Columns(0).Visible = False
    '            dgNews.Columns(1).Visible = False
    '            dgNews.Columns(2).Visible = False
    '            dgNews.Columns(7).Visible = False
    '        Else
    '            dgNews.Items(0).FindControl("lnkMoveUp").Visible = False
    '            dgNews.Items(dgNews.Items.Count - 1).FindControl("lnkMoveDown").Visible = False
    '            dgNews.Columns(0).Visible = True
    '            dgNews.Columns(1).Visible = True
    '            dgNews.Columns(2).Visible = True
    '            dgNews.Columns(7).Visible = True
    '        End If
    '    Else
    '        Dim dr As DataRow = NewDataRow()
    '        gdt.Rows.Add(dr)
    '        dgNews.DataSource = gds
    '        dgNews.DataBind()
    '    End If
    'End Sub
    
    Protected Sub btnAddRotator_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim btn As Button = sender
        If btn.ID = "btnAddRotatorLeft" Then
            If reLeft.Content.Trim <> String.Empty Then
                reLeft.Content = reLeft.Content & "<br>"
            End If
            reLeft.Content = reLeft.Content & "[Rotator: "
            If chkImage.Checked Then
                reLeft.Content = reLeft.Content & "Image"
            End If
            If chkTitle.Checked Then
                reLeft.Content = reLeft.Content & "," & "Title"
            End If
            If chkContent.Checked Then
                reLeft.Content = reLeft.Content & "," & "Content"
            End If
            
            ''''''''''''''''''''''''''''''''''''' Setting a data source ''''''''''''''''''''''''''''''''''
            reLeft.Content = reLeft.Content & ";" & " DataSource=" & ddlDataSourceTag.SelectedValue.ToString()
            reLeft.Content = reLeft.Content & ";" & " css=" & tbCss.Text.Trim
            If tbRotatorHeight.Text.Trim <> String.Empty Then
                reLeft.Content = reLeft.Content & ";" & " Height=" & tbRotatorHeight.Text.Trim
            End If
            If tbRotatorWidth.Text.Trim <> String.Empty Then
                reLeft.Content = reLeft.Content & ";" & " Width=" & tbRotatorWidth.Text.Trim
            End If
            reLeft.Content = reLeft.Content & ";" & " RotatorType=" & ddlRotatorType.SelectedValue
            reLeft.Content = reLeft.Content & ";" & " ScrollDirection=" & ddlScrollDirection.SelectedValue
            reLeft.Content = reLeft.Content & ";" & " AnimationType=" & ddlAnimationType.SelectedValue
            If tbScrollDuration.Text.Trim <> String.Empty Then    ' ??????????????????
                reLeft.Content = reLeft.Content & ";" & " ScrollDuration=" & tbScrollDuration.Text.Trim
            End If
             
            
            reLeft.Content = reLeft.Content & "]"
        ElseIf btn.ID = "btnAddRotatorCentre" Then
            If reCentre.Content.Trim <> String.Empty Then
                reCentre.Content = reCentre.Content & "<br>"
            End If
            reCentre.Content = reCentre.Content & "[Rotator: "
            If chkImage.Checked Then
                reCentre.Content = reCentre.Content & "Image"
            End If
            If chkTitle.Checked Then
                reCentre.Content = reCentre.Content & "," & "Title"
            End If
            If chkContent.Checked Then
                reCentre.Content = reCentre.Content & "," & "Content"
            End If
            ''''''''''''''''''''''''''''''''''''' Setting a data source ''''''''''''''''''''''''''''''''''
            reCentre.Content = reCentre.Content & ";" & " Data Source=" & ddlDataSourceTag.SelectedValue.ToString()
            reCentre.Content = reCentre.Content & ";" & " css=" & tbCss.Text.Trim
            If tbRotatorHeight.Text.Trim <> String.Empty Then
                reCentre.Content = reCentre.Content & ";" & " Height=" & tbRotatorHeight.Text.Trim
            End If
            If tbRotatorWidth.Text.Trim <> String.Empty Then
                reCentre.Content = reCentre.Content & ";" & " Width=" & tbRotatorWidth.Text.Trim
            End If
            reCentre.Content = reCentre.Content & ";" & " RotatorType=" & ddlRotatorType.SelectedValue
            reCentre.Content = reCentre.Content & ";" & " ScrollDirection=" & ddlScrollDirection.SelectedValue
            reCentre.Content = reCentre.Content & ";" & " AnimationType=" & ddlAnimationType.SelectedValue
            If tbScrollDuration.Text.Trim <> String.Empty Then
                reCentre.Content = reCentre.Content & ";" & " ScrollDuration=" & tbScrollDuration.Text.Trim
            End If
            reCentre.Content = reCentre.Content & "]"
        ElseIf btn.ID = "btnAddRotatorRight" Then
            If reRight.Content.Trim <> String.Empty Then
                reRight.Content = reRight.Content & "<br>"
            End If
            reRight.Content = reRight.Content & "[Rotator: "
            If chkImage.Checked Then
                reRight.Content = reRight.Content & "Image"
            End If
            If chkTitle.Checked Then
                reRight.Content = reRight.Content & "," & "Title"
            End If
            If chkContent.Checked Then
                reRight.Content = reRight.Content & "," & "Content"
            End If
            
            ''''''''''''''''''''''''''''''''''''' Setting a data source ''''''''''''''''''''''''''''''''''
            reRight.Content = reRight.Content & ";" & " Data Source=" & ddlDataSourceTag.SelectedValue.ToString()
            reRight.Content = reRight.Content & ";" & " css=" & tbCss.Text.Trim
            If tbRotatorHeight.Text.Trim <> String.Empty Then
                reRight.Content = reRight.Content & ";" & " Height=" & tbRotatorHeight.Text.Trim
            End If
            If tbRotatorWidth.Text.Trim <> String.Empty Then
                reRight.Content = reRight.Content & ";" & " Width=" & tbRotatorWidth.Text.Trim
            End If
            reRight.Content = reRight.Content & ";" & " RotatorType=" & ddlRotatorType.SelectedValue
            reRight.Content = reRight.Content & ";" & " ScrollDirection=" & ddlScrollDirection.SelectedValue
            reRight.Content = reRight.Content & ";" & " AnimationType=" & ddlAnimationType.SelectedValue
            If tbScrollDuration.Text.Trim <> String.Empty Then
                reRight.Content = reRight.Content & ";" & " ScrollDuration=" & tbScrollDuration.Text.Trim
            End If
            reRight.Content = reRight.Content & "]"
        ElseIf btn.ID = "btnAddRotatorHeader" Then
            If tbHeaderPanel.Text.Trim <> String.Empty Then
                tbHeaderPanel.Text = tbHeaderPanel.Text & "<br>"
            End If
            tbHeaderPanel.Text = tbHeaderPanel.Text & "[Rotator: "
            If chkImage.Checked Then
                tbHeaderPanel.Text = tbHeaderPanel.Text & "Image"
            End If
            If chkTitle.Checked Then
                tbHeaderPanel.Text = tbHeaderPanel.Text & "," & "Title"
            End If
            If chkContent.Checked Then
                tbHeaderPanel.Text = tbHeaderPanel.Text & "," & "Content"
            End If
            ''''''''''''''''''''''''''''''''''''' Setting a data source ''''''''''''''''''''''''''''''''''
            tbHeaderPanel.Text = tbHeaderPanel.Text & ";" & " Data Source=" & ddlDataSourceTag.SelectedValue.ToString()
            If tbRotatorHeight.Text.Trim <> String.Empty Then
                tbHeaderPanel.Text = tbHeaderPanel.Text & ";" & " Height=" & tbRotatorHeight.Text.Trim
            End If
            If tbRotatorWidth.Text.Trim <> String.Empty Then
                tbHeaderPanel.Text = tbHeaderPanel.Text & ";" & " Width=" & tbRotatorWidth.Text.Trim
            End If
            tbHeaderPanel.Text = tbHeaderPanel.Text & ";" & " RotatorType=" & ddlRotatorType.SelectedValue
            tbHeaderPanel.Text = tbHeaderPanel.Text & ";" & " ScrollDirection=" & ddlScrollDirection.SelectedValue
            tbHeaderPanel.Text = tbHeaderPanel.Text & ";" & " AnimationType=" & ddlAnimationType.SelectedValue
            If tbScrollDuration.Text.Trim <> String.Empty Then
                tbHeaderPanel.Text = tbHeaderPanel.Text & ";" & " ScrollDuration=" & tbScrollDuration.Text.Trim
            End If
            tbHeaderPanel.Text = tbHeaderPanel.Text & "]"
        End If
    End Sub
    
    Protected Sub radAsyncUpload_FileUploaded(sender As Object, e As FileUploadedEventArgs)
        imgUploadedImage.Visible = False
        lblUploadedImgMessage.Text = String.Empty
        Dim sSQL As String
        tbImageTag.Text = tbImageTag.Text.Trim.Replace(" ", "_")
        sSQL = "SELECT * FROM RotatorImages WHERE CustomerKey = " & Session("CustomerKey") & " AND ImageTag = '" & tbImageTag.Text.Replace("'", "''") & "'"
        Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
        If dt.Rows.Count > 0 Then
            WebMsgBox.Show("An image with that tag name already exists. Please select another tag name for the image.")
            tbImageTag.Focus()
            Exit Sub
        End If
        sSQL = "SELECT MAX(Position) FROM RotatorImages WHERE CustomerKey = " & Session("CustomerKey").ToString()
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As New SqlCommand(sSQL, oConn)
        Dim nPosition As Integer = 1
        Dim sCustomerKey As String = ""
        Dim sPosition As String = ""
        Dim sFileName As String = ""
        Try
            oConn.Open()
            If Not IsDBNull(oCmd.ExecuteScalar) Then
                nPosition = Convert.ToInt16(oCmd.ExecuteScalar())
                nPosition = nPosition + 1
            End If
        Catch ex As Exception
            WebMsgBox.Show(ex.Message.ToString())
        Finally
            oConn.Close()
        End Try
        Dim nCustomerKey = Convert.ToInt16(Session("CustomerKey"))
        If nCustomerKey < 100 Then
            sCustomerKey = "00" + Session("CustomerKey").ToString()
        ElseIf nCustomerKey < 1000 Then
            sCustomerKey = "0" + Session("CustomerKey").ToString()
        Else
            sCustomerKey = Session("CustomerKey").ToString()
        End If
        If nPosition < 100 Then
            sPosition = "00" + nPosition.ToString()
        ElseIf nPosition < 1000 Then
            sPosition = "0" + nPosition.ToString()
        Else
            sPosition = nPosition.ToString()
        End If
        sFileName = sCustomerKey + sPosition + ".jpg"
        Try
            oCmd = New SqlCommand("spASPNET_RotatorInsertUpdateCustomerImages", oConn)
            oCmd.CommandType = CommandType.StoredProcedure

            oCmd.Parameters.Add(New SqlParameter("@CustomerImageID", SqlDbType.Int))
            oCmd.Parameters("@CustomerImageID").Value = -1
                
            oCmd.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
            oCmd.Parameters("@CustomerKey").Value = Session("CustomerKey")
                
            oCmd.Parameters.Add(New SqlParameter("@FileName", SqlDbType.VarChar, 50))
            oCmd.Parameters("@FileName").Value = sFileName
            
            oCmd.Parameters.Add(New SqlParameter("@Position", SqlDbType.Int))
            oCmd.Parameters("@Position").Value = nPosition
            
            oCmd.Parameters.Add(New SqlParameter("@ImageTag", SqlDbType.VarChar, 50))
            oCmd.Parameters("@ImageTag").Value = tbImageTag.Text.Trim
            
            oCmd.Parameters.Add(New SqlParameter("@Notes", SqlDbType.VarChar, 50))
            oCmd.Parameters("@Notes").Value = ""
            
            oConn.Open()
            oCmd.ExecuteNonQuery()
        Catch ex As SqlException
            WebMsgBox.Show("RestoreDefaultSettings: " & ex.Message)
        Finally
            oConn.Close()
        End Try
        
        Try
            e.File.SaveAs(MapPath(UPLOADED_IMAGES_PATH) + sFileName)
            imgUploadedImage.Visible = True
            imgUploadedImage.ImageUrl = UPLOADED_IMAGES_PATH + sFileName
            lblUploadedImgMessage.Text = "Image uploaded successfully."
        Catch ex As Exception
            WebMsgBox.Show("Error uploading image. Please contact development.")
            lblUploadedImgMessage.Text = ex.Message.ToString
        End Try
    End Sub
    
    Protected Sub chkHeaderImage_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs)
        Dim chkHeaderImage As CheckBox = sender
        Dim imgCustomer As Image = chkHeaderImage.NamingContainer.FindControl("imgCustomer")
        Dim sql As String
        Dim sb As New StringBuilder
        sb.Append("Update UserProfile set RunningHeaderImage = '" & imgCustomer.ImageUrl & "' where [Key] = " & Session("UserKey"))
        sql = sb.ToString
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As New SqlCommand(sql, oConn)
        Try
            oConn.Open()
            oCmd.ExecuteNonQuery()
            WebMsgBox.Show("This image has been set as the header image. You must log out and log back in see the new header image.")
        Catch ex As Exception
            WebMsgBox.Show(ex.Message.ToString())
        Finally
            oConn.Close()
        End Try
    End Sub
   
    Protected Sub imgCustomer_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim imgCustomer As Image = sender
        radImageEditor.ImageUrl = imgCustomer.ImageUrl
        pnlShowImages.Visible = False
        pnlImageEditor.Visible = True
    End Sub
    
    Protected Sub imgUploadedImage_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim imgUploadedImage As Image = sender
        radImageEditor.ImageUrl = imgUploadedImage.ImageUrl
        pnlUploadImage.Visible = False
        pnlImageEditor.Visible = True
    End Sub
    
    Protected Sub LoadImages()
        Dim oDataTable As New DataTable
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter("spASPNET_RotatorGetAllCustomerImages", oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Convert.ToInt32(Session("CustomerKey"))
        oAdapter.Fill(oDataTable)
        dlCustomerImages.DataSource = oDataTable
        dlCustomerImages.DataBind()
    End Sub
    
    Protected Sub dlCustomerImages_ItemDataBound(ByVal sender As Object, ByVal e As DataListItemEventArgs) Handles dlCustomerImages.ItemDataBound
        Dim chkHeaderImage As CheckBox = e.Item.FindControl("chkHeaderImage")
        Dim imgCustomer As ImageButton = e.Item.FindControl("imgCustomer")
        If chkHeaderImage IsNot Nothing And imgCustomer IsNot Nothing Then
            If Session("RunningHeaderImage").ToString() = imgCustomer.ImageUrl Then
                chkHeaderImage.Checked = True
            End If
        End If
    End Sub
    
    Protected Sub HideAllPanels()
        pnlLoginPageEditor.Visible = False
        pnlDataSourceEditor.Visible = False
        pnlStyleSheetEditor.Visible = False
        pnlNoticeBoard.Visible = False
        pnlUploadImage.Visible = False
        pnlShowImages.Visible = False
        pnlImageEditor.Visible = False
        tblRssEditor.Visible = False
        accordion.Visible = False
    End Sub
    
    Protected Sub btnShowImages_Click(ByVal sender As Object, ByVal e As EventArgs)
        Call HideAllPanels()
        pnlShowImages.Visible = True
        Call LoadImages()
    End Sub
    
    Protected Sub btnUploadImages_Click(ByVal sender As Object, ByVal e As EventArgs)
        Call HideAllPanels()
        pnlUploadImage.Visible = True
        tbImageTag.Text = String.Empty
        imgUploadedImage.Visible = False
    End Sub
    
    Sub btnLoginPageEditor_Click(ByVal sender As Object, ByVal e As EventArgs)
        Call HideAllPanels()
        pnlLoginPageEditor.Visible = True
    End Sub
    
    Sub btnDataSourceEditor_Click(ByVal sender As Object, ByVal e As EventArgs)
        Call HideAllPanels()
        pnlDataSourceEditor.Visible = True
    End Sub
    
    Protected Sub btnStyleSheetEditor_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call HideAllPanels()
        pnlStyleSheetEditor.Visible = True
    End Sub
    
    Protected Sub btnHelp_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call HideAllPanels()
        accordion.Visible = True
    End Sub
    
    Protected Sub btnDynamicControls_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call HideAllPanels()
        pnlNoticeBoard.Visible = True
        Call LoadPagesPanel()
    End Sub

    Protected Sub rbNoticeBoardPage_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call InitArea()
    End Sub

    Protected Sub rbLoginPage_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call InitArea()
    End Sub

    Protected Sub ddlAreaSelector_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call InitArea()
    End Sub
    
    Protected Sub btnSaveChanges_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If pnlLoginPageEditor.Visible Then
            Call SaveLoginPageChanges()
        ElseIf pnlDataSourceEditor.Visible Then
            Call SaveRssDatatable()
        ElseIf pnlNoticeBoard.Visible Then
            Call SaveNoticeBoardPanel()
        ElseIf pnlImageEditor.Visible Then
            Call SaveUploadedImage()
        End If
    End Sub
    
    Protected Sub SaveDataSourceName()
        Dim liDataSource As ListItem = ddlDataSource.Items.FindByText(tbDataSource.Text.Trim)
        If liDataSource Is Nothing Then
            ddlDataSource.Items.Insert(ddlDataSource.Items.Count - 1, tbDataSource.Text.Trim)
            ddlDataSource.SelectedValue = tbDataSource.Text
        Else
            WebMsgBox.Show("A data source with that name already exists - please choose another name")
            tbDataSource.Text = String.Empty
        End If
    End Sub
    
    Protected Sub btnSaveDataSource_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnSaveDataSource.Click
        Call SaveDataSourceName()
        tblAddDataSource.Visible = False
        tblDataSourceEditor.Visible = True
    End Sub
    
    Protected Sub cbShowAdvancedFeatures_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim cb As CheckBox = sender
        If cb.Checked Then
            lblLegendCSSAttributes.Visible = True
            tbCSSAttributes.Visible = True
            lblLegendLoginBoxPosition.Visible = True
            ddlLoginBoxPosition.Visible = True
            trLoginPageAdvancedControls.Visible = True
            trNoticeBoard1AdvancedControls.Visible = True
        Else
            lblLegendCSSAttributes.Visible = False
            tbCSSAttributes.Visible = False
            lblLegendLoginBoxPosition.Visible = False
            ddlLoginBoxPosition.Visible = False
            trLoginPageAdvancedControls.Visible = False
            trNoticeBoard1AdvancedControls.Visible = False
        End If
        Call InitAreaSelector()
    End Sub
    
    Protected Sub btnRestoreDefaultStyleSheet_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        ' My.Computer.FileSystem.CopyFile(MapPath(STYLESHEET_FILENAME_DEFAULT), MapPath(STYLESHEET_FILENAME_WORKING), True)
        Session("StyleSheetPath") = DEFAULT_STYLESHEET_PATH
        Call PopulateCSSEditor()
    End Sub
    
    Protected Sub btnSaveStyleSheetChanges_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        ' My.Computer.FileSystem.WriteAllText(MapPath(STYLESHEET_FILENAME_WORKING), tbStyleSheet.Text, False)
        Dim sStyleSheetPath As String = "~\css\sprint_" & sGetPath() & ".css"
        If sStyleSheetPath <> String.Empty Then
            My.Computer.FileSystem.WriteAllText(Request.MapPath(sStyleSheetPath), tbStyleSheet.Text, False)
            Session("StyleSheetPath") = sStyleSheetPath
            Call SetStyleSheet()
        End If
    End Sub
    
    Protected Sub PopulateCSSEditor()
        ' tbStyleSheet.Text = My.Computer.FileSystem.ReadAllText(MapPath(STYLESHEET_FILENAME_WORKING))
        If Not IsNothing(Session("StyleSheetPath")) Then
            tbStyleSheet.Text = My.Computer.FileSystem.ReadAllText(Request.MapPath(Session("StyleSheetPath")))
        End If
    End Sub
    
    Protected Function sGetPath() As String
        Dim sPathInfo As String = Request.Path
        sGetPath = String.Empty
        If sPathInfo <> String.Empty Then
            sPathInfo = sPathInfo.Substring(1)
            Dim sPos As Integer = sPathInfo.IndexOf("/")
            If sPos > 0 Then
                sGetPath = sPathInfo.Substring(0, sPos)
            End If
        End If
    End Function
    
    Protected Sub lnkbtnRestoreDefaultSettings_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call RestoreDefaultSettings()
    End Sub
    
    Protected Sub RestoreDefaultSettings()
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As New SqlCommand("spASPNET_SiteContent", oConn)
        Try
            oCmd.CommandType = CommandType.StoredProcedure
            oCmd.Parameters.Add(New SqlParameter("@Action", SqlDbType.NVarChar, 50))
            oCmd.Parameters("@Action").Value = "RESET"
            oCmd.Parameters.Add(New SqlParameter("@SiteKey", SqlDbType.Int))
            oCmd.Parameters("@SiteKey").Value = Session("SiteKey")
            oCmd.Parameters.Add(New SqlParameter("@ContentType", SqlDbType.NVarChar, 50))
            oCmd.Parameters("@ContentType").Value = "RESET"
            oConn.Open()
            oCmd.ExecuteNonQuery()
        Catch ex As SqlException
            WebMsgBox.Show("RestoreDefaultSettings: " & ex.Message)
        Finally
            oConn.Close()
        End Try
        RadEditor1.Content = String.Empty
    End Sub

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
    
#Region "Preview Rotator"
    
    Protected Sub ConfigureRotatorPreview()
        Dim sAnimationType As String = ddlAnimationType.SelectedValue
        Dim sScrollDirection As String = ddlScrollDirection.SelectedValue
        Dim sRotatorType As String = ddlRotatorType.SelectedValue
        If sAnimationType.ToLower() = "none" Then
            rrPreview.SlideShowAnimation.Type = Telerik.Web.UI.Rotator.AnimationType.None
        ElseIf sAnimationType.ToLower() = "fade" Then
            rrPreview.SlideShowAnimation.Type = Telerik.Web.UI.Rotator.AnimationType.Fade
        ElseIf sAnimationType.ToLower() = "pulse" Then
            rrPreview.SlideShowAnimation.Type = Telerik.Web.UI.Rotator.AnimationType.Pulse
        ElseIf sAnimationType.ToLower() = "crossfade" Then
            rrPreview.SlideShowAnimation.Type = Telerik.Web.UI.Rotator.AnimationType.CrossFade
        End If
        
        If sScrollDirection.ToLower() = "up" Then
            rrPreview.ScrollDirection = RotatorScrollDirection.Up
        ElseIf sScrollDirection.ToLower() = "down" Then
            rrPreview.ScrollDirection = RotatorScrollDirection.Down
        ElseIf sScrollDirection.ToLower() = "left" Then
            rrPreview.ScrollDirection = RotatorScrollDirection.Left
        ElseIf sScrollDirection.ToLower() = "right" Then
            rrPreview.ScrollDirection = RotatorScrollDirection.Right
        Else
            rrPreview.ScrollDirection = RotatorScrollDirection.Up
        End If
      
        If sRotatorType.ToLower() = "buttons" Then
            rrPreview.RotatorType = RotatorType.Buttons
        ElseIf sRotatorType.ToLower() = "automaticadvance" Then
            rrPreview.RotatorType = RotatorType.AutomaticAdvance
        ElseIf sRotatorType.ToLower() = "buttonsover" Then
            rrPreview.RotatorType = RotatorType.ButtonsOver
        ElseIf sRotatorType.ToLower() = "carousel" Then
            rrPreview.RotatorType = RotatorType.Carousel
        ElseIf sRotatorType.ToLower() = "carouselbuttons" Then
            rrPreview.RotatorType = RotatorType.CarouselButtons
        ElseIf sRotatorType.ToLower() = "coverflow" Then
            rrPreview.RotatorType = RotatorType.CoverFlow
        ElseIf sRotatorType.ToLower() = "coverflowbuttons" Then
            rrPreview.RotatorType = RotatorType.CoverFlowButtons
        ElseIf sRotatorType.ToLower() = "slideshow" Then
            rrPreview.RotatorType = RotatorType.SlideShow
        ElseIf sRotatorType.ToLower() = "slideshowbuttons" Then
            rrPreview.RotatorType = RotatorType.SlideShowButtons
        Else
            rrPreview.RotatorType = RotatorType.Carousel
        End If
        
        If tbScrollDuration.Text.Trim <> String.Empty Then
            rrPreview.ScrollDuration = tbScrollDuration.Text.Trim
        End If
        
        If tbRotatorHeight.Text.Trim <> String.Empty Then
            rrPreview.Height = tbRotatorHeight.Text.Trim
        End If
        
        If tbRotatorWidth.Text.Trim <> String.Empty Then
            rrPreview.Width = tbRotatorWidth.Text.Trim
        End If
        
        rrPreview.ItemHeight = 50
        
    End Sub
    
    Protected Sub btnPreview_Click(ByVal sender As Object, ByVal e As EventArgs)
        If chkContent.Checked Or chkImage.Checked Or chkTitle.Checked Then
            Dim IListOfArgs As New List(Of String)
            If chkTitle.Checked Then
                IListOfArgs.Add("Title")
            End If
            If chkContent.Checked Then
                IListOfArgs.Add("Content")
            End If
            If ddlDataSourceTag.SelectedValue.ToLower <> "rss" Then
                If chkImage.Checked Then
                    IListOfArgs.Add("Image")
                End If
            End If
            Dim nArrayLength As Integer = IListOfArgs.Count - 1
            Dim sArgs(nArrayLength) As String
            Dim i As Integer = 0
            For Each arg As String In IListOfArgs
                sArgs(i) = arg
                i += 1
            Next
            rrPreview.ItemTemplate = New RadRotatorTemplate(sArgs)
        End If
        ConfigureRotatorPreview()
        rrPreview.DataSource = BindRotator(ddlDataSourceTag.SelectedValue)
        rrPreview.DataBind()
    End Sub
    
    Protected Function BindRotator(ByVal sDataSourceTag As String) As DataTable
        Dim oDataTable As New DataTable
        BindRotator = Nothing
        If sDataSourceTag.ToLower = "rss" Then
            Dim dt As DataTable = CreateProductsDataTable()
            Dim dr As DataRow
            Dim nRssCount As Integer = 0
            Dim sRssUrl As String = String.Empty
            Dim nCustomerKey As Integer = Convert.ToInt64(Session("CustomerKey"))
            If nCustomerKey > 0 Then
                Dim sQuery As String = "SELECT RSSURL, RSSCount FROM RSSFeed WHERE CustomerKey = " & nCustomerKey
                'Dim oDataTable As New DataTable
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
                If oDataTable IsNot Nothing AndAlso oDataTable.Rows.Count <> 0 Then
                    dr = oDataTable.Rows(0)
                    If Not IsDBNull(dr("RssCount")) Then
                        nRssCount = Convert.ToInt32(dr("RssCount"))
                    End If
                    If Not IsDBNull(dr("RssUrl")) Then
                        sRssUrl = dr("RssUrl").ToString()
                    End If
                End If
            End If
        
            Dim reader As XmlReader = XmlReader.Create(sRssUrl)
            Dim LoadRSSFeed As SyndicationFeed = SyndicationFeed.Load(reader)
            
            Dim sf As SyndicationFeed = LoadRSSFeed
            If Not sf Is Nothing Then
                For Each item As SyndicationItem In sf.Items
                    'Dim item As SyndicationItem = sf.Items(i)
                    dr = dt.NewRow
                    dr("Title") = item.Title.Text
                    dr("Date") = Convert.ToDateTime(item.PublishDate.LocalDateTime).ToString("dd-MMM-yyyy HH:mm")
                    dr("Content") = item.Summary.Text
                    If Not item.Links Is Nothing Then
                        For Each sl As SyndicationLink In item.Links
                            dr("BaseURI") = sl.Uri.AbsoluteUri
                        Next
                    End If
                    Dim bFirst As Boolean = True
                    For Each sc As SyndicationCategory In item.Categories
                        If Not bFirst Then
                            dr("Categories") += ", "
                        End If
                        'dr("Categories") += sc.Name
                        dr("Categories") += sc.Name
                        bFirst = False
                    Next
                    dt.Rows.Add(dr)
                    BindRotator = dt
                    'If i = CInt(nRssCount) - 1 Then
                    '    Exit For
                    'End If
                Next
            End If
        Else
            Dim oConn As New SqlConnection(gsConn)
            Dim oAdapter As New SqlDataAdapter("spASPNET_RotatorGetContentFromCustomerKey", oConn)
            oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
            oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Convert.ToInt32(Session("CustomerKey"))
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@DataSourceTag", SqlDbType.VarChar))
            oAdapter.SelectCommand.Parameters("@DataSourceTag").Value = sDataSourceTag
            Try
                oAdapter.Fill(oDataTable)
                BindRotator = oDataTable
            Catch ex As Exception
                WebMsgBox.Show("GetPageContent: " & ex.Message)
            Finally
                oConn.Close()
            End Try
        End If
    End Function
    
#End Region
    
#Region "Products Editor"
    
    Protected Sub LoadRssDatatable()
        Dim nCustomerKey As Integer = Convert.ToInt64(Session("CustomerKey"))
        If nCustomerKey > 0 Then
            Dim sQuery As String = "SELECT RSSURL, RSSCount FROM RotatorRssFeed WHERE CustomerKey = " & nCustomerKey
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
            If oDataTable IsNot Nothing AndAlso oDataTable.Rows.Count <> 0 Then
                Dim dr As DataRow = oDataTable.Rows(0)
                If Not IsDBNull(dr("RssCount")) Then
                    tbEntryCount.Text = Convert.ToInt32(dr("RssCount"))
                End If
                If Not IsDBNull(dr("RssUrl")) Then
                    tbURL.Text = dr("RssUrl").ToString()
                End If
            End If
        End If
        'Dim fs As New System.IO.FileStream(MapPath(gsXMLProductsContentFilePath), System.IO.FileMode.Open)
        'pdsProducts = New DataSet("ProductSettings")
        'pdsProducts.ReadXml(fs)
        'fs.Close()
    End Sub
    
    Protected Sub SaveRssDatatable()
        Dim nCustomerKey As Integer = Convert.ToInt64(Session("CustomerKey"))
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("RotatorInsertUpdateRssFeedUrl", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        oCmd.Parameters.Add(New SqlParameter("@RssFeedUrl", SqlDbType.VarChar, 255))
        oCmd.Parameters("@RssFeedUrl").Value = tbURL.Text.Trim
        oCmd.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
        oCmd.Parameters("@CustomerKey").Value = Session("CustomerKey")
        oCmd.Parameters.Add(New SqlParameter("@RssCount", SqlDbType.Int))
        oCmd.Parameters("@RssCount").Value = tbEntryCount.Text.Trim
        Try
            oConn.Open()
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            WebMsgBox.Show("Error in ExecuteNonQuery executing: " & " : " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
    
    'Protected Sub btnGo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnGo.Click
    '    Call DisplayRSSFeed()
    '    Dim dt As DataTable = ExtractFromRSSFeed()
    '    pdsProducts = New DataSet()
    '    pdsProducts.DataSetName = "ProductSettings"
    '    pdsProducts.Tables.Add(dt)
    '    gvRSSFeed.DataSource = dt
    '    gvRSSFeed.DataBind()
    'End Sub
    
    Protected Function CreateProductsDataTable() As DataTable
        Dim oDataTable As New DataTable("RotatorRssFeed")
        oDataTable.Columns.Add(New DataColumn("Title", GetType(String)))
        oDataTable.Columns.Add(New DataColumn("Content", GetType(String)))
        oDataTable.Columns.Add(New DataColumn("ImageTag", GetType(String)))
        oDataTable.Columns.Add(New DataColumn("Date", GetType(String)))
        oDataTable.Columns.Add(New DataColumn("BaseUri", GetType(String)))
        oDataTable.Columns.Add(New DataColumn("Categories", GetType(String)))
        CreateProductsDataTable = oDataTable
    End Function
    
    Protected Function ExtractFromRSSFeed() As DataTable
        Dim dt As DataTable = CreateProductsDataTable()
        Dim dr As DataRow
        Dim sf As SyndicationFeed = LoadRSSFeed()
        If Not sf Is Nothing Then
            For Each item As SyndicationItem In sf.Items
                'Dim item As SyndicationItem = sf.Items(i)
                dr = dt.NewRow
                dr("Title") = item.Title.Text
                dr("Date") = Convert.ToDateTime(item.PublishDate.LocalDateTime).ToString("dd-MMM-yyyy HH:mm")
                dr("Content") = item.Summary.Text
                If Not item.Links Is Nothing Then
                    For Each sl As SyndicationLink In item.Links
                        dr("BaseURI") = sl.Uri.AbsoluteUri
                    Next
                End If
                Dim bFirst As Boolean = True
                For Each sc As SyndicationCategory In item.Categories
                    If Not bFirst Then
                        dr("Categories") += ", "
                    End If
                    'dr("Categories") += sc.Name
                    dr("Categories") += sc.Name
                    bFirst = False
                Next
                dt.Rows.Add(dr)
                'If i = CInt(tbEntryCount.Text) - 1 Then
                '    Exit For
                'End If
            Next
        End If
        ExtractFromRSSFeed = dt
    End Function
    
    Protected Function LoadRSSFeed() As SyndicationFeed
        LoadRSSFeed = Nothing
        Try
            Using reader As XmlReader = XmlReader.Create(tbURL.Text)
                LoadRSSFeed = SyndicationFeed.Load(reader)
            End Using
        Catch ex As WebException
        Catch ex As XmlException
        Catch ex As Exception
        End Try
    End Function
    
    Protected Sub DisplayRSSFeed()
        tbTest.Text = String.Empty
        Dim sf As SyndicationFeed = LoadRSSFeed()
        If Not sf Is Nothing Then
            For Each item As SyndicationItem In sf.Items
                tbTest.Text += "ITEM.TITLE: " & item.Title.Text & Environment.NewLine
                tbTest.Text += "ITEM.Date: " & Convert.ToDateTime(item.PublishDate.LocalDateTime).ToString("dd-MMM-yyyy HH:mm") & Environment.NewLine
                tbTest.Text += "ITEM.LASTUPDATEDTIME: " & item.LastUpdatedTime.ToString & Environment.NewLine
                For Each s As SyndicationPerson In item.Authors
                    tbTest.Text += "AUTHOR: " & s.Name & Environment.NewLine

                Next
                If Not item.Summary Is Nothing Then
                    tbTest.Text += "ITEM.Text: " & item.Summary.Text & Environment.NewLine
                End If
                If Not item.Content Is Nothing Then
                    tbTest.Text += "ITEM.CONTENT: " & item.Content.ToString & Environment.NewLine
                End If
                If Not item.Authors Is Nothing Then
                    For Each author As SyndicationPerson In item.Authors
                        tbTest.Text += "AUTHOR: " & author.Name & Environment.NewLine
                    Next
                End If
                If Not item.BaseUri Is Nothing Then
                    Dim suri As System.Uri = item.BaseUri
                    tbTest.Text += "BASEURI: " & suri.AbsoluteUri & Environment.NewLine
                End If
                If Not item.Categories Is Nothing Then
                    For Each sc As SyndicationCategory In item.Categories
                        tbTest.Text += "CATEGORY: " & sc.Name & Environment.NewLine
                    Next
                End If
                If Not item.Contributors Is Nothing Then
                    For Each author As SyndicationPerson In item.Contributors
                        tbTest.Text += "CONTRIBUTOR: " & author.Name & Environment.NewLine
                    Next
                End If
                If Not item.Links Is Nothing Then
                    For Each sl As SyndicationLink In item.Links
                        tbTest.Text += "LINK.TITLE: " & sl.Title & Environment.NewLine
                        tbTest.Text += "LINK.URI: " & sl.Uri.AbsoluteUri & Environment.NewLine
                    Next
                End If
                If Not item.SourceFeed Is Nothing Then
                    tbTest.Text += "SOURCEFEED.TITLE: " & item.SourceFeed.Title.Text & Environment.NewLine
                End If
            Next
        End If
    End Sub

    Protected Sub SaveLoginPageChanges()
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_SiteContent", oConn)
        oCmd.CommandType = CommandType.StoredProcedure

        Dim paramAction As New SqlParameter("@Action", SqlDbType.NVarChar, 50)
        paramAction.Value = "SET"
        oCmd.Parameters.Add(paramAction)

        Dim paramSiteKey As New SqlParameter("@SiteKey", SqlDbType.Int)
        paramSiteKey.Value = Session("SiteKey")
        oCmd.Parameters.Add(paramSiteKey)

        Dim paramContentType As New SqlParameter("@ContentType", SqlDbType.NVarChar, 50)
        Dim paramContent As New SqlParameter("@PageContentArea", SqlDbType.NVarChar, 3000)
        Dim paramContent2 As New SqlParameter("@PageContentArea2", SqlDbType.NVarChar, 3000)
        Dim paramContent3 As New SqlParameter("@PageContentArea3", SqlDbType.NVarChar, 3000)
        Select Case psLastFieldEdited
            Case "LPleft"
                paramContentType.Value = "LP1Content"
                paramContent.Value = RadEditor1.Content.TrimEnd
            Case "LPleft+"
                paramContentType.Value = "LP1Content+"
                paramContent.Value = RadEditor1.Content.TrimEnd
                paramContent2.Value = tbCSSAttributes.Text
            Case "LPtop"
                paramContentType.Value = "LPTopContent"
                paramContent.Value = RadEditor1.Content.TrimEnd
                paramContent2.Value = tbCSSAttributes.Text
            Case "LPbottom"
                paramContentType.Value = "LPBottomContent"
                paramContent.Value = RadEditor1.Content.TrimEnd
                paramContent2.Value = tbCSSAttributes.Text
            Case "LPright"
                paramContentType.Value = "LP4Content"
                paramContent.Value = RadEditor1.Content.TrimEnd
                paramContent2.Value = tbCSSAttributes.Text
            Case "LPlayout"
                paramContentType.Value = "LPAttr"
                paramContent.Value = tbLPLeftAttributes.Text
                paramContent2.Value = tbLPRightAttributes.Text
                Dim paramPane As New SqlParameter("@Pane", SqlDbType.Int)
                paramPane.Value = tbLPLeftAttributes.Text
            Case "SiteLogo"
                paramContentType.Value = "DefaultRunningHeaderImage"
                Dim paramDefaultRunningHeaderImage As New SqlParameter("@DefaultRunningHeaderImage", SqlDbType.VarChar, 100)
                paramDefaultRunningHeaderImage.Value = tbSiteLogoURL.Text
                oCmd.Parameters.Add(paramDefaultRunningHeaderImage)
        End Select
        oCmd.Parameters.Add(paramContentType)
        oCmd.Parameters.Add(paramContent)
        oCmd.Parameters.Add(paramContent2)
        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As SqlException
            WebMsgBox.Show("SaveLoginPageChanges: " & ex.Message)
        Finally
            oConn.Close()
        End Try
        psLastFieldEdited = ddlAreaSelector.SelectedValue
    End Sub

    Protected Sub lnkbtnRemoveImage_Click(sender As Object, e As System.EventArgs)
        Dim lnkbtn As LinkButton = sender
        Dim nID As Int32 = lnkbtn.CommandArgument
        Dim dt As DataTable = ExecuteQueryToDataTable("SELECT * FROM RotatorImages WHERE [id] = " & nID)
        If dt.Rows.Count > 0 Then
            Dim dr As DataRow = dt.Rows(0)
            My.Computer.FileSystem.DeleteFile(MapPath(UPLOADED_IMAGES_PATH) + dr("FileName"))
            Call ExecuteQueryToDataTable("DELETE FROM  RotatorImages WHERE [id] = " & nID)
            Call LoadImages()
        End If
    End Sub

#End Region

    Property pdsProducts() As DataSet
        Get
            Dim o As Object = ViewState("Products")
            If o Is Nothing Then
                Return Nothing
            End If
            Return CType(o, DataSet)
        End Get
        Set(ByVal Value As DataSet)
            ViewState("Products") = Value
        End Set
    End Property
  
    Property psLastFieldEdited() As String
        Get
            Dim o As Object = ViewState("NBE_LastFieldEdited")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("NBE_LastFieldEdited") = Value
        End Set
    End Property

</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.5/jquery.min.js"></script>
    <script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jqueryui/1.8/jquery-ui.min.js"></script>
    <link href="http://ajax.googleapis.com/ajax/libs/jqueryui/1.8/themes/base/jquery-ui.css"
        rel="stylesheet" type="text/css" />
    <title>Notice Board Editor</title>
    <link rel="Stylesheet" type="text/css" href="~/css/sprint_rotator.css" />
    <style type="text/css" media="screen">
        BODY
        {
            font-family: Verdana;
        }
        UNKNOWN
        {
            window: editorInstance.Description;
        }
        TABLE
        {
            font-size: 7pt;
            font-family: verdana;
        }
        TD.small
        {
            font-size: 8pt;
            font-family: verdana;
        }
        TD.subheading
        {
            font-size: 14pt;
            font-family: sans-serif;
        }
        TR.darkbackground
        {
            background-color: silver;
        }
        bold
        {
        }
    </style>
    <script type="text/javascript">
        function validationFailed(sender, eventArgs) {
            $telerik.$("#radAsyncUpload").html("Validation failed for '" + eventArgs.get_fileName() + "'.").fadeIn("slow");
        }
        function fileRemoved(sender, eventArgs) {
            $telerik.$("#radAsyncUpload").html('').fadeOut("slow");
        }        
    </script>
    <script type="text/javascript">
        $(document).ready(function () {
            $("#accordion").accordion({
                autoheight: false,
                active: false,
                collapsible: true,
                alwaysOpen: false
            });
        });
    </script>
</head>
<body class="sf">
    <form id="frmSiteEditor" runat="server">
    <main:Header ID="ctlHeader" runat="server"></main:Header>
    <table width="100%" cellpadding="0" cellspacing="0">
        <tr class="bar_siteeditor">
            <td style="width: 50%; white-space: nowrap">
            </td>
            <td style="width: 50%; white-space: nowrap" align="right">
            </td>
        </tr>
        <tr>
            <td style="width: 85%; white-space: nowrap">
                <asp:Button ID="btnLoginPageEditor" runat="server" Text="login page editor" OnClick="btnLoginPageEditor_Click"
                    CausesValidation="False" /><asp:Button ID="btnNoticeBoardEditor" runat="server" Text="noticeboard editor"
                        OnClick="btnDynamicControls_Click" CausesValidation="False" /><asp:Button ID="btnDataSourceEditor"
                            runat="server" Text="data source editor" OnClick="btnDataSourceEditor_Click"
                            CausesValidation="False" /><asp:Button ID="btnStyleSheetEditor" runat="server" Text="style sheet editor"
                                OnClick="btnStyleSheetEditor_Click" CausesValidation="False" />
                &nbsp;&nbsp;
                <asp:Button ID="btnUploadImages" runat="server" Text="upload images" OnClick="btnUploadImages_Click"
                    CausesValidation="False" /><asp:Button ID="btnShowImages" runat="server" Text="show images"
                        OnClick="btnShowImages_Click" CausesValidation="False" />
                &nbsp;&nbsp;
                <asp:Button ID="btnSaveChanges" runat="server" Text="save changes" OnClick="btnSaveChanges_Click" />
                <asp:Button ID="btnHelp" runat="server" Text="help" OnClick="btnHelp_Click" CausesValidation="False" />
            </td>
            <td align="right" style="width: 15%; white-space: nowrap">
                <%--<asp:LinkButton ID="lnkbtnHelp" runat="server" OnClientClick='window.open("help_SiteEditor.pdf", "NBHelp","top=10,left=10,width=700,height=450,status=no,toolbar=yes,address=no,menubar=yes,resizable=yes,scrollbars=yes");'>site editor help</asp:LinkButton>&nbsp;--%>
            </td>
        </tr>
    </table>
    &nbsp; &nbsp;&nbsp;
    <asp:Panel ID="pnlLoginPageEditor" runat="server" Visible="true" Width="100%">
        <table width="100%">
            <tr>
                <td style="width: 33%">
                    <asp:Label ID="Label19" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="Login Page Editor" />
                </td>
                <td style="width: 34%">
                </td>
                <td style="width: 33%" align="right">
                    <asp:LinkButton ID="lnkbtnRestoreDefaultSettings" runat="server" OnClick="lnkbtnRestoreDefaultSettings_Click"
                        OnClientClick='return confirm("Are you sure you remove all page layout and text customisation and return to the default settings?");'>restore default settings</asp:LinkButton>&nbsp;
                </td>
            </tr>
        </table>
        <table width="100%">
            <tr class="darkbackground">
                <td class="small">
                    &nbsp;
                    <asp:DropDownList ID="ddlAreaSelector" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        AutoPostBack="True" OnSelectedIndexChanged="ddlAreaSelector_SelectedIndexChanged">
                    </asp:DropDownList>
                    &nbsp; &nbsp;
                    <asp:CheckBox ID="cbShowAdvancedFeatures" runat="server" AutoPostBack="True" Font-Names="Verdana"
                        Font-Size="XX-Small" OnCheckedChanged="cbShowAdvancedFeatures_CheckedChanged"
                        Text="show advanced features" />
                    &nbsp;&nbsp;&nbsp;
                    <asp:Label ID="lblLegendCSSAttributes" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="CSS attributes:" Visible="False" />
                    &nbsp;<asp:TextBox ID="tbCSSAttributes" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Visible="False" Width="500px" />
                    &nbsp;&nbsp;&nbsp;
                </td>
            </tr>
            <tr id="trLoginPageAdvancedControls" visible="false" class="darkbackground" runat="server">
                <td class="small" style="height: 23px">
                    <asp:Label ID="Label2" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Login page CSS:"
                        Font-Bold="True" />
                    &nbsp;
                    <asp:Label ID="Label4" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="left attributes:" />
                    &nbsp;<asp:TextBox ID="tbLPLeftAttributes" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Width="200px" />
                    &nbsp;
                    <asp:Label ID="Label5" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="right attributes:" />
                    &nbsp;<asp:TextBox ID="tbLPRightAttributes" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Width="200px" />
                    &nbsp;&nbsp;
                    <asp:Label ID="lblLegendLoginBoxPosition" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="Login box position:" Font-Bold="True" />&nbsp;<asp:DropDownList ID="ddlLoginBoxPosition"
                            runat="server" Font-Names="Verdana" Font-Size="XX-Small">
                            <asp:ListItem Value="0">top</asp:ListItem>
                            <asp:ListItem Value="1">left</asp:ListItem>
                            <asp:ListItem Value="2">left centre</asp:ListItem>
                            <asp:ListItem Value="3">right centre</asp:ListItem>
                            <asp:ListItem Value="4">right</asp:ListItem>
                            <asp:ListItem Value="5">bottom</asp:ListItem>
                        </asp:DropDownList>
                </td>
            </tr>
            <tr id="trNoticeBoard1AdvancedControls" runat="server" class="darkbackground" visible="false">
                <td class="small">
                    <asp:Label ID="Label3" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Notice board CSS:"
                        Font-Bold="True" />&nbsp;
                    <asp:Label ID="Label12" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="left attributes:" />&nbsp;<asp:TextBox ID="tbNB1LeftAttributes" runat="server"
                            Font-Names="Verdana" Font-Size="XX-Small" Width="200px" />
                    &nbsp;&nbsp;<asp:Label ID="Label13" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="centre attributes:" />
                    &nbsp;<asp:TextBox ID="tbNB1CentreAttributes" runat="server" Font-Names="Verdana"
                        Font-Size="XX-Small" Width="200px" />
                    &nbsp;
                    <asp:Label ID="Label6" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="right attributes:" />
                    &nbsp;<asp:TextBox ID="tbNB1RightAttributes" runat="server" Font-Names="Verdana"
                        Font-Size="XX-Small" Width="200px" />
                </td>
            </tr>
            <tr id="trSiteLogoURL" runat="server" class="darkbackground" visible="false">
                <td class="small" style="height: 22px">
                    <asp:Label ID="Label7" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Site logo URL:"
                        Font-Bold="True" />
                    &nbsp;<asp:TextBox ID="tbSiteLogoURL" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Width="400px" />
                </td>
            </tr>
        </table>
        <telerik:RadEditor ID="RadEditor1" SkinID="DefaultSetOfTools" runat="server" Width="100%" />
        <%--        <telerik:RadEditor ID="RadEditor1" Skin="Vista" SkinID="DefaultSetOfTools" runat="server"/>--%>
    </asp:Panel>
    <asp:Panel ID="pnlDataSourceEditor" runat="server" Visible="false" Width="100%">
        <table width="95%">
            <tr>
                <td class="subheading" style="width: 33%">
                    <asp:Label ID="Label1" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Data Source Editor"
                        Font-Bold="True" />
                </td>
                <td style="width: 34%">
                </td>
                <td style="width: 33%">
                </td>
            </tr>
        </table>
        <table width="95%">
            <tr>
                <td>
                    <table>
                        <tr>
                            <td>
                                Select data source:
                            </td>
                            <td>
                                <asp:DropDownList ID="ddlDataSource" AutoPostBack="true" runat="server" Font-Names="Verdana"
                                    Font-Size="XX-Small" />
                            </td>
                            <td>
                                <asp:LinkButton ID="lnkDeleteDataSource" OnClientClick="javascript:return confirm('Are you sure you want to delete this data source ?');"  Text="" runat="server"></asp:LinkButton>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td>
                    <table id="tblAddDataSource" visible="false" runat="server">
                        <tr>
                            <td>
                                Enter data source name:
                            </td>
                            <td>
                                <asp:TextBox ID="tbDataSource" MaxLength="50" runat="server" />
                                &nbsp;<asp:RequiredFieldValidator ID="rfvDataSource" ControlToValidate="tbDataSource"
                                    ErrorMessage="## enter data source name" runat="server" Font-Bold="True" />
                            </td>
                        </tr>
                        <tr>
                            <td>
                            </td>
                            <td>
                                <asp:Button ID="btnSaveDataSource" Text="save" runat="server" />
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td>
                    <table width="100%" id="tblDataSourceEditor" visible="false" runat="server">
                        <tr>
                            <td id="tdNews">
                                <telerik:RadGrid ID="rgDataSourceEditor" runat="server" CellPadding="2" Font-Names="Verdana"
                                    AllowPaging="true" PageSize="10" AllowSorting="true" Font-Size="XX-Small" Width="100%"
                                    AutoGenerateColumns="False">
                                    <MasterTableView CommandItemDisplay="Top">
                                        <Columns>
                                            <telerik:GridEditCommandColumn ButtonType="ImageButton" UniqueName="EditCommandColumn"
                                                ItemStyle-Width="30px">
                                            </telerik:GridEditCommandColumn>
                                            <telerik:GridTemplateColumn>
                                                <ItemTemplate>
                                                    <asp:LinkButton ID="lnkMoveUp" CommandName="MoveUp" Text="up" runat="server"></asp:LinkButton>
                                                    <asp:LinkButton ID="lnkMoveDown" CommandName="MoveDown" Text="down" runat="server"></asp:LinkButton>
                                                </ItemTemplate>
                                            </telerik:GridTemplateColumn>
                                            <telerik:GridBoundColumn DataField="ID" HeaderText="ID" Visible="false" ReadOnly="true">
                                            </telerik:GridBoundColumn>
                                            <telerik:GridBoundColumn DataField="DataSourceTag" HeaderText="Data Source Tag" Visible="false">
                                            </telerik:GridBoundColumn>
                                            <telerik:GridBoundColumn DataField="ImageTag" HeaderText="Image Tag">
                                            </telerik:GridBoundColumn>
                                            <telerik:GridBoundColumn DataField="Title" HeaderText="Title">
                                            </telerik:GridBoundColumn>
                                            <telerik:GridBoundColumn DataField="Content" HeaderText="Content">
                                            </telerik:GridBoundColumn>
                                            <telerik:GridBoundColumn DataField="Url" HeaderText="Url">
                                            </telerik:GridBoundColumn>
                                            <telerik:GridTemplateColumn>
                                                <ItemTemplate>
                                                    <asp:LinkButton ID="lnkDelete" runat="server" OnClick="lnkDelete_Click" CommandArgument='<%# Bind("ID") %>'
                                                        Text="Delete" ToolTip="Delete" CommandName="delete" OnClientClick="return confirm('Are you sure you want to delete this record?');">                                
                                                    </asp:LinkButton>
                                                    <asp:HiddenField ID="hidPosition" Value='<%# Bind("Position") %>' runat="server" />
                                                    <asp:HiddenField ID="hidID" Value='<%# Bind("ID") %>' runat="server" />
                                                </ItemTemplate>
                                            </telerik:GridTemplateColumn>
                                        </Columns>
                                        <EditFormSettings EditFormType="Template" InsertCaption="Add New record">
                                            <FormTemplate>
                                                <table width="100%">
                                                    <tr>
                                                        <td>
                                                            <label>
                                                                Image Tag</label>
                                                        </td>
                                                        <td>
                                                            <asp:TextBox ID="txtImageTag" Text='<%# Bind("ImageTag") %>' MaxLength="50" runat="server"></asp:TextBox>
                                                        </td>
                                                    </tr>
                                                    <tr>
                                                        <td>
                                                            <label>
                                                                Title</label>
                                                        </td>
                                                        <td>
                                                            <telerik:RadEditor Width="500px" Height="300px" ID="reTitle" Content='<%# Bind("Title") %>'
                                                                ToolbarMode="Default" runat="server">
                                                                <Tools>
                                                                    <telerik:EditorToolGroup>
                                                                        <telerik:EditorTool Name="Bold" />
                                                                        <telerik:EditorTool Name="Italic" />
                                                                        <telerik:EditorTool Name="Underline" />
                                                                        <telerik:EditorSeparator />
                                                                        <telerik:EditorTool Name="ForeColor" />
                                                                        <telerik:EditorTool Name="BackColor" />
                                                                        <telerik:EditorSeparator />
                                                                        <telerik:EditorTool Name="FontName" />
                                                                        <telerik:EditorTool Name="RealFontSize" />
                                                                    </telerik:EditorToolGroup>
                                                                </Tools>
                                                            </telerik:RadEditor>
                                                        </td>
                                                    </tr>
                                                    <tr>
                                                        <td>
                                                            <label>
                                                                Content</label>
                                                        </td>
                                                        <td>
                                                            <telerik:RadEditor Width="500px" Height="300px" ID="reContent" Content='<%# Bind("Content") %>'
                                                                ToolbarMode="Default" runat="server">
                                                                <Tools>
                                                                    <telerik:EditorToolGroup>
                                                                        <telerik:EditorTool Name="Bold" />
                                                                        <telerik:EditorTool Name="Italic" />
                                                                        <telerik:EditorTool Name="Underline" />
                                                                        <telerik:EditorSeparator />
                                                                        <telerik:EditorTool Name="ForeColor" />
                                                                        <telerik:EditorTool Name="BackColor" />
                                                                        <telerik:EditorSeparator />
                                                                        <telerik:EditorTool Name="FontName" />
                                                                        <telerik:EditorTool Name="RealFontSize" />
                                                                    </telerik:EditorToolGroup>
                                                                </Tools>
                                                            </telerik:RadEditor>
                                                        </td>
                                                    </tr>
                                                    <tr>
                                                        <td>
                                                            <label>
                                                                Url</label>
                                                        </td>
                                                        <td>
                                                            <asp:TextBox ID="txtUrl" Text='<%# Bind("Url") %>' MaxLength="255" runat="server"></asp:TextBox>
                                                        </td>
                                                    </tr>
                                                    <tr>
                                                        <td>
                                                        </td>
                                                        <td>
                                                            <asp:LinkButton ID="lnkbtnUpdate" runat="server" CausesValidation="True" Text='<%# IIf( DataBinder.Eval(Container, "OwnerTableView.IsItemInserted"), "Insert", "Update") %>'
                                                                CommandName='<%# IIf( DataBinder.Eval(Container, "OwnerTableView.IsItemInserted"), "PerformInsert", "Update") %>'></asp:LinkButton>
                                                            <asp:LinkButton ID="lnkbtnCancel" runat="server" Text="Cancel" CausesValidation="False"
                                                                CommandName="Cancel"></asp:LinkButton>
                                                            <asp:HiddenField ID="hidID" Value='<%# Bind("ID") %>' runat="server" />
                                                        </td>
                                                    </tr>
                                                </table>
                                            </FormTemplate>
                                        </EditFormSettings>
                                    </MasterTableView>
                                </telerik:RadGrid>
                                <%-- <asp:DataGrid ID="dgNews" runat="server" Height="280px" CellSpacing="1" AutoGenerateColumns="False"
                                    OnItemCommand="Item_Button" BorderWidth="0px" BorderStyle="Dotted" BorderColor="Silver"
                                    Font-Names="Verdana" Font-Size="XX-Small" OnPageIndexChanged="PageIndexChanged"
                                    Width="100%">
                                    <AlternatingItemStyle BackColor="#E0E0E0"></AlternatingItemStyle>
                                    <Columns>
                                        <asp:TemplateColumn>
                                            <ItemTemplate>
                                                <asp:Button ID="lnkInsert" Text="Insert" CommandName="Insert" runat="server" />
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:EditCommandColumn ButtonType="PushButton" UpdateText="update" CancelText="cancel"
                                            EditText="edit">
                                            <HeaderStyle Width="39px"></HeaderStyle>
                                        </asp:EditCommandColumn>
                                        <asp:TemplateColumn>
                                            <ItemTemplate>
                                                <asp:LinkButton ID="lnkMoveUp" CommandName="MoveUp" Text="up" runat="server"></asp:LinkButton>
                                                <asp:LinkButton ID="lnkMoveDown" CommandName="MoveDown" Text="down" runat="server"></asp:LinkButton>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn DataField="ImageTag" HeaderText="Image Tag" HeaderStyle-HorizontalAlign="Center" />
                                        <asp:TemplateColumn HeaderText="Title" HeaderStyle-HorizontalAlign="Center">
                                            <EditItemTemplate>
                                                <asp:HiddenField ID="hidID" Value='<%# Bind("ID") %>' runat="server" />
                                                <telerik:RadEditor Width="200px" Height="300px" ID="reTitle" Content='<%# Bind("Title") %>'
                                                    ToolbarMode="Default" runat="server">
                                                    <Tools>
                                                        <telerik:EditorToolGroup>
                                                            <telerik:EditorTool Name="Bold" />
                                                            <telerik:EditorTool Name="Italic" />
                                                            <telerik:EditorTool Name="Underline" />
                                                            <telerik:EditorSeparator />
                                                            <telerik:EditorTool Name="ForeColor" />
                                                            <telerik:EditorTool Name="BackColor" />
                                                            <telerik:EditorSeparator />
                                                            <telerik:EditorTool Name="FontName" />
                                                            <telerik:EditorTool Name="RealFontSize" />
                                                        </telerik:EditorToolGroup>
                                                    </Tools>
                                                </telerik:RadEditor>
                                            </EditItemTemplate>
                                            <ItemTemplate>
                                                <asp:HiddenField ID="hidPosition" Value='<%# Bind("Position") %>' runat="server" />
                                                <asp:HiddenField ID="hidID" Value='<%# Bind("ID") %>' runat="server" />
                                                <asp:Label ID="lblTitle" Text='<%# Bind("Title") %>' runat="server"></asp:Label>
                                            </ItemTemplate>
                                            <HeaderStyle HorizontalAlign="Center" />
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="Content" HeaderStyle-HorizontalAlign="Center">
                                            <EditItemTemplate>
                                                <telerik:RadEditor Width="200px" Height="300px" ID="reContent" Content='<%# Bind("Content") %>'
                                                    ToolbarMode="Default" runat="server">
                                                    <Tools>
                                                        <telerik:EditorToolGroup>
                                                            <telerik:EditorTool Name="Bold" />
                                                            <telerik:EditorTool Name="Italic" />
                                                            <telerik:EditorTool Name="Underline" />
                                                            <telerik:EditorSeparator />
                                                            <telerik:EditorTool Name="ForeColor" />
                                                            <telerik:EditorTool Name="BackColor" />
                                                            <telerik:EditorSeparator />
                                                            <telerik:EditorTool Name="FontName" />
                                                            <telerik:EditorTool Name="RealFontSize" />
                                                        </telerik:EditorToolGroup>
                                                    </Tools>
                                                </telerik:RadEditor>
                                            </EditItemTemplate>
                                            <ItemTemplate>
                                                <asp:Label ID="lblContent" Text='<%# Bind("Content") %>' runat="server"></asp:Label>
                                            </ItemTemplate>
                                            <HeaderStyle HorizontalAlign="Center" />
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="URL" HeaderStyle-HorizontalAlign="Center">
                                            <ItemTemplate>
                                                <asp:Label ID="lblUrl" Text='<%# Bind("Url") %>' runat="server"></asp:Label>
                                            </ItemTemplate>
                                            <EditItemTemplate>
                                                <asp:TextBox ID="txtUrl" Text='<%# Bind("Url") %>' MaxLength="255" runat="server"></asp:TextBox>
                                            </EditItemTemplate>
                                            <HeaderStyle HorizontalAlign="Center" />
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn>
                                            <ItemTemplate>
                                                <asp:ImageButton ID="ImageButton1" runat="server" ImageUrl="./images/delete.gif"
                                                    CommandName="Delete"></asp:ImageButton>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                    <HeaderStyle Font-Bold="True" Font-Italic="False" Font-Overline="False" Font-Strikeout="False"
                                        Font-Underline="False" />
                                </asp:DataGrid>--%>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td width="100%">
                    <table width="100%" id="tblRssEditor" runat="server">
                        <tr>
                            <td>
                                <asp:Label ID="Label11" runat="server" Text="URL:" />
                            </td>
                            <td>
                                <asp:TextBox ID="tbURL" runat="server" Width="500px" Font-Names="Verdana" Font-Size="XX-Small" />
                                <asp:Label ID="Label10" runat="server" Text="# Entries:"></asp:Label>
                                <asp:TextBox ID="tbEntryCount" MaxLength="4" runat="server" Width="50px">6</asp:TextBox>
                                <asp:RegularExpressionValidator ID="revURL" runat="server" ValidationExpression="http(s)?://([\w-]+\.)+[\w-]+(/[\w- ./?%&=]*)?"
                                    ControlToValidate="tbURL" ErrorMessage="Invalid URL"></asp:RegularExpressionValidator>
                            </td>
                        </tr>
                        <tr>
                            <td>
                            </td>
                            <td>
                                <asp:TextBox runat="server" ID="tbTest" Rows="10" TextMode="MultiLine" Width="600px"
                                    Font-Names="Verdana" Font-Size="XX-Small" />
                            </td>
                        </tr>
                        <tr>
                            <td>
                            </td>
                            <td>
                                <asp:Button ID="btnGo" runat="server" Text="go" Width="70px" />
                            </td>
                        </tr>
                        <tr>
                            <td colspan="2">
                                <asp:GridView ID="gvRSSFeed" runat="server" AutoGenerateColumns="False" CellPadding="2"
                                    Font-Size="XX-Small">
                                    <Columns>
                                        <asp:BoundField DataField="Date" HeaderText="Date Published" ReadOnly="True" SortExpression="Date"
                                            DataFormatString="{0:d-MMM-yyyy hh:mm}" />
                                        <asp:BoundField DataField="Title" HeaderText="Title" ReadOnly="True" SortExpression="Title" />
                                        <asp:TemplateField HeaderText="Summary" SortExpression="Text">
                                            <ItemTemplate>
                                                <asp:Label ID="Label1" runat="server" Text='<%# Bind("Content") %>' />
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="URL" SortExpression="BaseURI">
                                            <ItemTemplate>
                                                <asp:HyperLink ID="HyperLink1" runat="server" Text='<%# Bind("BaseURI") %>' NavigateUrl='<%# Bind("BaseURI") %>'
                                                    Target="_blank" />
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:BoundField DataField="Categories" HeaderText="Categories" ReadOnly="True" SortExpression="Categories" />
                                    </Columns>
                                </asp:GridView>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </asp:Panel>
    <asp:Panel ID="pnlStyleSheetEditor" Visible="false" runat="server" Width="100%">
        <table style="width: 100%">
            <tr>
                <td style="width: 1%">
                </td>
                <td style="width: 20%">
                    <asp:Label ID="Label14" runat="server" Text="Style Sheet Editor" Font-Names="Verdana"
                        Font-Size="XX-Small" Font-Bold="True"></asp:Label>
                </td>
                <td style="width: 29%">
                </td>
                <td style="width: 20%">
                </td>
                <td style="width: 29%">
                </td>
                <td style="width: 1%">
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td>
                </td>
                <td colspan="3">
                    <asp:TextBox ID="tbStyleSheet" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Rows="15" TextMode="MultiLine" Width="100%"></asp:TextBox>
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td>
                </td>
                <td colspan="2">
                    <asp:Button ID="btnSaveStyleSheetChanges" runat="server" Text="save style sheet changes"
                        OnClick="btnSaveStyleSheetChanges_Click" />
                    <asp:Button ID="btnRestoreDefaultStyleSheet" runat="server" Text="restore default style sheet"
                        OnClick="btnRestoreDefaultStyleSheet_Click" OnClientClick='return confirm("WARNING: This will overwrite any changes you have made to the style sheet. Are you sure you want to restore the default style sheet?");' />
                </td>
                <td>
                </td>
                <td>
                </td>
            </tr>
        </table>
    </asp:Panel>
    <asp:Panel Width="100%" Visible="false" ID="pnlNoticeBoard" runat="server">
        <asp:Label ID="Label18" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
            Text="Panel Editor" Font-Bold="True" />
        <table width="100%">
            <tr>
                <td style="text-align: center" id="pnlNoticeBoardHeader">
                    Header panel:
                </td>
            </tr>
            <tr>
                <td>
                    <asp:TextBox ID="tbHeaderPanel" TextMode="MultiLine" Height="100px" Width="95%" runat="server"></asp:TextBox>
                </td>
            </tr>
        </table>
        <br />
        <table width="100%" border="0" cellspacing="1px" cellpadding="1px">
            <tr>
                <td width="33%">
                    Left panel:
                    <telerik:RadEditor Width="80%" ID="reLeft" runat="server">
                        <Tools>
                            <telerik:EditorToolGroup>
                                <telerik:EditorTool Name="Bold" />
                                <telerik:EditorTool Name="Italic" />
                                <telerik:EditorTool Name="Underline" />
                                <telerik:EditorSeparator />
                                <telerik:EditorTool Name="ForeColor" />
                                <telerik:EditorTool Name="BackColor" />
                                <telerik:EditorSeparator />
                                <telerik:EditorTool Name="FontName" />
                                <telerik:EditorTool Name="RealFontSize" />
                            </telerik:EditorToolGroup>
                        </Tools>
                    </telerik:RadEditor>
                </td>
                <td width="33%">
                    Centre panel:
                    <telerik:RadEditor Width="80%" ID="reCentre" runat="server">
                        <Tools>
                            <telerik:EditorToolGroup>
                                <telerik:EditorTool Name="Bold" />
                                <telerik:EditorTool Name="Italic" />
                                <telerik:EditorTool Name="Underline" />
                                <telerik:EditorSeparator />
                                <telerik:EditorTool Name="ForeColor" />
                                <telerik:EditorTool Name="BackColor" />
                                <telerik:EditorSeparator />
                                <telerik:EditorTool Name="FontName" />
                                <telerik:EditorTool Name="RealFontSize" />
                            </telerik:EditorToolGroup>
                        </Tools>
                    </telerik:RadEditor>
                </td>
                <td width="33%">
                    Right panel:
                    <telerik:RadEditor Width="80%" ID="reRight" runat="server">
                        <Tools>
                            <telerik:EditorToolGroup>
                                <telerik:EditorTool Name="Bold" />
                                <telerik:EditorTool Name="Italic" />
                                <telerik:EditorTool Name="Underline" />
                                <telerik:EditorSeparator />
                                <telerik:EditorTool Name="ForeColor" />
                                <telerik:EditorTool Name="BackColor" />
                                <telerik:EditorSeparator />
                                <telerik:EditorTool Name="FontName" />
                                <telerik:EditorTool Name="RealFontSize" />
                            </telerik:EditorToolGroup>
                        </Tools>
                    </telerik:RadEditor>
                </td>
            </tr>
        </table>
        <br />
        <fieldset style="font-family: Verdana; font-size: xx-small">
            <legend>Add Rotator</legend>
            <table width="95%" id="tblRotatorSettings" runat="server">
                <tr>
                    <td width="60%">
                        <asp:ValidationSummary ID="vsPagesPanel" ValidationGroup="vgPagesPanel" runat="server"
                            BorderStyle="Solid" BorderColor="Red" />
                    </td>
                    <td>
                        &nbsp;
                    </td>
                </tr>
            </table>
            <table width="95%">
                <tr>
                    <td width="60%">
                        <table width="95%" id="tblRotatorConfiguration" runat="server">
                            <tr>
                                <td width="23%">
                                    Rotator template
                                </td>
                                <td>
                                    <asp:CheckBox ID="chkImage" Text="Image" Checked="true" runat="server" />
                                    <asp:CheckBox ID="chkTitle" Text="Title" Checked="true" runat="server" />
                                    <asp:CheckBox ID="chkContent" Text="Content" Checked="true" runat="server" />
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    Rotator type
                                </td>
                                <td>
                                    <asp:DropDownList ID="ddlRotatorType" runat="server" Font-Names="Verdana" Font-Size="XX-Small">
                                        <asp:ListItem Text="Automatic Advance" Value="AutomaticAdvance" />
                                        <asp:ListItem Text="Buttons" Value="Buttons" />
                                        <asp:ListItem Text="Buttons Over" Value="ButtonsOver" />
                                        <asp:ListItem Text="Carousel" Value="Carousel" />
                                        <asp:ListItem Text="Carousel Buttons" Value="CarouselButtons" />
                                        <asp:ListItem Text="Cover Flow" Value="CoverFlow" />
                                        <asp:ListItem Text="Cover Flow Buttons" Value="CoverFlowButtons" />
                                        <asp:ListItem Text="Slide Show" Value="SlideShow" />
                                        <asp:ListItem Text="Slide Show Buttons" Value="SlideShowButtons" />
                                    </asp:DropDownList>
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    Animation type
                                </td>
                                <td>
                                    <asp:DropDownList ID="ddlAnimationType" runat="server" Font-Names="Verdana" Font-Size="XX-Small">
                                        <asp:ListItem Text="None" Value="None" />
                                        <asp:ListItem Text="Pulse" Value="Pulse" />
                                        <asp:ListItem Text="Fade" Value="Fade" />
                                        <asp:ListItem Text="CrossFade" Value="CrossFade" />
                                    </asp:DropDownList>
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    Rotator width
                                </td>
                                <td>
                                    <asp:TextBox ID="tbRotatorWidth" ValidationGroup="vgPagesPanel" MaxLength="5" runat="server"
                                        Font-Names="Verdana" Font-Size="XX-Small" />
                                    <asp:RegularExpressionValidator ID="revRotatorWidth" Display="None" ValidationGroup="vgPagesPanel"
                                        ErrorMessage="Enter a valid width" ControlToValidate="tbRotatorWidth" ValidationExpression="^\d+$"
                                        runat="server" />
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    Rotator height
                                </td>
                                <td>
                                    <asp:TextBox ID="tbRotatorHeight" ValidationGroup="vgPagesPanel" MaxLength="5" runat="server"
                                        Font-Names="Verdana" Font-Size="XX-Small" />
                                    <asp:RegularExpressionValidator ID="revRotatorHeight" Display="None" ValidationGroup="vgPagesPanel"
                                        ErrorMessage="Enter a valid height" ControlToValidate="tbRotatorHeight" ValidationExpression="^\d+$"
                                        runat="server" />
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    Scroll direction
                                </td>
                                <td>
                                    <asp:DropDownList ID="ddlScrollDirection" runat="server" Font-Names="Verdana" Font-Size="XX-Small">
                                        <asp:ListItem Text="Up" Value="Up"></asp:ListItem>
                                        <asp:ListItem Text="Down" Value="Down"></asp:ListItem>
                                        <asp:ListItem Text="Left" Value="Left"></asp:ListItem>
                                        <asp:ListItem Text="Right" Value="Right"></asp:ListItem>
                                    </asp:DropDownList>
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    Scroll duration
                                </td>
                                <td>
                                    <asp:TextBox ID="tbScrollDuration" ValidationGroup="vgPagesPanel" MaxLength="5" runat="server"
                                        Font-Names="Verdana" Font-Size="XX-Small" />
                                    <asp:RegularExpressionValidator ID="revScrollDuration" Display="None" ValidationGroup="vgPagesPanel"
                                        ErrorMessage="Enter a valid scroll duration" ControlToValidate="tbScrollDuration"
                                        ValidationExpression="^\d+$" runat="server" />
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    Select data source
                                </td>
                                <td>
                                    <asp:DropDownList ID="ddlDataSourceTag" runat="server" Font-Names="Verdana" Font-Size="XX-Small" />
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    Style sheet (css):
                                </td>
                                <td>
                                    <asp:TextBox ID="tbCss" MaxLength="50" runat="server" Font-Names="Verdana" Font-Size="XX-Small" />
                                </td>
                            </tr>
                            <tr>
                                <td>
                                </td>
                                <td>
                                    <asp:Button ID="btnAddRotatorLeft" ValidationGroup="vgPagesPanel" Text="add to left panel"
                                        OnClick="btnAddRotator_Click" runat="server" />
                                    <asp:Button ID="btnAddRotatorCentre" ValidationGroup="vgPagesPanel" Text="add to centre panel"
                                        OnClick="btnAddRotator_Click" runat="server" />
                                    <asp:Button ID="btnAddRotatorRight" ValidationGroup="vgPagesPanel" Text="add to right panel"
                                        OnClick="btnAddRotator_Click" runat="server" />
                                    <asp:Button ID="btnAddRotatorHeader" ValidationGroup="vgPagesPanel" Text="add to header panel"
                                        OnClick="btnAddRotator_Click" runat="server" />
                                </td>
                            </tr>
                        </table>
                    </td>
                    <td style="vertical-align: top">
                        <table width="100%">
                            <tr>
                                <td style="vertical-align: top">
                                    <telerik:RadRotator ID="rrPreview" runat="server" />
                                </td>
                                <td style="vertical-align: top; text-align: right">
                                    <asp:Button ID="btnPreview" Text="Display Preview" runat="server" OnClick="btnPreview_Click" />
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
        </fieldset>
        <br />
    </asp:Panel>
    <asp:Panel ID="pnlUploadImage" Visible="false" Width="100%" runat="server">
        <asp:Label ID="Label15" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
            Text="Upload Image" Font-Bold="True" />
        <table width="100%">
            <tr>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td id="tdImgTag">
                    Enter image tag:&nbsp;
                    <asp:TextBox ID="tbImageTag" MaxLength="50" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Width="150px" />
                    &nbsp;<asp:RequiredFieldValidator ID="rfvImageTag" Display="None" ValidationGroup="vgUploadImage"
                        ControlToValidate="tbImageTag" ErrorMessage="Enter image tag" runat="server"
                        Font-Names="Verdana" Font-Size="XX-Small" />
                    <asp:ValidationSummary ID="vs" runat="server" BorderColor="Red" BorderStyle="Solid"
                        ValidationGroup="vgUploadImage" Width="275" Font-Names="Verdana" Font-Size="XX-Small" />
                </td>
            </tr>
            <tr>
                <td>
                    <telerik:RadAsyncUpload ID="radAsyncUpload" runat="server" OnFileUploaded="radAsyncUpload_FileUploaded"
                        MaxFileSize="524288" OnClientValidationFailed="validationFailed" AllowedFileExtensions="jpg,jpeg"
                        AutoAddFileInputs="false" OnClientFileUploadRemoved="fileRemoved" Font-Names="Verdana"
                        Font-Size="XX-Small" />
                    <asp:Button ID="btnUploadImageFile" Text="Upload Image" runat="server" ValidationGroup="vgUploadImage"
                        CausesValidation="true" /><br />
                </td>
            </tr>
            <tr>
                <td>
                    <asp:Label ID="Label9" Text="File Size limit 500KB, (jpg, jpeg)" runat="server" Style="line-height: 30px"
                        Font-Names="Verdana" Font-Size="XX-Small" />
                </td>
            </tr>
            <tr>
                <td>
                    <asp:ImageButton ID="imgUploadedImage" OnClick="imgUploadedImage_Click" runat="server" />
                </td>
            </tr>
            <tr>
                <td>
                    <asp:Label ID="lblUploadedImgMessage" runat="server" Font-Names="Verdana" Font-Size="XX-Small" />
                </td>
            </tr>
        </table>
    </asp:Panel>
    <asp:Panel ID="pnlShowImages" runat="server" Visible="false">
        <asp:Label ID="Label16" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
            Text="Show Images" Font-Bold="True" />
        <asp:DataList ID="dlCustomerImages" RepeatColumns="3" Width="800px" HorizontalAlign="Center"
            RepeatDirection="Horizontal" BorderStyle="Solid" BorderWidth="2px" runat="server">
            <ItemTemplate>
                <table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                        <td style="text-align: center">
                            <asp:ImageButton ID="imgCustomer" OnClick="imgCustomer_Click" ImageUrl='<%# Bind("FileName","~/Images/UploadedImages/{0}") %>'
                                runat="server" />
                        </td>
                    </tr>
                    <tr>
                        <td>
                        </td>
                    </tr>
                    <tr>
                        <td style="text-align: center">
                            <asp:Label ID="lblTagName" Font-Bold="true" Text='<%# Bind("ImageTag") %>' runat="server" />
                        </td>
                    </tr>
                </table>
                <asp:CheckBox ID="chkHeaderImage" AutoPostBack="true" Text="Set as header image"
                    OnCheckedChanged="chkHeaderImage_CheckedChanged" runat="server" />
                <br />
                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                <asp:LinkButton ID="lnkbtnRemoveImage" runat="server" OnClick="lnkbtnRemoveImage_Click"
                    CommandArgument='<%# Bind("ID") %>'>remove&nbsp;image</asp:LinkButton>
            </ItemTemplate>
            <FooterTemplate>
                <asp:Label ID="lblEmpty" ForeColor="Red" Text="No images found" Visible='<%# dlCustomerImages.Items.Count=0 %>'
                    runat="server" />
            </FooterTemplate>
            <ItemStyle Width="100px" HorizontalAlign="Center" Height="100px" BorderWidth="1px"
                BorderStyle="Solid" BorderColor="Black" />
        </asp:DataList>
    </asp:Panel>
    <asp:Panel ID="pnlImageEditor" Visible="false" runat="server">
        <asp:Label ID="Label17" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
            Text="Image Editor" Font-Bold="True" />
        <br />
        <telerik:RadImageEditor ID="radImageEditor" runat="server" Height="430px" Style="top: 0px;
            left: 0px" Width="720px">
            <Tools>
                <telerik:ImageEditorToolGroup>
                    <telerik:ImageEditorToolStrip CommandName="Undo" />
                    <telerik:ImageEditorToolStrip CommandName="Redo" />
                    <telerik:ImageEditorTool CommandName="Reset" />
                    <telerik:ImageEditorToolSeparator />
                    <telerik:ImageEditorTool CommandName="Crop" ToolTip="Crop" />
                    <telerik:ImageEditorTool CommandName="Resize" ToolTip="Resize" />
                    <telerik:ImageEditorTool CommandName="Zoom" ToolTip="Zoom" />
                    <telerik:ImageEditorTool CommandName="ZoomIn" ToolTip="Zoom in" />
                    <telerik:ImageEditorTool CommandName="ZoomOut" ToolTip="Zoom out" />
                    <telerik:ImageEditorTool CommandName="Rotate" IsToggleButton="true" ToolTip="Rotate image" />
                    <telerik:ImageEditorTool CommandName="RotateRight" ToolTip="Rotate right 90 degrees" />
                    <telerik:ImageEditorTool CommandName="RotateLeft" ToolTip="Rotate left 90 degrees" />
                    <telerik:ImageEditorTool CommandName="Flip" IsToggleButton="true" ToolTip="Flip Image" />
                    <telerik:ImageEditorTool CommandName="FlipVertical" ToolTip="Flip image vertically" />
                    <telerik:ImageEditorTool CommandName="FlipHorizontal" ToolTip="Flip image horizontally" />
                    <telerik:ImageEditorTool CommandName="Opacity" IsToggleButton="true" />
                    <telerik:ImageEditorTool CommandName="AddText" />
                </telerik:ImageEditorToolGroup>
            </Tools>
        </telerik:RadImageEditor>
    </asp:Panel>
    <div runat="server" id="accordion" style="width:90%; text-align:left" visible="false">
            <h3><a href="#">Introduction to the Site Editorr</a></h3>
            <div style="text-align: left">
                The Site Editor lets you customise the Login Page, the Notice Board and the page 
                header area.<br />
                <br />
                You can 
                add text and images to the Login page or the Notice Board page, add an animated message 
                rotator, link to an RSS feed and more.<br />
                <br />
                The Site Editor 
                tab provides several different editors to achieve the results you 
                want. Select the editor you require using the buttons at the top of the page.
                The help topics below describe how to use these editors.<br />
                <br />
                Whenever you make a change to your content, you must save it by clicking the <b>
                save changes</b> 
                button before continuing.<br />
                <br />
                The Site Editor makes extensive use of a sophisticated web-oriented HTML editor window, which allows you to enter text and images, set text colour, 
                font and size, layout options and many other features typically found on web 
                pages. Full end-user documentation for the HTML editor can be found <a href="http://www.telerik.com/documents/radeditorajaxendusermanual.pdf">here.</a></div>
            <h3><a href="#">Login Page Editor</a></h3>
            <div style="text-align: left">
                The Login Page editor allows you to add text and images to the Login Page, 
                using the HTML editor.<br />
                <br />
                The page is divided into four areas - header, left hand column, right hand 
                column, footer.&nbsp;
                By default you add text to the left hand column of the page.<br />
                <br />
                To add to another area, click the <b>show advanced features</b> check box, then 
                select the area to which you want to add text and images.<br />
                <br />
                You can add images you have uploaded using the <b>upload images</b> screen, or 
                you can link to any image on the Transworld server or on the public internet.<br />
                <br />
                Currently the </div>
            <h3><a href="#">Notice Board Editor</a></h3>
            <div style="text-align: left">
            <p>The Notice Board page is divided into four panels - the Header panel and three 
                columns: Left panel, Centre panel, Right panel.</p>
            </div>
            <div style="text-align: left">
                <p>&nbsp;</p>
            </div>
            <div style="text-align: left">
                <p>The rotator only scrolls if there are more entries than can be displayed.</p>
            </div>
            <div style="text-align: left">
                <p>Use the HTML editors to add text to the columns.</p>
            </div>
            <h3><a href="#">Data Source Editor</a></h3>
            <div style="text-align: left">
            <p>Creating content that is animated, or where text comes from an external 
                web page via an RSS feed, is a two-stage process. First you specify a data source and give it a 
                name using the Data Source Editor, then you specify that named data source you 
                want in the Notice Board Editor.</p>
            </div>
            <div style="text-align: left">
                <p>&nbsp;</p>
            </div>
            <div style="text-align: left">
                <p>&nbsp;</p>
            </div>
            <h3><a href="#">Style Sheet Editor</a></h3>
            <div style="text-align: left">
            <p>some more content here</p>
            </div>
            <h3><a href="#">Upload Image</a></h3>
            <div style="text-align: left">
            <p>some more content here</p>
            </div>
            <h3><a href="#">Show Images</a></h3>
            <div style="text-align: left">
            <p>some more content here</p>
            </div>
            <h3><a href="#">Save Changes</a></h3>
            <div style="text-align: left">
            <p>some more content here</p>
            </div>
            <h3><a href="#">Help</a></h3>
            <div style="text-align: left">
            <p>some more content here</p>
            </div>

    </div>
    </form>
</body>
</html>