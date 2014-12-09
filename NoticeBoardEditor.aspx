<%@ Page Language="VB" ValidateRequest="false" Theme="AIMSDefault" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ Register TagPrefix="ComponentArt" Namespace="ComponentArt.Web.UI" Assembly="ComponentArt.Web.UI" %>
<%@ Register TagPrefix="FCKeditorV2" Namespace="FredCK.FCKeditorV2" Assembly="FredCK.FCKeditorV2" %>
<%@ import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.XML" %>
<%@ Import Namespace="System.IO" %>

<script runat="server">

    Const TABLENAME_HEADER_ROTATOR As String = "HeaderRotator"
    Const TABLENAME_NOTICEBOARD_ROTATOR As String = "NoticeBoard1Rotator"
    Const TABLENAME_LHSBODY As String = "LHSBody"
    Const TABLENAME_PAGETITLE As String = "PageTitle"
    
    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    Dim gsXMLRotatorConfigFilePath As String
    Dim gsXMLNewsContentFilePath As String
    Dim gds As New DataSet
    Dim gdt As DataTable
    Dim gcol As DataColumn
    
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)
        If Not IsNumeric(Session("SiteKey")) Then
            Server.Transfer("session_expired.aspx")
            Exit Sub
        End If
        
        gsXMLRotatorConfigFilePath = ".\rotator\news_config" & Session("SiteKey") & ".xml"
        gsXMLNewsContentFilePath = ".\rotator\news" & Session("SiteKey") & ".xml"

        If Not IsPostBack Then
            Call InitAreaSelector()
            psLastFieldEdited = ddlAreaSelector.SelectedValue

            If Not File.Exists(Server.MapPath(gsXMLRotatorConfigFilePath)) Then
                CreateNewConfigFile()
            End If
            Call WriteRotatorConfigsFromDatabase()
            Call InitFromXML()
            Call SaveCurrentTargetsInViewState()
            If Not File.Exists(MapPath(gsXMLNewsContentFilePath)) Then
                Call CreateDataset()
            Else
                Call LoadDataset()
            End If
            dgNews.DataSource = gds
            dgNews.DataBind()
        End If
    
        If Not rblRotatorTarget.SelectedItem.Value = ViewState("RotatorTarget") Then
            Dim val, vs As String
            val = rblRotatorTarget.SelectedItem.Value
            vs = ViewState("RotatorTarget")
            rblRotatorTarget.SelectedItem.Value = vs
            Call WriteConfigFile()
            rblRotatorTarget.SelectedItem.Value = val
            Call InitFromXML()
        End If
        Call ApplySettings()
        Call SaveCurrentTargetsInViewState()
        
        Call SetTitle()
        Call SetStyleSheet()
    End Sub
    
    Protected Sub SetTitle()
        Dim sTitle As String = Session("SiteTitle")
        If sTitle <> String.Empty Then
            sTitle += " - "
        End If
        Page.Header.Title = sTitle & "Notice Board Editor"
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
        ddlAreaSelector.Items.Add(New ListItem("notice board - main text area", "NB1left"))
        ddlAreaSelector.Items.Add(New ListItem("login page - main text area", "LPleft"))
        If cbShowAdvancedFeatures.Checked Then
            ddlAreaSelector.Items.Add(New ListItem("", ""))
            ddlAreaSelector.Items.Add(New ListItem("notice board - top text area", "NB1top"))
            ddlAreaSelector.Items.Add(New ListItem("notice board - bottom text area", "NB1bottom"))
            ddlAreaSelector.Items.Add(New ListItem("", ""))
            ddlAreaSelector.Items.Add(New ListItem("login page - top text area", "LPtop"))
            ddlAreaSelector.Items.Add(New ListItem("login page - bottom text area", "LPbottom"))
            ddlAreaSelector.Items.Add(New ListItem("login page - right text area", "LPright"))
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
            Case "LPleft"
                FCKeditor1.Value = (dr("LP1Content") & String.Empty).ToString.TrimEnd
                tbCSSAttributes.Text = dr("LP1Attr") & String.Empty
            Case "LPtop"
                FCKeditor1.Value = (dr("LPTopContent") & String.Empty).ToString.TrimEnd
                tbCSSAttributes.Text = dr("LPTopAttr") & String.Empty
            Case "LPbottom"
                FCKeditor1.Value = (dr("LPBottomContent") & String.Empty).ToString.TrimEnd
                tbCSSAttributes.Text = dr("LPBottomAttr") & String.Empty
            Case "LPright"
                FCKeditor1.Value = (dr("LP4Content") & String.Empty).ToString.TrimEnd
                tbCSSAttributes.Text = dr("LP4Attr") & String.Empty
            Case "NB1left"
                FCKeditor1.Value = (dr("NB1_BodyContent") & String.Empty).ToString.TrimEnd
                tbCSSAttributes.Text = dr("NB1_BodyAttr") & String.Empty
            Case "NB1top"
                FCKeditor1.Value = (dr("NB1_TopContent") & String.Empty).ToString.TrimEnd
                tbCSSAttributes.Text = dr("NB1_TopAttr") & String.Empty
            Case "NB1bottom"
                FCKeditor1.Value = (dr("NB1_BottomContent") & String.Empty).ToString.TrimEnd
                tbCSSAttributes.Text = dr("NB1_BottomAttr") & String.Empty
        End Select
    End Sub
    
    Protected Sub SaveNewsChanges()
        
    End Sub
    
    Protected Sub SaveRotatorChanges()
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
        Select Case rblRotatorTarget.SelectedValue
            Case "NoticeBoard1Rotator"
                paramContentType.Value = "NB1Rtr"
            Case "HeaderRotator"
                paramContentType.Value = "HdrRtr"
        End Select
        oCmd.Parameters.Add(paramContentType)

        Dim paramVisible As New SqlParameter("@RtrVisible", SqlDbType.Bit)
        paramVisible.Value = pbVisible
        oCmd.Parameters.Add(paramVisible)

        Dim paramContinuousLoop As New SqlParameter("@RtrContinuousLoop", SqlDbType.Bit)
        paramContinuousLoop.Value = pbContinuousLoop
        oCmd.Parameters.Add(paramContinuousLoop)

        Dim paramPauseOnMouseOver As New SqlParameter("@RtrPauseOnMouseOver", SqlDbType.Bit)
        paramPauseOnMouseOver.Value = pbPauseOnMouseOver
        oCmd.Parameters.Add(paramPauseOnMouseOver)

        Dim paramScrollDirection As New SqlParameter("@RtrScrollDirection", SqlDbType.VarChar, 10)
        paramScrollDirection.Value = ViewState("ED_ScrollDirection")
        oCmd.Parameters.Add(paramScrollDirection)

        Dim paramSlidePause As New SqlParameter("@RtrSlidePause", SqlDbType.SmallInt)
        paramSlidePause.Value = pnSlidePause
        oCmd.Parameters.Add(paramSlidePause)
        
        Dim paramScrollInterval As New SqlParameter("@RtrScrollInterval", SqlDbType.SmallInt)
        paramScrollInterval.Value = pnScrollInterval
        oCmd.Parameters.Add(paramScrollInterval)
        
        Dim paramRotationType As New SqlParameter("@RtrRotationType", SqlDbType.VarChar, 15)
        paramRotationType.Value = ViewState("ED_RotationType")
        oCmd.Parameters.Add(paramRotationType)

        Dim paramSmoothScrollSpeed As New SqlParameter("@RtrSmoothScrollSpeed", SqlDbType.VarChar, 10)
        paramSmoothScrollSpeed.Value = ViewState("ED_SmoothScrollSpeed")
        oCmd.Parameters.Add(paramSmoothScrollSpeed)

        Dim paramShowEffect As New SqlParameter("@RtrShowEffect", SqlDbType.VarChar, 10)
        paramShowEffect.Value = ViewState("ED_ShowEffect")
        oCmd.Parameters.Add(paramShowEffect)

        Dim paramShowEffectDuration As New SqlParameter("@RtrShowEffectDuration", SqlDbType.SmallInt)
        paramShowEffectDuration.Value = pnShowEffectDuration
        oCmd.Parameters.Add(paramShowEffectDuration)
        
        Dim paramHideEffect As New SqlParameter("@RtrHideEffect", SqlDbType.VarChar, 10)
        paramHideEffect.Value = ViewState("ED_HideEffect")
        oCmd.Parameters.Add(paramHideEffect)

        Dim paramHideEffectDuration As New SqlParameter("@RtrHideEffectDuration", SqlDbType.SmallInt)
        paramHideEffectDuration.Value = pnHideEffectDuration
        oCmd.Parameters.Add(paramHideEffectDuration)
        
        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As SqlException
            WebMsgBox.Show("SaveRotatorChanges: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub SaveTextChanges()
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
        Select Case psLastFieldEdited
            Case "LPleft"
                paramContentType.Value = "LP1Content"
            Case "LPtop"
                paramContentType.Value = "LPTopContent"
            Case "LPbottom"
                paramContentType.Value = "LPBottomContent"
            Case "LPright"
                paramContentType.Value = "LP4Content"
            Case "NB1left"
                paramContentType.Value = "NB1_BodyContent"
            Case "NB1top"
                paramContentType.Value = "NB1_TopContent"
            Case "NB1bottom"
                paramContentType.Value = "NB1_BottomContent"
        End Select
        oCmd.Parameters.Add(paramContentType)

        Dim paramSiteTitle As New SqlParameter("@PageContentArea", SqlDbType.NVarChar, 3000)
        paramSiteTitle.Value = FCKeditor1.Value.TrimEnd
        oCmd.Parameters.Add(paramSiteTitle)

        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As SqlException
            WebMsgBox.Show("SaveTextChanges: " & ex.Message)
        Finally
            oConn.Close()
        End Try
        psLastFieldEdited = ddlAreaSelector.SelectedValue
    End Sub

    Protected Sub ApplySettings() ' equivalent of property sets for each property, done each Page Load
        setVisible(rblVisible.SelectedItem.Value.ToString)
        setContinuousLoop(rblContinuousLoop.SelectedItem.Value.ToString)
        setPauseOnMouseOver(rblPauseOnMouseOver.SelectedItem.Value.ToString)
        setScrollDirection(rblScrollDirection.SelectedItem.Value.ToString)
        setRotationType(rblDisplayType.SelectedItem.Value.ToString)
        setSlidePause(tbSlidePause.Text)
        setScrollInterval(tbScrollInterval.Text)
        setSmoothScrollSpeed(rblSmoothScrollSpeed.SelectedItem.Value.ToString)
        setShowEffect(rblShowEffect.SelectedItem.Value.ToString)
        setShowEffectDuration(tbShowEffectDuration.Text)
        setHideEffect(rblHideEffect.SelectedItem.Value.ToString)
        setHideEffectDuration(tbHideEffectDuration.Text)
    
        WriteConfigFile()
    End Sub 'ApplySettings
    
    Protected Sub SaveCurrentTargetsInViewState()
        ViewState("RotatorTarget") = rblRotatorTarget.SelectedItem.Value
    End Sub
    
    Protected Sub WriteRotatorConfigsFromDatabase()
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
            WebMsgBox.Show("GetPageContent: " & ex.Message)
        Finally
            oConn.Close()
        End Try

        Dim drDatabase As DataRow = oDataTable.Rows(0)
        Dim sXMLRotatorConfigFilePath As String = Server.MapPath(gsXMLRotatorConfigFilePath)
        Dim fs As New System.IO.FileStream(sXMLRotatorConfigFilePath, System.IO.FileMode.Open)
    
        Dim ds As DataSet
        Dim oXMLDataTable As DataTable
        Dim drXML As DataRow
    
        ds = New DataSet
        ds.ReadXml(fs)
        fs.Close()
    
        oXMLDataTable = ds.Tables("NoticeBoard1Rotator")            ' "NoticeBoard1Rotator" | "HeaderRotator"
        drXML = oXMLDataTable.Rows(0)
    
        drXML("Visible") = drDatabase("NB1RtrVisible")
        drXML("ContinuousLoop") = drDatabase("NB1RtrContinuousLoop")
        drXML("PauseOnMouseOver") = drDatabase("NB1RtrPauseOnMouseOver")
        drXML("ScrollDirection") = drDatabase("NB1RtrScrollDirection")
        drXML("SlidePause") = drDatabase("NB1RtrSlidePause")
        drXML("ScrollInterval") = drDatabase("NB1RtrScrollInterval")
        drXML("RotationType") = drDatabase("NB1RtrRotationType")
        drXML("SmoothScrollSpeed") = drDatabase("NB1RtrSmoothScrollSpeed")
        drXML("ShowEffect") = drDatabase("NB1RtrShowEffect")
        drXML("ShowEffectDuration") = drDatabase("NB1RtrShowEffectDuration")
        drXML("HideEffect") = drDatabase("NB1RtrHideEffect")
        drXML("HideEffectDuration") = drDatabase("NB1RtrHideEffectDuration")
    
        oXMLDataTable = ds.Tables("HeaderRotator")            ' "NoticeBoard1Rotator" | "HeaderRotator"
        drXML = oXMLDataTable.Rows(0)
    
        drXML("Visible") = drDatabase("HdrRtrVisible")
        drXML("ContinuousLoop") = drDatabase("HdrRtrContinuousLoop")
        drXML("PauseOnMouseOver") = drDatabase("HdrRtrPauseOnMouseOver")
        drXML("ScrollDirection") = drDatabase("HdrRtrScrollDirection")
        drXML("SlidePause") = drDatabase("HdrRtrSlidePause")
        drXML("ScrollInterval") = drDatabase("HdrRtrScrollInterval")
        drXML("RotationType") = drDatabase("HdrRtrRotationType")
        drXML("SmoothScrollSpeed") = drDatabase("HdrRtrSmoothScrollSpeed")
        drXML("ShowEffect") = drDatabase("HdrRtrShowEffect")
        drXML("ShowEffectDuration") = drDatabase("HdrRtrShowEffectDuration")
        drXML("HideEffect") = drDatabase("HdrRtrHideEffect")
        drXML("HideEffectDuration") = drDatabase("HdrRtrHideEffectDuration")

        fs = New System.IO.FileStream(sXMLRotatorConfigFilePath, IO.FileMode.Create)
        ds.WriteXml(fs)
        ds = Nothing
        fs.Close()
    End Sub
    
    Protected Sub WriteConfigFile()
        Dim sXMLRotatorConfigFilePath As String = Server.MapPath(gsXMLRotatorConfigFilePath)
        Dim fs As New System.IO.FileStream(sXMLRotatorConfigFilePath, System.IO.FileMode.Open)
    
        Dim ds As DataSet
        Dim oDataTable As DataTable
        Dim dr As DataRow
    
        ds = New DataSet
        ds.ReadXml(fs)
        fs.Close()
    
        oDataTable = ds.Tables(rblRotatorTarget.SelectedItem.Value) '           "NoticeBoard1Rotator" | "HeaderRotator"
        dr = oDataTable.Rows(0)
    
        dr("Visible") = rblVisible.SelectedItem.Value.ToString
        dr("ContinuousLoop") = rblContinuousLoop.SelectedItem.Value.ToString
        dr("PauseOnMouseOver") = rblPauseOnMouseOver.SelectedItem.Value.ToString
        dr("ScrollDirection") = rblScrollDirection.SelectedItem.Value.ToString
        dr("SlidePause") = tbSlidePause.Text
        dr("ScrollInterval") = tbScrollInterval.Text
        dr("RotationType") = rblDisplayType.SelectedItem.Value.ToString
        dr("SmoothScrollSpeed") = rblSmoothScrollSpeed.SelectedItem.Value.ToString
        dr("ShowEffect") = rblShowEffect.SelectedItem.Value.ToString
        dr("ShowEffectDuration") = tbShowEffectDuration.Text
        dr("HideEffect") = rblHideEffect.SelectedItem.Value.ToString
        dr("HideEffectDuration") = tbHideEffectDuration.Text
    
        fs = New System.IO.FileStream(sXMLRotatorConfigFilePath, IO.FileMode.Create)
        ds.WriteXml(fs)
    
        ds = Nothing
        fs.Close()
    End Sub
    
    Protected Sub InitFromXML()
        Dim sFilename As String = gsXMLRotatorConfigFilePath
        sFilename = Server.MapPath(sFilename)
        Dim fs As New System.IO.FileStream(sFilename, System.IO.FileMode.Open)
    
        Dim ds As DataSet
        Dim t As DataTable
        Dim dr As DataRow
        Dim sValue As String
        Dim i As Integer
    
        ds = New DataSet
        ds.ReadXml(fs)
        fs.Close()
    
        t = ds.Tables(rblRotatorTarget.SelectedItem.Value)
    
        dr = t.Rows(0)
    
        sValue = dr("Visible")
        setVisible(sValue)
        For i = 0 To rblVisible.Items.Count - 1
            If rblVisible.Items(i).Value = sValue Then
                rblVisible.SelectedIndex = i
            End If
        Next
    
        sValue = dr("ContinuousLoop")
        setContinuousLoop(sValue)
        For i = 0 To rblContinuousLoop.Items.Count - 1
            If rblContinuousLoop.Items(i).Value = sValue Then
                rblContinuousLoop.SelectedIndex = i
            End If
        Next
    
        sValue = dr("PauseOnMouseOver")
        setPauseOnMouseOver(sValue)
        For i = 0 To rblPauseOnMouseOver.Items.Count - 1
            If rblPauseOnMouseOver.Items(i).Value = sValue Then
                rblPauseOnMouseOver.SelectedIndex = i
            End If
        Next
    
        sValue = dr("ScrollDirection")
        setScrollDirection(sValue)
        For i = 0 To rblScrollDirection.Items.Count - 1
            If rblScrollDirection.Items(i).Value = sValue Then
                rblScrollDirection.SelectedIndex = i
            End If
        Next
    
        sValue = dr("RotationType")
        setRotationType(sValue)
        For i = 0 To rblDisplayType.Items.Count - 1
            If rblDisplayType.Items(i).Value = sValue Then
                rblDisplayType.SelectedIndex = i
            End If
        Next
    
        sValue = dr("SmoothScrollSpeed")
        setSmoothScrollSpeed(sValue)
        For i = 0 To rblSmoothScrollSpeed.Items.Count - 1
            If rblSmoothScrollSpeed.Items(i).Value = sValue Then
                rblSmoothScrollSpeed.SelectedIndex = i
            End If
        Next
    
        sValue = dr("SlidePause")
        setSlidePause(sValue)
        tbSlidePause.Text = sValue
    
        sValue = dr("ScrollInterval")
        setScrollInterval(sValue)
        tbScrollInterval.Text = sValue
    
        sValue = dr("ShowEffect")
        setShowEffect(sValue)
        For i = 0 To rblShowEffect.Items.Count - 1
            If rblShowEffect.Items(i).Value = sValue Then
                rblShowEffect.SelectedIndex = i
            End If
        Next
    
        sValue = dr("ShowEffectDuration")
        setShowEffectDuration(sValue)
        tbShowEffectDuration.Text = sValue
    
        sValue = dr("HideEffect")
        setHideEffect(sValue)
        For i = 0 To rblHideEffect.Items.Count - 1
            If rblShowEffect.Items(i).Value = sValue Then
                rblHideEffect.SelectedIndex = i
            End If
        Next
    
        sValue = dr("HideEffectDuration")
        setHideEffectDuration(sValue)
        tbHideEffectDuration.Text = sValue
    
        ds = Nothing
    End Sub
    
    Protected Sub CreateNewConfigFile()
        Dim ds As DataSet
        Dim oDataTable1, oDataTable2 As DataTable
        Dim dc As DataColumn
        Dim dr As DataRow
    
        Dim fname As String = Server.MapPath(gsXMLRotatorConfigFilePath)
        Dim fs As New System.IO.FileStream(fname, System.IO.FileMode.Create)
    
        oDataTable1 = New DataTable
    
        dc = New DataColumn
        dc.ColumnName = "Visible"
        dc.DataType = Type.GetType("System.String")
        dc.ColumnMapping = MappingType.Element
        oDataTable1.Columns.Add(dc)
    
        dc = New DataColumn
        dc.ColumnName = "ContinuousLoop"
        dc.DataType = Type.GetType("System.String")
        oDataTable1.Columns.Add(dc)
    
        dc = New DataColumn
        dc.ColumnName = "PauseOnMouseOver"
        dc.DataType = Type.GetType("System.String")
        oDataTable1.Columns.Add(dc)
    
        dc = New DataColumn
        dc.ColumnName = "ScrollDirection"
        dc.DataType = Type.GetType("System.String")
        oDataTable1.Columns.Add(dc)
    
        dc = New DataColumn
        dc.ColumnName = "SlidePause"
        dc.DataType = Type.GetType("System.String")
        oDataTable1.Columns.Add(dc)
    
        dc = New DataColumn
        dc.ColumnName = "ScrollInterval"
        dc.DataType = Type.GetType("System.String")
        oDataTable1.Columns.Add(dc)
    
        dc = New DataColumn
        dc.ColumnName = "RotationType"
        dc.DataType = Type.GetType("System.String")
        oDataTable1.Columns.Add(dc)
    
        dc = New DataColumn
        dc.ColumnName = "SmoothScrollSpeed"
        dc.DataType = Type.GetType("System.String")
        oDataTable1.Columns.Add(dc)
    
        dc = New DataColumn
        dc.ColumnName = "ShowEffect"
        dc.DataType = Type.GetType("System.String")
        oDataTable1.Columns.Add(dc)
    
        dc = New DataColumn
        dc.ColumnName = "ShowEffectDuration"
        dc.DataType = Type.GetType("System.String")
        oDataTable1.Columns.Add(dc)
    
        dc = New DataColumn
        dc.ColumnName = "HideEffect"
        dc.DataType = Type.GetType("System.String")
        oDataTable1.Columns.Add(dc)
    
        dc = New DataColumn
        dc.ColumnName = "HideEffectDuration"
        dc.DataType = Type.GetType("System.String")
        oDataTable1.Columns.Add(dc)
    
        ' build row
        dr = oDataTable1.NewRow
        dr("Visible") = True
        dr("ContinuousLoop") = True
        dr("PauseOnMouseOver") = True
        dr("ScrollDirection") = "Left"
        dr("SlidePause") = "5000"
        dr("ScrollInterval") = "15"
        dr("RotationType") = "SmoothScroll"
        dr("SmoothScrollSpeed") = "Medium"
        dr("ShowEffect") = "None"
        dr("ShowEffectDuration") = "250"
        dr("HideEffect") = "None"
        dr("HideEffectDuration") = "250"
    
        oDataTable1.TableName = TABLENAME_HEADER_ROTATOR
        oDataTable1.Rows.Add(dr)
    
        ds = New DataSet                ' create dataset and add table to it
        ds.DataSetName = "DisplaySettings"
        ds.Tables.Add(oDataTable1)
        
        oDataTable2 = oDataTable1.Copy    ' build next table as a copy of existing table
        oDataTable2.TableName = TABLENAME_NOTICEBOARD_ROTATOR
    
        ds.Tables.Add(oDataTable2)
    
        ' build LHS Body text table
        oDataTable1 = New DataTable
        oDataTable1.TableName = TABLENAME_LHSBODY
    
        dc = New DataColumn
        dc.ColumnName = "Text"
        dc.DataType = Type.GetType("System.String")
        dc.ColumnMapping = MappingType.Element
        oDataTable1.Columns.Add(dc)
    
        dr = oDataTable1.NewRow
        dr(0) = "Your Text Here"
        oDataTable1.Rows.Add(dr)
    
        ds.Tables.Add(oDataTable1)
    
        oDataTable2 = oDataTable1.Copy
        oDataTable2.TableName = TABLENAME_PAGETITLE
    
        ds.Tables.Add(oDataTable2)
    
        ds.WriteXml(fs)
    
        ds = Nothing
        fs.Close()
    End Sub ' CreateNewConfigFile
    
    Protected Sub Item_Button(ByVal sender As Object, ByVal e As DataGridCommandEventArgs)
        Dim tb As TextBox
        LoadDataset()
        Select Case e.CommandName
            Case "InsertBefore"
                Dim dr As DataRow
                dr = NewNewsRow()
                gdt.Rows.InsertAt(dr, e.Item.ItemIndex)
                dgNews.EditItemIndex = e.Item.ItemIndex
    
            Case "InsertAfter"
                Dim dr As DataRow
                dr = NewNewsRow()
                gdt.Rows.InsertAt(dr, e.Item.ItemIndex + 1)
                dgNews.EditItemIndex = e.Item.ItemIndex + 1
            Case "Edit"
                dgNews.EditItemIndex = e.Item.ItemIndex
                If gds.Tables("News").Rows.Count = 1 Then
                    If gds.Tables("News").Rows(0).Item("Date") = "" _
                      And gds.Tables("News").Rows(0).Item("Title") = "" _
                        And gds.Tables("News").Rows(0).Item("Text") = "" Then
                        gds.Tables("News").Clear()
                        Dim dr As DataRow
                        dr = NewNewsRow()
                        gdt.Rows.InsertAt(dr, e.Item.ItemIndex)
                    End If
                End If
    
            Case "Update"
                tb = e.Item.Cells(2).Controls(0)
                gds.Tables("News").Rows(e.Item.ItemIndex).Item("Date") = tb.Text
                tb = e.Item.Cells(3).Controls(0)
                gds.Tables("News").Rows(e.Item.ItemIndex).Item("Title") = tb.Text
                tb = e.Item.Cells(4).Controls(0)
                gds.Tables("News").Rows(e.Item.ItemIndex).Item("Text") = tb.Text
                dgNews.EditItemIndex = -1
    
            Case "Cancel"
                dgNews.EditItemIndex = -1
            Case "Delete"
                If gds.Tables("News").Rows.Count > 1 Then
                    gds.Tables("News").Rows.RemoveAt(e.Item.ItemIndex)
                Else
                    gds.Tables("News").Rows(0).Item("Date") = ""
                    gds.Tables("News").Rows(0).Item("Title") = ""
                    gds.Tables("News").Rows(0).Item("Text") = ""
    
                End If
            Case Else
    
        End Select
        dgNews.DataSource = gds
        dgNews.DataBind()
        SaveDataset()
        If e.CommandName = "Edit" Then
            tb = dgNews.Items(e.Item.ItemIndex).Cells(2).FindControl("tbDate")
        End If
    End Sub
    
    Protected Function NewNewsRow() As DataRow
        gdt = gds.Tables("News")
        Dim row As DataRow = gdt.NewRow
        row(0) = Date.Now.ToLongDateString
        row(1) = "+++ Title +++"
        row(2) = "+++ Text +++"
        Return row
    End Function
    
    Protected Sub LoadDataset()
        Dim fs As New System.IO.FileStream(MapPath(gsXMLNewsContentFilePath), System.IO.FileMode.Open)
        gds.ReadXml(fs)
        fs.Close()
    End Sub
    
    Protected Sub SaveDataset()
        Dim fs As New System.IO.FileStream(MapPath(gsXMLNewsContentFilePath), System.IO.FileMode.Create)
        gds.WriteXml(fs)
        fs.Close()
    End Sub
    
    Protected Sub CreateDataset()
        Dim fs As New System.IO.FileStream(MapPath(gsXMLNewsContentFilePath), System.IO.FileMode.Create)
    
        gds.DataSetName = "DisplaySettings"
        gdt = New DataTable
    
        gcol = New DataColumn
        gcol.ColumnName = "Date"
        gdt.Columns.Add(gcol)
    
        gcol = New DataColumn
        gcol.ColumnName = "Title"
        gdt.Columns.Add(gcol)
    
        gcol = New DataColumn
        gcol.ColumnName = "Text"
        gdt.Columns.Add(gcol)
    
        Dim row As DataRow = gdt.NewRow
    
        row("Date") = "1st January 2009"
        row("Title") = "This is the Title field"
        row("Text") = "This is the Text field"
    
        gdt.TableName = "News"
        gdt.Rows.Add(row)
        gds.Tables.Add(gdt)
    
        gds.WriteXml(fs)
        fs.Close()
    End Sub

    Protected Sub PageIndexChanged(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs)
        dgNews.CurrentPageIndex = e.NewPageIndex
        LoadDataset()
        dgNews.DataSource = gds
        dgNews.DataBind()
    End Sub
    
    Private Sub setVisible(ByVal value As String)
        ViewState("ED_Visible") = value
    End Sub
    
    ReadOnly Property pbVisible() As Boolean
        Get
            Dim o As Object = ViewState("ED_Visible")
            If o Is Nothing Then
                Return True
            End If
            Select Case CStr(o)
                Case "True"
                    Return True
                Case "False"
                    Return False
                Case Else
                    Return True
            End Select
        End Get
    End Property
    
    Private Sub setContinuousLoop(ByVal value As String)
        ViewState("ED_ContinuousLoop") = value
    End Sub
    
    ReadOnly Property pbContinuousLoop() As Boolean
        Get
            Dim o As Object = ViewState("ED_ContinuousLoop")
            If o Is Nothing Then
                Return True
            End If
            Select Case CStr(o)
                Case "True"
                    Return True
                Case "False"
                    Return False
                Case Else
                    Return True
            End Select
        End Get
    End Property
    
    Private Sub setPauseOnMouseOver(ByVal value As String)
        ViewState("ED_PauseOnMouseOver") = value
    End Sub
    
    ReadOnly Property pbPauseOnMouseOver() As Boolean
        Get
            Dim o As Object = ViewState("ED_PauseOnMouseOver")
            If o Is Nothing Then
                Return True
            End If
            Select Case CStr(o)
                Case "True"
                    Return True
                Case "False"
                    Return False
                Case Else
                    Return True
            End Select
        End Get
    End Property
    
    ' Sets property from RadioButtonList - Can't do this as a Set Property because we are passing in a string but it's encapsulated like this to keep it with the Get Property
    
    Private Sub setScrollDirection(ByVal value As String)
        ViewState("ED_ScrollDirection") = value.ToString
    End Sub
    
    ReadOnly Property penumScrollDirection() As ComponentArt.Web.UI.ScrollDirection
        Get
            Dim o As Object = ViewState("ED_ScrollDirection")
            If o Is Nothing Then
                Return ComponentArt.Web.UI.ScrollDirection.Up
            End If
            Select Case CStr(o)
                Case "Up"
                    Return ComponentArt.Web.UI.ScrollDirection.Up
                Case "Left"
                    Return ComponentArt.Web.UI.ScrollDirection.Left
                Case Else
                    Return ComponentArt.Web.UI.ScrollDirection.Left
            End Select
        End Get
    End Property
    
    Private Sub setSmoothScrollSpeed(ByVal value As String)
        ViewState("ED_SmoothScrollSpeed") = value.ToString
    End Sub
    
    ReadOnly Property penumSmoothScrollSpeed() As ComponentArt.Web.UI.SmoothScrollSpeed
        Get
            Dim o As Object = ViewState("ED_SmoothScrollSpeed")
            If o Is Nothing Then
                Return ComponentArt.Web.UI.SmoothScrollSpeed.Medium
            End If
            Select Case CStr(o)
                Case "Slow"
                    Return ComponentArt.Web.UI.SmoothScrollSpeed.Slow
                Case "Medium"
                    Return ComponentArt.Web.UI.SmoothScrollSpeed.Medium
                Case "Fast"
                    Return ComponentArt.Web.UI.SmoothScrollSpeed.Fast
                Case Else
                    Return ComponentArt.Web.UI.SmoothScrollSpeed.Medium
            End Select
        End Get
    End Property
    
    Private Sub setRotationType(ByVal value As String)
        ViewState("ED_RotationType") = value.ToString
    End Sub
    
    ReadOnly Property penumRotationType() As ComponentArt.Web.UI.RotationType
        Get
            Dim o As Object = ViewState("ED_RotationType")
            If o Is Nothing Then
                Return ComponentArt.Web.UI.RotationType.ContentScroll
            End If
            Select Case CStr(o)
                Case "ContentScroll"
                    Return ComponentArt.Web.UI.RotationType.ContentScroll
                Case "SlideShow"
                    Return ComponentArt.Web.UI.RotationType.SlideShow
                Case Else
                    Return ComponentArt.Web.UI.RotationType.ContentScroll
            End Select
        End Get
    End Property
    
    Private Sub setSlidePause(ByVal value As Integer)
        ViewState("ED_SlidePause") = value
    End Sub
    
    ReadOnly Property pnSlidePause() As Integer
        Get
            Dim o As Object = ViewState("ED_SlidePause")
            If o Is Nothing Then
                Return 15
            End If
            Return CInt(o)
        End Get
    End Property
    
    Private Sub setScrollInterval(ByVal value As Integer)
        ViewState("ED_ScrollInterval") = value
    End Sub
    
    ReadOnly Property pnScrollInterval() As Integer
        Get
            Dim o As Object = ViewState("ED_ScrollInterval")
            If o Is Nothing Then
                Return 15
            End If
            Return CInt(o)
        End Get
    End Property
    
    Private Sub setShowEffect(ByVal value As String)
        ViewState("ED_ShowEffect") = value.ToString
    End Sub
    
    ReadOnly Property penumShowEffect() As ComponentArt.Web.UI.RotationEffect
        Get
            Dim o As Object = ViewState("ED_ShowEffect")
            If o Is Nothing Then
                Return ComponentArt.Web.UI.RotationEffect.None
            End If
            Select Case CStr(o)
                Case "None"
                    Return ComponentArt.Web.UI.RotationEffect.None
                Case "Fade"
                    Return ComponentArt.Web.UI.RotationEffect.Fade
                Case "Pixelate"
                    Return ComponentArt.Web.UI.RotationEffect.Pixelate
                Case "Dissolve"
                    Return ComponentArt.Web.UI.RotationEffect.Dissolve
                Case "GradientWipe"
                    Return ComponentArt.Web.UI.RotationEffect.GradientWipe
                Case Else
                    Return ComponentArt.Web.UI.RotationEffect.None
            End Select
        End Get
    End Property
    
    Private Sub setShowEffectDuration(ByVal value As Integer)
        ViewState("ED_ShowEffectDuration") = value
    End Sub
    
    ReadOnly Property pnShowEffectDuration() As Integer
        Get
            Dim o As Object = ViewState("ED_ShowEffectDuration")
            If o Is Nothing Then
                Return 250
            End If
            Return CInt(o)
        End Get
    End Property
    
    Private Sub setHideEffect(ByVal value As String)
        ViewState("ED_HideEffect") = value.ToString
    End Sub
    
    ReadOnly Property penumHideEffect() As ComponentArt.Web.UI.RotationEffect
        Get
            Dim o As Object = ViewState("ED_HideEffect")
            If o Is Nothing Then
                Return ComponentArt.Web.UI.RotationEffect.None
            End If
            Select Case CStr(o)
                Case "None"
                    Return ComponentArt.Web.UI.RotationEffect.None
                Case "Fade"
                    Return ComponentArt.Web.UI.RotationEffect.Fade
                Case "Pixelate"
                    Return ComponentArt.Web.UI.RotationEffect.Pixelate
                Case "Dissolve"
                    Return ComponentArt.Web.UI.RotationEffect.Dissolve
                Case "GradientWipe"
                    Return ComponentArt.Web.UI.RotationEffect.GradientWipe
                Case Else
                    Return ComponentArt.Web.UI.RotationEffect.None
            End Select
        End Get
    End Property
    
    Private Sub setHideEffectDuration(ByVal value As Integer)
        ViewState("ED_HideEffectDuration") = value
    End Sub
    
    ReadOnly Property pnHideEffectDuration() As Integer
        Get
            Dim o As Object = ViewState("ED_HideEffectDuration")
            If o Is Nothing Then
                Return 250
            End If
            Return CInt(o)
        End Get
    End Property
    
    Sub btnShowTextEditor_Click(ByVal sender As Object, ByVal e As EventArgs)
        pnlTextEditor.Visible = True
        pnlNewsEditor.Visible = False
        pnlRotatorEditor.Visible = False
    End Sub
    
    Sub btnShowNewsEditor_Click(ByVal sender As Object, ByVal e As EventArgs)
        pnlNewsEditor.Visible = True
        pnlTextEditor.Visible = False
        pnlRotatorEditor.Visible = False
    End Sub
    
    Sub btnShowRotatorEditor_Click(ByVal sender As Object, ByVal e As EventArgs)
        pnlRotatorEditor.Visible = True
        pnlTextEditor.Visible = False
        pnlNewsEditor.Visible = False
    End Sub
    
    Sub btnResetToDefaults_Click(ByVal sender As Object, ByVal e As EventArgs)
        rblVisible.SelectedIndex = 0
        rblContinuousLoop.SelectedIndex = 0
        rblPauseOnMouseOver.SelectedIndex = 0
        rblScrollDirection.SelectedIndex = 0
        rblDisplayType.SelectedIndex = 0
        rblSmoothScrollSpeed.SelectedIndex = 0
        tbSlidePause.Text = "5000"
        tbScrollInterval.Text = "15"
        rblShowEffect.SelectedIndex = 0
        tbShowEffectDuration.Text = "500"
        rblHideEffect.SelectedIndex = 0
        tbHideEffectDuration.Text = "500"
    End Sub

    Protected Sub rbNoticeBoardPage_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call SaveTextChanges()
        Call InitArea()
    End Sub

    Protected Sub rbLoginPage_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call SaveTextChanges()
        Call InitArea()
    End Sub

    Protected Sub ddlAreaSelector_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call SaveTextChanges()
        Call InitArea()
    End Sub
    
    Protected Sub btnSaveChanges_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If pnlTextEditor.Visible Then
            Call SaveTextChanges()
        ElseIf pnlNewsEditor.Visible Then
            Call SaveNewsChanges()
        ElseIf pnlRotatorEditor.Visible Then
            Call SaveRotatorChanges()
        End If
    End Sub
    
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
        Call SaveTextChanges()
        Call InitAreaSelector()
    End Sub
    
</script>

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Notice Board Editor</title>
    <style type="text/css" media="screen">
    BODY {
	FONT-FAMILY: Verdana
}
UNKNOWN {
	window: editorInstance.Description
}
TABLE {
	FONT-SIZE: 7pt; FONT-FAMILY: verdana
}
TD.small {
	FONT-SIZE: 8pt; FONT-FAMILY: verdana
}
TD.subheading {
	FONT-SIZE: 14pt; FONT-FAMILY: sans-serif
}
TR.darkbackground {
	BACKGROUND-COLOR: silver
}
bold {
}
</style>
    <link href="~/css/sprint.css" rel="stylesheet" type="text/css" />
    <link href="~/css/sprint.css" rel="stylesheet" type="text/css" />
    <link href="~/css/sprint.css" rel="stylesheet" type="text/css" />
    <link href="~/css/sprint.css" rel="stylesheet" type="text/css" />
</head>
<body class="sf">
    <form id="frmNoticeBoardEditor" runat="server">
        <main:Header ID="ctlHeader" runat="server"></main:Header>
        <table width="100%" cellpadding="0" cellspacing="0">
            <tr class="bar_noticeboardeditor">
                <td style="width:50%; white-space:nowrap; height: 13px;">
                </td>
                <td style="width:50%; white-space:nowrap; height: 13px;" align="right">
                </td>
            </tr>
            <tr>
                <td style="width: 75%; white-space: nowrap; height: 24px;">
        <asp:Button ID="btnShowTextEditor" OnClick="btnShowTextEditor_Click" runat="server"
            Text="text editor" />
        <asp:Button ID="btnShowNewsEditor" OnClick="btnShowNewsEditor_Click" runat="server"
            Text="news editor" />
        <asp:Button ID="btnShowRotatorEditor" OnClick="btnShowRotatorEditor_Click" runat="server"
            Text="rotator editor" />
                    &nbsp; &nbsp;
        <asp:Button ID="btnSaveChanges" runat="server" Text="save changes" OnClick="btnSaveChanges_Click" />
        <asp:Label ID="lblMessage" runat="server" Font-Names="Verdana" Font-Size="XX-Small"></asp:Label></td>
                <td align="right" style="width: 25%; white-space: nowrap; height: 24px;">
                    <asp:LinkButton ID="lnkbtnHelp" runat="server" OnClientClick='window.open("help_noticeboardeditor.pdf", "NBHelp","top=10,left=10,width=700,height=450,status=no,toolbar=yes,address=no,menubar=yes,resizable=yes,scrollbars=yes");'>editor help</asp:LinkButton></td>
            </tr>
        </table>
        &nbsp; &nbsp;&nbsp;
        <asp:Panel ID="pnlTextEditor" runat="server" Visible="true" Width="100%">
            <table width="100%">
                <tr>
                    <td style="width: 33%">
                        <strong>Text Editor</strong>
                    </td>
                    <td style="width: 34%">
                    </td>
                    <td style="width: 33%">
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
                            Text="show advanced features" Enabled="False" />
                        &nbsp;&nbsp;&nbsp;
                        <asp:Label ID="lblLegendCSSAttributes" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="CSS attributes:" Visible="False"/>
                        <asp:TextBox ID="tbCSSAttributes" runat="server" Font-Names="Arial" Font-Size="XX-Small"
                            Width="500px" Visible="False"/>&nbsp;&nbsp;&nbsp;
                    </td>
                </tr>
                <tr id="trLoginPageAdvancedControls" visible="false" class="darkbackground" runat="server">
                    <td class="small"><asp:Label ID="Label2" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="Login page column split:" Visible="False" Font-Bold="True"/>
                        &nbsp;
                        <asp:Label ID="Label4" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="left %:" Visible="False"/>
                        <asp:TextBox ID="tbLPLeftSplit" runat="server" Font-Names="Arial" Font-Size="XX-Small"
                            Width="30px" Visible="False"/>&nbsp;
                        <asp:Label ID="Label5" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="right %:" Visible="False"/><asp:TextBox ID="tbLPRightSplit" runat="server" Font-Names="Arial" Font-Size="XX-Small"
                            Width="30px" Visible="False"/>
                        &nbsp;&nbsp;
                        <asp:Label ID="Label8" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="left subsplit:" Visible="False" Font-Bold="True"/>
                        <asp:Label ID="Label6" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="left %:" Visible="False"/><asp:TextBox ID="tbLPLeftLeft" runat="server" Font-Names="Arial" Font-Size="XX-Small"
                            Width="30px" Visible="False"/><asp:Label ID="Label7" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="right %:" Visible="False"/><asp:TextBox ID="tbLPLeftRight" runat="server" Font-Names="Arial" Font-Size="XX-Small"
                            Width="30px" Visible="False"/>
                        &nbsp;&nbsp;
                        <asp:Label ID="Label9" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="right subsplit:" Visible="False" Font-Bold="True"/>
                        <asp:Label ID="Label10" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="left %:" Visible="False"/><asp:TextBox ID="tbLPRightLeft" runat="server" Font-Names="Arial" Font-Size="XX-Small"
                            Width="30px" Visible="False"/><asp:Label ID="Label11" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="right %:" Visible="False"/><asp:TextBox ID="tbLPRightRight" runat="server" Font-Names="Arial" Font-Size="XX-Small"
                            Width="30px" Visible="False"/>
                        &nbsp;&nbsp; &nbsp;
                        <asp:Label ID="lblLegendLoginBoxPosition" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="Login box position:" Visible="False" Font-Bold="True"/><asp:DropDownList ID="ddlLoginBoxPosition" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Visible="False">
                            <asp:ListItem>top</asp:ListItem>
                            <asp:ListItem>left</asp:ListItem>
                            <asp:ListItem>left centre</asp:ListItem>
                            <asp:ListItem>right centre</asp:ListItem>
                            <asp:ListItem>right</asp:ListItem>
                            <asp:ListItem>bottom</asp:ListItem>
                        </asp:DropDownList></td>
                </tr>
                <tr id="trNoticeBoard1AdvancedControls" runat="server" class="darkbackground" visible="false">
                    <td class="small"><asp:Label ID="Label3" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="Notice board column split:" Visible="False" Font-Bold="True"/>&nbsp;
                        <asp:Label ID="Label12" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="left %:" Visible="False"/><asp:TextBox ID="tbNB1LeftSplit" runat="server" Font-Names="Arial" Font-Size="XX-Small"
                            Width="30px" Visible="False"/><asp:Label ID="Label13" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="right %:" Visible="False"/><asp:TextBox ID="tbNB1RightSplit" runat="server" Font-Names="Arial" Font-Size="XX-Small"
                            Width="30px" Visible="False"/></td>
                </tr>
            </table>
            <FCKeditorV2:FCKeditor ID="FCKeditor1" runat="server" ToolbarSet="CourierSoftware"
                BasePath="./fckeditor/" Value='This is some <strong>sample text</strong>. You are using <a href="http://www.fckeditor.net/">FCKeditor</a>.'>
            </FCKeditorV2:FCKeditor>
        </asp:Panel>
        <br />
        <asp:Panel ID="pnlNewsEditor" runat="server" Visible="false" Width="100%">
            <table width="95%">
                <tr>
                    <td class="subheading" style="width:33%">
                        <strong>
                            <asp:Label ID="Label1" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Scrolling News Editor"></asp:Label></strong></td>
                    <td style="width:34%">
                    </td>
                    <td style="width:33%">
                    </td>
                </tr>
            </table>
            <asp:DataGrid ID="dgNews" runat="server" Height="280px" CellSpacing="1" AutoGenerateColumns="False"
                OnItemCommand="Item_Button" BorderWidth="0px" BorderStyle="Dotted" BorderColor="Silver"
                Font-Names="Verdana" Font-Size="XX-Small" OnPageIndexChanged="PageIndexChanged"
                Width="100%">
                <AlternatingItemStyle BackColor="#E0E0E0"></AlternatingItemStyle>
                <Columns>
                    <asp:TemplateColumn>
                        <ItemTemplate>
                            <asp:ImageButton ID="ImageButton2" runat="server" ImageUrl="./images/newentryabove.gif"
                                CommandName="InsertBefore"></asp:ImageButton>
                            <asp:ImageButton ID="ImageButton3" runat="server" ImageUrl="./images/newentrybelow.gif"
                                CommandName="InsertAfter"></asp:ImageButton>
                        </ItemTemplate>
                    </asp:TemplateColumn>
                    <asp:EditCommandColumn ButtonType="PushButton" UpdateText="update" CancelText="cancel"
                        EditText="edit">
                        <HeaderStyle Width="39px"></HeaderStyle>
                    </asp:EditCommandColumn>
                    <asp:BoundColumn DataField="Date" HeaderText="&lt;b&gt;Date&lt;/b&gt;"></asp:BoundColumn>
                    <asp:BoundColumn DataField="Title" HeaderText="&lt;b&gt;Item Title&lt;/b&gt;"></asp:BoundColumn>
                    <asp:BoundColumn DataField="Text" HeaderText="&lt;b&gt;Item Text&lt;/b&gt;"></asp:BoundColumn>
                    <asp:TemplateColumn>
                        <ItemTemplate>
                            <asp:ImageButton ID="ImageButton1" runat="server" ImageUrl="./images/delete.gif"
                                CommandName="Delete"></asp:ImageButton>
                        </ItemTemplate>
                    </asp:TemplateColumn>
                </Columns>
            </asp:DataGrid>
        </asp:Panel>
        <br />
        <asp:Panel ID="pnlRotatorEditor" runat="server" Visible="false" Width="100%">
            <table width="95%">
                <tr>
                    <td style="width: 33%; height: 14px;">
                        <strong>Text Rotator Editor</strong></td>
                    <td style="width: 34%; height: 14px;">
                    </td>
                    <td style="width: 33%; height: 14px;">
                    </td>
                </tr>
            </table>
            <asp:RequiredFieldValidator ID="rfvScrollInterval" runat="server" Display="None"
                ControlToValidate="tbScrollInterval" ErrorMessage="Value required for Scroll Interval" Font-Names="Vrinda" Font-Size="XX-Small"></asp:RequiredFieldValidator>
            &nbsp;<asp:RangeValidator ID="rvScrollInterval" runat="server" Display="None" ControlToValidate="tbScrollInterval"
                ErrorMessage="Scroll Interval must be numeric" MaximumValue="99999" MinimumValue="1" Font-Names="Vrinda" Font-Size="XX-Small"></asp:RangeValidator>
            &nbsp;&nbsp;<asp:RequiredFieldValidator ID="rfvSlidePause" runat="server" Display="None"
                ControlToValidate="tbSlidePause" ErrorMessage="Value required for Slide Pause" Font-Names="Vrinda" Font-Size="XX-Small"></asp:RequiredFieldValidator>
            <asp:RangeValidator ID="rvSlidePause" runat="server" Display="None" ControlToValidate="tbSlidePause"
                ErrorMessage="Slide Pause must be numeric" MaximumValue="9999" MinimumValue="1" Font-Names="Vrinda" Font-Size="XX-Small"></asp:RangeValidator>
            &nbsp;<asp:RequiredFieldValidator ID="rfvShowEffectDuration" runat="server" Display="None"
                ControlToValidate="tbShowEffectDuration" ErrorMessage="Value required for Show Effect Duration" Font-Names="Vrinda" Font-Size="XX-Small"></asp:RequiredFieldValidator>
            <asp:RangeValidator ID="rvShowEffectDuration" runat="server" Display="None" ControlToValidate="tbShowEffectDuration"
                ErrorMessage="Show Effect Duration must be numeric" MaximumValue="99999" MinimumValue="1" Font-Names="Vrinda" Font-Size="XX-Small"></asp:RangeValidator>
            &nbsp;<asp:RequiredFieldValidator ID="rfvHideEffectDuration" runat="server" Display="None"
                ControlToValidate="tbHideEffectDuration" ErrorMessage="Value required for Hide Effect Duration" Font-Names="Vrinda" Font-Size="XX-Small"></asp:RequiredFieldValidator>
            <asp:RangeValidator ID="rvHideEffectDuration" runat="server" Display="None" ControlToValidate="tbHideEffectDuration"
                ErrorMessage="Hide Effect Duration must be numeric" MaximumValue="99999" MinimumValue="1" Font-Names="Vrinda" Font-Size="XX-Small"></asp:RangeValidator>
            <br />
            <asp:ValidationSummary ID="ValidationSummary1" runat="server" Font-Names="Verdana" Font-Size="XX-Small"></asp:ValidationSummary>
            <table width="100%">
                <tbody>
                    <tr class="darkbackground">
                        <td class="small" align="right">
                            <strong> Edit:</strong></td>
                        <td class="small">
                            <asp:RadioButtonList ID="rblRotatorTarget" runat="server"
                                AutoPostBack="True" RepeatDirection="Horizontal">
                                <asp:ListItem Value="NoticeBoard1Rotator" Selected="True">notice board rotator</asp:ListItem>
                                <asp:ListItem Value="HeaderRotator">top of page rotator (visible on all tabs)</asp:ListItem>
                            </asp:RadioButtonList>
                        </td>
                    </tr>
                </tbody>
            </table>
            <ComponentArt:Rotator ID="Rotator1" runat="server" Visible="<%# pbVisible %>" XmlContentFile="example_rotator.xml"
                SlidePause="<%# pnSlidePause %>" PauseOnMouseOver="<%# pbPauseOnMouseOver %>"
                CssClass="Rotator" HideEffectDuration="<%# pnHideEffectDuration %>" HideEffect="<%# penumHideEffect %>"
                ShowEffectDuration="<%# pnShowEffectDuration %>" ShowEffect="<%# penumShowEffect %>"
                SmoothScrollSpeed="<%# penumSmoothScrollSpeed %>" RotationType="<%# penumRotationType %>"
                ScrollInterval="<%# pnScrollInterval %>" ScrollDirection="<%# penumScrollDirection %>"
                Loop="<%# pbContinuousLoop %>" Height="50" Width="500">
                <SlideTemplate>
                    <table cellspacing="1" cellpadding="0" width="100%" border="0">
                        <tr>
                            <td class="RotatorMain">
                                <span>
                                    <img alt="" src='./images/rotatorExampleImage.jpg' height="44" /><img alt="" src="./images/blank.gif"
                                        width="10" border="0" />
                                </span>
                            </td>
                            <td class="RotatorMain" style="white-space:nowrap">
                                <span class="AdRotatorText">This is example text to show the selected effect</span>
                            </td>
                        </tr>
                    </table>
                </SlideTemplate>
            </ComponentArt:Rotator>
            <p>
                <asp:Table ID="Table3" runat="server">
                    <asp:TableRow Width="95%">
                        <asp:TableCell Width="50%">
                            <asp:Table ID="Table2" runat="server" Width="362px">
                                <asp:TableRow ID="Vsble" BackColor="#E0E0E0">
                                    <asp:TableCell Text="Visible:"></asp:TableCell>
                                    <asp:TableCell>
                                        <asp:RadioButtonList runat="server" ID="rblVisible" RepeatDirection="Horizontal">
                                            <asp:ListItem Value="True" Selected="True">True</asp:ListItem>
                                            <asp:ListItem Value="False">False</asp:ListItem>
                                        </asp:RadioButtonList>
                                    </asp:TableCell>
                                    <asp:TableCell>
                                        <asp:Button ID="btnResetToDefaults" OnClick="btnResetToDefaults_Click" runat="server"
                                            Text="reset to defaults"></asp:Button>
                                    </asp:TableCell>
                                </asp:TableRow>
                                <asp:TableRow ID="Loop" BackColor="#E0E0E0">
                                    <asp:TableCell Text="Continuous Loop:
                    "></asp:TableCell>
                                    <asp:TableCell>
                                        <asp:RadioButtonList runat="server" ID="rblContinuousLoop" RepeatDirection="Horizontal">
                                            <asp:ListItem Value="True" Selected="True">True</asp:ListItem>
                                            <asp:ListItem Value="False">False</asp:ListItem>
                                        </asp:RadioButtonList>
                                    </asp:TableCell>
                                    <asp:TableCell></asp:TableCell>
                                </asp:TableRow>
                                <asp:TableRow BackColor="Silver">
                                    <asp:TableCell Text="
                    Pause On&lt;br /&gt;
                        Mouse Over:
                    "></asp:TableCell>
                                    <asp:TableCell>
                                        <asp:RadioButtonList runat="server" ID="rblPauseOnMouseOver" RepeatDirection="Horizontal">
                                            <asp:ListItem Value="True" Selected="True">True</asp:ListItem>
                                            <asp:ListItem Value="False">False</asp:ListItem>
                                        </asp:RadioButtonList>
                                    </asp:TableCell>
                                    <asp:TableCell></asp:TableCell>
                                </asp:TableRow>
                                <asp:TableRow BackColor="#E0E0E0">
                                    <asp:TableCell Text="
                    Scroll Direction:
                    "></asp:TableCell>
                                    <asp:TableCell>
                                        <asp:RadioButtonList runat="server" ID="rblScrollDirection" RepeatDirection="Horizontal">
                                            <asp:ListItem Value="Up" Selected="True">Up</asp:ListItem>
                                            <asp:ListItem Value="Left">Left</asp:ListItem>
                                        </asp:RadioButtonList>
                                    </asp:TableCell>
                                    <asp:TableCell></asp:TableCell>
                                </asp:TableRow>
                                <asp:TableRow BackColor="#E0E0E0">
                                    <asp:TableCell Text="
                    Smooth Scroll Speed:
                    "></asp:TableCell>
                                    <asp:TableCell>
                                        <asp:RadioButtonList runat="server" ID="rblSmoothScrollSpeed" RepeatDirection="Horizontal">
                                            <asp:ListItem Value="Slow" Selected="True">Slow</asp:ListItem>
                                            <asp:ListItem Value="Medium">Medium</asp:ListItem>
                                            <asp:ListItem Value="Fast">Fast</asp:ListItem>
                                        </asp:RadioButtonList>
                                    </asp:TableCell>
                                    <asp:TableCell>
                                        Scroll Interval (ms)
                                        <asp:TextBox runat="server" ID="tbScrollInterval">15</asp:TextBox>
                                    </asp:TableCell>
                                </asp:TableRow>
                                <asp:TableRow BackColor="Silver">
                                    <asp:TableCell Text="Display Type:"></asp:TableCell>
                                    <asp:TableCell>
                                        <asp:RadioButtonList runat="server" ID="rblDisplayType" RepeatDirection="Horizontal">
                                            <asp:ListItem Value="SmoothScroll" Selected="True">Scrolling</asp:ListItem>
                                            <asp:ListItem Value="SlideShow">Slide Show</asp:ListItem>
                                        </asp:RadioButtonList>
                                    </asp:TableCell>
                                    <asp:TableCell>
                                        Slide Pause (ms)
                                        <asp:TextBox runat="server" ID="tbSlidePause">5000</asp:TextBox>
                                    </asp:TableCell>
                                </asp:TableRow>
                                <asp:TableRow BackColor="#E0E0E0">
                                    <asp:TableCell Text="
                        Show Effect:
                    "></asp:TableCell>
                                    <asp:TableCell>
                                        <asp:RadioButtonList runat="server" ID="rblShowEffect" RepeatDirection="Horizontal">
                                            <asp:ListItem Value="None" Selected="True">None</asp:ListItem>
                                            <asp:ListItem Value="Fade">Fade</asp:ListItem>
                                            <asp:ListItem Value="Pixelate">Pixelate</asp:ListItem>
                                            <asp:ListItem Value="Dissolve">Dissolve</asp:ListItem>
                                            <asp:ListItem Value="GradientWipe">Gradient&#160;Wipe</asp:ListItem>
                                        </asp:RadioButtonList>
                                    </asp:TableCell>
                                    <asp:TableCell>
                                        Show Effect Duration (ms)
                                        <asp:TextBox runat="server" ID="tbShowEffectDuration">250</asp:TextBox>
                                    </asp:TableCell>
                                </asp:TableRow>
                                <asp:TableRow BackColor="#E0E0E0">
                                    <asp:TableCell Text="
                        HideEffect:
                    "></asp:TableCell>
                                    <asp:TableCell>
                                        <asp:RadioButtonList runat="server" ID="rblHideEffect" RepeatDirection="Horizontal">
                                            <asp:ListItem Value="None" Selected="True">None</asp:ListItem>
                                            <asp:ListItem Value="Fade">Fade</asp:ListItem>
                                            <asp:ListItem Value="Pixelate">Pixelate</asp:ListItem>
                                            <asp:ListItem Value="Dissolve">Dissolve</asp:ListItem>
                                            <asp:ListItem Value="GradientWipe">Gradient&#160;Wipe</asp:ListItem>
                                        </asp:RadioButtonList>
                                    </asp:TableCell>
                                    <asp:TableCell>
                                        Hide Effect Duration (ms)
                                        <asp:TextBox runat="server" ID="tbHideEffectDuration">250</asp:TextBox>
                                    </asp:TableCell>
                                </asp:TableRow>
                            </asp:Table>
                        </asp:TableCell>
                        <asp:TableCell Width="50%"></asp:TableCell>
                    </asp:TableRow>
                </asp:Table>
            </p>
        </asp:Panel>
        <asp:Panel ID="HelpTextPanel" runat="server" Width="100%">
        </asp:Panel>
    </form>
</body>
</html>