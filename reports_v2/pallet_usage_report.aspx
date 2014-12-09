<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>
<%@ import Namespace="System.Collections.Generic" %>
<script runat="server">

    '   Pallet Usage Report
    '
    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString()
    Private gsMonthNames() As String = {"", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"}
    Dim gnPrevYear As Int32
    Dim gnPrevMonth As Int32

    Protected Sub Page_Load(ByVal Source As Object, ByVal E As EventArgs)
        If Not IsNumeric(Session("CustomerKey")) Then
            Server.Transfer("../session_expired.aspx")
        End If
        If Not IsPostBack Then
            lblReportGeneratedDateTime.Text = "Report generated: " & Now().ToString("dd-MMM-yy HH:mm")
            Call InitData()
        End If
    End Sub
    
    Protected Sub InitData()
        Dim sSQL As String = String.Empty
        Dim dt As DataTable
        Dim sMonths As String() = {"", "J", "F", "M", "A", "M", "J", "J", "A", "S", "O", "N", "D", "J", "F", "M", "A", "M", "J", "J", "A", "S", "O", "N", "D", "J", "F", "M", "A", "M", "J", "J", "A", "S", "O", "N", "D", "J", "F", "M", "A", "M", "J", "J", "A", "S", "O", "N", "D", "J", "F", "M", "A", "M", "J", "J", "A", "S", "O", "N", "D"}
        sSQL = "SELECT SUM(Quantity) 'Total', Year, Month FROM PalletUsage WHERE CustomerKey = " & Session("CustomerKey") & " GROUP BY Month, Year"
        dt = ComLib.ExecuteQueryToDataTable(sSQL)
        Dim lstPalletUsage As New List(Of Double)
        Dim nStartMonth As Int32
        Dim nStartYear As Int32

        If dt.Rows.Count > 0 Then
            nStartMonth = dt.Rows(0).Item("Month")
            nStartYear = dt.Rows(0).Item("Year")
            If dt.Rows(0).Item("Month") = 1 Then
                gnPrevYear = nStartYear - 1
                gnPrevMonth = 12
            Else
                gnPrevYear = nStartYear
                gnPrevMonth = dt.Rows(0).Item("Month") - 1
            End If
            
            For Each dr As DataRow In dt.Rows
                Dim nYear As Int32 = dr("Year")
                Dim nMonth As Int32 = dr("Month")
                Dim dblTotal As Double = dr("Total")
            
                If nMonth = 1 Then
                    While Not gnPrevMonth = 12 And gnPrevYear = (nYear - 1)
                        lstPalletUsage.Add(-1)
                        Call IncrementPrevMonthYear()
                    End While
                Else
                    While Not gnPrevMonth = (nMonth - 1) And gnPrevYear = nYear
                        lstPalletUsage.Add(-1)
                        Call IncrementPrevMonthYear()
                    End While
                End If
                lstPalletUsage.Add(dblTotal)
                Call IncrementPrevMonthYear()
            Next

            Dim sMonthAxis As String = String.Empty
            For i As Int32 = nStartMonth To (nStartMonth + lstPalletUsage.Count)-1
                sMonthAxis += "|" & sMonths(i)
            Next
            Dim nAxisYear As Int32 = nStartYear
            Dim sYearAxis As String = "|" & nAxisYear.ToString
            For i As Int32 = nStartMonth + 1 To (nStartMonth + lstPalletUsage.Count) - 1
                sYearAxis += "|"
                If (i - 1) Mod 12 = 0 Then
                    nAxisYear += 1
                    sYearAxis += nAxisYear.ToString
                End If
            Next
            gnPrevYear = dt.Rows(0).Item("Year")   ' reuse these variables as increment routine already present
            gnPrevMonth = dt.Rows(0).Item("Month")

            Dim dblMinValue As Double = Double.MaxValue
            Dim dblMaxValue As Double = Double.MinValue

            For Each dblPalletCount As Double In lstPalletUsage
                If dblPalletCount < dblMinValue Then
                    dblMinValue = dblPalletCount
                End If
                If dblPalletCount > dblMaxValue Then
                    dblMaxValue = dblPalletCount
                End If
            Next
            Dim nPixelsX As Int32 = 800
            Dim nPixelsY As Int32 = 300 + CInt(dblMaxValue / 2)
            Dim nPixelsTotal As Int32 = nPixelsX * nPixelsY
            While nPixelsTotal > 300000
                nPixelsX = nPixelsX - 10
                nPixelsY = nPixelsY - 5
                nPixelsTotal = nPixelsX * nPixelsY
            End While
            
            Dim sbChart As New StringBuilder
            sbChart.Append("http://chart.apis.google.com/chart?")
            sbChart.Append("cht=bvs")
            sbChart.Append("&")
            sbChart.Append("chbh=25,6")
            sbChart.Append("&")
            ' sbChart.Append("chs=800x" & (300 + CInt(dblMaxValue / 2)).ToString)
            sbChart.Append("chs=" & nPixelsX.ToString & "x" & nPixelsY.ToString)
            sbChart.Append("&")
            sbChart.Append("chxt=x,y,x,y")
            sbChart.Append("&")
            sbChart.Append("chd=")
            sbChart.Append("t:")
            sbChart.Append(CommaSeparateList(lstPalletUsage))    ' data series
            sbChart.Append("&")
            sbChart.Append("chds=0," & (dblMaxValue + 10).ToString)
            sbChart.Append("&")
            sbChart.Append("chxr=1,0," & (dblMaxValue + 10).ToString & ",10")    ' axis index of Y axis
            sbChart.Append("&")
            sbChart.Append("chco=008000")
            sbChart.Append("&")
            sbChart.Append("chm=B,A5CE84,0,0,0")
            sbChart.Append("&")
            sbChart.Append("chtt=Pallet+storage+by+month")
            sbChart.Append("&")
            sbChart.Append("chm=N,000000,0,-1,11")
            sbChart.Append("&")
            sbChart.Append("chxl=")
            sbChart.Append("0:")
            sbChart.Append(sMonthAxis)  ' eg |J|F|M|A|M
            ' sbChart.Append("|2:|" & nStartYear.ToString)
            sbChart.Append("|2:" & sYearAxis)
            sbChart.Append("|3:||No+of+pallets")
            imgChart.ImageUrl = sbChart.ToString
            imgChart.Visible=True
        lblNoInformationAvailable.Visible=false
        Else
            lblNoInformationAvailable.Visible = True
            imgChart.Visible = False
        End If
    End Sub
    
    Protected Sub IncrementPrevMonthYear()
        If gnPrevMonth < 12 Then
            gnPrevMonth += 1
        Else
            gnPrevMonth = 1
            gnPrevYear += 1
        End If
    End Sub
    
    Private Function CommaSeparateList(ByVal lstList As List(Of Integer)) As String
        Dim s As String = String.Empty
        For i As Integer = 0 To lstList.Count - 2
            s = s & lstList(i) & ","
        Next
        s = s & lstList(lstList.Count - 1)
        CommaSeparateList = s
    End Function
            
    Private Function CommaSeparateList(ByVal lstList As List(Of String)) As String
        Dim s As String = String.Empty
        For i As Integer = 0 To lstList.Count - 2
            s = s & lstList(i) & ","
        Next
        s = s & lstList(lstList.Count - 1)
        CommaSeparateList = s
    End Function

    Private Function CommaSeparateList(ByVal lstList As List(Of Double)) As String
        Dim s As String = String.Empty
        For i As Integer = 0 To lstList.Count - 2
            s = s & CInt(lstList(i) + 0.4) & ","
            
        Next
        s = s & CInt(lstList(lstList.Count - 1) + 0.4)
        CommaSeparateList = s
    End Function

    Protected Sub HideAllPanels()
        'pnlProductList.Visible = False
        'pnlMovementList.Visible = False
    End Sub
    
    Private Function DisplayDate(ByVal dtDate As Date) As String
        Dim nMonth As Integer = dtDate.Month
        Dim nYear As Integer = dtDate.Year
        DisplayDate = gsMonthNames(nMonth) & " " & nYear.ToString.Substring(2, 2)
    End Function
    
</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>Product History Report</title>
    <link rel="stylesheet" type="text/css" href="../css/sprint.css" />
</head>
<body>
    <form id="Form1" runat="server">
        <asp:Panel id="pnlProductList" runat="server" Width="100%">
            <table width="100%">
                <tr>
                    <td valign="Bottom" width="5%" style="height: 45px"></td>
                    <td Wrap="False" width="50%" style="height: 45px">
                        <asp:Label ID="Label1" runat="server" forecolor="Silver" font-size="Small" 
                            font-bold="True" font-names="Verdana">Pallet Usage Report</asp:Label>
                        <br /><br />
                    </td>
                    <td Wrap="False" align="Right" width="45%" style="height: 45px"></td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td wrap="False">
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td></td>
                    <td Wrap="False" colspan="2">
                        <asp:Image ID="imgChart" runat="server" />
                        <asp:Label ID="lblNoInformationAvailable" runat="server" 
                            Text="Sorry, no pallet usage data is available"></asp:Label>
                        </td>
                </tr>
                <tr>
                    <td>
                        &nbsp;</td>
                    <td colspan="2" Wrap="False">
                        <asp:Label ID="lblReportGeneratedDateTime" runat="server" 
                            font-names="Verdana,Sans-Serif" font-size="XX-Small" forecolor="Green"></asp:Label>
                        <br />
                        <asp:Label ID="lblError" runat="server" font-names=",Sans-Serif" 
                            font-size="XX-Small" forecolor="red" />
                    </td>
                </tr>
            </table>
        </asp:Panel>
    </form>
</body>
</html>