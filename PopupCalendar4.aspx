<%@ Page Language="VB" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsPostBack Then
            lblToday.Text = Now.ToString("dddd, dd-MMM-yy")
            lnkbtnYear1.Text = Date.Today.Year - 4
            lnkbtnYear2.Text = Date.Today.Year - 3
            lnkbtnYear3.Text = Date.Today.Year - 2
            lnkbtnYear4.Text = Date.Today.Year - 1
            lnkbtnYear5.Text = Date.Today.Year
            lnkbtnYear6.Text = Date.Today.Year + 1
            lnkbtnYear7.Text = Date.Today.Year + 2
            lnkbtnYear8.Text = Date.Today.Year + 3
            lnkbtnYear9.Text = Date.Today.Year + 4
        End If
        hidTextboxName.Value = Request.QueryString("textbox").ToString()
        Try
            hidMode.Value = Request.QueryString("mode").ToString() ' "replace" or "append - replace if blank"
        Catch
            hidMode.Value = String.Empty
        End Try
        Try
            hidPrefixText.Value = Request.QueryString("prefixtext").ToString() ' text to place before date
        Catch
            hidPrefixText.Value = String.Empty
        End Try
    End Sub

    Protected Sub cal_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call BuildScript()
    End Sub
    
    Protected Sub BuildScript()
        Dim strScript As String = "<script>"
        If hidMode.Value = "append" Then
            strScript += "window.opener.document.forms(0)." + hidTextboxName.Value + ".value = window.opener.document.forms(0)." + hidTextboxName.Value + ".value + '" & hidPrefixText.Value & "' + '"
        Else
            strScript += "window.opener.document.forms(0)." + hidTextboxName.Value + ".value = '"
        End If
        strScript += cal.SelectedDate.ToString("dd-MMM-yyyy")
        strScript += "';self.close()"
        strScript += "</" + "script>"
        'RegisterClientScriptBlock("anything", strScript)
        ClientScript.RegisterClientScriptBlock(GetType(Page), "noname", strScript)
    End Sub

    Protected Sub ChangeMonth(ByVal sMonth As String)
        cal.TodaysDate = Date.Parse(Day(cal.TodaysDate) & sMonth & Year(cal.TodaysDate))
    End Sub

    Protected Sub ChangeYear(ByVal sYear As String)
        cal.TodaysDate = Date.Parse(Day(cal.TodaysDate) & Format(cal.TodaysDate, "MMM") & sYear)
    End Sub

    Protected Sub lnkbtnMonth_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim lb As LinkButton = sender
        Call ChangeMonth(lb.Text)
    End Sub

    Protected Sub lnkbtnYear_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim lb As LinkButton = sender
        Call ChangeYear(lb.Text)
    End Sub

    Protected Sub lnkbtnToday_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        cal.SelectedDate = Today
        Call BuildScript()
    End Sub
</script>

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Select Date</title>
    <LINK rel="stylesheet" type="text/css" href="Reports.css" />
</head>
<body>
    <form id="form1" runat="server">
    <div style="font-size: x-small; font-family: Verdana, Sans-Serif">
        <p align="center">
            <asp:LinkButton ID="lnkbtnToday" runat="server" OnClick="lnkbtnToday_Click">Today</asp:LinkButton>
            is
            <asp:Label ID="lblToday" runat="server"></asp:Label><br />
            Select a date by choosing a <b>month</b> & <b>year</b>, then clicking on a <b>day</b></p>
        <table style="width: 100%">
            <tr>
                <td align="center" valign=top >
        <asp:Calendar ID="cal" runat="server" BackColor="White" BorderColor="#999999"
            CellPadding="4" DayNameFormat="Shortest" Font-Names="Verdana" Font-Size="8pt"
            ForeColor="Black" Height="180px" Width="200px" OnSelectionChanged="cal_SelectionChanged">
            <SelectedDayStyle BackColor="#666666" Font-Bold="True" ForeColor="White" />
            <TodayDayStyle BackColor="#CCCCCC" ForeColor="Black" />
            <SelectorStyle BackColor="#CCCCCC" />
            <WeekendDayStyle BackColor="#FFFFCC" />
            <OtherMonthDayStyle ForeColor="#808080" />
            <NextPrevStyle VerticalAlign="Bottom" />
            <DayHeaderStyle BackColor="#CCCCCC" Font-Bold="True" Font-Size="7pt" />
            <TitleStyle BackColor="#999999" BorderColor="Black" Font-Bold="True" />
        </asp:Calendar>
                </td>
            </tr>
            <tr>
                <td align="center" valign=top nowrap style="font-size: xx-small; font-family: Verdana">
                    <asp:LinkButton ID="lnkbtnJan" runat="server" OnClick="lnkbtnMonth_Click">Jan</asp:LinkButton>
                    <asp:LinkButton ID="lnkbtnFeb" runat="server" OnClick="lnkbtnMonth_Click">Feb</asp:LinkButton>
                    <asp:LinkButton ID="lnkbtnMar" runat="server" OnClick="lnkbtnMonth_Click">Mar</asp:LinkButton>
                    <asp:LinkButton ID="lnkbtnApr" runat="server" OnClick="lnkbtnMonth_Click">Apr</asp:LinkButton>
                    <asp:LinkButton ID="lnkbtnMay" runat="server" OnClick="lnkbtnMonth_Click">May</asp:LinkButton>
                    <asp:LinkButton ID="lnkbtnJun" runat="server" OnClick="lnkbtnMonth_Click">Jun</asp:LinkButton>
                    <asp:LinkButton ID="lnkbtnJul" runat="server" OnClick="lnkbtnMonth_Click">Jul</asp:LinkButton>
                    <asp:LinkButton ID="lnkbtnAug" runat="server" OnClick="lnkbtnMonth_Click">Aug</asp:LinkButton>
                    <asp:LinkButton ID="lnkbtnSep" runat="server" OnClick="lnkbtnMonth_Click">Sep</asp:LinkButton>
                    <asp:LinkButton ID="lnkbtnOct" runat="server" OnClick="lnkbtnMonth_Click">Oct</asp:LinkButton>
                    <asp:LinkButton ID="lnkbtnNov" runat="server" OnClick="lnkbtnMonth_Click">Nov</asp:LinkButton>
                    <asp:LinkButton ID="lnkbtnDec" runat="server" OnClick="lnkbtnMonth_Click">Dec</asp:LinkButton>
                    </td>
            </tr>
            <tr >
                <td align="center" valign="top" style="font-size: xx-small; font-family: Verdana; white-space:nowrap; height: 18px;">
                    <asp:LinkButton ID="lnkbtnYear1" runat="server" OnClick="lnkbtnYear_Click">Year1</asp:LinkButton>
                    <asp:LinkButton ID="lnkbtnYear2" runat="server" OnClick="lnkbtnYear_Click">Year2</asp:LinkButton>
                    <asp:LinkButton ID="lnkbtnYear3" runat="server" OnClick="lnkbtnYear_Click">Year3</asp:LinkButton>
                    <asp:LinkButton ID="lnkbtnYear4" runat="server" OnClick="lnkbtnYear_Click">Year4</asp:LinkButton>
                    <asp:LinkButton ID="lnkbtnYear5" runat="server" OnClick="lnkbtnYear_Click">Year5</asp:LinkButton>
                    <asp:LinkButton ID="lnkbtnYear6" runat="server" OnClick="lnkbtnYear_Click">Year6</asp:LinkButton>
                    <asp:LinkButton ID="lnkbtnYear7" runat="server" OnClick="lnkbtnYear_Click">Year7</asp:LinkButton>
                    <asp:LinkButton ID="lnkbtnYear8" runat="server" OnClick="lnkbtnYear_Click">Year8</asp:LinkButton>
                    <asp:LinkButton ID="lnkbtnYear9" runat="server" OnClick="lnkbtnYear_Click">Year9</asp:LinkButton>
                </td>
            </tr>
        </table>
        <br />
        <asp:HiddenField ID="hidTextboxName" runat="server" />
        <asp:HiddenField ID="hidMode" runat="server" />
        <asp:HiddenField ID="hidPrefixText" runat="server" />
    </div>
    </form>
</body>
</html>
