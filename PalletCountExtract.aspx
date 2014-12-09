<%@ Page Language="VB" AutoEventWireup="true" CodeFile="PalletCountExtract.aspx.vb" Inherits="PalletCountExtract" Theme="AIMSDefault" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Monthly Pallet Count Extraction</title>
    <style type="text/css">
        .stylered
        {
            color: red;
            font-weight: bold;
        }
        .stylepurple
        {
            color: #F13CFC;
            font-weight: bold;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
    <main:header id="ctlHeader" runat="server"></main:header>
    <table style="width: 100%" cellpadding="0" cellspacing="0">
        <tr class="bar_reports">
            <td style="width: 50%; white-space: nowrap">
            </td>
            <td style="width: 50%; white-space: nowrap" align="right">
            </td>
        </tr>
    </table>
    &nbsp;
    <asp:Label ID="lblMessage8" runat="server" Font-Names="Verdana" Font-Size="Small"
        Font-Bold="True">Extract monthly pallet count</asp:Label>
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    <asp:Button ID="btnReadExcelFile" runat="server" Text="Read Spreadsheet" Enabled="False"
        Width="150px" />
    &nbsp;
    <asp:Button ID="btnSaveData" runat="server" Text="Save Data" Enabled="False" Width="150px" />
    &nbsp;&nbsp;
    <asp:Label ID="lblSpreadsheetLocation" runat="server" Font-Names="Verdana"
                        Font-Size="XX-Small" >Ensure spreadsheet is in <b>\\SPRINT_DATA2\PalletCountReport</b></asp:Label>
    <br />
    &nbsp;<table>
        <tr>
            <td style="width: 32%" valign="top">
                <asp:GridView ID="gvRawData" runat="server" AutoGenerateColumns="False" CellPadding="2"
                    Font-Names="Verdana" Font-Size="XX-Small">
                    <Columns>
                        <asp:BoundField DataField="Account" HeaderText="Acct in spreadsheet" />
                        <asp:BoundField DataField="Total" HeaderText="Total" />
                        <asp:BoundField DataField="CustomerAccountCode" HeaderText="Matched to AIMS Acct"
                            ReadOnly="True" />
                    </Columns>
                </asp:GridView>
                <br />
                <asp:Panel ID="pnlDataIntegrityChecks" BorderWidth="3" BorderColor="White" Font-Names="Verdana"
                    Font-Size="XX-Small" Width="100%" runat="server" Visible="False">
                    <b>DATA INTEGRITY CHECKS</b><br />
                    <br />
                    1.&nbsp; Nothing in the <b>Ignore spreadsheet cells</b> list is causing data to
                    be omitted that should be included.<br />
                    <br />
                    2.&nbsp; All customer names read from the spreadsheet are mapped to the correct
                    METACS / AIMS customer code.<br />
                    <br />
                    3.&nbsp; No customers are marked as <span class="stylered">not matched!</span> (the
                    <b>Save Data</b> button is disabled until all cells are successfully matched).<br />
                    <br />
                    4.&nbsp; Any customers marked as <span class="stylepurple">UNPROCESSED</span> are
                    done so intentionally.<br />
                    <br />
                    5.&nbsp; The date has been deduced from the filename correctly.<br />
                    <br />
                    If all these checks pass, click <b>Save Data</b> at the top of the page.<br />
                </asp:Panel>
            </td>
            <td style="width: 1%">
            </td>
            <td style="width: 67%" valign="top">
                <br />
                <fieldset id="Fieldset5">
                    <legend id="Legend4">
                        <asp:Label ID="Label5" runat="server" Text="Saved data report" Font-Names="Verdana" Font-Bold="true"
                            Font-Size="XX-Small" />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:LinkButton ID="lnkbtnRefreshSavedDataReport" runat="server" Font-Names="Verdana" Font-Size="XX-Small">refresh</asp:LinkButton>
                    </legend>
                    <br />
                    &nbsp;<asp:Label ID="lblMostRecentSavedData" runat="server" Font-Names="Verdana" Font-Size="XX-Small" />
                    <br />
                    <br />
                    &nbsp;<asp:Label ID="lblSavedData" runat="server" Font-Names="Verdana" Font-Size="XX-Small" />
                    <br />
                    <br />
&nbsp;<asp:Label ID="lblCustomers" runat="server" Font-Names="Verdana" Font-Size="XX-Small" />
                    <br />
                    <br />
                </fieldset>
                <br />
                <fieldset id="Fieldset3">
                    <legend id="Legend2">
                        <asp:Label ID="Label3" runat="server" Text="File to process" Font-Names="Verdana" Font-Size="XX-Small" Font-Bold="true" />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <asp:LinkButton ID="lnkbtnRecheckFile" runat="server" Font-Names="Verdana" Font-Size="XX-Small">refresh</asp:LinkButton>
                    </legend>&nbsp;<br />
&nbsp;<asp:Label ID="lblFileToProcess" runat="server" Font-Names="Verdana"
                        Font-Size="XX-Small">No file found!</asp:Label>
                    
                    <br />
                    <br />
                </fieldset>
                <br />
                <fieldset id="Fieldset2">
                    <legend id="Legend1">
                        <asp:Label ID="Label2" runat="server" Text="Report date" Font-Names="Verdana" Font-Size="XX-Small" Font-Bold="true" />
                    </legend>
                    <br />
                    &nbsp;<asp:Label ID="lblLegendThisSpreadsheetContainsDataFor" runat="server" Font-Names="Verdana"
                        Font-Size="XX-Small">This spreadsheet contains data for: </asp:Label>
                    <asp:Label ID="lblDate" runat="server" Font-Names="Verdana" Font-Size="X-Small" Font-Bold="True"></asp:Label>
                    <br />
                    <br />
                    <asp:CheckBox ID="cbSetDateManually" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="That's wrong!" AutoPostBack="True" />
                    &nbsp;&nbsp;<asp:DropDownList ID="ddlYear" runat="server" Font-Names="Verdana" Font-Size="XX-Small">
                        <asp:ListItem Selected="True" Value="0">- select year -</asp:ListItem>
                        <asp:ListItem>2009</asp:ListItem>
                        <asp:ListItem>2010</asp:ListItem>
                        <asp:ListItem>2011</asp:ListItem>
                        <asp:ListItem>2012</asp:ListItem>
                        <asp:ListItem>2013</asp:ListItem>
                        <asp:ListItem>2014</asp:ListItem>
                        <asp:ListItem>2015</asp:ListItem>
                        <asp:ListItem>2016</asp:ListItem>
                        <asp:ListItem>2017</asp:ListItem>
                        <asp:ListItem>2018</asp:ListItem>
                        <asp:ListItem>2019</asp:ListItem>
                        <asp:ListItem>2020</asp:ListItem>
                    </asp:DropDownList>
                    &nbsp;<asp:DropDownList ID="ddlMonth" runat="server" Font-Names="Verdana" Font-Size="XX-Small">
                        <asp:ListItem Selected="True" Value="0">- select month -</asp:ListItem>
                        <asp:ListItem Value="1">January</asp:ListItem>
                        <asp:ListItem Value="2">February</asp:ListItem>
                        <asp:ListItem Value="3">March</asp:ListItem>
                        <asp:ListItem Value="4">April</asp:ListItem>
                        <asp:ListItem Value="5">May</asp:ListItem>
                        <asp:ListItem Value="6">June</asp:ListItem>
                        <asp:ListItem Value="7">July</asp:ListItem>
                        <asp:ListItem Value="8">August</asp:ListItem>
                        <asp:ListItem Value="9">September</asp:ListItem>
                        <asp:ListItem Value="10">October</asp:ListItem>
                        <asp:ListItem Value="11">November</asp:ListItem>
                        <asp:ListItem Value="12">December</asp:ListItem>
                    </asp:DropDownList>
                    <br />
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                    <asp:Button ID="btnSaveDate" runat="server" Text="Save" Width="100px" />
                    <br />
                </fieldset>
                <br />
                <fieldset id="Fieldset1">
                    <legend id="Ignore spreadsheet cells">
                        <asp:Label ID="Label1" runat="server" Text="Ignore spreadsheet cells" Font-Names="Verdana" Font-Bold="true"
                            Font-Size="XX-Small" />
                    </legend>
                    <br />
&nbsp;<asp:Label ID="lblMessage5" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Width="400px">Data is extracted from the worksheet <b>Display Report</b>. Cells in this worksheet will <b>not</b> be extracted if they contain text in this list.</asp:Label>
                    <br />
                    <br />
                    &nbsp;<asp:Label ID="lblMessage4" runat="server" Font-Names="Verdana" Font-Size="XX-Small">Ignore cells containing the text:</asp:Label>
                    &nbsp;<asp:TextBox ID="tbIgnoreText" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Width="250px" />
                    <br />
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                    <asp:Button ID="btnAddIgnoreText" runat="server" Text="Add" Width="100px" />
                    <br />
                    <asp:GridView ID="gvIgnoreText" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        CellPadding="2" AutoGenerateColumns="False">
                        <Columns>
                            <asp:TemplateField>
                                <ItemTemplate>
                                    <asp:LinkButton ID="lnkbtnRemoveIgnoreText" runat="server" Font-Names="Verdana" CommandArgument='<%# DataBinder.Eval(Container.DataItem,"id") %>'
                                        Font-Size="XX-Small" OnClick="lnkbtnRemoveIgnoreText_Click">remove</asp:LinkButton>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:BoundField DataField="IgnoreString" HeaderText="Ignore Text" ReadOnly="True">
                                <ItemStyle Width="400px" />
                            </asp:BoundField>
                        </Columns>
                        <EmptyDataTemplate>
                            <asp:Label ID="lblLegendNoData" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                Text="No data" />
                        </EmptyDataTemplate>
                    </asp:GridView>
                    <br />
                </fieldset>
                <br />
                <fieldset id="fsMain">
                    <legend id="fslgndMain">
                        <asp:Label ID="lblLegend01" runat="server" Text="Account identification" Font-Names="Verdana" Font-Bold="true"
                            Font-Size="XX-Small" />
                    </legend>
                    <br />
&nbsp;<asp:Label ID="lblMessage6" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Width="400px">Once the <b>Account Name</b> and <b>Total</b> columns have been extracted from the worksheet, the account text is matched to an account on AIMS using the matching rules in this list.</asp:Label>
                    <br />
                    <br />
                    &nbsp;<asp:Label ID="lblMessage0" runat="server" Font-Names="Verdana" Font-Size="XX-Small">Associate the account text:</asp:Label>
                    &nbsp;<asp:TextBox ID="tbAccountString" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Width="250px" />
                    <br />
                    &nbsp;<asp:Label ID="lblMessage02" runat="server" Font-Names="Verdana" Font-Size="XX-Small">...with the AIMS account:</asp:Label>
                    &nbsp;<asp:DropDownList ID="ddlAimsAccount" runat="server" Font-Names="Verdana" Font-Size="XX-Small">
                    </asp:DropDownList>
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                    <asp:Button ID="btnAddAccountMatchingText" runat="server" Text="Add" Width="100px" />
                    <br />
                    <br />
                    <asp:GridView ID="gvMapping" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        CellPadding="2" AutoGenerateColumns="False">
                        <Columns>
                            <asp:TemplateField>
                                <ItemTemplate>
                                    <asp:LinkButton ID="lnkbtnRemoveMapping" runat="server" Font-Names="Verdana" CommandArgument='<%# DataBinder.Eval(Container.DataItem,"id") %>'
                                        Font-Size="XX-Small" OnClick="lnkbtnRemoveMapping_Click">remove</asp:LinkButton>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:BoundField DataField="ReportString" HeaderText="Match Text" ReadOnly="True">
                                <ItemStyle Width="300px" />
                            </asp:BoundField>
                            <asp:BoundField DataField="ClientName" HeaderText="To Customer">
                                <ItemStyle Width="100px" />
                            </asp:BoundField>
                        </Columns>
                        <EmptyDataTemplate>
                            <asp:Label ID="lblLegendNoData" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                Text="No data" />
                        </EmptyDataTemplate>
                    </asp:GridView>
                    <br />
                    <asp:Panel ID="pnlNotes" BorderWidth="3" BorderColor="White" Font-Names="Verdana"
                        Font-Size="XX-Small" Width="100%" runat="server">
                        NOTES<br />
                        <br />
                        Longer strings are compared first. For example if both <b>Atkins HR</b> and <b>Atkins</b>
                        are defined, <b>Atkins HR</b> takes precedence.<br />
                        <br />
                        Comparison is <b>case insensitive</b>.<br />
                        <br />
                        To skip processing / saving for an entry, associate it with the pseudo customer
                        code <span class="stylepurple">UNPROCESSED</span>, at the bottom of the dropdown
                        list.<br />
                        <br />
                    </asp:Panel>
                </fieldset>
                <br />
                <fieldset id="Fieldset4">
                    <legend id="Legend3">
                        <asp:Label ID="Label4" runat="server" Text="Instructions" Font-Names="Verdana" Font-Size="XX-Small" Font-Bold="true" />
                    </legend>
                    <asp:Panel ID="pnlInstructions" BorderWidth="3" BorderColor="White" Font-Names="Verdana"
                        Font-Size="XX-Small" Width="100%" runat="server">
                        <br />
                        1.&nbsp; Copy the spreadsheet containing the pallet count data to<b> 
                        \\SPRINT_DATA2\PalletCountReport</b>.<br />
                        <br />
                        2.&nbsp; Run this utility.<br />
                        <br />
                        The Saved Data Report above tells you when the last extraction was done, and what
                        data is already in the database.<br />
                        <br />
                        3.&nbsp; Check that the correct file has been located (the filename must end in
                        .xls or .xlsx), and that date has been identified correctly from the filename (the
                        program looks for &lt;3 letter month&gt;&lt;2 digit year&gt;. If it hasn&#39;t,
                        click the check box marked <b>That&#39;s wrong</b>, enter the correct date, then
                        click <b>Save</b>.<br />
                        <br />
                        4.&nbsp; Click <b>Read Spreadsheet</b>.&nbsp; The utility finds the spreadsheet
                        and reads the workbook <b>Display Report</b>.<br />
                        <br />
                        5.&nbsp; Check any cells that do not contain customer data are ignored. Adjust the
                        ignore text if necessary and re-read the spreadsheet.<br />
                        <br />
                        6.&nbsp; Check all Customer names have been correctly identified and mapped to their
                        equivalent METACS/AIMS abbreviation. Add/change/remove name mappings if necessary
                        and re-read the spreadsheet.<br />
                        <br />
                        7.&nbsp; Click <b>Save Data</b>.&nbsp; The data is written to the database, after
                        which <b>the spreadsheet is deleted</b>. An archive is saved in 
                        <b>\\SPRINT_DATA2\PalletCountReport\backup</b>. To retrieve a backup, navigate 
                        to this directory, identify your file (extra characters whill hve been added to 
                        the filename to ensure it is unique within the backup directory) then copy and 
                        rename it.
                        <br />
                        
                        <br />
                        <i>NOTES<br />
                        </i>
                        <br />
                        If you save data for a month that you have already previously saved, the previously
                        saved data is overwritten.<br />
                    </asp:Panel>
                </fieldset>
            </td>
        </tr>
    </table>
    &nbsp;<asp:Label ID="lblMessage" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Font-Bold="True" ForeColor="Red"/>
    </form>
</body>
</html>
