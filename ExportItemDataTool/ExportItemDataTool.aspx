<%@ Page Language="C#" AutoEventWireup="true"
    Inherits="SharedSource.Verndale.ExportData.ExportItemDataTool" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Export Item Data</title>
    <script type="text/javascript">
        function ClearStatus() {
            var labelObj = document.getElementById("<%= lblStatus.ClientID %>");
            labelObj.innerHTML = "";
            return true;
        }
    </script>
</head>
<body>
    <form id="form1" runat="server" style="background: white; font-family: sans-serif; font-size: 13px">
        <table style="padding: 10px;">
            <tr>
                <td>
                    <asp:Label ID="lblLanguage" runat="server" Text="Language:"></asp:Label>
                </td>
                <td>
                    <asp:TextBox ID="txtLanguage" runat="server" Width="350px" Text="en"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:Label ID="lblIndexName" runat="server" Text="Index:"></asp:Label>
                </td>
                <td>
                    <asp:TextBox ID="txtIndexName" runat="server" Width="350px" Text="sitecore_web_index"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:Label ID="lblLocation" runat="server" Text="Location (Guid):"></asp:Label>
                </td>
                <td>
                    <asp:TextBox ID="txtLocation" runat="server" Width="350px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:Label ID="lblTemplate" runat="server" Text="Template (Comma Separated List of Template IDs):"></asp:Label>
                </td>
                <td>
                    <asp:TextBox ID="txtTemplate" runat="server" Width="350px" TextMode="MultiLine" Height="70px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:Label ID="lblInclude" runat="server" Text="Include:"></asp:Label>
                </td>
                <td>
                    <asp:CheckBoxList ID="cboxListStandardFields" runat="server">
                        <asp:ListItem Value="0" Selected="true">Item Id</asp:ListItem>
                        <asp:ListItem Value="1" Selected="true">Item Name</asp:ListItem>
                        <asp:ListItem Value="2">Template Id</asp:ListItem>
                        <asp:ListItem Value="3">Path</asp:ListItem>
                    </asp:CheckBoxList>
                </td>
            </tr>
            <tr>
                <td style="padding-right: 10px;">
                    <asp:Label ID="lblFields1" runat="server" Text="Field Names (Comma Separated):"></asp:Label>
                    <br />
                    <asp:Label ID="lblFields2" runat="server" Text="If you need the value from a reference field item, use nested brackets."></asp:Label>
                    <br />
                    <asp:Label ID="lblFields3" runat="server" Text="eg: if Article Type, Necessary Products & ProductType were droplinks / multilists, "></asp:Label>
                    <br />
                    <asp:Label ID="lblFields4" runat="server"
                        Text="Headline,Article Type(Title),Necessary Products(Sku,ProductType(Title)),Skill Level(Title)"
                        Style="font-style: italic; color: #0033CC"></asp:Label>
                </td>
                <td>
                    <asp:TextBox ID="txtFieldNames" runat="server" TextMode="MultiLine" Height="70px" Width="350px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:Label ID="lblFileFormat" runat="server" Text="File Format:"></asp:Label>
                </td>
                <td>
                    <asp:RadioButtonList ID="rbtnList" runat="server">
                        <asp:ListItem Value="0" Selected="True">CSV</asp:ListItem>
                        <asp:ListItem Value="1">XML</asp:ListItem>
                        <asp:ListItem Value="2">Json</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td colspan="2" style="text-align: center">
                    <asp:Button ID="btnExport" runat="server" Text="Export" OnClick="btnExport_Click" OnClientClick="return ClearStatus();" />
                </td>
            </tr>
            <tr>
                <td style="height: 15px;"></td>
            </tr>
            <tr>
                <td colspan="2" style="text-align: center; color: red">
                    <asp:Label runat="server" ID="lblStatus"></asp:Label>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
