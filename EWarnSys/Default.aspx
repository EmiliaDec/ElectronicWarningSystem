<%@ Page Title="主页" Language="C#" MasterPageFile="~/Site.master" AutoEventWireup="true"
    CodeBehind="Default.aspx.cs" Inherits="warnSysforElec._Default" %>

<asp:Content ID="HeaderContent" runat="server" ContentPlaceHolderID="HeadContent">
</asp:Content>
<asp:Content ID="BodyContent" runat="server" ContentPlaceHolderID="MainContent">
<script type="text/javascript" src="http://ajax.microsoft.com/ajax/jquery/jquery-1.7.min.js";;;></script>
<script src='http://code.jquery.com/jquery-latest.min.js' type='text/javascript'></script>

    <p>
        <asp:FileUpload ID="FileUpload1" runat="server" />
        <asp:Button ID="Button_OK" runat="server" Text="确定" OnClick="Button_OK_Click" />
        <asp:Button  ID="Button1" runat="server" Text="实时电流检测" OnClick="ButtonTimelyICheck"/>
        <asp:Button  ID="Button4" runat="server" Text="实时线损检测" OnClick="ButtonTimelyWWCheck"/>
        <br/><br/>明天的温度：
        <asp:TextBox ID="TB_temprature" runat="server"></asp:TextBox>
        <asp:Button  ID="Button2" runat="server" Text="跨天电流预测" OnClick="ButtonTomorrowPredict_I"/>
        <asp:Button  ID="Button3" runat="server" Text="跨天有功预测" OnClick="ButtonTomorrowPredict_W"/>
    </p>
    <p>
        <asp:Label ID="Lable_I" Text="电流阈值：405" runat="server"></asp:Label>
        <asp:Label ID="Label_W" Text="有功功率阈值：6" runat="server"></asp:Label>
    </p>
    <asp:GridView ID="gridview_W" runat="server"></asp:GridView>
</asp:Content>
