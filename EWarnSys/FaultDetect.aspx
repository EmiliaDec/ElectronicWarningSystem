<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="FaultDetect.aspx.cs" Inherits="warnSysforElec.FaultDetect" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<form id="abcd" runat="server">
    <asp:fileupload ID="fileuploadB" runat="server"></asp:fileupload> 
    <asp:button runat="server" text="确定" id="Button_Check" 
        onclick="Button_Check_Click" />
    <asp:GridView ID="gridview_C" runat="server"></asp:GridView>
</form>