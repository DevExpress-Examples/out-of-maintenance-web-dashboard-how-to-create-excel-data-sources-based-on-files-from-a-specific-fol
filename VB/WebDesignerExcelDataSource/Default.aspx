<%@ Page Language="vb" AutoEventWireup="true" CodeBehind="Default.aspx.vb" Inherits="WebDesignerExcelDataSource.Default" %>

<%@ Register Assembly="DevExpress.Dashboard.v17.2.Web, Version=17.2.5.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.DashboardWeb" TagPrefix="dx" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div style="position: absolute; top: 0; left: 0; right: 0; bottom: 0;">
            <dx:ASPxDashboard ID="ASPxDashboard1" runat="server" Width="100%" Height="100%">
            </dx:ASPxDashboard>
        </div>
    </form>
</body>
</html>