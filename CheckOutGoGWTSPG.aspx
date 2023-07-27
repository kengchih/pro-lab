<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CheckOutGoGWTSPG.aspx.vb" Inherits="PP01SSV2.CheckOutGoGWTSPG" %>

<%@ Register TagPrefix="wc" Namespace="Wellan.Web.UI.WebControls" Assembly="Wellan.Web" %>
<!DOCTYPE html>
<html lang="zh-Hant">
<head id="Head1" runat="server">
    <meta content="text/html; charset=utf-8" />
    <meta name="robots" content="noindex">
    <meta name="googlebot" content="noindex">
    <title><%Wellan.Web.Application.ResourceString("RPT_Title")%></title>
    <script type="text/javascript" src="https://code.jquery.com/jquery-2.2.4.min.js"></script>
    <script src="https://code.jquery.com/jquery-migrate-1.4.1.js"></script>
    <script type="text/javascript" src="../js/default.js"></script>
</head>
<body oncontextmenu="norightclick();" oncopy="nocopy();" leftmargin="0" topmargin="0">
    <wc:WUserLog ID="WUserLog1" runat="server" FunctionID="StoreFront_CheckOutGoGWP">
    </wc:WUserLog>
    <div id="MsgWait" style="z-index: 1; left: 300px; width: 200px; position: absolute; top: 200px; height: 100px; text-align: center;">
        <img alt="Wait" style="vertical-align:middle;" src="../Images/Product/PPwait.gif"  />
    </div>
    <div>
    </div>
</body>
</html>
