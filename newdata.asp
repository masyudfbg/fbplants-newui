<%Option Explicit%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include virtual="txtFunction.inc"-->
<html>
<head>
<title>iPS - New Data</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js"></script>
<script language="JavaScript" src="date-picker.js"></script> 
</head>
<body Onload=TeksJalan>
<!--#include file="top.jsp"-->
<table width="100%">
<tr>
 <td width="22%" valign="top">
 <!--#include file="menu.asp"-->
 </td>
 <td width="78%" valign="top">
 <%Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,fddate,fccode
 fddate   =Request("fddate")
 fccode   =Request("fccode")
 txtTable =Request("txtTable")
 txtModule=Request("txtModule")
 txtSearch=Request("txtSearch")
 txtField =Request("txtField")
 if txtTable="PaymentVoucher" then%>
    <!--#include file="form/PaymentVoucher.asp"-->
 <%end if%>
 </td>
</tr>
<tr>
  <td colspan="2">
  <!--#include file="bottom.jsp"-->
  </td>
</tr>
</table>
</body>
</html>