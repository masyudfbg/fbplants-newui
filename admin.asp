<%Option Explicit%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="include/db_param.asp"-->
<html>
<head>
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>iPS - Login</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js"></script>
</head>
<body onLoad="document.frmLogin.txtUser.focus()">
<!--#include file="top.jsp"-->
<table width="100%">
<tr>
 <td width="22%" valign="top">
 <!--#include file="menu.asp"-->
 </td>
 <td width="78%" valign="top">
  <%
  '--ATAS ISI--
  Dim txtUser,txtBA,txtPassword,q
  txtUser=Request("txtUser")
  txtBA=Request("txtBA")
  txtPassword=Request("txtPassword")

	Session("fcBa")= ""
	Session("Crop")=""
	Session("BA") = ""
	Session("COMPANY")=""
	Session("FCMANAGER") = ""
	Session("KTU") = ""
	Session("Estate_Mill") = ""
	Session("v_Bank")=""
	Session("v_Purchase")=""
	Session("v_Stores")=""
	Session("v_Employee")=""
	Session("v_Workshop")=""
	Session("v_Vehicle")=""
	Session("v_Contract")=""
	Session("v_Nursery")=""
	Session("v_Project")=""
	Session("v_Building")=""
	Session("v_General_Charges")=""
	Session("v_Costbook")=""
	Session("v_Harvesting")=""
	Session("v_WeighBridge")=""
	Session("v_Mill_Production")=""
	Session("v_Delivery")=""
    Session("USERNAME") = ""
	Session("CodeName") = ""

	Session("cls_Bank")= ""
	Session("cls_Purchase")= ""
	Session("cls_Stores")= ""
	Session("cls_Employee")= ""
	Session("cls_Workshop")= ""
	Session("cls_Vehicle")= ""
	Session("cls_Contract")= ""
	Session("cls_Nursery")= ""
	Session("cls_Project")= ""
	Session("cls_Building")= ""
	Session("cls_General_Charges")= ""
	Session("cls_Costbook")= ""
	Session("cls_Harvesting")= ""
	Session("cls_WeighBridge")= ""
	Session("cls_Mill_Production")= ""
	Session("cls_Delivery")= ""
	
	Session("NextMonth")
	 
 IF Request("cmdok")="cmdok" and txtuser<>"" and txtpassword<>"" then
  ' Perbaikan keamanan: username/password/BA dikirim sebagai parameter (dulu disambung ke SQL)
  q="Select fccode,FULLNAME from USERPROFILE where fccode = ? AND fcPassword = ? and fcBA = ?"
  'q="Select fccode,FULLNAME from USERPROFILE where fccode='"& txtUser & "' and fcBA='" & txtBA & "'"
  Set rs = DbQuery(q, Array(txtUser, encrypt(txtPassword,1), txtBA))
  if not rs.eof then
    Session("USERNAME") = RS("FCCODE")
	Session("CodeName") = RS("fccode")
	q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode = ?"
 

    set rsCat = DbQuery(q, Array(txtBA))
	if not rsCat.eof then
	
	   Session("fcBa")= rsCat("fcCode")
	   Session("BA") = rsCat("fcName")
	   Session("COMPANYCODE")=rsCat("fccompanycode")	   
	   Session("COMPANY")=rsCat("FCCOMPANYNAME")
	   Session("FCMANAGER") = rsCat("FCMANAGER")
	   Session("KTU") = rsCat("FCKTU")
	   Session("Crop")=rsCat("FCCROP")
	   Session("Estate_Mill")=rsCat("fcType")
	   
	   Session("v_Bank")=cmdFindPeriod("Bank")
	   Session("v_Purchase")=cmdFindPeriod("Purchase")
	   Session("v_Stores")=cmdFindPeriod("Stores")
	   Session("v_Employee")=cmdFindPeriod("Employee")
	   Session("v_Workshop")=cmdFindPeriod("Workshop")
	   Session("v_Vehicle")=cmdFindPeriod("Vehicle")
	   Session("v_Contract")=cmdFindPeriod("Contract")
	   Session("v_Nursery")=cmdFindPeriod("Nursery")
	   Session("v_Project")=cmdFindPeriod("Project")
	   Session("v_Building")=cmdFindPeriod("Building")
	   Session("v_General_Charges")=cmdFindPeriod("General_Charges")
	   Session("v_Costbook")=cmdFindPeriod("Costbook")
	   Session("v_Harvesting")=cmdFindPeriod("Harvesting")
	   Session("v_WeighBridge")=cmdFindPeriod("WeighBridge") 
	   Session("v_Mill_Production")=cmdFindPeriod("Mill_Production") 
	   Session("v_Delivery")=cmdFindPeriod("Delivery") 

	   Session("cls_Bank")=cmdPeriodClose("Bank")
	   Session("cls_Purchase")=cmdPeriodClose("Purchase")
	   Session("cls_Stores")=cmdPeriodClose("Stores")
	   Session("cls_Employee")=cmdPeriodClose("Employee")
	   Session("cls_Workshop")=cmdPeriodClose("Workshop")
	   Session("cls_Vehicle")=cmdPeriodClose("Vehicle")
	   Session("cls_Contract")=cmdPeriodClose("Contract")
	   Session("cls_Nursery")=cmdPeriodClose("Nursery")
	   Session("cls_Project")=cmdPeriodClose("Project")
	   Session("cls_Building")=cmdPeriodClose("Building")
	   Session("cls_General_Charges")=cmdPeriodClose("General_Charges")
	   Session("cls_Costbook")=cmdPeriodClose("Costbook")
	   Session("cls_Harvesting")=cmdPeriodClose("Harvesting")
	   Session("cls_WeighBridge")=cmdPeriodClose("WeighBridge") 
	   Session("cls_Mill_Production")=cmdPeriodClose("Mill_Production") 
	   Session("cls_Delivery")=cmdPeriodClose("Delivery") 

       Session("NextMonth") = cmdPeriodClose("Financial")
	   
	   if request("persistent") = "y" Then
		  response.Cookies("myCookies") = ""   ' kosongkan dulu: ASP ikut menyalin sub-key lama (termasuk password) dari cookie request
		  response.Cookies("myCookies")("txtUser") = trim(request("txtUser"))
		  response.Cookies("myCookies")("txtBA") = trim(request("txtBA"))
		  ' Keamanan: password TIDAK disimpan di cookie (dulu password asli tersimpan 365 hari)
		  response.Cookies("myCookies").Expires = Now() + 365
	   End if
	   Response.Redirect("login.asp")
	end if
  End if
  rs.close()
  rs=null  
  End if

  txtBA=Request.Cookies("myCookies")("txtBA")
  txtUser = Request.Cookies("myCookies")("txtUser")
  txtPassword = ""   ' password tidak lagi diisi otomatis dari cookie
  %>
  <form action="logout.asp" method="post" name="frmLogin">
  
      <img src="ICON/login.bmp">
	  <table>
		  <tr>
		     <td colspan="2" align="left"><font face="Arial" size="-1">Login Here : <br></td>
		  </tr>
		  <tr>
		     <td align="left"><font face="Arial" size="-1">User Name </td>
			 <td align="left"><input type="text" name="txtUser" maxlength="20" value="<%=txtUser%>"></td>
		  </tr>
		  <tr>
		     <td align="left"><font face="Arial" size="-1">Password</td><td align="left">
		     <input type="password"  name="txtPassword" maxlength="20" value="<%=txtPassword%>">
			 </td>
		  </tr>

		  <tr>
		  <input type="hidden" value="KSB" name="txtBA">
		  <td colspan="2">
		  <input type="checkbox" name="persistent" value="y"><font face="Arial" size="-3">Remember my ID on this computer</font>
		  </td>
		  </tr>
		  <tr>
		  <td colspan="2" align="center"><input type="submit" value="Login" style="background-color:#CCCCCC" style="border:#0000FF">
		  </td>
		  </tr>
	  </table>
	  <br>
		<font  face="verdana, arial"  size="-3">The User Name & password is case sensitive<br>
		<span style="color: #0000CC">Requirement : 
		</span>Explorer 6.0 or Higher<br>
		<span style="color: #0033FF">Best View</span> : 1024x768<br><font color="#009966"><b>
		</b></font>
	  
	  <input type="hidden" value="cmdok" name="cmdok">
  </form>

  <%
  '--BAWAH ISI--
  %>
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