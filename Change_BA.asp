<%Option Explicit%>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<html>
<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>iPS - Change BA</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js"></script>
</head>
<body>
<!--#include file="top.jsp"-->
<table width="100%">
<tr>
 <td width="22%" valign="top">
 <!--#include file="menu.asp"-->
 </td>
 <td width="78%" valign="top">
  <%' ----------------- ATAS ISI %>
  <%
  Dim txtUser,txtBA,txtPassword,q
  txtUser=Request("txtUser")
  txtBA=Request("txtBA")
  txtPassword=Request("txtPassword")
  
  q="Select fcCode from USERPROFILE where fccode='"+txtUser+"' AND fcPassword='"+txtPassword+"'"
  rs=Server.CreateObject("ADODB.Recordset")
  Set rs = connStr.Execute(q)
  if not rs.eof then
    Session("USERNAME") = RS("FULLNAME")
	q = "Select FCNAME,FCCOMPANYNAME,FCMANAGER from BusinessUnit where fcCode='" & txtBA & "'"
    set rsCat = connStr.Execute(q)
	if not rsCat.eof then
	   Session("BA") = rsCat("fcName")
	   Session("COMPANY")=rsCat("FCCOMPANYNAME")
	   Session("FCMANAGER") = rsCat("FCMANAGER")
	   Response.Redirect("login.asp")
	end if
  End if
  rs.close()
  rs=null%>

  <form action="logout.asp" method="post">
	  <font face="Verdana, Arial, Helvetica, sans-serif" style="font-size:10px">Change BA Here : </font><br>
	  <table class="data">
		  <tr>
		     <td colspan="2">Case Sensitive</td>
		  </tr>
		  <tr>
		     <td align="left">User Name </td>
			 <td align="left"><input type="text" name="txtUser" maxlength="20"></td>
		  </tr>
		  <tr>
		     <td align="left">Password</td><td align="left">
		     <input type="password"  name="txtPassword" maxlength="20" value="<%=Request.Form("txtPassword")%>">
			 </td>
		  </tr>
		  <tr>
		     <td align="left">Business Unit</td>
		     <td align="left">
			 <select name="txtBA">
			 <%
			 q="Select fccode,fcname from BusinessUnit"
			 Set rs = connStr.Execute(q)
			 Do while not rs.EOF%>
			    <option value="<%=rs("fccode")%>"><%=rs("fcname")%></option>
			    <%rs.moveNext()
			 loop%>
			 rs.close()
			 rs=null%>
			</select>
		    </td>
		  </tr>
		  <tr>
		  <td colspan="2"><input type="submit" value="Login" style="background-color:#CCCCCC" style="border:#0000FF">
		  </td>
		  </tr>
	  </table>
	  <input type="hidden" value="cmdok" name="cmdok">
  </form>

  <%'------------------ BAWAH ISI%>
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