<%Option Explicit%>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<html>
<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>ePlaS - Integrated Plantation System</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js"></script>
<link rel="StyleSheet" href="style-content.css" type="text/css" />
<link rel="StyleSheet" href="dtree.css" type="text/css" />
<script type="text/javascript" src="dtree.js"></script>

</head>
<body>
<%if Session("USERNAME")="" then%>
	      <!--#include file="toplogin.jsp"-->
<%else%>
	     <!--#include file="top.jsp"-->
<%end if%>


 <% if Session("fcBa") <> "" then %>
 <div style=" width:21.7%; float:left;margin-left:0.3%; margin-top:0.3%; margin-bottom:0.3%;">
     <!--#include file="menu.asp"-->
   </div>
   <div class="judul_head"  valign="top" bgcolor="#fff" style=" width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;"> 
 <div class="judul">
	   iPlaS ( integrated Plantation System )
	</div>
   </div>
 <% end if %>
 
     <div class="box_both" style="overflow:hidden;" > 
            <div>
            </br> </br>
            <img src="MSGBOD.png" border="0">
            </br> </br>
            </div>
     </div>
  <!--#include file="bottom.jsp"-->

</body>
</html>