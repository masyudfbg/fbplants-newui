<%
Response.Buffer = True
%>
<HTML>
<HEAD>
	<TITLE>iPS</TITLE>
</HEAD>
<SCRIPT LANGUAGE="JavaScript">
// This script is intended for use with a minimum of Netscape 4 or IE 4.
if(document.getElementById) {
	var upLevel = true;
	}
else if(document.layers) {
	var ns4 = true;
	}
else if(document.all) {
	var ie4 = true;
	}

function showObject(obj) {
	if (ns4) obj.visibility = "show";
	else if (ie4 || upLevel) obj.style.visibility = "visible";
	}
function hideObject(obj) {
	if (ns4) {
		obj.visibility = "hide";
		}
	if (ie4 || upLevel) {
		obj.style.visibility = "hidden";
		}
	}

</SCRIPT>
<BODY BGCOLOR="#FFFFFF">
		<DIV ID="splashScreen" STYLE="position:absolute;z-index:5;top:30%;left:35%;">
			<TABLE BGCOLOR="#000000" BORDER=1 BORDERCOLOR="#000000" CELLPADDING=0 CELLSPACING=0 HEIGHT=200 WIDTH=300>
				<TR>
					<TD WIDTH="100%" HEIGHT="100%" BGCOLOR="#CCCCCC" ALIGN="CENTER" VALIGN="MIDDLE">
						<BR><BR> &nbsp; &nbsp; 
						<FONT FACE="Helvetica,Verdana,Arial" SIZE=3 COLOR="#000066"><B>Page Loading.  Please wait...</B></FONT> 
						&nbsp; &nbsp; <BR> 
						<IMG SRC="wait.gif" BORDER=1 WIDTH=75 HEIGHT=15><BR><BR> 
					</TD>
				</TR>
			</TABLE>
		</DIV>
<%Response.Flush%>


All processing and content here...


<%Response.Flush%>
<SCRIPT LANGUAGE="JavaScript">
if(upLevel) {
	var splash = document.getElementById("splashScreen");
	}
else if(ns4) {
	var splash = document.splashScreen;
	}
else if(ie4) {
	var splash = document.all.splashScreen;
	}
hideObject(splash);
</SCRIPT>
</BODY>
</HTML>
