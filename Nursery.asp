<%Option Explicit%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<html>
<head>
<title>ePlaS - Nursery Batch</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js" ></script>
<link rel="StyleSheet" href="dtree.css" type="text/css" />
<link rel="StyleSheet" href="style-content.css" type="text/css" />
<link rel="stylesheet" href="css/jquery-ui.css" />
<script type="text/javascript" src="dtree.js"></script>
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>
<script language="JavaScript">
xAddEventListener(window, 'load',
  function() {
    new xTableHeaderFixed('xthf-site', window);
    new xTableHeaderFixed('xthf-green', window);
    new xTableHeaderFixed('xthf-blue', window);
  }, false
);
</script>
 <script>
$(function() {
$( ".datepicker" ).datepicker();
});
</script>
<!--<script language="JavaScript" src="/js/date-picker.js"></script>-->
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>
<% if Request("oldui") <> "1" then %>
<!-- Tampilan baru (css/fbp-shell.css + css/fbp-forms.css + css/fbp-docdetail.css). ?oldui=1 = tampilan lama -->
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
<link rel="StyleSheet" href="css/login/font-awesome.min.css" type="text/css" />
<link rel="StyleSheet" href="css/fbp-shell.css?v=20261009d" type="text/css" />
<link rel="StyleSheet" href="css/fbp-forms.css?v=20261009e" type="text/css" />
<link rel="StyleSheet" href="css/fbp-docdetail.css?v=20261009b" type="text/css" />
<script type="text/javascript" src="js/fbp-shell.js?v=20261014a"></script>
<% end if %>
</head>
<body<% if Request("oldui") <> "1" then %> class="dash page-show page-legacy page-doc"<% end if %>>
<!--#include file="top.jsp"-->
	<div style=" width:21.7%; float:left;margin-left:0.3%; margin-top:0.3%; margin-bottom:0.3%;">
			<!--#include file="menu.asp"-->
		</div>
   <div class="judul_head" valign="top" bgcolor="#fff" style=" width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;"> 
 <div class="judul">
	   Module : <%=request("txtModule")%> / <%=request("txtTable")%>
	</div>
 <div class="box_both" style="overflow:scroll;">
 <!-------------awal!------------->
 <%


Dim v_FCENTRY, v_FCEDIT, v_FCIP, v_FCBA
DIM v_NURSERYCODE,v_DATEPLANTED
DIM v_DESCRIPTION,  v_CROP, v_PLANTINGMATERIAL, v_CLONE, v_SOURCE, v_TYPE, v_STATUS, v_StoreCode, v_StockCode, v_Stage
Dim cmdView
v_NURSERYCODE =request("v_NURSERYCODE")
v_DATEPLANTED =request("v_DATEPLANTED")
v_DESCRIPTION=request("v_DESCRIPTION")
v_DATEPLANTED=request("v_DATEPLANTED")
v_CROP=request("v_CROP")
v_PLANTINGMATERIAL=request("v_PLANTINGMATERIAL")
v_CLONE=request("v_CLONE")
v_SOURCE=request("v_SOURCE")
v_TYPE=request("v_TYPE")
v_StoreCode=request("v_StoreCode")
v_StockCode=request("v_StockCode")
v_STATUS=request("v_STATUS")
v_Stage=request("v_Stage")


if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then
	if Request("txtJenis")= "Delete" then
		q	=	" Delete from Nursery where" & _
				" NURSERYCODE = '" & v_NURSERYCODE & "'" & _
				" and fcBA = '" & Session("fcBA") & "'"
				
		'Response.Write q
		connStr.execute(q)
		cmdView	= "True"
	elseif Request("txtJenis")="New" then	
		q	=	" Insert into Nursery (" & _
				" NURSERYCODE" & _
				" , DESCRIPTION" & _
				" , DATEPLANTED" & _
				" , CROP" & _
				" , PLANTINGMATERIAL" & _
				" , CLONE" & _
				" , SOURCE" & _
				" , TYPE" & _
				" , STATUS" & _
				" , FCENTRY" & _
				" , FCEDIT" & _
				" , FCIP" & _
				" , FCBA" & _
				" , STOCKCODE" & _
				" , STORECODE" & _
				" , LASTUPDATE" & _
				" , LASTTIME" & _
				" , STAGE" & _
				" )" & _
				" values (" & _
				" '" & v_DESCRIPTION & "'" & _
				" , '" & v_DESCRIPTION & "'" & _
				" , to_date('" & v_DATEPLANTED & "','mm/dd/yyyy')" & _
				" , '" & v_CROP & "'" & _
				" , '" & v_PLANTINGMATERIAL & "'" & _
				" , '" & v_CLONE & "'" & _
				" , '" & v_SOURCE & "'" & _
				" , '" & v_TYPE & "'" & _
				" , '" & v_STATUS & "'" & _
				" , '" & Session("CodeName") & "'" & _
				" , '" & Session("CodeName") & "'" & _
				" , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
				" , '" & Session("fcBA") & "'" & _
				" , '" & v_StockCode & "'" & _
				" , '" & v_StoreCode & "'" & _				
				" , sysdate" & _
				" , To_Char(sysdate,'HH24:MI')" & _		
				" , '" & v_Stage & "'" & _								
				" )"
				
	'	Response.Write q
	    connStr.execute(q)		
        cmdView	= "True"
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
		q	=	" Update Nursery set" & _
				" DESCRIPTION = '" & v_DESCRIPTION & "'" & _
				" , CROP = '" & v_CROP & "'" & _
				" , DATEPLANTED = to_date('" & v_DATEPLANTED & "','mm/dd/yyyy')" & _
				" , PLANTINGMATERIAL = '" & v_PLANTINGMATERIAL & "'" & _
				" , CLONE = '" & v_CLONE & "'" & _
				" , SOURCE = '" & v_SOURCE & "'" & _
				" , TYPE = '" & v_TYPE & "'" & _
				" , STATUS = '" & v_STATUS & "'" & _
				" , LASTUPDATE = sysdate " & _
				" , LASTTIME = To_Char(sysdate,'HH24:MI') " & _
				" , STAGE = '" & v_Stage & "'" & _
				" where" & _
				" NURSERYCODE = '" & v_NURSERYCODE & "'" & _
				" and fcBA = '" & Session("fcBA") & "'"
				
		'Response.Write q
		'response.Flush()
		connStr.execute(q)
		connStr.execute("commit")
		cmdView	= "True"
    else
        cmdView	= "True"
	end if								 
	'Response.Write "sql " & q
else
    cmdView	= request("cmdView")
End if
Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis

txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")
txtJenis = request("txtJenis")
If Err.Number <> 0 then
   Response.Write "<font color=red>" & "Error Number : " & Err.Number &  " Desc. " & Err.Description & "</font>"
End if
%>
 <!-------------ISI Payment Form!------------->
 
 <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-leaf" aria-hidden="true"></i>Header Nursery</div><% end if %>
 <form name="frmEntry" method="post" action="Nursery.asp">
	<table class="data" width="100%">
		<tr><td align="left" width="20%">Nursery Code</td>
		<td align="left"><input type="text" maxlength="10" name="v_NURSERYCODE" size="10" value="<%=v_NURSERYCODE%>"></td></tr>	
		<tr><td colspan="2"  align="left" class="fbp-doc-actions">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
		<a href="PrintNursery.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&v_NURSERYCODE=<%=v_NURSERYCODE%>&v_DATEPLANTED=<%=v_DATEPLANTED%>" 
		title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a>
		</td>
		</tr>
    </table>  
 </form>
<%if cmdView = "True" Then
   Dim rsVH,intRecordsPerPage,cColor
   set rsVH = server.CreateObject("Adodb.recordset")
   'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
	q	=	" select" & _
			" NURSERYCODE" & _
			" , DESCRIPTION" & _
			" , TO_CHAR(DATEPLANTED,'MM/DD/YYYY') DATEPLANTED" & _
			" , STAGE" & _
			" , CROP" & _
			" , PLANTINGMATERIAL" & _
			" , CLONE" & _
			" , SOURCE" & _
			" , TYPE" & _
			" , STATUS" & _
			" , STORECODE" & _
			" , STOCKCODE" & _
			" , FCENTRY" & _
			" , FCEDIT" & _
			" , FCIP" & _
			" , FCBA" & _
			" from  Nursery where" & _
			" fcBA ='" & session("fcBA") & "'"
			
If Len(Trim(v_NurseryCode)) > 0 Then
	q	=	q & " And Nurserycode = '" & v_NurseryCode & "'"
End If

	q	=	q & " Order By Nurserycode"
   'response.Write q
   Set rsVH = Server.CreateObject("ADODB.Recordset")
   rsVH.CursorType = 3
   rsVH.CursorLocation = 3
   rsVH.Open q, connstr%>
   <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-edit" aria-hidden="true"></i>Detail Item</div><% end if %>
   <form action="Nursery.Asp" method="post" name="eForm">
   <table class="data" cellpadding="0" cellspacing="0" width="100%">
		<tr><td align="left" width="20%">Description</td><td align="left"><input type="text" size="50" name="v_DESCRIPTION" value="<%=v_DESCRIPTION%>" maxlength="50"></td></tr>
		<tr><td align="left" width="20%">Date Planted</td><td align="left"><input class="datepicker" type="text" size="15" name="v_DATEPLANTED" value="<%=v_DATEPLANTED%>"></td></tr>
		<tr><td align="left" width="20%">Stage</td><td align="left">
		<select name="v_Stage">
					<option value="PN" <%if v_Stage = "PN" then%>selected<%end if%>>PN</option>
					<option value="MN" <%if v_Stage = "MN" then%>selected<%end if%>>MN</option>
	   </select>
		</td></tr>
		<tr><td align="left" width="20%">Crop</td><td align="left"><input type="text" size="15" name="v_CROP" value="<%=v_CROP%>" maxlength="15"></td></tr>
		<tr><td align="left" width="20%">Planting Material</td><td align="left"><input type="text" size="15" name="v_PLANTINGMATERIAL" value="<%=v_PLANTINGMATERIAL%>"  maxlength="50"></td></tr>
		<tr><td align="left" width="20%">Clone</td><td align="left"><input type="text" size="15" name="v_CLONE" value="<%=v_CLONE%>"  maxlength="50"></td></tr>
		<tr><td align="left" width="20%">Source</td><td align="left"><input type="text" size="15" name="v_SOURCE" value="<%=v_SOURCE%>"  maxlength="50" ></td></tr>
		<tr><td align="left" width="20%">Type</td><td align="left"><input type="text" size="15" name="v_TYPE" value="<%=v_TYPE%>"  maxlength="15"></td></tr>
		
		

		<tr>
			<td align="left" width="20%">Store Code</td>
			<td align="left">       
				<select name="v_StoreCode" value="<%=v_StoreCode%>"> 
					<%cmdCallCombo "store",v_StoreCode,"",""%>
				</select>
			</td>
		</tr>
		
		<tr>
			<td align="left" width="20%">Stock Code</td>
			<td align="left"><input type="text" size="15" name="v_StockCode" value="<%=v_StockCode%>"  maxlength="15"></td> 
		</tr>
		
		
		<tr>
		    <td align="left" width="20%">Status</td>
			<td align="left"><input type="text" size="5" name="v_STATUS" value="<%=v_STATUS%>"  maxlength="1"></td>
	    </tr>
		<tr>
		<td colspan="2" align="left">
		  <input type="hidden" name="v_NURSERYCODE" value="<%=v_NURSERYCODE%>">
		  <input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		  <input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		  <%IF Request("txtJenis")="Edit" then%>
		       <input type="hidden" name="txtJenis" value="Edit">
		       <input Type="Submit"  name="Add" value="Update">
		  <%Else%>
		       <input type="hidden" name="txtJenis" value="New">
		       <input Type="Submit"  name="Add" value="Add">
		  <%End if%>
		  <input Type="Reset"  name="Reset" value="Reset">
		  <input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)">
		</td>
	  </tr>
   </table>

   <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-list" aria-hidden="true"></i>Daftar Data</div><% end if %>
   <div class="fbp-doc-scroll">
   <table class="data fbp-doc-lines">
	<thead>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">CODE</TD>	
	<TD align="left">DESCRIPTION</TD>	
	<TD align="left">DATE PLANTED</TD>	
	<TD align="left">STAGE</TD>		
	<TD align="left">CROP</TD>	
	<TD align="left">PLANTING MATERIAL</TD>	
	<TD align="left">CLONE</TD>	
	<TD align="left">SOURCE</TD>	
	<TD align="left">TYPE</TD>	
	<TD align="left">STATUS</TD>	
	<TD align="left">STORE CODE</TD>	
	<TD align="left">STOCK CODE</TD>	
	</tr>
	</thead>
	<tbody>
   
   <%
   if Not rsVH.eof then
      nCountofField=rsVH.fields.count-1
	  no=1
	  While Not rsVH.EOF
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
		  <td align="left" valign="top"><%=NO%></td>
			<td valign="middle" width="30">

			<a href="Nursery.Asp?v_NURSERYCODE=<%=rsVH("NURSERYCODE")%>&v_DESCRIPTION=<%=rsVH("DESCRIPTION")%>&v_DATEPLANTED=<%=rsVH("DATEPLANTED")%>&v_Stage=<%=rsVH("STAGE")%>&v_CROP=<%=rsVH("CROP")%>&v_PLANTINGMATERIAL=<%=rsVH("PLANTINGMATERIAL")%>&v_CLONE=<%=rsVH("CLONE")%>&v_SOURCE=<%=rsVH("SOURCE")%>&v_TYPE=<%=rsVH("TYPE")%>&v_STATUS=<%=rsVH("STATUS")%>&v_StoreCode=<%=rsVH("STORECODE")%>&v_StockCode=<%=rsVH("STOCKCODE")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=request("v_NURSERYCODE")%>" border="0"></a></td><td>
			<a href="Nursery.Asp?v_NURSERYCODE=<%=rsVH("NURSERYCODE")%>&v_DESCRIPTION=<%=rsVH("DESCRIPTION")%>&v_DATEPLANTED=<%=rsVH("DATEPLANTED")%>&v_Stage=<%=rsVH("STAGE")%>&v_CROP=<%=rsVH("CROP")%>&v_PLANTINGMATERIAL=<%=rsVH("PLANTINGMATERIAL")%>&v_CLONE=<%=rsVH("CLONE")%>&v_SOURCE=<%=rsVH("SOURCE")%>&v_TYPE=<%=rsVH("TYPE")%>&v_STATUS=<%=rsVH("STATUS")%>&v_StoreCode=<%=rsVH("STORECODE")%>&v_StockCode=<%=rsVH("STOCKCODE")%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=request("v_NURSERYCODE")%>" border="0" onClick="return confirm('Are you sure to delete');"></a>
			</td>
		    <%for i = 0 to nCountofField-4%>
			    <td align="left" title="<%=rsVH(i).Name%>" valign="top"><%=rsVH(i)%></td>
		    <%next%>
		
		  </tr>
		  <%
		  rsVH.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   end if 
   rsVH.Close
   set rsVH = Nothing%>
   </tbody>
   </TABLE>
   </div>
   </form>
<%End if%>  
 
  <!-------------akhir!------------->
</div></div>
  <!--#include file="bottom.jsp"-->

</body>
</html>
