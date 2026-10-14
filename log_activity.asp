<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<html>
<head>
    <title>ePlaS - Chart</title>

		<!--<link rel="StyleSheet" href="dtree.css" type="text/css" />--->
	<link rel="StyleSheet" href="style-content.css" type="text/css" />
	<link class="include" rel="stylesheet" type="text/css" href="dist/jquery.jqplot.min.css" />
		<!--<link rel="stylesheet" type="text/css" href="dist/examples/examples.min.css" />--->
		<!--<script type="text/javascript" src="dtree.js"></script>-->
		<!--<script src="/js/jquery-1.9.1.js"></script>-->
		<!--<script src="/js/jquery-ui.js"></script>-->
		<!--[if lt IE 9]><script language="javascript" type="text/javascript" src="dist/excanvas.js"></script><![endif]-->
    <script class="include" type="text/javascript" src="js/jquery-1.9.1.js"></script>
	<script language="JavaScript" src='js/x.js'></script>
	<script language="JavaScript" src='js/xtableheaderfixed.js'></script>
	<script language="JavaScript">
	xAddEventListener(window, 'load',
	  function() {
		new xTableHeaderFixed('xthf-site', 'tabs-1');
		new xTableHeaderFixed('xthf-year', 'tabs-3');
	  }, false
	);
	</script>
	<script>
		$(function() {
		$( ".datepicker" ).datepicker();
		});
		$(function() {
			var h = window.innerHeight - 170;	
			var tabs_1 = document.getElementById("tabs-1");
			var tabs_2 = document.getElementById("tabs-2");
			var tabs_3 = document.getElementById("tabs-3");
			var tabs_4 = document.getElementById("tabs-4");
				tabs_1.style.height = h + "px";
				tabs_2.style.height = h + "px";
				tabs_3.style.height = h + "px";
				tabs_4.style.height = h + "px";
		}); 
	</script>

<style>

html, body {
    background: none repeat scroll 0 0 #fff;
    overflow: hidden;
}
.xthf-site thead th, .xthf-year thead th {
    border: 1px solid #fff;
}
.xthf-site th, .xthf-year th {
    background-color: #0c5b24;
    color: #ffffff;
    padding: 10px 20px 10px 10px;
    text-align: center;
}
.xthf-site, .xthf-year {
    border-collapse: collapse;
}
.xthf-site, .xthf-year {
    color: #000000;
    font-family: Verdana,Arial,Helvetica,sans-serif;
    font-size: 11px;
    text-align: right;
	text-align:left;
}
.xthf-site, .xthf-year {
    border-spacing: 1px;
}

		.none_style { border:none; padding:none;}
		.jqplot-image-button { font-size:11px;}
		table.data tr:nth-of-type(odd) {
			background-color:#FFF;
		}
		table.data tr {
			background-color:#fff;
		}
		.minion {
		}
		.ui-widget-content{color: #ffffff;}
		#tabs .jqplot-image-container {
			background-color:rgba(0,0,0,0.9);
			position:fixed;
			top:0px;
			bottom:0px;
			right:0px;
			left:0px;
			z-index: 1000;
			padding-top:20px;
		}
		#tabs .jqplot-image-container .jqplot-image-container-content {
			text-align: center;
		}
		.jqplot-image-button{
		}
		.ui-widget-header{overflow: hidden;}
		#tabs a.jqplot-image-container-close:after {
			display:none;
		}
		#tabs a.jqplot-image-container-close {
			background: linear-gradient(to bottom, #fff, #ddd) repeat scroll 0 0 #ddd;
			border-radius: 0;
			box-shadow: 0 2px 2px rgba(0, 0, 0, 0.4);
			color: #444;
			float: right;
			font-size: 12px;
			padding: 0.5em;
			margin-top:-5px;
			position: relative;
			text-decoration: none;
			text-shadow: 0 1px 0 rgba(255, 255, 255, 0.8);
		}
		.jqplot-image-container-header{
			Padding: 10px 20px;
			background:green;
			margin-bottom:10px;
			font-size:14px;
		}
		#product-table {
		 border-collapse: collapse;
		 }
		#product-table thead th{ border: 1px #fff solid;}
		#product-table tbody td{ border: 1px #ccc solid;}
		#product-table tfoot td{ border: 1px #FFCC00 solid;background-color:#FFCC00;}
	</style>
</head>

<body>
<!--#include file="top.jsp"-->
<% if Session("fcBa") <> "" then %>		
<div class="judul_head" valign="top" bgcolor="#fff" style=" width:100%;LEFT:0%;bottom:0%; border:none;overflow:hidden;"> 
<!-- Example scripts go here -->
<%
Dim txttitle,wheretitle
Dim txtPRDATE_to,height_pie
dim count_log : count_log = 0
height_pie = 300
height_plus = 20

	function GetLastDayByDate(aDate)
		dim intMonth
		dim dteFirstDayNextMonth

		dtefirstdaynextmonth = dateserial(year(adate),month(adate) + 1, 1)
		GetLastDayByDate = Day(DateAdd ("d", -1, dteFirstDayNextMonth))
	end function
	'FUNGSI ARRAY----------------------------
		Function GetRSArray(strSQL)
			'Response.Write(strSQL)
			set objRSArray = connStr.Execute(strSQL)
			If Not objRSArray.EOF Then
				arrRS = objRSArray.GetRows()
				GetRSArray = arrRS
			Else
				'Response.Write("There are currently no records returned.<br/>")
                Exit Function
			End If
			objRSArray.Close()
			Set objRSArray = Nothing
		End Function
		
		Function GetRSArray_DETAIL(strSQL)
			'Response.Write(strSQL)
			set objRSArray = connStr.Execute(strSQL)
			If Not objRSArray.EOF Then
				arrRSDETAIL = objRSArray.GetRows()
				GetRSArray_DETAIL = arrRSDETAIL
			Else
				'Response.Write("There are currently no records returned.<br/>")
                Exit Function
			End If
			objRSArray.Close()
			Set objRSArray = Nothing
		End Function
	'END FUNGSI ARRAY----------------------------	
txttitle = Request("title")
txtPRDATE_to =Request("PRDATE_to")
	'---- Data Header
 if Request("PRDATE_to")<>"" then
   txtPRDATE_to   =Request("PRDATE_to")
 else
   txtPRDATE_to   =textDate(date)
 end if
 
 if txttitle <> "" then
	if txttitle = "ST_GM" then
		wheretitle = "And b.Title in ('ST_GMM','ST_GEM') "
	else
		wheretitle = "And b.Title = '" & txttitle & "'"
	end if
 elseif txttitle = "" then
	wheretitle = "And b.Title IN (select fcCode from usertitle)"
 end if

 
 
stringUSER  = 	"SELECT distinct (USERNAME), BU, title,fullname,region,fccompanyname, COUNT(USERNAME) AKTIFITAS " & _
				"FROM (select a.*, b.title,b.fullname, c.region,c.fccompanyname from historical_user_tab a, USERS_TAB b,businessunit c " & _
				"where " & _
				"a.username = b.FCCODE  and a.BU = c.FCCODE and a.username <> 'ADMIN' and c.isactive = 'TRUE' and b.isactive = 'TRUE'" & wheretitle & ") "
				
if  LEN(TRIM(txtPRDATE_to)) > 0 then
	stringUSER	=	stringUSER & "WHERE TO_CHAR(LASTUPDATE,'YYYY') = '"& YEAR(txtPRDATE_to) &"' " & _
					"GROUP BY USERNAME, BU, title,fullname,region,fccompanyname " & _
					"ORDER BY USERNAME,region,fccompanyname,BU"
	else
	stringUSER	=	stringUSER & "WHERE TO_CHAR(LASTUPDATE,'YYYY') = '"& YEAR(textDate(date)) &"' " & _
					"GROUP BY USERNAME, BU, title,fullname,region,fccompanyname " & _
					"ORDER BY USERNAME,region,fccompanyname,BU"
	end if		

stringDATAUSER  = 	"SELECT USERNAME,BU,TO_CHAR(LASTUPDATE,'DD') TGL,TO_NUMBER(TO_CHAR(LASTUPDATE,'MM')) BLN,TO_CHAR(LASTUPDATE,'YYYY') THN, COUNT(USERNAME) AKTIFITAS,TITLE " & _
				"FROM (select a.*,b.title from historical_user_tab a,USERS_TAB b " & _
				"where " & _
				"a.username = b.FCCODE and a.username <> 'ADMIN' and b.isactive = 'TRUE' " & wheretitle & ") "
			  
if  LEN(TRIM(txtPRDATE_to)) > 0 then
	stringDATAUSER	=	stringDATAUSER & "WHERE TO_CHAR(LASTUPDATE,'YYYY') = '"& YEAR(txtPRDATE_to) &"' " & _
					"GROUP BY TO_CHAR(LASTUPDATE,'DD'),TO_CHAR(LASTUPDATE,'MM'),TO_CHAR(LASTUPDATE,'YYYY'),USERNAME,BU,TITLE " & _
					"ORDER BY USERNAME,TO_CHAR(LASTUPDATE,'MM'),TO_CHAR(LASTUPDATE,'DD')"
	else
	stringDATAUSER	=	stringDATAUSER & "WHERE TO_CHAR(LASTUPDATE,'MM') = '"& MONTH(textDate(date)) &"' AND TO_CHAR(LASTUPDATE,'YYYY') = '"& YEAR(txtPRDATE_to) &"' " & _
					"GROUP BY TO_CHAR(LASTUPDATE,'DD'),TO_CHAR(LASTUPDATE,'MM'),TO_CHAR(LASTUPDATE,'YYYY'),USERNAME,BU,TITLE " & _
					"ORDER BY USERNAME,TO_CHAR(LASTUPDATE,'MM'),TO_CHAR(LASTUPDATE,'DD')"
	end if		

'response.write(stringDATAUSER)
	
dim SPACE,test
SPACE = "  "

arrRS = GetRSArray(stringUSER)
arrRSDetail = GetRSArray_DETAIL(stringDATAUSER)	
%>

 <script>
	var data_history = [<% If IsArray(arrRS) Then
		For i = LBound(arrRS, 2) To UBound(arrRS, 2)
		%>{'id':['<%=arrRS(0, i)%>','<%=arrRS(1, i)%>','<%=arrRS(2, i)%>','<%=arrRS(6, i)%>','<%=arrRS(3, i)%>','<%=arrRS(4, i)%>','<%=arrRS(5, i)%>'],
	<% For m = 1 to 12 %>
		'<%=MonthName(m)%>':{'data_tgl':[
		<%	If IsArray(arrRSDetail) Then
				For x = LBound(arrRSDetail, 2) To UBound(arrRSDetail, 2)
					IF m = cint(arrRSDetail(3, x)) and arrRSDetail(0, x) = arrRS(0, i) then 
					%>
					'<%=cint(arrRSDetail(2, x))%>',
					<%
					end if
				Next
			End If 
			%>],
			'data_act':[<%
				If IsArray(arrRSDetail) Then
				For x = LBound(arrRSDetail, 2) To UBound(arrRSDetail, 2)
					IF m = cint(arrRSDetail(3, x)) and arrRSDetail(0, x) = arrRS(0, i) then 
					%>
					'<%=arrRSDetail(5, x)%>',
					<%
					end if
				Next
			End If %>]
	},
	<%
	Next
	%>
	},
	<%
		Next
		'Erase arrRS
		Erase arrRSDetail
	End If 

	%>];
</script>


 <table style="float:right; position:relative; z-index:1; top:2.5px; right:10px;">
 <tr>
	<form class="none_style" action="log_activity.Asp" method="post" name="eForm">
		<td align="left"><font face="Verdana" size="1">Title</font></td>
		<td align="left">
			<select name="TITLE">
			
			<%
			
			q = "select fcCode,description from usertitle"
			Set rstitle = Server.CreateObject("ADODB.Recordset")
			rstitle.CursorType = 3
			rstitle.CursorLocation = 3
			rstitle.Open q, connstr
			if not rstitle.eof then
			%>
			<OPTION value="">-- Title --</option>
			<option value="ST_GM" <%if txttitle = "ST_GM" then%> Selected <%End if%>>ST_GM  : GEM/GMM</option>
			<%
			 Do While Not rstitle.eof
				if rstitle("fcCode") = "ST_GEM" or rstitle("fcCode") = "ST_GMM" then
				else
			%>
					<option value="<%=rstitle("fcCode")%>" <%if txttitle = rstitle("fcCode") then%> Selected <%End if%>><%=rstitle("fcCode")%> : <%=rstitle("description")%></option>
			<%
				end if
			rstitle.movenext
			 loop
			End if
			rstitle.Close
			set rstitle = Nothing
			q=""
			%>
			</SELECT>
		</td>
		<td>
			<font style="font-size:10px;">Date : </font></td><td><input class="datepicker" type="text" maxlength="10" name="PRDATE_To" size="10" value="<%=txtPRDATE_to%>">
		</td>
		<td><input type="Submit" name="View" value="View"></td>
	  </form> 
	  </tr>
  </table>

  <div id="tabs" style="padding:0px;border:none;">

	<!-- Menu Tab -->
    <ul style="margin-top:-3px;">
      <li style="margin: 0em 0.2em 0em 0em;"><a href="#tabs-1">Month</a></li>	  
      <li style="margin: 0em 0.2em 0em 0em;"><a href="#tabs-2">Graphic Chart / Month</a></li>
	  <li style="margin: 0em 0.2em 0em 0em;"><a href="#tabs-3">Year</a></li>
      <li style="margin: 0em 0.2em 0em 0em;"><a href="#tabs-4">Pie Chart / Year</a></li>

    </ul>
	

	<!--#Tab 1 -->
    <div id="tabs-1" style=" overflow: auto; width: 100%;position:relative;">
			<%
			dim jml_hari
			jml_hari = GetLastDayByDate(txtPRDATE_to)
			%>
		<table border="0" id="product-table" class="xthf-site data" style="font-size: 11px; text-align:left; float:left; width:50%;">
		<thead>	
			<tr>
			<th rowspan="2" style="text-align:center;border-bottom:1px #ccc solid; width:5px;">No</th>
			<th rowspan= "2" style="width:auto;border-bottom:1px #ccc solid;">Nama</th>
			<%  
			if txttitle = "ST_GM" then %>
			<th rowspan= "2" style="border-bottom:1px #ccc solid;">REGION</th>
			<th rowspan= "2" style="border-bottom:1px #ccc solid;">PT</th>
			<%
			END IF
			%>
			<th rowspan= "2" style="text-align:center;border-bottom:1px #ccc solid;width:5px;">BU</th>
			<th colspan="<%=jml_hari%>" style="text-align:left;"> <%=MonthName(month(txtPRDATE_to))%>&nbsp;&nbsp;<%=Year(txtPRDATE_to)%></th>
			<th rowspan= "2" style="width:50px; border-bottom:1px #ccc solid;">Jumlah Log</th>
			</tr>
			<tr>
			<%
			For i = 1 To jml_hari
				if i = day(txtPRDATE_to) then
					response.write("<th style='background:red;width:20px; border-bottom:1px #ccc solid;'>" & i & "</th>")
				else
					response.write("<th style='width:20px; border-bottom:1px #ccc solid;'>" & i & "</th>")
				end if
			Next
			%>
			</tr>
		</thead>
		<tbody>
		
		<script>
			var ST_GEM = "<%=txttitle%>";
			var jml_data = data_history.length;
			var rating = [];
			total = 0
			//var strmnt = "May";
			//alert(data_history[0].<%=MonthName(month(txtPRDATE_to))%>.data_tgl[0]);
			var user_name_and_sum = [];
			no = 1;
			id_name = [];
			id_row = [];
			
			if (ST_GEM == "ST_GM"){
				//== region ===
				id_region = [];
				row_region = [];
				//=== PT ====
				id_pt = [];
				row_pt = [];
			}
			
			no = 1;
			prev = 0 ;
			for (i = 0; i < jml_data; i++) {
				var jml_data_bln = data_history[i].<%=MonthName(month(txtPRDATE_to))%>.data_tgl.length;
				//var id_row = "name_" + no + "";
				//var name_id = document.getElementById(id_row);
				
				//if (jml_data_bln > 0 ){
					document.write("<tr>");
					
					if (ST_GEM == "ST_GM"){
					
							if (i === 0 ){
								rspn = 0;
								rspn_region = 0;
								rspn_pt = 0;
								document.write("<td id='" + data_history[i].id[0] + "_no'>" + no + "</td>");
								document.write("<td id='" + data_history[i].id[0].replace(/ /g,"_") + "'>" + data_history[i].id[4] + "</td>");	
								document.write("<td id='" + data_history[i].id[0].replace(/ /g,"_") + data_history[i].id[5].replace(/ /g,"_") +  "_region'>" +data_history[i].id[5] + "</td>");
								document.write("<td id='" + data_history[i].id[0].replace(/ /g,"_") + data_history[i].id[5].replace(/ /g,"_") + data_history[i].id[6].replace(/ /g,"_") +  "_pt'>" + data_history[i].id[6] + "</td>");
								document.write("<td>" + data_history[i].id[1] + "</td>");
								rspn = rspn + 1 ;
								no = no + 1;
								rspn_region = rspn_region + 1 ;
								rspn_pt = rspn_pt + 1 ;
							}else{
								if (data_history[i].id[0] != data_history[prev].id[0]){
									//array ============
									id_row.push(rspn);
									id_name.push(data_history[prev].id[0].replace(/ /g,"_"));
									
									//END:array ========
								rspn = 0;
									document.write("<td id='" + data_history[i].id[0].replace(/ /g,"_") + "_no'>" + no + "</td>");
									document.write("<td id='" + data_history[i].id[0].replace(/ /g,"_") + "'>" + data_history[i].id[4] + "</td>");
									
								rspn = rspn + 1 ;
								no = no + 1;
								}else if (data_history[i].id[0] == data_history[prev].id[0] && i == jml_data-1){
									rspn = rspn + 1 ;
									//array Terakhir============
									id_row.push(rspn);
									id_name.push(data_history[prev].id[0].replace(/ /g,"_"));
									//END:array Terakhir========
									document.write("");	
								}else if (data_history[i].id[0] == data_history[prev].id[0]){
									document.write("");
									rspn = rspn + 1 ;	
								}else{
									alert("error : Name " + i + "/" + prev)
								}
								//=========================Region====================================
								if (data_history[i].id[5] != data_history[prev].id[5] || data_history[i].id[0] != data_history[prev].id[0]){
									row_region.push(rspn_region);
									id_region.push(data_history[prev].id[0].replace(/ /g,"_") + data_history[prev].id[5].replace(/ /g,"_") + "_region");
									rspn_region = 0;
									document.write("<td id='" + data_history[i].id[0].replace(/ /g,"_") + data_history[i].id[5].replace(/ /g,"_") + "_region'>" + data_history[i].id[5] + "</td>");
									rspn_region = rspn_region + 1 ;
								}else if (data_history[i].id[5] === data_history[prev].id[5] && data_history[i].id[0] === data_history[prev].id[0] && i == jml_data-1){
									rspn_region = rspn_region + 1 ;
									row_region.push(rspn_region);
									id_region.push(data_history[prev].id[0].replace(/ /g,"_") + data_history[prev].id[5].replace(/ /g,"_") + "_region");
									document.write("");
									
								}else if (data_history[i].id[5] === data_history[prev].id[5] && data_history[i].id[0] === data_history[prev].id[0]){
									document.write("");
									rspn_region = rspn_region + 1 ;
								}else{
									alert("error : region" + i + "/" + prev)
								}

								//============================PT=================================
								if ((data_history[i].id[6] != data_history[prev].id[6] || data_history[i].id[5] != data_history[prev].id[5] || data_history[i].id[0] != data_history[prev].id[0]) && i != jml_data-1 ){
									row_pt.push(rspn_pt);
									id_pt.push(data_history[prev].id[0].replace(/ /g,"_") + data_history[prev].id[5].replace(/ /g,"_") + data_history[prev].id[6].replace(/ /g,"_") + "_pt");
									rspn_pt = 0;
									document.write("<td id='" + data_history[i].id[0].replace(/ /g,"_") + data_history[i].id[5].replace(/ /g,"_") +data_history[i].id[6].replace(/ /g,"_") + "_pt' >" + data_history[i].id[6] + "</td>");
									rspn_pt = rspn_pt + 1 ;
								}else if (data_history[i].id[6] === data_history[prev].id[6] && data_history[i].id[5] === data_history[prev].id[5] && data_history[i].id[0] === data_history[prev].id[0] && i == jml_data-1 ){
									rspn_pt = rspn_pt + 1 ;
									row_pt.push(rspn_pt);
									id_pt.push(data_history[prev].id[0].replace(/ /g,"_") + data_history[prev].id[5].replace(/ /g,"_") + data_history[prev].id[6].replace(/ /g,"_") + "_pt");
									document.write("");
								}else if (data_history[i].id[6] === data_history[prev].id[6] && data_history[i].id[5] === data_history[prev].id[5] && data_history[i].id[0] === data_history[prev].id[0]){
									document.write("");
									rspn_pt = rspn_pt + 1 ;
								}else{
									alert("error : PT year " + data_history[i].id[6] +"/" + i + "/" + no)
									//document.write("<td id='" + data_history[i].id[6].replace(/ /g,"_") + "_pt' >" + data_history[i].id[6] + i + "</td>");
								}
								//================================BU=============================
								if (data_history[i].id[1] != data_history[prev].id[1]){
									document.write("<td>" + data_history[i].id[1] + "</td>");
								}else if (data_history[i].id[1] === data_history[prev].id[1] && data_history[i].id[6] === data_history[prev].id[6] && data_history[i].id[5] === data_history[prev].id[5] && data_history[i].id[0] === data_history[prev].id[0]){
									document.write("");
								}else{
									alert("error : " + i + "/" + prev + 1)
								}
							}
					}else{
								if (i === 0 ){
									rspn = 0;
									document.write("<td id='" + data_history[i].id[0] + "_no'>" + no + "</td>");
									document.write("<td id='" + data_history[i].id[0].replace(/ /g,"_") + "'>" + data_history[i].id[4] + "</td>");	
									document.write("<td>" + data_history[i].id[1] + "</td>");
									rspn = rspn + 1 ;
									no = no + 1;
								
								}else{
									if (data_history[i].id[0] != data_history[prev].id[0]){
									//array ============
									id_row.push(rspn);
									id_name.push(data_history[prev].id[0].replace(/ /g,"_"));
									
									//END:array ========
									rspn = 0;
										document.write("<td id='" + data_history[i].id[0].replace(/ /g,"_") + "_no'>" + no + "</td>");
										document.write("<td id='" + data_history[i].id[0].replace(/ /g,"_") + "'>" + data_history[i].id[4] + "</td>");
										
									rspn = rspn + 1 ;
									no = no + 1;
									}else if (data_history[i].id[0] == data_history[prev].id[0] && i == jml_data-1){
										rspn = rspn + 1 ;
										//array Terakhir============
										id_row.push(rspn);
										id_name.push(data_history[prev].id[0].replace(/ /g,"_"));
										//END:array Terakhir========
										document.write("");	
									}else if (data_history[i].id[0] == data_history[prev].id[0]){
										document.write("");
										rspn = rspn + 1 ;	
									}else{
										alert("error : Name " + i + "/" + prev)
									}
									//=============================================================
									if (data_history[i].id[1] != data_history[prev].id[1] || data_history[i].id[0] != data_history[prev].id[0]){
										document.write("<td>" + data_history[i].id[1] + "</td>");
									}else if (data_history[i].id[1] === data_history[prev].id[1] && data_history[i].id[0] === data_history[prev].id[0]){
										document.write("");
									}else{
										alert("error : " + data_history[i].id[0] + data_history[i].id[1] +  i + "/" + prev + 1)
									}
								}
								//document.write("<td>" + no + "</td><td>" + data_history[i].id[4] + "</td><td>" + data_history[i].id[1] + "</td>");
								//no = no + 1;
					}
						
					sum = "";
						var rate = [];
						
						for (x = 1; x <= <%=jml_hari%>; x++) {
						var d = new Date(<%=year(txtPRDATE_to)%>, <%=(month(txtPRDATE_to)-1)%>, x); 
						var weekly = d.getDay();  
						//alert();
								if (weekly === 0){
									minggu = "style='background-color:#f1f1f1; color:#000;'";
								}else{
									minggu = "";
								}
							document.write("<td " + minggu +" align='center'>");
							//document.write(data_history[i].data_tgl.indexOf("" + x + ""));
							//alert(data_history[i].<%=MonthName(month(txtPRDATE_to))%>.data_tgl.indexOf("" + x + ""));
							if(data_history[i].<%=MonthName(month(txtPRDATE_to))%>.data_tgl.indexOf("" + x + "") < 0){
								document.write(".");
								rate.push(0);
							}else{
								document.write(data_history[i].<%=MonthName(month(txtPRDATE_to))%>.data_act[data_history[i].<%=MonthName(month(txtPRDATE_to))%>.data_tgl.indexOf("" + x + "")]);
									
								sum = Number(sum) + Number(data_history[i].<%=MonthName(month(txtPRDATE_to))%>.data_act[data_history[i].<%=MonthName(month(txtPRDATE_to))%>.data_tgl.indexOf("" + x + "")]);
									
								rate.push(data_history[i].<%=MonthName(month(txtPRDATE_to))%>.data_act[data_history[i].<%=MonthName(month(txtPRDATE_to))%>.data_tgl.indexOf("" + x + "")]);
									
							}
							document.write("</td>");
						}
					if (sum == "" || sum == 0){
						document.write("<td align='Right'>0</td>");
					}else{
						document.write("<td align='Right'>" + sum + "</td>");
					}
					document.write("</tr>");
					total = Number(total) + Number(sum);
					rating.push(rate);
					user_name_and_sum.push([data_history[i].id[0],sum]);
					prev = i;
					//alert(prev + "/" + rate);
				}
//}
			
			//alert(id_name[0]);
			for (c = 0; c < id_name.length; c++) {
				document.getElementById(id_name[c] + "_no").rowSpan = id_row[c];
				document.getElementById(id_name[c] + "").rowSpan = id_row[c];
				//alert(id_pt[c]);
			}
			if (ST_GEM == "ST_GM"){
				for (c = 0; c < id_region.length; c++) {
					document.getElementById(id_region[c]).rowSpan = row_region[c];
				//alert(id_region[c]);
				}

				//alert("test");
				for (c = 0; c < id_pt.length; c++) {
				//alert(id_pt[c]);
					document.getElementById(id_pt[c]).rowSpan = row_pt[c];
				//alert(id_pt[c]);
				}
			}
			
				var jml_rating = [];
			
				for (h = 0; h < <%=jml_hari-1%>; h++) {
					jml_perrow = "";
					for (r = 0; r < rating.length; r++) {
						jml_perrow = Number(jml_perrow) + Number(rating[r][h]);
					}
					jml_rating.push(jml_perrow)
				}
			
		</script>
		</tbody>
		<tfoot>
			<tr class="nothover">
			<% if txttitle = "ST_GM" then %>
				<td align="left" style="font-weight:bold;<%=cmdstyleLeft%>" colspan="<%=5+jml_hari%>"> <font face="Verdana" size="1">T o t a l</td>
			<% else %>
				<td align="left" style="font-weight:bold;<%=cmdstyleLeft%>" colspan="<%=3+jml_hari%>"> <font face="Verdana" size="1">T o t a l</td>
			<% end if %>
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><script>document.write(total);</script></td>
			</tr>
		</tfoot>	
			
		</table>
		<div style="clear:both; content:' ';"> </div>
    </div>
	
	
   
	<div id="tabs-2" style=" overflow: scroll; width: 100%;">
      <div id="chart1" style="margin-top:20px; margin-left:20px;"></div>
    </div>
 
   
 <!--#Tab 4 -->
	<div id="tabs-3" style=" overflow: auto; width: 100%;position:relative;">
		<%
			dim jml_bln
			jml_bln = 12
			%>
		<table id="product-table" class="xthf-year data" border="0"  style="font-size: 11px; text-align:left; float:left; width:100%;">
		<thead>	
			<tr>
			<th rowspan="2" style="text-align:center;border-bottom:1px #ccc solid; width:5px;">No</th>
			<th rowspan= "2" style="border-bottom:1px #ccc solid;">Nama</th>
			<%  
			if txttitle = "ST_GM" then %>
			<th rowspan= "2" style="border-bottom:1px #ccc solid;">REGION</th>
			<th rowspan= "2" style="border-bottom:1px #ccc solid;">PT</th>
			<%
			END IF
			%>
			<th rowspan= "2" style="text-align:center;border-bottom:1px #ccc solid;width:5px;">BU</th>
			<th colspan="<%=jml_bln%>" style="text-align:left;"><%=Year(txtPRDATE_to)%></th>
			<th rowspan= "2" style="width:50px; border-bottom:1px #ccc solid;">Jumlah Log</th>
			</tr>
			<tr>
			<%
			For i = 1 To jml_bln
			if i = Month(txtPRDATE_to) then
					response.write("<th style='background:red;width:20px; border-bottom:1px #ccc solid;'>" & MonthName(i) & "</th>")
				else
					response.write("<th style='width:20px; border-bottom:1px #ccc solid;'>" & MonthName(i) & "</th>")
				end if
			Next
			%>
			</tr>
		</thead>
		<tbody>

		<script>
			var jml_data = data_history.length;
			//alert("ST : " + ST_GEM);
			//var rating = [];	
			total_mth = 0;			
			//alert(data_history[0].<%=MonthName(month(txtPRDATE_to))%>.data_tgl[0]);
			//== nama ========
			rating_year = [];
			id_name_year = [];
			id_row_year = [];
			if (ST_GEM == "ST_GM"){
				//== region ===
				id_region_year = [];
				row_region_year = [];
				//=== PT ====
				id_pt_year = [];
				row_pt_year = [];
			}
			
			no_year = 1;
			prev_year = 0 ;
			for (i = 0; i < jml_data; i++) {	
					document.write("<tr>");
					if (ST_GEM == "ST_GM"){
							if (i === 0 ){
								rspn_year = 0;
								rspn_year_region = 0;
								rspn_year_pt = 0;
								document.write("<td id='" + data_history[i].id[0] + "_no_year'>" + no_year + "</td>");
								document.write("<td id='" + data_history[i].id[0].replace(/ /g,"_") + "_year'>" + data_history[i].id[4] + "</td>");	
								document.write("<td id='" + data_history[i].id[0].replace(/ /g,"_") + data_history[i].id[5].replace(/ /g,"_") +  "_region_year'>" + data_history[i].id[5] + "</td>");
								document.write("<td id='" + data_history[i].id[0].replace(/ /g,"_") + data_history[i].id[5].replace(/ /g,"_") + data_history[i].id[6].replace(/ /g,"_") +  "_pt_year'>" + data_history[i].id[6] + "</td>");
								document.write("<td>" + data_history[i].id[1] + "</td>");
								rspn_year = rspn_year + 1 ;
								no_year = no_year + 1;
								rspn_year_region = rspn_year_region + 1 ;
								rspn_year_pt = rspn_year_pt + 1 ;
							}else{
								if (data_history[i].id[0] != data_history[prev_year].id[0]){
									//array ============
									id_row_year.push(rspn_year);
									id_name_year.push(data_history[prev_year].id[0].replace(/ /g,"_"));
									
									//END:array ========
								rspn_year = 0;
									document.write("<td id='" + data_history[i].id[0].replace(/ /g,"_") + "_no_year'>" + no_year + "</td>");
									document.write("<td id='" + data_history[i].id[0].replace(/ /g,"_") + "_year'>" + data_history[i].id[4] + "</td>");
									
								rspn_year = rspn_year + 1 ;
								no_year = no_year + 1;
								}else if (data_history[i].id[0] == data_history[prev_year].id[0] && i == jml_data-1){
									rspn_year = rspn_year + 1 ;
									document.write("");
									//array Terakhir============
									id_row_year.push(rspn_year);
									id_name_year.push(data_history[i].id[0].replace(/ /g,"_"));
									//END:array Terakhir========
										
								}else if (data_history[i].id[0] == data_history[prev_year].id[0]){
									document.write("");
									rspn_year = rspn_year + 1 ;
								}else{
									alert("error : Name " + i + "/" + prev_year)
								}
								//=============================================================
								if ((data_history[i].id[5] != data_history[prev_year].id[5] || data_history[i].id[0] != data_history[prev_year].id[0]) && i != jml_data-1){
									row_region_year.push(rspn_year_region);
									id_region_year.push(data_history[prev_year].id[0].replace(/ /g,"_") + data_history[prev_year].id[5].replace(/ /g,"_") + "_region_year");
									rspn_year_region = 0;
									document.write("<td id='" + data_history[i].id[0].replace(/ /g,"_") + data_history[i].id[5].replace(/ /g,"_") + "_region_year'>" + data_history[i].id[5] + "</td>");
									rspn_year_region = rspn_year_region + 1 ;
								}else if (data_history[i].id[5] === data_history[prev_year].id[5] && data_history[i].id[4] === data_history[prev_year].id[4] && i == jml_data-1){
									document.write("");
									rspn_year_region = rspn_year_region + 1 ;
									row_region_year.push(rspn_year_region);
									id_region_year.push(data_history[prev_year].id[0].replace(/ /g,"_") + data_history[prev_year].id[5].replace(/ /g,"_") + "_region_year");
								}else if (data_history[i].id[5] === data_history[prev_year].id[5] && data_history[i].id[4] === data_history[prev_year].id[4]){
									document.write("");
									rspn_year_region = rspn_year_region + 1 ;
									
								}else{
									alert("error : region_year" + i + "/" + prev_year)
								}

								//=============================================================
								if (data_history[i].id[6] != data_history[prev_year].id[6] || data_history[i].id[5] != data_history[prev_year].id[5] || data_history[i].id[0] != data_history[prev_year].id[0] ){
									row_pt_year.push(rspn_year_pt);
									id_pt_year.push(data_history[prev_year].id[0].replace(/ /g,"_") + data_history[prev_year].id[5].replace(/ /g,"_") + data_history[prev_year].id[6].replace(/ /g,"_") + "_pt_year");
									rspn_year_pt = 0;
									document.write("<td id='" + data_history[i].id[0].replace(/ /g,"_") + data_history[i].id[5].replace(/ /g,"_") + data_history[i].id[6].replace(/ /g,"_") + "_pt_year' >" + data_history[i].id[6] + "</td>");
									rspn_year_pt = rspn_year_pt + 1 ;
								}else if (data_history[i].id[6] === data_history[prev_year].id[6] && data_history[i].id[5] === data_history[prev_year].id[5] && data_history[i].id[0] === data_history[prev_year].id[0] && i == jml_data-1){
									document.write("");
									rspn_year_pt = rspn_year_pt + 1 ;
									row_pt_year.push(rspn_year_pt);
									id_pt_year.push(data_history[prev_year].id[0].replace(/ /g,"_") + data_history[prev_year].id[5].replace(/ /g,"_") + data_history[prev_year].id[6].replace(/ /g,"_") + "_pt_year");
								}else if (data_history[i].id[6] === data_history[prev_year].id[6] && data_history[i].id[5] === data_history[prev_year].id[5] && data_history[i].id[0] === data_history[prev_year].id[0]){
									document.write("");
									rspn_year_pt = rspn_year_pt + 1 ;
								}else{
									alert("error : PT year " + data_history[i].id[6] +"/" + i + "/" + no_year)
									//document.write("<td id='" + data_history[i].id[6].replace(/ /g,"_") + "_pt_year' >" + data_history[i].id[6] + i + "</td>");
								}
								//=============================================================
								if (data_history[i].id[1] != data_history[prev_year].id[1]){
									document.write("<td>" + data_history[i].id[1] + "</td>");
								}else if (data_history[i].id[1] === data_history[prev_year].id[1] && data_history[i].id[6] === data_history[prev_year].id[6] && data_history[i].id[5] === data_history[prev_year].id[5] && data_history[i].id[0] === data_history[prev_year].id[0]){
									document.write("");
								}else if (data_history[i].id[1] = data_history[prev_year].id[1] && data_history[i].id[0] != data_history[prev_year].id[0]){
									document.write("<td>" + data_history[i].id[1] + "</td>");
								}else{
									alert("error Year : " + data_history[i].id[0])
								}
							}
					} else {
							if (i === 0 ){
								rspn_year = 0;
								document.write("<td id='" + data_history[i].id[0] + "_no_year'>" + no_year + "</td>");
								document.write("<td id='" + data_history[i].id[0].replace(/ /g,"_") + "_year'>" + data_history[i].id[4] + "</td>");	
								document.write("<td>" + data_history[i].id[1] + "</td>");
								rspn_year = rspn_year + 1 ;
								no_year = no_year + 1;

							}else{
								if (data_history[i].id[0] != data_history[prev_year].id[0]){
									//array ============
									id_row_year.push(rspn_year);
									id_name_year.push(data_history[prev_year].id[0].replace(/ /g,"_"));
									//END:array ========
								rspn_year = 0;
									document.write("<td id='" + data_history[i].id[0].replace(/ /g,"_") + "_no_year'>" + no_year + "</td>");
									document.write("<td id='" + data_history[i].id[0].replace(/ /g,"_") + "_year'>" + data_history[i].id[4] + "</td>");
									
								rspn_year = rspn_year + 1 ;
								no_year = no_year + 1;
								}else if (data_history[i].id[0] == data_history[prev_year].id[0] && i == jml_data-1){
									rspn_year = rspn_year + 1 ;
									document.write("");
									//array Terakhir============
									id_row_year.push(rspn_year);
									id_name_year.push(data_history[i].id[0].replace(/ /g,"_"));
									//END:array Terakhir========
										
								}else if (data_history[i].id[0] == data_history[prev_year].id[0]){
									document.write("");
									rspn_year = rspn_year + 1 ;
								}else{
									alert("error : Name " + i + "/" + prev_year)
								}
								//=============================================================
								if (data_history[i].id[1] != data_history[prev_year].id[1] || data_history[i].id[0] != data_history[prev_year].id[0]){
									document.write("<td>" + data_history[i].id[1] + "</td>");
								}else if (data_history[i].id[1] === data_history[prev_year].id[1] && data_history[i].id[0] === data_history[prev_year].id[0]){
									document.write("");
								}else{
									alert("error : " + data_history[i].id[0] + data_history[i].id[1] +  i + "/" + prev_year + 1)
								}
							}
							//document.write("<td>" + no_year + "</td><td>" + data_history[i].id[4] + "</td><td>" + data_history[i].id[1] + "</td>");
							//no_year = no_year + 1;
					}
					rate_year = [];
					month
					sum_moth = "";
						<% For m = 1 To 12 %>
								var jml_tgl = data_history[i].<%=MonthName(m)%>.data_tgl.length;
								document.write("<td align='center'>");	
								sum_day = "";
								if(data_history[i].<%=MonthName(m)%>.data_tgl.length > 0){
									for (t = 0; t < jml_tgl; t++){			
										sum_day = Number(sum_day) + Number(data_history[i].<%=MonthName(m)%>.data_act[t]);							
									}						
									document.write(sum_day);
									sum_moth = Number(sum_moth) + Number(sum_day);
									rate_year.push(sum_day);
								}else{
									document.write(".");
									rate_year.push(0);
								}
								document.write("</td>");
						<% next %>

					document.write("<td  align='Right'>" + sum_moth + "</td>");
					document.write("</tr>");
					total_mth = Number(total_mth) + Number(sum_moth);
					rating_year.push(rate_year);
					prev_year = i ;
			}
			//document.getElementById("IDRIS_DAMANIK_no_year").rowSpan = 6;
			//alert(id_name_year[id_name_year.length-1] + "_year")
			for (cy = 0; cy < id_name_year.length; cy++) {
				
				document.getElementById(id_name_year[cy] + "_no_year").rowSpan = id_row_year[cy];
				document.getElementById(id_name_year[cy] + "_year").rowSpan = id_row_year[cy];
			//alert(id_name_year[cy] + "_year : " + id_row_year[cy]);
			}
			if (ST_GEM == "ST_GM"){

				for (cyr = 0; cyr < id_region_year.length; cyr++) {
					document.getElementById(id_region_year[cyr]).rowSpan = row_region_year[cyr];
				//alert(id_region_year);
				}
				//alert("test");
				for (cyp = 0; cyp < id_pt_year.length; cyp++) {
				//alert(id_pt_year);
					document.getElementById(id_pt_year[cyp]).rowSpan = row_pt_year[cyp];
				
				}
			}
			var jml_rating_year = [];
			var month = new Array();
			month[0] = "January";
			month[1] = "February";
			month[2] = "March";
			month[3] = "April";
			month[4] = "May";
			month[5] = "June";
			month[6] = "July";
			month[7] = "August";
			month[8] = "September";
			month[9] = "October";
			month[10] = "November";
			month[11] = "December";
				for (y = 0; y < 11; y++) {
					jml_perrow_year = "";
					for (r = 0; r < rating_year.length; r++) {
						jml_perrow_year = Number(jml_perrow_year) + Number(rating_year[r][y]);
					}
					jml_rating_year.push([month[y],jml_perrow_year])
				}
			//alert(jml_rating_year);
			//rating_year.push([data_history[i].id[0],sum]);
		</script>
		</tbody>
		<tfoot>
			<tr class="nothover">
			
			<% if txttitle = "ST_GM" then %>
				<td align="left" style="font-weight:bold;<%=cmdstyleLeft%>" colspan="<%=17%>"> <font face="Verdana" size="1">T o t a l</td>
			<% else %>
				<td align="left" style="ont-weight:bold;<%=cmdstyleLeft%>" colspan="<%=15%>"> <font face="Verdana" size="1">T o t a l</td>
			<% end if %>
			<td align="Right" style="background-color:#FFCC00;<%=cmdstyle%>"> <font face="Verdana" size="1"><script>document.write(total_mth);</script></td>
			</tr>
			
		 </tfoot>
		</table>
		<div style="clear:both; content:' ';"> </div>
    </div>    
	
	 <div id="tabs-4" style=" overflow: scroll; width: 100%;">
       <div id="chart2" style="margin-top:20px; margin-left:20px;"></div>
    </div>
</div> 
 
<!-- End example scripts -->
 <script type="text/javascript">
	
    $(document).ready(function() {
	//alert(jml_rating_year);
        $.jqplot.config.enablePlugins = false;
		var ticks = [
			<% For i = 1 To jml_hari
				RESPONSE.write(i & ",")
			Next
			%>];
		var TT = "1 - <%=jml_hari%><%=SPACE%><%=MonthName(month(txtPRDATE_to))%> <%=SPACE%> <%=Year(txtPRDATE_to)%>";
        var l1 = jml_rating;
        var l2 = [[2, 14], [7, 2], [8,5]];
        var l3 = [4, 7, 9, 2, 11, 5, 9, 13, 8, 7];
        var l4 = jml_rating_year;
        $("#tabs").tabs();
        $("#accordion").accordion();

        var plot1 = $.jqplot('chart1', [l1],  {
          title: "Log Activity User Graph, " + TT,
		  height: 400,
		  width: 1000,
          lengend:{show:true},
		  animate: true,
			// Will animate plot on calls to plot1.replot({resetAxes:true})
			animateReplot: true,
			cursor: {
				show: true,
				zoom: true,
				looseZoom: true,
				showTooltip: false
			},
          seriesDefaults:{
                renderer:$.jqplot.BarRenderer,
                pointLabels: { show: true }
            },
            axes: {
                xaxis: {
                    renderer: $.jqplot.CategoryAxisRenderer,
                    ticks: ticks
                }
            },
          cursor:{show:true, zoom:true},
          axesDefaults:{useSeriesColor:true, rendererOptions: { alignTicks: true}}
        });
	
	var plot2 = jQuery.jqplot ('chart2', [l4], 
    { 
	  title: "Log Activity User Pie, " + TT,
	  height: 500,
      width: 1000,
      seriesDefaults: {
        // Make this a pie chart.
        renderer: jQuery.jqplot.PieRenderer, 
        rendererOptions: {
          // Put data labels on the pie slices.
          // By default, labels show the percentage of the slice.
          showDataLabels: true
        }
      }, 
      legend: { show:true, location: 'e' }
    });
        $('#tabs').bind('tabsactivate', function(event, ui) {
          if (ui.newTab.index() === 1 && plot1._drawCount === 0) {
            plot1.replot();
          }
          else if (ui.newTab.index() === 3 && plot2._drawCount === 0) {
            plot2.replot();
          }
        });

        $('#accordion').bind('accordionactivate', function(event, ui) {
          var index = $(this).find("h3").index ( ui.newHeader[0] );
          if (index === 1) {
            plot3.replot();
          }
        });
    
    });
</script> 
<!-- Don't touch this! -->


    <script class="include" type="text/javascript" src="dist/jquery.jqplot.min.js"></script>
   <!-- <script type="text/javascript" src="dist/examples/syntaxhighlighter/scripts/shCore.min.js"></script>
    <script type="text/javascript" src="dist/examples/syntaxhighlighter/scripts/shBrushJScript.min.js"></script>
    <script type="text/javascript" src="dist/examples/syntaxhighlighter/scripts/shBrushXml.min.js"></script>-->
<!-- End Don't touch this! -->

<!-- Additional plugins go here -->
  <script class="include" type="text/javascript" src="dist/plugins/jqplot.cursor.min.js"></script>
  <script class="include" type="text/javascript" src="dist/plugins/jqplot.pieRenderer.min.js"></script>
  <script class="include" type="text/javascript" src="dist/plugins/jqplot.ohlcRenderer.min.js"></script>
  <script class="include" type="text/javascript" src="dist/plugins/jqplot.categoryAxisRenderer.min.js"></script>
  <link class="include" type="text/css" href="css/jquery-ui_tab.css" rel="Stylesheet" /> 
  <script class="include" type="text/javascript" src="js/jquery-1.10.0.js"></script>

<!-- End additional plugins -->

 
    <script type="text/javascript" src="dist/capture.min.js"></script>

</div>
<%else%>

<%end if%>



    </body>
</html>
