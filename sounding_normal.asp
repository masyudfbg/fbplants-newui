<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<html>
	<head>
		<title>ePlaS-<%=request("txtTable")%></title>
		
		<!-- CSS -->
		<link rel="StyleSheet" href="style-content.css" type="text/css" />
		<link rel="StyleSheet" href="2col_leftNav.css" type="text/css" />
		<link rel="StyleSheet" href="dtree.css" type="text/css" />
		<link rel="stylesheet" href="css/jquery-ui.css" />
		<link rel="stylesheet" href="css/ns-window.css" />
		<link rel="stylesheet" href="css/font/font-awesome.css"/>
		
		<!-- JQUERY ENGINE -->
		<script type="text/javascript" src="js/jquery-1.8.3.min.js"></script>	
		
		<!-- FLEXIGRID -->
		<link rel="stylesheet" type="text/css" href="css/flexigrid.pack.css" />
		<script type="text/javascript" src="js/flexigrid.pack.js"></script>
		<!-- END FLEXIGRID -->
		
		<script type="text/javascript" src="dtree.js"></script>
		
		<!-- DATE PICKER -->
		<script src="js/jquery-ui.js"></script>
		<script language="JavaScript" src="/js/date-picker.js"></script> 
		<script>
			$(function() {
			$( ".datepicker" ).datepicker();
			});
		</script>
		<!-- END DATE PICKER -->
		
		<script type="text/javascript">
			function reloadPage()
			  {location.reload();}
		</script>
		
		<!-- FUNCTION TOGGLE "CLICK TO HIDE " -->
		<script>
					function toggle(id){
						var hfi = $("#WARP_input").height();
						var ns = document.getElementById("nswindow");
						var e = document.getElementById(id);
						if(e.style.display == 'table-row'){
							e.style.display = 'none';
						}else if(e.style.display == 'none'){
							e.style.display = 'table-row';
							ns.style.height = hfi + 80;	
						}
					}
					function toggle_headDetail(id){
						var button_On = document.getElementById("icon_hideOn");
						var button_Off = document.getElementById("icon_hideOff");
						var e = document.getElementById(id);
						if(e.style.display == 'block'){
							e.style.display = 'none';
							button_On.style.display = 'none';
							button_Off.style.display = 'block';
						}else if(e.style.display == 'none'){
							e.style.display = 'block';
							button_On.style.display = 'block';
							button_Off.style.display = 'none';
						}
					}
		</script>
		
	</head>

<!----------------------- ALL ABOUT TO ORACLE DATA QUERY  ------------------------------------------>

<%
dim v_QUANTITY,v_DOCDATE,v_PSTORAGE_STORAGECODE,NORMALSOUNDINGNUMBER,v_FFA_QUALITY,v_MOISTURE_QUALITY,v_DIRT_QUALITY,v_HIGHLEVELTANK,v_TEMPERATURE
dim AddNew

v_QUANTITY = request("QUANTITY")
v_DOCDATE = request("DOCDATE")
v_FFA_QUALITY = request("FFA_QUALITY")
v_MOISTURE_QUALITY = request("MOISTURE_QUALITY")
v_DIRT_QUALITY = request("DIRT_QUALITY")
v_HIGHLEVELTANK = request("HIGHLEVELTANK")
v_TEMPERATURE = request("TEMPERATURE")


if Len(trim(request("NORMALSOUNDINGNUMBER"))) > 0 then
NORMALSOUNDINGNUMBER = request("NORMALSOUNDINGNUMBER")
else
NORMALSOUNDINGNUMBER = "(NEW)"
end if 

if Len(trim(request("DOCNUMBER"))) > 0 then
v_DOCNUMBER = request("DOCNUMBER")
else
v_DOCNUMBER = "(NEW)"
end if 

if Len(trim(request("PSTORAGE_STORAGECODE"))) > 0 then
v_PSTORAGE_STORAGECODE = request("PSTORAGE_STORAGECODE")
else
v_PSTORAGE_STORAGECODE = ""
end if 

if Request("txttable") ="Sounding_Normal" then
	if Request("txtJenis")="Delete" then
	
	connstr.BeginTrans
		q = " delete from NORMALSOUNDING" & _
		    " where NORMALSOUNDINGNUMBER = '" & NORMALSOUNDINGNUMBER & "'" & _
			"       and PSTORAGE_STORAGECODE= '" & v_PSTORAGE_STORAGECODE & "'" & _
			"       and fcba = '" & Session("fcBA") & "'"
   	   'response.Write(q)
	   connStr.execute(q)
	   
		if err.number = 0 then
		   connStr.CommitTrans
		end if
		Response.Redirect "?txtModule=Mill_Production&txtTable=Sounding_Normal&txtJenis=view" 
	elseif Request("txtJenis")="New" then
			if NORMALSOUNDINGNUMBER = "(NEW)" then
				
					'auto number (NEW)
					dim rsAUTONUM,AutoNum,numq

					 Set rsAUTONUM = Server.CreateObject("ADODB.Recordset")
					 numq = "select nvl(controlnumber_api.get_nextnumber(upper('DOCCSNDNOR'), " & year(date) & ", " & month(date) & ", '" & session("fcBA") & "'), 'FALSE') prcode from dual "
					 'response.write(numq)
					 'response.write (numq)
					 rsAUTONUM.Open  numq, connstr	   
					if Not rsAUTONUM.Eof then
						AutoNum = rsAUTONUM("prcode")
					end if
								
					NORMALSOUNDINGNUMBER = AutoNum		
					rsAUTONUM.Close
					set rsAUTONUM = Nothing
				else
					NORMALSOUNDINGNUMBER = Request("DOCNUMBER")	
				end if
	
		connstr.BeginTrans	
		 q= "Insert into NORMALSOUNDING (NORMALSOUNDINGNUMBER, DOCDATE, PSTORAGE_STORAGECODE, QUANTITY, FFA_QUALITY, MOISTURE_QUALITY, DIRT_QUALITY, ROWSTATEID,DOCNUMBER, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME)" & _
				  "values ('" &  trim(NORMALSOUNDINGNUMBER) & "'" & _ 
				  ",to_date('" & Request("DOCDATE") & "','mm/dd/yyyy')" & _
				  ",'"  &  Request("PSTORAGE_STORAGECODE") & "'" & _
				  ",'" & Request("QUANTITY") & "'" & _
				  ",'" & Request("FFA_QUALITY") & "'" & _
				  ",'"  &  Request("MOISTURE_QUALITY") & "'" & _
				  ",'"  &  Request("DIRT_QUALITY") & "'" & _
				  ",'CLOSE'" & _
				  ",'" & Request("DOCNUMBER") & "'" & _
				  ",'" & Session("CodeName") & "'" & _
				  ",'" & Session("CodeName") & "'" & _
				  ",'" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
				  ",'" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI') )"
			''response.write(q)
			connStr.execute(q)
					
			q2 = " update DOCSOUNDING" & _
					" set NORMALSOUNDINGNUMBER = '" &  trim(NORMALSOUNDINGNUMBER) & "'" & _
					"     , fcentry = '" & Session("CodeName") & "'" & _
					"     , fcedit = '" & Session("CodeName") & "'" & _
					"     , fcip = '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _	
					"     , fcba = '" & Session("fcBA") & "'" & _	
					"     , lastupdate = sysdate" & _	
					"     , lasttime = to_char(sysdate,'hh24:mi')" & _						
					" where DOCNUMBER = '" & trim(v_DOCNUMBER) & "'" & _ 
					"       and fcba = '" & Session("fcBA") & "'"
	
				'response.Write(q2)
				connStr.execute(q2)
				
			q3 = " delete from productstoragereading" & _
		    " where readingdate = to_date('" & Request("DOCDATE") & "', 'mm/dd/yyyy')" & _
			"       and PSTORAGE_STORAGECODE= '" & Request("PSTORAGE_STORAGECODE") & "'" & _
			"       and fcba = '" & Session("fcBA") & "'"
		   'response.Write(q3)
		   connStr.execute(q3)	
				
			q4 = " insert into productstoragereading (" & _
						" readingdate, pstorage_storagecode, quantity, fcentry, fcedit" & _
						" , fcip, fcba, lastupdate, lasttime, ffa_quality" & _
						" , moisture_quality, dirt_quality, isautosounding, highleveltank, temperature,DOCNUMBER,NORMALSOUNDINGNUMBER" & _
						" ) values (" & _
						" to_date('" & Request("DOCDATE") & "', 'mm/dd/yyyy')" & _
						" , '" & Request("PSTORAGE_STORAGECODE") & "'" & _
						" , '" & Request("QUANTITY") & "'" & _
						" , '" & Session("CodeName") & "'" & _
						" , '" & Session("CodeName") & "'" & _
						" , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
						" , '" & Session("fcBA") & "'" & _
						" , sysdate" & _
						" , to_char(sysdate,'hh24:mi')" & _
						" , '" & Request("FFA_QUALITY") & "'" & _
						" , '" & Request("MOISTURE_QUALITY") & "'" & _
						" , '" & Request("DIRT_QUALITY") & "'" & _
						" , 'FALSE'" & _
						" , '" & Request("HIGHLEVELTANK") & "'" & _
						" , '" & Request("TEMPERATURE") & "'" & _
						" , '" & trim(v_DOCNUMBER) & "'" & _
						" , '" & trim(NORMALSOUNDINGNUMBER) & "'" & _
						" )"
					response.Write(q4)
					'response.Flush()
					connStr.execute(q4)	
			if err.number = 0 then
				connStr.CommitTrans
			end if
			'RESPONSE.REDIRECT "?txtModule=" & Request("txtModule") & "&txtTable=" & Request("txtTable") & "&txtJenis=view"
	elseif Request("txtJenis")="Update" then	 
		connstr.BeginTrans	
		
			q = " update NORMALSOUNDING" & _
					" set FFA_QUALITY = '" &  trim(Request("FFA_QUALITY")) & "'" & _
					"     , fcentry = '" & Session("CodeName") & "'" & _
					"     , fcedit = '" & Session("CodeName") & "'" & _
					"     , fcip = '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _	
					"     , fcba = '" & Session("fcBA") & "'" & _	
					"     , lastupdate = sysdate" & _	
					"     , lasttime = to_char(sysdate,'hh24:mi')" & _						
					" where NORMALSOUNDINGNUMBER = '" & trim(NORMALSOUNDINGNUMBER) & "'" & _ 
					"		and PSTORAGE_STORAGECODE = '" & v_PSTORAGE_STORAGECODE & "'" & _
					"       and fcba = '" & Session("fcBA") & "'"
			'response.Write(q)
			connStr.execute(q)	
			
			q2 = " update productstoragereading" & _
					" set FFA_QUALITY = '" &  trim(Request("FFA_QUALITY")) & "'" & _
					"     , fcentry = '" & Session("CodeName") & "'" & _
					"     , fcedit = '" & Session("CodeName") & "'" & _
					"     , fcip = '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _	
					"     , fcba = '" & Session("fcBA") & "'" & _	
					"     , lastupdate = sysdate" & _	
					"     , lasttime = to_char(sysdate,'hh24:mi')" & _						
					" where readingdate = to_date('" & Request("DOCDATE") & "', 'mm/dd/yyyy')" & _
					"       and PSTORAGE_STORAGECODE= '" & Request("PSTORAGE_STORAGECODE") & "'" & _
					"       and fcba = '" & Session("fcBA") & "'"
			'response.Write(q2)
			connStr.execute(q2)	
			
		if err.number = 0 then
		   connStr.CommitTrans
		end if
	
	end if

end if

if Request("txtModule") = "Mill_Production" then

Dim txtTable,txtModule,txtJenis,Searching,cmdView,crashId,qUpdateString
dim settingColum,settingHeightTable
settingHeightTable = 0
txtTable = Request("txtTable")
txtModule = Request("txtModule")
Searching=Request("Searching")
txtJenis=Request("txtJenis")


if Request("txttable") ="Sounding_Normal" then

	q	=	"SELECT " & _
			"NORMALSOUNDINGNUMBER, DOCDATE, PSTORAGE_STORAGECODE,QUANTITY,FFA_QUALITY,MOISTURE_QUALITY,DIRT_QUALITY," & _
			"ROWSTATEID,DOCNUMBER, FCENTRY, FCEDIT, FCIP,FCBA, LASTUPDATE,LASTTIME FROM NORMALSOUNDING "
			
	if  LEN(TRIM(Searching)) > 0 then
	q	=	q & "WHERE (UPPER(NORMALSOUNDINGNUMBER) LIKE UPPER('%"& trim(Searching)&"%') or " & _
            	"UPPER(PSTORAGE_STORAGECODE) LIKE UPPER('%"& Searching & "%') or " & _
				"UPPER(to_char(TO_DATE(DOCDATE), 'YYYY')) LIKE UPPER('%"& Searching & "%') or " & _
				"UPPER(to_char(TO_DATE(DOCDATE), 'MONTH')) LIKE UPPER('%"& Searching & "%') or " & _
				"UPPER(to_char(TO_DATE(DOCDATE), 'DAY')) LIKE UPPER('%" & Searching & "%')) and " & _
				"FCBA = '" & Session("fcBA") & "' "
		
		'response.write(q)
	else
	q	=	q & "WHERE FCBA = '" & Session("fcBA") & "' "
	end if
elseif Request("txttable") ="" then
	
else 
	response.write("No Data..")
	
end if
   Set rsGAD = Server.CreateObject("ADODB.Recordset")
   rsGAD.CursorType = 3
   rsGAD.CursorLocation = 3 
   rsGAD.Open q, connstr
   JumRec=1
   intRecordsPerPage = 13
	
	Dim intRecordsPerPage,intNumberOfPages,intRecords,nPage,howmanyfields,JumRec
	howmanyfields=rsGAD.fields.count-1
	
	JumRec=1
	intRecordsPerPage = 13
	
	rsGAD.PageSize = intRecordsPerPage
	intNumberOfPages = 0
	intNumberOfPages =  rsGAD.PageCount
	intRecords = rsGAD.RecordCount
	'Set our AbsolutePage property.
	if not rsGAD.EOF then
		If (isEmpty(Request("iPageNo")) or trim(Request("iPageNo"))="") Then
		  rsGAD.AbsolutePage = 1
		  
		Else
		  rsGAD.AbsolutePage = Request("iPageNo") 
		End If 
	else
		'response.write("<script language=""javascript"">alert('Data Not Found,');window.location.href='?txtModule=" & Request("txtModule") & "&txtTable=Sounding_Normal';</script>")
	end if
	if Request("iPageNo") = "" then
		nPage = 1
	Else
		nPage = Request("iPageNo")
	End if
	Dim cColor

   rsGAD.PageSize = intRecordsPerPage
   intNumberOfPages = 0
   intNumberOfPages =  rsGAD.PageCount
   intRecords = rsGAD.RecordCount   
	%>
<!----------------------- END SQL Edit or Update Data ------------------------------------------>

	<body>

		<!--#include file="top.jsp"-->
		<div id="lenght_menu" style=" width:21.7%; float:left;margin-left:0.3%; margin-top:0.3%; margin-bottom:0.3%;">
		 <% if Session("fcBa") <> "" then %>
			 <!--#include file="menu.asp"-->
		 <% end if %>
		</div>

	<div class="judul_head"  valign="top" bgcolor="#fff" style=" width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;"> 

		<div class="judul">
		   Module : <%=request("txtModule")%> / <%=request("txtTable")%>
		</div>

<!-- NEVER ON DELETE ---->
	<div style="width:200px; height:32px;">
		<!-- KOSONGKAN -->
	</div>
<!-- END NEVER ON DELETE -->	

<!-- DISPLAY PARENT LIST AND MENU TAB 
	 --------------------------------
	 Masukkan Data parent kedalam Comment Tag "Parent Data"------->
	 
	<%if Request("txttable") ="ProjectAccountingDetail" and not rsGAD.EOF then%>
		<div id="parentDetail">
			<div style="background-color:#0C5B24; padding:7px 3px;margin-top:2px;border-bottom:1px #000 solid;">
				<a class="head_link" href="?txtModule=Mill_Production&txtTable=Sounding_Normal" >&nbsp;
					<i class="fa fa-chevron-left"></i>&nbsp;&nbsp;Project Code&nbsp;:&nbsp;<%=rsGAD(6)%>
				</a>
				<a id="hide_headDetail" onclick="javascript:toggle_headDetail('data_headDetail','hide_headDetail')">
					<i id="icon_hideOn" style="display:block;" class="fa fa-chevron-up"></i>
					<i id="icon_hideOff" style="display:none;" class="fa fa-chevron-down"></i>
				</a>
			</div>
		<!-- Parent Data ------->
			<table id="data_headDetail" style="display:block;margin-left:10px;">
				<tr>
					<td>Name</td><td>:&nbsp;<%=rsGAD(7)%></td>
					<td rowspan="3" style="width:2px;"></td>
					<td rowspan="3" style="background-color:#fff; width:1px; padding:0px;"></td>
					<td rowspan="3" style="width:2px;"></td>
					<td>Project Type</td><td>:&nbsp;<%=rsGAD(9)%></td>
					<td rowspan="3" style="width:2px;"></td>
					<td rowspan="3" style="background-color:#fff; width:1px; padding:0px;"></td>
					<td rowspan="3" style="width:2px;"></td>
					<td>Valid Form</td><td>:&nbsp;<%=rsGAD(12)%></td>
					<td rowspan="3" style="width:2px;"></td>
					<td rowspan="3" style="background-color:#fff; width:1px; padding:0px;"></td>
					<td rowspan="3" style="width:2px;"></td>
					<td>Plan Start Date</td><td>:&nbsp;<%=rsGAD(10)%></td>
					<td rowspan="3" style="width:2px;"></td>
					<td rowspan="3" style="background-color:#fff; width:1px; padding:0px;"></td>
					<td rowspan="3" style="width:2px;"></td>
					<td>Project Group</td><td>:&nbsp;<%=rsGAD(8)%></td>
				</tr>
				<tr>
					<td>Project Manager</td><td>:&nbsp;<%=rsGAD(15)%></td>
					<td>Control Job</td><td>:&nbsp;<%=rsGAD(11)%></td>
					<td>Valid Until</td><td>:&nbsp;<%=rsGAD(13)%></td>
					<td>Remarks</td><td>:&nbsp;<%=rsGAD(16)%></td>
					<td>Row State Id</td><td>:&nbsp;<%=rsGAD(14)%></td>
				</tr>
			</table>
		<!-- END Parent Data ------->
		</div>
		
	<!-- MENU TAB ------>	
		<ul id="tabs">
			<li id="current"><a href="#" name="tab1">Project Accounting Detail</a></li>
		</ul>
	<!-- END MENU TAB -->
	
	<%elseif Request("txttable") ="ProjectAccountingDetail" and rsGAD.EOF then%>
		<div id="parentDetail">
			<div style="background-color:#0C5B24; padding:7px 3px;margin-top:2px;border-bottom:1px #000 solid;">
				<a class="head_link" href="?txtModule=Mill_Production&txtTable=Sounding_Normal" >&nbsp;
					<i class="fa fa-chevron-left"></i>&nbsp;&nbsp;Project Code&nbsp;:&nbsp;<%=request("txtIdCode")%>
				</a>
				<a id="hide_headDetail" onclick="javascript:toggle_headDetail('data_headDetail','hide_headDetail')">
					<i class="fa fa-ban"></i>
				</a>
			</div>
		</div>
	<%end if%>
<!-- END DISPLAY PARENT LIST AND MENU TAB ---->	
	
<!-- FLEXIGRID ----------------------------------->
		<table class="flexme">
				<thead>
					<tr>
					<% if Request("txttable") ="Sounding_Normal" then %>
						<th width="10"  align="center" style="width:10px;">No</th>
						<th width="120" align="left">Sounding Number</th>	
						<th width="100" align="left">Doc Date</th>
						<th width="100" align="left">Storage Code</th>
						<th width="100" align="left">Quantity</th>
						<th width="50" align="left">FFA</th>
						<th width="50" align="left">Moisture</th>
						<th width="50" align="left">Dirt</th>
						<th width="100" align="left">Rowsteid</th>
						<th width="100" align="left">Doc Sounding</th>
						<th width="30" align="center">Edit</th>
					<!--<th width="30" align="center">Delete</th>-->
					<%end if %>	
					</tr>
				</thead>
				<tbody>					
<% 	
	if Not rsGAD.eof then
      nCountofField=rsGAD.fields.count-1
	  if Request("iPageNo") = "" then
	  no_lanjutan = 1
	  else
	  no_lanjutan = 13 * Request("iPageNo") - 12
	  end if
	  no=no_lanjutan
				
	  While Not rsGAD.EOF 
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if

		  %>
		<tr bgcolor="<%=cColor%>">
			  <td align="left" valign="top"><%=NO%></td>
			  <%for i = 0 to nCountofField - (6 + settingColum)%>
					<td align="left" title="<%=rsGAD(i).Name%>" valign="top"><%=rsGAD(i)%></td>
				<%next%>
			  <td align="center" valign="top">
			  <a id="edit" title="Edit" class="pencil" href="sounding_normal.asp?txtModule=<%=request("txtModule")%>&txtTable=<%=request("txtTable")%>&NORMALSOUNDINGNUMBER=<%=rsGAD("NORMALSOUNDINGNUMBER")%>&DOCDATE=<%=rsGAD("DOCDATE")%>&PSTORAGE_STORAGECODE=<%=rsGAD("PSTORAGE_STORAGECODE")%>&QUANTITY=<%=rsGAD("QUANTITY")%>&FFA_QUALITY=<%=rsGAD("FFA_QUALITY")%>&MOISTURE_QUALITY=<%=rsGAD("MOISTURE_QUALITY")%>&DIRT_QUALITY=<%=rsGAD("DIRT_QUALITY")%>&DOCNUMBER=<%=rsGAD("DOCNUMBER")%>&HIGHLEVELTANK=<%=HIGHLEVELTANK %>&TEMPERATURE=<%=TEMPERATURE %>&formwindow=open&txtJenis=Edit&ViewSounding=True">
			  <i class="fa fa-pencil"></i></a>
			  </td>
		 <% if Request("txttable") ="Sounding_Normal" then %>
		 <!--	  <td align="center" valign="top">
			  <a class="pencil" title="Delete"
			  href="sounding_normal.asp?txtModule=<%=request("txtModule")%>&txtTable=<%=request("txtTable")%>&NORMALSOUNDINGNUMBER=<%=rsGAD("NORMALSOUNDINGNUMBER")%>&PSTORAGE_STORAGECODE=<%=rsGAD("PSTORAGE_STORAGECODE")%>&txtJenis=Delete">
			  <i class="fa fa-trash-o"></i></a>
			  </td>-->
		 <%end if %>	  
		</tr>
	
		 <%
			
			  rsGAD.MoveNext
			  intRecordsPerPage = intRecordsPerPage - 1
			  no=no+1
		  Wend
	   else%>
		<tr style="display:table-row;">
			<td>
			1
			</td>
			<td colspan="8" align="left">
			No Data..
			</td>
		</tr>
	<%   end if 
	   rsGAD.Close
	   set rsGAD = Nothing
    %>
   
		</tbody>
	</table>
	
<!-- END FLEXIGRID ----------------------------------->
	
	
	
<!-- MENU : ADDLIST / SEARCING / REFRESH ------------>

	<div style="background-color:#f1f1f1; padding:0.3%; border-top:1px #ccc solid; border-bottom:1px #ccc solid; position:fixed; top:114px; right:0.3%; width:76.8%;">
		<!-- BUTTON ADD LIST -->
	    <a id="ButtonOpenWindow" class="add_list" href="?txtModule=Mill_Production&txtTable=Sounding_Normal&txtJenis=Addlist&formwindow=open" value="Window - yellow"
		title="New :<%=txtTable%>"><i class="fa fa-plus-circle"></i><span>ADD LIST</span></a>
		
		<!-- BUTTON REFRESH --->
		<a href="#" class="add_list" onclick="reloadPage()"><i class="fa fa-refresh"></i><span>Refresh </span></a>&nbsp;
		
		<!-- SEARCHING -------->
		<form style="padding:0px; border:none;display:inline;margin:0px;" method="post" action="sounding_normal.asp?txtModule=<%=request("txtModule")%>&txtTable=<%=request("txtTable")%>&txtIdCode=<%=request("txtIdCode")%>"> 
		<input class="search" type="text" name="Searching" value="" placeholder="Search Data">
		<input class="button_search" type="submit" value=""> 
		<i class="fa fa-search icon_search_data"></i>
		</form>
		
<!-- END MENU : ADDLIST / SEARCING / REFRESH ------------>		


<!-- FORMAT PAGE NUMBER --------------------------------->	

	<div style="float:right; margin-left:10px;">
		<% dim sebelum, lanjut
		if intNumberOfPages < 2 then 
		
		else
		sebelum = int(Request("iPageNo")) - 1
		lanjut = int(Request("iPageNo")) + 1
		response.write "<font size=""2"" face=""Arial Narrow"" color=""#333"">Page :</font> "
		if int(Request("iPageNo")) = "" or int(Request("iPageNo")) = 1 then
		%>
		<span class="button_prev"><i class="fa fa-chevron-left"></i></span>
		<%else%>
		<a class="button_prev" href="process_app.asp?txtTable=<%=request("txtTable")%>&iPageNo=<%=sebelum%>"><i class="fa fa-chevron-left"></i></a>
		<%end if%>
		<%if Request("iPageNo") = "" then%>
				<%nPage = 1%>
		   <%Else%>
				<%nPage = Request("iPageNo")%>
		   <%End if
		  
		   if intNumberOfPages<>1 then
		   
			  For i = 1 to intNumberOfPages
					 if i = int(Request("iPageNo")) then
					 Response.Write "<a class=""button_page_current"">" & i & "</a> "  
					 else
					 Response.Write "<a class=""button_page"" href=""process_app.asp?txtTable="& request("txtTable")& "&iPageNo=" &  i & _
					  """>" & i & "</a> "
					 end if
			  Next
		   End if
		   
		   if int(Request("iPageNo")) = intNumberOfPages then %>
			<span class="button_next"><i class="fa fa-chevron-right"></i></span>
			<%else%>
			<a class="button_next" href="process_app.asp?txtTable=<%=request("txtTable")%>&iPageNo=<%=lanjut%>"><i class="fa fa-chevron-right"></i></a>
		<%  end if
		end if%>
			<span style="margin-top:15px;"><font size="1" color="#000000" face="Arial, Helvetica, sans-serif">Total Records : <%=intRecords%> </font></span>
	</div>
<!-- END FORMAT PAGE NUMBER ---------------------------------->	  	
</div>


<!------- BOTTOM ----------->

	<!--#include file="bottom.jsp"-->

<!------- END BOTTOM ------->

<%end if%>


 
<!--TEMPLATE ADD LIST FORM 
	-----------------------
	Masukkan format inputan kedalam comment tag "INPUT FORM ADD LIST " ------>
	<%dim display__window,open_window 
	
		open_window = request("formwindow")
	
		if open_window = "open" then
		display__window = "display:table-row;"
		else
		display__window = "display:none;"
		end if
		
	dim title_window,height_window	
	if Len(Request("txtGetSounding")) = 0 then
		title_window = "Add Sounding List"
	elseif 	Request("txtGetSounding") = "GetSounding1st" then
		title_window = "Get Sounding"
	else
		title_window = "Add list"
	end if
	if  Len(Request("Height_form")) > 0 then 
		height_window = Request("Height_form")
	else
		height_window = "420px"
	end if
	if txtJenis = "Edit" then
		read_only_edit = "readonly='True' style='background:#FFFFCC'"

	else
		read_only_edit = ""

	end if
	%>
	<div id="form_window" class="nswindowContainer" style="<%=display__window%> z-index: 1010; background-color:rgba(0,0,0,0.4);">
		<div id="nswindow" class="yellow" style="width: 800px; height:<%=height_window%>; left: 270px; top: 25px; opacity: 1;">
			<div id="ctrlbar" class="controlbar">
				<div class="title" style="margin-top: 4.5px;"><%=title_window%></div>
				<a id="ButtonCloseWindow" class="close" href="?txtModule=Mill_Production&txtTable=Sounding_Normal&txtJenis=View" title="close">Close</a>
			</div>
			<div id="nswindowContent" style="background:#f1f1f1;top: 26px;">
			
			<!-- INPUT FORM ADD LIST --------------------------------->		
			
			<% if Request("txttable") ="Sounding_Normal" and Len(Request("txtGetSounding")) = 0 then %>
				<div id="WARP_input" style="height:340px;">
				<script language="javascript">


	function ValidasiSoundingNormal(form){
		if (form.NORMALSOUNDINGNUMBER.value == 0){

		alert("Please enter Normal Sounding Number");

		form.NORMALSOUNDINGNUMBER.focus();

		return (false);

		}
		
		if (form.DOCDATE.value == 0){

		alert("Please enter Document Number");

		form.DOCDATE.focus();

		return (false);

		}
		if (form.FFA_QUALITY.value == 0){

		alert("Please enter FFA");

		form.FFA_QUALITY.focus();

		return (false);

		}
		if (form.MOISTURE_QUALITY.value == 0){

		alert("Please enter Moisture");

		form.MOISTURE_QUALITY.focus();

		return (false);

		}
		if (form.DIRT_QUALITY.value == 0){

		alert("Please enter Dirt");

		form.DIRT_QUALITY.focus();

		return (false);

		}
		if (form.QUANTITY.value == 0){

		alert("Please enter Get Sounding");

		form.QUANTITY.focus();

		return (false);

		}

	  return (true);

	}
	
	
</script>
<% 

if Request("txtJenis") = "Addlist" or Request("txtJenis") = "" then 
AddNew = "New"
elseif Request("txtJenis") = "Edit" then
AddNew = "Update"
end if

%> 
			<form class="form_input" onSubmit="javascript:return ValidasiSoundingNormal(this)" style="background:#ffffff;" method="post" action="sounding_normal.asp?txtModule=<%=Request("txtModule")%>&txtTable=<%=Request("txtTable")%>&txtJenis=<%=AddNew%>">
				
				<table id="table_Input" style=" float:left;">
					<tr>
						<td>Doc Number Normal</td>
						<td><input id="NORMALSOUNDINGNUMBER" type="text" maxlength="50" name="NORMALSOUNDINGNUMBER" size="30" value="<%=NORMALSOUNDINGNUMBER%>"  <%=read_only_edit%>><font color="#FF0000">&nbsp;&nbsp;*</font></td>
					</tr>
					<tr>
						<td>Doc Date</td>
						<td><input type="text" class="datepicker" name="DOCDATE" placeholder="Date" value="<%=v_DOCDATE%>" <%=read_only_edit%>></td>
					</tr>
					<tr>
						<td>Doc Number</td>
						<td>
						
						<% IF v_DOCNUMBER = "(NEW)" THEN %>
						<input type="button" value="GET SOUNDING" onClick="
						location.href='sounding_normal.asp?txtModule=Mill_Production&txtTable=Sounding_Normal&txtGetSounding=GetSounding1st&NORMALSOUNDINGNUMBER=(NEW)&formwindow=open'
						">
						<% ELSE 
							if txtJenis = "Edit" then%>
							<input type="button" value="<%=v_DOCNUMBER%>" onClick="
							location.href='sounding_normal.asp?txtModule=Mill_Production&txtTable=Sounding_Normal&txtGetSounding=GetSounding1st&NORMALSOUNDINGNUMBER=<%=NORMALSOUNDINGNUMBER%>&DOCDATE=<%=v_DOCDATE%>&PSTORAGE_STORAGECODE=<%=v_PSTORAGE_STORAGECODE%>&FFA_QUALITY=<%=v_FFA_QUALITY%>&MOISTURE_QUALITY=<%=v_MOISTURE_QUALITY%>&DIRT_QUALITY=<%=v_DIRT_QUALITY%>&DOCNUMBER=<%=v_DOCNUMBER%>&formwindow=open&txtJenis=Edit'" disabled>
							<% else %>
							<input type="button" value="<%=v_DOCNUMBER%>" onClick="
							location.href='sounding_normal.asp?txtModule=Mill_Production&txtTable=Sounding_Normal&txtGetSounding=GetSounding1st&DOCNUMBER=<%=v_DOCNUMBER%>&formwindow=open'">
						<%  end if
						END IF %>
						
						</td>
						
					</tr>
					
					<tr>
						<td>Quantity</td>
						<td>
						<input id="QUANTITY" type="text" maxlength="50" name="QUANTITY" size="30" value="<%=v_QUANTITY%>" readonly='True' style='background:#FFFFCC'>
						<input id="DOCNUMBER" type="hidden" name="DOCNUMBER" value="<%=v_DOCNUMBER%>">
						<input id="PSTORAGE_STORAGECODE" type="hidden" name="PSTORAGE_STORAGECODE" value="<%=v_PSTORAGE_STORAGECODE%>">
						<input id="HIGHLEVELTANK" type="hidden" name="HIGHLEVELTANK" value="<%=REQUEST("HIGHLEVELTANK")%>">
						<input id="TEMPERATURE" type="hidden" name="TEMPERATURE" value="<%=REQUEST("TEMPERATURE")%>">
						</td>
					
					</tr>
					<tr>
						<td>FFA</td>
						<td><input type="text" name="FFA_QUALITY" placeholder="FFA Quality" value="<%=v_FFA_QUALITY%>"></td>
					</tr>
					<tr>
						<td>Moisture</td>
						<td><input type="text" name="MOISTURE_QUALITY" placeholder="Moisture Quality" value="<%=v_MOISTURE_QUALITY%>" <%=read_only_edit%>></td>
					</tr>
					<tr>
						<td>Dirt</td>
						<td><input type="text" name="DIRT_QUALITY" placeholder="Dirt Quality" value="<%=v_DIRT_QUALITY%>" <%=read_only_edit%>></td>
					</tr>
		
				</table>
				<% if request("ViewSounding") = "True" then %>	
					<div style="margin-top:10px;margin-left:20px;padding:10px;min-width:150px;float:left;border:1px #ccc solid;box-shadow:2px 2px 0px rgba(0,0,0,0.6);">
					<% if Len(trim(v_HIGHLEVELTANK)) = 0 or Len(trim(v_TEMPERATURE)) = 0 then
					
						Set rsht = Server.CreateObject("ADODB.Recordset")
						qht = "select HIGHLEVELTANK, TEMPERATURE from DOCSOUNDINGDETAIL where DOCNUMBER = '" & v_DOCNUMBER & "'"
						 rsht.Open  qht, connstr	   
						if Not rsht.Eof then
							v_HIGHLEVELTANK = rsht("HIGHLEVELTANK")
							v_TEMPERATURE = rsht("TEMPERATURE")
						end if
						rsht.Close
						set rsht = Nothing
						
						end if
					%>
					<table id="table_Input">
						<tr>
							<td colspan="2"><b>Sounding</b><hr/></td>
						</tr>

						<tr>
							<td>High level Tank :</td>
							<td align="right"><%=v_HIGHLEVELTANK%> Cm</td>
						</tr>
						<tr>
							<td>Temperature :</td>
							<td align="right"><%=v_TEMPERATURE%> &#186;</td>
						</tr>
						<tr>
							<td>Quantity :</td>
							<td align="right"><%=v_QUANTITY%> Kg</td>
						</tr>
					</table>
					</div>
				<% end if%>	
				<div style="clear:both;"></div>
				<div style="position:absolute; left:10px; bottom:10px;">
				
				<% if Request("txtJenis")="Edit" then %>
				<input type="submit" name="submit" value="Update">
				<% else %>
				<input type="submit" name="submit" value="Add List">
				<% end if %>
				<input id="ButtonCloseWindow" type="button" onclick="location.href='?txtModule=Mill_Production&txtTable=Sounding_Normal&txtJenis=View'" value="Cancel">
				</div>
				</form>
			<% elseif Request("txttable") ="Sounding_Normal" and Request("txtGetSounding") = "GetSounding1st" then %>	
				<!--#include file="sounding_form.asp"-->
				
			<% end if%>

			<!-- END INPUT FORM ADD LIST --------------------------------->	

			</div>	
		</div>
	</div>
<!-- END TEMPLATE ADD LIST FORM --------------------------------------->


<!-- JAVASCRIPT STILL PUT THIS PAGE IS UNDER -------------------------->
	<script type="text/javascript">
		//SET FLEXIBLE HEIGHT FOR FLEXIGRID(TABLE)
		var h_menu = $(".body_menu").height();
		var h_HeadParent = $("#parentDetail").height();
		var h_TabMenu = $("#tabs").height();
		var result = h_menu - (41 + h_HeadParent + h_TabMenu);
		//GET FLEXIGRID(TABLE)
        $('.flexme').flexigrid({
			height:result
		});

    </script>
<!-- END JAVASCRIPT STILL PUT THIS PAGE IS UNDER ----------------------->		
	</body>
</html>

