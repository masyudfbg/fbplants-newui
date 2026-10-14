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
		
		<script language="javascript">
			function ControlJob(str,No)
				{
				var grpcode = document.getElementById("PROJECTGROUPCODE"+No).value;
				
				if 	(grpcode == ""){
					document.getElementById("controljob"+No).value="";
					}else{
						if (str.length==0)
						  {
						  document.getElementById("controljob"+No).value="0";
						  return;
						  }
						  
						if (window.XMLHttpRequest)
						  {// code for IE7+, Firefox, Chrome, Opera, Safari
						  xmlhttp=new XMLHttpRequest();
						  }
						else
						  {// code for IE6, IE5
						  xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");
						  }
						  
						xmlhttp.onreadystatechange=function()
						  {
						  if (xmlhttp.readyState==4 && xmlhttp.status==200)
							{
							document.getElementById("controljob"+No).value=xmlhttp.responseText;
							}
						  }
						xmlhttp.open("GET","ajax/ProjectAccJob_ajx.asp?initialFcba=<%=session("FCBA")%>&PROJECTGROUPCODE="+str,true);
						xmlhttp.send();
					}
					
					
					
				}
				
</script>
	</head>

<!----------------------- ALL ABOUT TO ORACLE DATA QUERY  ------------------------------------------>

<%

if Request("txttable") ="ProjectAccounting" then
	if Request("txtJenis")="Delete" then

	qCheckData = " select count(*) totalData from PROJECTACCDETAIL " & _
					  " where " & _
					  " PROJECTCODE = '" & Request("txtIdCode") &  "' and " & _ 
					  " FCBA = '" & Session("fcBA") & "'"
					  
		 Set rsCheckData = Server.CreateObject("ADODB.Recordset")
		   rsCheckData.CursorType = 3
		   rsCheckData.CursorLocation = 3 
		   rsCheckData.Open qCheckData, connstr	

		 if not rsCheckData.EOF then
		     totalData = cdbl(rsCheckData("totalData"))			  
		 else
		     totalData = 0
		 end if
		   
		   
		 rsCheckData.Close
		 set rsCheckData = Nothing		   

		 if totalData = 0 then 	
			 q="delete from PROJECTACC where FCCODE = '" & Request("txtIdCode")  & "' AND FCBA = '" & Session("fcBA") & "'"
			 'response.write("test" & q)
			 connStr.execute(q)
			 RESPONSE.REDIRECT "?txtModule=" & Request("txtModule") & "&txtTable=" & Request("txtTable") & ""
	     else%>
		 <script language="javascript">
		 var r=confirm('This records contained detailed data, if you want to delete all?');
		 if (r==true){
		 window.location.href='?txtModule=<%=Request("txtModule")%>&txtTable=<%=Request("txtTable")%>&txtIdCode=<%=Request("txtIdCode")%>&txtJenis=DeleteAll';
		 }else{
		 window.location.href='?txtModule=<%=Request("txtModule")%>&txtTable=<%=Request("txtTable")%>';
		 }
		 </script>
		 <%end if
	elseif Request("txtJenis")="DeleteAll" then	
		q="delete from PROJECTACCDETAIL where PROJECTCODE = '" & Request("txtIdCode")  & "' AND FCBA = '" & Session("fcBA") & "'"
		connStr.execute(q)
		
		q="delete from PROJECTACC where FCCODE = '" & Request("txtIdCode")  & "' AND FCBA = '" & Session("fcBA") & "'"
		connStr.execute(q)
		
		RESPONSE.REDIRECT "?txtModule=" & Request("txtModule") & "&txtTable=" & Request("txtTable") & ""
		 
	elseif Request("txtJenis")="New" then
		 q= "Insert into PROJECTACC (FCCODE, FCNAME, PROJECTGROUPCODE, PROJECTTYPECODE, PLANSTARTDATE, BUSINESSUNITREQUESTOR, PROJECTQTYREQUEST, CONTROLJOBCODE, VALID_FROM, VALID_UNTIL, ROWSTATEID,PROJECTMANAGER,REMARKS, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME)" & _
				  "values ('" &  trim(Request("FCCODE")) & "'" & _ 
				  ",'" & Request("FCNAME") & "'" & _
				  ",'"  &  Request("PROJECTGROUPCODE") & "'" & _
				  ",'" & Request("PROJECTTYPECODE") & "'" & _
				  ",to_date('" & Request("PLANSTARTDATE") & "','mm/dd/yyyy')" & _
				  ",'"  &  Request("CUSTOMER") & "'" & _
				  ",'"  &  Request("QUANTITY") & "'" & _
				  ",'" & Request("CONTROLJOBCODE") & "'" & _
				  ",to_date('" & Request("VALID_FROM") & "','mm/dd/yyyy')" & _
				  ",to_date('" & Request("VALID_UNTIL") & "','mm/dd/yyyy')" & _
				  ",'" & Request("ROWSTATEID")  & "'" & _
				  ",'" & Request("PROJECTMANAGER")  & "'" & _
				  ",'" & Request("REMARKS") & "'" & _
				  ",'" & Session("CodeName") & "'" & _
				  ",'" & Session("CodeName") & "'" & _
				  ",'" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
				  ",'" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI') )"
				  'response.write(q)
			connStr.execute(q)
			RESPONSE.REDIRECT "?txtModule=" & Request("txtModule") & "&txtTable=" & Request("txtTable") & ""
	elseif Request("txtJenis")="Edit" then	 
	DIM dVALID_FROM,dVALID_UNTIL,dPLANSTARTDATE
		IF LEN(TRIM(Request("MONTH_VALID_FROM") )) = 0 or LEN(TRIM(Request("DAY_VALID_FROM") )) = 0 or LEN(TRIM(Request("YEAR_VALID_FROM") )) = 0 THEN
			dVALID_FROM = ""
		ELSE
			dVALID_FROM = "VALID_FROM = to_date('" & Request("MONTH_VALID_FROM") & "/" & Request("DAY_VALID_FROM") & "/" & Request("YEAR_VALID_FROM") & "','mm/dd/yyyy'),"
		END IF
		IF LEN(TRIM(Request("MONTH_VALID_UNTIL") )) = 0 or LEN(TRIM(Request("DAY_VALID_UNTIL") )) = 0 or LEN(TRIM(Request("YEAR_VALID_UNTIL") )) = 0  THEN
			dVALID_UNTIL = ""
		ELSE
			dVALID_UNTIL = "VALID_UNTIL = to_date('" & Request("MONTH_VALID_UNTIL") & "/" & Request("DAY_VALID_UNTIL") & "/" & Request("YEAR_VALID_UNTIL") & "','mm/dd/yyyy'), "
		END IF
		IF LEN(TRIM(Request("MONTH_PLANSTARTDATE") )) = 0 or LEN(TRIM(Request("DAY_PLANSTARTDATE") )) = 0 or LEN(TRIM(Request("YEAR_PLANSTARTDATE") )) = 0  THEN
			dPLANSTARTDATE = ""
		ELSE
			dPLANSTARTDATE = "PLANSTARTDATE = to_date('" & Request("MONTH_PLANSTARTDATE") & "/" & Request("DAY_PLANSTARTDATE") & "/" & Request("YEAR_PLANSTARTDATE") & "','mm/dd/yyyy'), "
		END IF
		 q= "UPDATE PROJECTACC SET " & _
			"FCNAME = '" & Request("FCNAME") & "', " & _
			"PROJECTGROUPCODE = '" & Request("PROJECTGROUPCODE") & "', " & _
			"PROJECTTYPECODE = '" & Request("PROJECTTYPECODE") & "', " & _
			"" & dPLANSTARTDATE & "" & _
			"BUSINESSUNITREQUESTOR = '" & Request("CUSTOMER") &"'," & _
			"PROJECTQTYREQUEST = '" & Request("QUANTITY") &"'," & _
			"CONTROLJOBCODE = '" & Request("CONTROLJOBCODE") & "', " & _
			"" & dVALID_FROM & "" & _
			"" & dVALID_UNTIL & "" & _
			"ROWSTATEID 	= '" & Request("ROWSTATEID") & "', " & _
			"PROJECTMANAGER = '" & Request("PROJECTMANAGER") & "', " & _
			"REMARKS 		= '" & Request("REMARKS") & "', " & _ 
			"FCENTRY 		= '" & Session("CodeName") & "', " & _
			"FCEDIT 		= '" & Session("CodeName") & "', " & _
			"FCIP 			= '" & Request.ServerVariables("REMOTE_ADDR") & "', " & _
			"FCBA 			= '" & Session("FCBA") & "', " & _
			"LASTUPDATE 	= sysdate, " & _
			"LASTTIME 		= To_Char(sysdate,'HH24:MI') " & _
			"WHERE FCCODE 	= '" & Request("FCCODE") & "' AND FCBA = '" & Session("FCBA") & "'"
		 response.write (q)  
			connStr.execute(q)
			
			if err.number <> 0 then
			   response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & err.description & "</font>"
			End if
			
			RESPONSE.REDIRECT "?txtModule=" & Request("txtModule") & "&txtTable=" & Request("txtTable") & ""
	end if
elseif Request("txttable") ="ProjectAccountingDetail" then
	if Request("txtJenis")="Delete" then
	
		 q="delete from PROJECTACCDETAIL where PROJECTCODE = '" & Request("txtIdCode")  & "' AND LINENO = '" & Request("LINENO")  & "' AND FCBA = '" & Session("fcBA") & "'"
		 connStr.execute(q)
		 RESPONSE.REDIRECT "?txtModule=" & Request("txtModule") & "&txtTable=" & Request("txtTable") & "&txtIdCode=" & Request("txtIdCode") & ""
		 
	elseif Request("txtJenis")="New" then
	'SELECT LINENO
		 q_lineno = "select nvl(MAX(LINENO),0)+1 LINENO from PROJECTACCDETAIL where fcba = '" & session("fcba") & "'"
		  Set rsLineno = Server.CreateObject("ADODB.Recordset")
		  rsLineno.Open q_lineno, connstr
		     'INSERT PROJECTACCDETAIL
			 q= "Insert into PROJECTACCDETAIL (PROJECTCODE, LINENO, JOBREFCODE, VALID_INPUT_FROM, VALID_INPUT_UNTIL,REMARKS, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME)" & _
					  "values ('" &  trim(Request("PROJECTCODE")) & "','" & rsLineno("LINENO") & "','"  &  Request("JOBREFCODE") & "'," & _
					  "to_date('" & Request("VALID_INPUT_FROM") & "','mm/dd/yyyy')," & _
					  "to_date('" & Request("VALID_INPUT_UNTIL") & "','mm/dd/yyyy'),'" & _
					  Request("REMARKS") & "','" & _
					  Session("CodeName") & "','" & Session("CodeName") & "','" & _
					  Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI') )"
				
				connStr.execute(q)
		  rsLineno.Close
		  set rsLineno = Nothing
			RESPONSE.REDIRECT "?txtModule=" & Request("txtModule") & "&txtTable=" & Request("txtTable") & "&txtIdCode=" & Request("PROJECTCODE") & ""
	elseif Request("txtJenis")="Edit" then
		 
		 q= "UPDATE PROJECTACCDETAIL SET " & _
			"JOBREFCODE = '" & Request("JOBREFCODE") & "', " & _
			"VALID_INPUT_FROM = to_date('" & Request("MONTH_VALID_INPUT_FROM") & "/" & Request("DAY_VALID_INPUT_FROM") & "/" & Request("YEAR_VALID_INPUT_FROM") & "','mm/dd/yyyy'), " & _
			"VALID_INPUT_UNTIL = to_date('" & Request("MONTH_VALID_INPUT_UNTIL") & "/" & Request("DAY_VALID_INPUT_UNTIL") & "/" & Request("YEAR_VALID_INPUT_UNTIL") & "','mm/dd/yyyy'), " & _
			"REMARKS 	= '" & Request("REMARKS") & "', " & _ 
			"FCENTRY 	= '" & Session("CodeName") & "', " & _
			"FCEDIT 	= '" & Session("CodeName") & "', " & _
			"FCIP 		= '" & Request.ServerVariables("REMOTE_ADDR") & "', " & _
			"FCBA 		= '" & Session("FCBA") & "', " & _
			"LASTUPDATE = sysdate, " & _
			"LASTTIME 	= To_Char(sysdate,'HH24:MI') " & _
			"WHERE PROJECTCODE = '" & Request("PROJECTCODE") & "' AND LINENO = '" & Request("LINENO") & "' AND FCBA = '" & Session("FCBA") & "'"
			connStr.execute(q)
			
			if err.number <> 0 then
			   response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & err.description & "</font>"
			End if
			
			RESPONSE.REDIRECT "?txtModule=" & Request("txtModule") & "&txtTable=" & Request("txtTable") & "&txtIdCode=" & Request("PROJECTCODE") & ""
	end if


end if

if Request("txtModule") = "Project" then

Dim txtTable,txtModule,txtJenis,Searching,cmdView,crashId,qUpdateString
dim settingColum,settingHeightTable
settingHeightTable = 0
txtTable = Request("txtTable")
txtModule = Request("txtModule")
Searching=Request("Searching")
txtJenis=Request("txtJenis")


if Request("txttable") ="ProjectAccounting" then

	q	=	"SELECT " & _
			"FCCODE, FCNAME, PROJECTGROUPCODE,PROJECTTYPECODE,PLANSTARTDATE,BUSINESSUNITREQUESTOR,PROJECTQTYREQUEST,CONTROLJOBCODE,VALID_FROM,VALID_UNTIL,ROWSTATEID, " & _
			"PROJECTMANAGER,REMARKS,JOB_API.get_jobcode_name(CONTROLJOBCODE, FCBA) CONTROLJOBNAME, FCENTRY, FCEDIT, FCIP,FCBA, LASTUPDATE,LASTTIME FROM PROJECTACC "
			
	if  LEN(TRIM(Searching)) > 0 then
	q	=	q & "WHERE (UPPER(FCCODE) LIKE UPPER('%"& trim(Searching)&"%') or " & _
            	"UPPER(FCNAME) LIKE UPPER('%"& Searching & "%') or " & _
				"UPPER(PROJECTGROUPCODE) LIKE UPPER('%"& Searching & "%') or " & _
				"UPPER(PROJECTTYPECODE) LIKE UPPER('%"& Searching & "%') or " & _
				"UPPER(to_char(TO_DATE(PLANSTARTDATE), 'YYYY')) LIKE UPPER('%"& Searching & "%') or " & _
				"UPPER(to_char(TO_DATE(PLANSTARTDATE), 'MONTH')) LIKE UPPER('%"& Searching & "%') or " & _
				"UPPER(to_char(TO_DATE(PLANSTARTDATE), 'DAY')) LIKE UPPER('%" & Searching & "%')) and " & _
				"FCBA = '" & Session("fcBA") & "' "
		
		'response.write(q)
	else
	q	=	q & "WHERE FCBA = '" & Session("fcBA") & "' "
	end if
elseif Request("txttable") ="ProjectAccountingDetail" then
	
    settingColum = 5
	settingHeightTable = 85
	q	=	"SELECT " & _
			"e1.PROJECTCODE,e1.LINENO,e1.JOBREFCODE,e1.VALID_INPUT_FROM,e1.VALID_INPUT_UNTIL, " & _
			"e1.REMARKS,e2.FCCODE,e2.FCNAME,e2.PROJECTGROUPCODE,e2.PROJECTTYPECODE,e2.PLANSTARTDATE,e2.CONTROLJOBCODE," & _
			"e2.VALID_FROM,e2.VALID_UNTIL,e2.ROWSTATEID,e2.PROJECTMANAGER,e2.REMARKS,JOB_API.get_jobcode_name(e1.JOBREFCODE, e1.FCBA) CONTROLJOBNAME " & _
			"FROM PROJECTACCDETAIL e1,PROJECTACC e2 "
			
	if  LEN(TRIM(Searching)) > 0 then
	q	=	q & "WHERE e1.PROJECTCODE = '" & REQUEST("txtIdCode")& "' and e1.PROJECTCODE=e2.FCCODE AND e1.FCBA = e2.FCBA and " & _
				"(e1.LINENO = '" & Searching & "' or " & _
				"UPPER(e1.JOBREFCODE) LIKE UPPER('%"& Searching & "%') or " & _
				"UPPER(to_char(TO_DATE(e1.VALID_INPUT_FROM), 'YYYY')) LIKE UPPER('%"& Searching & "%') or " & _
				"UPPER(to_char(TO_DATE(e1.VALID_INPUT_FROM), 'MONTH')) LIKE UPPER('%"& Searching & "%') or " & _
				"UPPER(to_char(TO_DATE(e1.VALID_INPUT_FROM), 'DAY')) LIKE UPPER('%" & Searching & "%'))" 
		
		'response.write(q)
	else
	q	=	q & "WHERE e1.PROJECTCODE = '" & REQUEST("txtIdCode")& "' AND e1.PROJECTCODE=e2.FCCODE AND e1.FCBA = e2.FCBA "
	end if		
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
		'response.write("<script language=""javascript"">alert('Data Not Found,');window.location.href='?txtModule=" & Request("txtModule") & "&txtTable=ProjectAccounting';</script>")
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
				<a class="head_link" href="?txtModule=Project&txtTable=ProjectAccounting" >&nbsp;
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
				<a class="head_link" href="?txtModule=Project&txtTable=ProjectAccounting" >&nbsp;
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
					<% if Request("txttable") ="ProjectAccounting" then %>
						<th width="10"  align="center" style="width:10px;">No</th>
						<th width="70" style="width:100px;" align="left">Project Code</th>	
						<th width="100" align="left">Name</th>
						<th width="100" align="left">Project Group</th>
						<th width="100" align="left">Project Type</th>
						<th width="50" align="left">Plan Start Date</th>
						<th width="50" align="left">Customer</th>
						<th width="50" align="left">Quantity</th>
						<th width="100" align="left">Control Job</th>
						<th width="50" align="left">Valid Form</th>
						<th width="50" align="left">Valid Until</th>
						<th width="100" align="left">Row State Id</th>
						<th width="100" align="left">Project Manager</th>
						<th width="100" align="left">Remarks</th>
						<th width="30" align="center">Edit</th>
						<th width="30" align="center">Delete</th>
						<th width="30" align="center">Detail</th>
					<%elseif Request("txttable") ="ProjectAccountingDetail" then%>
						<th width="10"  align="center" style="width:10px;">No</th>
						<th width="100" style="width:100px;" align="left">Project Code</th>	
						<th width="100" align="left">Line Number</th>
						<th width="100" align="left">Job Ref</th>
						<th width="100" align="left">Valid Input From</th>
						<th width="100" align="left">Valid Input Until</th>
						<th width="150" align="left">Remarks</th>
						<th width="30" align="center">Edit</th>
						<th width="30" align="center">Delete</th>
					<%end if %>	
					</tr>
				</thead>
				<tbody>					
<% 	
	DIM PROJECTGROUPCODE
	if Not rsGAD.eof then
	  PROJECTGROUPCODE = rsGAD("PROJECTGROUPCODE")
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
			  <%
		

			  for i = 0 to nCountofField - (7 + settingColum)
					if i = 2 and Request("txttable") ="ProjectAccountingDetail" then %>
					<td align="left" title="<%=rsGAD(i).Name%>" valign="top"><%=rsGAD(i)%>&nbsp;:&nbsp;<%=rsGAD("CONTROLJOBNAME")%> </td>
					<%elseif i = 7 and Request("txttable") ="ProjectAccounting" then %>
					<td align="left" title="<%=rsGAD(i).Name%>" valign="top"><%=rsGAD(i)%>&nbsp;:&nbsp;<%=rsGAD("CONTROLJOBNAME")%> </td>
					<%else%>
					<td align="left" title="<%=rsGAD(i).Name%>" valign="top"><%=rsGAD(i)%></td>
					<%end if%>
			  <%next%>
	
			  <td align="center" valign="top">
			  <a id="edit<%=NO%>" title="Edit" class="pencil" onclick="javascript:toggle('editing_form_<%=NO%>','edit<%=NO%>')">
			  <i class="fa fa-pencil"></i></a>
			  </td>
		<% if Request("txttable") ="ProjectAccounting" then %>
			  <td align="center" valign="top">
			  <a class="pencil" title="Delete" onclick="return confirm('Are you sure you want to Delete this record?');"
			  href="ProjectAccounting.asp?txtModule=<%=request("txtModule")%>&txtTable=<%=request("txtTable")%>&txtIdCode=<%=rsGAD("FCCODE")%>&txtJenis=Delete">
			  <i class="fa fa-trash-o"></i></a>
			  </td>
		
			  <td align="center" valign="top">
			  <a id="detail_acc" class="pencil" title="Detail"
			  href="ProjectAccounting.asp?txtModule=<%=request("txtModule")%>&txtTable=<%=request("txtTable")%>Detail&txtIdCode=<%=rsGAD("FCCODE")%>&PROJECTGROUPCODE=<%=rsGAD("PROJECTGROUPCODE")%>">
			  <i class="fa fa-list"></i>
			  </a>
			  </td>
		<%elseif Request("txttable") ="ProjectAccountingDetail" then%>
			  <td align="center" valign="top">
			  <a class="pencil" title="Delete" onclick="return confirm('Are you sure you want to Delete this record?');"
			  href="ProjectAccounting.asp?txtModule=<%=request("txtModule")%>&txtTable=<%=request("txtTable")%>&txtIdCode=<%=rsGAD("PROJECTCODE")%>&LINENO=<%=rsGAD("LINENO")%>&txtJenis=Delete">
			  <i class="fa fa-trash-o"></i></a>
			  </td>
			  
		    <%end if %>	  
		  </tr>
			  
	<!-- FORM EDIT / ROW ---->
		<% if Request("txttable") ="ProjectAccounting" then %>
		
			<tr id="editing_form_<%=NO%>" style="display:none;">
			  <td class="arrow"><br/><br/><i class="fa fa-reply"></i></td>
			  <td colspan="<%=nCountofField%>" align="left">
			 <form method="post" action="ProjectAccounting.asp?txtModule=<%=request("txtModule")%>&txtTable=<%=request("txtTable")%>&txtJenis=Edit" style="border:none;padding:0px;"> 
			 
			  <div style=" width:100%; display: inline-block;">
				  <div class="edit_form">
					<div class="edit_mark">Project Code :</div>
					<div>
						<input type="text" value="<%=rsGAD("FCCODE")%>" disabled>
						<input type="hidden" name="FCCODE" value="<%=rsGAD("FCCODE")%>" placeholder="Code" >
					</div>
				  </div>
				  <div class="edit_form">
					<div class="edit_mark">Plan Start Date :</div>
					<div><%dateCallComboUpdate "MONTH_PLANSTARTDATE","DAY_PLANSTARTDATE","YEAR_PLANSTARTDATE",rsGAD("PLANSTARTDATE")%></div>
				   </div>
				   
				  <div class="edit_form">
					<div class="edit_mark">Valid Form :</div>
					<div><%dateCallComboUpdate "MONTH_VALID_FROM","DAY_VALID_FROM","YEAR_VALID_FROM",rsGAD("VALID_FROM")%></div>
				  </div>
				  <div class="edit_form">
					<div class="edit_mark">Remarks :</div>
					<div><input type="text" name="REMARKS" value="<%=rsGAD("REMARKS")%>" placeholder="" ></div>
				  </div><br/>
				  <!-- New Row (3) -->
				  <div class="edit_form">
					<div class="edit_mark">Project Name :</div>
					<div><input type="text" name="FCNAME" value="<%=rsGAD("FCNAME")%>" placeholder="Code" ></div>
				  </div>
				  <div class="edit_form">
					<div class="edit_mark">Customer :</div>
					<div>
						<input type="text" value="<%=rsGAD("BUSINESSUNITREQUESTOR")%>" disabled>
						<input type="hidden" name="CUSTOMER" value="<%=rsGAD("BUSINESSUNITREQUESTOR")%>" placeholder="CUSTOMER" >
					</div>
				  </div>
				   
				  <div class="edit_form">
					<div class="edit_mark">Valid Until :</div>
					<div><%dateCallComboUpdate "MONTH_VALID_UNTIL","DAY_VALID_UNTIL","YEAR_VALID_UNTIL",rsGAD("VALID_UNTIL")%></div>
				  </div>
				  <div class="edit_form">
					<div class="edit_mark">Project Group :</div>
					<div>
			
					<select id="PROJECTGROUPCODE<%=NO%>" name="PROJECTGROUPCODE" onchange="ControlJob(this.value,'<%=NO%>')" style="width:208px" <%=disabled_edit%>>

						<%	

						q = "select FCCODE,FCNAME from PROJECTACCGROUP Where fcba = '" & session("fcba") & "' order by fccode"
							Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
							rsCallCombo.CursorType = 3
							rsCallCombo.CursorLocation = 3
							rsCallCombo.Open q, connstr
							if not rsCallCombo.eof then%>
							 <option value=""></option><%
							 Do While Not rsCallCombo.eof
								if rsCallCombo("FCCODE") = rsGAD("PROJECTGROUPCODE")then
							%>	
								<option value="<%=rsCallCombo("FCCODE")%>" selected><%=rsCallCombo("FCCODE")%>&nbsp;:&nbsp;<%=rsCallCombo("FCNAME")%></option>  	
							<% else %>
								<option value="<%=rsCallCombo("FCCODE")%>"><%=rsCallCombo("FCCODE")%>&nbsp;:&nbsp;<%=rsCallCombo("FCNAME")%></option>  	
							<%
							end if
							rsCallCombo.movenext
							loop
						End if
						rsCallCombo.Close
						set rsCallCombo = Nothing
						q=""%>	
						</select>
					</div>
				  </div><br/>
				  <!-- New Row (3) -->
				  <div class="edit_form">
					<div class="edit_mark">Project Type :</div>
					<div><input type="text" name="PROJECTTYPECODE" value="<%=rsGAD("PROJECTTYPECODE")%>" placeholder="Code" ></div>
				  </div>
				  <div class="edit_form">
					<div class="edit_mark">Quantity :</div>
					<div>
						<input type="text" name="QUANTITY" value="<%=rsGAD("PROJECTQTYREQUEST")%>" placeholder="QUANTITY"  >
						<!--<input type="hidden" name="QUANTITY" value="<%=rsGAD("PROJECTQTYREQUEST")%>" placeholder="QUANTITY" >-->
					</div>
				  </div> 
				  <div class="edit_form">
					<div class="edit_mark">Row State Id :</div>
					<div><input type="text" name="ROWSTATEID" value="<%=rsGAD("ROWSTATEID")%>" placeholder="" ></div>
				   </div>
				  
				   <!-- New Row (3) -->
				   <div class="edit_form">
					<div class="edit_mark">Project Manager :</div>
					<div><input type="text" name="PROJECTMANAGER" value="<%=rsGAD("PROJECTMANAGER")%>" placeholder="" ></div>
				   </div><br/>
				   <div class="edit_form">
					<div class="edit_mark">Control Job :</div>
					<div>
					<input id="controljob<%=NO%>" type="text" value="<%=rsGAD("CONTROLJOBCODE")%>" name="CONTROLJOBCODE" placeholder="control job" readonly='True' style='background:#FFFFCC'>
					
				   </div><br/>
				   <!-- Submit / Button -->
				  <div style="display: inline-block;">
				  <input onclick="return confirm('Are you sure you want to Update this record?');" style="width:30px;font-family:'FontAwesome'; padding:3px 5px;" type="submit" value="&#xf0c7;">&nbsp;&nbsp;
				  <input style="width:30px;font-family:'FontAwesome'; padding:3px 5px;" type="button" onclick="javascript:toggle('editing_form_<%=NO%>','edit<%=NO%>')" value="&#xf05e;">
				  </div>
				  </form>
				</td>
		    </tr>
		<% elseif Request("txttable") ="ProjectAccountingDetail" then %>
			<tr id="editing_form_<%=NO%>" style="display:none;">
			  <td class="arrow"><br/><br/><i class="fa fa-reply"></i></td>
			  <td colspan="<%=nCountofField%>" align="left">
			 <form method="post" action="ProjectAccounting.asp?txtModule=<%=request("txtModule")%>&txtTable=<%=request("txtTable")%>&txtIdCode=<%=rsGAD("PROJECTCODE")%>&txtJenis=Edit" style="border:none;padding:0px;"> 
			  <div style=" width:100%; display: inline-block;">
				  <div class="edit_form">
					<div class="edit_mark">Project Code :</div>
					<div>
						<input type="text" value="<%=rsGAD("PROJECTCODE")%>" disabled>
						<input type="hidden" name="PROJECTCODE" value="<%=rsGAD("PROJECTCODE")%>">
					</div>
				  </div>

				  <div class="edit_form">
					<div class="edit_mark">Line Number :</div>
					<div>
						<input type="text" value="<%=rsGAD("LINENO")%>" disabled>
						<input type="hidden" name="LINENO" value="<%=rsGAD("LINENO")%>">
					</div>
				  </div><br/>
				  <div class="edit_form">
					<div class="edit_mark">Valid Input From :</div>
					<div><%dateCallComboUpdate "MONTH_VALID_INPUT_FROM","DAY_VALID_INPUT_FROM","YEAR_VALID_INPUT_FROM",rsGAD("VALID_INPUT_FROM")%></div>
				  </div>
				 <div class="edit_form">
					<div class="edit_mark">Valid Input Until :</div>
					<div><%dateCallComboUpdate "MONTH_VALID_INPUT_UNTIL","DAY_VALID_INPUT_UNTIL","YEAR_VALID_INPUT_UNTIL",rsGAD("VALID_INPUT_UNTIL")%></div>
				  </div><br/>
				  <div class="edit_form">
					<div class="edit_mark">Job Ref Code :</div>
					<div>
					<select id="JOBREFCODE" name="JOBREFCODE" style="width:208px" <%=disabled_edit%>>

						<%	

						q = "select JOBREFCODE,JOB_API.get_jobcode_name(JOBREFCODE, FCBA) FCNAME from projectaccgroupdetail Where PROJECTGROUPCODE = '" & rsGAD("PROJECTGROUPCODE") & "' AND fcba = '" & session("fcba") & "' order by JOBREFCODE"
							Set rsCombo = Server.CreateObject("ADODB.Recordset")
							rsCombo.CursorType = 3
							rsCombo.CursorLocation = 3
							rsCombo.Open q, connstr
							'RESPONSE.WRITE q
						if not rsCombo.eof then%>
							 <%
							 Do While Not rsCombo.eof
							 if rsCombo("JOBREFCODE") = rsGAD("JOBREFCODE")then
							%>	
								<option value="<%=rsCombo("JOBREFCODE")%>" selected>
									<%=rsCombo("JOBREFCODE")%>&nbsp;:&nbsp;<%=rsCombo("FCNAME")%>
								</option>												
							<% else %>	
								<option value="<%=rsCombo("JOBREFCODE")%>"><%=rsCombo("JOBREFCODE")%>&nbsp;:&nbsp;<%=rsCombo("FCNAME")%></option>
							<%
								end if
							rsCombo.movenext
							loop
						End if
						rsCombo.Close
						set rsCombo = Nothing
						q=""%>	
						</select>
				  </div>
				  <div class="edit_form">
					<div class="edit_mark">Remarks :</div>
					<div><input type="text" name="REMARKS" value="<%=rsGAD(5)%>"></div>
				  </div>
				  
				  <br/>
				  
				   <!-- Submit / Button -->
				  <div style="display: inline-block;">
				  <input onclick="return confirm('Are you sure you want to Update this record?');" style="width:30px;font-family:'FontAwesome'; padding:3px 5px;" type="submit" value="&#xf0c7;">&nbsp;&nbsp;
				  <input style="width:30px;font-family:'FontAwesome'; padding:3px 5px;" type="button" onclick="javascript:toggle('editing_form_<%=NO%>','edit<%=NO%>')" value="&#xf05e;">
				  </div>
				  </form>
				</td>
		    </tr>

		<% end if %>
	<!-- END FORM EDIT / ROW ------------>	   
	
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
	    <a id="ButtonOpenWindow" class="add_list" href="#" onclick="javascript:toggle('form_window','ButtonOpenWindow')" value="Window - yellow"
		title="New :<%=txtTable%>"><i class="fa fa-plus-circle"></i><span>ADD LIST</span></a>
		
		<!-- BUTTON REFRESH --->
		<a href="#" class="add_list" onclick="reloadPage()"><i class="fa fa-refresh"></i><span>Refresh </span></a>&nbsp;
		
		<!-- SEARCHING -------->
		<form style="padding:0px; border:none;display:inline;margin:0px;" method="post" action="ProjectAccounting.asp?txtModule=<%=request("txtModule")%>&txtTable=<%=request("txtTable")%>&txtIdCode=<%=request("txtIdCode")%>"> 
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

	<div id="form_window" class="nswindowContainer" style="display:none;z-index: 1010; background-color:rgba(0,0,0,0.4);">
		<div id="nswindow" class="yellow" style="width: 800px; height:380px; left: 270px; top: 25px; opacity: 1;">
			<div id="ctrlbar" class="controlbar">
				<div class="title" style="margin-top: 4.5px;">Add Project list</div>
				<a id="ButtonCloseWindow" class="close" href="javascript:toggle('form_window','ButtonCloseWindow')" title="close">Close</a>
			</div>
			<div id="nswindowContent" style="background:#f1f1f1;top: 26px;">
			
			<!-- INPUT FORM ADD LIST --------------------------------->		
			
			<% if Request("txttable") ="ProjectAccounting" then %>
				<div id="WARP_input" style="height:340px;">
				<form class="form_input" style="background:#ffffff;" method="post" action="ProjectAccounting.asp?txtModule=<%=Request("txtModule")%>&txtTable=<%=Request("txtTable")%>&txtJenis=New">
				
				<table id="table_Input">
					<tr>
						<td>Project Code</td>
						<td><input type="text" name="FCCODE" value="" placeholder="Code"></td>
						<td></td>
						<td>Valid Form</td>
						<td><input type="text" class="datepicker" name="VALID_FROM" placeholder="Date"></td>
					</tr>
					<tr>
						<td>Project Name</td>
						<td><input type="text" name="FCNAME" value="" placeholder="Name"></td>
						<td></td>
						<td>Valid Until</td>
						<td><input type="text" class="datepicker" name="VALID_UNTIL" placeholder="Date"></td>
					</tr>
					<tr>
						<td>Project Group</td>
						<td>
						<select id="PROJECTGROUPCODE" name="PROJECTGROUPCODE" onchange="ControlJob(this.value,'')" style="width:208px" <%=disabled_edit%>>

						<%	

						q = "select * from PROJECTACCGROUP Where fcba = '" & session("fcba") & "' order by fccode"
							Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
							rsCallCombo.CursorType = 3
							rsCallCombo.CursorLocation = 3
							rsCallCombo.Open q, connstr
							if not rsCallCombo.eof then%>
							 <option value=""></option><%
							 Do While Not rsCallCombo.eof
							%>	
								<option value="<%=rsCallCombo("FCCODE")%>"><%=rsCallCombo("FCCODE")%>&nbsp;:&nbsp;<%=rsCallCombo("FCNAME")%></option>  	
							<%rsCallCombo.movenext
							loop
						End if
						rsCallCombo.Close
						set rsCallCombo = Nothing
						q=""%>	
						</select>
						</td>
						<td></td>
						<td>Row State Id</td>
						<td><input type="text" name="ROWSTATEID" value="" placeholder="Row State Id"></td>
					</tr>
					<tr>
						<td>Project Type</td>
						<td><input type="text" name="PROJECTTYPECODE" value="" placeholder="Project Type"></td>
						<td></td>
						<td>Project Manager</td>
						<td><input type="text" name="PROJECTMANAGER" value="" placeholder="Project Manager"></td>		
					</tr>
					
					<tr>
						<td>Plan Start Date</td>
						<td><input type="text" class="datepicker" name="PLANSTARTDATE" placeholder="Date"></td>
						<td></td>
						<td>Control Job</td>
						<td>
						<input id="controljob" type="text" name="CONTROLJOBCODE" placeholder="control job" readonly='True' style='background:#FFFFCC'>
						
					</tr>
					<tr>
						<td>Customer</td>
						<td><input type="text" name="CUSTOMER" value="" placeholder="Customer"></td>
						<td></td>
						<td>Remarks</td>
						<td><input type="text" name="REMARKS" value="" placeholder="Remarks"></td> 
					</tr>
					<tr>
						<td>Quantity</td>
						<td><input type="text" name="QUANTITY" value="" placeholder="Quantity"></td>
						<td></td>
						<td></td>
						<td></td> 
					</tr>
				</table>

				<div style="position:absolute; left:10px; bottom:10px;">
				<input type="submit" name="submit" value="Add List">
				<input id="ButtonCloseWindow" type="button" onclick="javascript:toggle('form_window','ButtonCloseWindow')" value="Cancel">
				</div>
				</form>
				</div>
			<% elseif Request("txttable") ="ProjectAccountingDetail" then %>
				<div id="WARP_input" style="height:170px;">
				<form class="form_input" style="background:#ffffff;" method="post" action="ProjectAccounting.asp?txtModule=<%=Request("txtModule")%>&txtTable=<%=Request("txtTable")%>&txtJenis=New">
				<table id="table_Input">
					<tr>
						<td>Project Code</td>
						<td>
						<input type="hidden" name="PROJECTCODE" value="<%=REQUEST("txtIdCode")%>">
						<input type="text" value="<%=REQUEST("txtIdCode")%>" disabled></td>
						<td></td>
						<td>Valid Input From</td>
						<td><input type="text" class="datepicker" name="VALID_INPUT_FROM" value="" placeholder="Date"></td>
					</tr>
					<tr>
						<td>Job Ref Code</td>
						<td>
						<select id="JOBREFCODE" name="JOBREFCODE" style="width:208px" <%=disabled_edit%>>

						<%	

						q = "select JOBREFCODE,JOB_API.get_jobcode_name(JOBREFCODE, FCBA) FCNAME from projectaccgroupdetail Where PROJECTGROUPCODE = '" & REQUEST("PROJECTGROUPCODE") & "' AND fcba = '" & session("fcba") & "' order by JOBREFCODE"
							Set rsCombo = Server.CreateObject("ADODB.Recordset")
							rsCombo.CursorType = 3
							rsCombo.CursorLocation = 3
							rsCombo.Open q, connstr
							RESPONSE.WRITE q
						if not rsCombo.eof then%><%
							 Do While Not rsCombo.eof
							%>
								<option value="<%=rsCombo("JOBREFCODE")%>"><%=rsCombo("JOBREFCODE")%>&nbsp;:&nbsp;<%=rsCombo("FCNAME")%></option>
							<%
							rsCombo.movenext
							loop
						End if
						rsCombo.Close
						set rsCombo = Nothing
						q=""%>	
						</select>
						
						<td></td>
						<td>Valid Input Until</td>
						<td><input type="text" class="datepicker" name="VALID_INPUT_UNTIL" placeholder="Date"></td>
					</tr>
					<tr>
						<td>Remarks</td>
						<td><input type="text" name="REMARKS" value="" placeholder="Remarks"></td>
						<td></td>
						<td></td>
						<td></td>
					</tr>
				</table>

				<div style="position:absolute; left:10px; bottom:10px;">
				<input type="submit" name="submit" value="Add List" placeholder="">
				<input id="ButtonCloseWindow" type="button" onclick="javascript:toggle('form_window','ButtonCloseWindow')" value="Cancel">
				</div>
				</form>
				</div>
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

