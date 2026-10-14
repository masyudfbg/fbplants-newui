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
		<script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.8.3/jquery.min.js"></script>	
		
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
dim Searching
Searching = request("Searching")

if len(trim(request("txtModule"))) = 0 and len(trim(request("txtTable"))) = 0 then
response.redirect "?txtModule=Data Setup&txtTable=Process" 
end if

if Request("txttable") ="Process" then
    if Request("txtJenis")="Delete" then
		qCheckData = " select count(*) totalData from PROCESS_DETAIL_TAB " & _
					  " where " & _
					  " PROCESSCODE = '" & Request("txtIdCode") &  "'"
					  
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
			 q="delete from PROCESS_TAB where FCCODE = '" & Request("txtIdCode")  & "'"
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
		q="delete from PROCESS_DETAIL_TAB where PROJECTCODE = '" & Request("txtIdCode")  & "'"
		connStr.execute(q)
		
		q="delete from PROCESS_TAB where FCCODE = '" & Request("txtIdCode")  & "'"
		connStr.execute(q)
		RESPONSE.REDIRECT "?txtModule=" & Request("txtModule") & "&txtTable=" & Request("txtTable") & ""
		 
	elseif Request("txtJenis")="New" then
		 q= "Insert into PROCESS_TAB (FCCODE, FCNAME,REMARKS, FCENTRY, FCEDIT, FCIP, LASTUPDATE, LASTTIME)" & _
				  "values ('" &  trim(Request("FCCODE")) & "','" & Request("FCNAME") & "','" & Request("REMARKS") & "','" & _
				  Session("CodeName") & "','" & Session("CodeName") & "','" & _
				  Request.ServerVariables("REMOTE_ADDR") &"', sysdate, To_Char(sysdate,'HH24:MI') )"
				  'response.write(q)
			connStr.execute(q)
			RESPONSE.REDIRECT "?txtModule=" & Request("txtModule") & "&txtTable=" & Request("txtTable") & ""
			
	elseif Request("txtJenis")="Edit" then	 
		 q= "UPDATE PROCESS_TAB SET " & _
			"FCNAME = '" & Request("FCNAME") & "', " & _
			"REMARKS = '" & Request("REMARKS") & "', " & _ 
			"FCENTRY = '" & Session("CodeName") & "', " & _
			"FCEDIT = '" & Session("CodeName") & "', " & _
			"FCIP = '" & Request.ServerVariables("REMOTE_ADDR") & "', " & _
			"LASTUPDATE = sysdate, " & _
			"LASTTIME = To_Char(sysdate,'HH24:MI') " & _
			"WHERE FCCODE = '" & Request("FCCODE") & "'"
		   
			connStr.execute(q)
			
			if err.number <> 0 then
			   response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & err.description & "</font>"
			End if
			
			RESPONSE.REDIRECT "?txtModule=" & Request("txtModule") & "&txtTable=" & Request("txtTable") & ""
	end if
	
	q	=	"SELECT " & _
			"FCCODE, FCNAME, REMARKS, " & _
			"FCENTRY, FCEDIT, FCIP, LASTUPDATE FROM PROCESS_TAB "
			
	if  LEN(TRIM(Searching)) > 0 then
	q	=	q & "WHERE UPPER(FCCODE) LIKE UPPER('%" & request("Searching") & "%') " & _
				"OR UPPER(FCNAME) LIKE UPPER('%" & request("Searching") & "%')" 
		
		'response.write(q)
	else
	q	=	q 
	end if 		
	'Response.Write q 
elseif Request("txttable") ="process_detail" then
	if Request("txtJenis")="Delete" then
	
			 q="delete from PROCESS_DETAIL_TAB where PROCESSCODE = '" & Request("txtProcess")  & "' AND FCCODE = '" & Request("txtIdCode")  & "'"
			 connStr.execute(q)
			 RESPONSE.REDIRECT "?txtModule=" & Request("txtModule") & "&txtTable=" & Request("txtTable") & "&txtIdCode=" & Request("txtProcess") & ""
	    
	elseif Request("txtJenis")="New" then
		 q= "Insert into PROCESS_DETAIL_TAB (PROCESSCODE,FCCODE,FCNAME,REMARKS,FCENTRY,FCEDIT,FCIP,LASTUPDATE,LASTTIME)" & _
				  "values ('" &  trim(Request("PROCESSCODE")) & "','" &  trim(Request("FCCODE")) & "','" & Request("FCNAME") & "','" & Request("REMARKS") & "','" & _
				  Session("CodeName") & "','" & Session("CodeName") & "','" & _
				  Request.ServerVariables("REMOTE_ADDR") &"', sysdate, To_Char(sysdate,'HH24:MI') )"
				  'response.write(q)
			connStr.execute(q)
			RESPONSE.REDIRECT "?txtModule=" & Request("txtModule") & "&txtTable=" & Request("txtTable") & "&txtIdCode=" & Request("PROCESSCODE") & ""
			
	elseif Request("txtJenis")="Edit" then	 
		 q= "UPDATE PROCESS_DETAIL_TAB SET " & _
			"FCNAME = '" & Request("FCNAME") & "', " & _
			"REMARKS = '" & Request("REMARKS") & "', " & _ 
			"FCENTRY = '" & Session("CodeName") & "', " & _
			"FCEDIT = '" & Session("CodeName") & "', " & _
			"FCIP = '" & Request.ServerVariables("REMOTE_ADDR") & "', " & _
			"LASTUPDATE = sysdate, " & _
			"LASTTIME = To_Char(sysdate,'HH24:MI') " & _
			"WHERE PROCESSCODE = '" & Request("PROCESSCODE")  & "' AND FCCODE = '" & Request("FCCODE") & "'"
		   
			connStr.execute(q)
			
			if err.number <> 0 then
			   response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & err.description & "</font>"
			End if
			
			RESPONSE.REDIRECT "?txtModule=" & Request("txtModule") & "&txtTable=" & Request("txtTable") & "&txtIdCode=" & Request("PROCESSCODE") & ""
	end if
	
    q	=	" SELECT" & _
				" PROCESSCODE, FCCODE, FCNAME, REMARKS" & _
				" , FCENTRY, FCEDIT, FCIP, LASTUPDATE FROM PROCESS_DETAIL_TAB "
				
				
	if  LEN(TRIM(Searching)) > 0 then
	q	=	q & "WHERE PROCESSCODE = '" &request("txtIdCode") & "' AND " & _
				"(UPPER(FCCODE) LIKE UPPER('%" & request("Searching") & "%') " & _
				"OR UPPER(FCNAME) LIKE UPPER('%" & request("Searching") & "%'))" 
		
		'response.write(q)
	else
	q	=	q & "WHERE PROCESSCODE = '" &request("txtIdCode") & "'"
	end if 		
		
	'Response.Write q 
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
	
	<%if Request("txttable") ="process_detail" then%>
		<div id="parentDetail">
			<div style="background-color:#0C5B24; padding:7px 3px;margin-top:2px;border-bottom:1px #000 solid;">
				<a class="head_link" href="?txtModule=<%=request("txtModule")%>&txtTable=Process" >&nbsp;
					<i class="fa fa-chevron-left"></i>&nbsp;&nbsp;Process&nbsp;:&nbsp;<%=request("txtIdCode")%>
				</a>
				<a id="hide_headDetail" onclick="javascript:toggle_headDetail('data_headDetail','hide_headDetail')">
					<i class="fa fa-ban"></i>
				</a>
			</div>
		</div>
	<%end if%>
<!-- END DISPLAY PARENT LIST AND MENU TAB ---->	
	
<!-- FLEXIGRID ----------------------------------->
	<table  class="flexme">
		<thead>
			<tr >
			<% if Request("txttable") ="Process" then %>
				<th width="10"  align="center" style="width:10px;">No</th>
				<th width="150" align="left">Code</th>	
				<th width="150" align="left">Name</th>
				<th width="400" align="left">Remarks</th>
				<th width="30" align="center">Edit</th>
				<th width="30" align="center">Delete</th>
				<th width="30" align="center">Detail</th>
			<% elseif Request("txttable") ="process_detail" then%>
				<th width="10"  align="center" style="width:10px;">No</th>
				<th width="150" align="left">Process</th>
				<th width="150" align="left">Code</th>	
				<th width="150" align="left">Name</th>
				<th width="300" align="left">Remarks</th>
				<th width="30" align="center">Edit</th>
				<th width="30" align="center">Delete</th>
			<% end if %>
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
		if Request("txttable") ="Process" then
				crashId = "txtIdCode=" & rsGAD(0) &""
				elseif Request("txttable") ="process_detail" then
				crashId = "txtProcess=" & rsGAD(0) & "&txtIdCode=" & rsGAD(1) & ""
				end if
		  %>
		  <tr class="linker" href="login.asp"  bgcolor="<%=cColor%>">
		  <td align="left" valign="top"><%=NO%></td>
		  <%for i = 0 to nCountofField - 4%>
					<td align="left" title="<%=rsGAD(i).Name%>" valign="top"><%=rsGAD(i)%> </td>
				<%next %>
		  <td align="center" valign="top">
		  <a id="edit<%=NO%>" class="pencil" onclick="javascript:toggle('editing_form_<%=NO%>','edit<%=NO%>')">
		  <i class="fa fa-pencil"></i>
		  </a></td>
		  <td align="center" valign="top"><a class="trash" href="process.asp?txtModule=<%=request("txtModule")%>&txtTable=<%=request("txtTable")%>&<%=crashId%>&txtJenis=Delete" onclick="return confirm('Are you sure you want to Delete this record?');"><i class="fa fa-trash-o"></i></a></td>
		   <% if Request("txttable") ="Process" then %>
		  <td>
		  <a class="pencil" href="process.asp?txtModule=<%=request("txtModule")%>&txtTable=process_detail&<%=crashId%>"><i class="fa fa-list"></i></i></a>
		  </td>
		  <% end if%>
		  </tr>
		  
		  <tr id="editing_form_<%=NO%>" style="display:none;">
		  
		<% if Request("txttable") ="process_detail" then %>
		  
		  <td class="arrow"><i class="fa fa-reply"></i></td>
		  <td colspan="<%=nCountofField%>" align="left" >
		  <form method="post" action="process.asp?txtModule=<%=request("txtModule")%>&txtTable=<%=request("txtTable")%>&txtJenis=Edit"  style="border:none;padding:0px;">
		  Code :
		<%for i = 0 to nCountofField - 6%>
				<input type="hidden" name="<%=rsGAD(i).Name%>" value="<%=rsGAD(i)%>"><%=rsGAD(i)%> /

		<%next %>
			<br/><br/>
			<div style=" width:120px; display: inline-block;">Name :</div>
			<div style="display: inline-block;">
				<input type="text" name="FCNAME" style="width:300px;" value="<%=rsGAD("FCNAME")%>">
			</div><br/>
			<div style=" width:120px; display: inline-block;">Remarks :</div>
			
			<div style="display: inline-block;">
				<textarea name="REMARKS" style="width:300px;" ><%=rsGAD("REMARKS")%></textarea>
			</div><br/>
			
			<div style="display: inline-block;">
			  <input onclick="return confirm('Are you sure you want to Update this record?');" style="font-family:'FontAwesome'; padding:3px 5px;" type="submit" value="&#xf0c7;">
			  <input style="font-family:'FontAwesome'; padding:3px 5px;" type="button" onclick="javascript:toggle('editing_form_<%=NO%>','edit<%=NO%>')" value="&#xf05e;">
			</div>
			
			
			</form>  
		  </td>
		 
		<% elseif Request("txttable") ="Process" then %>
		 
		  <td class="arrow"><br/><br/><i class="fa fa-reply"></i></td>
		  <td colspan="<%=nCountofField%>" align="left">
			<form method="post" action="process.asp?txtModule=<%=request("txtModule")%>&txtTable=<%=request("txtTable")%>&txtJenis=Edit" style="border:none;padding:0px;">  
			  <!--<td colspan="<%=nCountofField%>"  >-->
			<%for i = 0 to nCountofField - 6%>
					   
						<input type="hidden" name="<%=rsGAD(i).Name%>" value="<%=rsGAD(i)%>">Edit Data : <%=rsGAD(i)%> <br/>
				
					<%next %>
			  <br/>
			  <div style=" width:120px; display: inline-block;">
			  Name :
			  </div>
			  <div style="display: inline-block;">
			  <input type="text" name="FCNAME" value="<%=rsGAD("FCNAME")%>">
			  </div><br/>
			  <div style="width:120px; display: inline-block;">
			  Remarks :
			  </div>
			  <div style="display: inline-block;">
			  <textarea name="REMARKS" style="width:300px;" ><%=rsGAD("REMARKS")%></textarea>
			  </div><br/>
			  <div style="display: inline-block;">
			  <input onclick="return confirm('Are you sure you want to Update this record?');" style="font-family:'FontAwesome'; padding:3px 5px;" type="submit" value="&#xf0c7;">
			  <input style="font-family:'FontAwesome'; padding:3px 5px;" type="button" onclick="javascript:toggle('editing_form_<%=NO%>','edit<%=NO%>')" value="&#xf05e;">
			  </div>
			  </form>
			  </td>
		  <%end if%>
		  </tr>
		   
		 <%
			
		  rsGAD.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   end if 
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
		<form style="padding:0px; border:none;display:inline;margin:0px;" method="post" action="process.asp?txtModule=<%=request("txtModule")%>&txtTable=<%=request("txtTable")%>&txtIdCode=<%=request("txtIdCode")%>"> 
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



 
<!--TEMPLATE ADD LIST FORM 
	-----------------------
	Masukkan format inputan kedalam comment tag "INPUT FORM ADD LIST " ------>

	<div id="form_window" class="nswindowContainer" style="display:none;z-index: 1010; background-color:rgba(0,0,0,0.4);">
		<div id="nswindow" class="yellow" style="width: 700px; height:380px; left: 333px; top: 25px; opacity: 1;">
			<div id="ctrlbar" class="controlbar">
				<div class="title" style="margin-top: 4.5px;">Add Project list</div>
				<a id="ButtonCloseWindow" class="close" href="javascript:toggle('form_window','ButtonCloseWindow')" title="close">Close</a>
			</div>
			<div id="nswindowContent" style="background:#f1f1f1;top: 26px;">
			
			<!-- INPUT FORM ADD LIST --------------------------------->		
			
			<% if Request("txttable") ="Process" then %>
				<div id="WARP_input" style="height:130px;">
				<form class="form_input" style="background:#ffffff;" method="post" action="process.asp?txtModule=<%=Request("txtModule")%>&txtTable=<%=Request("txtTable")%>&txtJenis=New">
				
				<table id="table_Input">
				<tr>
					<td>Code</td>
					<td><input type="text" name="FCCODE" value="" placeholder="Code"></td>
					<td></td>
					<td>Name</td>
					<td><input type="text" name="FCNAME" value="" placeholder="Name"></td>
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
				<input type="submit" name="submit" value="Add List">
				<input id="ButtonCloseWindow" type="button" onclick="javascript:toggle('form_window','ButtonCloseWindow')" value="Cancel">
				</div>
				</form>
				</div>
			<% elseif Request("txttable") ="process_detail" then %>
				<div id="WARP_input" style="height:130px;">
				<form class="form_input" style="background:#ffffff;" method="post" action="process.asp?txtModule=<%=Request("txtModule")%>&txtTable=<%=Request("txtTable")%>&txtJenis=New">
				<table id="table_Input">
				<tr>
					<td>Process Code</td>
					<td>
						<input type="text" value="<%=request("txtIdCode")%>" disabled>
						<input type="hidden" name="PROCESSCODE" value="<%=request("txtIdCode")%>" placeholder="Code">
					</td>
					<td></td>
					<td>Code</td>
					<td><input type="text" name="FCCODE" value="" placeholder="Code"></td>
				</tr>
				<tr>
					<td>Name</td>
					<td><input type="text" name="FCNAME" value="" placeholder="Name"></td>
					<td></td>
					<td>Remarks</td>
					<td><input type="text" name="REMARKS" value="" placeholder="Remarks"></td>
					
				</tr>

				<div style="position:absolute; left:10px; bottom:10px;">
				<input type="submit" name="submit" value="Add List" placeholder="">
				<input id="ButtonCloseWindow" type="button" onclick="javascript:toggle('form_window','ButtonCloseWindow')" value="Cancel">
				</div>
				</form>
				</div>
			<% else %>
			
				<div style="text-align:center; margin-top:10px;">
				<div class="loader">No Form</div>
				</div>
			<%end if%>

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

