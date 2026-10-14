<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<html>
<head>
<title>ePlaS-<%=request("txtTable")%></title>
<link rel="StyleSheet" href="dtree.css" type="text/css" />
<link rel="StyleSheet" href="style-content.css" type="text/css" />
<link rel="StyleSheet" href="2col_leftNav.css" type="text/css" />
<link rel="stylesheet" href="css/jquery-ui.css" />
<link rel="stylesheet" href="css/ns-window.css" />
<link rel="stylesheet" href="css/font/font-awesome.css"/>
<script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.8.3/jquery.min.js"></script>
<!-- check box True / false  -->

<!--End check box True / false	 -->	
<link rel="stylesheet" type="text/css" href="css/flexigrid.pack.css" />
<script type="text/javascript" src="js/flexigrid.pack.js"></script>
<script type="text/javascript" src="js/ns-window.js"></script>
<script type="text/javascript" src="dtree.js"></script>
<!--
<script src="http://code.jquery.com/jquery-latest.min.js"></script>
<script language="JavaScript" src='/js/x.js'></script>
<script language="JavaScript" src='/js/xtableheaderfixed.js'></script>
-->
<!-- Link table(tr)
<script language="JavaScript">
$(document).ready(function(){
    $('table tr.linker').click(function(){
        window.location = $(this).attr('href');
        return false;
    });
});
</script>-->
<script language="JavaScript">
xAddEventListener(window, 'load',
  function() {
    new xTableHeaderFixed('xxxx', 'xxxx',0);
  }, false
);
</script>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js" ></script>
<script src="js/jquery-ui.js"></script>
<script language="JavaScript" src="/js/date-picker.js"></script> 
<script>
$(function() {
$( ".datepicker" ).datepicker();
});
</script>
<script language="JavaScript">
var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);
}
</script>
<script type="text/javascript">
function reloadPage()
  {
  location.reload();
  }
</script>
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
<body>

<%
if Request("txttable") ="ProjectAccGroup" then
	if Request("txtJenis")="Delete" then
	
	     qCheckData = " select count(*) totalData from projectaccgroupdetail " & _
					  " where " & _
					  " PROJECTGROUPCODE = '" & Request("txtIdCode") &  "' and " & _ 
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
			 q="delete from PROJECTACCGROUP where FCCODE = '" & Request("txtIdCode")  & "' AND FCBA = '" & Session("fcBA") & "'"
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
			q="delete from PROJECTACCGROUPDETAIL where PROJECTGROUPCODE = '" & Request("txtIdCode")  & "' AND FCBA = '" & Session("fcBA") & "'"
			 'response.write("test" & q)
			 connStr.execute(q)		 
			 
			 q="delete from PROJECTACCGROUP where FCCODE = '" & Request("txtIdCode")  & "' AND FCBA = '" & Session("fcBA") & "'"
			 'response.write("test" & q)
			 connStr.execute(q)		 
		 
	elseif Request("txtJenis")="New" then
		 q= "Insert into PROJECTACCGROUP (FCCODE, FCNAME,CONTROLJOBCODE, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME)" & _
				  "values ('" &  trim(Request("FCCODE")) & "','" & Request("FCNAME") & "','" & Request("CONTROLJOBCODE") & "','" & _
				  Session("CodeName") & "','" & Session("CodeName") & "','" & _
				  Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI') )"
				  'response.write(q)
			connStr.execute(q)
			if err.number <> 0 then
			   response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & err.description & "</font>"
			End if
			RESPONSE.REDIRECT "?txtModule=" & Request("txtModule") & "&txtTable=" & Request("txtTable") & ""
	elseif Request("txtJenis")="Edit" then	 
		 q= "UPDATE PROJECTACCGROUP SET " & _
			"FCNAME = '" & Request("FCNAME") & "', " & _
			"CONTROLJOBCODE = '" & Request("CONTROLJOBCODE") & "', " & _
			"FCENTRY = '" & Session("CodeName") & "', " & _
			"FCEDIT = '" & Session("CodeName") & "', " & _
			"FCIP = '" & Request.ServerVariables("REMOTE_ADDR") & "', " & _
			"FCBA = '" & Session("FCBA") & "', " & _
			"LASTUPDATE = sysdate, " & _
			"LASTTIME = To_Char(sysdate,'HH24:MI') " & _
			"WHERE FCCODE = '" & Request("FCCODE") & "' AND FCBA = '" & Session("FCBA") & "'"
			connStr.execute(q)
			
			if err.number <> 0 then
			   response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & err.description & "</font>"
			End if
			
			RESPONSE.REDIRECT "?txtModule=" & Request("txtModule") & "&txtTable=" & Request("txtTable") & ""
	end if
elseif Request("txttable") ="ProjectAccGroupDetail" then
	if Request("txtJenis")="Delete" then
	
		 q="delete from PROJECTACCGROUPDETAIL where PROJECTGROUPCODE = '" & Request("txtIdCode")  & "' " & _
			"AND JOBREFCODE = '" & Request("txtJobRefCode")  & "' AND FCBA = '" & Session("fcBA") & "'"
		 connStr.execute(q)
		 RESPONSE.REDIRECT "?txtModule=" & Request("txtModule") & "&txtTable=" & Request("txtTable") & "&txtIdCode=" & Request("txtIdCode") & ""
	elseif Request("txtJenis")="New" then
			 q= "Insert into PROJECTACCGROUPDETAIL (PROJECTGROUPCODE,JOBREFCODE,FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
					  "values ('" &  trim(Request("PROJECTGROUPCODE")) & "','"  &  Request("JOBREFCODE") & "','" & _
					  Session("CodeName") & "','" & Session("CodeName") & "','" & _
					  Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI'))"
			connStr.execute(q)
			RESPONSE.REDIRECT "?txtModule=" & Request("txtModule") & "&txtTable=" & Request("txtTable") & "&txtIdCode=" & Request("PROJECTGROUPCODE") & ""
	elseif Request("txtJenis")="Edit" then
		 q= "UPDATE PROJECTACCGROUPDETAIL SET " & _
			"JOBREFCODE = '" & Request("JOBREFCODE") & "', " & _
			"FCENTRY = '" & Session("CodeName") & "', " & _
			"FCEDIT = '" & Session("CodeName") & "', " & _
			"FCIP = '" & Request.ServerVariables("REMOTE_ADDR") & "', " & _
			"FCBA = '" & Session("FCBA") & "', " & _
			"LASTUPDATE = sysdate, " & _
			"LASTTIME = To_Char(sysdate,'HH24:MI') " & _
			"WHERE PROJECTGROUPCODE = '" & Request("PROJECTGROUPCODE") & "' AND FCBA = '" & Session("FCBA") & "'"
			connStr.execute(q)
			
			if err.number <> 0 then
			   response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & err.description & "</font>"
			End if
			
			RESPONSE.REDIRECT "?txtModule=" & Request("txtModule") & "&txtTable=" & Request("txtTable") & "&txtIdCode=" & Request("PROJECTGROUPCODE") & ""
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

%>
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
	
<!--<div id="box_both_fixed" class="box_both" style="display: table-cell; overflow:scroll;padding:0px; top: 139px; margin-right:0.3%;">-->

<!----------------------- SQL Edit or Update Data ------------------------------------------>

<!----------------------- Tampilan select form --------------------------------------------->


<%if Request("txttable") ="ProjectAccGroup" then

	q	=	"SELECT " & _
			"FCCODE, FCNAME,CONTROLJOBCODE, FCENTRY, FCEDIT, FCIP,FCBA, LASTUPDATE,LASTTIME FROM PROJECTACCGROUP "
			
	if  LEN(TRIM(Searching)) > 0 then
	q	=	q & "WHERE (UPPER(FCCODE) LIKE UPPER('%"& Searching &"%') or " & _
            	"UPPER(FCNAME) LIKE UPPER('%"& Searching & "%') or and " & _
				"FCBA = '" & Session("fcBA") & "' "
		
		'response.write(q)
	else
	q	=	q & "WHERE FCBA = '" & Session("fcBA") & "' "
	end if
	q   =   q & "order by FCCODE"
elseif Request("txttable") ="ProjectAccGroupDetail" then
    settingColum = -4
	settingHeightTable = 58
	q	=	"SELECT " & _
			"e1.PROJECTGROUPCODE,e1.JOBREFCODE," & _
			"e2.FCCODE,e2.FCNAME " & _
			"FROM PROJECTACCGROUPDETAIL e1,PROJECTACCGROUP e2 " & _
			"WHERE e1.PROJECTGROUPCODE = '" & REQUEST("txtIdCode")& "' AND e1.FCBA = '" & Session("fcBA") & "' " & _
			"AND e1.PROJECTGROUPCODE=e2.FCCODE AND e1.FCBA = e2.FCBA "
	
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
		'response.write("<script language=""javascript"">alert('Data Not Found,');window.location.href='?txtModule=" & Request("txtModule") & "&txtTable=ProjectAccGroup';</script>")
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
	<div style="width:200px; height:32px;">
	<!-- KOSONGKAN -->
	</div>
	<%if Request("txttable") ="ProjectAccGroupDetail" and not rsGAD.EOF then%>
		<div id="parentDetail">
			<div style="background-color:#0C5B24; padding:7px 3px;margin-top:2px;border-bottom:1px #000 solid;">
			<a class="head_link" href="?txtModule=Project&txtTable=ProjectAccGroup" >&nbsp;<i class="fa fa-chevron-left"></i>&nbsp;&nbsp;Project Code&nbsp;:&nbsp;<%=rsGAD(2)%></a>
			<a id="hide_headDetail" onclick="javascript:toggle_headDetail('data_headDetail','hide_headDetail')">
				<i id="icon_hideOn" style="display:block;" class="fa fa-chevron-up"></i>
				<i id="icon_hideOff" style="display:none;" class="fa fa-chevron-down"></i>
			</a>
			</div>
			<table id="data_headDetail" style="display:block;margin-left:10px;">
				<tr>
					<td>Name</td><td>:&nbsp;<%=rsGAD(3)%></td>
				</tr>
			</table>
		</div>
	<!-- Menu Tab -->	
		<ul id="tabs">
			<li id="current"><a href="#" name="tab1">Project Group Detail</a></li>
		</ul>
	<!-- end Menu Tab -->	
		<%elseif Request("txttable") ="ProjectAccGroupDetail" then%>
		<div id="parentDetail">
			<div style="background-color:#0C5B24; padding:7px 3px;margin-top:2px;border-bottom:1px #000 solid;">
			<a class="head_link" href="?txtModule=Project&txtTable=ProjectAccGroup" >&nbsp;<i class="fa fa-chevron-left"></i>&nbsp;&nbsp;Project Code&nbsp;:&nbsp;<%=request("txtIdCode")%></a>
			<a id="hide_headDetail" onclick="javascript:toggle_headDetail('data_headDetail','hide_headDetail')">
				<i class="fa fa-ban"></i>
			</a>
			</div>
		</div>
	<%end if%>
	
		<table class="flexme">
				<thead>
					<tr>
					<% if Request("txttable") ="ProjectAccGroup" then %>
						<th width="10"  align="center" style="width:10px;">No</th>
						<th width="200" style="width:100px;" align="left">Project Group Code</th>	
						<th width="200" align="left">Project Group Name</th>
						<th width="200" align="left">Control Job</th>
						<th width="30" align="center">Edit</th>
						<th width="30" align="center">Delete</th>
						<th width="30" align="center">Detail</th>
					<%elseif Request("txttable") ="ProjectAccGroupDetail" then%>
						<th width="10"  align="center" style="width:10px;">No</th>
						<th width="200" style="width:100px;" align="left">Project Group Code</th>	
						<th width="400" align="left">Job Ref</th>
						<th width="30" align="center">Edit</th>
						<th width="30" align="center">Delete</th>
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
			  <a id="edit<%=NO%>" title="Edit" class="pencil" onclick="javascript:toggle('editing_form_<%=NO%>','edit<%=NO%>')">
			  <i class="fa fa-pencil"></i></a>
			  </td>
		<% if Request("txttable") ="ProjectAccGroup" then %>
			  <td align="center" valign="top">
			  <a class="pencil" title="Delete" onclick="return confirm('Are you sure you want to Delete this record?');"
			  href="ProjectAccGroup.asp?txtModule=<%=request("txtModule")%>&txtTable=<%=request("txtTable")%>&txtIdCode=<%=rsGAD("FCCODE")%>&txtJenis=Delete">
			  <i class="fa fa-trash-o"></i></a>
			  </td>
		
			  <td align="center" valign="top">
			  <a id="detail_acc" class="pencil" title="Detail"
			  href="ProjectAccGroup.asp?txtModule=<%=request("txtModule")%>&txtTable=<%=request("txtTable")%>Detail&txtIdCode=<%=rsGAD("FCCODE")%>">
			  <i class="fa fa-list"></i>
			  </a>
			  </td>
		<%elseif Request("txttable") ="ProjectAccGroupDetail" then%>
			  <td align="center" valign="top">
			  <a class="pencil" title="Delete" onclick="return confirm('Are you sure you want to Delete this record?');"
			  href="ProjectAccGroup.asp?txtModule=<%=request("txtModule")%>&txtTable=<%=request("txtTable")%>&txtIdCode=<%=rsGAD("PROJECTGROUPCODE")%>&txtJobRefCode=<%=rsGAD("JOBREFCODE")%>&txtJenis=Delete">
			  <i class="fa fa-trash-o"></i></a>
			  </td>
			  
		    <%end if %>	  
		  </tr>
			  
	<!-- FORM EDIT / ROW ---->
		<% if Request("txttable") ="ProjectAccGroup" then %>
			<tr id="editing_form_<%=NO%>" style="display:none;">
			  <td class="arrow"><br/><br/><i class="fa fa-reply"></i></td>
			  <td colspan="<%=nCountofField%>" align="left">
			 <form method="post" action="ProjectAccGroup.asp?txtModule=<%=request("txtModule")%>&txtTable=<%=request("txtTable")%>&txtJenis=Edit" style="border:none;padding:0px;"> 
			  <div style=" width:100%; display: inline-block;">
				  <div class="edit_form">
					<div class="edit_mark">Project Code :</div>
					<div>
						<input type="text" value="<%=rsGAD("FCCODE")%>" disabled>
						<input type="hidden" name="FCCODE" value="<%=rsGAD("FCCODE")%>" placeholder="Code" >
					</div>
				  </div>
				  <div class="edit_form">
					<div class="edit_mark">Project Name :</div>
					<div>
						<input type="text" name="FCNAME" value="<%=rsGAD("FCNAME")%>" placeholder="Code" >
					</div>
				  </div><br/>
				   <div class="edit_form">
					<div class="edit_mark">Control Job :</div>
					<div>
						<select name="CONTROLJOBCODE" title="CONTROLJOB">
							<%cmdCallCombo "Job",rsGAD("CONTROLJOBCODE"),"",""%>
						</select>
						
					</div>
				  </div><br/>
				   <!-- Submit / Button -->
				  <div style="display: inline-block;">
				  <input onclick="return confirm('Are you sure you want to Update this record?');" style="font-family:'FontAwesome'; padding:3px 5px;" type="submit" value="&#xf0c7;">&nbsp;&nbsp;
				  <input style="font-family:'FontAwesome'; padding:3px 5px;" type="button" onclick="javascript:toggle('editing_form_<%=NO%>','edit<%=NO%>')" value="&#xf05e;">
				  </div>
				  </form>
				</td>
		    </tr>
		<% elseif Request("txttable") ="ProjectAccGroupDetail" then %>
			<tr id="editing_form_<%=NO%>" style="display:none;">
			  <td class="arrow"><br/><br/><i class="fa fa-reply"></i></td>
			  <td colspan="<%=nCountofField + 1%>" align="left">
			 <form method="post" action="ProjectAccGroup.asp?txtModule=<%=request("txtModule")%>&txtTable=<%=request("txtTable")%>&txtIdCode=<%=rsGAD("PROJECTGROUPCODE")%>&txtJenis=Edit" style="border:none;padding:0px;"> 
			  <div style=" width:100%; display: inline-block;">
				  <div class="edit_form">
					<div class="edit_mark">Project Code :</div>
					<div>
						<input type="text" value="<%=rsGAD("PROJECTGROUPCODE")%>" disabled>
						<input type="hidden" name="PROJECTGROUPCODE" value="<%=rsGAD("PROJECTGROUPCODE")%>">
					</div>
				  </div>
				  <div class="edit_form">
					<div class="edit_mark">Job Ref Code :</div>
					<div>
					<select name="JOBREFCODE" title="JOBREFCODE">
							<%cmdCallCombo "Job",rsGAD("JOBREFCODE"),"",""%>
					</select>
					</div>
				  </div>
				  <br/>
				   <!-- Submit / Button -->
				  <div style="display: inline-block;">
				  <input onclick="return confirm('Are you sure you want to Update this record?');" style="font-family:'FontAwesome'; padding:3px 5px;" type="submit" value="&#xf0c7;">&nbsp;&nbsp;
				  <input style="font-family:'FontAwesome'; padding:3px 5px;" type="button" onclick="javascript:toggle('editing_form_<%=NO%>','edit<%=NO%>')" value="&#xf05e;">
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
	

			<!-- Akhir Tampilan Select form -->

		<input  type="hidden" value="<%=txtTable%>" name="txtTable">
		<input  type="hidden" value="<%=txtModule%>" name="txtModule">
		<!--</form>-->

<!--</div>-->

	<div style="background-color:#f1f1f1; padding:0.3%; border-top:1px #ccc solid; border-bottom:1px #ccc solid; position:fixed; top:114px; right:0.3%; width:76.8%;">
	    <a id="ButtonOpenWindow" class="add_list" href="#" onclick="javascript:toggle('form_window','ButtonOpenWindow')" value="Window - yellow"
		title="New : <%=txtTable%>"><i class="fa fa-plus-circle"></i><span>ADD LIST</span></a>
		<!--<a class="add_list" href="#" onclick="showWindow('yellow');" value="Window - yellow"
		title="New : <%=txtTable%>"><i class="fa fa-plus-circle"></i><span>ADD LIST</span></a>-->
		<a href="#" class="add_list" onclick="reloadPage()"><i class="fa fa-refresh"></i><span>Refresh </span></a>&nbsp;
		<form style="padding:0px; border:none;display:inline;margin:0px;" method="post" action="ProjectAccGroup.asp?txtModule=<%=request("txtModule")%>&txtTable=<%=request("txtTable")%>"> 
		<input class="search" type="text" name="Searching" value="" placeholder="Search Data">
		<input class="button_search" type="submit" value=""> 
		<i class="fa fa-search icon_search_data"></i>
		</form>
		


<!--
Format Page ------------------------------------------------------>	
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
		end if
		
			%>
		   
			   <span style="margin-top:15px;"><font size="1" color="#000000" face="Arial, Helvetica, sans-serif">Total Records : <%=intRecords%> </font></span>
		</div>
<!-- End Format Page ------------------------------------------------------>	  	
	</div>


</div>
<!------------------ BAWAH ISI------->

  <!--#include file="bottom.jsp"-->
<%end if%>
<!-- Add list form ------------------------------------------------------>

<div id="form_window" class="nswindowContainer" style="display:none;z-index: 1010; background-color:rgba(0,0,0,0.4);">
	<div id="nswindow" class="yellow" style="width: 800px; height:380px; left: 280px; top: 25px; opacity: 1;">
		<div id="ctrlbar" class="controlbar">
			<div class="title" style="margin-top: 4.5px;">Add Project list</div>
			<a id="ButtonCloseWindow" class="close" href="javascript:toggle('form_window','ButtonCloseWindow')" title="close">Close</a>
		</div>
		<div id="nswindowContent" style="background:#f1f1f1;top: 26px;">
		<% if Request("txttable") ="ProjectAccGroup" then %>
			<div id="WARP_input" style="height:130px;">
			<form class="form_input" style="background:#ffffff;" method="post" action="ProjectAccGroup.asp?txtModule=<%=Request("txtModule")%>&txtTable=<%=Request("txtTable")%>&txtJenis=New">
			<table id="table_Input">
				<tr>
					<td>Project Group Code</td>
					<td><input type="text" name="FCCODE" value="" placeholder="Code"></td>
					<td></td>
					<td>Project Group Name</td>
					<td><input type="text" name="FCNAME" value="" placeholder="Name"></td>
				</tr>
				<tr>
					<td>Control Code</td>
					<td colspan="4">
					<select name="CONTROLJOBCODE" title="CONTROL JOB">
							<%cmdCallCombo "Job","","",""%>
					</select>
					</td>
				</tr>
			</table>

			<div style="position:absolute; left:10px; bottom:10px;">
			<input type="submit" name="submit" value="Add List" placeholder="">
			<input id="ButtonCloseWindow" type="button" onclick="javascript:toggle('form_window','ButtonCloseWindow')" value="Cancel">
			</div>
			</form>
			</div>
		<% elseif Request("txttable") ="ProjectAccGroupDetail" then %>
			<div id="WARP_input" style="height:85px;">
			<form class="form_input" style="background:#ffffff;" method="post" action="ProjectAccGroup.asp?txtModule=<%=Request("txtModule")%>&txtTable=<%=Request("txtTable")%>&txtJenis=New">
			<table id="table_Input">
				<tr>
					<td>Project Group Code</td>
					<td>
					<input type="hidden" name="PROJECTGROUPCODE" value="<%=REQUEST("txtIdCode")%>">
					<input type="text" value="<%=REQUEST("txtIdCode")%>" disabled></td>
					<td></td>
					<td>Job Ref Code</td>
					<td>
						<select name="JOBREFCODE" title="JOBREFCODE">
							<%cmdCallCombo "Job",fccode,"",""%>
						</select>
					</td>
				</tr>
			</table>

			<div style="position:absolute; left:10px; bottom:10px;">
			<input type="submit" name="submit" value="Add List" placeholder="">
			<input id="ButtonCloseWindow" type="button" onclick="javascript:toggle('form_window','ButtonCloseWindow')" value="Cancel">
			</div>
			</form>
			</div>
			
			
			
			
			
		<% end if%>
		</div>	
	</div>
</div>
<!-- end Add list form ------------------------------------------------------>

<script type="text/javascript">

			var h_menu = $(".body_menu").height();
			var h_HeadParent = $("#parentDetail").height();
			var h_TabMenu = $("#tabs").height();
			var result = h_menu - (41 + h_HeadParent + h_TabMenu);
            $('.flexme').flexigrid({
				height:result
			});

        </script>
</body>
</html>

