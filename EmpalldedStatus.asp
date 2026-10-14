<%response.buffer=True%>
<%validatesession%>

<!--#include file="conndB1.Inc"-->
<!--#include file="txtFunction.inc"-->
<html>
<head>
<title>ePlaS</title>

<link rel="StyleSheet" href="dtree.css" type="text/css" />
<link rel="StyleSheet" href="style-content.css" type="text/css" />
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
<script language="JavaScript">
xAddEventListener(window, 'load',
  function() {
    new xTableHeaderFixed('xxxx', 'xxxx',0);
  }, false
);

$(document).ready(function(){
    $('table tr.linker').click(function(){
        window.location = $(this).attr('href');
        return false;
    });
});
</script>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js" ></script>
<script language="JavaScript" src="/js/date-picker.js"></script> 
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
	function showWindow(theme) {
		$.nsWindow.open({
			title: 'Add Process',
			width: 700,
			height: 500,
			top:20,
			dataUrl: 'module/mod_process_tab/action.asp',
			theme:theme
        });
	}
	function reloadPage()
  {
  location.reload();
  }
</script>
<script>
            function toggle(id){
				var e = document.getElementById(id);
				if(e.style.display == 'table-row'){
					e.style.display = 'none';
				}else if(e.style.display == 'none'){
					e.style.display = 'table-row';
					
				}
			}
</script>	
</head>
<body>
<!--#include file="top.jsp"-->

<%


Dim txtTable,txtModule,txtJenis,Searching,cmdView,crashId,qUpdateString,rsFrez,q,aksesfinal,ScriptName
ScriptName	=	Request.ServerVariables("SCRIPT_NAME")
txtTable = Request("txtTable")
txtModule = Request("txtModule")
Searching=Request("Searching")
txtJenis=Request("txtJenis")
if txtTable = "" then
response.redirect "?txtTable=Payroll_Control_Status"
else

end if

if Len(trim(Searching))=0 then
	Searching = year(date)
end if	


set rsFrez = server.CreateObject("Adodb.recordset")
q = " SELECT COUNT(1) total" & _
	" FROM authorizelevel al" & _
	" INNER JOIN authorizestructure aus ON aus.authorizelevelcode = al.fccode" & _
	" INNER JOIN authorizetemplate at ON aus.authorizetemplatecode = at.fccode" & _
	" INNER JOIN authorizelevel_users alu ON (alu.authorizelevelcode = al.fccode)" & _
	" WHERE at.document_type = 'XX'" & _
	"       AND al.isactive = 'TRUE'" & _
	"       AND al.finallevel = 'TRUE'" & _
	"		AND alu.isactive = 'TRUE'" & _
	"       AND aus.fcusercode = '" & session("CodeName") & "'" 
rsFrez.Open q, connstr
if Not rsFrez.Eof then
	aksesfinal =  cdbl(to_nol(rsFrez("total")))
else
	aksesfinal = 0
End if
rsFrez.Close : set rsFrez = Nothing 

%>
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

<%if Request("txttable") ="EmpalldedStatus" then
        if Request("txtJenis")= "generate"  then
			if LEN(TRIM(Searching)) > 0  Then
				qUpdateString = "BEGIN EMPALLDEDSTATUS_API.GenerateStatus (" & Searching & ",'" & _
																			   Session("fcBA") & "','" & _
																			   Session("CodeName") & "','" & _
																			   Session("CodeName") & "','" & _
																			   Request.ServerVariables("REMOTE_ADDR") & "'); END;"

					'response.write qUpdateString
					ExecuteAndLog ScriptName, qUpdateString
					response.redirect "?txtTable=EmpalldedStatus"
			end if
		end if
		if Request("txtJenis")= "edit"  then
				'DIM txtUpdateStatus
			'-- Check for Empallded PreProcess
				'qCheckStatus = "select EMPALLDEDSTATUS_API.GetCheckPreUpdate (" & request("YEAR") & "," & _
				'																  request("MONTH") & ",'" & _
				'																  Session("fcBA") & "') UpdateStatus from dual "
				'Set rsCheckStatus = Server.CreateObject("ADODB.Recordset")
				'rsCheckStatus.CursorType = 3
				'rsCheckStatus.CursorLocation = 3
				'rsCheckStatus.Open qCheckStatus, connstr
				'if not rsCheckStatus.eof then
				'   txtUpdateStatus = rsCheckStatus("UpdateStatus")
				'else
				'   Response.Write  "<font face=Verdana size=1 color=red>" & "EmpalldedStatus are not defined "
				'   txtUpdateStatus ="TRUE"
				'end if
				'rsCheckStatus.close
				'set rsCheckStatus = Nothing
				
				qUpdateString = " UPDATE EMPALLDEDSTATUS " & _ 
				                " SET " & _ 
								"     PAYROLL_EXECUTED='" & request("PAYROLL_EXECUTED") & "'," & _
								"     CLOSING_EXECUTED='" & request("CLOSING_EXECUTED") & "'," & _
								"     CLOSING_ACC='" & request("CLOSING_ACC") & "'," & _
								"     ARESTA_CLOSING='" & request("ARESTA_CLOSING") & "'," & _
								"     KASBANK_CLOSING='" & request("KASBANK_CLOSING") & "'," & _
								"     STOCK_CLOSING='" & request("STOCK_CLOSING") & "'," & _
								"     WIP_CLOSING='" & request("WIP_CLOSING") & "'," & _
								"     WSVRA_CLOSING='" & request("WSVRA_CLOSING") & "'," & _
								"     PAYROLL_CLOSING='" & request("PAYROLL_CLOSING") & "'," & _
								"     UNIT_CLOSING='" & request("UNIT_CLOSING") & "'," & _
								"     Start_date=to_date('" & request("START_DATE") & "','MM/DD/YYYY')," & _ 
								"     end_date=to_date('" & request("END_DATE") & "','MM/DD/YYYY') " & _
								"	  ,FCEDIT = '" & Session("CodeName")  & "'" & _
								"     ,FCIP = '" & Request.ServerVariables("REMOTE_ADDR")  & "'" & _
								"	  ,LASTUPDATE = sysdate " & _
								"	  ,LASTTIME = To_Char(sysdate,'HH24:MI') "& _
								" WHERE FCBA='" & Session("fcBA") & "' AND month=" & request("MONTH") & " AND year=" & request("YEAR") & ""


			''	response.write(qUpdateString)
				ExecuteAndLog ScriptName, qUpdateString
			response.redirect "?txtTable=EmpalldedStatus"
		end if
		
		Set dbHandler = Nothing
		
		if err.number <> 0 then
		   response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & err.description & "</font>"
		End if
   
	q	=	"SELECT " & _
			"YEAR, MONTH, PAYROLL_EXECUTED,CLOSING_EXECUTED,PROCESS_DATE,INVALID_DATE,closing_acc,closing_final,START_DATE,END_DATE, ARESTA_CLOSING, KASBANK_CLOSING, STOCK_CLOSING, WIP_CLOSING, WSVRA_CLOSING,PAYROLL_CLOSING,UNIT_CLOSING, " & _
			"FCENTRY, FCEDIT, FCIP,FCBA, LASTUPDATE,LASTTIME FROM EMPALLDEDSTATUS "
			
	if  LEN(TRIM(Searching)) > 0 then
	q	=	q & "WHERE (UPPER(YEAR) LIKE UPPER('%"& Searching &"%') or " & _
            	"UPPER(PAYROLL_EXECUTED) LIKE UPPER('%" & Searching & "%')) and " & _
				"FCBA = '" & Session("fcBA") & "' "
		
		' response.write(q)
	else
	q	=	q & "WHERE FCBA = '" & Session("fcBA") & "' "
	end if
else 
	response.write("No Data..")
	
end if
  q = q & "ORDER BY YEAR,MONTH "
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
	<div style="width:200px; height:32px;">
	
	</div>
		<table class="flexme" style="">
					<thead>
					<tr >
					<th width="10"  align="center" style="width:10px;">No</th>
					<th width="70"  align="left">Year</th>	
					<th width="70"  align="left">Month</th>
					<th width="100" align="left">Payroll Calculate</th>
					<th width="100" align="left">Payroll Lock</th>
					<th width="100" align="left">Accounting Lock</th>
					<th width="100" align="left">Final Lock</th>
					<th width="70"  align="left">Start Date</th>
					<th width="70"  align="left">End Date</th>
					<th width="100"  align="left">ARESTA CLOSING</th>
					<th width="100"  align="left">KASBANK CLOSING</th>
					<th width="100"  align="left">STOCK CLOSING</th>
					<th width="100"  align="left">WIP CLOSING</th>
					<th width="100"  align="left">WSVRA CLOSING</th>
					<th width="100"  align="left">PAYROLL CLOSING</th>
					<th width="100"  align="left">UNIT CLOSING</th>
					<th width="30" align="center">Edit</th>
						</tr>
					</thead>
					<tbody>			
					
<% 
				if Request("txttable") ="process_tab" then
				crashId = "FCCODE=" & rsGAD(0) &"&"
				elseif Request("txttable") ="process_detail_tab" then
				crashId = "PROCESSCODE=" & rsGAD(0) &"&FCCODE=" & rsGAD(1) &"&"
				end if

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
		  <tr class="linker" href="login.asp"  bgcolor="<%=cColor%>">
		  <td align="left" valign="top"><%=NO%></td>
		  <td align="left" title="YEAR" valign="top"><%=rsGAD("YEAR")%></td>
		  <td align="left" title="PAYROLL_EXECUTED" valign="top"><%=textMonth(rsGAD("MONTH"))%></td>
		  <td align="left" title="YEAR" valign="top"><%=rsGAD("PAYROLL_EXECUTED")%></td>
		  <td align="left" title="CLOSING_EXECUTED" valign="top"><%=rsGAD("CLOSING_EXECUTED")%></td>
		  <td align="left" title="CLOSING_ACC" valign="top"><%=rsGAD("CLOSING_ACC")%></td>
		  <td align="left" title="CLOSING_FINAL" valign="top"><%=rsGAD("CLOSING_FINAL")%></td>
		  
		  <td align="left" title="Start_date" valign="top"><%=rsGAD("START_DATE")%></td>
		  <td align="left" title="End_date" valign="top"><%=rsGAD("END_DATE")%></td>
		  <td align="left" title="ARESTA_CLOSING" valign="top"><%=rsGAD("ARESTA_CLOSING")%></td>
		  <td align="left" title="KASBANK_CLOSING" valign="top"><%=rsGAD("KASBANK_CLOSING")%></td>
		  <td align="left" title="STOCK_CLOSING" valign="top"><%=rsGAD("STOCK_CLOSING")%></td>
		  <td align="left" title="WIP_CLOSING" valign="top"><%=rsGAD("WIP_CLOSING")%></td>
		  <td align="left" title="WSVRA_CLOSING" valign="top"><%=rsGAD("WSVRA_CLOSING")%></td>
		  <td align="left" title="PAYROLL_CLOSING" valign="top"><%=rsGAD("PAYROLL_CLOSING")%></td>
		  <td align="left" title="UNIT_CLOSING" valign="top"><%=rsGAD("UNIT_CLOSING")%></td>
		  
		  <td align="center" valign="top">
		  <a id="edit<%=NO%>" class="pencil" onclick="javascript:toggle('editing_form_<%=NO%>','edit<%=NO%>')">
		  <i class="fa fa-pencil"></i>
		  </a></td>
		  </td>
		  </tr>
		  <tr id="editing_form_<%=NO%>" style="display:none;">
		  <td class="arrow"><br/><br/><i class="fa fa-reply"></i></td>
		   <td colspan="<%=nCountofField%>" align="left">
		 <form method="post" action="EmpalldedStatus.asp?txtTable=<%=request("txtTable")%>&txtJenis=edit" style="border:none;padding:0px;">  
		  <!--<td colspan="<%=nCountofField%>"  >-->
		  <input type="hidden" name="<%=rsGAD("MONTH").Name%>" value="<%=rsGAD("MONTH")%>"><%=textMonth(rsGAD("MONTH"))%>,&nbsp;
		  <input type="hidden" name="<%=rsGAD("YEAR").Name%>" value="<%=rsGAD("YEAR")%>"><%=rsGAD("YEAR")%> &nbsp;&nbsp;<br/><br/>
		  <div style=" width:120px; display: inline-block;">
		  Payroll Process :
		  </div>
		  
		  <div style="display: inline-block;">
			<%if rsGAD("CLOSING_FINAL")="TRUE" Then %>
			  <%if rsGAD("PAYROLL_EXECUTED")="TRUE" Then %>
			  <input type="radio" checked="checked" name="PAYROLL_EXECUTED" disabled ="TRUE" value="TRUE">&nbsp;<label for="square-checkbox-2">True</label>
			  <input type="radio" value="FALSE" disabled ="TRUE" name="PAYROLL_EXECUTED">&nbsp;<label for="square-checkbox-2">False</label>
			  <%else%>
			  <input type="radio" name="PAYROLL_EXECUTED" disabled ="TRUE" value="TRUE">&nbsp;<label for="square-checkbox-2">True</label>
			  <input type="radio" checked="checked" disabled ="TRUE" name="PAYROLL_EXECUTED" value="FALSE" >&nbsp;<label for="square-checkbox-2">False</label>
			  <%end if%>
		    <%else%>
			   <%if rsGAD("PAYROLL_EXECUTED")="TRUE" Then %>
			  <input type="radio" checked="checked" name="PAYROLL_EXECUTED" value="TRUE">&nbsp;<label for="square-checkbox-2">True</label>
			  <input type="radio" value="FALSE" name="PAYROLL_EXECUTED">&nbsp;<label for="square-checkbox-2">False</label>
			  <%else%>
			  <input type="radio" name="PAYROLL_EXECUTED" value="TRUE">&nbsp;<label for="square-checkbox-2">True</label>
			  <input type="radio" checked="checked" name="PAYROLL_EXECUTED" value="FALSE" >&nbsp;<label for="square-checkbox-2">False</label>
			  <%end if%>
			<%end if%>  
			
		  </div><br/>
		  <div style=" width:120px; display: inline-block;">
		 
   		 Payroll Lock :
		  </div>
		  <div style="display: inline-block;">
		   <%if rsGAD("CLOSING_FINAL")="TRUE" Then %>
			  <%if rsGAD("CLOSING_EXECUTED")="TRUE" Then %>
			  <input type="radio" checked="checked" name="CLOSING_EXECUTED"  disabled ="TRUE" value="TRUE">&nbsp;<label for="square-checkbox-2">True</label>
			  <input type="radio" name="CLOSING_EXECUTED" disabled ="TRUE" value="FALSE">&nbsp;<label for="square-checkbox-2">False</label>
			  <%else%>
			  <input type="radio" name="CLOSING_EXECUTED" disabled ="TRUE" value="TRUE">&nbsp;<label for="square-checkbox-2">True</label>
			  <input type="radio" checked="checked" name="CLOSING_EXECUTED" disabled ="TRUE" value="FALSE">&nbsp;<label for="square-checkbox-2">False</label>
			  <%end if%>
		    <%else%>	
			  <%if rsGAD("CLOSING_EXECUTED")="TRUE" Then %>
			  <input type="radio" checked="checked" name="CLOSING_EXECUTED" value="TRUE">&nbsp;<label for="square-checkbox-2">True</label>
			  <input type="radio" name="CLOSING_EXECUTED" value="FALSE">&nbsp;<label for="square-checkbox-2">False</label>
			  <%else%>
			  <input type="radio" name="CLOSING_EXECUTED" value="TRUE">&nbsp;<label for="square-checkbox-2">True</label>
			  <input type="radio" checked="checked" name="CLOSING_EXECUTED" value="FALSE">&nbsp;<label for="square-checkbox-2">False</label>
			  <%end if%>
		    <%end if%>
		  </div><br/>
		  
		  <div style=" width:120px; display: inline-block;">
		  Accounting Lock :
		  </div>
		  <div style="display: inline-block;">
			<%if rsGAD("CLOSING_FINAL")="TRUE" Then %>
			  <%if rsGAD("CLOSING_ACC")="TRUE" Then %>
			  <input type="radio" checked="checked" name="CLOSING_ACC" disabled ="TRUE" value="TRUE">&nbsp;<label for="square-checkbox-2">True</label>
			  <input type="radio" name="CLOSING_ACC" disabled ="TRUE" value="FALSE">&nbsp;<label for="square-checkbox-2">False</label>
			  <%else%>
			  <input type="radio" name="CLOSING_ACC" disabled ="TRUE" value="TRUE">&nbsp;<label for="square-checkbox-2">True</label>
			  <input type="radio" checked="checked" name="CLOSING_ACC" disabled ="TRUE" value="FALSE">&nbsp;<label for="square-checkbox-2">False</label>
			  <%end if%>
		    <%else%>	 
			  <%if rsGAD("CLOSING_ACC")="TRUE" Then %>
			  <input type="radio" checked="checked" name="CLOSING_ACC" value="TRUE">&nbsp;<label for="square-checkbox-2">True</label>
			  <input type="radio" name="CLOSING_ACC" value="FALSE">&nbsp;<label for="square-checkbox-2">False</label>
			  <%else%>
			  <input type="radio" name="CLOSING_ACC" value="TRUE">&nbsp;<label for="square-checkbox-2">True</label>
			  <input type="radio" checked="checked" name="CLOSING_ACC" value="FALSE">&nbsp;<label for="square-checkbox-2">False</label>
			  <%end if%>			 
		    <%end if%>
		  </div><br/>

		  <div style=" width:120px; display: inline-block;">
		  ARESTA CLOSING :
		  </div>
		  
		  <div style="display: inline-block;">
			<%if rsGAD("CLOSING_FINAL")="TRUE" Then %>
			  <%if rsGAD("ARESTA_CLOSING")="TRUE" Then %>
				<input type="radio" checked="checked" name="ARESTA_CLOSING" disabled ="TRUE" value="TRUE">&nbsp;<label for="square-checkbox-2">True</label>
				<input type="radio" value="FALSE" disabled ="TRUE" name="ARESTA_CLOSING">&nbsp;<label for="square-checkbox-2">False</label>
			  <%else%>
				<input type="radio" name="ARESTA_CLOSING" disabled ="TRUE" value="TRUE">&nbsp;<label for="square-checkbox-2">True</label>
				<input type="radio" checked="checked" disabled ="TRUE" name="ARESTA_CLOSING" value="FALSE" >&nbsp;<label for="square-checkbox-2">False</label>
			  <%end if%>
		    <%else%>
			   <%if rsGAD("ARESTA_CLOSING")="TRUE" Then %>
				<input type="radio" checked="checked" name="ARESTA_CLOSING" value="TRUE">&nbsp;<label for="square-checkbox-2">True</label>
				<input type="radio" value="FALSE" name="ARESTA_CLOSING">&nbsp;<label for="square-checkbox-2">False</label>
			  <%else%>
				<input type="radio" name="ARESTA_CLOSING" value="TRUE">&nbsp;<label for="square-checkbox-2">True</label>
				<input type="radio" checked="checked" name="ARESTA_CLOSING" value="FALSE" >&nbsp;<label for="square-checkbox-2">False</label>
			  <%end if%>
			<%end if%>  
			
		  </div><br/>

		  <div style=" width:120px; display: inline-block;">
		  KASBANK CLOSING :
		  </div>
		  
		  <div style="display: inline-block;">
			<%if rsGAD("CLOSING_FINAL")="TRUE" Then %>
			  <%if rsGAD("KASBANK_CLOSING")="TRUE" Then %>
				<input type="radio" checked="checked" name="KASBANK_CLOSING" disabled ="TRUE" value="TRUE">&nbsp;<label for="square-checkbox-2">True</label>
				<input type="radio" value="FALSE" disabled ="TRUE" name="KASBANK_CLOSING">&nbsp;<label for="square-checkbox-2">False</label>
			  <%else%>
				<input type="radio" name="KASBANK_CLOSING" disabled ="TRUE" value="TRUE">&nbsp;<label for="square-checkbox-2">True</label>
				<input type="radio" checked="checked" disabled ="TRUE" name="KASBANK_CLOSING" value="FALSE" >&nbsp;<label for="square-checkbox-2">False</label>
			  <%end if%>
		    <%else%>
			   <%if rsGAD("KASBANK_CLOSING")="TRUE" Then %>
				<input type="radio" checked="checked" name="KASBANK_CLOSING" value="TRUE">&nbsp;<label for="square-checkbox-2">True</label>
				<input type="radio" value="FALSE" name="KASBANK_CLOSING">&nbsp;<label for="square-checkbox-2">False</label>
			  <%else%>
				<input type="radio" name="KASBANK_CLOSING" value="TRUE">&nbsp;<label for="square-checkbox-2">True</label>
				<input type="radio" checked="checked" name="KASBANK_CLOSING" value="FALSE" >&nbsp;<label for="square-checkbox-2">False</label>
			  <%end if%>
			<%end if%>  
			
		  </div><br/>

		  <div style=" width:120px; display: inline-block;">
		  STOCK CLOSING :
		  </div>
		  
		  <div style="display: inline-block;">
			<%if rsGAD("CLOSING_FINAL")="TRUE" Then %>
			  <%if rsGAD("STOCK_CLOSING")="TRUE" Then %>
			  <input type="radio" checked="checked" name="STOCK_CLOSING" disabled ="TRUE" value="TRUE">&nbsp;<label for="square-checkbox-2">True</label>
			  <input type="radio" value="FALSE" disabled ="TRUE" name="STOCK_CLOSING">&nbsp;<label for="square-checkbox-2">False</label>
			  <%else%>
			  <input type="radio" name="STOCK_CLOSING" disabled ="TRUE" value="TRUE">&nbsp;<label for="square-checkbox-2">True</label>
			  <input type="radio" checked="checked" disabled ="TRUE" name="STOCK_CLOSING" value="FALSE" >&nbsp;<label for="square-checkbox-2">False</label>
			  <%end if%>
		    <%else%>
			   <%if rsGAD("STOCK_CLOSING")="TRUE" Then %>
			  <input type="radio" checked="checked" name="STOCK_CLOSING" value="TRUE">&nbsp;<label for="square-checkbox-2">True</label>
			  <input type="radio" value="FALSE" name="STOCK_CLOSING">&nbsp;<label for="square-checkbox-2">False</label>
			  <%else%>
			  <input type="radio" name="STOCK_CLOSING" value="TRUE">&nbsp;<label for="square-checkbox-2">True</label>
			  <input type="radio" checked="checked" name="STOCK_CLOSING" value="FALSE" >&nbsp;<label for="square-checkbox-2">False</label>
			  <%end if%>
			<%end if%>  
			
		  </div><br/>

		  <div style=" width:120px; display: inline-block;">
		  WIP CLOSING :
		  </div>
		  
		  <div style="display: inline-block;">
			<%if rsGAD("CLOSING_FINAL")="TRUE" Then %>
			  <%if rsGAD("WIP_CLOSING")="TRUE" Then %>
			  <input type="radio" checked="checked" name="WIP_CLOSING" disabled ="TRUE" value="TRUE">&nbsp;<label for="square-checkbox-2">True</label>
			  <input type="radio" value="FALSE" disabled ="TRUE" name="WIP_CLOSING">&nbsp;<label for="square-checkbox-2">False</label>
			  <%else%>
			  <input type="radio" name="WIP_CLOSING" disabled ="TRUE" value="TRUE">&nbsp;<label for="square-checkbox-2">True</label>
			  <input type="radio" checked="checked" disabled ="TRUE" name="WIP_CLOSING" value="FALSE" >&nbsp;<label for="square-checkbox-2">False</label>
			  <%end if%>
		    <%else%>
			   <%if rsGAD("WIP_CLOSING")="TRUE" Then %>
			  <input type="radio" checked="checked" name="WIP_CLOSING" value="TRUE">&nbsp;<label for="square-checkbox-2">True</label>
			  <input type="radio" value="FALSE" name="WIP_CLOSING">&nbsp;<label for="square-checkbox-2">False</label>
			  <%else%>
			  <input type="radio" name="WIP_CLOSING" value="TRUE">&nbsp;<label for="square-checkbox-2">True</label>
			  <input type="radio" checked="checked" name="WIP_CLOSING" value="FALSE" >&nbsp;<label for="square-checkbox-2">False</label>
			  <%end if%>
			<%end if%>  
			
		  </div><br/>

		  <div style=" width:120px; display: inline-block;">
		  WSVRA CLOSING :
		  </div>

		   <div style="display: inline-block;">
			<%if rsGAD("CLOSING_FINAL")="TRUE" Then %>
			  <%if rsGAD("WSVRA_CLOSING")="TRUE" Then %>
			  <input type="radio" checked="checked" name="WSVRA_CLOSING" disabled ="TRUE" value="TRUE">&nbsp;<label for="square-checkbox-2">True</label>
			  <input type="radio" value="FALSE" disabled ="TRUE" name="WSVRA_CLOSING">&nbsp;<label for="square-checkbox-2">False</label>
			  <%else%>
			  <input type="radio" name="WSVRA_CLOSING" disabled ="TRUE" value="TRUE">&nbsp;<label for="square-checkbox-2">True</label>
			  <input type="radio" checked="checked" disabled ="TRUE" name="WSVRA_CLOSING" value="FALSE" >&nbsp;<label for="square-checkbox-2">False</label>
			  <%end if%>
		    <%else%>
			   <%if rsGAD("WSVRA_CLOSING")="TRUE" Then %>
			  <input type="radio" checked="checked" name="WSVRA_CLOSING" value="TRUE">&nbsp;<label for="square-checkbox-2">True</label>
			  <input type="radio" value="FALSE" name="WSVRA_CLOSING">&nbsp;<label for="square-checkbox-2">False</label>
			  <%else%>
			  <input type="radio" name="WSVRA_CLOSING" value="TRUE">&nbsp;<label for="square-checkbox-2">True</label>
			  <input type="radio" checked="checked" name="WSVRA_CLOSING" value="FALSE" >&nbsp;<label for="square-checkbox-2">False</label>
			  <%end if%>
			<%end if%>  
			
		  </div><br/>

		  <div style=" width:120px; display: inline-block;">
		  PAYROLL CLOSING :
		  </div>
		  
		  <div style="display: inline-block;">
			<%if rsGAD("CLOSING_FINAL")="TRUE" Then %>
			  <%if rsGAD("PAYROLL_CLOSING")="TRUE" Then %>
				<input type="radio" checked="checked" name="PAYROLL_CLOSING" disabled ="TRUE" value="TRUE">&nbsp;<label for="square-checkbox-2">True</label>
				<input type="radio" value="FALSE" disabled ="TRUE" name="PAYROLL_CLOSING">&nbsp;<label for="square-checkbox-2">False</label>
			  <%else%>
				<input type="radio" name="PAYROLL_CLOSING" disabled ="TRUE" value="TRUE">&nbsp;<label for="square-checkbox-2">True</label>
				<input type="radio" checked="checked" disabled ="TRUE" name="PAYROLL_CLOSING" value="FALSE" >&nbsp;<label for="square-checkbox-2">False</label>
			  <%end if%>
		    <%else%>
			   <%if rsGAD("PAYROLL_CLOSING")="TRUE" Then %>
			  <input type="radio" checked="checked" name="PAYROLL_CLOSING" value="TRUE">&nbsp;<label for="square-checkbox-2">True</label>
			  <input type="radio" value="FALSE" name="PAYROLL_CLOSING">&nbsp;<label for="square-checkbox-2">False</label>
			  <%else%>
			  <input type="radio" name="PAYROLL_CLOSING" value="TRUE">&nbsp;<label for="square-checkbox-2">True</label>
			  <input type="radio" checked="checked" name="PAYROLL_CLOSING" value="FALSE" >&nbsp;<label for="square-checkbox-2">False</label>
			  <%end if%>
			<%end if%>  
			
		  </div><br/>

		  <div style=" width:120px; display: inline-block;">
		  UNIT CLOSING :
		  </div>
		  
		  <div style="display: inline-block;">
			<%if rsGAD("CLOSING_FINAL")="TRUE" Then %>
			  <%if rsGAD("UNIT_CLOSING")="TRUE" Then %>
				<input type="radio" checked="checked" name="UNIT_CLOSING" disabled ="TRUE" value="TRUE">&nbsp;<label for="square-checkbox-2">True</label>
				<input type="radio" value="FALSE" disabled ="TRUE" name="UNIT_CLOSING">&nbsp;<label for="square-checkbox-2">False</label>
			  <%else%>
				<input type="radio" name="UNIT_CLOSING" disabled ="TRUE" value="TRUE">&nbsp;<label for="square-checkbox-2">True</label>
				<input type="radio" checked="checked" disabled ="TRUE" name="UNIT_CLOSING" value="FALSE" >&nbsp;<label for="square-checkbox-2">False</label>
			  <%end if%>
		    <%else%>
			   <%if rsGAD("UNIT_CLOSING")="TRUE" Then %>
			  <input type="radio" checked="checked" name="UNIT_CLOSING" value="TRUE">&nbsp;<label for="square-checkbox-2">True</label>
			  <input type="radio" value="FALSE" name="UNIT_CLOSING">&nbsp;<label for="square-checkbox-2">False</label>
			  <%else%>
			  <input type="radio" name="UNIT_CLOSING" value="TRUE">&nbsp;<label for="square-checkbox-2">True</label>
			  <input type="radio" checked="checked" name="UNIT_CLOSING" value="FALSE" >&nbsp;<label for="square-checkbox-2">False</label>
			  <%end if%>
			<%end if%>  
			
		  </div><br/>

		  <div style=" width:120px; display: inline-block;">
		  Payroll Period :
		  </div>
		  <div style="display: inline-block;">
		  <%if aksesfinal = 0 Then %>
		  <input type="text" name="<%=rsGAD("START_DATE").Name%>" readonly ="TRUE" value="<%=rsGAD("START_DATE")%>">&nbsp;&nbsp;
		  <input type="text" name="<%=rsGAD("End_date").Name%>" readonly ="TRUE" value="<%=rsGAD("End_date")%>"><br/><br/>
		  <%else%>
		  <input type="text" name="<%=rsGAD("START_DATE").Name%>" value="<%=rsGAD("START_DATE")%>">&nbsp;&nbsp;
		  <input type="text" name="<%=rsGAD("End_date").Name%>" value="<%=rsGAD("End_date")%>"><br/><br/>
		   <%end if%>
		  </div><br/>
		 
 <%if aksesfinal <> 0 then%>
		 <div style=" width:120px; display: inline-block;">
		  FINAL LOCK :
		  </div>
		  <div style="display: inline-block;">
		  <%if rsGAD("CLOSING_FINAL")="TRUE" Then %>
		  <input type="radio" checked="checked" name="CLOSING_FINAL" value="TRUE" disabled ="TRUE">&nbsp;<label for="square-checkbox-2">True</label>
		  <input type="radio" name="CLOSING_FINAL" value="FALSE" disabled ="TRUE">&nbsp;<label for="square-checkbox-2">False</label>
		  <%else%>
		  <input type="radio" name="CLOSING_FINAL" value="TRUE">&nbsp;<label for="square-checkbox-2">True</label>
		  <input type="radio" checked="checked" name="CLOSING_FINAL" value="FALSE">&nbsp;<label for="square-checkbox-2">False</label>
		  <%end if%>
		  
		  </div><br/>	
<%end if%>		  

          <%if rsGAD("CLOSING_FINAL")="FALSE" Then %>
		  <div style="display: inline-block;">
		  <input onclick="return confirm('Are you sure you want to Update this record?');" style="font-family:'FontAwesome'; padding:3px 5px;" type="submit" value="&#xf0c7;">&nbsp;&nbsp;
		  <input style="font-family:'FontAwesome'; padding:3px 5px;" type="button" onclick="javascript:toggle('editing_form_<%=NO%>','edit<%=NO%>')" value="&#xf05e;">
		  </div><br/>	
          <%end if%>	  


		 </form>
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
	
	

			<!-- Akhir Tampilan Select form -->

		<input  type="hidden" value="<%=txtTable%>" name="txtTable">
		<input  type="hidden" value="<%=txtModule%>" name="txtModule">
		<!--</form>-->

<!--</div>-->

	<div style="background-color:#f1f1f1; padding:0.3%; border-top:1px #ccc solid; border-bottom:1px #ccc solid; position:fixed; top:114px; right:0.3%; width:76.8%;">
		<!--<a class="add_list" href="#" onclick="showWindow('yellow');" value="Window - yellow"
		title="New : <%=txtTable%>"><i class="fa fa-plus-circle"></i><span>ADD LIST</span></a>-->
		<a href="#" class="add_list" onclick="reloadPage()"><i class="fa fa-refresh"></i><span>Refresh </span></a>&nbsp;
		<form style="padding:0px; border:none;display:inline;margin:0px;" method="post" action="EmpalldedStatus.asp?txtTable=<%=request("txtTable")%>"> 
		<input class="search" type="text" name="Searching" value="<%=Searching%>"">
		<input class="button_search" type="submit" value=""> 
		<i class="fa fa-search icon_search_data"></i>
		<a href="EmpalldedStatus.asp?txtTable=<%=request("txtTable")%>&txtJenis=generate"" class="add_list"><i class="fa fa-flask"></i><span>Generate </span></a>&nbsp;
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
		<a class="button_prev" href="EmpalldedStatus.asp?txtTable=<%=request("txtTable")%>&iPageNo=<%=sebelum%>"><i class="fa fa-chevron-left"></i></a>
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
					 Response.Write "<a class=""button_page"" href=""EmpalldedStatus.asp?txtTable="& request("txtTable")& "&iPageNo=" &  i & _
					  """>" & i & "</a> "
					 end if
			  Next
		   End if
		   
		   if int(Request("iPageNo")) = intNumberOfPages then %>
			<span class="button_next"><i class="fa fa-chevron-right"></i></span>
			<%else%>
			<a class="button_next" href="EmpalldedStatus.asp?txtTable=<%=request("txtTable")%>&iPageNo=<%=lanjut%>"><i class="fa fa-chevron-right"></i></a>
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
<script type="text/javascript">

			var result = $("#lenght_menu").height();
			var result = result - 5;
            $('.flexme').flexigrid({
				height:result,
				combo_flag: false

			});
        </script>
</body>
</html>
