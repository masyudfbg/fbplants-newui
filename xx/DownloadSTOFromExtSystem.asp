<%
'Option Explicit
%>
<%
'On Error resume Next
%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<html>
<head>
<title>ePlaS - Download Stock Transfer Order From Ext System</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js" ></script>
<link rel="StyleSheet" href="Table_Html.css" type="text/css" /> <!-- For Hidden Column UNIQUE_KEY  ON TABLE HTML, Because Field UNIQUE_KEY Not View on Table HTML, Create BY AST 20180514 -->
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
<script language="JavaScript" src="/js/date-picker.js"></script>  
<script language="JavaScript"> 
var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}


function SetCheckBox(CbCode, CbOriCode){    

    var StringAll;
	var value_;


	var CID = document.getElementById(CbOriCode).checked;
			
	if(document.getElementById(CbOriCode).checked == 1) {
		StringAll = "<input id='data_" + CbCode + "' name='data_" + CbCode + "' type='hidden' value='on'>";
	}
	else {
		StringAll = "<input id='data_" + CbCode + "' name='data_" + CbCode + "' type='hidden' value='off'>";
	}
	
    document.getElementById("div_" + CbCode ).innerHTML=StringAll;
}

function SetCheckBoxDefault(CbCode, CbOriCode){    

    var StringAll;
	var value_;


	var CID = document.getElementById(CbOriCode).checked;
			
	if(document.getElementById(CbOriCode).checked == 1) {
		StringAll = "<input id='data_default_" + CbCode + "' name='data_default_" + CbCode + "' type='hidden' value='on'>";
	}
	else {
		StringAll = "<input id='data_default_" + CbCode + "' name='data_default_" + CbCode + "' type='hidden' value='off'>";
	}
	
    document.getElementById("div_default_" + CbCode ).innerHTML=StringAll;
}

function checkAll(formname, checktoggle, vNilai)
{

	var StringCheck;
	var checkboxes = new Array(); 
		 
	checkboxes = document[formname].getElementsByClassName('ok');
	var myTab = document.getElementById('entry');
		
	for (var i=0; i <= checkboxes.length; i++)  {		
		var objCells = myTab.rows.item(i+1).cells; /*Setting ID yang ingin di ambil dari kolom berapa..*/
		/*=====================================*/						
		/*Pengambilan ID untuk filter*/
		var id_key = objCells.item(1).innerHTML;   /* Get ID_Key */
		var string1 = id_key;
		/*=====================================*/	
		if (checkboxes[i].type == 'checkbox') {
			checkboxes[i].checked = checktoggle;
			
			if (checkboxes[i].checked == true) {
				//alert('True : '+i+','+checkboxes.length);
				StringCheck = "<input id='data_" + string1 + "' name='data_" + string1 + "' type='hidden' value='on'>";
			}else{
				//alert('False : '+i+','+checkboxes.length);
				StringCheck = "<input id='data_" + string1 + "' name='data_" + string1 + "' type='hidden' value='off'>";
			}
									
		}
			//console.log(document.getElementsByClassName("test")[i].innerHTML=StringCheck);
			document.getElementsByClassName("test")[i].innerHTML=StringCheck;
	}
		
}

/*Selected ALL Checkbox Event */
function SelectedAll_Checkbox(formname)
{
	var checktoggle;
	var vNilai;
	
	if (cb_SelectedAll.checked) { 
		checktoggle = true;
		vNilai = 'on';
	}else{
		checktoggle = false;
		vNilai = 'off';
	}
	
	console.log(checktoggle + vNilai);

	var StringCheck;
	var checkboxes = new Array(); 
		 
	checkboxes = document[formname].getElementsByClassName('ok');
	var myTab = document.getElementById('entry');
		
	for (var i=0; i <= checkboxes.length; i++)  {		
		var objCells = myTab.rows.item(i+1).cells; /*Setting ID yang ingin di ambil dari kolom berapa..*/
		/*=====================================*/						
		/*Pengambilan ID untuk filter*/
		var id_key = objCells.item(1).innerHTML;   /* Get ID_Key */
		var string1 = id_key;
		
		/*Pengambilan ID untuk filter*/
		/*
		var Pono = objCells.item(4).innerHTML;
		var LineNo = objCells.item(8).innerHTML;
		var Dono = objCells.item(14).innerHTML;
		var string1 = Pono + LineNo + Dono;
		*/
		/*=====================================*/	
		if (checkboxes[i].type == 'checkbox') {
			checkboxes[i].checked = checktoggle;
			
			if (checkboxes[i].checked == true) {
				//alert('True : '+i+','+checkboxes.length);
				StringCheck = "<input id='data_" + string1 + "' name='data_" + string1 + "' type='hidden' value='on'>";
			}else{
				//alert('False : '+i+','+checkboxes.length);
				StringCheck = "<input id='data_" + string1 + "' name='data_" + string1 + "' type='hidden' value='off'>";
			}
									
		}
			//console.log(document.getElementsByClassName("test")[i].innerHTML=StringCheck);
			document.getElementsByClassName("test")[i].innerHTML=StringCheck;
	}
	
}


</script> 
</head>
<body>
<!--#include file="top.jsp"-->
<div style=" width:21.7%; float:left;margin-left:0.3%; margin-top:0.3%; margin-bottom:0.3%;">
		<% if Session("fcBa") <> "" then %>
         <!--#include file="menu.asp"-->
        <% end if %>
	</div>
    <div class="judul_head" valign="top" bgcolor="#fff" style=" width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;"> 
    <div class="judul">
	   Module : <%=request("txtModule")%> / <%=request("txtTable")%>
	</div>
    <div class="box_both" style="overflow:scroll;"> 
 <%
 '-------------awal-----------------------
 %>
 <%


Dim v_FCENTRY, v_FCEDIT, v_FCIP
DIM v_FCCODE
Dim cmdView
Dim defaultCompany, qCompanyAll
Dim tempCounterDefault, doneContinue
Dim rsModule,qModule,qInsert, rsModuleCheck
Dim txtFcba,v_FCBA,txtpono,v_TxtPONO
Dim v_txtErrorFlag, txtErrorFlag
		

v_TxtPONO = request("txtpono")
v_FCBA    = request("txtFcba")
v_txtErrorFlag   = request("txtErrorFlag")

If v_TxtPONO <> "--ALL--" then
    whereCons = whereCons & " And  p.pono = '" & v_TxtPONO & "' "
end if

If v_FCBA <> "--ALL--" then
    whereCons = whereCons & " And  BusinessUnitExt_Api.GetFcba(p.PLANT,p.STORLOC) = '" & v_FCBA & "' "
end if

if v_txtErrorFlag <> "--ALL--" then
   whereCons = whereCons & " And  p.ERRORFLAG = '" & v_txtErrorFlag & "' "
end if

 if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then

	if Request("txtJenis")= "Delete" then
	    q = " delete from user_fcba_tab " & _
			" where " & _
			" fcusercode = '" & request("v_FCCODE") & "'"	

		   connStr.execute(q)
		   cmdView	= "True"
	elseif Request("txtJenis")="New" or Request("txtJenis")="Edit" then
	
		   qCompanyAll = "  Select BusinessUnitExt_Api.GetFcba(p.PLANT,p.STORLOC) Fcba,p.* " & _
						 "  from bgprodsapiplas.PurchaseOrder p " & _
						 "  where p.bridgestate=0 And p.Iplasstate=1  And  " & _
						 "  p.IsRelease=1  And p.potype='STO' and p.appsys='IPLAS' " & whereCons
			  
	    Set rsModuleCheck = Server.CreateObject("ADODB.Recordset")
		rsModuleCheck.CursorType = 3
		rsModuleCheck.CursorLocation = 3
		rsModuleCheck.Open qCompanyAll, connstr
		Dim msgresult
		msgresult =""
		doneContinue = "TRUE"
		tempCounterDefault = 0
		if Not rsModuleCheck.eof then 
		    While Not rsModuleCheck.EOF
			    v_UnixCode  = rsModuleCheck("UNIQUEKEY") 
				'response.Write v_UnixCode & " "
				if request(trim("data_" & v_UnixCode))="on" Then
					
					'response.Write "BEGIN BridgingSAP_PRPOGRN_API.InsertSTO('" & rsModuleCheck("pono") & "'); END; "
                    msgresult =  msgresult & rsModuleCheck("PONO") & ", "

					strQueryStatus = "BEGIN BridgingSAP_PRPOGRN_API.InsertSTO('" & rsModuleCheck("PONO") & "'); END; "
	                connStr.execute(strQueryStatus)	 
				end if
				rsModuleCheck.MoveNext
		    wend	
		end if	
		
		rsModuleCheck.Close
		set rsModuleCheck = Nothing
		if msgresult <> ""  then
		  response.Write "<font face=Verdana size=1 color=blue>Download Stock Transport Order (STO) ," & msgresult & " From External System Completed..</font>"
		end if
		cmdView = "True"
		
    else
        cmdView	= "True"
	end if								 
	'Response.Write "sql " & q
else
    cmdView	= request("cmdView")
End if
Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis
if Request("PeriodDate")<>"" then
   PeriodDate   =Request("PeriodDate")
else
   PeriodDate   =textDate(date)
end if

txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")
'fcCode=request("fcCode")
txtJenis = request("txtJenis")
If Err.Number <> 0 then
   Response.Write "<font color=red>" & "Error Number : " & Err.Number &  " Desc. " & Err.Description & "</font>"
End if

 '-------------ISI Payment Form
 %>
 
 <form name="frmEntry" method="post" action="DownloadSTOFromExtSystem.Asp">
	<table class="data" width="500">
		<h2>DOWNLOAD STOCK TRANSPORT ORDER FROM EXTERNAL SYSTEM</h2>
				
	        &nbsp;&nbsp;
		<tr>
		  <td><font face="verdana" size="1">Fcba</font></td>
		  <td>
			<select name="txtFcba" >
			<option value="--ALL--">--ALL--</option>
			   <%
				Dim rsFcba, strFcba
				strFcba =   " Select distinct BusinessUnitExt_Api.GetFcba(p.PLANT,p.STORLOC) Fcba " & _
				             "  from bgprodsapiplas.PurchaseOrder p " & _
							 "  where p.bridgestate=0 And p.Iplasstate=1  And  " & _
				             "  p.IsRelease=1   And p.potype='STO' and p.appsys='IPLAS'"
								
				Set rsFcba = Server.CreateObject("ADODB.Recordset")
				rsFcba.Open strFcba, connstr, adOpenStatic
							
				if not rsFcba.eof then
				   Do While Not rsFcba.eof
					  if v_txtFcba = rsFcba("Fcba") then %>
							<option value="<%=rsFcba("Fcba")%>" Selected><%=rsFcba("Fcba")%></option>
					<%else%>
					    	<option value="<%=rsFcba("Fcba")%>"><%=rsFcba("Fcba")%></option>
	    			<%end if			
					   rsFcba.movenext
	   			   loop
				end if
				rsFcba.Close
				set rsFcba = Nothing
				%>
			</select>
	       </td>
		   <td><font face="verdana" size="1">PO No</font></td>
		   <td>
			<select name="TxtPONO" >
			<option value="--ALL--">--ALL--</option>
			   <%
				Dim rspono, strpono
				strpono =   " Select distinct p.Pono " & _
				             "  from bgprodsapiplas.PurchaseOrder p " & _
							 "  where p.bridgestate=0 And p.Iplasstate=1  And  " & _
				             "  p.IsRelease=1   And p.potype='STO' and p.appsys='IPLAS' order by p.Pono"
								
				Set rspono = Server.CreateObject("ADODB.Recordset")
				rspono.Open strpono, connstr, adOpenStatic
							
				if not rspono.eof then
				   Do While Not rspono.eof
					  if v_TxtPONO = rspono("Pono") then %>
							<option value="<%=rspono("Pono")%>" Selected><%=rspono("Pono")%></option>
					<%else%>
					    	<option value="<%=rspono("Pono")%>"><%=rspono("Pono")%></option>
	    			<%end if			
					   rspono.movenext
	   			   loop
				end if
				rspono.Close
				set rspono = Nothing
				%>
			</select>
	       </td>
		   
			<td><font face="verdana" size="1">ERROR FLAG</font></td>
			<td>
				<select name="txtErrorFlag" >
				<!-- <option value="--ALL--">--ALL--</option> -->
				   <%
					strError = " SELECT distinct '--ALL--' as errorflag " & _
							   " FROM DUAL " & _
							   " UNION ALL " & _
							   " Select distinct TO_CHAR(b.errorflag) as errorflag " & _
							   " From bgprodsapiplas.PurchaseOrder b " & _
							   " where b.bridgestate = 0 " & _
							   " And b.appsys = 'IPLAS' " & _
							   " AND b.potype = 'STO' " & _
							   " AND b.Iplasstate = 1 " & _
							   " AND b.IsRelease = 1 "
									
					Set rsEF = Server.CreateObject("ADODB.Recordset")
					rsEF.Open strError, connstr, adOpenStatic
								
					if not rsEF.eof then
					   Do While Not rsEF.eof
						  if v_txtErrorFlag = rsEF("errorflag") then %>
								<option value="<%=rsEF("errorflag")%>" Selected><%=rsEF("errorflag")%></option>
						<%else%>
								<option value="<%=rsEF("errorflag")%>"><%=rsEF("errorflag")%></option>
						<%end if			
						   rsEF.movenext
					   loop
					end if
					rsEF.Close
					set rsEF = Nothing
					%>
				</select>
			</td>
		   
		  <td>
			<input Type="Hidden" Value="True" Name="cmdView">
			<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
			<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
			<input Type="Submit"  name="Add" value="Select">
		  </td>
		</tr>
    </table>  
 </form>
<%

   txtCekUsers="True"


if cmdView = "True" and txtCekUsers="True" Then 
   Dim rsVH,intRecordsPerPage,cColor
   set rsVH = server.CreateObject("Adodb.recordset")
   'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
  
	  
             q = "  Select p.UNIQUEKEY AS ID_KEY, BusinessUnitExt_Api.GetFcba(p.PLANT,p.STORLOC) Fcba,p.* " & _
				 "  from bgprodsapiplas.PurchaseOrder p " & _
				 "  where p.bridgestate=0 And p.Iplasstate=1  And  " & _
				 "  p.IsRelease=1  And p.potype='STO' and p.appsys='IPLAS' " & whereCons
                    
   Set rsVH = Server.CreateObject("ADODB.Recordset")
   rsVH.CursorType = 3
   rsVH.CursorLocation = 3
   rsVH.Open q, connstr%>
   <form action="DownloadSTOFromExtSystem.Asp" method="post" name="eForm">
   <table id="entry" class="entry" cellpadding="1" cellspacing="1" width="620">
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;">
		<!--<td align="center" valign="top" >DOWNLOAD</td>-->
		<td align="center" valign="top" width="60px" class="coba">
			<input 	type="checkbox" id="cb_SelectedAll" name="cb_SelectedAll" onchange="SelectedAll_Checkbox('eForm')" /> DOWNLOAD ALL
		</td>		
		<td id="para1" class="para1" align="center" valign="top" >ID_KEY</td>				
		<td align="center" valign="top" >NO</td>
		<td align="center" valign="top" >COMPANY</td>
		<td align="center" valign="top" >FCBA</td>
		<td align="center" valign="top" >PONO</td>
		<td align="center" valign="top" >VENDOR</td>
		<td align="center" valign="top" >POTYPE</td>
		<td align="center" valign="top" >PODATE</td>
		<td align="center" valign="top" >LINENO</td>
		<td align="center" valign="top" >MATERIAL</td>
		<td align="center" valign="top" >MATERIALDESC</td>
		<td align="center" valign="top" >PLANT</td>
		<td align="center" valign="top" >STORLOC</td>
		<td align="center" valign="top" >POQTY</td>
		<td align="center" valign="top" >DONO</td>
		<td align="center" valign="top" >EXTCTRNO</td>
		<td align="center" valign="top" >DOCTYPE</td>
		<td align="center" valign="top" >DOCTYPEDESC</td>
		<td align="center" valign="top" >PURCHORG</td>
		<td align="center" valign="top" >PURCHGRP</td>
		<td align="center" valign="top" >PAYTERM</td>
		<td align="center" valign="top" >INCO1</td>
		<td align="center" valign="top" >INCO2</td>
		<td align="center" valign="top" >CURR</td>
		<td align="center" valign="top" >CURRRATE</td>
		<td align="center" valign="top" >ACCCAT</td>
		<td align="center" valign="top" >ITCAT</td>
		<td align="center" valign="top" >MATGRP</td>
		<td align="center" valign="top" >UOM</td>
		<td align="center" valign="top" >DELIVDATE</td>
		<td align="center" valign="top" >NETPRICE</td>
		<td align="center" valign="top" >PER</td>
		<td align="center" valign="top" >PERUOM</td>
		<td align="center" valign="top" >TRACKNO</td>
		<td align="center" valign="top" >TRACKLINENO</td>
		<td align="center" valign="top" >PRNO</td>
		<td align="center" valign="top" >PRLINENO</td>
		<td align="center" valign="top" >AFCENO</td>
		<td align="center" valign="top" >INTORDER</td>
		<td align="center" valign="top" >GLACCOUNT</td>
		<td align="center" valign="top" >ISDELETE</td>
		<td align="center" valign="top" >ISCOMPLETE</td>
		<td align="center" valign="top" >ISRELEASE</td>
		<td align="center" valign="top" >BRIDGESTATE</td>
		<td align="center" valign="top" >IPLASSTATE</td>
		<td align="center" valign="top" >APPSYS</td>
		<td align="center" valign="top" >ERRORFLAG</td>
		<td align="center" valign="top" >ERRORDESC</td>

	</tr>
   <%
   if Not rsVH.eof then
      nCountofField=rsVH.fields.count
	  no=1
	  While Not rsVH.EOF

		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if
		  
          v_UnixCode  = rsVH("UNIQUEKEY")  
          'response.Write v_UnixCode
		  %>
		  <tr bgcolor="<%=cColor%>">
			  
			  
			<td align="center" valign="top" title="Please,Selected" id="coba" class="coba">
				  <input class="ok" name="cb_<%=v_UnixCode%>" id="cb_<%=v_UnixCode%>" type="checkbox" value="off" onClick="SetCheckBox('<%=v_UnixCode%>', 'cb_<%=v_UnixCode%>')">
				  
				  <div class="test" id="div_<%=v_UnixCode%>" name="div_<%=v_UnixCode%>">
					 <input id="data_<%=v_UnixCode%>" name="data_<%=v_UnixCode%>" type="hidden" value="off" >
				  </div>			
			</td>
			<td id="para1" class="para1" align="Right" valign="top" ><%=rsVH("ID_KEY")%></td>			
            <td align="Right" valign="top"><%=NO%></td>
			<td align="Right" valign="top"><%=rsVH("FCBA")%></td>
			<td align="Right" valign="top"><%=rsVH("COMPANY")%></td>
			<td align="Right" valign="top"><%=rsVH("PONO")%></td>
			<td align="Right" valign="top"><%=rsVH("VENDOR")%></td>
			<td align="Right" valign="top"><%=rsVH("POTYPE")%></td>
			<td align="Right" valign="top"><%=rsVH("PODATE")%></td>
			<td align="Right" valign="top"><%=rsVH("LINENO")%></td>
			<td align="Right" valign="top"><%=rsVH("MATERIAL")%></td>
			<td align="Right" valign="top"><%=rsVH("MATERIALDESC")%></td>
			<td align="Right" valign="top"><%=rsVH("PLANT")%></td>
			<td align="Right" valign="top"><%=rsVH("STORLOC")%></td>
			<td align="Right" valign="top"><%=rsVH("POQTY")%></td>
			<td align="Right" valign="top"><%=rsVH("DONO")%></td>
			<td align="Right" valign="top"><%=rsVH("EXTCTRNO")%></td>
			<td align="Right" valign="top"><%=rsVH("DOCTYPE")%></td>
			<td align="Right" valign="top"><%=rsVH("DOCTYPEDESC")%></td>
			<td align="Right" valign="top"><%=rsVH("PURCHORG")%></td>
			<td align="Right" valign="top"><%=rsVH("PURCHGRP")%></td>
			<td align="Right" valign="top"><%=rsVH("PAYTERM")%></td>
			<td align="Right" valign="top"><%=rsVH("INCO1")%></td>
			<td align="Right" valign="top"><%=rsVH("INCO2")%></td>
			<td align="Right" valign="top"><%=rsVH("CURR")%></td>
			<td align="Right" valign="top"><%=rsVH("CURRRATE")%></td>
			<td align="Right" valign="top"><%=rsVH("ACCCAT")%></td>
			<td align="Right" valign="top"><%=rsVH("ITCAT")%></td>
			<td align="Right" valign="top"><%=rsVH("MATGRP")%></td>
			<td align="Right" valign="top"><%=rsVH("UOM")%></td>
			<td align="Right" valign="top"><%=rsVH("DELIVDATE")%></td>
			<td align="Right" valign="top"><%=rsVH("NETPRICE")%></td>
			<td align="Right" valign="top"><%=rsVH("PER")%></td>
			<td align="Right" valign="top"><%=rsVH("PERUOM")%></td>
			<td align="Right" valign="top"><%=rsVH("TRACKNO")%></td>
			<td align="Right" valign="top"><%=rsVH("TRACKLINENO")%></td>
			<td align="Right" valign="top"><%=rsVH("PRNO")%></td>
			<td align="Right" valign="top"><%=rsVH("PRLINENO")%></td>
			<td align="Right" valign="top"><%=rsVH("AFCENO")%></td>
			<td align="Right" valign="top"><%=rsVH("INTORDER")%></td>
			<td align="Right" valign="top"><%=rsVH("GLACCOUNT")%></td>
			<td align="Right" valign="top"><%=rsVH("ISDELETE")%></td>
			<td align="Right" valign="top"><%=rsVH("ISCOMPLETE")%></td>
			<td align="Right" valign="top"><%=rsVH("ISRELEASE")%></td>
			<td align="Right" valign="top"><%=rsVH("BRIDGESTATE")%></td>
			<td align="Right" valign="top"><%=rsVH("IPLASSTATE")%></td>
			<td align="Right" valign="top"><%=rsVH("APPSYS")%></td>
			<td align="Right" valign="top"><%=rsVH("ERRORFLAG")%></td>
			<td align="Right" valign="top"><%=rsVH("ERRORDESC")%></td>
		  </tr>
		  <%
		  
		  
		  rsVH.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   end if 
   rsVH.Close
   set rsVH = Nothing%>
	<tr bgcolor="#EEEEEE">
	<!--
		<td colspan="4" align="center">
			<a onclick="javascript:checkAll('eForm', true, 'on');" href="javascript:void();">check all</a> &nbsp &nbsp &nbsp
			<a onclick="javascript:checkAll('eForm', false, 'off');" href="javascript:void();">uncheck all</a>
		</td>
	-->
	
	<td colspan="6" align="center">
	  <input type="hidden" name="txtpono" value="<%=v_TxtPONO%>">
	  <input type="hidden" name="txtFcba" value="<%=v_FCBA%>">
	  <input type="hidden" name="txtErrorFlag" value="<%=v_txtErrorFlag%>">
		  
	  <input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
	  <input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
	  <input type="hidden" name="txtJenis" value="Edit">
	  <input Type="Submit"  name="Add" value="Submit" class="IgnoreButton">
	  <!--#include file="./print_excel_update.asp"-->
	</td>
  </tr>

   </TABLE>
   </div>
   </form>
<%elseif cmdView = "True" and txtCekRoles="False" Then %>
   <P><font face="Verdana" size="1" color="#FF0000">User Name Not Found.. Please entry from Data Setup</P>
<%End if%>  
 
  <%
  '-------------akhir----------------------
  %>
</div></div>
  <!--#include file="bottom.jsp"-->

</body>
</html>
