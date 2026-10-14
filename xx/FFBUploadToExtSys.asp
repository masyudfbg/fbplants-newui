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
<title>ePlaS - FFB Upload To Ext System</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js" ></script>
<link rel="StyleSheet" href="Table_Html.css" type="text/css" /> <!-- For Hidden Column ID_KEY  ON TABLE HTML, Because Field ID_KEY Not View on Table HTML, Create BY AST 20180514 -->
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


Dim v_FCENTRY, v_FCEDIT, v_FCIP, v_FCBA
DIM v_FCCODE
Dim cmdView
Dim defaultCompany, qCompanyAll
Dim tempCounterDefault, doneContinue
Dim rsModule,qModule,qInsert, rsModuleCheck
		


dim whereCons,v_txtfcba,txtfcba
		



 whereCons =""
 v_txtfcba	    = request("txtfcba")
 v_txtDateFrom	= trim(Request("txtDateFrom"))
 v_txtDateTo	= trim(Request("txtDateTo"))
 
 response.write Request("txtJenis")
 
 If v_txtfcba <> "--ALL--" then
    whereCons = whereCons & " And  WB.FCBA = '" & v_txtfcba & "' "
 end if
 if  len(trim(v_txtDateFrom)) > 0 then
	whereCons = whereCons & "  AND WB.POSTINGDATE >=	to_Date('" & v_txtDateFrom & "','MM/DD/YYYY')"
	else
	whereCons = whereCons& "  AND WB.POSTINGDATE >=	to_Date('" & DateAdd("YYYY",-1,date) & "','MM/DD/YYYY')"
end if
if  len(trim(v_txtDateTo)) > 0 then
	whereCons = whereCons &  " AND WB.POSTINGDATE <= to_Date('" & v_txtDateTo & "','MM/DD/YYYY')"
	else
	whereCons = whereCons &  " AND WB.POSTINGDATE <= to_Date('" & date & "','MM/DD/YYYY')"
end if
 

 if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then

	if Request("txtJenis")="New" or Request("txtJenis")="Edit" then
	
		   qCompanyAll = " SELECT  WB.CHITNO," & _
						 "         Businessunit_api.get_SapLocation(WB.FCBA) Location," & _
						 "		   WB.POSTINGDATE," & _
						 "		   WB.TIMEOUT," & _
						 "		   CU.DESCRIPTION CUSTOMERNAME," & _
						 "		  ( CASE WB.FIELD_ESTATE_ESTATECODE WHEN 'LUAR' THEN" & _
						 "          WB.CUST_CUSTOMERCODE  " & _
						 "			 ELSE " & _
						 "				WB.FIELD_ESTATE_ESTATECODE " & _
						 "			 END )  CUSTOMERCODE," & _
						 "		   'TBS' BATCH," & _
						 "		   WB.WEIGHTIN  BRUTOQTY," & _
						 "         WB.WEIGHTOUT TARRAQTY," & _
						 "         NVL(GR.WEIGHT, 0) SORTASE," & _
                         "        (WB.WEIGHTIN - WB.WEIGHTOUT) - NVL(GR.WEIGHT, 0) AFTSRTQTY," & _
						 "		   WB.VEHICLENO," & _
						 "		   WB.DRIVERNAME," & _
						 "		   WB.SPBNO," & _
						 "		   ''  Grade," & _
						 "		   ''  rcpo," & _
						 "		   NVL(GR.BUNCHES, 0) JJG," & _
						 "		   GR.F00 unripbunch," & _
						 "		   GR.F5  rottenbunch," & _
						 "		   GR.Tkosong Emptybunch," & _
						 "		   GR.Tpanjang Tallbunch," & _
						 "		   NVL(GR.Air, 0) AIR," & _
						 "		   NVL(GR.Sampah, 0) Sampah," & _
						 "		   NVL(GR.Pasir, 0) Pasir," & _
						 "		   nvl(GR.EATENBYRAT, 0) + nvl(GR.Brondolan, 0) + nvl(GR.Bkecil, 0) Lain," & _
						 "		   NVL(GR.dura, 0) Dura," & _
						 "  	   'A01'  GRtype," & _
						 "		   ''      NegoNo," & _
						 "		   '' STO," & _
						 "		   '' STOLINE," & _
						 "		   0  price," & _
						 "		   '' Curr," & _
						 "		   0 Bridgestate," & _
						 "		   GR.Remarks" & _	
						 " FROM WEIGHBRIDGETICKET WB" & _
						 " LEFT JOIN GRADINGRECORD GR ON (WB.FCBA = GR.FCBA AND WB.CHITNO = GR.WBT_CHITNO)" & _
						 " LEFT JOIN CUSTOMER CU ON (WB.FCBA = CU.FCBA AND WB.CUST_CUSTOMERCODE = CU.CUSTOMERCODE)" & _
						 " WHERE WB.PRODUCT_PRODUCTCODE = 'FFB' AND WB.isinbridge <> '1' "  & whereCons & " " & _
						 " ORDER BY WB.CHITNO"	
		
		'response.Write	 qCompanyAll 
		
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
			    v_UnixCode  = rsModuleCheck("CHITNO") & "_" & rsModuleCheck("LOCATION")
				if request(trim("data_" & v_UnixCode))="on" Then
					'response.Write v_UnixCode
					msgresult = msgresult & ", " & v_UnixCode
					strQueryStatus = "BEGIN BRIDGINSAP_FFB_API.SendFFBToExtSys('" & rsModuleCheck("CHITNO") & "','" & rsModuleCheck("LOCATION") & "'); END; "
	                connStr.execute(strQueryStatus)	
				end if
				rsModuleCheck.MoveNext
		    wend	
		end if	
		
		rsModuleCheck.Close
		set rsModuleCheck = Nothing
		If msgresult <> ""  then
			response.Write "<font face=Verdana size=1 color=blue>Update FFB Ticket No "& msgresult &" From External System Completed..</font>"
		End If
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
 
 <form name="frmEntry" method="post" action="FFBUploadToExtSys.Asp">
	<table class="data" width="500">
		
		<h2>UPDATE FFB UPLOAD TO EXT SYSTEM</h2>
				
	        &nbsp;&nbsp;
		<tr>
			<td><font face="verdana" size="1">FCBA</font></td>
			<td>
				<select name="txtfcba" >
				<option value="--ALL--">--ALL--</option>
				   <%
					Dim rsfcba, strfcba
					strfcba =  "select b.fccode,b.fcname from businessunit b where b.fctype='M' and b.isactive ='TRUE'  "
									
					Set rsfcba = Server.CreateObject("ADODB.Recordset")
					rsfcba.Open strfcba, connstr, adOpenStatic
								
					if not rsfcba.eof then
					   Do While Not rsfcba.eof
						  if v_txtfcba = rsfcba("fccode") then %>
								<option value="<%=rsfcba("fccode")%>" Selected><%=rsfcba("fccode")%> - <%=rsfcba("fcname")%></option>
						<%else%>
								<option value="<%=rsfcba("fccode")%>"><%=rsfcba("fccode")%> - <%=rsfcba("fcname")%></option>
						<%end if			
						   rsfcba.movenext
					   loop
					end if
					rsfcba.Close
					set rsfcba = Nothing
					%>
				</select>
			</td>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<td>
			    <font face="verdana" size="1">Posting Date From</font>
			</td>
			<td align="center">
				<% if v_txtDateFrom <> "" and v_txtDateFrom = v_txtDateFrom then %>
					<input class="datepicker" type="text" maxlength="10" name="txtDateFrom" size="10" value="<%=v_txtDateFrom%>" title="Type Date mm/dd/yyyy" >
				<% else %>
					<input class="datepicker" type="text" maxlength="10" name="txtDateFrom" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
				<% end if%>
					
			</td>
			<td>
			    <font face="verdana" size="1">Posting Date To</font>
			</td>
			<td align="center">
				<% if v_txtDateTo <> "" and v_txtDateTo = v_txtDateTo then %>
					<input class="datepicker" type="text" maxlength="10" name="txtDateTo" size="10" value="<%=v_txtDateTo%>" title="Type Date mm/dd/yyyy" >
				<% else %>
					<input class="datepicker" type="text" maxlength="10" name="txtDateTo" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
				<% end if%>
					
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
  
	  
             q = " SELECT  WB.CHITNO, CONCAT(CONCAT(WB.CHITNO,'_'),Businessunit_api.get_SapLocation(WB.FCBA)) as ID_KEY," & _
                 "         Businessunit_api.get_SapLocation(WB.FCBA) Location," & _
				 "		   WB.POSTINGDATE," & _
				 "		   WB.TIMEOUT," & _
				 "		   CU.DESCRIPTION CUSTOMERNAME," & _
				 "		  ( CASE WB.FIELD_ESTATE_ESTATECODE WHEN 'LUAR' THEN" & _
				 "          WB.CUST_CUSTOMERCODE  " & _
				 "			 ELSE " & _
				 "				WB.FIELD_ESTATE_ESTATECODE " & _
				 "			 END )  CUSTOMERCODE," & _
				 "		   'TBS' BATCH," & _
				 "		   WB.WEIGHTIN  BRUTOQTY," & _
                 "         WB.WEIGHTOUT TARRAQTY," & _
				 "         NVL(GR.WEIGHT, 0) SORTASE," & _
                 "        (WB.WEIGHTIN - WB.WEIGHTOUT) - NVL(GR.WEIGHT, 0) AFTSRTQTY," & _
				 "		   WB.VEHICLENO," & _
				 "		   WB.DRIVERNAME," & _
				 "		   WB.SPBNO," & _
				 "		   ''  Grade," & _
				 "		   ''  rcpo," & _
				 "		   NVL(GR.BUNCHES, 0) JJG," & _
				 "		   GR.F00 unripbunch," & _
				 "		   GR.F5  rottenbunch," & _
				 "		   GR.Tkosong Emptybunch," & _
				 "		   GR.Tpanjang Tallbunch," & _
				 "		   NVL(GR.Air, 0) AIR," & _
				 "		   NVL(GR.Sampah, 0) Sampah," & _
				 "		   NVL(GR.Pasir, 0) Pasir," & _
				 "		   nvl(GR.EATENBYRAT, 0) + nvl(GR.Brondolan, 0) + nvl(GR.Bkecil, 0) Lain," & _
				 "		   NVL(GR.dura, 0) Dura," & _
				 "  	    ( CASE WB.FIELD_ESTATE_ESTATECODE WHEN 'LUAR' THEN " & _
				 "				  'A01'  " & _
				 "			   ELSE " & _
				 "				  'A02'" & _
				 "			   END ) GRtype, " & _
				 "			 ( CASE WB.FIELD_ESTATE_ESTATECODE WHEN 'LUAR' THEN " & _
				 "				  WB.Negono " & _
				 "			   ELSE " & _
				 "				  '' " & _
				 "			   END ) " & _
				 "				   NegoNo," & _
				 "		   '' STO," & _
				 "		   '' STOLINE," & _
				 "		   0  price," & _
				 "		   '' Curr," & _
				 "		   0 Bridgestate," & _
				 "		   GR.Remarks" & _	
                 " FROM WEIGHBRIDGETICKET WB" & _
                 " LEFT JOIN GRADINGRECORD GR ON (WB.FCBA = GR.FCBA AND WB.CHITNO = GR.WBT_CHITNO)" & _
                 " LEFT JOIN CUSTOMER CU ON (WB.FCBA = CU.FCBA AND WB.CUST_CUSTOMERCODE = CU.CUSTOMERCODE)" & _
                 " WHERE WB.PRODUCT_PRODUCTCODE = 'FFB' AND WB.isinbridge <> '1' "  & whereCons & " " & _
                 " ORDER BY WB.CHITNO"	    
    
	''response.write q
	
   Set rsVH = Server.CreateObject("ADODB.Recordset")
   rsVH.CursorType = 3
   rsVH.CursorLocation = 3
   rsVH.Open q, connstr%> 
   <form action="FFBUploadToExtSys.Asp" method="post" name="eForm">
   <table id="entry" class="entry" cellpadding="1" cellspacing="1" width="620">
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;">
	<!-- <td align="center" valign="top" >Upload</td> -->
	<td align="center" valign="top" width="60px" class="coba">
			<input 	type="checkbox" id="cb_SelectedAll" name="cb_SelectedAll" onchange="SelectedAll_Checkbox('eForm')" /> <br>UPLOAD ALL
	</td>
	<td id="para1" class="para1" align="center" valign="top" >ID_KEY</td>					
	<td align="center" valign="top" >NO</td>
	<td align="center" valign="top" >CHITNO</td>
	<td align="center" valign="top" >LOCATION</td>
	<td align="center" valign="top" >POSTINGDATE</td>
	<td align="center" valign="top" >TIMEOUT</td>
	<td align="center" valign="top" >CUSTOMERNAME</td>
	<td align="center" valign="top" >CUSTOMERCODE</td>
	<td align="center" valign="top" >BATCH</td>
	<td align="center" valign="top" >BRUTO</td>
	<td align="center" valign="top" >TARRA</td>
	<td align="center" valign="top" >SORTASE</td>
	<td align="center" valign="top" >AFTER SORTASE</td>
	<td align="center" valign="top" >VEHICLENO</td>
	<td align="center" valign="top" >DRIVERNAME</td>
	<td align="center" valign="top" >SPBNO</td>
	<td align="center" valign="top" >GRADE</td>
	<td align="center" valign="top" >RCPO</td>
	<td align="center" valign="top" >JJG</td>
	<td align="center" valign="top" >UNRIPBUNCH</td>
	<td align="center" valign="top" >ROTTENBUNCH</td>
	<td align="center" valign="top" >EMPTYBUNCH</td>
	<td align="center" valign="top" >TALLBUNCH</td>
	<td align="center" valign="top" >AIR</td>
	<td align="center" valign="top" >SAMPAH</td>
	<td align="center" valign="top" >PASIR,</td>
	<td align="center" valign="top" >LAIN</td>
	<td align="center" valign="top" >DURA</td>
	<td align="center" valign="top" >GRTYPE</td>
	<td align="center" valign="top" >NEGONO</td>
	<td align="center" valign="top" >STO</td>
	<td align="center" valign="top" >STOLINE</td>
	<td align="center" valign="top" >PRICE</td>
	<td align="center" valign="top" >CURR</td>
	<td align="center" valign="top" >BRIDGESTATE</td>
	<td align="center" valign="top" >REMARKS</td>



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
		  
          v_UnixCode  = rsVH("CHITNO") & "_" & rsVH("LOCATION")

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
            <td align="Right" valign="top"><%=rsVH("CHITNO")%></td>
			<td align="Right" valign="top"><%=rsVH("LOCATION")%></td>
			<td align="Right" valign="top"><%=rsVH("POSTINGDATE")%></td>
			<td align="Right" valign="top"><%=rsVH("TIMEOUT")%></td>
			<td align="Right" valign="top"><%=rsVH("CUSTOMERNAME")%></td>
			<td align="Right" valign="top"><%=rsVH("CUSTOMERCODE")%></td>
			<td align="Right" valign="top"><%=rsVH("BATCH")%></td>
			<td align="Right" valign="top"><%=rsVH("BRUTOQTY")%></td>
			<td align="Right" valign="top"><%=rsVH("TARRAQTY")%></td>
			<td align="Right" valign="top"><%=rsVH("SORTASE")%></td>
			<td align="Right" valign="top"><%=rsVH("AFTSRTQTY")%></td>
			<td align="Right" valign="top"><%=rsVH("VEHICLENO")%></td>
			<td align="Right" valign="top"><%=rsVH("DRIVERNAME")%></td>
			<td align="Right" valign="top"><%=rsVH("SPBNO")%></td>
			<td align="Right" valign="top"><%=rsVH("GRADE")%></td>
			<td align="Right" valign="top"><%=rsVH("RCPO")%></td>
			<td align="Right" valign="top"><%=rsVH("JJG")%></td>
			<td align="Right" valign="top"><%=rsVH("UNRIPBUNCH")%></td>
			<td align="Right" valign="top"><%=rsVH("ROTTENBUNCH")%></td>
			<td align="Right" valign="top"><%=rsVH("EMPTYBUNCH")%></td>
			<td align="Right" valign="top"><%=rsVH("TALLBUNCH")%></td>
			<td align="Right" valign="top"><%=rsVH("AIR")%></td>
			<td align="Right" valign="top"><%=rsVH("SAMPAH")%></td>
			<td align="Right" valign="top"><%=rsVH("PASIR")%></td>
			<td align="Right" valign="top"><%=rsVH("LAIN")%></td>
			<td align="Right" valign="top"><%=rsVH("DURA")%></td>
			<td align="Right" valign="top"><%=rsVH("GRTYPE")%></td>
			<td align="Right" valign="top"><%=rsVH("NEGONO")%></td>
			<td align="Right" valign="top"><%=rsVH("STO")%></td>
			<td align="Right" valign="top"><%=rsVH("STOLINE")%></td>
			<td align="Right" valign="top"><%=rsVH("PRICE")%></td>
			<td align="Right" valign="top"><%=rsVH("CURR")%></td>
			<td align="Right" valign="top"><%=rsVH("BRIDGESTATE")%></td>
			<td align="Right" valign="top"><%=rsVH("REMARKS")%></td>
 			  
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
		
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input type="hidden" name="txtfcba" value="<%=v_txtfcba%>">
	    <input type="hidden" name="txtDateFrom" value="<%=v_txtDateFrom%>">
		<input type="hidden" name="txtDateTo" value="<%=v_txtDateTo%>">
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
