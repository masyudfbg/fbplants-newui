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
<title>ePlaS - Update Status Purchase Requisition From Ext System</title>
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
		


v_FCCODE =request("v_FCCODE")

 if Request("txtJenis")="New" or Request("txtJenis")="Edit" then
	
		   qCompanyAll = " select FCBA,PRCODE,LINEID,ITEMCODE,stock_api.get_stockName(ITEMCODE,FCBA) ItemDesc ,APPROVEDQTY , ROWSTATEID " & _
						 " from Purchaserequestdetails t where  rowstateid in ('ReleaseOnProgress','CancelOnProgress','CloseOnProgress') and fcba = '" & Session("fcBA") & "' " & _
						 " order by  t.fcba,t.prcode,LINEID "
		
			''response.Write qCompanyAll 
		
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
			    v_UnixCode  = rsModuleCheck("PRCODE") & "_" & rsModuleCheck("LINEID")  & "_" & rsModuleCheck("FCBA") 
				if request(trim("data_" & v_UnixCode))="on" Then
					'response.Write v_UnixCode & " "
					msgresult = msgresult & ", " & v_UnixCode
					' response.Write "BEGIN BridgingSAP_PRPOGRN_API.UpdateStatusPRFcba('" & rsModuleCheck("PRCODE") & "','" & rsModuleCheck("FCBA") & "'); END; "
					 strQueryStatus = "BEGIN BridgingSAP_PRPOGRN_API.UpdateStatusPRFcba('" & rsModuleCheck("PRCODE") & "','" & rsModuleCheck("FCBA") & "'); END; "
	                 connStr.execute(strQueryStatus)
					
				end if
				rsModuleCheck.MoveNext
		    wend	
		end if	
		
		rsModuleCheck.Close
		set rsModuleCheck = Nothing
		If msgresult <> ""  then
			response.Write "<font face=Verdana size=1 color=blue>Update Status PR No "& msgresult &" From External System Completed..</font>"
		End If
		cmdView = "True"
								 
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
 
 <form name="frmEntry" method="post" action="UpdPurchaseReqFromExtSys.Asp">
	<table class="data" width="500">
		
		<tr><td colspan="2"  align="left">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
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
  
	  ' q = " select PRCODE||'_'||LINEID||'_'||FCBA  as ID_KEY, FCBA,PRCODE,LINEID,ITEMCODE,stock_api.get_stockName(ITEMCODE,FCBA) ItemDesc ,APPROVEDQTY , ROWSTATEID " & _
			' " from Purchaserequestdetails t WHERE ROWNUM < 10 " 
			
             q = " select PRCODE||'_'||LINEID||'_'||FCBA  as ID_KEY, FCBA,PRCODE,LINEID,ITEMCODE,stock_api.get_stockName(ITEMCODE,FCBA) ItemDesc ,APPROVEDQTY , ROWSTATEID " & _
				 " from Purchaserequestdetails t where  rowstateid in ('ReleaseOnProgress','CancelOnProgress','CloseOnProgress') and fcba = '" & Session("fcBA") & "' " & _
                 " order by  t.fcba,t.prcode,LINEID " 
                
		'	response.write q		
    
   Set rsVH = Server.CreateObject("ADODB.Recordset")
   rsVH.CursorType = 3
   rsVH.CursorLocation = 3
   rsVH.Open q, connstr%>
   <form action="UpdPurchaseReqFromExtSys.Asp" method="post" name="eForm">
   <table id="entry" class="entry" cellpadding="1" cellspacing="1" width="620">
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;">
		<!-- <td align="center" valign="top" >UPDATE</td> -->
		<td align="center" valign="top" width="60px" class="coba">
			<input 	type="checkbox" id="cb_SelectedAll" name="cb_SelectedAll" onchange="SelectedAll_Checkbox('eForm')" /> <br>UPDATE ALL
		</td>
		<td id="para1" class="para1" align="center" valign="top" >ID_KEY</td>						
		<td align="center" valign="top" >NO</td>
		<td align="center" valign="top" >FCBA</td>
		<td align="center" valign="top" >PR NO</td>
		<td align="center" valign="top" >LINEID</td>
		<td align="center" valign="top" >ITEM CODE</td>
		<td align="center" valign="top" >DESCRIPTION</td>
		<td align="center" valign="top" >APPROVEDQTY</td>
		<td align="center" valign="top" >ROWSTATEID</td>
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
		  
          v_UnixCode  = rsVH("PRCODE") & "_" & rsVH("LINEID") & "_" & rsVH("FCBA")  
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
			<td align="Right" valign="top"><%=rsVH("PRCODE")%></td>
			<td align="Right" valign="top"><%=rsVH("LINEID")%></td>
			<td align="Right" valign="top"><%=rsVH("ITEMCODE")%></td>
			<td align="Right" valign="top" width="500"><%=rsVH("ITEMDESC")%></td>
			<td align="Right" valign="top"><%=rsVH("APPROVEDQTY")%></td>
			<td align="Right" valign="top"><%=rsVH("ROWSTATEID")%></td>
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
	  <input type="hidden" name="v_FCCODE" value="<%=v_FCCODE%>">
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
