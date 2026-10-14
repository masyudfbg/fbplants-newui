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
<title>ePlaS - Update Vendor Master Per Company From Ext System</title>
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
dim Dataterupdate
dim txtVendor,v_txtVendor
dim whereCons
Dim v_txtErrorFlag, txtErrorFlag		



 whereCons =""
 v_txtVendor	  = request("txtVendor")
 v_txtErrorFlag   = request("txtErrorFlag")
 
 If v_txtVendor <> "--ALL--" then
    whereCons = whereCons & " And  A1.VENDORCODE = '" & v_txtVendor & "' "
 end if

 if v_txtErrorFlag <> "--ALL--" then
   whereCons = whereCons & " And  B1.ERRORFLAG = '" & v_txtErrorFlag & "' "
 end if
 
if Request("txtJenis")="Process"  Then
	
			qCompanyAll =" Select DISTINCT A1.* , B1.ERRORFLAG, B1.ERRORDESC From ( " & _
						 " Select CONCAT(CONCAT(a.VENDORCODE,'_'),b.COMPANY) as ID_KEY, " & _
						 " 			b.company,b.RECACCOUNT,b.suppliertype,b.supplierffbtype, a.* " & _
						 " From vendormastergeneral@IPLAS192L a " & _
						 " inner join vendormastercompany@IPLAS192L b on a.vendorcode = b.vendorcode " & _
						 " where a.vendorcode in ( select c.vendorcode " & _
						 " 						   from vendormastercompanyTrans@IPLAS192L c " & _
						 "						   where c.bridgestate=0 " & _
						 " 						   and c.appsys='IPLAS') " & _
						 "			) A1 " & _
						 " LEFT JOIN vendormastercompanyTrans@IPLAS192L B1 ON B1.vendorcode = A1.vendorcode " & _
						 " WHERE 1=1 " & whereCons & " " & _						 
						 " order by A1.VENDORCODE "
				 
	
		   ' qCompanyAll = " Select b.company,b.RECACCOUNT,b.suppliertype,b.supplierffbtype, a.* " & _
						 ' "  From vendormastergeneral@IPLAS192L a " & _
						 ' " inner join vendormastercompany@IPLAS192L b " & _
						 ' " 	on a.vendorcode = b.vendorcode " & _
						 ' " where a.vendorcode in ( select c.vendorcode from vendormastercompanyTrans@IPLAS192L c " & _
						 ' "	where c.bridgestate=0 and c.appsys='IPLAS'  and ( ERRORFLAG = 0  or  ERRORFLAG is null  ) and b.company = c.company "  & whereCons & " )" & _
						 ' " order by a.VENDORCODE "
						 
		 ''response.write qCompanyAll	  
		 
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
			    v_UnixCode  = rsModuleCheck("VENDORCODE")  & "_" & rsModuleCheck("COMPANY")
				if request(trim("data_" & v_UnixCode))="on" Then
					'response.Write v_UnixCode
					msgresult =  msgresult & rsModuleCheck("VENDORCODE")  & "_" & rsModuleCheck("COMPANY") & ", "
					' Dataterupdate = Dataterupdate & ", " & v_UnixCode
					strQueryStatus = "BEGIN BridgingSAP_MASTER_API.UpadteVendorCompMaster('" & rsModuleCheck("VENDORCODE") & "','" & rsModuleCheck("COMPANY") & "'); END; "
	                connStr.execute(strQueryStatus)	
				end if
				rsModuleCheck.MoveNext
		    wend	
		end if	
		
		rsModuleCheck.Close
		set rsModuleCheck = Nothing
		if msgresult <> ""  then
		   response.Write "<font face=Verdana size=1 color=blue>Download Vendor Per Company to All FCBA, Vendor Code ," & msgresult & " From External System Completed..</font>"
		end if
		cmdView = "True"
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
 
 <form name="frmEntry" method="post" action="DownloadVendCompFromExtSys.Asp">
	<table class="data" width="500">
		
			<h2>UPDATE VENDOR CODE PER COMPANY FROM EXTERNAL SYSTEM</h2>
				
	        &nbsp;&nbsp;
		<tr>
		<td><font face="verdana" size="1">VENDOR CODE</font></td>
		<td>
			<select name="txtVendor" >
			<option value="--ALL--">--ALL--</option>
			   <%
				Dim rsStock, strStock
				strStock =   " Select distinct a.VENDORCODE, " & _
							 "		  a.VENDORNAME  " & _
							 " From vendormastergeneral@IPLAS192L a " & _
							 " inner join vendormastercompany@IPLAS192L b " & _
							 " 	on a.vendorcode = b.vendorcode " & _
							 " where a.vendorcode in ( select c.vendorcode from vendormastercompanyTrans@IPLAS192L c " & _
							 "	                       where c.bridgestate=0 and c.appsys='IPLAS' and b.company = c.company )" & _
							 " order by a.VENDORCODE "
								
				Set rsStock = Server.CreateObject("ADODB.Recordset")
				rsStock.Open strStock, connstr, adOpenStatic
							
				if not rsStock.eof then
				   Do While Not rsStock.eof
					  if v_txtVendor = rsStock("VENDORCODE") then %>
							<option value="<%=rsStock("VENDORCODE")%>" Selected><%=rsStock("VENDORCODE")%> - <%=rsStock("VENDORNAME")%></option>
					<%else%>
					    	<option value="<%=rsStock("VENDORCODE")%>"><%=rsStock("VENDORCODE")%> - <%=rsStock("VENDORNAME")%></option>
	    			<%end if			
					   rsStock.movenext
	   			   loop
				end if
				rsStock.Close
				set rsStock = Nothing
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
						   " From vendormastercompanyTrans@IPLAS192L b " & _
						   " where b.bridgestate = 0 " & _
						   " And b.appsys = 'IPLAS' "
								
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
		<input Type="Submit"  name="Add" value="View Data">
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
  
			 q = " Select DISTINCT A1.* , B1.ERRORFLAG, B1.ERRORDESC From ( " & _
                 " Select CONCAT(CONCAT(a.VENDORCODE,'_'),b.COMPANY) as ID_KEY, " & _
				 " 			b.company,b.RECACCOUNT,b.suppliertype,b.supplierffbtype, a.* " & _
				 " From vendormastergeneral@IPLAS192L a " & _
				 " inner join vendormastercompany@IPLAS192L b on a.vendorcode = b.vendorcode " & _
				 " where a.vendorcode in ( select c.vendorcode " & _
				 " 						   from vendormastercompanyTrans@IPLAS192L c " & _
				 "						   where c.bridgestate=0 " & _
				 " 						   and c.appsys='IPLAS') " & _
				 "			) A1 " & _
				 " LEFT JOIN vendormastercompanyTrans@IPLAS192L B1 ON B1.vendorcode = A1.vendorcode " & _
				 " WHERE 1=1 " & whereCons & " " & _				 
				 " order by A1.VENDORCODE "
				
   ''response.write q
   
   Set rsVH = Server.CreateObject("ADODB.Recordset")
   rsVH.CursorType = 3
   rsVH.CursorLocation = 3
   rsVH.Open q, connstr%>
   <form action="DownloadVendCompFromExtSys.Asp" method="post" name="eForm">
   <table id="entry" class="entry" cellpadding="1" cellspacing="1" width="620">
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;">
		<!--<td align="center" valign="top" >Download</td>-->

		<td align="center" valign="top" width="60px" class="coba">
			<input 	type="checkbox" id="cb_SelectedAll" name="cb_SelectedAll" onchange="SelectedAll_Checkbox('eForm')" /> DOWNLOAD ALL
		</td>		
		<td id="para1" class="para1" align="center" valign="top" >ID_KEY</td>						
		<td align="center" valign="top" >NO</td>
		<td align="center" valign="top" >SAP COMPANY CODE</td>
		<td align="center" valign="top" >VENDOR CODE</td>
		<td align="center" valign="top" >VENDOR NAME</td>
		<td align="center" valign="top" >STREET</td>
		<td align="center" valign="top" >CITY</td>
		<td align="center" valign="top" >PROVINCE</td>
		<td align="center" valign="top" >POSTALCODE</td>
		<td align="center" valign="top" >TELEPHONE</td>
		<td align="center" valign="top" >FAX</td>
		<td align="center" valign="top" >EMAIL</td>
		<td align="center" valign="top" >VATREGNO</td>
		<td align="center" valign="top" >ISBLOCK</td>
		<td align="center" valign="top" >ISDELETE</td>
		<td align="center" valign="top" >USERINPUT</td>
		<td align="center" valign="top" >DATEINPUT</td>
		<td align="center" valign="top" >USERMODIFY</td>
		<td align="center" valign="top" >DATEMODIFY</td>
		<td align="center" valign="top" >ERROR FLAG</td>
		<td align="center" valign="top" >ERROR DESC</td>
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
		  
          v_UnixCode  = rsVH("VENDORCODE") & "_" & rsVH("COMPANY")

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
			<td align="Right" valign="top"><%=rsVH("COMPANY")%></td>
            <td align="Right" valign="top"><%=rsVH("VENDORCODE")%></td>
			<td align="Right" valign="top"><%=rsVH("VENDORNAME")%></td>
			<td align="Right" valign="top"><%=rsVH("STREET")%></td>
			<td align="Right" valign="top"><%=rsVH("CITY")%></td>
			<td align="Right" valign="top"><%=rsVH("PROVINCE")%></td>
			<td align="Right" valign="top"><%=rsVH("POSTALCODE")%></td>
			<td align="Right" valign="top"><%=rsVH("TELEPHONE")%></td>
			<td align="Right" valign="top"><%=rsVH("FAX")%></td>
			<td align="Right" valign="top"><%=rsVH("EMAIL")%></td>
			<td align="Right" valign="top"><%=rsVH("VATREGNO")%></td>
			<td align="Right" valign="top"><%=rsVH("ISBLOCK")%></td>
			<td align="Right" valign="top"><%=rsVH("ISDELETE")%></td>
			<td align="Right" valign="top"><%=rsVH("USERINPUT")%></td>
			<td align="Right" valign="top"><%=rsVH("DATEINPUT")%></td>
			<td align="Right" valign="top"><%=rsVH("USERMODIFY")%></td>
			<td align="Right" valign="top"><%=rsVH("DATEMODIFY")%></td>
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
	  <input type="hidden" name="txtVendor" value="<%=v_txtVendor%>">
	  <input type="hidden" name="txtErrorFlag" value="<%=v_txtErrorFlag%>">
	  
	  <input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
	  <input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
	  <input type="hidden" name="txtJenis" value="Process">
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

