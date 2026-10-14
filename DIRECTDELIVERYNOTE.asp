<%Option Explicit%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<html>
<head>
<title>ePlaS - Delivery Note</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js" ></script>
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
<!--<script language="JavaScript" src="/js/date-picker.js"></script> -->
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>
</head>
<body>
<!--#include file="top.jsp"-->
	<div style=" width:21.7%; float:left;margin-left:0.3%; margin-top:0.3%; margin-bottom:0.3%;">
			<!--#include file="menu.asp"-->
		</div>
   <div class="judul_head" valign="top" bgcolor="#fff" style=" width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;"> 
 <div class="judul">
	   Module : <%=request("txtModule")%> / <%=request("txtTable")%>
	</div>
 <div class="box_both" style="overflow:scroll;">
 <%
	'-------------awal-----------------------


Dim FCENTRY, FCEDIT, FCIP, FCBA
dim txt_NOTECODE,Perioddate

Dim DNOTE_DATE, DELIVERYNOTECODE, CONTRACTCODE, GATEPASS, WBT_CHITNO, STORAGECODE, QUANTITY, FFA, DIRT, MOISTURE, TRANSPORTCO, DRIVERNAME, ICNO

DIM SEALNO_A,SEALNO_B,SEALNO_C,SEALNO_D,SEALNO_E,SEALNO_F,SEALNO_G,SEALNO_H,SEALNO_I,SEALNO_J,POSTINGDATE
Dim SEALNO
Dim cmdView,rsDN
Dim recordDataCheck, rsrecordDataCheck, varSelisih
Dim varError

txt_NOTECODE=request("txt_NOTECODE")
Perioddate=request("Perioddate")
txttable=request("txttable")
DNOTE_DATE=request("DNOTE_DATE")
DELIVERYNOTECODE=request("DELIVERYNOTECODE")
CONTRACTCODE=request("CONTRACTCODE")
GATEPASS=request("GATEPASS")
WBT_CHITNO=request("WBT_CHITNO")
STORAGECODE=request("STORAGECODE")
QUANTITY=request("QUANTITY")
FFA=request("FFA")
DIRT=request("DIRT")
MOISTURE=request("MOISTURE")
TRANSPORTCO=request("TRANSPORTCO")
DRIVERNAME=request("DRIVERNAME")
POSTINGDATE =request("POSTINGDATE")
ICNO=request("ICNO")
	SEALNO_A=request("SEALNO_A")
	SEALNO_B=request("SEALNO_B")
	SEALNO_C=request("SEALNO_C")
	SEALNO_D=request("SEALNO_D")
	SEALNO_E=request("SEALNO_E")
	SEALNO_F=request("SEALNO_F")
	SEALNO_G=request("SEALNO_G")
	SEALNO_H=request("SEALNO_H")
	SEALNO_I=request("SEALNO_I")
	SEALNO_J=request("SEALNO_J")

 if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then

	if Request("txtJenis")= "Delete" then
       q="Delete from DELIVERYNOTE where DELIVERYNOTECODE='" & DELIVERYNOTECODE & "' And" & _
			" DNOTE_DATE=to_Date('" & DNOTE_DATE & "','MM/DD/YYYY') And  fcBA='" & Session("fcBA") & "' and " & _
			" isdirectform = 'TRUE' "
   	   connStr.execute(q)
       cmdView	= "True"
	elseif Request("txtJenis")="New" then
	  recordDataCheck = "SELECT dataContract.contractCode, dataContract.totalWeight, nvl(dataDeliver.totalDelivered,0) totalDelivered, "  & _
						"	  nvl((dataContract.totalWeight - nvl(dataDeliver.totalDelivered,0)),0) selisih " & _
						"  FROM (SELECT contractcode, totalweight " & _
						"		  FROM DELIVERYCONTRACT " & _
						"		 WHERE fcba =  '" & Session("fcBA") & "') dataContract, " & _
						"	   (SELECT a.contractcode, " & _
                        "     a.deliverynotecode,  "& _
                        "     a.totalDeliver,  " & _
                        "  nvl(totalReturn,0) totalReturn,  " & _
                        "      (nvl(a.totalDeliver,0) - nvl(b.totalReturn,0)) totalDelivered " & _
						"		  FROM (SELECT contractcode, deliverynotecode, quantity totalDeliver " & _
						"				  FROM deliverynote " & _
						"				 WHERE fcba =  '" & Session("fcBA") & "' and isdirectform = 'TRUE') a, " & _
						"			   (  SELECT dnote_deliverynotecode, SUM (quantity) totalReturn " & _
						"					FROM deliveryreturnnote " & _
						"				   WHERE fcba = '" & Session("fcBA") & "'  " & _
						"				GROUP BY dnote_deliverynotecode) b " & _ 
						"		 WHERE a.deliverynotecode = b.dnote_deliverynotecode (+)) dataDeliver " & _
						" WHERE dataContract.contractCode = dataDeliver.contractcode (+) and " & _
						"       dataContract.contractCode = '" & CONTRACTCODE & "'"
						
		Set rsrecordDataCheck = Server.CreateObject("ADODB.Recordset")
		rsrecordDataCheck.CursorType = 3
		rsrecordDataCheck.CursorLocation = 3 

		rsrecordDataCheck.Open recordDataCheck, connstr


		if Not rsrecordDataCheck.Eof then
		   varSelisih = cdbl(rsrecordDataCheck("selisih"))
		Else
		   varSelisih = 0
		End if	

        if (varSelisih  - cdbl(QUANTITY) ) >=  0 then 
	
			q="Insert into DELIVERYNOTE (DNOTE_DATE,POSTINGDATE, DELIVERYNOTECODE, CONTRACTCODE, GATEPASS, WBT_CHITNO, STORAGECODE, QUANTITY, FFA, DIRT, MOISTURE, TRANSPORTCO, DRIVERNAME, ICNO, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME, ISDIRECTFORM ) " & _
			"values (to_Date('" & DNOTE_DATE & "','MM/DD/YYYY')," & "to_Date('" & POSTINGDATE & "','MM/DD/YYYY'),'" &  DELIVERYNOTECODE & "','" &  CONTRACTCODE & "','" & GATEPASS & "','" &  WBT_CHITNO & "'," & _
			"'" &  STORAGECODE & "','" &  QUANTITY & "','" &  FFA & "','" &  DIRT & "','" &  MOISTURE & "','" & TRANSPORTCO & "','" &  DRIVERNAME & "','" &  ICNO & "'," & _
			"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
			Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI'), 'TRUE' )"
			
			'response.Write(q)
			'response.Flush()
			connStr.execute(q)
			cmdView	= "True"
			  IF REQUEST(SEALNO_A)<> "" THEN
				  q="Insert into DELIVERYNOTESEAL (DELIVERYNOTECODE, SEALNO,FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
					"values ('" &  DELIVERYNOTECODE & "','" &  SEALNO_A & "'," & _
					"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
					Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI') )"
					connStr.execute(q)
					cmdView	= "True"
			  ELSEIF REQUEST(SEALNO_B)<> "" THEN
				  q="Insert into DELIVERYNOTESEAL (DELIVERYNOTECODE, SEALNO,FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
					"values ('" &  DELIVERYNOTECODE & "','" &  SEALNO_B & "'," & _
					"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
					Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI') )"
					connStr.execute(q)
					cmdView	= "True"
			  ELSEIF REQUEST(SEALNO_C)<> "" THEN
				  q="Insert into DELIVERYNOTESEAL (DELIVERYNOTECODE, SEALNO,FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
					"values ('" &  DELIVERYNOTECODE & "','" &  SEALNO_C & "'," & _
					"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
					Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI') )"
					connStr.execute(q)
					cmdView	= "True"
			  ELSEIF REQUEST(SEALNO_D)<> "" THEN
				  q="Insert into DELIVERYNOTESEAL (DELIVERYNOTECODE, SEALNO,FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
					"values ('" &  DELIVERYNOTECODE & "','" &  SEALNO_D & "'," & _
					"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
					Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI') )"
					connStr.execute(q)
					cmdView	= "True"
			  ELSEIF REQUEST(SEALNO_E)<> "" THEN
				  q="Insert into DELIVERYNOTESEAL (DELIVERYNOTECODE, SEALNO,FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
					"values ('" &  DELIVERYNOTECODE & "','" &  SEALNO_E & "'," & _
					"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
					Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI') )"
					connStr.execute(q)
					cmdView	= "True"
			  ELSEIF REQUEST(SEALNO_F)<> "" THEN
				  q="Insert into DELIVERYNOTESEAL (DELIVERYNOTECODE, SEALNO,FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
					"values ('" &  DELIVERYNOTECODE & "','" &  SEALNO_F & "'," & _
					"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
					Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI') )"
					connStr.execute(q)
					cmdView	= "True"
			  ELSEIF REQUEST(SEALNO_G)<> "" THEN
				  q="Insert into DELIVERYNOTESEAL (DELIVERYNOTECODE, SEALNO,FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
					"values ('" &  DELIVERYNOTECODE & "','" &  SEALNO_G & "'," & _
					"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
					Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI') )"
					connStr.execute(q)
					cmdView	= "True"
			  ELSEIF REQUEST(SEALNO_H)<> "" THEN
				  q="Insert into DELIVERYNOTESEAL (DELIVERYNOTECODE, SEALNO,FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
					"values ('" &  DELIVERYNOTECODE & "','" &  SEALNO_H & "'," & _
					"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
					Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI') )"
					connStr.execute(q)
					cmdView	= "True"
			  ELSEIF REQUEST(SEALNO_I)<> "" THEN
				  q="Insert into DELIVERYNOTESEAL (DELIVERYNOTECODE, SEALNO,FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
					"values ('" &  DELIVERYNOTECODE & "','" &  SEALNO_I & "'," & _
					"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
					Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI') )"
					connStr.execute(q)
					cmdView	= "True"
			  ELSEIF REQUEST(SEALNO_J)<> "" THEN
				  q="Insert into DELIVERYNOTESEAL (DELIVERYNOTECODE, SEALNO,FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
					"values ('" &  DELIVERYNOTECODE & "','" &  SEALNO_J & "'," & _
					"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
					Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI') )"
					'response.Write(q)
					'response.Flush()
					connStr.execute(q)
					cmdView	= "True"
			  END IF
			varError = ""
	    else
			varError = "Total Quantity yang anda masukan melebihi dari kontrak yang disepakati "  
        end if
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
	
			q="Update DELIVERYNOTE set " & _ 
			" TRANSPORTCO ='" & TRANSPORTCO & "'" & _
			" ,DRIVERNAME='" & DRIVERNAME & "'" & _
			" ,ICNO ='" & ICNO & "'" & _
			" ,FFA ='" & FFA & "'" & _
			" ,DIRT ='" & DIRT & "'" & _
			" ,MOISTURE ='" & MOISTURE & "'" & _
			" ,STORAGECODE ='" & STORAGECODE & "'" & _
			" ,GATEPASS='" & GATEPASS & "'" & _
			" ,QUANTITY='" & QUANTITY & "'" & _
			" , LASTUPDATE = sysdate " & _
		    " , LASTTIME = To_Char(sysdate,'HH24:MI') " & _
			" where DELIVERYNOTECODE='" & DELIVERYNOTECODE & "' " & _
			" And DNOTE_DATE=to_Date('" & DNOTE_DATE & "','MM/DD/YYYY')" & _ 
			" And WBT_CHITNO='" & WBT_CHITNO & "'" & _
			" And CONTRACTCODE ='" & CONTRACTCODE & "'" & _ 
			" And  fcBA='" & Session("fcBA") & "'" & _
			" And ISDIRECTFORM = 'TRUE' "
			'response.write q
			connStr.execute(q)
		
			recordDataCheck = "SELECT dataContract.contractCode, dataContract.totalWeight, nvl(dataDeliver.totalDelivered,0) totalDelivered, "  & _
						"	  nvl((dataContract.totalWeight - nvl(dataDeliver.totalDelivered,0)),0) selisih " & _
						"  FROM (SELECT contractcode, totalweight " & _
						"		  FROM DELIVERYCONTRACT " & _
						"		 WHERE fcba =  '" & Session("fcBA") & "') dataContract, " & _
						"	   (SELECT a.contractcode, " & _
                        "     a.deliverynotecode,  "& _
                        "     a.totalDeliver,  " & _
                        "  nvl(totalReturn,0) totalReturn,  " & _
                        "      (nvl(a.totalDeliver,0) - nvl(b.totalReturn,0)) totalDelivered " & _
						"		  FROM (SELECT contractcode, deliverynotecode, quantity totalDeliver " & _
						"				  FROM deliverynote " & _
						"				 WHERE fcba =  '" & Session("fcBA") & "' and isdirectform = 'TRUE' ) a, " & _
						"			   (  SELECT dnote_deliverynotecode, SUM (quantity) totalReturn " & _
						"					FROM deliveryreturnnote " & _
						"				   WHERE fcba = '" & Session("fcBA") & "' " & _
						"				GROUP BY dnote_deliverynotecode) b " & _ 
						"		 WHERE a.deliverynotecode = b.dnote_deliverynotecode (+)) dataDeliver " & _
						" WHERE dataContract.contractCode = dataDeliver.contractcode (+) and " & _
						"       dataContract.contractCode = '" & CONTRACTCODE & "'"
						
			Set rsrecordDataCheck = Server.CreateObject("ADODB.Recordset")
			rsrecordDataCheck.CursorType = 3
			rsrecordDataCheck.CursorLocation = 3 

			rsrecordDataCheck.Open recordDataCheck, connstr


			if Not rsrecordDataCheck.Eof then
			   varSelisih = rsrecordDataCheck("selisih")
			Else
			   varSelisih = 0
			End if	

            if cdbl(varSelisih ) <  0 then 
			    connStr.execute("rollback")
				varError = "Total Quantity yang anda masukan melebihi dari kontrak yang disepakati "  
			else
		        connStr.execute("commit")			
				varError = ""  
			end if

        cmdView	= "True"
    else
        cmdView	= "True"
	end if								 
	'Response.Write "sql " & q
else
    cmdView	= request("cmdView")
End if

Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis
if Request("Perioddate")<>"" then
   Perioddate   =Request("Perioddate")
else
   Perioddate   =textDate(date)
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
 <form name="frmEntry" method="post" action="DIRECTDELIVERYNOTE.asp">
    <font color=red><%= varError  %></font>
	<table class="data" width="100%" cellpadding="0" cellspacing="0">
		<TR> 
		<td colspan="2" ><div align="left"><font size="2" face="Verdana"><strong>MAINTENANCE DIRECT - DELIVERY NOTE</strong></font></div></td>
		</TR>

		<tr><td align="left" width="20%">Date</td>
		<td align="left"><input class="datepicker" type="text" maxlength="10" name="Perioddate" size="10" value="<%=Perioddate%>">
		</td></tr>
		<TR>
		<td align="left">Delivery Note Code</td><td align="left"><input type="text" size="15" maxlength="30" name="txt_NOTECODE" value="<%=txt_NOTECODE%>"><font color="#FF0000">*</font></td>
		</TR>	
		<tr><td colspan="2"  align="left">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
		<a href="PrintDeliveryNote.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&Perioddate=<%=Perioddate%>&txt_NOTECODE=<%=txt_NOTECODE%>" 
		title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a>
		</td>
		</tr>
    </table>  
 </form>
<%if cmdView = "True" Then
   Dim rsGAD,intRecordsPerPage,cColor
   set rsGAD = server.CreateObject("Adodb.recordset")
   'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
   q="select DNOTE_DATE, DELIVERYNOTECODE, CONTRACTCODE, GATEPASS, WBT_CHITNO, STORAGECODE, QUANTITY, FFA, " & _
     "DIRT, MOISTURE, TRANSPORTCO, DRIVERNAME,ICNO,POSTINGDATE,  FCENTRY, FCEDIT, FCIP, FCBA from DELIVERYNOTE where " & _
     " fcBA='" & session("fcBA") & "' and isdirectform = 'TRUE' "
	if  len(trim(request("Perioddate"))) > 0 then
		q = q & " and to_char(DNOTE_DATE,'mm/yyyy')=to_char(to_date('" & request("Perioddate") & "','" & "MM/DD/YYYY" & "'),'mm/yyyy')"
	end if
	if  len(trim(request("txt_NOTECODE"))) > 0 then
		q = q & " and DELIVERYNOTECODE='" & request("txt_NOTECODE") & "'"
	end if
	if Request("txtJenis")="Edit"  then
	   q = q & " and to_char(DNOTE_DATE,'mm/yyyy')=to_char(to_date('" & DNOTE_DATE & "','" & "MM/DD/YYYY" & "'),'mm/yyyy')"
	end if
	q = q & " Order by DNOTE_DATE, WBT_CHITNO "
   'response.Write(q)
  'response.Flush()
   rsGAD.CursorType = 3
   rsGAD.CursorLocation = 3
   rsGAD.Open q, connstr %>

   <form action="DIRECTDELIVERYNOTE.asp" method="post" name="eForm">
   <table class="data" width="100%" cellpadding="0" cellspacing="0">
		<tr>
			<td align="left">Date</td>
				<td align="left"><input class="datepicker" type="text" maxlength="10" name="DNOTE_DATE" size="10" value="<%=DNOTE_DATE%>">
				<font color="#FF0000">*</font>
				</td>
			<td align="left">Seal No.01</td><td align="left"><input type="text" size="15" maxlength="15" name="SEALNO_A" value="<%=SEALNO_A%>"></td>		
		</tr>
		<tr>
		  <td align="left">Posting Date</td>
		  <td align="left"><input class="datepicker" type="text" maxlength="10" name="POSTINGDATE" size="10" value="<%=POSTINGDATE%>">
				<font color="#FF0000">*</font>
				</td>
           <td align="left">Seal No.02</td><td align="left"><input type="text" size="15" maxlength="15" name="SEALNO_B" value="<%=SEALNO_B%>"></td>
		</tr>
		<TR>
		<td align="left">Delivery Note Code</td><td align="left"><input type="text" size="15" maxlength="30" name="DELIVERYNOTECODE" value="<%=DELIVERYNOTECODE%>"><font color="#FF0000">*</font></td>
		<td align="left">Seal No.03</td><td align="left"><input type="text" size="15" maxlength="15" name="SEALNO_C" value="<%=SEALNO_C%>"></td>		
		</TR>	
		<TR>
		<%IF Request("txtJenis")="Edit" then%>
			<td align="left">Contract Code</td>
			<td align="left"> <select name="CONTRACTCODE"> 			
			<%cmdCallComboSetup "DELIVERYCONTRACT",request("CONTRACTCODE"),"fcba",session("fcba")%>
			</select> <font color="#FF0000">*</font></td>
		<%Else%>
			<td align="left">Contract Code</td>
			<td align="left"><select name="CONTRACTCODE">         
			<%cmdCallComboSetup "DELIVERYCONTRACT","CONTRACTCODE","",""%>
			</select> <font color="#FF0000">*</font></td>
		
		<%End IF%>
		<td align="left">Seal No.04</td><td align="left"><input type="text" size="15" maxlength="15" name="SEALNO_D" value="<%=SEALNO_D%>"></td>					
		</tr>
		<tr>
		<td align="left">Gate Pass</td><td align="left"><input type="text" size="30" maxlength="50" name="GATEPASS" value="<%=GATEPASS%>"></td>
		<td align="left">Seal No.05</td><td align="left"><input type="text" size="15" maxlength="15" name="SEALNO_E" value="<%=SEALNO_E%>"></td>					
		</tr>
		<TR>
		<td align="left">Sales No</td><td align="left"><input type="text" size="15" maxlength="15" name="WBT_CHITNO" value="<%=WBT_CHITNO%>"><font color="#FF0000">*</font></td>
		<td align="left">Seal No.06</td><td align="left"><input type="text" size="15" maxlength="15" name="SEALNO_F" value="<%=SEALNO_F%>"></td>							
		</tr>
		<tr>
		<%IF Request("txtJenis")="Edit" then%>		
            <td align="left"><font face="Verdana" size="1">From Storage Code</font></td>
            <td align="left"> <select name="STORAGECODE">
			    <%cmdCallComboSetup "PRODUCTSTORAGE",request("STORAGECODE"),"fcba",session("fcba")%>

              </select> <font color="#FF0000">*</font> </td>
		<%else%>
				<td align="left"><font face="Verdana" size="1">From Storage Code</font></td>
				<td align="left"> <select name="STORAGECODE">
				<%cmdCallComboSetup "PRODUCTSTORAGE","STORAGECODE","fcba",session("fcba")%>

				<!--%cmdCallComboWhereStatement "PRODUCTSTORAGE","STORAGECODE","fcba",session("fcba")%-->
				</select> <font color="#FF0000">*</font> </td>
		<%end if%>
		<td align="left">Seal No.07</td><td align="left"><input type="text" size="15" maxlength="15" name="SEALNO_H" value="<%=SEALNO_H%>"></td>					
		</tr>
		<tr>
		<td align="left">Quantity (KG)</td><td align="left"><input type="text" size="30" maxlength="50" name="QUANTITY" value="<%=QUANTITY%>"></td>	
		<td align="left">Seal No.08</td><td align="left"><input type="text" size="15" maxlength="15" name="SEALNO_G" value="<%=SEALNO_G%>"></td>					
		</tr>
		<tr>
		<td align="left">FFA</td><td align="left"><input type="text" size="15" name="FFA" value="<%=FFA%>"></td>
		<td align="left">Seal No.09</td><td align="left"><input type="text" size="15" maxlength="15" name="SEALNO_I" value="<%=SEALNO_I%>"></td>						
		</tr>
		<tr>
		<td align="left">Dirt</td><td align="left"><input type="text" size="15" name="DIRT" value="<%=DIRT%>"></td>
		<td align="left">Seal No.10</td><td align="left"><input type="text" size="15" maxlength="15" name="SEALNO_J" value="<%=SEALNO_J%>"></td>					
		</tr>
		<tr>
		<td align="left">Moisture</td><td align="left"><input type="text" size="15" name="MOISTURE" value="<%=MOISTURE%>"></td>
		<td align="left"></td><td align="left"></td>					
		</tr>
		<tr>
		<td align="left">Transport CO</td><td align="left"><input type="text" size="30" maxlength="50" name="TRANSPORTCO" value="<%=TRANSPORTCO%>"><font color="#FF0000">*</font></td>
		</tr>
		<tr>
		<td align="left">Driver Name</td><td align="left"><input type="text" size="30"  maxlength="50" name="DRIVERNAME" value="<%=DRIVERNAME%>"><font color="#FF0000">*</font></td>
		</tr>
		<tr>
		<td align="left">IC No</td><td align="left"><input type="text" size="20" maxlength="20" name="ICNO" value="<%=ICNO%>"></td>
		</tr>
	    <TR><td colspan="2" align="left">
		  <input type="hidden" name="Perioddate" value="<%=Perioddate%>">
		  <input type="hidden" name="txt_NOTECODE" value="<%=txt_NOTECODE%>">

		  <input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		  <input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		  <%IF Request("txtJenis")="Edit" then%>
		       <input type="hidden" name="txtJenis" value="Edit">
		       <input Type="Submit"  name="Add" value="Update">
		  <%Else%>
		       <input type="hidden" name="txtJenis" value="New">
		       <input Type="Submit"  name="Add" value="Add">
		  <%End if%>
		  <input Type="Reset"  name="Reset" value="Reset">
		  <input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
	  </tr>
   </table>
   </form>
  <!--#include file="validation_show.Asp"--> 
<%
   JumRec=1
   intRecordsPerPage = 20
	
	Dim intNumberOfPages,intRecords,nPage,howmanyfields,JumRec
	howmanyfields=rsGAD.fields.count-1
	
	JumRec=1
	intRecordsPerPage = 20
	
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

   rsGAD.PageSize = intRecordsPerPage
   intNumberOfPages = 0
   intNumberOfPages =  rsGAD.PageCount
   intRecords = rsGAD.RecordCount   
   %>
   <table width="737" class="data" >
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">DATE</TD>	
	<TD align="left">DELIVERY NOTE CODE</TD>	
	<TD align="left">CONTRACT CODE</TD>	
	<TD align="left">GATE PASS</TD>	
	<TD align="left">SALES NO</TD>	
	<TD align="left">FROM STORAGE CODE</TD>	
	<TD align="left">QUATITY</TD>	
	<TD align="left">FFA</TD>	
	<TD align="left">DIRT</TD>	
	<TD align="left">MOISTURE</TD>	
	<TD align="left">TRASPORT CO</TD>	
	<TD align="left">DRIVER</TD>	
	<TD align="left">ICNO</TD>	
	<TD align="left">POSTING DATE</TD>	
	</tr>
   
        <%
   if Not rsGAD.eof then
      nCountofField=rsGAD.fields.count-5
	  no=1
	  While Not rsGAD.EOF
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
		  <td width="79" align="left" valign="top"><%=NO%></td>
			<td valign="middle" width="21">

			<a href="DIRECTDELIVERYNOTE.asp?DELIVERYNOTECODE=<%=rsGAD("DELIVERYNOTECODE")%>&txttable=<%=request("txttable")%>&DNOTE_DATE=<%=rsGAD("DNOTE_DATE")%>&POSTINGDATE=<%=rsGAD("POSTINGDATE")%>&TRANSPORTCO=<%=rsGAD("TRANSPORTCO")%>&DRIVERNAME=<%=rsGAD("DRIVERNAME")%>&ICNO=<%=rsGAD("ICNO")%>&WBT_CHITNO=<%=rsGAD("WBT_CHITNO")%>&FFA=<%=rsGAD("FFA")%>&DIRT=<%=rsGAD("DIRT")%>&MOISTURE=<%=rsGAD("MOISTURE")%>&STORAGECODE=<%=rsGAD("STORAGECODE")%>&GATEPASS=<%=rsGAD("GATEPASS")%>&QUANTITY=<%=rsGAD("QUANTITY")%>&CONTRACTCODE=<%=rsGAD("CONTRACTCODE")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsGAD(1)%>" border="0"></a>
			</td>
			<td valign="middle" width="20">
			<a href="DIRECTDELIVERYNOTE.asp?DELIVERYNOTECODE=<%=rsGAD("DELIVERYNOTECODE")%>&txttable=<%=request("txttable")%>&DNOTE_DATE=<%=rsGAD("DNOTE_DATE")%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGAD(1)%>" border="0" onClick="return confirm('Are you sure to Delete');"></a>
			</td>
		    <%for i = 0 to nCountofField%>
			    <td width="597" align="left" valign="top" title="<%=rsGAD(i).Name%>"><%=rsGAD(i)%></td>
		    <%next%>
		  </tr>
		  <%
		  rsGAD.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   end if 
   rsGAD.Close
   set rsGAD = Nothing%>
   </div>
      </TABLE>
<%End if%>
  <%'-------------akhir----------------------%>
</div></div>
  <!--#include file="bottom.jsp"-->

</body>
</html>
