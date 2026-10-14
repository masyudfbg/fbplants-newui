<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Weighbridge Ticket</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js" ></script>
<link rel="StyleSheet" href="dtree.css" type="text/css" />
<link rel="StyleSheet" href="style-content.css" type="text/css" />
<link rel="stylesheet" href="css/jquery-ui.css" />
<!-- HADIR/ font ICON -->
<link rel="stylesheet" href="css/font/font-awesome.css" />
<!-- END:HADIR -->
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
dim txt_PRODUCTCODE,Perioddate

Dim TRANSACTIONTYPE, CHITNO, PRODUCT_PRODUCTCODE, VEHICLENO, CUST_CUSTOMERCODE, FIELD_ESTATE_ESTATECODE,DIVISION, TIMEIN, TIMEOUT,DRIVERNAME
Dim FIELD_FIELDCODE, DATEIN, WBRIDGE_WEIGHBRIDGECODEIN, WEIGHTIN, DATEOUT, WBRIDGE_WEIGHBRIDGECODEOUT, WEIGHTOUT,POSTINGDATE,SPBNO,STORAGECODE
Dim cmdView
Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""
txttable=request("txttable")

	txt_PRODUCTCODE = request("txt_PRODUCTCODE")
	Perioddate = request("Perioddate")
	TRANSACTIONTYPE = request("TRANSACTIONTYPE")
	CHITNO = request("CHITNO")
	PRODUCT_PRODUCTCODE = request("PRODUCT_PRODUCTCODE")
	VEHICLENO = request("VEHICLENO")
	CUST_CUSTOMERCODE = request("CUST_CUSTOMERCODE")
	FIELD_ESTATE_ESTATECODE = request("FIELD_ESTATE_ESTATECODE")
	FIELD_FIELDCODE = request("FIELD_FIELDCODE")
	DATEIN = request("DATEIN")
	WBRIDGE_WEIGHBRIDGECODEIN = request("WBRIDGE_WEIGHBRIDGECODEIN")
	WEIGHTIN = request("WEIGHTIN")
	DATEOUT = request("DATEOUT")
	WBRIDGE_WEIGHBRIDGECODEOUT = request("WBRIDGE_WEIGHBRIDGECODEOUT")
	WEIGHTOUT = request("WEIGHTOUT")
	'Add by Akbar
	DIVISION = request("DIVISION")
	TIMEIN = request("TIMEIN")
	TIMEOUT = request("TIMEOUT")
	DRIVERNAME = request("DRIVERNAME")
	POSTINGDATE = request("POSTINGDATE")
	SPBNO = request("SPBNO")
	STORAGECODE = request("STORAGECODE")
	TBSTYPE = request("TBSTYPE")
	
    ' Add by Akbar
	
 if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then

	if Request("txtJenis")= "Delete" then
	
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), POSTINGDATE)
		If Cek_IsDailyClosing = "FALSE" then
			q = " delete from weighbridgeticket" & _
			    " where fcba = '" & Session("fcBA") & "'" & _
				"       and chitno = '" & CHITNO & "'"
			   connStr.execute(q)
			   'response.Write(q)
			   'response.Flush()
		else
			err.number = 1001
			err.description = "Tanggal """ & POSTINGDATE & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		cmdView	= "True"
	   
	elseif Request("txtJenis")="New" then
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), POSTINGDATE)
		If Cek_IsDailyClosing = "FALSE" then
			  q=" Insert into WEIGHBRIDGETICKET (TRANSACTIONTYPE, CHITNO, TBSTYPE, PRODUCT_PRODUCTCODE, VEHICLENO, CUST_CUSTOMERCODE,DIVISION,SPBNO, FIELD_ESTATE_ESTATECODE, FIELD_FIELDCODE, DATEIN,TIMEIN, WBRIDGE_WEIGHBRIDGECODEIN, WEIGHTIN, DATEOUT,TIMEOUT, WBRIDGE_WEIGHBRIDGECODEOUT, WEIGHTOUT,DRIVERNAME,STORAGECODE,POSTINGDATE, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME,DATATYPE ) " & _
				" values ('" &  TRANSACTIONTYPE & "','" &  CHITNO & "','" &  TBSTYPE & "','" &  PRODUCT_PRODUCTCODE & "','" &  VEHICLENO & "','" &  CUST_CUSTOMERCODE & "','" &  DIVISION & "','" &  SPBNO & "','" &  FIELD_ESTATE_ESTATECODE & "'," & _
				"'" &  FIELD_FIELDCODE & "',to_Date('" & DATEIN & "','MM/DD/YYYY') ,'" &  TIMEIN & "','" &  WBRIDGE_WEIGHBRIDGECODEIN & "'," &  WEIGHTIN & ",to_Date('" & DATEOUT & "','MM/DD/YYYY') ,'" &  TIMEOUT & "'," & _
				"'" &  WBRIDGE_WEIGHBRIDGECODEOUT & "'," &  WEIGHTOUT & ",'" & DRIVERNAME & "','" &  STORAGECODE & "', to_Date('" & POSTINGDATE & "','MM/DD/YYYY') ," & _
				"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
				Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI'),'AM' )"
				'response.Write(q)
				'response.Flush()
				'potsting date di ubah menjadi to_date dan postingdate dihilangkan di update
				connStr.execute(q)
		else
			err.number = 1001
			err.description = "Tanggal """ & POSTINGDATE & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		cmdView	= "True"
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), POSTINGDATE)
		If Cek_IsDailyClosing = "FALSE" then
				q = " update weighbridgeticket" & _
                    " set TRANSACTIONTYPE ='" & TRANSACTIONTYPE & "'" & _
			        "     , PRODUCT_PRODUCTCODE ='" & PRODUCT_PRODUCTCODE & "'" & _
			        "     , VEHICLENO ='" & VEHICLENO & "'" & _
					"     , CUST_CUSTOMERCODE ='" & CUST_CUSTOMERCODE & "'" & _
					"     , SPBNO ='" & SPBNO & "'" & _
					"     , DIVISION ='" & DIVISION & "'" & _
					"     , TBSTYPE ='" & TBSTYPE & "'" & _
			        "     , FIELD_ESTATE_ESTATECODE ='" & FIELD_ESTATE_ESTATECODE & "'" & _
			        "     , FIELD_FIELDCODE ='" & FIELD_FIELDCODE & "'" & _
					"     , DATEIN=to_Date('" & DATEIN & "','MM/DD/YYYY')" & _
					"     , TIMEIN='" &  TIMEIN & "'" & _
			        "     , WBRIDGE_WEIGHBRIDGECODEIN ='" & WBRIDGE_WEIGHBRIDGECODEIN & "'" & _
			        "     , WEIGHTIN ='" & WEIGHTIN & "'" & _			        
			        "     , DATEOUT=to_Date('" & DATEOUT & "','MM/DD/YYYY')" & _					
			        "     , TIMEOUT='" &  TIMEOUT & "'" & _
			        "     , WBRIDGE_WEIGHBRIDGECODEOUT ='" & WBRIDGE_WEIGHBRIDGECODEOUT & "'" & _
			        "     , WEIGHTOUT='" & WEIGHTOUT & "'" & _
			        "     , DRIVERNAME='" & DRIVERNAME & "'"  & _ 
					"     , POSTINGDATE=to_Date('" & POSTINGDATE & "','MM/DD/YYYY')" & _
					"     , STORAGECODE= '" & STORAGECODE & "'" & _
					"     , fcedit = '" & Session("CodeName") & "'" & _
			        "     , LASTUPDATE = sysdate " & _
			        "     , LASTTIME = To_Char(sysdate,'HH24:MI') " & _
			        " where fcba = '" & Session("fcBA") & "'" & _
			        "       and chitno = '" & CHITNO & "'"
				' response.Write(q)
				'response.Flush()
				connStr.execute(q)
				connStr.execute("commit")
				'RESPONSE.Write "SQLnsddsdya : " & Q
		else
			err.number = 1001
			err.description = "Tanggal """ & POSTINGDATE & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
        cmdView	= "True"
    else
        cmdView	= "True"
	end if								 
	'Response.Write "sql " & 
else
    cmdView	= request("cmdView")
End if

Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis
if len(trim(Perioddate)) <> 0 then
   Perioddate = Request("Perioddate")
else
   Perioddate = textDate(date)
end if

txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")
'fcCode=request("fcCode")
txtJenis = request("txtJenis")
If Err.Number <> 0 then
   response.Write  "<div class = ""error"">" & err.description & "</div>"
End if

 '-------------ISI Payment Form
 %>

 <form name="frmEntry" method="post" action="WEIGHBRIDGETICKET.asp">
	<table class="data" width="70%" cellpadding="0" cellspacing="0">
		<TR> 
		<td colspan="2"><div align="left"><font size="2" face="Verdana"><strong>MAINTENANCE WEIGHBRIDGETICKET</strong></font></div></td>
		</TR>

		<tr><td align="left">Date</td>
		<td align="left"><input class="datepicker" type="text" maxlength="10" name="Perioddate" size="10" value="<%=Perioddate%>">
		
		</td></tr>
		<TR>
			<td align="left">Product Code</td>
			<td align="left"> <select name="txt_PRODUCTCODE">
			<%cmdCallComboSetup "PRODUCT",REQUEST("txt_PRODUCTCODE"),"FCBA",Session("fcBA")%>
			</select></td>
		</TR>	
		<tr><td colspan="2"  align="left">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
		<a href="PrintWEIGHBRIDGETICKETtest.asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&Perioddate=<%=Perioddate%>&txt_PRODUCTCODE=<%=REQUEST("txt_PRODUCTCODE")%>" 
		title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a>
		<!-- HADIR -->
		<%
		dim dataModule,dataSubmodule,dataDetail
		dataModule 		= "Weighbridge"
		dataSubmodule 	= "Weighbridge Ticket"
		dataDetail 		= "NULL"
		%>
		<a class="jqplot-image-button" style="" href="ajax/hadirAjax.asp?dataModule=<%=dataModule%>&dataSubmodule=<%=dataSubmodule%>&dataDetail=<%=dataDetail%>"><i class="fa fa-check-square-o"></i>&nbsp;Hadir</a>
		<!-- END:HADIR -->
		</td>
		</tr>
    </table>  
 </form>

<%if cmdView = "True" Then
   Dim rsGAD,intRecordsPerPage,cColor
   set rsGAD = server.CreateObject("Adodb.recordset")
   'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
   q="select TRANSACTIONTYPE, CHITNO, PRODUCT_PRODUCTCODE, VEHICLENO, CUST_CUSTOMERCODE,SPBNO,DIVISION, FIELD_ESTATE_ESTATECODE, FIELD_FIELDCODE, DATEIN, TIMEIN," & _ 
     "WBRIDGE_WEIGHBRIDGECODEIN, WEIGHTIN, DATEOUT,TIMEOUT, WBRIDGE_WEIGHBRIDGECODEOUT, WEIGHTOUT,DRIVERNAME,STORAGECODE,POSTINGDATE,FCENTRY, FCEDIT, FCIP, FCBA,datatype from WEIGHBRIDGETICKET where " & _
     " fcBA='" & session("fcBA") & "'"
	if  len(trim(Perioddate)) > 0 then
		q = q & " and postingdate = to_date('" & Perioddate & "','MM/DD/YYYY')"
	end if
	if  len(trim(txt_PRODUCTCODE)) > 0 then
		q = q & " and PRODUCT_PRODUCTCODE='" & txt_PRODUCTCODE & "'"
	end if
	if Request("txtJenis")="Edit"  then
	   q = q & " and PRODUCT_PRODUCTCODE='" & PRODUCT_PRODUCTCODE & "' and postingdate = to_Date('" & Perioddate & "','MM/DD/YYYY') "
	end if
	
	q = q & " Order by Chitno"
   'response.Write(q)
  'response.Flush()
   rsGAD.CursorType = 3
   rsGAD.CursorLocation = 3
   rsGAD.Open q, connstr %>

   <form action="WEIGHBRIDGETICKET.asp" method="post" name="eForm">
   <table class="data" cellpadding="0" cellspacing="0" width="70%">
		<tr>
		<%IF txtJenis="Edit" Then%>
		<td align="left">Transaction Type</td>
			<% 	Dim rsType
			q = "select P.FCCODE,P.FCNAME from  PARAMETER P  where P.PARCODE= 'W0001' and P.FCBA = '" & Session("fcBA") & "' order by P.FCNAME desc"
			Set rsType = Server.CreateObject("ADODB.Recordset")
			rsType.CursorType = 3
			rsType.CursorLocation = 3
			rsType.Open q, connstr
			if not rsType.eof then%>
     	    <td align="left"> <select name="TRANSACTIONTYPE" title="TRANSACTIONTYPE Character a s/d z or 1 s/d 9">         
			<option value=""></option> 
			<%			Do While Not rsType.eof%>
			<%if rsType("FCCODE")=request("TRANSACTIONTYPE") Then%>			
				<option value="<%=rsType("FCCODE")%>" selected><%=rsType("FCNAME")%></option> 
			<%Else%>
				<option value="<%=rsType("FCCODE")%>"><%=rsType("FCNAME")%></option> 
			<%End if%>
			<%rsType.movenext
			loop
			End if
			rsType.Close
			set rsType = Nothing
			q=""
			%>
			</select></td>
		<%Else%>
		<td align="left">Transaction Type</td>
			<% 	Dim rsTypeA
			q = "select P.FCCODE,P.FCNAME from  PARAMETER P  where P.PARCODE= 'W0001' and P.FCBA = '" & Session("fcBA") & "' order by P.FCNAME desc"
			Set rsTypeA = Server.CreateObject("ADODB.Recordset")
			rsTypeA.CursorType = 3
			rsTypeA.CursorLocation = 3
			rsTypeA.Open q, connstr
			if not rsTypeA.eof then%>
     	    <td align="left"> <select name="TRANSACTIONTYPE" title="TRANSACTIONTYPE Character a s/d z or 1 s/d 9">         
			<option value=""></option> 
			<%			Do While Not rsTypeA.eof%>
			<%if rsTypeA("FCCODE")=request("TRANSACTIONTYPE") Then%>			
				<option value="<%=rsTypeA("FCCODE")%>" selected><%=rsTypeA("FCNAME")%></option> 
			<%Else%>
				<option value="<%=rsTypeA("FCCODE")%>"><%=rsTypeA("FCNAME")%></option> 
			<%End if%>
			<%rsTypeA.movenext
			loop
			End if
			rsTypeA.Close
			set rsTypeA = Nothing
			q=""
			%>
			</select></td>
		<%End if%>
		</tr>
		<tr>
			<td align="left">Chit No</td>
			<td align="left">
				<%IF txtJenis="Edit" Then%>
				<input type="text" size="15" maxlength="15" name="CHITNO" value="<%=CHITNO%>" readonly="True" style="background:#999999">
				<%else%>
				<input type="text" size="15" maxlength="15" name="CHITNO" value="<%=CHITNO%>">
				<%end if%>
				<font color="#FF0000">*</font>
			</td>
		</tr>
		<tr>
		<%IF txtJenis="Edit" Then%>
			<td align="left">Product Code</td>
			<td align="left"> <select name="PRODUCT_PRODUCTCODE">
			<%cmdCallComboSetup "PRODUCT",request("PRODUCT_PRODUCTCODE"),"FCBA",Session("fcBA")%>
			</select> <font color="#FF0000">*</font> </td>
		<%Else%>
			<td align="left">Product Code</td>
			<td align="left"> <select name="PRODUCT_PRODUCTCODE">
			<%cmdCallComboSetup "PRODUCT",request("PRODUCT_PRODUCTCODE"),"FCBA",Session("fcBA")%>
			</select> <font color="#FF0000">*</font> </td>
		<%End if%>		
		</tr>
			<tr>
            <td align="left"><font face="Verdana" size="1">Vehicle</font></td><td align="left"><input type="text" size="15" maxlength="50" name="VEHICLENO" value="<%=VEHICLENO%>"><font color="#FF0000">*</font></td>
			</tr>
		<tr>
		<%IF txtJenis="Edit" Then%>
			<td align="left">Customer Code</td>
			<td align="left"> <select name="CUST_CUSTOMERCODE">
			<%cmdCallComboSetup "CUSTOMER",request("CUST_CUSTOMERCODE"),"FCBA",Session("fcBA")%>
			</select> <font color="#FF0000">*</font></td>
		<%Else%>
			<td align="left">Customer Code</td>
			<td align="left"> <select name="CUST_CUSTOMERCODE">
			<%cmdCallComboSetup "CUSTOMER",request("CUST_CUSTOMERCODE"),"FCBA",Session("fcBA")%>
			</select> <font color="#FF0000">*</font></td>
		<%End if%>
		<tr>
		<tr>
		<td align="left">SPBNO </td><td align="left"><input type="text" size="15" name="SPBNO" value="<%=SPBNO%>"></td>
		</tr>
		<tr>
		<td align="left">Division ( Afdeling ) </td><td align="left"><input type="text" size="15" name="DIVISION" value="<%=DIVISION%>"></td>
		</tr>
		<tr>
		<td align="left">Field Estate Code</td><td align="left"><input type="text" size="15" name="FIELD_ESTATE_ESTATECODE" value="<%=FIELD_ESTATE_ESTATECODE%>"></td>
		</tr>
		<tr>
		<td align="left"><font face="Verdana" size="1">Field Code</font></td><td align="left"><input type="text" size="15" maxlength="30" name="FIELD_FIELDCODE" value="<%=FIELD_FIELDCODE%>"></td>
		</tr>
		<tr>
				<td align="left">TBS Type</td>	
				<td align="left">
					<select name="tbstype" value="<%=tbstype%>">
						<option value="<%=tbstype%>"></option>
						<option value="TBS" <% If cstr(""&tbstype) = "TBS" Then %> Selected <%End if%>>TBS</option> 
						<option value="TBSB" <% If cstr(""&tbstype) = "TBSB" Then %> Selected <%End if%>>TBSB</option> 
					</select>
				</td>	
			</tr>
		<tr>
		
		<td align="left">Date In</td>
			<td align="left"><input class="datepicker" type="text" maxlength="10" name="DATEIN" size="10" value="<%=DATEIN%>">
			<font color="#FF0000">*</font>
			</td>
		</tr>
		
		<tr>
		<td align="left">Time In</td><td align="left"><input type="text" size="15" name="TIMEIN" value="<%=TIMEIN%>"></td>
		</tr>
		
		<tr>
		<td align="left">Weighbridge Code In</td>
		<%IF txtJenis="Edit" Then%>
			<td align="left"> <select name="WBRIDGE_WEIGHBRIDGECODEIN">
			<%cmdCallComboSetup "WEIGHBRIDGE",REQUEST("WBRIDGE_WEIGHBRIDGECODEIN"),"FCBA",Session("fcBA")%>
			</select> <font color="#FF0000">*</font></td>
		<%Else%>
			<td align="left"> <select name="WBRIDGE_WEIGHBRIDGECODEIN">
			<%cmdCallComboSetup "WEIGHBRIDGE",REQUEST("WBRIDGE_WEIGHBRIDGECODEIN"),"FCBA",Session("fcBA")%>
			</select> <font color="#FF0000">*</font></td>		
		<%End if%>
		</tr>
		<tr>
		<td align="left">Weight In (Kg)</td><td align="left"><input type="text" size="15" maxlength="15" name="WEIGHTIN" value="<%=WEIGHTIN%>"><font color="#FF0000">*</font></td>
		</tr>
		<tr>
		<td align="left">Date Out</td>
			<td align="left"><input class="datepicker" type="text" maxlength="10" name="DATEOUT" size="10" value="<%=DATEOUT%>">
			<font color="#FF0000">*</font>
			</td>
		</tr>
		
		<tr>
		<td align="left">Time Out</td><td align="left"><input type="text" size="15" name="TIMEOUT" value="<%=TIMEOUT%>"></td>
		</tr>
		
		<tr>
		<td align="left">Weighbridge Code Out</td>
		<%IF txtJenis="Edit" Then%>
			<td align="left"> <select name="WBRIDGE_WEIGHBRIDGECODEOUT">
			<%cmdCallComboSetup "WEIGHBRIDGE",REQUEST("WBRIDGE_WEIGHBRIDGECODEOUT"),"FCBA",Session("fcBA")%>
			</select> <font color="#FF0000">*</font></td>
		<%Else%>
			<td align="left"> <select name="WBRIDGE_WEIGHBRIDGECODEOUT">
			<%cmdCallComboSetup "WEIGHBRIDGE",REQUEST("WBRIDGE_WEIGHBRIDGECODEOUT"),"FCBA",Session("fcBA")%>
			</select> <font color="#FF0000">*</font></td>		
		<%End if%>

		</tr>
		
		<tr>
		<td align="left">Weight Out (Kg)</td><td align="left"><input type="text" size="15" maxlength="15" name="WEIGHTOUT" value="<%=WEIGHTOUT%>"><font color="#FF0000">*</font></td>
		</tr>
		<tr>
		<td align="left">Driver Name</td><td align="left"><input type="text" size="25" name="DRIVERNAME" value="<%=DRIVERNAME%>"></td>
		</tr>
		<tr>
		<%IF Request("txtJenis")="Edit" then%>		
            <td align="left"><font face="Verdana" size="1">Storage Code</font></td>
            <td align="left"> <select name="STORAGECODE">
                <%cmdCallComboSetup "PRODUCTSTORAGE",STORAGECODE,"fcba",session("fcba")%>
              </select> <font color="#FF0000">*</font> </td>
		<%else%>
				<td align="left"><font face="Verdana" size="1">Storage Code</font></td>
				<td align="left"> <select name="STORAGECODE">
				<%cmdCallComboSetup "PRODUCTSTORAGE",STORAGECODE,"fcba",session("fcba")%>
				</select> <font color="#FF0000">*</font> </td>
		<%end if%>
		</tr>
		<tr>
		<td align="left">Posting Date</td>
			<td align="left"><input class="datepicker" type="text" maxlength="10" name="POSTINGDATE" size="10" value="<%=POSTINGDATE%>">
			<font color="#FF0000">*</font>
			</td>
		</tr>
		<tr>
	    <TR><td colspan="2" align="left">
		  <input type="hidden" name="Perioddate" value="<%=Perioddate%>">
		  <input type="hidden" name="txt_PRODUCTCODE" value="<%=txt_PRODUCTCODE%>">
		
		  <input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		  <input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		  <%IF Request("txtJenis")="Edit" then%>
		       <input type="hidden" name="txtJenis" value="Edit">
		       <input Type="Submit"  name="Add" value="Update">
			   <input Type="Hidden" Value="<%=transactiontype%>" Name="transactiontype">
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
        <tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;"> 
          <TD align="left">NO</TD>	
		  <TD colspan="2">&nbsp;</TD>
          <TD align="left">TRANS TYPE</TD>
          <TD align="left">CHITNO</TD>
          <TD align="left">PRODUCT CODE</TD>
          <TD align="left">VEHICLE</TD>
          <TD align="left">CUSTOMER</TD>
		  <TD align="left">SPBNO</TD>
		  <TD align="left">DIVISION</TD>
          <TD align="left">FIELD ESTATE</TD>
          <TD align="left">FIELD CODE</TD>
          <TD align="left">DATE IN</TD>
		  <TD align="left">TIME IN</TD>
          <TD align="left">WEIGHBRIDGE CODE IN</TD>
          <TD align="left">WEIGHT IN</TD>
          <TD align="left">DATE OUT</TD>
		  <TD align="left">TIME OUT</TD>
          <TD align="left">WEIGHBRIDGE CODE OUT</TD>
          <TD align="left">WEIGHT OUT</TD>
		  <TD align="left">DRIVER NAME</TD>
		  <TD align="left">STORAGE CODE</TD>	
		  <TD align="left">POSTING DATE</TD>
        </tr>
        <%
   if Not rsGAD.eof then
      nCountofField=rsGAD.fields.count-6
	  no=1
	  While Not rsGAD.EOF
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
        <tr bgcolor="<%=cColor%>"> 
          <td width="79" align="left" valign="top"><%=NO%></td>
		<% if rsGAD("DATATYPE") = "AT"  then %>
			<td colspan="2" align="center">Automate</td>  
		<% ELSE %>	
			<td valign="middle" width="21">
			  <a href="WEIGHBRIDGETICKET.asp?PRODUCT_PRODUCTCODE=<%=rsGAD("PRODUCT_PRODUCTCODE")%>&txttable=<%=request("txttable")%>&SPBNO=<%=rsGAD("SPBNO")%>&DIVISION=<%=rsGAD("DIVISION")%>&POSTINGDATE=<%=rsGAD("POSTINGDATE")%>&DRIVERNAME=<%=rsGAD("DRIVERNAME")%>&CHITNO=<%=rsGAD("CHITNO")%>&VEHICLENO=<%=rsGAD("VEHICLENO")%>&WEIGHTIN=<%=rsGAD("WEIGHTIN")%>&WEIGHTOUT=<%=rsGAD("WEIGHTOUT")%>&CUST_CUSTOMERCODE=<%=rsGAD("CUST_CUSTOMERCODE")%>&DATEIN=<%=rsGAD("DATEIN")%>&TIMEIN=<%=rsGAD("TIMEIN")%>&DATEOUT=<%=rsGAD("DATEOUT")%>&TIMEOUT=<%=rsGAD("TIMEOUT")%>&WBRIDGE_WEIGHBRIDGECODEIN=<%=rsGAD("WBRIDGE_WEIGHBRIDGECODEIN")%>&TRANSACTIONTYPE=<%=rsGAD("TRANSACTIONTYPE")%>&FIELD_ESTATE_ESTATECODE=<%=rsGAD("FIELD_ESTATE_ESTATECODE")%>&FIELD_FIELDCODE=<%=rsGAD("FIELD_FIELDCODE")%>&WBRIDGE_WEIGHBRIDGECODEOUT=<%=rsGAD("WBRIDGE_WEIGHBRIDGECODEOUT")%>&Perioddate=<%=rsGAD("POSTINGDATE")%>&txt_PRODUCTCODE=<%=rsGAD("PRODUCT_PRODUCTCODE")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsGAD(1)%>" border="0"></a> 
			  </td>
			  <td valign="middle" width="20"> <a href="WEIGHBRIDGETICKET.asp?PRODUCT_PRODUCTCODE=<%=rsGAD("PRODUCT_PRODUCTCODE")%>&txttable=<%=request("txttable")%>&CHITNO=<%=rsGAD("CHITNO")%>&CUST_CUSTOMERCODE=<%=rsGAD("CUST_CUSTOMERCODE")%>&SPBNO=<%=rsGAD("SPBNO")%>&DATEIN=<%=rsGAD("DATEIN")%>&DATEOUT=<%=rsGAD("DATEOUT")%>&POSTINGDATE=<%=rsGAD("POSTINGDATE")%>&Perioddate=<%=rsGAD("POSTINGDATE")%>&txt_PRODUCTCODE=<%=rsGAD("PRODUCT_PRODUCTCODE")%>&txtJenis=Delete" onClick="return confirm('Are you sure to Delete');"><img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGAD(1)%>" border="0"></a> 
			  </td>
		<% end if %>	
          <%for i = 0 to nCountofField%>
          <td width="597" align="left" valign="top" title="<%=rsGAD(i).Name%>"><%=rsGAD(i)%></td>
          <%next%>
        </tr>
        <%
		  rsGAD.MoveNext
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
