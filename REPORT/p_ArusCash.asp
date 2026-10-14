<!--#include file="../connDb1.Inc"-->
<!--#include file="../txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>ePlaS Reporting - Arus Cash</title>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<script language="JavaScript" src='/js/x.js'></script>
<script language="JavaScript" src='/js/xtableheaderfixed.js'></script>
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

</head>
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

Dim ACCOUNTINGYEARPERIOD


DATECREATED_From= trim(request("DATECREATED_From"))
DATECREATED_To=trim(request("DATECREATED_To"))

BANKCODE=trim(request("BANKCODE"))
txtControlJob=""
nOpening=0

' Added By Sendy 09/25/2018
			v_txtfcba 			= Request("txtfcba")
	 		
	 		strDefaultBa =  " SELECT" & _
							"     FCCODE as FCBA" & _
							" FROM BUSINESSUNIT BU" & _
							" INNER JOIN USER_FCBA_TAB UFT ON (BU.FCCODE = UFT.FCBA) " & _
							" WHERE" & _
							"     ISDEFAULT = 'TRUE'" & _
							"     AND FCTYPE = 'E'" & _
							"     AND FCUSERCODE = '" & Session("CodeName") & "'"

			Set rsDefaultBa = Server.CreateObject("ADODB.Recordset")
			rsDefaultBa.Open strDefaultBa, connstr, adOpenStatic

			If Not rsDefaultBa.EOF Then
				varDefaultBa = rsDefaultBa("FCBA")
			Else
				varDefaultBa = ""
			End If
			rsDefaultBa.Close
			set rsDefaultBa = Nothing

			If (v_txtFcba = "" Or Len(Trim(v_txtFcba)) = 0) Then
				v_txtFcba = varDefaultBa
			End If	
	 		if  v_txtfcba <> "--ALL--"  and v_txtfcba <> "" then
		       whereCons = whereCons & " And pay.fcba = '" & v_txtfcba & "'"
			end if
		    ' selectCons = selectCons & ",fcba"
		    ' orderCons  = orderCons & ",BUSINESSUNIT_API.get_OrderData(fcba)"
			q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & v_txtfcba & "' "
			set rsCat = connStr.Execute(q)
			if not rsCat.eof then
			   fcbaname = rsCat("fcName")
			   CompanyName=rsCat("FCCOMPANYNAME")
			   MgrName = rsCat("FCMANAGER")
			   KTUName = rsCat("FCKTU")
			   CropName=rsCat("FCCROP")
			   FcbaType=rsCat("fcType")
		else
		    nSpanTotal = nSpanTotal - 1 
		end if


	  '  if  BANKCODE <> "--ALL--"  and BANKCODE <> "" then
			' whereCons = whereCons & " AND p.BANKCODE = '" & BANKCODE & "'"
	  '  else
	  '  nSpanTotal = nSpanTotal - 1  
	  '  end if   
' End Added

q="SELECT CONTROLJOB FROM BANK WHERE FCCODE ='" & BANKCODE & _
  "' AND  FCBA= '" & v_txtfcba & "'"
Set rsCARI = Server.CreateObject("ADODB.Recordset")
rsCARI.CursorType = 3
rsCARI.CursorLocation = 3
rsCARI.Open q, connstr	 
if Not rsCARI.eof Then
   txtControlJob=rsCARI("CONTROLJOB")
else
   txtControlJob=""   
End if
rsCari.Close
set rsCari=nothing

q="SELECT nvl(sum(OPENING_BALANCE),0) as OPENING_BALANCE  FROM TRIALBALANCE WHERE FCCODE ='" & txtControlJob & _
  "' AND  FCBA= '" & v_txtfcba & "'" & _
  "  AND PERIOD_MONTH = to_char(to_Date('" & DATECREATED_From & "','MM/DD/YYYY'),'MM')" & _
  "  AND PERIOD_year = to_char(to_Date('" & DATECREATED_From & "','MM/DD/YYYY'),'YYYY')"
Set rsCARI = Server.CreateObject("ADODB.Recordset")
rsCARI.CursorType = 3
rsCARI.CursorLocation = 3
'response.Write q
response.Flush()
rsCARI.Open q, connstr	 
if Not rsCARI.eof Then
   nOpening=rsCARI("OPENING_BALANCE")
else
   nOpening=0
End if
rsCari.Close
set rsCari=nothing

' Amount diminuskan
' Karena kalau debet adalah kas bekurang
' Kalau Kredit maka kas bertambah
'-------------------------------- Control Job  -----------------------------------------------------------------------------------
txtSQL1 = "SELECT distinct p.datecreated,p.bankcode,pay.vouchercode, pay.jobcode, job.fcname, pay.locationtype," & _
		 " pay.locationcode, nvl(sum(pay.amount)*-1,0) as amount, pay.REFERENCE, pay.remarks, pay.FCBA" & _
		 " FROM paymentvoucherdetail pay, job job,paymentvoucher p" & _
		 " WHERE ( ((pay.jobcode = job.fccode) AND (pay.FCBA = job.FCBA)) and " & _
		 "((pay.vouchercode = p.vouchercode) AND (pay.FCBA = p.FCBA)) )" & whereCons &" " 	
if  len(trim(DATECREATED_From)) > 0 then
    txtSQL1 = txtSQL1 & " and p.datecreated >=to_Date('" & DATECREATED_From & "','MM/DD/YYYY')"
	txtSelect1 = " From " & FormatDateTime(DATECREATED_From,1)
end if
if  len(trim(DATECREATED_To)) > 0 then
    txtSQL1 =txtSQL1 &  " and p.datecreated <=to_Date('" & DATECREATED_To & "','MM/DD/YYYY')"
	txtSelect1 = txtSelect1 & "  to " & FormatDateTime(DATECREATED_To,1)
end if
if  len(trim(BANKCODE)) > 0 then
    txtSQL1 = txtSQL1 & " and (p.BANKCODE='" & BANKCODE & "')"
	txtSelect1 = txtSelect1 & " Bank Code " & BANKCODE
end if


txtsql1=(txtsql1) & " group by p.datecreated,p.bankcode,pay.vouchercode, pay.jobcode, job.fcname, pay.locationtype," & _
		 " pay.locationcode,pay.REFERENCE, pay.remarks, pay.FCBA"

'-------------------------------- Jurnal ----------------------------------------------------------------------------------------

txtSQL2 = "SELECT distinct p.datecreated,p.bankcode,pay.vouchercode, pay.jobcode, job.fcname, pay.locationtype," & _
		 " pay.locationcode, nvl(sum(pay.amount),0) as amount, pay.REFERENCE, pay.remarks, pay.FCBA" & _
		 " FROM paymentvoucherdetail pay, job job,paymentvoucher p" & _
		 " WHERE ( ((pay.jobcode = job.fccode) AND (pay.FCBA = job.FCBA)) and " & _
		 "((pay.vouchercode = p.vouchercode) AND (pay.FCBA = p.FCBA)) )" & whereCons &" " 
if  len(trim(DATECREATED_From)) > 0 then
    txtSQL2 = txtSQL2 & " and p.datecreated >=to_Date('" & DATECREATED_From & "','MM/DD/YYYY')"
	txtSelect2 = " From " & FormatDateTime(DATECREATED_From,1)
end if
if  len(trim(DATECREATED_To)) > 0 then
    txtSQL2 =txtSQL2 &  " and p.datecreated <=to_Date('" & DATECREATED_To & "','MM/DD/YYYY')"
	txtSelect2 = txtSelect2 & "  to " & FormatDateTime(DATECREATED_To,1)
end if
if  len(trim(BANKCODE)) > 0 then
    txtSQL2 = txtSQL2 & " and (pay.jobcode='" & txtControlJob & "')"
	txtSelect2 = txtSelect2 & " Bank Code " & BANKCODE
end if

txtsql2=(txtsql2) & " group by p.datecreated,p.bankcode,pay.vouchercode, pay.jobcode, job.fcname, pay.locationtype," & _
		 " pay.locationcode,pay.REFERENCE, pay.remarks,pay.FCBA order by FCBA, datecreated,bankcode,vouchercode"

txtsql = txtsql1 & " union all " & txtsql2
txtSelect = txtSelect1
if len(trim(txtSelect)) > 0 then
   txtSelect = "Selected  by : " & txtSelect
else
   txtSelect = "Selected All Data"
end if

Set rscat = Server.CreateObject("ADODB.Recordset")
	rscat.Open txtSelect, connstr, adOpenStatic 
	
response.Write(txtsql)
response.Flush()
%>
 <body>
		<div id="wrapper">
		<!--#include file="../print_excel.asp"-->
		<form name="frmEntry" method="post" action="p_ArusCash.asp">
		<table id="menu-table" class="centered-margin">
			<tr>
				<td align="center"><a href="/login.asp">Back to Home</a></td>
				<td align="center"><a href="/report/Dashboard.asp">Back to Report</a></td>
			</tr>
		</table>
	<table border="0" id="form-table" class="centered-margin">
		<tr>
			<td>
				<font face="Verdana" size="1">Kebun</font> 
			</td>
			<td>
				<select name="txtFcba" id="vFcba" onchange="javascript:filtering()">
				<%
				Dim rsFcba
				Set rsFcba = Server.CreateObject("ADODB.Recordset")
				rsFcba.CursorLocation = adUseClient
				strsql =    " SELECT '--ALL--' as fcba from DUAL UNION " & _  
				            " select fccode fcba from businessunit " & _
							" where " & _
							" USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fccode) IS NOT NULL " 
				  
				
				rsFcba.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
				if not rsFcba.eof then
					DO while not rsFcba.eof
						if v_txtFcba = rsFcba("fcba") then
						%>			
							<option value="<%=rsFcba("fcba")%>" Selected><%=rsFcba("fcba")%></option>
						<% Else %>	
							<option value="<%=rsFcba("fcba")%>"><%=rsFcba("fcba")%></option>
						<% End If
						
						rsFcba.movenext
					loop
				End if
				rsFcba.Close
				set rsFcba = Nothing
				%>
				</select>	
			</td>
			<td>
				<font face="Verdana" size="1">From</font> 
			</td>
			<td >
				<input class="datepicker" type="text" maxlength="10" name="DATECREATED_From" size="10" value="<%=DATECREATED_From%>" title="Type Date mm/dd/yyyy" autocomplete ="off" >
			</td>
			<td>
				<font face="Verdana" size="1">To</font>
			</td>
			<td>
				<input class="datepicker" type="text" maxlength="10" name="DATECREATED_To" size="10" value="<%=DATECREATED_To%>"title="Type Date mm/dd/yyyy" autocomplete ="off">
			</td>
		
			<td>
				<font face="Verdana" size="1">Bank</font>
			</td>
			<td align="center">
				<select name="BANKCODE">	
					<option></option>
					<%
						Dim rsBlok
						Set rsBlok = Server.CreateObject("ADODB.Recordset")
						rsBlok.CursorLocation = adUseClient
						' if len(trim(BANKCODE)) > 0 then 
							strsql =  " SELECT CONCAT(CONCAT(FCCODE,' : '),FCNAME) AS BANKCODE, fccode FROM BANK WHERE FCBA ='" & v_txtFcba & "'"
						' else
						' 	strsql =  " SELECT '--ALL--' as BANKCODE,'' fccode from DUAL UNION SELECT CONCAT(CONCAT(FCCODE,' : '),FCNAME) AS BANKCODE, fccode FROM BANK WHERE FCBA ='" & Session("fcBA") & "'"
						' end if
						
						rsBlok.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
						if not rsBlok.eof then
							DO while not rsBlok.eof
							if BANKCODE = rsBlok("BANKCODE") then
					%>			<option value="<%=rsBlok("fcCode")%>" Selected><%=rsBlok("BANKCODE")%></option>
							<%else%>
								<option value="<%=rsBlok("fcCode")%>"><%=rsBlok("BANKCODE")%></option>
							<%end if%>
					<%			rsBlok.movenext
							loop
						End if
						rsBlok.Close
						set rsBlok = nothing
					%>
				</select>		
			</td>

	   		 <td colspan="2"><input type="submit" value="View" ></td>
	</tr>
</table>
         <table border="0" width="100%" id="product-table" class="xthf-site">
         <thead>
         <tr>
		   
         <th colspan="11" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
               
                <div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h1 class="judul">LAPORAN ARUS CASH</h1>
				</div>
				<div style="clear:both;"></div>
			</div>
            </th>
			</tr>

				<!--<div class="printhide">-->
				<tr>
					<th colspan="11" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></th>
				</tr>
				<!--</div>-->
		<th> Kebun</th>
		<th> Bank</th>
		<th> Date</th>
		<th> Voucher Code </th>
		<th> Activity <br> Code</th>
		<th> Activity<br>Name</th>
		<th> Target<br>Type</th>
		<th> Target<br>Code</th>
		<th> Reference</th>
		<th> Remarks</th>
		<th> Amount</th>
	</tr>
	</thead>
	<tbody>
	<%
	Dim v_Opening,v_Debit,v_Credit,v_Saldo	
	
	if Not rsCat.eof then
	   nCountofField=rsCat.fields.count-1
	NO=1
    v_Amount = 0%>
	<tr>
	<td align="left" style="<%=cmdstyleLeft%>" colspan="10"><font face="Verdana" size="1">Opening Balance.........</td>
	<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(nOpening,2)%></td>
	
	<%
	v_Amount=cdbl(nOpening)
	While Not rsCat.EOF%>
		<tr>
		<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fcba")%></td>
		<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("bankcode")%></td>
		<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Datecreated")%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Vouchercode")%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Jobcode")%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Fcname")%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Locationtype")%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Locationcode")%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Reference")%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Remarks")%></td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("Amount"),2)%></td>
		</tr>
	  <%
	  v_Amount=v_Amount+trim(cdbl(rsCat("Amount")))
	  'response.write cdbl(trim(rsCat("MTHOPENBALANCE")))+1
	  rsCat.MoveNext
	  'response.Flush()
	  no=no+1
	Wend
	rsCat.close()
	rsCat=null%>
	<input  type="hidden" value="<%=txtTable%>" name="txtTable">
	<input  type="hidden" value="<%=txtModule%>" name="txtModule">
	</form>
	<tr>
	<td align="left" style="<%=cmdstyleLeft%>" colspan="10"> <font face="Verdana" size="1">Ending Balance.........</td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_Amount,2)%></td>
	</tr>
	</tbody>
		
    <%		
	end if %>
</table>
<!--#include file="../Bottom_Report.asp"-->
		</div>
	</body>
</html>