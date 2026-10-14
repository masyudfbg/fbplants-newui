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

	Dim rsEPS, strEPS,rscomp , strcomp,	strEPS2
	Dim v_txtyear, v_txtmonth, v_txttodaydate, v_txtDivision, v_txtBlok, v_txtPlantingDate, NO, cClass
	Dim strMain, nSpanTotal
	Dim txtyear, txtmonth,txtday, txttodaydate, txtDivision, txtBlok, txtPlantingDate,txtComp,v_txtComp
    Dim tempDivision, tempBlok, tempPlantingDate ,tempownership
	Dim tempfcba,tempCompany,tempRegion,tempGroupRegion,tempPlantation,tempGroupPlantation
	Dim nSubMandays_Amount_hi, nSubMandays_Amount_shi
	Dim nOutput_Mandays_Amount_hi, nOutput_Mandays_Amount_shi
	DIM q,FcbaName,CompanyName,MgrName,KTUName,CropName,FcbaType
	Dim SprevFcba, sprevDivision, sprevPlantingYear, sprevFieldCode,SprevPlantingDate
	Dim SprevCompany,SprevRegion,SprevGroupRegion,SprevPlantation,Sprevownership,SprevBlok


		SprevPlantation= ""
		SprevGroupRegion= ""
		SprevRegion= ""
		SprevCompany= ""
		sprevFcba = ""
		sprevDivision = ""
		sprevPlantingYear = ""
		sprevFieldCode = ""
		Sprevownership = ""
		SprevPlantingDate = ""
		SprevBlok = ""
		
		
	        v_txtGroupPlantation  = Request("txtGroupPlantation")
			v_txtPlantation  	= Request("txtPlantation")
			v_txtGroupRegion  	= Request("txtGroupRegion")
			v_txtRegion  		= Request("txtRegion")
			v_txtcompany  		= Request("txtCompany")
			v_txtfcba 			= Request("txtfcba")
			v_txtDivision		= Request("txtDivision")
			v_txtBlok			= Request("txtBlok")
			v_txtPlantingDate	= Request("txtPlantingDate")
			v_txttodaydate		= Request("txttodaydate")
			
			txtyear 		= YEAR(Request("txttodaydate"))
			txtmonth 		= MONTH(Request("txttodaydate"))
			txtday			= DAY(Request("txttodaydate")) 
			txttodaydate 	= Request("txttodaydate")
			
			txtComp         = Request("txtCompany")
			nSpanTotal = 8
			v_txtComp           =   txtComp
			v_txtyear			=	txtyear
			v_txtmonth			=	txtmonth
			
			
			
			'response.write  "GP " & v_txtGroupPlantation & "  P " & v_txtPlantation & "  GR " & v_txtGroupRegion & " R " & v_txtRegion & " C " & v_txtcompany & " S " & v_txtfcba  &  " D " & v_txtDivision & " B "  & v_txtBlok & " TT" & v_txtPlantingDate
			
			Dim selectFields,strFieldCons,selectCons,whereCons,GroupCons,orderCons,orderBlok,orderPlantingDate
			strFieldCons = " USER_COMPANY_API.GetGroupPlantationCode('" & Session("CodeName") & "',fcba) GroupPlantation,USER_COMPANY_API.GetPlantationCode('" & Session("CodeName") & "',fcba) Plantation, " & _
						   " USER_COMPANY_API.GetGroupRegionCode('" & Session("CodeName") & "',fcba) groupregion ,USER_COMPANY_API.GetRegionCode('" & Session("CodeName") & "',fcba) region,USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fcba) company, fcba, BUSINESSUNIT_API.get_OrderData(fcba) norut, " 

						   
			wherecons = " where company is not null "
			if v_txtGroupPlantation <> "--SUMMARY--"  and v_txtGroupPlantation <> "--ALL--" and v_txtGroupPlantation <> "" then
			   whereCons = wherecons & " And GroupPlantation = '" & v_txtGroupPlantation & "' "
			 
			end if
			if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
			    if v_txtPlantation <> "--ALL--" then
				   whereCons = whereCons & " And  Plantation = '" & v_txtPlantation & "' "
				end if   
				selectCons = selectCons & ",Plantation"
				orderCons  = orderCons & ", PLANTATION_API.get_OrderData(Plantation) "
			else
			    nSpanTotal = nSpanTotal - 1 
			end if
			if v_txtGroupRegion <> "--SUMMARY--"   and v_txtGroupRegion <> "" then
			    if v_txtGroupRegion <> "--ALL--" then
				   whereCons = whereCons & " And  GroupRegion = '" & v_txtGroupRegion & "' "
				end if
				selectCons = selectCons & " ,GroupRegion"
				orderCons  = orderCons & ", GROUPREGION_API.get_OrderData(GroupRegion) "
			else
			    nSpanTotal = nSpanTotal - 1 
			end if
			if  v_txtRegion <> "--SUMMARY--"  and v_txtRegion <> "" then
			    if v_txtRegion <> "--ALL--" then
				   whereCons = whereCons & " And  Region = '" & v_txtRegion & "' "
				end if
				selectCons = selectCons & " ,Region"
				orderCons  = orderCons & ", REGION_API.get_OrderData(Region) "
			else
			    nSpanTotal = nSpanTotal - 1 
			end if
			if  v_txtCompany <> "--SUMMARY--" and v_txtCompany <> "" then
			    if v_txtCompany  <> "--ALL--" then
				   whereCons = whereCons & " And  company = '" & v_txtCompany & "' "
				end if
				selectCons = selectCons & " ,company"
			    orderCons  = orderCons & ", COMPANY_API.get_OrderData(company) "
				q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where FCCOMPANYCODE = '" & v_txtCompany & "'"	
				set rsCat = connStr.Execute(q)
				if not rsCat.eof then
				   CompanyName=rsCat("FCCOMPANYNAME")			   
			    end if
			else
			    nSpanTotal = nSpanTotal - 1 	
			end if
	        if  v_txtfcba <> "--SUMMARY--"  and v_txtfcba <> "" then
			    if v_txtfcba  <> "--ALL--" then
			       whereCons = whereCons & " And  ff.fcba = '" & v_txtfcba & "' "
				end if
			    selectCons = selectCons & ",ff.fcba"
			    orderCons  = orderCons & ",BUSINESSUNIT_API.get_OrderData(fcba)"
				q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & v_txtfcba & "' "
				set rsCat = connStr.Execute(q)
				if not rsCat.eof then
				   fcbaname = rsCat("fcName")
				   CompanyName=rsCat("FCCOMPANYNAME")
				   MgrName = rsCat("FCMANAGER")
				   KTUName = rsCat("FCKTU")
				   CropName=rsCat("FCCROP")
				   FcbaType=rsCat("fcType")
			    end if
			else
			    nSpanTotal = nSpanTotal - 1 
			end if

Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

Dim v_txtACCOUNTINGYEARPERIOD


v_txtDATECREATED_From= trim(request("DATECREATED_From"))
v_txtDATECREATED_To=trim(request("DATECREATED_To"))
v_txtACCOUNTINGYEARPERIOD=trim(request("ACCOUNTINGYEARPERIOD"))
v_txtBANKCODE=trim(request("BANKCODE"))
txtFCNAME=trim(request("FCNAME"))
txtControlJob=""
nOpening=0

q="SELECT CONTROLJOB,FCNAME FROM BANK WHERE FCCODE ='" & v_txtBANKCODE & _
  "' AND  FCBA= '" & v_txtfcba & "'"
Set rsCARI = Server.CreateObject("ADODB.Recordset")
rsCARI.CursorType = 3
rsCARI.CursorLocation = 3
rsCARI.Open q, connstr	 
if Not rsCARI.eof Then
   txtControlJob=rsCARI("CONTROLJOB")
   txtFCNAME=rsCARI("FCNAME")
else
   txtControlJob=""   
End if
rsCari.Close
set rsCari=nothing

q="SELECT nvl(sum(OPENING_BALANCE),0) as OPENING_BALANCE  FROM TRIALBALANCE WHERE FCCODE ='" & txtControlJob & _
  "' AND  FCBA= '" & v_txtfcba & "'" & _
  "  AND PERIOD_MONTH = to_char(to_Date('" & v_txtDATECREATED_From & "','MM/DD/YYYY'),'MM')" & _
  "  AND PERIOD_year = to_char(to_Date('" & v_txtDATECREATED_From & "','MM/DD/YYYY'),'YYYY')"
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
txtSQL1 = "SELECT p.datecreated,p.bankcode,pay.vouchercode, pay.jobcode, job.fcname, pay.locationtype," & _
		 " pay.locationcode, nvl(sum(pay.amount)*-1,0) as amount, pay.REFERENCE, pay.remarks, pay.fcba" & _
		 " FROM paymentvoucherdetail pay, job job,paymentvoucher p" & _
		 " WHERE ( ((pay.jobcode = job.fccode) AND (pay.FCBA = job.FCBA)) and " & _
		 "((pay.vouchercode = p.vouchercode) AND (pay.FCBA = p.FCBA)) )" 
if  len(trim(v_txtDATECREATED_From)) > 0 then
    txtSQL1 = txtSQL1 & " and p.datecreated >=to_Date('" & v_txtDATECREATED_From & "','MM/DD/YYYY')"
	txtSelect1 = " From " & FormatDateTime(v_txtDATECREATED_From,1)
end if
if  len(trim(v_txtDATECREATED_To)) > 0 then
    txtSQL1 =txtSQL1 &  " and p.datecreated <=to_Date('" & v_txtDATECREATED_To & "','MM/DD/YYYY')"
	txtSelect1 = txtSelect1 & "  to " & FormatDateTime(v_txtDATECREATED_To,1)
end if
if  len(trim(v_txtBANKCODE)) > 0 then
    txtSQL1 = txtSQL1 & " and (p.BANKCODE='" & v_txtBANKCODE & "')"
	txtSelect1 = txtSelect1 & " Bank Code " & v_txtBANKCODE
end if

if  len(trim(v_txtACCOUNTINGYEARPERIOD)) > 0 then
    txtSQL1 = txtSQL1 & " and p.accountingyear||lpad(p.accountingperiod,2,0) =" & v_txtACCOUNTINGYEARPERIOD 
	txtSelect1 = txtSelect1 & " , ACCOUNTINGYEAR PERIOD " & left(v_txtACCOUNTINGYEARPERIOD,4) & " - " & Right(v_txtACCOUNTINGYEARPERIOD,2)
end if
txtsql1=(txtsql1) & " group by p.datecreated,p.bankcode,pay.vouchercode, pay.jobcode, job.fcname, pay.locationtype," & _
		 " pay.locationcode,pay.REFERENCE, pay.remarks, pay.fcba"

'-------------------------------- Jurnal ----------------------------------------------------------------------------------------

txtSQL2 = "SELECT p.datecreated,p.bankcode,pay.vouchercode, pay.jobcode, job.fcname, pay.locationtype," & _
		 " pay.locationcode, nvl(sum(pay.amount),0) as amount, pay.REFERENCE, pay.remarks, pay.fcba" & _
		 " FROM paymentvoucherdetail pay, job job,paymentvoucher p" & _
		 " WHERE ( ((pay.jobcode = job.fccode) AND (pay.FCBA = job.FCBA)) and " & _
		 "((pay.vouchercode = p.vouchercode) AND (pay.FCBA = p.FCBA)) )" 
if  len(trim(v_txtDATECREATED_From)) > 0 then
    txtSQL2 = txtSQL2 & " and p.datecreated >=to_Date('" & v_txtDATECREATED_From & "','MM/DD/YYYY')"
	txtSelect2 = " From " & FormatDateTime(v_txtDATECREATED_From,1)
end if
if  len(trim(v_txtDATECREATED_To)) > 0 then
    txtSQL2 =txtSQL2 &  " and p.datecreated <=to_Date('" & v_txtDATECREATED_To & "','MM/DD/YYYY')"
	txtSelect2 = txtSelect2 & "  to " & FormatDateTime(v_txtDATECREATED_To,1)
end if
if  len(trim(v_txtBANKCODE)) > 0 then
    txtSQL2 = txtSQL2 & " and (pay.jobcode='" & txtControlJob & "')"
	txtSelect2 = txtSelect2 & " Bank Code " & v_txtBANKCODE
end if
if  len(trim(v_txtACCOUNTINGYEARPERIOD)) > 0 then
    txtSQL2 = txtSQL2 & " and p.accountingyear||lpad(p.accountingperiod,2,0) =" & v_txtACCOUNTINGYEARPERIOD 
end if
txtsql2=(txtsql2) & " group by p.datecreated,p.bankcode,pay.vouchercode, pay.jobcode, job.fcname, pay.locationtype," & _
		 " pay.locationcode,pay.REFERENCE, pay.remarks, pay.fcba order by fcba,bankcode,datecreated,vouchercode"

txtsql = txtsql1 & " union all " & txtsql2
txtSelect = txtSelect1
if len(trim(txtSelect)) > 0 then
   txtSelect = "Selected  by : " & txtSelect
else
   txtSelect = "Selected All Data"
end if
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
	<!--#include file="FormSelectNewSum.html"-->
<table border="0" id="form-table" class="centered-margin">
    
	<tr>
		<th>TANGGAL</th>
		<th>ACCOUNTING PERIOD</th>
		<th>&nbsp;</th>
	</tr>
	<input type="hidden" name="txtComp" value="<%=v_txtComp%>">
	<tr>
	    <!--%response.write v_txtfcba %-->
	    <!--%response.write v_txtyear %-->
	    <!--%response.write v_txtmonth %-->
		<!--%response.write len(trim(request("txtYear"))) %-->
		
		<td align="center">
		<% if v_txttodaydate <> "" and v_txttodaydate = v_txttodaydate then %>
			<input class="datepicker" type="text" maxlength="10" name="txttodaydate" size="10" value="<%=v_txttodaydate%>" title="Type Date mm/dd/yyyy" >
		<% else%>
			<input class="datepicker" type="text" maxlength="10" name="txttodaydate" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
		<% end if%>
			<!--<a class="calendar" href="javascript:show_calendar('frmEntry.txttodaydate');" onMouseOver="window.status='Date Picker';return true;" onMouseOut="window.status='';return true;">
				<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top>
			</a>-->
		</td>
		
		<td align="center">
			<% 	Dim rsComboA,Data_ACCOUNTINGYEARPERIOD
				q = "select accountingyear||lpad(accountingperiod,2,0) ACCOUNTINGYEARPERIOD,Description from  ACCOUNTINGYEARPERIOD where fcBA='" &  Session("fcBA") & "' order by accountingyear||lpad(accountingperiod,2,0)"

				Set rsComboA = Server.CreateObject("ADODB.Recordset")
				'response.write q
				rsComboA.CursorType = 3
				rsComboA.CursorLocation = 3
				rsComboA.Open q, connstr
				if not rsComboA.eof then%>
				<select name="ACCOUNTINGYEARPERIOD">    
					<% Do While Not rsComboA.eof %>
					<% If Request("ACCOUNTINGYEARPERIOD") = rsComboA("ACCOUNTINGYEARPERIOD") Then 
						Data_ACCOUNTINGYEARPERIOD = "true"
					else
						Data_ACCOUNTINGYEARPERIOD = "false"
					End if%>
					<option value="<%=rsComboA("ACCOUNTINGYEARPERIOD")%>" <% If Request("ACCOUNTINGYEARPERIOD") = rsComboA("ACCOUNTINGYEARPERIOD") Then %> Selected <%End if%>><%=rsComboA("Description")%></option> 
					<%  rsComboA.movenext
					   loop
				if Data_ACCOUNTINGYEARPERIOD = "false" then %>
					<option value="<%=ACCOUNTINGYEAR&ACCOUNTINGPERIOD%>" Selected>Tahun <%=ACCOUNTINGYEAR%> Bulan <%=ACCOUNTINGPERIOD%></option>
				<% end if %>
				
				<%
					End if
					rsComboA.Close
					set rsComboA = Nothing
					q=""
					%>
			</select> 		
		</td>
		<td align="center"><input type="submit" value="View" /></td>

	</tr>
</table>
</form>
         <table border="0" width="100%" id="product-table" class="xthf-site">
         <thead>
         <tr>
		   
         <th colspan="11" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
		<div id="header">
			<% if  CompanyName <>  "" then %>
				<h4>PT.<%=CompanyName%></h4> 
			<% End if%>
			<% If  FcbaName <> ""  then %>
				   <h5><%=FcbaName%></h5>   
			<% End if%>
		</div>
		<div id="subheader" class="centered">
			<h2>LAPORAN ARUS CASH</h2>
			<% if v_txttodaydate <> "" and v_txtFcba <> "" then %>
				<h5>PERIODE : <%=v_txttodaydate%>, KEBUN : <%=v_txtFcba%></h5>
			<% end if %>
			<h3></br></h3>
			<h3></br></h3>
		</div>
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
		<th>Kebun</th>
		<th>Bank</th>
		<th>Date</th>
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
	if  len(trim(v_txtACCOUNTINGYEARPERIOD)) > 0 then
        Set rsCat = connStr.Execute(txtSQL)
	
	
	if Not rsCat.eof then
	   nCountofField=rsCat.fields.count-1
	end if
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