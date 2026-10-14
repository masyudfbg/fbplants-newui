<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - Gang Activity and Attendance</title>
<script language="JavaScript" src='/js/x.js'></script>
<script language="JavaScript" src='/js/xtableheaderfixed.js'></script>
<script language="JavaScript">
xAddEventListener(window, 'load',
  function() {
    new xTableHeaderFixed('xthf-site', window);
    new xTableHeaderFixed('xthf-green', window);
    new xTableHeaderFixed('xthf-blue', window);
  }, false
);
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
</script>

<style type="text/css">
@media print
{
.printhide{display:none;}
}
</style>
</head>

<%
Dim Columnx : Columnx = 0
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3

Dim tbl 
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

v_txtFDDATE_From=trim(REQUEST("FDDATE_From"))
v_txtFDDATE_To=trim(REQUEST("FDDATE_To"))
v_txtGANGCODE=trim(REQUEST("GANGCODE"))
v_txtJOBCODE=trim(REQUEST("JOBCODE"))
v_txtEMPLOYEECODE=trim(REQUEST("EMPLOYEECODE"))
v_txtLOCATIONTYPE=trim(REQUEST("LOCATIONTYPE"))
v_txtLOCATIONCODE=trim(REQUEST("LOCATIONCODE"))
v_txtsectionname = Request("txtsectionname")
v_cField=trim(REQUEST("cField"))
ATTENDANCE= trim(request("ATTENDANCE"))
businessunit= trim(request("businessunit"))
rowstate= trim(request("rowstate"))

if rowstate="Approved" Then
	tbl = "attendance_gad"
Else
    tbl = "attendance_gad_temp"
End if

		q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & businessunit & "' "
			' response.write q
			set rsCat = connStr.Execute(q)
			if not rsCat.eof then
			   fcbaname = rsCat("fcName")
			   CompanyName=rsCat("FCCOMPANYNAME")
			   MgrName = rsCat("FCMANAGER")
			   KTUName = rsCat("FCKTU")
			   CropName=rsCat("FCCROP")
			   FcbaType=rsCat("fcType")
			end if
if  len(trim(v_cField)) = 0 then
	txtSQL = 	" select fcba" & _
				"        , nvl(MAX(mandays),0) mandays" & _
				"        , nvl(MAX(othrs),0) overtime" & _
				"        , nvl(MAX(unit),0) unit" & _
				"        , nvl(MAX(kg_brondolan),0) kg_brondolan" & _
				"        , nvl(MAX(output),0) output" & _
				"        , nvl(MAX(rate),0) rate" & _
				"        , nvl(MAX(bucket),0) bucket" & _
				"        , nvl(MAX(amount),0) amount_allowded" & _
				"        , nvl(MAX(kg_janjang),0) jumlah_janjang" & _
				"        , nvl(MAX(bunch_weight),0) bunch_weight" & _
				" from (" & _
				"     select gad.gangcode" & _
				"            , gad.fddate" & _
				"            , gad.supervision_1" & _
				"            , gad.supervision_2" & _
				"            , gad.supervision_3" & _
				"            , gad.supervision_4" & _
				"            , gad.employeecode" & _
				"            , EMPLOYEE_API.get_EmployeeName(gad.fcba, gad.employeecode) employeename" & _
				"            , EMPLOYEE_API.get_sectionname(gad.fcba, gad.employeecode) sectionname" & _
				"            , gad.attendance" & _
				"            , gad.jobcode" & _
				"            , JOB_API.get_jobcode_name(gad.jobcode, gad.fcba) jobname" & _
				"            , gad.locationtype" & _
				"            , gad.locationcode" & _
				"            , gad.chargejob" & _
				"            , gad.chargetype" & _
				"            , gad.chargecode" & _
				"            , nvl(gad.mandays, 0) mandays" & _
				"            , nvl(gad.othrs, 0) othrs" & _
				"            , nvl(gad.rate, 0) rate" & _
				"            , nvl(gad.unit, 0) unit" & _
				"            , nvl(gad.output, 0) output" & _
				"            , gad.reference" & _
				"            , gad.linenokey" & _
				"            , gad.spbno" & _
				"            , gad.remarks" & _
				"            , gad.bucket" & _
				"            , gad.kg_janjang" & _
				"            , gad.kg_brondolan" & _
				"            , nvl(ae.amount,0) amount" & _
				"            , ae.remarksallowded" & _
				"            , gad.fcba" & _
				"            , gad.rowstate" & _
				"            , nvl(fin.datatype,'NO Finger') Fin_type " & _
				"            , decode(fin.timein,NULL,'-',to_char(fin.timein,'hh24:mi')) Fin_Timein " & _
				"            , decode(fin.timeout,NULL,'-',to_char(fin.timeout,'hh24:mi')) Fin_Timeout " & _
				"            , nvl(gad.bunch_weight, 0) bunch_weight" & _
				"   from " & tbl & " gad " & _
				" 	LEFT JOIN attendance_empallded ae ON (ae.linenokey = gad.linenokey AND ae.fcba = gad.fcba)"& _
				" 	LEFT JOIN (select distinct empno,tdate,timein,timeout,datatype from userfinger.attendanceview) fin ON (gad.employeecode=fin.empno and gad.fddate=fin.tdate)"& _
				"     where gad.fcba = '" & businessunit & "'" & _
				"   UNION " & _
				"     select b.gangcode" & _
				"            , b.bkmdate fddate" & _
				"            , b.supervisi1 supervision_1" & _
				"            , b.supervisi2 supervision_2" & _
				"            , b.supervisi3 supervision_3" & _
				"            , b.supervisi4 supervision_4" & _
				"            , a.employeecode" & _
				"            , EMPLOYEE_API.get_EmployeeName(a.fcba, a.employeecode) employeename" & _
				"            , EMPLOYEE_API.get_sectionname(a.fcba, a.employeecode) sectionname" & _
				"            , a.attendance" & _
				"            , a.actcode jobcode" & _
				"            , jobmaster_api.get_fcname(a.actcode) jobname" & _
				"            , jobmaster_api.get_job_category(a.actcode) locationtype" & _
				"            , a.locationcode" & _
				"            , '' chargejob" & _
				"            , '' chargetype" & _
				"            , '' chargecode" & _
				"            , nvl(a.mandays, 0) mandays" & _
				"            , nvl(a.othconvertion, 0) othrs" & _
				"            , nvl(a.rate, 0) rate" & _
				"            , nvl(a.unit, 0) unit" & _
				"            , nvl(a.output, 0) output" & _
				"            , a.bkmno reference" & _
				"            , a.linenokey" & _
				"            , '' spbno" & _
				"            , a.remarks" & _
				"            , 0 bucket" & _
				"            , 0 kg_janjang" & _
				"            , a.kg_brondolan" & _
				"            , a.fcba" & _
				"            , a.allowance_amt amount" & _
				"            , allowded_api.get_fcname(a.allowance_code,a.fcba) remarksallowded" & _				
				"            , b.rowstate" & _
				"            , '' Fin_type " & _
				"            , decode(a.in_fin,NULL,'-',to_char(a.in_fin,'hh24:mi')) Fin_Timein " & _
				"            , decode(a.out_fin,NULL,'-',to_char(a.out_fin,'hh24:mi')) Fin_Timeout " & _
				"            , nvl(a.bunch_weight, 0) bunch_weight" & _
				"   from bkmupahdtl a inner join bkmhdr b on (a.fcba=b.fcba AND a.bkmno=b.bkmno) " & _
				"   where a.fcba = '" & businessunit & "'" & _
				" ) where fcba = '" & businessunit & "'"
else
	txtSQL =	" select fcba" & _
				"        , rowstate" & _
				"        , " & v_cField & "" & _
				"        , Fin_type" & _
				"        , Fin_Timein" & _
				"        , Fin_Timeout" & _
				"        , nvl(MAX(mandays),0) mandays" & _
				"        , nvl(MAX(othrs),0) overtime" & _
				"        , nvl(MAX(unit),0) unit" & _
				"        , nvl(MAX(output),0) output" & _
				"        , nvl(MAX(rate),0) rate" & _
				"        , nvl(MAX(bucket),0) bucket" & _
				"        , nvl(MAX(amount),0) amount_allowded" & _
				"        , nvl(MAX(kg_brondolan),0) kg_brondolan" & _
				"        , nvl(MAX(kg_janjang),0) jumlah_janjang" & _
				"        , nvl(MAX(bunch_weight),0) bunch_weight" & _
				" from (" & _
				"     select gad.gangcode" & _
				"            , gad.fddate" & _
				"            , gad.supervision_1" & _
				"            , gad.supervision_2" & _
				"            , gad.supervision_3" & _
				"            , gad.supervision_4" & _
				"            , gad.employeecode" & _
				"            , EMPLOYEE_API.get_EmployeeName(gad.fcba, gad.employeecode) employeename" & _
				"            , EMPLOYEE_API.get_sectionname(gad.fcba, gad.employeecode) sectionname" & _
				"            , gad.attendance" & _
				"            , gad.jobcode" & _
				"            , JOB_API.get_jobcode_name(gad.jobcode, gad.fcba) jobname" & _
				"            , gad.locationtype" & _
				"            , gad.locationcode" & _
				"            , gad.chargejob" & _
				"            , gad.chargetype" & _
				"            , gad.chargecode" & _
				"            , nvl(gad.mandays, 0) mandays" & _
				"            , nvl(gad.othrs, 0) othrs" & _
				"            , nvl(gad.rate, 0) rate" & _
				"            , nvl(gad.unit, 0) unit" & _
				"            , nvl(gad.output, 0) output" & _
				"            , gad.reference" & _
				"            , gad.linenokey" & _
				"            , gad.spbno" & _
				"            , gad.remarks" & _
				"            , gad.bucket" & _
				"            , gad.kg_janjang" & _
				"            , gad.kg_brondolan" & _
				"            , gad.fcba" & _
				"            , nvl(ae.amount,0) amount" & _
				"            , ae.remarksallowded" & _
				"            , gad.rowstate" & _
				"            , nvl(fin.datatype,'NO Finger') Fin_type " & _
				"            , decode(fin.timein,NULL,'-',to_char(fin.timein,'hh24:mi')) Fin_Timein " & _
				"            , decode(fin.timeout,NULL,'-',to_char(fin.timeout,'hh24:mi')) Fin_Timeout " & _
				"            , nvl(gad.bunch_weight, 0) bunch_weight" & _
				"     from " & tbl & " gad " & _
				" 	LEFT JOIN attendance_empallded ae ON (ae.linenokey = gad.linenokey AND ae.fcba = gad.fcba)"& _
				" 	LEFT JOIN (select distinct empno,tdate,timein,timeout,datatype from userfinger.attendanceview) fin ON (gad.employeecode=fin.empno and gad.fddate=fin.tdate)"& _
				"   where gad.fcba = '" & businessunit & "'" & _
				"   UNION " & _
				"     select b.gangcode" & _
				"            , b.bkmdate fddate" & _
				"            , b.supervisi1 supervision_1" & _
				"            , b.supervisi2 supervision_2" & _
				"            , b.supervisi3 supervision_3" & _
				"            , b.supervisi4 supervision_4" & _
				"            , a.employeecode" & _
				"            , EMPLOYEE_API.get_EmployeeName(a.fcba, a.employeecode) employeename" & _
				"            , EMPLOYEE_API.get_sectionname(a.fcba, a.employeecode) sectionname" & _
				"            , a.attendance" & _
				"            , a.actcode jobcode" & _
				"            , jobmaster_api.get_fcname(a.actcode) jobname" & _
				"            , jobmaster_api.get_job_category(a.actcode) locationtype" & _
				"            , a.locationcode" & _
				"            , '' chargejob" & _
				"            , '' chargetype" & _
				"            , '' chargecode" & _
				"            , nvl(a.mandays, 0) mandays" & _
				"            , nvl(a.othconvertion, 0) othrs" & _
				"            , nvl(a.rate, 0) rate" & _
				"            , nvl(a.unit, 0) unit" & _
				"            , nvl(a.output, 0) output" & _
				"            , a.bkmno reference" & _
				"            , a.linenokey" & _
				"            , '' spbno" & _
				"            , a.remarks" & _
				"            , 0 bucket" & _
				"            , 0 kg_janjang" & _
				"            , a.kg_brondolan" & _
				"            , a.fcba" & _
				"            , a.allowance_amt amount" & _
				"            , allowded_api.get_fcname(a.allowance_code,a.fcba) remarksallowded" & _				
				"            , b.rowstate" & _
				"            , '' Fin_type " & _
				"            , decode(a.in_fin,NULL,'-',to_char(a.in_fin,'hh24:mi')) Fin_Timein " & _
				"            , decode(a.out_fin,NULL,'-',to_char(a.out_fin,'hh24:mi')) Fin_Timeout " & _
				"            , nvl(a.bunch_weight, 0) bunch_weight" & _
				"   from bkmupahdtl a inner join bkmhdr b on (a.fcba=b.fcba AND a.bkmno=b.bkmno) " & _
				"   where a.fcba = '" & businessunit & "'" & _
				" ) where fcba = '" & businessunit & "'"


end if

if  len(trim(v_txtFDDATE_From)) > 0 then
    txtSQL = txtSQL & " and fddate >= to_date('" & v_txtFDDATE_From & "','mm/dd/yyyy')"
	txtSelect = " From " & FormatDateTime(v_txtFDDATE_From,1)
end if
if  len(trim(v_txtFDDATE_To)) > 0 then
    txtSQL =txtSQL &  " and fddate <= to_date('" & v_txtFDDATE_To & "','mm/dd/yyyy')"
	txtSelect = txtSelect & "  to " & FormatDateTime(v_txtFDDATE_To,1)
end if
if  len(trim(v_txtsectionname)) > 0 then
    txtSQL = txtSQL & " and UPPER(sectionname) like UPPER ('%" & v_txtsectionname & "%')"
	txtSelect = txtSelect & "  Section Name " & (v_txtsectionname)
end if
if  len(trim(v_txtGANGCODE)) > 0 then
    txtSQL = txtSQL & " and UPPER(gangcode) like UPPER ('%" & v_txtGANGCODE & "%')"
	txtSelect = txtSelect & "  Gang Code " & (v_txtGANGCODE)
end if
if  len(trim(v_txtJOBCODE)) > 0 then
    txtSQL =txtSQL &  " and jobcode = '" & v_txtJOBCODE & "'"
	txtSelect = txtSelect & "  Job Code " & (v_txtJOBCODE)
end if 
if  len(trim(v_txtEMPLOYEECODE)) > 0 then
    txtSQL =txtSQL &  " and (UPPER(employeecode) like UPPER ('%" & v_txtEMPLOYEECODE & "%') OR employeename like UPPER ('%"& v_txtEMPLOYEECODE &"%'))"
	txtSelect = txtSelect & "  Emp. Code " & (v_txtEMPLOYEECODE)
end if 
if  len(trim(v_txtLOCATIONTYPE)) > 0 then
    txtSQL =txtSQL &  " and locationtype = '" & v_txtLOCATIONTYPE & "'"
	txtSelect = txtSelect & "  TT " & (v_txtLOCATIONTYPE)
end if 
if  len(trim(v_txtLOCATIONCODE)) > 0 then
    txtSQL =txtSQL &  " and locationcode = '" & v_txtLOCATIONCODE & "'"
	txtSelect = txtSelect & "  TC " & (v_txtLOCATIONCODE)
end if
if  len(trim(ATTENDANCE)) > 0 then
    txtSQL =txtSQL &  " and ATTENDANCE = '" & ATTENDANCE & "'"
	txtSelect = txtSelect & "  ATTENDANCE " & (ATTENDANCE)
end if
if  len(trim(rowstate)) > 0 then
    txtSQL =txtSQL &  " and rowstate = '" & rowstate & "'"
	txtSelect = txtSelect & "  rowstate" & (rowstate)
end if
txtsql	=	txtsql & " group by fcba, fddate, rowstate,Fin_type,Fin_Timeout,Fin_Timein"

if  len(trim(v_cField)) = 0 then
	txtSQL = txtSQL
Else
	txtsql = txtsql & ", " & v_cField	
End if

If len(trim(v_cField)) = 0 Then
	txtsql	=	txtsql & " order by fcba"
Else
	txtsql	=	txtsql & " order by fcba, fddate"
	if instr(v_cField,"gangcode") > 0 then txtsql = txtsql & " , gangcode" End if
	if instr(v_cField,"jobcode") > 0 then txtsql = txtsql & " , jobcode" End if
	if instr(v_cField,"employeecode") > 0 then txtsql = txtsql & " , employeecode" End if	
End If

'response.Write(txtsql)
'response.Flush()
%>
<% Dim Kolom : Kolom = 0
if instr(v_cField,"fddate") > 0 then Kolom = Kolom + 1 End if
if instr(v_cField,"gangcode") > 0 then Kolom = Kolom + 1 End if
if instr(v_cField,"supervision_1") > 0 then Kolom = Kolom + 1 End if
if instr(v_cField,"supervision_2") > 0 then Kolom = Kolom + 1 End if
if instr(v_cField,"supervision_3") > 0 then Kolom = Kolom + 1 End if
if instr(v_cField,"supervision_4") > 0 then Kolom = Kolom + 1 End if
if instr(v_cField,"sectionname") > 0 then Kolom = Kolom + 1 End if
if instr(v_cField,"employeecode") > 0 then Kolom = Kolom + 1 End if
if instr(v_cField,"employeename") > 0 then Kolom = Kolom + 1 End if
if instr(v_cField,"attendance") > 0 then Kolom = Kolom + 1 End if
if instr(v_cField,"jobcode") > 0 then Kolom = Kolom + 1 End if
if instr(v_cField,"jobname") > 0 then Kolom = Kolom + 1 End if
if instr(v_cField,"locationtype") > 0 then Kolom = Kolom + 1 End if
if instr(v_cField,"locationcode") > 0 then Kolom = Kolom + 1 End if
if instr(v_cField,"chargejob") > 0 then Kolom = Kolom + 1 End if
if instr(v_cField,"chargetype") > 0 then Kolom = Kolom + 1 End if
if instr(v_cField,"chargecode") > 0 then Kolom = Kolom + 1 End if
if instr(v_cField,"allowance_code") > 0 then Kolom = Kolom + 1 End if
if instr(v_cField,"amount") > 0 then Kolom = Kolom + 1 End if
if instr(v_cField,"remarks") > 0 then Kolom = Kolom + 1 End if
if instr(v_cField,"reference") > 0 then Kolom = Kolom + 1 End if
if instr(v_cField,"spbno") > 0 then Kolom = Kolom + 1 End if
if instr(v_cField,"remarksallowded") > 0 then Kolom = Kolom + 1 End if
if instr(v_cField,"linenokey") > 0 then Kolom = Kolom + 1 End if

Kolom = Kolom
%>
<body>
	<div id="wrapper">
        <!--#include file="print_excel_blob.asp"-->
		<table border="0" width="100%" id="product-table" class="xthf-site">
			<thead>
				<tr>
					<%if len(trim(v_cField)) = 0 then%>
					<th colspan="<%=20+Kolom%>" style=" background-color:#FFFFFF; color:#000000; border:none">
					<%Else%>
					<th colspan="<%=20+Kolom%>" style=" background-color:#FFFFFF; color:#000000; border:none">
					<%End If%>
						<div id="header-wrapper">
							<div id="head_judul" style="float:left; text-align:left;">
								<h4>PT. <%=COmpanyName%></h4>

								<h5><%=fcbaname%></h5>
							</div>
							<div id="subheader" class="centered" style="width:50%; margin:auto;">
								<h2>EMPLOYEE ACTIVITY / BKM</h2>
							</div>
						</div>
					</th>
				</tr>
						<!--<div class="printhide">-->
						<tr>
						<%if len(trim(v_cField)) = 0 then%>
						<th colspan="<%=20+Kolom%>" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></th>
						<%Else%>
						<th colspan="<%=20+Kolom%>" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></th>
						<%End If%>
						</tr>
						<!--</div>-->


	<%	
	Set rsCat = connStr.Execute(txtSQL)
	if Not rsCat.eof then
	   nCountofField=rsCat.fields.count-1
	
	end if
	NO=1%>
	<tr>
		<%if len(trim(v_cField)) = 0 then%>
		<th style="color:#fff;">No</th>
		<%End If%>
		<%if instr(v_cField,"fddate") > 0 then%> <th style="color:#fff;">Date</th><%End if%>
		<%if instr(v_cField,"linenokey") > 0 then%> <th style="color:#fff;">Line No Key</th> <%End if%>
		<%if instr(v_cField,"gangcode") > 0 then%> <th style="color:#fff;">Gang Code</th> <%End if%>
		<%if instr(v_cField,"supervision_1") > 0 then%> <th style="color:#fff;">Supervisi 1</th> <%End if%>
		<%if instr(v_cField,"supervision_2") > 0 then%> <th style="color:#fff;">Supervisi 2</th> <%End if%>
		<%if instr(v_cField,"supervision_3") > 0 then%> <th style="color:#fff;">Supervisi 3</th> <%End if%>
		<%if instr(v_cField,"supervision_4") > 0 then%> <th style="color:#fff;">Supervisi 4</th> <%End if%>
		<%if instr(v_cField,"sectionname") > 0 then%> <th style="color:#fff;">Section Name</th> <%End if%>
		<%if instr(v_cField,"employeecode") > 0 then%> <th style="color:#fff;">Employee Code</th> <%End if%>
		<%if instr(v_cField,"employeename") > 0 then%> <th style="color:#fff;">Employee Name</th> <%End if%>
		<%if instr(v_cField,"attendance") > 0 then%> <th style="color:#fff;">Attendance</th> <%End if%>
		<%if instr(v_cField,"jobcode") > 0 then%> <th style="color:#fff;">Job Code</th> <%End if%>
		<%if instr(v_cField,"jobname") > 0 then%> <th style="color:#fff;">Job Desc</th> <%End if%>
		<%if instr(v_cField,"locationtype") > 0 then%> <th style="color:#fff;">Target Type</th> <%End if%>
		<%if instr(v_cField,"locationcode") > 0 then%> <th style="color:#fff;">Target Code</th> <%End if%>
		<%if instr(v_cField,"chargejob") > 0 then%> <th style="color:#fff;">Charge Job</th> <%End if%>
		<%if instr(v_cField,"chargetype") > 0 then%> <th style="color:#fff;">Charge Type</th> <%End if%>
		<%if instr(v_cField,"chargecode") > 0 then%> <th style="color:#fff;">Charge Code</th> <%End if%>
		<%if instr(v_cField,"allowance_code") > 0 then%> <th style="color:#fff;">Allowded</th> <%End if%>
		<%if instr(v_cField,"amount") > 0 then%> <th style="color:#fff;">Amount</th> <%End if%>
		<%if instr(v_cField,"reference") > 0 then%> <th style="color:#fff;">Reference</th> <%End if%>
		<%if instr(v_cField,"spbno") > 0 then%> <th style="color:#fff;">SPB No</th> <%End if%>
		<%if instr(v_cField,"remarks") > 0 then%> <th style="color:#fff;">Remarks</th> <%End if%>
		<th style="color:#fff;">Mandays</th>
		<th style="color:#fff;">Over Times</th>
		<th style="color:#fff;">Unit</th> 
		<th style="color:#fff;">Bunch Weight</th> 
		<th style="color:#fff;">KG Brondolan</th>
		<th style="color:#fff;">Output</th>
		<th style="color:#fff;">Rate</th>
		<th style="color:#fff;">Amount Allowded</th>
		<%if instr(v_cField,"remarksallowded") > 0 then%> <th style="color:#fff;">Remarks Allowded</th> <%End if%>
		<th style="color:#fff;">Status</th>
		<th style="color:#fff;">Finger Type</th>
		<th style="color:#fff;">Finger Time</th>
	</tr>
</thead>
<tbody>
	<%
		Set rsCat = Server.CreateObject("ADODB.Recordset")
			rsCat.CursorType = 3
			rsCat.CursorLocation = 3
			rsCat.Open txtSql, connstr
	   if Not rsCat.eof then
			nCountofField=rsCat.fields.count-1
			no=1
		
		Dim v_mandays,v_unit,v_Rate,v_Output,v_othrs, v_bucket
			v_mandays=0
			v_unit=0
			v_Rate=0
			v_Output=0
			v_othrs=0
			v_bucket=0
			v_amount_allowded= 0
			v_kg_brondolan = 0
			v_jumlah_janjang = 0

		While Not rsCat.EOF%>

		<tr>
			<%if len(trim(v_cField)) = 0 then%>
				<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=no%></td>
			<%End If%> 
			<%if instr(v_cField,"fddate") > 0 then%> 
				<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fddate")%></td>
			<%end if%>
			<%if instr(v_cField,"linenokey") > 0 then%>
				<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("linenokey")%></td>
			<%end if%>
			<%if instr(v_cField,"gangcode") > 0 then%>
				<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("gangcode")%></td>
			<%end if%>
			<%if instr(v_cField,"supervision_1") > 0 then%>
				<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("supervision_1")%></td>
			<%end if%>
			<%if instr(v_cField,"supervision_2") > 0 then%>
				<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("supervision_2")%></td>
			<%end if%>
			<%if instr(v_cField,"supervision_3") > 0 then%>
				<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("supervision_3")%></td>
			<%end if%>
			<%if instr(v_cField,"supervision_4") > 0 then%>
				<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("supervision_4")%></td>
			<%end if%>
			<%if instr(v_cField,"sectionname") > 0 then%>
				<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("sectionname")%></td>
			<%end if%>
			<%if instr(v_cField,"employeecode") > 0 then%>
				<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("employeecode")%></td>
			<%end if%>
			<%if instr(v_cField,"employeename") > 0 then%>
				<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("employeename")%></td>	
			<%end if%>
			<%if instr(v_cField,"attendance") > 0 then%>
				<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("attendance")%></td>	
			<%end if%>
			<%if instr(v_cField,"jobcode") > 0 then%>
				<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("jobcode")%></td>	
			<%end if%>
			<%if instr(v_cField,"jobname") > 0 then%>
				<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("jobname")%></td>
			<%end if%>
			<%if instr(v_cField,"locationtype") > 0 then%>
				<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("locationtype")%></td>
			<%end if%>
			<%if instr(v_cField,"locationcode") > 0 then%>
				<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("locationcode")%></td>
			<%end if%>
			<%if instr(v_cField,"chargejob") > 0 then%>
				<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("chargejob")%></td>
			<%end if%>
			<%if instr(v_cField,"chargetype") > 0 then%>
				<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("chargetype")%></td>
			<%end if%>
			<%if instr(v_cField,"chargecode") > 0 then%>
				<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("chargecode")%></td>
			<%end if%>
			<%if instr(v_cField,"allowance_code") > 0 then%>
				<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("allowance_code")%></td>
			<%end if%>
			<%if instr(v_cField,"amount") > 0 then%>
				<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("amount"),2)%></td>
			<%end if%>
			<%if instr(v_cField,"reference") > 0 then%>
				<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("reference")%></td>
			<%end if%>
			<%if instr(v_cField,"spbno") > 0 then%>
				<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("spbno")%></td>
			<%end if%>
			<%if instr(v_cField,"remarks") > 0 then%>
				<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("remarks")%></td>
			<%end if%>
			<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("mandays"), 2)%></td>
			<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("overtime"), 2)%></td>
			<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("unit"), 2)%></td> 
			<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("bunch_weight"), 2)%></td> 
			<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("kg_brondolan"), 2)%></td>
			<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("Output"), 2)%></td>
			<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("rate"), 2)%></td>
			<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("amount_allowded"), 2)%></td>
			<%if instr(v_cField,"remarksallowded") > 0 then%>
				<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("remarksallowded")%></td>
			<%end if%>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("rowstate")%></td>
			
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Fin_type")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Fin_timein") & "-" & rscat("Fin_timeout")%></td>
		</tr>
	  <%
	  v_bucket	= v_bucket + Cdbl(rsCat("bucket"))
	  v_mandays	= v_mandays + Cdbl(rsCat("mandays"))
	  v_othrs	= v_othrs + Cdbl(rsCat("overtime"))
	  v_unit	= v_unit + Cdbl(rsCat("unit"))
	  v_bunch_weight	= v_bunch_weight + Cdbl(rsCat("bunch_weight"))
	  v_jumlah_janjang	= v_jumlah_janjang + Cdbl(rsCat("jumlah_janjang"))
	  v_Output	= v_Output + Cdbl(rsCat("Output"))
	  v_rate	= v_rate + Cdbl(rsCat("rate"))
	  v_kg_brondolan	= v_kg_brondolan + Cdbl(rsCat("kg_brondolan"))
	  v_amount_allowded	= v_amount_allowded + Cdbl(rsCat("amount_allowded"))
	  ' v_bucket	= v_bucket + Cdbl(to_nol(rsCat("bucket")))
	  rsCat.MoveNext
	  'response.Flush()
	  no=no+1
	Wend
	rsCat.close()
	rsCat=null
	END IF
%>

		<tr bgcolor="#F4F4A0">
			<td align="Center" style="<%=cmdstyle%>" colspan= "<%=0+Kolom%>"><font face="Verdana" size="1"><strong>T o t a l</strong></td>
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(v_mandays, 2)%></strong></td>
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(v_othrs, 2)%></strong></td>
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(v_unit, 2)%></strong></td> 
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(v_bunch_weight, 2)%></strong></td> 
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(v_kg_brondolan, 2)%></strong></td>
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(v_Output, 2)%></strong></td>
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(v_rate, 2)%></strong></td>
			<%if instr(v_cField,"remarksallowded") > 0 then%>
				<td></td>
			<%End If%>
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(v_amount_allowded, 2)%></strong></td>
			<td></td>
		</tr>
	</tbody>
</table>
<!--#include file="Bottom_Report.asp"---------->

		</DIV>
	</BODY>
</HTML>	