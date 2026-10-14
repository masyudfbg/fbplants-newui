<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
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
<title>ePlaS Reporting - Contract Work In Progress Final</title>
</head>

<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3

v_txtyear = Request("txtyear")
v_txtmonth = Request("txtmonth")
AGREEMENTCODE = Request("AGREEMENTCODE")
CONTRACTORCODE = Request("CONTRACTORCODE")
LINENO = Request("LINENO")
JOBCODE = Request("JOBCODE")
LOCATIONTYPE = Request("LOCATIONTYPE")
LOCATIONCODE = Request("LOCATIONCODE")
businessunit= trim(request("businessunit"))

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
	
txtSQL = 	 "   SELECT wf.ID,bappno,wf.wipfinaldate," & _
			 "          wf.fcba," & _
			 "          wf.agreementcode," & _
			 "          ca.contractorcode," & _
			 "          cr.fcname contractorname," & _
			 "          wf.agreementlineno," & _
			 "          wf.wipdate," & _
			 "          wf.CHARGEUNIT," & _
			 "          wf.locationtype," & _
			 "          wf.locationcode," & _
			 "          wf.jobcode," & _
			 "          jb.fcname jobname," & _
			 "          jb.uom_unit," & _
			 "          wf.qty_plan plan_qty," & _
			 "          wf.unit_plan plan_unit," & _
			 "          wf.output_plan plan_output," & _
			 "          wf.rate_final rate," & _
			 "          wf.qty_final final_qty," & _
			 "          wf.unit_final final_unit," & _
			 "          wf.output_final final_output," & _
			 "          (wf.rate_final	* wf.qty_final) total" & _
			 "          ,decode(NVL(wf.isprocessing,'FALSE'),'FALSE','PLANNED','APPROVED')  isprocessing " & _ 
			 "    FROM workinprogressfinal_new wf " & _ 
			 "    INNER JOIN jobmaster jb ON (wf.jobcode=jb.fccode) " & _ 
			 "    INNER JOIN contractagreement ca ON (wf.agreementcode=ca.agreementcode and wf.fcba=ca.fcba) " & _
			 "    INNER JOIN contractor cr ON (cr.fccode = ca.contractorcode AND cr.fcba = ca.fcba) " & _ 
			 "   WHERE wf.fcba = '"& session("Fcba") &"' AND to_number(to_char(wipfinaldate,'mm')) = '"& v_txtmonth &"' AND to_number(to_char(wipfinaldate,'yyyy')) = '"& v_txtyear &"'"  

 
			' "   SELECT fcba," & _
			' "          agreementcode," & _
			' "          contractorcode," & _
			' "          contractorname," & _
			' "          agreementlineno," & _
			' "          wipdate," & _
			' "          locationtype," & _
			' "          locationcode," & _
			' "          jobcode," & _
			' "          jobname," & _
			' "          wipfinaldate," & _
			' "          plan_qty," & _
			' "          plan_unit," & _
			' "          plan_output," & _
			' "          rate," & _
			' "          final_qty," & _
			' "          final_unit," & _
			' "          final_output," & _
			' "          (rate * final_qty) total" & _
			' "     FROM (SELECT wf.fcba," & _
			' "                  wf.agreementcode," & _
			' "                  ca.contractorcode," & _
			' "                  ca.contractorname," & _
			' "                  wf.agreementlineno," & _
			' "                  wf.wipdate," & _
			' "                  wf.locationtype," & _
			' "                  wf.locationcode," & _
			' "                  wf.jobcode," & _
			' "                  wf.jobname, " & _
			' "                  wf.wipfinaldate, " & _
			' "                  NVL (wf.qty, 0) plan_qty," & _
			' "                  NVL (wf.unit, 0) plan_unit," & _
			' "                  NVL (wf.output, 0) plan_output," & _
			' "                  NVL (ca.rate, 0) rate," & _
			' "                  NVL (final_qty, 0) final_qty," & _
			' "                  NVL (final_unit, 0) final_unit," & _
			' "                  NVL (final_output, 0) final_output," & _
			' "                  NVL (ca.qty, 0) agreement_qty," & _
			' "                  NVL (final_qty, 0) progress_qty" & _
			' "             FROM (  SELECT wp.fcba," & _
			' "                            wp.agreementcode," & _
			' "                            wp.agreementlineno," & _
			' "                            wpp.wipdate," & _
			' "                            wpp.jobcode," & _
			' "                            wpp.locationcode," & _
			' "                            wpp.locationtype," & _
			' "                            jb.fcname jobname," & _
			' "                            wf.wipfinaldate," & _			
			' "                            NVL (SUM (wp.qty), 0) qty," & _
			' "                            NVL (SUM (wpp.unit), 0) unit," & _
			' "                            NVL (SUM (wpp.output), 0) output," & _
			' "                            NVL (SUM (wf.qty), 0) final_qty," & _
			' "                            NVL (SUM (wf.unit), 0) final_unit," & _
			' "                            NVL (SUM (wf.output), 0) final_output" & _
			' "                       FROM workinprogress wp" & _
			' "                            INNER JOIN workinprogressplan wpp" & _
			' "                               ON (    wp.agreementcode = wpp.agreementcode" & _
			' "                                   AND wp.fcba = wpp.fcba" & _
			' "                                   AND wp.agreementlineno = wpp.agreementlineno" & _
			' "                                   AND wp.agreementfinalno = wpp.agreementfinalno)" & _
			' "                            INNER JOIN workinprogressfinal wf" & _
			' "                               ON (    wp.fcba = wf.fcba" & _
			' "                                   AND wp.agreementcode = wf.agreementcode" & _
			' "                                   AND wp.agreementlineno = wf.agreementlineno" & _
			' "                                   AND wp.agreementfinalno = wf.agreementfinalno)" & _
			' "                            INNER JOIN job jb" & _
			' "                               ON (jb.fccode = wpp.jobcode AND jb.fcba = wpp.fcba)" & _ 
			' "                   GROUP BY wp.fcba," & _
			' "                            wpp.wipdate," & _
			' "                            wp.agreementcode," & _
			' "                            wp.agreementlineno," & _
			' "                            wpp.jobcode," & _
			' "                            wpp.locationcode," & _
			' "                            wpp.locationtype," & _
			' "                            jb.fcname,wf.wipfinaldate) wf" & _
			' "                  INNER JOIN" & _
			' "                  (  SELECT ca.fcba," & _
			' "                            ca.agreementcode," & _
			' "                            ca.contractorcode," & _
			' "                            cr.fcname contractorname," & _
			' "                            cad.lineno," & _
			' "                            NVL (cad.rate, 0) rate," & _
			' "                            NVL (SUM (cad.qty), 0) qty" & _
			' "                       FROM contractagreementdetail cad" & _
			' "                            INNER JOIN contractagreement ca" & _
			' "                               ON (    ca.agreementcode = cad.agreementcode" & _
			' "                                   AND ca.fcba = cad.fcba)" & _
			' "                            INNER JOIN contractor cr" & _
			' "                               ON (    cr.fccode = ca.contractorcode" & _
			' "                                   AND cr.fcba = ca.fcba)" & _
			' "                   GROUP BY ca.fcba," & _
			' "                            ca.agreementcode," & _
			' "                            cad.lineno," & _
			' "                            ca.contractorcode," & _
			' "                            cr.fcname," & _
			' "                            cad.rate) ca" & _
			' "                     ON (    wf.fcba = ca.fcba" & _
			' "                         AND wf.agreementcode = ca.agreementcode" & _
			' "                         AND wf.agreementlineno = ca.lineno))" & _
			
if  len(trim(AGREEMENTCODE)) > 0 then
	txtSQL = txtSQL & " and wf.agreementcode = '" & AGREEMENTCODE & "'"
end if
if  len(trim(CONTRACTORCODE)) > 0 then
	txtSQL = txtSQL & " and UPPER(ca.contractorcode) like UPPER ('%" & CONTRACTORCODE & "%')"
end if
if  len(trim(LINENO)) > 0 then
	txtSQL = txtSQL & " and wf.agreementlineno = '" & LINENO & "'"
end if
if  len(trim(JOBCODE)) > 0 then
	txtSQL = txtSQL & " and UPPER (wf.jobcode) like UPPER ('%" & JOBCODE & "%')"
end if
if  len(trim(LOCATIONTYPE)) > 0 then
	txtSQL = txtSQL & " and targettype = '" & LOCATIONTYPE & "'"
end if 
if  len(trim(LOCATIONCODE)) > 0 then
	txtSQL = txtSQL & " and targetcode = '" & LOCATIONCODE & "'"
END IF	

txtsql = txtsql & " "
					
' txtsql = txtsql & " ORDER BY contractorcode, bappno, agreementcode, LPAD (agreementlineno, 3, 0)" 
					
'response.Write txtsql
' response.Flush()
%>
<body>
	<div id="wrapper">
        <!--#include file="print_excel.asp"-->
        <table border="0" width="100%" id="product-table" class="xthf-site">
			<thead>
			<tr>
				<th colspan="22" style=" background-color:#FFFFFF; color:#000000; border:none">
					<div id="header-wrapper">
						<div id="head_judul" style="float:left; text-align:left;">
							<h4><%=Session("COMPANY")%></h4>
							<h5><%=Session("BA")%></h5>
						</div>
						<div id="subheader" class="centered" style="width:50%; margin:auto;">
							<h2>WORK IN PROGRESS FINAL</h2>
						</div>
					</div>
				</th>
			</tr>
			<tr>
				<td colspan="22" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></td>
			</tr>
<%	
	Set rsCat = connStr.Execute(txtSQL)
	if Not rsCat.eof then
	   nCountofField = rsCat.fields.count-1
	end if
	NO = 1
%>
			<tr>   
			    <th rowspan="2"></th>
				<th rowspan="2">No.BAPP</th>
				<th rowspan="2">Status</th>
				<th rowspan="2">Date.BAPP</th>	   
				<th rowspan="2">Agreement Code</th>
				<th rowspan="2">Contractor Code</th>
				<th rowspan="2">Contractor Name</th>
				<th rowspan="2">Line No</th>
				<th rowspan="2">WIP Date</th>
				<th rowspan="2">Charge Unit</th>	   
				<th rowspan="2">Target Type</th>
				<th rowspan="2">Target Code</th>
				<th rowspan="2">Job Code</th>
				<th rowspan="2">Job Name</th>
				<th rowspan="2">Job Uom</th>
				<th colspan="3">Plan</th>
				<th colspan="5">Final</th> 
			</tr>
			<tr>
				<th>Qty</th>
				<th>Unit</th>
				<th>Output(Ha)</th>
				<th>Rate</th>
				<th>Qty</th>
				<th>Unit</th>
				<th>Output(Ha)</th>
				<th>Total</th> 
			</tr>
			</thead>
			<tbody>
	<%
	dim v_plan_qty : v_plan_qty = 0
	dim v_plan_unit : v_plan_unit = 0
	dim v_plan_output : v_plan_output = 0
	dim v_rate : v_rate = 0
	dim v_final_qty : v_final_qty = 0
	dim v_final_unit : v_final_unit = 0
	dim v_final_output : v_final_output = 0
	dim v_total : v_total = 0
	dim v_agreement_qty : v_agreement_qty = 0
	dim v_progress_qty : v_progress_qty = 0
	dim v_outstanding_qty : v_outstanding_qty = 0

	
	While Not rsCat.EOF%>
			<tr bgcolor="<%=cColor%>">	
				<td valign="middle" width="10">		
							<a href="wipfinal.Asp?txtbappno=<%=rscat("BAPPNO")%>&txtagreementcode=<%=rscat("AGREEMENTCODE")%>&cmdView=True">
								<img src="IMAGES/edit.gif" alt="Edit Data : <%=rscat(1)%>" border="0">
							</a>
				</td>			
				<td align="left" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("BAPPNO")%></td>
				<td align="left" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("isprocessing")%></td>
				<td align="left" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("WIPFINALDATE")%></td>
				<td align="left" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("AGREEMENTCODE")%></td>
				<td align="left" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("CONTRACTORCODE")%></td>
				<td align="left" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("CONTRACTORNAME")%></td>
				<td align="left" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("AGREEMENTLINENO")%></td>
				<td align="left" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("WIPDATE")%></td>
				<td align="left" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("CHARGEUNIT")%></td>
				<td align="left" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("locationtype")%></td>
				<td align="left" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("locationcode")%></td>
				<td align="left" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("JOBCODE")%></td>
				<td align="left" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("JOBNAME")%></td>
				<td align="left" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("uom_unit")%></td>
				<td align="Right" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rsCat("PLAN_QTY"))), 2)%></td>
				<td align="Right" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rsCat("PLAN_UNIT"))), 2)%></td>
				<td align="Right" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rsCat("PLAN_OUTPUT"))), 2)%></td>
				<td align="Right" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rsCat("RATE"))), 2)%></td>
				<td align="Right" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rsCat("FINAL_QTY"))), 2)%></td>
				<td align="Right" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rsCat("FINAL_UNIT"))), 2)%></td>
				<td align="Right" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rsCat("FINAL_OUTPUT"))), 2)%></td>
				<td align="Right" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rsCat("TOTAL"))), 2)%></td> 
			</tr>
<%
		v_plan_qty = v_plan_qty + cdbl(rscat("PLAN_QTY"))
		v_plan_unit = v_plan_unit + cdbl(rscat("PLAN_UNIT"))
		v_plan_output = v_plan_output + cdbl(rscat("PLAN_OUTPUT"))
		v_rate = v_rate + cdbl(rscat("RATE"))
		v_final_qty = v_final_qty + cdbl(rscat("FINAL_QTY"))
		v_final_unit = v_final_unit + cdbl(rscat("FINAL_UNIT"))
		v_final_output = v_final_output + cdbl(rscat("FINAL_OUTPUT"))
		v_total = v_total + cdbl(rscat("TOTAL"))
		' v_agreement_qty = v_agreement_qty + cdbl(rscat("AGREEMENT_QTY"))
		' v_progress_qty = v_progress_qty + cdbl(rscat("PROGRESS_QTY"))
		' v_outstanding_qty = v_outstanding_qty + cdbl(rscat("OUTSTANDING_QTY"))

		rsCat.MoveNext
		no = no + 1
	Wend
	rsCat.close()
	rsCat=null
%>
	<input  type="hidden" value="<%=txtTable%>" name="txtTable">
	<input  type="hidden" value="<%=txtModule%>" name="txtModule">
	</form>
		<tr bgcolor="#F4F4A0">
			<td align="Center" style="<%=cmdstyleLeft%>" colspan="15"> <font face="Verdana" size="1"><strong>T o t a l</strong></td>			
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(v_plan_qty, 2)%></strong></td>
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(v_plan_unit, 2)%></strong></td>
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(v_plan_output, 2)%></strong></td>
			<td align="Right" style="<%=cmdstyle%>"></td>
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(v_final_qty, 2)%></strong></td>
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(v_final_unit, 2)%></strong></td>
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(v_final_output, 2)%></strong></td>
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(v_total, 2)%></strong></td> 		
		</tr>
    </tbody>
</table>
<!--#include file="Bottom_Report.asp"-->
</DIV>
</BODY>
</HTML>