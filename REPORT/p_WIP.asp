<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - Contract Work In Progress</title>
</head>
<body>
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

    v_Contractor_From=trim(request("DATEACTIVITY_From"))
	v_Contractor_To=trim(request("DATEACTIVITY_To"))
	v_JOBCODE=trim(request("JOB"))
	v_LOCATIONTYPE=trim(request("LOCATIONTYPE"))
	v_LocationCode=trim(request("LocationCode"))
	
	txtSQL = "select wip.AGREEMENTCODE,to_char(wip.WIPDATE,'mm/dd/yyyy') as WIPDATE, wip.AGREEMENTLINENO, ag.LOCATIONTYPE,ag.LOCATIONCODE,ag.job,job.FCNAME,sum(wip.QTY) as QTY " & _
     " from WORKINPROGRESS wip ,AGREEMENTDETAIL AG,JOB  where (wip.AGREEMENTCODE=AG.AGREEMENTCODE and wip.AGREEMENTLINENO=ag.LINENO and wip.fcba=ag.fcba)" & _
	 " and (ag.job = job.FCCODE and ag.fcba = job.fcba) " & _
	 " and wip.fcBA='" & session("fcBA") & "' "
	if  len(trim(v_Contractor_From)) > 0 then
		txtSQL = txtSQL & " and wip.WIPDATE >=to_Date('" & v_Contractor_From & "','MM/DD/YYYY')"
	end if
	if  len(trim(v_Contractor_To)) > 0 then
		txtSQL =txtSQL &  " and wip.WIPDATE <=to_Date('" & v_Contractor_To & "','MM/DD/YYYY')"
	end if
	if  len(trim(v_JOBCODE)) > 0 then
		txtSQL = txtSQL & " and aG.JOB='" & v_JOBCODE & "'"
	end if
	if  len(trim(v_LOCATIONTYPE)) > 0 then
		txtSQL =txtSQL &  " and AG.LOCATIONTYPE='" & v_LOCATIONTYPE & "'"
	end if 
	if  len(trim(v_LocationCode)) > 0 then
		txtSQL =txtSQL &  " and AG.LOCATIONCODE='" & v_LocationCode & "'"
	END IF	

    txtsql=(txtsql) & " group by wip.AGREEMENTCODE,to_char(wip.WIPDATE,'mm/dd/yyyy'), wip.AGREEMENTLINENO, ag.LOCATIONTYPE,ag.LOCATIONCODE,ag.job,job.FCNAME"
    'response.Write(txtsql)
	'response.Flush()
%>
 <body>
		<div id="wrapper">
			<div id="header-wrapper">
				<div id="header">
					<h4>PT. <%=Session("COMPANY")%></h4>
					<h5><%=Session("BA")%></h5>
				</div>
				<div id="subheader" class="centered">
					<h2>WORK IN PROGRESS</h2>
				</div>
			</div>
					<table border="0" width="100%" id="product-table" class="headertable">
						<div class="printhide">
						<tr>
							<td colspan="8"><%=txtSelect%></td>
						</tr>
						</div>
	<%	
	Set rsCat = connStr.Execute(txtSQL)
	if Not rsCat.eof then
	   nCountofField=rsCat.fields.count-1
	end if
	NO=1%>
	</tr>
   <tr>   
	   <th>Code</th>
	   <th>WIP Date</th>
	   <th>Line</th>
	   <th>TT</th>
	   <th>TC</th>
	   <th>Job Code</th>
	   <th>Job Desc.</th>
	   <th>Qty</th>
   </tr>
	<%
	Dim v_Quantity
	v_Quantity=0
	While Not rsCat.EOF%>
		  <tr bgcolor="<%=cColor%>">
		   
          <td align="left" title="WS.AGREEMENTCODE" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("AGREEMENTCODE")%></td>
	      <td align="left" title="WS.LINENO" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("WIPDATE")%></td>
	      <td align="left" title="WS.JOB" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("AGREEMENTLINENO")%></td>
	      <td align="left" title="WS.JOB" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("LOCATIONTYPE")%></td>
	      <td align="left" title="WS.JOB" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("LOCATIONCODE")%></td>
	      <td align="left" title="WS.JOB" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("job")%></td>
	      <td align="left" title="WS.JOB" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("FCNAME")%></td>
	      <td align="Right" title="WS.QTY" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsCat("QTY"),2)%></td>
		  </tr>
	  <%
	  v_Quantity=v_Quantity+trim(cdbl(rsCat("QTY")))
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
	<tr bgcolor="#F4F4A0">
	<td align="Center" style="<%=cmdstyleLeft%>" colspan="7"> <font face="Verdana" size="1">T o t a l</td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_Quantity,2)%> </td>
	</tr>
</table>
<!--#include file="Bottom_Report.asp"-->
</DIV>
</BODY>
</HTML>