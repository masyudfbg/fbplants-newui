<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
	Dim v_txtYear, v_txtMonth
	Dim ScriptName
	Dim rsGrnInv, strGrnInv
	Dim v_txtPOCode
	Dim NO, cClass
	
	ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

	v_txtFcba = Request("txtFcba")
	v_txtMonth = Request("txtMonth")
	v_txtYear = Request("txtYear")
	v_txtsectionname = Request("txtsectionname")
	v_txtempcode = Request("txtempcode")

	txttdate = txtMonth & "" & txtYear
	If "" & txtYear = "" Then txtYear = Year(Session("v_Costbook"))
	If "" & txtMonth = "" Then txtMonth = Month(Session("v_Costbook"))
	If "" & txttdate = "" Then txttdate = textDate(date)
	varDefaultBa = Session("fcBA")

	v_txtdatefrom = Request("txtdatefrom")
	v_txtdateuntil = Request("txtdateuntil")
	If (v_txtdatefrom = "" Or Len(Trim(v_txtdatefrom)) = 0) Then v_txtdatefrom = textdate(date)
	If (v_txtdateuntil = "" Or Len(Trim(v_txtdateuntil)) = 0) Then v_txtdateuntil = textdate(date)

	If (v_txtFcba = "" Or Len(Trim(v_txtFcba)) = 0) Then
		v_txtFcba = varDefaultBa
	End If	
	nSpanTotal = 0
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
 
		strGrnInv	=	" SELECT emp.fcba," & _
						"        emp.FCCODE," & _
						"        EMPLOYEE_API.GET_EMPLOYEENAME (emp.fcba, emp.FCCODE) employeename," & _
						"        att_gad.FDDATE," & _
						"        att_gad.LOCATIONCODE," & _
						"        att_gad.BUCKET," & _
						"        ROUND (MWN.BBR, 2) BBR," & _
						"        ROUND ( (att_gad.bucket * MWN.BBR * mhr.rate_brondolan), 2) AS amount" & _
						"   FROM attendance_gad att_gad" & _
						"        INNER JOIN Employee Emp" & _
						"           ON     att_gad.EMPLOYEECODE = emp.FCCODE" & _
						"              AND att_gad.FCBA = emp.FCBA" & _
						"              AND emp.DATETERMINATE IS NULL" & _
						"        INNER JOIN HARVESTING_MILL_WEIGHT_NETTO MWN" & _
						"           ON     (    att_gad.fcba = MWN.FCBA_ESTATE" & _
						"                   AND att_gad.LOCATIONCODE = MWN.FIELDCODE)" & _
						"              AND TO_CHAR (att_gad.FDDATE, 'MM/DD/YYYY') =" & _
						"                     TO_CHAR (MWN.HARVESTDATE, 'MM/DD/YYYY')" & _
						"        INNER JOIN FIELD ff" & _
						"           ON     att_gad.LOCATIONCODE = ff.FCCODE" & _
						"              AND att_gad.fcba = ff.fcba" & _
						"              AND ff.ACTIVATION = 'Y'" & _
						"        INNER JOIN (SELECT DISTINCT fcba," & _
						"                                    division," & _
						"                                    planting_year," & _
						"                                    rate_brondolan" & _
						"                      FROM MULTILEVEL_HARVESTING_RATE" & _
						"                     WHERE fcba = '"& v_txtfcba &"') mhr" & _
						"           ON (    att_gad.fcba = mhr.fcba" & _
						"               AND ff.DIVISION = mhr.DIVISION" & _
						"               AND TO_CHAR (ff.PLANTINGDATE, 'yyyy') = mhr.PLANTING_YEAR)" & _
						"  WHERE     att_gad.Attendance IN ('KJ', 'WH', 'WS')" & _
						"        AND att_gad.Jobcode IN ('112030101', '505030101')" & _
						"        AND att_gad.LOCATIONTYPE = 'FF'" & _
						"        AND att_gad.MANDAYS > 0" & _
						"        AND att_gad.BUCKET > 0" & _
						"        AND att_gad.rate = 0" & _ 
						"        AND att_gad.fcBA = '"& v_txtfcba &"'"  
 
		if  len(trim(v_txtdatefrom)) > 0 then
			strGrnInv = strGrnInv & " AND att_gad.FDDATE >= TO_DATE('" & v_txtdatefrom & "', 'mm/dd/yyyy')"
			txtSelect = " From " & FormatDateTime(v_txtdatefrom,1)
		end if
		if  len(trim(v_txtdateuntil)) > 0 then
			strGrnInv = strGrnInv & " AND att_gad.FDDATE <= TO_DATE('" & v_txtdateuntil & "', 'mm/dd/yyyy')"
			txtSelect = txtSelect & " To " & FormatDateTime(v_txtdateuntil,1)
		end if 
		if  len(trim(v_txtempcode)) > 0 then
		    strGrnInv	=	strGrnInv & " AND operatorcode = '"& v_txtempcode &"'" 
		    txtSelect = txtSelect & " Employee Code " & v_txtempcode
		end if 
									
		strGrnInv	=	strGrnInv & " ORDER BY fcba," & _
									"          FCCODE," & _ 
									"          FDDATE," & _
									" 		   LOCATIONCODE"

		' response.write strGrnInv	
		Set rsGrnInv = Server.CreateObject("ADODB.Recordset")
		rsGrnInv.Open strGrnInv, connstr, adOpenStatic 
 
%>


<html>

<head>
<title>ePlaS Reporting - Premi Brondolan </title>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>
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
 <script>
$(function() {
$( ".datepicker" ).datepicker();
});
</script>
<!--<script language="JavaScript" src="/js/date-picker.js"></script> 
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>-->

<style type="text/css">
@media print
{
.printhide{display:none;}
}
</style>
</head>

<body>
	<div id="wrapper">
	<!--#include file="../print_excel.asp"-->
		<div id="header-wrapper">
		
			<div class="printhide">
				<!--#include file="menutop.asp"-->
			</div>
            	<form action="<%=ScriptName%>" method="post" name="frmEntry">
			<div class="printhide">
				<table id="form-table" class="centered-margin">
					<tr>
						<td>
							<font face="verdana" size="1">KEBUN</font>
						</td>
						<td>
							<select name="txtFcba" id="vFcba" onchange="javascript:filtering()">
							<%
							Dim rsFcba
							Set rsFcba = Server.CreateObject("ADODB.Recordset")
							rsFcba.CursorLocation = adUseClient
							strsql =   " select fccode fcba from businessunit " & _
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
							<font face="verdana" size="1">FROM</font>
						</td>
						<td width="4">
							<input class="datepicker" type="text" maxlength="10" name="txtdatefrom" size="10" value="<%=v_txtdatefrom%>">
						</td>
						<td>
							<font face="verdana" size="1">TO</font>
						</td>
						<td width="2">
							<input class="datepicker" type="text" maxlength="10" name="txtdateuntil" size="10" value="<%=v_txtdateuntil%>">
						</td> 
						<td>
							<font face="verdana" size="1">EMPLOYEE</font>
						</td>
						<td>
							<select name="txtempcode" >
								<option value="">-- ALL --</option>
									<% 
									strsectionname	=	"select fccode, fcname from employee where fcba = '"& v_txtFcba &"' and dateterminate is null order by fccode"
									
									Set rssectionname = Server.CreateObject("ADODB.Recordset")
									rssectionname.Open strsectionname, connstr, adOpenStatic 
									
									if not rssectionname.eof then
									Do While Not rssectionname.eof
									if rssectionname("fccode") = v_txtempcode then%>
											<option value="<%=rssectionname("fccode")%>" Selected><%=rssectionname("fccode")%> : <%=rssectionname("fcname")%></option> 
										 <%Else%>
											<option value="<%=rssectionname("fccode")%>"><%=rssectionname("fccode")%> : <%=rssectionname("fcname")%></option>  
										<%End if
										rssectionname.movenext
									loop
									end if
									rssectionname.Close
									set rssectionname = Nothing
									%>
							</select>
						</td> 
						<td>
							<input type="submit" value="View" />
						</td>
					</tr>
				</table>
			</div>
			</form>
			
			</div>
<table border="0" width="100%" id="product-table" class="xthf-site">
	<thead>
		<tr>
    		<th colspan="7" style="background-color:#FFFFFF; border:none;">  
				<div id="header-wrapper">
					<div id="head_judul" style="float:left; text-align:left;">
						<h4>PT. <%=CompanyName%></h4>
						<h5><%=fcbaname%></h5>
					</div>
					<div id="subheader" class="centered" tyle="width:50%; margin:auto;">
						<h2>LAPORAN PREMI BRONDOLAN</h2>
		                <br/>
						<h4>Selected By :<%=txtSelect%></h4>
					</div> 
				</div>
			</th>
		</tr>
		<tr>
			<th style="background-color:#00A300;color:#fff;">EMPLOYEE CODE</th> 
			<th style="background-color:#00A300;color:#fff;">EMPLOYEE NAME</th>
			<th style="background-color:#00A300;color:#fff;">DATE</th>
			<th style="background-color:#00A300;color:#fff;">BLOCK</th>
			<th style="background-color:#00A300;color:#fff;">BUCKET</th>
			<th style="background-color:#00A300;color:#fff;">BBR</th>
			<th style="background-color:#00A300;color:#fff;">AMOUNT</th>
		</tr>   
    </thead>
<tbody>
<%if rsGrnInv.BOF and rsGrnInv.EOF then%>
		<tr>
			<td colspan="7" align="left">NO DATA AVAILABLE</td>
		</tr>
<%end if
		NO = 1
		rows = 1
		while not rsGrnInv.eof
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
		<tr class="<%=cClass%>">
			<% if rows = 1 then
					sqlrow = "  select count(fccode) rowsd  from ( " & strGrnInv & " ) where fccode = '" & rsGrnInv("fccode") & "'"
					' response.write sqlrow
					Set rsrow = Server.CreateObject("ADODB.Recordset")
					rsrow.Open sqlrow, connstr, adOpenStatic
					if not rsrow.EOF then
						temprows = CDbl(rsrow("rowsd"))
					else
						temprows = 0
					end if
				end if 
			%>
			<%if rows = 1 then%>
				<td align="center" Rowspan="<%=0+temprows%>"><%=rsGrnInv("fccode")%></td>
				<td align="center" Rowspan="<%=0+temprows%>"><%=rsGrnInv("employeename")%></td>
			<%end if%>
			<td align="center"><%=rsGrnInv("fddate")%></td>
			<td align="center"><%=rsGrnInv("LOCATIONCODE")%></td>
			<td align="right"><%=FormatNumber(cdbl(to_nol(rsGrnInv("bucket"))),0)%></td>
			<td align="right"><%=FormatNumber(cdbl(to_nol(rsGrnInv("bbr"))),2)%></td>
			<td align="right"><%=FormatNumber(cdbl(to_nol(rsGrnInv("amount"))),2)%></td>  
		</tr>
<%	if rows = temprows then
			nbucket = nbucket + cdbl(to_nol(rsGrnInv("bucket")))
			nbbr = nbbr + cdbl(to_nol(rsGrnInv("bbr")))	
			namount = namount + cdbl(to_nol(rsGrnInv("amount"))) 
			rows = 1
			
	%>
		<tr bgcolor="#F4F4A0"> 
			<td align="center" colspan="4"><font style="font-weight: bold;">TOTAL</font></td>				
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(nbucket)),0)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(nbbr)),2)%></font></td>  
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(namount)),2)%></font></td>  
		</tr>
		<tr>
			<td align="center" colspan="<%= 7+ nSpanTotal %>" ></td>
		</tr>
		<%
			nbucket = cdbl("0")
			nbbr = cdbl("0")
			namount = cdbl("0") 
		else
		' cari sub total
			nbucket = nbucket + cdbl(to_nol(rsGrnInv("bucket")))
			nbbr = nbbr + cdbl(to_nol(rsGrnInv("bbr")))	
			namount = namount + cdbl(to_nol(rsGrnInv("amount")))
			rows = rows + 1
		End if
							  
			' cari grand total 
				gt_bucket = cdbl(gt_bucket) + cdbl(rsGrnInv("bucket"))
				gt_bbr = cdbl(gt_bbr) + cdbl(rsGrnInv("bbr"))
				gt_amount = cdbl(gt_amount) + cdbl(rsGrnInv("amount")) 

			rsGrnInv.MoveNext			
			Wend
		rsGrnInv.Close
		set rsGrnInv = Nothing
	%> 
		<tr bgcolor="#FFFF00">	 
			<td align="center" colspan="4"><b>GRAND TOTAL</b></td>
			<td align="right"><b><%=formatnumber(gt_bucket,0)%></td>
			<td align="right"><b><%=formatnumber(gt_bbr,2)%></td>
			<td align="right"><b><%=formatnumber(gt_amount,2)%></td> 
		</tr>
	</tbody>
</table>