<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>FBPlants Reporting - Fin Mod</title>
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


</head>
<style>
	.tebal{
		font-weight: bold;
	}
</style>
<body>
<%
dim region,fcbaname,CompanyName,MgrName,KTUName,CropName,FcbaType
 dim regn,cmpy,pyear,tipe,filter,filtercompany
regn     = Request("region")
cmpy     = Request("company")
pyear    = Request("year")
tipe     = Request("type")
if trim(Request("company")) = "-" then 
	cmpy = ""
else
	cmpy = trim(Request("company"))
end if
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"



	if trim(company) = "all" then 
		filtercompany = " WHERE region = '" & region & "'"
	else
		if trim(fcba) = "-" then 
	    	filtercompany = " WHERE fccompanycode ='" & company & "' "
		else 
			filtercompany = " WHERE fccode ='" & fcba & "' "
		end if
	end if
	sqlexec = "Select region, fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit " & filtercompany 
	

	' response.write sqlexec
		set rsCat = connStr.Execute(sqlexec)
		if not rsCat.eof then
			region = rsCat("region")
			fcbaname = rsCat("fcName")
			CompanyName=rsCat("FCCOMPANYNAME")
			MgrName = rsCat("FCMANAGER")
			KTUName = rsCat("FCKTU")
			CropName=rsCat("FCCROP")
			FcbaType=rsCat("fcType")
		end if


   
    ' if trim(company) = "" then 
	' 	filter = " AND hdr.sectioncode in(select SECTIONNAME  from user_section_tab where FCUSERCODE='" & Session("CodeName") & "' ) "
	' else
	'     filter = " AND hdr.sectioncode in(select SECTIONNAME  from user_section_tab where FCUSERCODE='" & Session("CodeName") & "' AND sectionname='" & txtafd & "') "
	' end if

	q = "SELECT a.fccompanycode,a.period_year,a.jobgroupcode,a.jobgroupname,a.job_forfieldstatus,a.qty,a.price,a.amount,nvl(remarks,' ')  " & _ 
	" remarks,a.ownership,a.qtyunit,Round(Nvl(b.qty,0),2) qtyact,Decode(  Nvl(b.qty,0),0,0, Round(Nvl(b.amount,0) /  Nvl(b.qty,0),2)) priceact,  " & _ 
	" Round(Nvl(b.amount,0),2) amountact " & _ 
	" from bud_finmod a " & _
	"	LEFT JOIN " & _
	"	 ( " & _
	"		SELECT " & _
	"		c.fccompanycode,accountingyear period_year,d.job_forbudgetcode jobgroupcode,d.job_forfieldstatus,f.ownership," & _
	"		Sum(Decode(d.budgetqty,'N',0,a.unit)) qty,Sum(amountdebit+amountkredit) amount " & _
	"		FROM bkmjournaldtl a " & _
	"		INNER JOIN bkmjournalhdr b ON (a.docno=b.docno AND a.fcba=b.fcba) " & _
	"		INNER JOIN businessunit c ON (a.fcba=c.fccode) " & _
	"		INNER JOIN jobmaster d ON (a.jobcode=d.fccode) " & _
	"		INNER JOIN field f ON (a.locationcode = f.fccode AND a.fcba=f.fcba ) " & _
	"		WHERE " & _
	"		c.fccompanycode in (select fccompanycode from businessunit where region like '" & regn & "%' AND fccompanycode like '" & cmpy & "%') " & _ 
	"		AND b.accountingyear=" & pyear & "" & _
	"		AND Nvl(d.job_forfieldstatus,'-') <> '-' " & _
	"		AND b.postingflg='1' " & _
	"		GROUP BY  c.fccompanycode,b.accountingyear,d.job_forbudgetcode ,d.job_forfieldstatus,f.ownership " & _
	"	 )b ON (a.fccompanycode = b.fccompanycode AND a.period_year=b.period_year AND  a.jobgroupcode=b.jobgroupcode AND a.job_forfieldstatus=b.job_forfieldstatus AND a.ownership=b.ownership) " & _
	"	WHERE  " & _
	" 	a.fccompanycode in (select fccompanycode from businessunit where region like '" & regn & "%' AND fccompanycode like '" & cmpy & "%') " & _
	"	AND (a.Amount<>0 OR b.amount <> 0 ) AND a.job_forfieldstatus like '" & tipe & "%' AND a.period_year='" & pyear & "'"
						
	Set rsCat = Server.CreateObject("ADODB.Recordset")
	rsCat.CursorType = 3
	rsCat.CursorLocation = 3
	rsCat.Open q, connstr


							
	' if len(trim(txtSelect)) > 0 then
	' txtSelect = "Selected  by : " & txtSelect
	' else
	' txtSelect = "Selected by All Data"
	' end if

' response.Write(q)
'response.Flush()


%>
<script src="/distexcel/tableToExcel.js"></script>
<script type="text/javascript">

	function getExport()
	{
	
		var tahun 	= '<%= pyear %>';
		var pcompanygl = '<%= CompanyName %>';
		var regn    = '<%= regn %>';
		var nm = 'FinMod_' + pcompanygl + '_' + regn + '_'+tahun+".xlsx";  
	//alert(nm);
	 
		TableToExcel.convert(document.getElementById("product-table"), {
            name: nm,
            sheet: {
            name: "FinMod"
            }
          });
		//$('table').tblToExcel();
		
	}
</script>	
<script language="JavaScript" type="text/javascript" >
	document.getElementById('tableview').style.display = "none";
</script>
	<body>
		<div id="wrapper">
         <table border="0" width="100%" id="product-table" class="xthf-site">
         <thead>
         <tr>
         <th colspan="13" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>Region. <%=region%></h4>
					<h5><%=CompanyName%></h5>
				</div>
				<div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h2>FIN MOD</h2>
				</div>
			</div>
					
		</th>
        </tr>
        		<!--<div class="printhide">-->
						<tr>
							<th colspan="13"  style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></th>
						</tr>
						<!--</div>-->
						<tr>
							
						</tr>
						<tr>
							
							<th colspan="6">Description</th>
							<th colspan="3">FinMod</th>
							<th colspan="3">Actual</th>
							<th rowspan="2">Var.%</th>
						</tr>
						<tr>
							<th>PT</th>
							<th>Type</th>
							<th>Ownership</th>
							<th>Code</th>
							<th>Name</th>
							<th>Ha. Gross</th>
							<th>Qty</th>
							<th>Price</th>
							<th>Amount</th>
							<th>Qty</th>
							<th>Price</th>
							<th>Amount</th>
						</tr>
                        </thead>
                        <tbody>
							<%
							dim varian,qtyunit,pricefin,priceact
							dim ttlqtyunit,ttlqty,ttlamount,ttlqtyacct,ttlamountact,ttlprice,ttlpriceact
							ttlqtyunit = 0
							ttlqty	= 0
							ttlamount = 0
							ttlqtyacct = 0
							ttlamountact = 0
							ttlprice = 0
							ttlpriceact = 0
							While Not rsCat.EOF
								 qtyunit 	= rscat("qtyunit")
								 pricefin 	= rscat("price")
								 priceact 	= rscat("priceact")
								if pricefin = "0" then 
								  	varianp = 1
								else
									varianp = cdbl(priceact)/cdbl(pricefin)
								end if
								varian	 = cdbl(varianp)*100
								
							%>
							<tr>
								
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fccompanycode")%> </td>
								<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("job_forfieldstatus")%> </td>
								<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("ownership")%> </td>
								<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("jobgroupcode")%> </td>
								<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("jobgroupname")%> </td>
								<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(rscat("qtyunit"),0)%> </td>
								<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(rscat("qty"),0)%> </td>
								<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(rscat("price"),0)%> </td>
								<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(rscat("amount"),0)%> </td>
								<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(rscat("qtyact"),0)%> </td>
								<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(rscat("priceact"),0)%> </td>
								<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(rscat("amountact"),0)%> </td>
								<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(varian,0)%> </td>
								
							</tr>
							  <%
									
									ttlqty	= cdbl(ttlqty) + cdbl(rscat("qty"))
									ttlprice	= cdbl(ttlprice) + cdbl(rscat("price"))
									ttlamount = cdbl(ttlamount) + cdbl(rscat("amount"))
									ttlqtyacct = cdbl(ttlqtyacct) + cdbl(rscat("qtyact"))
									ttlpriceact = cdbl(ttlpriceact) + cdbl(rscat("priceact"))
									ttlamountact = cdbl(ttlamountact) + cdbl(rscat("amountact"))
							  rsCat.MoveNext
							  'response.Flush()
							  no=no+1
							Wend
							rsCat.close()
							rsCat=null%>
							<tr class="tebal" >
								<td colspan="6" align="right" ><font face="Verdana" size="1">Total </td>
								<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(ttlqty,0)%> </td>
								<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(ttlprice,0)%></td>
								<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(ttlamount,0)%> </td>
								<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(ttlqtyacct,0)%> </td>
								<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(ttlpriceact,0)%></td>
								<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(ttlamountact,0)%> </td>
								<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"></td>
							</tr>
							</form>
                            </tbody>
						</table>
					<%	Response.Write("<script language='javascript'> getExport();</script>")%>
						<!--#include file="Bottom_Report.asp"-->
		</DIV>
	</BODY>
</HTML>	