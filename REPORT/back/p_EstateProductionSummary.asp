<!--#include file="../connDb1.Inc"-->
<!--#include file="../txtFunction.inc"-->
<%

	'---- CursorTypeEnum Values ----
	Const adOpenForwardOnly = 0
	Const adOpenKeyset = 1
	Const adOpenDynamic = 2
	Const adOpenStatic = 3
	
	Dim rsEPS, strEPS

	v_txtyear			=	Cint(trim(request("txtyear")))
	v_txtmonth			=	Cint(trim(request("txtmonth")))
	v_txttodaydate		=	request("txttodaydate")
	
strEPS	=	" select" & _
			" fcba," & _
			" division," & _
			" fieldcode," & _
			" fieldname," & _
			" plantingyear," & _
			" sum(hi) hi," & _
			" sum(shi) shi," & _
			" sum(bi) bi," & _
			" sum(sbi) sbi," & _
			" sum(tonase_hi) tonase_hi," & _
			" sum(tonase_shi) tonase_shi," & _
			" sum(tonase_bi) tonase_bi," & _
			" sum(tonase_sbi) tonase_sbi," & _
			" sum(pom_hi) pom_hi," & _
			" sum(pom_shi) pom_shi," & _
			" sum(pom_bi) pom_bi," & _
			" sum(pom_sbi) pom_sbi," & _
			" sum(pom_tonase_hi) pom_tonase_hi," & _
			" sum(pom_tonase_shi) pom_tonase_shi," & _
			" sum(pom_tonase_bi) pom_tonase_bi," & _
			" sum(pom_tonase_sbi) pom_tonase_sbi," & _
			" sum(restan) restan" & _
			" from (" & _
			" SELECT fcba," & _
			" division," & _
			" fieldcode," & _
			" fieldname," & _
			" plantingyear," & _
			" totaldailyproduction hi," & _
			" 0 shi," & _
			" 0 bi," & _
			" 0 sbi," & _
			" totaldailyprodtonase tonase_hi," & _
			" 0 tonase_shi," & _
			" 0 tonase_bi," & _
			" 0 tonase_sbi," & _
			" 0 pom_hi," & _
			" 0 pom_shi," & _
			" 0 pom_bi," & _
			" 0 pom_sbi," & _
			" 0 pom_tonase_hi," & _
			" 0 pom_tonase_shi," & _
			" 0 pom_tonase_bi," & _
			" 0 pom_tonase_sbi," & _
			" 0 Restan" & _
			" FROM view_dailyharvesting_prod" & _
			" WHERE fcba = '" & Session("fcBA")& "'" & _
			" AND accountingyear = " & v_txtyear & _
			" AND accountingperiod = " & v_txtmonth & _
			" AND fddate = TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY')" & _
			" UNION" & _
			" SELECT fcba," & _
			" division," & _
			" fieldcode," & _
			" fieldname," & _
			" plantingyear," & _
			" 0 hi," & _
			" sum(totaldailyproduction) shi," & _
			" 0 bi," & _
			" 0 sbi," & _
			" 0 tonase_hi," & _
			" sum(totaldailyprodtonase) tonase_shi," & _
			" 0 tonase_bi," & _
			" 0 tonase_sbi," & _
			" 0 pom_hi," & _
			" 0 pom_shi," & _
			" 0 pom_bi," & _
			" 0 pom_sbi," & _
			" 0 pom_tonase_hi," & _
			" 0 pom_tonase_shi," & _
			" 0 pom_tonase_bi," & _
			" 0 pom_tonase_sbi," & _
			" 0 Restan" & _
			" FROM view_dailyharvesting_prod" & _
			" WHERE fcba = '" & Session("fcBA")& "'" & _
			" AND accountingyear = " & v_txtyear & _
			" AND accountingperiod = " & v_txtmonth & _
			" AND fddate <= TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY')" & _
			" group by" & _
			" fcba," & _
			" accountingyear," & _
			" accountingperiod," & _
			" division," & _
			" fieldcode," & _
			" fieldname," & _
			" plantingyear" & _
			" UNION" & _
			" SELECT fcba," & _
			" division," & _
			" fieldcode," & _
			" fieldname," & _
			" plantingyear," & _
			" 0 hi," & _
			" 0 shi," & _
			" 0 bi," & _
			" sum(totaldailyproduction) sbi," & _
			" 0 tonase_hi," & _
			" 0 tonase_shi," & _
			" 0 tonase_bi," & _
			" sum(totaldailyprodtonase) tonase_sbi," & _
			" 0 pom_hi," & _
			" 0 pom_shi," & _
			" 0 pom_bi," & _
			" 0 pom_sbi," & _
			" 0 pom_tonase_hi," & _
			" 0 pom_tonase_shi," & _
			" 0 pom_tonase_bi," & _
			" 0 pom_tonase_sbi," & _
			" 0 Restan" & _
			" FROM view_dailyharvesting_prod" & _
			" WHERE fcba = '" & Session("fcBA")& "'" & _
			" AND accountingyear = " & v_txtyear & _
			" AND fddate <= TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY')" & _
			" group by" & _
			" fcba," & _
			" division," & _
			" fieldcode," & _
			" fieldname," & _
			" plantingyear" & _
			" UNION" & _
			" SELECT fcba," & _
			" division," & _
			" fieldcode," & _
			" fieldname," & _
			" plantingyear," & _
			" 0 hi," & _
			" 0 shi," & _
			" 0 bi," & _
			" 0 sbi," & _
			" 0 tonase_hi," & _
			" 0 tonase_shi," & _
			" 0 tonase_bi," & _
			" 0 tonase_sbi," & _
			" totaldailyreceive pom_hi," & _
			" 0 pom_shi," & _
			" 0 pom_bi," & _
			" 0 pom_sbi," & _
			" totaldailyreceivetonase1 pom_tonase_hi," & _
			" 0 pom_tonase_shi," & _
			" 0 pom_tonase_bi," & _
			" 0 pom_tonase_sbi," & _
			" 0 Restan" & _
			" FROM view_dailyharvesting_mill" & _
			" WHERE fcba = '" & Session("fcBA")& "'" & _
			" AND accountingyear = " & v_txtyear & _
			" AND accountingperiod = " & v_txtmonth & _
			" AND receptiondate = TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY')" & _
			" UNION" & _
			" SELECT fcba," & _
			" division," & _
			" fieldcode," & _
			" fieldname," & _
			" plantingyear," & _
			" 0 hi," & _
			" 0 shi," & _
			" 0 bi," & _
			" 0 sbi," & _
			" 0 tonase_hi," & _
			" 0 tonase_shi," & _
			" 0 tonase_bi," & _
			" 0 tonase_sbi," & _
			" 0 pom_hi," & _
			" sum(totaldailyreceive) pom_shi," & _
			" 0 pom_bi," & _
			" 0 pom_sbi," & _
			" 0 pom_tonase_hi," & _
			" sum(totaldailyreceivetonase1) pom_tonase_shi," & _
			" 0 pom_tonase_bi," & _
			" 0 pom_tonase_sbi," & _
			" 0 Restan" & _
			" FROM view_dailyharvesting_mill" & _
			" WHERE fcba = '" & Session("fcBA")& "'" & _
			" AND accountingyear = " & v_txtyear & _
			" AND accountingperiod = " & v_txtmonth & _
			" AND receptiondate <= TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY')" & _
			" group by" & _
			" fcba," & _
			" accountingyear," & _
			" accountingperiod," & _
			" division," & _
			" fieldcode," & _
			" fieldname," & _
			" plantingyear" & _
			" UNION" & _
			" SELECT fcba," & _
			" division," & _
			" fieldcode," & _
			" fieldname," & _
			" plantingyear," & _
			" 0 hi," & _
			" 0 shi," & _
			" 0 bi," & _
			" 0 sbi," & _
			" 0 tonase_hi," & _
			" 0 tonase_shi," & _
			" 0 tonase_bi," & _
			" 0 tonase_sbi," & _
			" 0 pom_hi," & _
			" 0 pom_shi," & _
			" 0 pom_bi," & _
			" sum(totaldailyreceive) pom_sbi," & _
			" 0 pom_tonase_hi," & _
			" 0 pom_tonase_shi," & _
			" 0 pom_tonase_bi," & _
			" sum(totaldailyreceivetonase1) pom_tonase_sbi," & _
			" 0 Restan" & _
			" FROM view_dailyharvesting_mill" & _
			" WHERE fcba = '" & Session("fcBA")& "'" & _
			" AND accountingyear = " & v_txtyear & _
			" AND receptiondate <= TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY')" & _
			" group by" & _
			" fcba," & _
			" division," & _
			" fieldcode," & _
			" fieldname," & _
			" plantingyear" & _
			" UNION" & _
			" SELECT fcba," & _
			" division," & _
			" fieldcode," & _
			" fieldname," & _
			" plantingyear," & _
			" 0 hi," & _
			" 0 shi," & _
			" 0 bi," & _
			" 0 sbi," & _
			" 0 tonase_hi," & _
			" 0 tonase_shi," & _
			" 0 tonase_bi," & _
			" 0 tonase_sbi," & _
			" 0 pom_hi," & _
			" 0 pom_shi," & _
			" 0 pom_bi," & _
			" 0 pom_sbi," & _
			" 0 pom_tonase_hi," & _
			" 0 pom_tonase_shi," & _
			" 0 pom_tonase_bi," & _
			" 0 pom_tonase_sbi," & _
			" (SUM (totaldailyproduction) - SUM (totaldailyreceive)) + ( SUM (prevtotaldailyproduction) - SUM (prevtotaldailyreceive)) Restan" & _
			" FROM (SELECT fcba, accountingyear, accountingperiod, division, fieldcode," & _
			" fieldname, plantingyear, fddate dateprod," & _
			" totaldailyproduction, 0 totaldailyreceive, 0 prevtotaldailyproduction, 0 prevtotaldailyreceive" & _
			" FROM view_dailyharvesting_prod" & _
			" WHERE fcba = '" & Session("fcBA")& "'" & _
			" AND accountingyear = " & v_txtyear & _
			" AND accountingperiod = " & v_txtmonth & _
			" AND fddate = TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY')" & _
			" UNION" & _
			" SELECT fcba, accountingyear, accountingperiod, division, fieldcode," & _
			" fieldname, plantingyear, receptiondate dateprod," & _
			" 0 totaldailyproduction, totaldailyreceive, 0 prevtotaldailyproduction, 0 prevtotaldailyreceive" & _
			" FROM view_dailyharvesting_mill" & _
			" WHERE fcba = '" & Session("fcBA")& "'" & _
			" AND accountingyear = " & v_txtyear & _
			" AND accountingperiod = " & v_txtmonth & _
			" AND receptiondate = TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY')" & _
			" UNION" & _
			" SELECT fcba, accountingyear, accountingperiod, division, fieldcode," & _
			" fieldname, plantingyear, fddate dateprod," & _
			" 0 totaldailyproduction, 0 totaldailyreceive, totaldailyproduction Prevtotaldailyproduction, 0 Prevtotaldailyreceive" & _
			" FROM view_dailyharvesting_prod" & _
			" WHERE fcba = '" & Session("fcBA")& "'" & _
			" AND accountingyear = " & v_txtyear & _
			" AND accountingperiod = " & v_txtmonth & _
			" AND fddate = TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY') - 1" & _
			" UNION" & _
			" SELECT fcba, accountingyear, accountingperiod, division, fieldcode," & _
			" fieldname, plantingyear, receptiondate dateprod," & _
			" 0 totaldailyproduction, 0 totaldailyreceive, 0 Prevtotaldailyproduction, totaldailyreceive Prevtotaldailyreceive" & _
			" FROM view_dailyharvesting_mill" & _
			" WHERE fcba = '" & Session("fcBA")& "'" & _
			" AND accountingyear = " & v_txtyear & _
			" AND accountingperiod = " & v_txtmonth & _
			" AND receptiondate = TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY') - 1" & _
			" )" & _
			" GROUP BY fcba," & _
			" division," & _
			" fieldcode," & _
			" fieldname," & _
			" plantingyear" & _
			" )group by" & _
			" fcba," & _
			" division," & _
			" fieldcode," & _
			" fieldname," & _
			" plantingyear"					

%>

<html>

<head>
<title>iPS Reporting - Production Summary</title>
</head>

<body>

<table border="1" width="100%" id="table1" height="160">
	<tr>
		<td rowspan="4">NO</td>
		<td rowspan="4">AFDELING</td>
		<td rowspan="4">BLOK</td>
		<td rowspan="4">TAHUN TANAM</td>
		<td colspan="8">PRODUKSI TBS KEBUN</td>
		<td colspan="8">TBS DITERIMA PMKS</td>
		<td rowspan="2">RESTAN</td>
	</tr>
	<tr>
		<td colspan="4">JANJANG PANEN</td>
		<td colspan="4">TONASE</td>
		<td colspan="4">JANJANG KIRIM</td>
		<td colspan="4">TONASE</td>
	</tr>
	<tr>
		<td colspan="2">BI</td>
		<td>BUDGET</td>
		<td rowspan="2">SBI</td>
		<td colspan="2">BI</td>
		<td>BUDGET</td>
		<td rowspan="2">SBI</td>
		<td colspan="2">BI</td>
		<td>BUDGET</td>
		<td rowspan="2">SBI</td>
		<td colspan="2">BI</td>
		<td>BUDGET</td>
		<td rowspan="2">SBI</td>
		<td rowspan="2">HI</td>
	</tr>
	<tr>
		<td>HI</td>
		<td>SHI</td>
		<td>BI</td>
		<td>HI</td>
		<td>SHI</td>
		<td>BI</td>
		<td>HI</td>
		<td>SHI</td>
		<td>BI</td>
		<td>HI</td>
		<td>SHI</td>
		<td>BI</td>
	</tr>
<%	
	Set rsEPS = connStr.Execute(strEPS)
	if Not rsEPS.eof then
	   nCountofField = rsEPS.fields.count-1
	end if
	NO = 1
%>
<%While Not rsEPS.EOF%>
	<tr>
	
	
	<!--
				" sum(tonase_hi) tonase_hi," & _
			" sum(tonase_shi) tonase_shi," & _
			" sum(tonase_bi) tonase_bi," & _
			" sum(tonase_sbi) tonase_sbi," & _
			" sum(pom_hi) pom_hi," & _
			" sum(pom_shi) pom_shi," & _
			" sum(pom_bi) pom_bi," & _
			" sum(pom_sbi) pom_sbi," & _
			" sum(pom_tonase_hi) pom_tonase_hi," & _
			" sum(pom_tonase_shi) pom_tonase_shi," & _
			" sum(pom_tonase_bi) pom_tonase_bi," & _
			" sum(pom_tonase_sbi) pom_tonase_sbi," & _
			" sum(restan) restan" & _-->
		<td><%=NO%></td>
		<td><%=rsEPS("division")%></td>
		<td><%=rsEPS("fieldcode")%></td>
		<td><%=rsEPS("plantingyear")%></td>
		<td><%=rsEPS("hi")%></td>
		<td><%=rsEPS("shi")%></td>
		<td><%=rsEPS("bi")%></td>
		<td><%=rsEPS("sbi")%></td>
		<td><%=rsEPS("tonase_hi")%></td>
		<td><%=rsEPS("tonase_shi")%></td>
		<td><%=rsEPS("tonase_bi")%></td>
		<td><%=rsEPS("tonase_sbi")%></td>
		<td><%=rsEPS("pom_hi")%></td>
		<td><%=rsEPS("pom_shi")%></td>
		<td><%=rsEPS("pom_bi")%></td>
		<td><%=rsEPS("pom_sbi")%></td>
		<td><%=rsEPS("pom_tonase_hi")%></td>
		<td><%=rsEPS("pom_tonase_shi")%></td>
		<td><%=rsEPS("pom_tonase_bi")%></td>
		<td><%=rsEPS("pom_tonase_sbi")%></td>
		<td><%=rsEPS("restan")%></td>
	</tr>
<%rsEPS.MoveNext%>
<%NO = NO +1%>
<%Wend%>
<%rsEPS.close%>
<%rsEPS=null%>
</table>


</body>

</html>
