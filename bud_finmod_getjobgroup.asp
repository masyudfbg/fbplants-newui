
<%
Response.CharSet="euc-kr"
Session.codepage="949"
Response.codepage="949"
Response.ContentType="text/html;charset=euc-kr"
%>
<%
Response.LCID = 1046 
%>

<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->

<!--#includes file="include/JSON.asp" --> 
<!--#includes file="include/json.objectClass.asp"-->
<!--#includes file="include/aspJSON.asp" -->

<%


dim txtcari,fcba,gangcode,upahedit,bkmno,tglbkm,ovtimehours,typeovrtime
'fcba = session("fcba")
'txtcari = Request("keyword")
'gangcode = Request("gangcode")
'vehiclecode = Request("vehiclecode")
'itemcode = Request("itemcode")
'satuan = Request("satuan")
'tglbkm = Request("tglbkm")
'ovtimehours = Request("ovtimehours")
'typeovrtime = Request("typeovrtime")
'employee    = Request("employee")
'sectionname    = Request("sectionname")
regn    = Request("region")
pyear   = Request("year")
pmonth   = Request("month")

cmpy    = Request("company")
fcba    = Request("fcba")

if trim(Request("region")) = "-" then 
	regn = ""
else
	regn = trim(Request("region"))
end if


if trim(Request("company")) = "-" then 
	cmpy = ""
else
	cmpy = trim(Request("company"))
end if
ty    = Request("type")
id    = Request("id")

if id=1 then
 sqlview =		" SELECT  DISTINCT a.fccode,a.fcname,Nvl(b.job_forfieldstatus,'GENERAL') job_forfieldstatus," & _ 
 			" CASE WHEN (c.ownership)='INTI' THEN 'INTI' WHEN (c.ownership)='PLASMA' THEN 'PLASMA' ELSE 'INTI' END  ownership, NVL(qtyunit,0) qtyunit " & _
			" FROM  jobbtmaster a  " & _
			" INNER JOIN jobmaster b ON (a.fccode=b.job_forbudgetcode) " & _
			" LEFT JOIN " & _
			" (SELECT DX_BLOCK_STATUS,ownership,Round(Sum(nvl(luas_gross,0)),2) qtyunit " & _
			"  FROM field " & _
			"  WHERE activation='Y' AND fcba IN (SELECT fccode FROM businessunit WHERE  fccompanycode='" & cmpy & "' AND isactive='TRUE') " & _
			"  GROUP BY DX_BLOCK_STATUS,ownership " & _
			"  ) c ON (trim(b.job_forfieldstatus)=trim(c.DX_BLOCK_STATUS)) " & _
			" WHERE  " & _
			" Nvl(b.job_forfieldstatus,'-') <> '-' " & _
			" ORDER BY a.fccode,job_forfieldstatus " 

end if

if id=0 then
 sqlview =	" SELECT  DISTINCT a.fccode,a.fcname,'HA'  uom, nvl(b.job_forfieldstatus,'GC') job_forfieldstatus ,(c.ownership)ownership, nvl(qtyunit,0) qtyunit " & _
			" FROM jobmaster b INNER JOIN jobgroupmaster a ON (b.jobgroupcode=a.fccode) " & _
			" LEFT JOIN " & _
			" (SELECT DX_BLOCK_STATUS,ownership,Round(Sum(nvl(luas_gross,0)),2) qtyunit " & _
			"  FROM field_hist_tab " & _
			"  WHERE year=" & pyear & " AND month=1 AND activation='Y' AND fcba IN (SELECT fccode FROM businessunit WHERE  fccompanycode='" & cmpy & "' AND fccode='" & fcba & "' AND isactive='TRUE') " & _
			"  GROUP BY DX_BLOCK_STATUS,ownership " & _
			"  ) c ON (trim(b.job_forfieldstatus)=trim(c.DX_BLOCK_STATUS)) " & _
			" WHERE  " & _
			" job_forbudgetcode IN " & _
			"   ( " & _
			"   '101010'," & _
			"   '101020'," & _
			"   '101030'," & _
			"   '101040'," & _
			"   '101050'," & _
			"   '201101'," & _
			"   '201201'," & _
			"   '201301'," & _
			"   '201401'," & _
			"   '201501' " & _
			"   ) " & _
			" /*Nvl(b.job_forfieldstatus,'-') <> '-' */" & _
			" ORDER BY job_forfieldstatus,a.fccode,nvl(c.ownership) " 

end if

if id=2 then
sqlview =   " WITH fin AS (" & _
			" SELECT  DISTINCT  '" & cmpy & "' fccompanycode," & pyear &" period_year,a.fccode,a.fcname,Nvl(b.job_forfieldstatus,'GENERAL') job_forfieldstatus," & _ 
			" CASE WHEN (c.ownership)='INTI' THEN 'INTI' WHEN (c.ownership)='PLASMA' THEN 'PLASMA' ELSE 'INTI' END  ownership, Nvl(qtyunit,0) qtyunit " & _
			"  FROM  jobbtmaster a " & _
			"  INNER JOIN jobmaster b ON (a.fccode=b.job_forbudgetcode)" & _
			"  LEFT JOIN" & _
			"  (" & _
			" SELECT status,ownership,Round(Sum(nvl(luas_gross,0)),2) qtyunit" & _
			"   FROM field" & _
			"  WHERE activation='Y' AND fcba IN (SELECT fccode FROM businessunit WHERE  fccompanycode='" & cmpy & "' AND isactive='TRUE')" & _
			"   GROUP BY status,ownership" & _
			"   ) c ON (trim(b.job_forfieldstatus)=trim(c.dx_block_status))" & _
			"  ORDER BY a.fccode,job_forfieldstatus " & _
			" )" & _
			" SELECT a.fccompanycode,a.period_year,a.fccode jobgroupcode,a.fcname jobgroupname,a.job_forfieldstatus,a.ownership,a.qtyunit,nvl(b.qty,0) qty,nvl(b.price,0) price,nvl(b.amount,0) amount,nvl(b.remarks,' ') remarks" & _
			" FROM fin a " & _
			" LEFT JOIN  bud_finmod b ON (a.fccompanycode=b.fccompanycode AND a.period_year=b.period_year AND a.fccode=b.jobgroupcode AND a.ownership=b.ownership) " & _
			" Order By a.job_forfieldstatus,a.fccode"

end if

if id=4 then
sqlview = 	"   SELECT a.fccompanycode,a.period_year,a.jobgroupcode,a.jobgroupname,a.job_forfieldstatus,a.qty,a.price,a.amount,nvl(remarks,' ') remarks,a.ownership,a.qtyunit," & _
			"	Round(Nvl(b.qtybi,0),2) qtyactBi,Decode(  Nvl(b.qtybi,0),0,0, Round(Nvl(b.amountbi,0) /  Nvl(b.qtybi,0),2)) priceactBi, Round(Nvl(b.amountbi,0),2) amountactBi, " & _
			"	Round(Nvl(c.qtysbi,0),2) qtyactSBi,Decode(  Nvl(c.qtysbi,0),0,0, Round(Nvl(c.amountsbi,0) /  Nvl(c.qtysbi,0),2)) priceactSBi, Round(Nvl(c.amountsbi,0),2) amountactSBi " & _
			"	FROM bud_finmod a " & _
			"	LEFT JOIN " & _
			"	 ( " & _
			"		SELECT " & _
			"		c.fccompanycode,accountingyear period_year,d.job_forbudgetcode jobgroupcode,Nvl(d.job_forfieldstatus,'GENERAL') job_forfieldstatus,(f.ownership) ownership," & _
			"		Sum(Decode(d.budgetqty,'YY',a.unit,0)) qtybi,Sum(amountdebit+amountkredit) amountbi " & _
			"		FROM bkmjournaldtl a " & _
			"		INNER JOIN bkmjournalhdr b ON (a.docno=b.docno AND a.fcba=b.fcba) " & _
			"		INNER JOIN businessunit c ON (a.fcba=c.fccode) " & _
			"		INNER JOIN ( SELECT fccode,fcname,job_forbudgetcode,job_forfieldstatus,trim(budgetqty)||Nvl(bgtqty,'N') budgetqty" & _
			"					 FROM " & _
			"					 jobmaster x left JOIN " & _
			"					 (SELECT refno,'Y' bgtqty FROM  finstate_templatedtl WHERE reportcode='LM_PROD'  AND refno <> '-' ) y ON (x.fccode=y.refno)  " & _
			"				   ) d ON (a.jobcode=d.fccode) " & _
			"		LEFT JOIN field_hist_tab f ON (a.locationcode = f.fccode AND a.fcba=f.fcba AND b.accountingperiod=f.MONTH AND b.accountingyear=f.year ) " & _
			"		WHERE " & _
			"		c.fccompanycode in (select fccompanycode from businessunit where region like '" & regn & "%' AND fccompanycode like '" & cmpy & "%') " & _ 
			"		AND b.accountingyear=" & pyear & "" & _
			"		AND b.accountingperiod=" & pmonth & "" & _
			"		/*AND Nvl(d.job_forfieldstatus,'-') <> '-' */" & _
			"		AND b.postingflg='1' " & _
			"		GROUP BY  c.fccompanycode,b.accountingyear,d.job_forbudgetcode ,Nvl(d.job_forfieldstatus,'GENERAL'),(f.ownership) " & _
			"	 )b ON (a.fccompanycode = b.fccompanycode AND a.period_year=b.period_year AND  a.jobgroupcode=b.jobgroupcode AND a.job_forfieldstatus=b.job_forfieldstatus AND a.ownership=b.ownership) " & _
			"	LEFT JOIN " & _
			"	 ( " & _
			"		SELECT " & _
			"		c.fccompanycode,accountingyear period_year,d.job_forbudgetcode jobgroupcode,Nvl(d.job_forfieldstatus,'GENERAL') job_forfieldstatus,(f.ownership) ownership," & _
			"		Sum(Decode(d.budgetqty,'YY',a.unit,0)) qtysbi,Sum(amountdebit+amountkredit) amountsbi " & _
			"		FROM bkmjournaldtl a " & _
			"		INNER JOIN bkmjournalhdr b ON (a.docno=b.docno AND a.fcba=b.fcba) " & _
			"		INNER JOIN businessunit c ON (a.fcba=c.fccode) " & _
			"		INNER JOIN ( SELECT fccode,fcname,job_forbudgetcode,job_forfieldstatus,trim(budgetqty)||Nvl(bgtqty,'N') budgetqty" & _
			"					 FROM " & _
			"					 jobmaster x left JOIN " & _
			"					 (SELECT refno,'Y' bgtqty FROM  finstate_templatedtl WHERE reportcode='LM_PROD'  AND refno <> '-' ) y ON (x.fccode=y.refno)  " & _
			"				   ) d ON (a.jobcode=d.fccode) " & _
			"		LEFT JOIN field_hist_tab f ON (a.locationcode = f.fccode AND a.fcba=f.fcba AND b.accountingperiod=f.MONTH AND b.accountingyear=f.year ) " & _
			"		WHERE " & _
			"		c.fccompanycode in (select fccompanycode from businessunit where region like '" & regn & "%' AND fccompanycode like '" & cmpy & "%') " & _ 
			"		AND b.accountingyear=" & pyear & "" & _
			"		AND b.accountingperiod <=" & pmonth & "" & _
			"		/*AND Nvl(d.job_forfieldstatus,'-') <> '-' */" & _
			"		AND b.postingflg='1' " & _
			"		GROUP BY  c.fccompanycode,b.accountingyear,d.job_forbudgetcode ,Nvl(d.job_forfieldstatus,'GENERAL'),(f.ownership) " & _
			"	 )c ON (a.fccompanycode = c.fccompanycode AND a.period_year=c.period_year AND  a.jobgroupcode=c.jobgroupcode AND a.job_forfieldstatus=c.job_forfieldstatus AND a.ownership=c.ownership) " & _
			"	WHERE  " & _
			" 	a.fccompanycode in (select fccompanycode from businessunit where region like '" & regn & "%' AND fccompanycode like '" & cmpy & "%') " & _
			"	AND (a.Amount<>0 OR b.amountbi <> 0 OR c.amountsbi <> 0 ) AND a.job_forfieldstatus like '" & ty & "%' AND a.period_year=" & pyear 


end if

if id=3 then
sqlview =   " SELECT fccompanycode,fcbusinessunit,period_year,jobgroupcode,jobgroupname,uom,job_forfieldstatus,qty,price,amount,nvl(remarks,' ') remarks,ownership,qtyunit " & _ 
			" FROM bud_rkt " & _
			" WHERE " & _ 
			" fccompanycode in (select fccompanycode from businessunit where region like '" & regn & "%' AND fccompanycode like '" & cmpy & "%' AND fccode = '" & fcba & "') AND job_forfieldstatus like '" & ty & "%' AND period_year=" & pyear 

end if

if id=5 then
sqlview =   " SELECT a.fccompanycode,a.fcbusinessunit, a.period_year,a.jobgroupcode,a.jobgroupname,a.job_forfieldstatus,round(a.qty,2) qty,a.price,a.amount,nvl(a.remarks,' ') remarks,a.ownership,a.qtyunit, " & _ 
			" nvl(b.qty,0) qtyfinmod,nvl(b.price,0) pricefinmod,nvl(b.amount,0) amountfinmod,  " & _ 
			" Round(Nvl(c.qty,0),2) qtyact,Decode(  Nvl(c.qty,0),0,0, Round(Nvl(c.amount,0) /  Nvl(c.qty,0),2)) priceact, Round(Nvl(c.amount,0),2) amountact " & _       
			" FROM bud_rkt a " & _
			" LEFT JOIN bud_finmod b on (a.fccompanycode=b.fccompanycode AND a.period_year=b.period_year AND a.jobgroupcode=b.jobgroupcode AND a.job_forfieldstatus=b.job_forfieldstatus AND a.ownership=b.ownership) " & _
			" LEFT JOIN " & _
			"	 ( " & _
			"		SELECT " & _
			"		c.fccompanycode,c.fccode fcbusinessunit,accountingyear period_year,d.jobgroupcode jobgroupcode,NVL(d.job_forfieldstatus,'GC') job_forfieldstatus,NVL(f.ownership) ownership," & _
			"		Sum(Decode(d.budgetqty,'N',0,a.unit)) qty,Sum(amountdebit+amountkredit) amount " & _
			"		FROM bkmjournaldtl a " & _
			"		INNER JOIN bkmjournalhdr b ON (a.docno=b.docno AND a.fcba=b.fcba) " & _
			"		INNER JOIN businessunit c ON (a.fcba=c.fccode) " & _
			"		INNER JOIN jobmaster d ON (a.jobcode=d.fccode) " & _
			"		INNER JOIN field f ON (a.locationcode = f.fccode AND a.fcba=f.fcba ) " & _
			"		WHERE " & _
			"       c.fccode like '" & fcba & "%'" & _
			"		AND c.fccompanycode in (select fccompanycode from businessunit where region like '" & regn & "%' AND fccompanycode like '" & cmpy & "%') " & _ 
			"		AND b.accountingyear=" & pyear & "" & _
			"		/*AND Nvl(d.job_forfieldstatus,'-') <> '-' */ " & _
			"		AND b.postingflg='1' " & _
			"		GROUP BY  c.fccompanycode,c.fccode,b.accountingyear,d.jobgroupcode ,d.job_forfieldstatus,f.ownership " & _
			"	 )c ON (a.fccompanycode = c.fccompanycode AND a.fcbusinessunit=c.fcbusinessunit AND a.period_year=c.period_year AND  a.jobgroupcode=c.jobgroupcode AND a.job_forfieldstatus=c.job_forfieldstatus AND a.ownership=c.ownership) " & _
			" WHERE " & _ 
			"  a.fcbusinessunit Like '" & fcba & "%' AND a.fccompanycode in (select fccompanycode from businessunit where region like '" & regn & "%' AND fccompanycode like '" & cmpy & "%') AND a.job_forfieldstatus like '" & ty & "%' AND a.period_year=" & pyear 

end if


'Response.write sqlview
 
Set List = Server.CreateObject("ADODB.Recordset")
List.Open sqlview, connStr, 1
ListCount = List.RecordCount 

Set caljson = new JSON  'JSON.asp가 필요함
Set rs = connStr.Execute(sqlview)

Dim List_jsonString
List_jsonString = caljson.toJSON("results",rs, false)

Dim jsonObj, outputObj
set jsonObj = new JSONobject
set outputObj = jsonObj.parse(List_jsonString)  'json.objectClass.asp가 필요함(파싱하려면)


Set oJSON = New aspJSON     'aspJSON.asp 필요함(for문)
oJSON.loadJSON(List_jsonString)

Response.write List_jsonString



%>