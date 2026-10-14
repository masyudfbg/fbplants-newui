
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

dim region,company,bulan,tahun,costelemen,t1,r1
dim txtcari,fcba,bkmfrom,bkmto,idview,actfrom,actto,sectioncode,statustm
dim filterdiv,filtersource,filterfcba,filterprd,filterprdlalu,filtercostelemen,filterprdsbi,filtertmtbm
dim prev_bulanfrom,prev_bulanto
dim bulansplitfrom,tglsplitfrom,thnsplitfrom,tglfrom,tglto,tglawalsbi
dim bulansplitto,tglsplitto,thnsplitto,xbln1,xthn1
idview   = Request("idview")

region 	= Request("region")
company 	= Request("company")
fcba 		= Request("fcba")
actfrom   	= Request("actfrom")
actto   	= Request("actto")
bulan   	= Request("bulan")
tahun   	= Request("tahun")
costelemen  = Request("costelemen")
sectioncode	= Request("section")
statustm	= Request("statustm")
t1   		= Request("t1")
r1   		= Request("r1")
'tampilkan data listing bkm'

bulansplitfrom  = left(actfrom,2)
tglsplitfrom    = mid(actfrom,4,2)
thnsplitfrom    = right(actfrom,4)
tglfrom     	= tglsplitfrom & "/" & bulansplitfrom & "/" & thnsplitfrom

'rubah format tgl to jd dd/mm/yyyy
bulansplitto  	= left(actto,2)
tglsplitto    	= mid(actto,4,2)
thnsplitto    	= right(actto,4)
tglto     		= tglsplitto & "/" & bulansplitto & "/" & thnsplitto

'tglawal tahun
tglawalsbi = "01" & "/" & "01" & "/" & thnsplitfrom


prev_bulanfrom  = dateadd("m", -1,tglfrom)
prev_bulanto	= dateadd("m", -1,tglto)

xbln1       = bulan
xthn1       = tahun
if  xbln1   = 1 then
	xbln2   = 12
	xthn2   = xthn1-1
else   
	xbln2   = xbln1-1
	xthn2   = xthn1
end if


'qry laporan act report
if idview="1" then
    
   

	'rubah format tgl awal jd dd/mm/yyyy
	' bulansplitfrom  = left(actfrom,2)
    ' tglsplitfrom    = mid(actfrom,4,2)
    ' thnsplitfrom    = right(actfrom,4)
    ' tglfrom     = tglsplitfrom & "/" & bulansplitfrom & "/" & thnsplitfrom

	' 'rubah format tgl to jd dd/mm/yyyy
	' bulansplitto  	= left(actto,2)
    ' tglsplitto    	= mid(actto,4,2)
    ' thnsplitto    	= right(actto,4)
    ' tglto     		= tglsplitto & "/" & bulansplitto & "/" & thnsplitto

	' 'tglawal tahun
	' tglawalsbi = "01" & "/" & "01" & "/" & thnsplitfrom
    
	
	' prev_bulanfrom  = dateadd("m", -1,tglfrom)
	' prev_bulanto	= dateadd("m", -1,tglto)

	' xbln1       = bulan
    ' xthn1       = tahun
    ' if  xbln1   = 1 then
    '     xbln2   = 12
    '     xthn2   = xthn1-1
    ' else   
    '     xbln2   = xbln1-1
    '     xthn2   = xthn1
    ' end if
	


	

	'filter by company
	if trim(company) = "all" then 
		filtercompany = " WHERE t.region = '" & region & "'"
	else
		if trim(fcba) = "-" then 
	    	filtercompany = " WHERE t.fccompanycode ='" & company & "' "
		else 
			filtercompany = " WHERE t.fccode ='" & fcba & "' "
		end if
	end if

	'bulan lalu
    if trim(t1) = "Tgl" THEN
		filterprdlalu = " (o.transdate between to_date('" & prev_bulanfrom & "','dd/mm/yyyy') and to_date('" & prev_bulanto & "','dd/mm/yyyy')) " 			   
	else
		filterprdlalu = " o.period_month='" & xbln2 & "' " & _ 
					" AND o.period_year ='" & xthn2 & "' " 		   
	end if

    if trim(t1) = "Tgl" THEN
		filterprd = " (o.transdate between to_date('" & actfrom & "','mm/dd/yyyy') and to_date('" & actto & "','mm/dd/yyyy')) " 			   
	else
		filterprd = " o.period_month='" & bulan & "' " & _ 
					" AND o.period_year ='" & tahun & "' " 		   
	end if

	'filter sbi
    if trim(t1) = "Tgl" THEN
		filterprdsbi = " (o.transdate between to_date('" & tglawalsbi & "','mm/dd/yyyy') and to_date('" & actto & "','mm/dd/yyyy')) " 			   
	else
		filterprdsbi = " o.period_month <= '" & bulan & "' AND o.period_year ='" & tahun & "' " 		   
	end if

	'filter by cost elemen
	if trim(costelemen) <> "all" then 
		filtercostelemen = " AND t.source='" & costelemen & "' "
	end if

	if trim(sectioncode) <> "all" then
		filterdiv = " AND t.sectioncode='" & sectioncode & "'"
	end if

	if trim(statustm) <> "all" then
		filtertmtbm = " AND j.job_forfieldstatus='" & statustm & "'"
	end if


    sqlview =  "select t.region,t.fccompanycode,t.fccode,j.job_forfieldstatus fieldstatus,t.source,t.sectioncode,t.fccode jobcode,t.jobname, " & _
	" NVL(sum(t.hkblnlalu),0)hkblnlalu,NVL(sum(t.fisikblnlalu),0)fisikblnlalu,NVL(sum(t.unitblnlalu),0)unitblnlalu,NVL(sum(t.amountblnlalu),0)amountblnlalu, " & _
	" sum(t.hkblnini)hkblnini,sum(t.fisikblnini)fisikblnini,sum(t.unitblnini)unitblnini,sum(t.amountblnini)amountblnini, " & _
	" sum(t.hksbi)hksbi,sum(t.fisiksbi)fisiksbi,sum(t.unitsbi)unitsbi,sum(t.amountsbi)amountsbi, " & _
	" NVL(SUM(t.rkb_hkblnlalu),0)rkb_hkblnlalu,NVL(SUM(t.rkb_outputblnlalu),0)rkb_outputblnlalu,NVL(SUM(t.rkb_amountblnlalu),0)rkb_amountblnlalu, " & _
	" NVL(SUM(t.rkb_hkblnini),0)rkb_hkblnini,NVL(SUM(t.rkb_outputblnini),0)rkb_outputblnini,NVL(SUM(t.rkb_amountblnini),0)rkb_amountblnini, " & _
	" NVL(SUM(t.rkb_hksbi),0)rkb_hksbi,NVL(SUM(t.rkb_outputsbi),0)rkb_outputsbi,NVL(SUM(t.rkb_amountsbi),0)rkb_amountsbi " & _
	" from( " & _
			" SELECT b.region,b.fccompanycode,b.fccode,o.sectioncode,o.source,o.fccode jobcode, o.fccode || ' - ' || j.fcname jobname, " & _
			" SUM(o.hk) hkblnlalu,SUM(o.OUTPUT) fisikblnlalu,SUM(o.unit) unitblnlalu,sum(o.amount)amountblnlalu, " & _
			" 0 hkblnini,0 fisikblnini,0 unitblnini,0 amountblnini, " & _
			" 0 hksbi,0 fisiksbi,0 unitsbi,0 amountsbi, " & _
			" sum(o.rkb_hk) rkb_hkblnlalu,sum(o.rkb_output) rkb_outputblnlalu,sum(o.rkb_amount) rkb_amountblnlalu, " & _
			" 0 rkb_hkblnini,0 rkb_outputblnini,0 rkb_amountblnini, " & _
			" 0 rkb_hksbi,0 rkb_outputsbi,0 rkb_amountsbi " & _
			" FROM OPR_ACCOUNTREPORT_BULANAN  o " & _
			" LEFT JOIN BUSINESSUNIT b on o.FCBA=b.fccode " & _
			" LEFT JOIN JOBMASTER j on j.fccode=o.fccode  " & _
			" LEFT JOIN FIELD f on f.fccode=o.FIELDCODE and f.fcba=o.fcba " & _
			" WHERE  " & _
			filterprdlalu & _
			" group by  b.region,b.fccompanycode,b.fccode,o.sectioncode,o.source,o.fccode, o.fccode || ' - ' || j.fcname " & _
			"  union all " & _
			" SELECT b.region,b.fccompanycode,b.fccode,o.sectioncode,o.source,o.fccode jobcode, o.fccode || ' - ' || j.fcname jobname, " & _
			" 0 hkblnlalu,0 fisikblnlalu,0 unitblnlalu,0 amountblnlalu, " & _
			" SUM(o.hk) hkblnini,SUM(o.OUTPUT) fisikblnini,SUM(o.unit) unitblnini,sum(o.amount)amountblnini, " & _
			" 0 hksbi,0 fisiksbi,0 unitsbi,0 amountsbi, " & _
			" 0 rkb_hkblnlalu,0 rkb_outputblnlalu,0 rkb_amountblnlalu, " & _
			" SUM(o.rkb_hk)rkb_hkblnini,SUM(o.rkb_output)rkb_outputblnini,SUM(o.rkb_amount)rkb_amountblnini, " & _
			" 0 rkb_hksbi,0 rkb_outputsbi,0 rkb_amountsbi " & _
			" FROM OPR_ACCOUNTREPORT_BULANAN  o " & _
			" LEFT JOIN BUSINESSUNIT b on o.FCBA=b.fccode " & _
			" LEFT JOIN JOBMASTER j on j.fccode=o.fccode " & _
			" LEFT JOIN FIELD f on f.fccode=o.FIELDCODE and f.fcba=o.fcba " & _
			" WHERE " & _
			filterprd & _
			" group by  b.region,b.fccompanycode,b.fccode,o.sectioncode,o.source,o.fccode, o.fccode || ' - ' || j.fcname " & _
			"  union all " & _
			" SELECT b.region,b.fccompanycode,b.fccode,o.sectioncode,o.source,o.fccode jobcode, o.fccode || ' - ' || j.fcname jobname, " & _
			" 0 hkblnlalu,0 fisikblnlalu,0 unitblnlalu,0 amountblnlalu, " & _
			" 0 hkblnini,0 fisikblnini,0 unitblnini,0 amountblnini, " & _
			" SUM(o.hk) hksbi,SUM(o.OUTPUT) fisiksbi,SUM(o.unit) unitsbi,sum(o.amount)amountsbi, " & _
			" 0 rkb_hkblnlalu,0 rkb_outputblnlalu,0 rkb_amountblnlalu, " & _
			" 0 rkb_hkblnini,0 rkb_outputblnini,0 rkb_amountblnini, " & _
			" sum(o.rkb_hk) rkb_hksbi,sum(o.rkb_output) rkb_outputsbi,sum(o.rkb_amount) rkb_amountsbi " & _
			" FROM OPR_ACCOUNTREPORT_BULANAN  o " & _
			" LEFT JOIN BUSINESSUNIT b on o.FCBA=b.fccode " & _
			" LEFT JOIN JOBMASTER j on j.fccode=o.fccode " & _
			" LEFT JOIN FIELD f on f.fccode=o.FIELDCODE and f.fcba=o.fcba " & _
			" WHERE " & _
			filterprdsbi & _
			" group by  b.region,b.fccompanycode,b.fccode,o.sectioncode,o.source,o.fccode, o.fccode || ' - ' || j.fcname " & _
	" )t LEFT JOIN jobmaster j ON t.jobcode=j.fccode" & _
	filtercompany & _
	filtercostelemen & _
	filterdiv & _
	filtertmtbm & _

	" group by t.region,t.fccompanycode,t.fccode,j.job_forfieldstatus,t.sectioncode,t.source,t.jobname " & _
	" order by t.region,t.fccompanycode,t.fccode,t.jobname,j.job_forfieldstatus,t.sectioncode,t.source"

	Set List = Server.CreateObject("ADODB.Recordset")
	List.Open sqlview, connStr, 1
    ListCount = List.RecordCount 

    Set caljson = new JSON  'JSON.asp가 필요함P
    Set rs = connStr.Execute(sqlview)

    Dim List_jsonString
    List_jsonString = caljson.toJSON("results",rs, false)

    Dim jsonObjg, outputObjg
    set jsonObjg = new JSONobject
    set outputObjg = jsonObjg.parse(List_jsonString)  'json.objectClass.asp가 필요함(파싱하려면)


    Set oJSON = New aspJSON     'aspJSON.asp 필요함(for문)
    oJSON.loadJSON(List_jsonString)

    Response.write List_jsonString
end if

if idview="2" then
	' region 		= Request("region")
	' company 	= Request("company")
    ' fcba 		= Request("fcba")
    ' actfrom   	= Request("actfrom")
    ' actto   	= Request("actto")
    ' bulan   	= Request("bulan")
    ' tahun   	= Request("tahun")
    ' costelemen  = Request("costelemen")
    ' sectioncode	= Request("section")
    ' statustm	= Request("statustm")
    ' t1   		= Request("t1")
    ' r1   		= Request("r1")

	'rubah format tgl awal jd dd/mm/yyyy
	

	

	'filter by company
	if trim(company) = "all" then 
		filtercompany = " WHERE t.region = '" & region & "'"
	else
		if trim(fcba) = "-" then 
	    	filtercompany = " WHERE t.fccompanycode ='" & company & "' "
		else 
			filtercompany = " WHERE t.fccode ='" & fcba & "' "
		end if
	end if

	'bulan lalu
    if trim(t1) = "Tgl" THEN
		filterprdlalu = " (o.transdate between to_date('" & prev_bulanfrom & "','dd/mm/yyyy') and to_date('" & prev_bulanto & "','dd/mm/yyyy')) " 			   
	else
		filterprdlalu = " o.period_month='" & xbln2 & "' " & _ 
					" AND o.period_year ='" & xthn2 & "' " 		   
	end if

    if trim(t1) = "Tgl" THEN
		filterprd = " (o.transdate between to_date('" & actfrom & "','mm/dd/yyyy') and to_date('" & actto & "','mm/dd/yyyy')) " 			   
	else
		filterprd = " o.period_month='" & bulan & "' " & _ 
					" AND o.period_year ='" & tahun & "' " 		   
	end if


	'filter sbi
    if trim(t1) = "Tgl" THEN
		filterprdsbi = " (o.transdate between to_date('" & tglawalsbi & "','mm/dd/yyyy') and to_date('" & actto & "','mm/dd/yyyy')) " 			   
	else
		filterprdsbi = " o.period_month <= '" & bulan & "' AND o.period_year ='" & tahun & "' " 		   
	end if

	'filter by cost elemen
	if trim(costelemen) <> "all" then 
		filtercostelemen = " AND t.source='" & costelemen & "' "
	end if

	if trim(sectioncode) <> "all" then
		filterdiv = " AND t.sectioncode='" & sectioncode & "'"
	end if

	if trim(statustm) <> "all" then
		filtertmtbm = " AND t.job_forfieldstatus='" & statustm & "'"
	end if


	sqlview1 =  "select t.region,t.fccompanycode,t.fccode,t.source,t.job_forfieldstatus fieldstatus,t.fccode jobcode,t.jobname, " & _
	" t.sectioncode, t.fieldcode,t.fieldarea,t.fieldpokok, " & _
	" NVL(sum(t.hkblnlalu),0)hkblnlalu,NVL(sum(t.fisikblnlalu),0)fisikblnlalu,NVL(sum(t.unitblnlalu),0)unitblnlalu,NVL(sum(t.amountblnlalu),0)amountblnlalu, " & _
	" sum(t.hkblnini)hkblnini,sum(t.fisikblnini)fisikblnini,sum(t.unitblnini)unitblnini,sum(t.amountblnini)amountblnini, " & _
	" sum(t.hksbi)hksbi,sum(t.fisiksbi)fisiksbi,sum(t.unitsbi)unitsbi,sum(t.amountsbi)amountsbi, " & _
	" NVL(SUM(t.rkb_hkblnlalu),0)rkb_hkblnlalu,NVL(SUM(t.rkb_outputblnlalu),0)rkb_outputblnlalu,NVL(SUM(t.rkb_amountblnlalu),0)rkb_amountblnlalu, " & _
	" NVL(SUM(t.rkb_hkblnini),0)rkb_hkblnini,NVL(SUM(t.rkb_outputblnini),0)rkb_outputblnini,NVL(SUM(t.rkb_amountblnini),0)rkb_amountblnini, " & _
	" NVL(SUM(t.rkb_hksbi),0)rkb_hksbi,NVL(SUM(t.rkb_outputsbi),0)rkb_outputsbi,NVL(SUM(t.rkb_amountsbi),0)rkb_amountsbi " & _
	" from( " & _
			" SELECT b.region,b.fccompanycode,b.fccode,o.source,o.fccode jobcode, o.fccode || ' - ' || j.fcname jobname,j.job_forfieldstatus, " & _
			" o.sectioncode,o.fieldcode,o.fieldarea,o.fieldpokok, " & _
			" SUM(o.hk) hkblnlalu,SUM(o.OUTPUT) fisikblnlalu,SUM(o.unit) unitblnlalu,sum(o.amount)amountblnlalu, " & _
			" 0 hkblnini,0 fisikblnini,0 unitblnini,0 amountblnini, " & _
			" 0 hksbi,0 fisiksbi,0 unitsbi,0 amountsbi, " & _
			" sum(o.rkb_hk) rkb_hkblnlalu,sum(o.rkb_output) rkb_outputblnlalu,sum(o.rkb_amount) rkb_amountblnlalu, " & _
			" 0 rkb_hkblnini,0 rkb_outputblnini,0 rkb_amountblnini, " & _
			" 0 rkb_hksbi,0 rkb_outputsbi,0 rkb_amountsbi " & _
			" FROM OPR_ACCOUNTREPORT_BULANAN  o " & _
			" LEFT JOIN BUSINESSUNIT b on o.FCBA=b.fccode  and b.fccode=o.FCBA " & _
			" LEFT JOIN JOBMASTER j on j.fccode=o.fccode  " & _
			
			" WHERE  " & _
			filterprdlalu & _
			" group by b.region,b.fccompanycode,b.fccode,o.source,o.fccode, o.fccode || ' - ' || j.fcname,j.job_forfieldstatus,o.sectioncode,o.fieldcode,o.fieldarea,o.fieldpokok " & _
			
			"  union all " & _
			" SELECT b.region,b.fccompanycode,b.fccode,o.source,o.fccode jobcode, o.fccode || ' - ' || j.fcname jobname,j.job_forfieldstatus, " & _
			" o.sectioncode,o.fieldcode,o.fieldarea,o.fieldpokok, " & _
			" 0 hkblnlalu,0 fisikblnlalu,0 unitblnlalu,0 amountblnlalu, " & _
			" SUM(o.hk) hkblnini,SUM(o.OUTPUT) fisikblnini,SUM(o.unit) unitblnini,sum(o.amount)amountblnini, " & _
			" 0 hksbi,0 fisiksbi,0 unitsbi,0 amountsbi, " & _
			" 0 rkb_hkblnlalu,0 rkb_outputblnlalu,0 rkb_amountblnlalu, " & _
			" SUM(o.rkb_hk)rkb_hkblnini,SUM(o.rkb_output)rkb_outputblnini,SUM(o.rkb_amount)rkb_amountblnini, " & _
			" 0 rkb_hksbi,0 rkb_outputsbi,0 rkb_amountsbi " & _
			" FROM OPR_ACCOUNTREPORT_BULANAN  o " & _
			" LEFT JOIN BUSINESSUNIT b on o.FCBA=b.fccode and b.fccode=o.FCBA " & _
			" LEFT JOIN JOBMASTER j on j.fccode=o.fccode " & _
			
			" WHERE " & _
			filterprd & _
			" group by  b.region,b.fccompanycode,b.fccode,o.source,o.fccode, o.fccode || ' - ' || j.fcname,j.job_forfieldstatus,o.sectioncode,o.fieldcode,o.fieldarea,o.fieldpokok " & _
			
			"  union all " & _
			" SELECT b.region,b.fccompanycode,b.fccode,o.source,o.fccode jobcode, o.fccode || ' - ' || j.fcname jobname,j.job_forfieldstatus, " & _
			" o.sectioncode,o.fieldcode,o.fieldarea,o.fieldpokok, " & _
			" 0 hkblnlalu,0 fisikblnlalu,0 unitblnlalu,0 amountblnlalu, " & _
			" 0 hkblnini,0 fisikblnini,0 unitblnini,0 amountblnini, " & _
			" SUM(o.hk) hksbi,SUM(o.OUTPUT) fisiksbi,SUM(o.unit) unitsbi,sum(o.amount)amountsbi, " & _
			" 0 rkb_hkblnlalu,0 rkb_outputblnlalu,0 rkb_amountblnlalu, " & _
			" 0 rkb_hkblnini,0 rkb_outputblnini,0 rkb_amountblnini, " & _
			" sum(o.rkb_hk) rkb_hksbi,sum(o.rkb_output) rkb_outputsbi,sum(o.rkb_amount) rkb_amountsbi " & _
			" FROM OPR_ACCOUNTREPORT_BULANAN  o " & _
			" LEFT JOIN BUSINESSUNIT b on o.FCBA=b.fccode  and b.fccode=o.FCBA " & _
			" LEFT JOIN JOBMASTER j on j.fccode=o.fccode " & _
			
			" WHERE " & _
			filterprdsbi & _
			" group by  b.region,b.fccompanycode,b.fccode,o.source,o.fccode, o.fccode || ' - ' || j.fcname,j.job_forfieldstatus,o.sectioncode,o.fieldcode,o.fieldarea,o.fieldpokok " & _
			
	" )t " & _
	filtercompany & _
	filtercostelemen & _
	filterdiv & _
	filtertmtbm & _

	" group by t.region,t.fccompanycode,t.fccode,t.source,t.sectioncode, t.job_forfieldstatus,t.jobname,t.fieldcode,t.fieldarea,t.fieldpokok" & _
	" order by t.region,t.fccompanycode,t.fccode,t.jobname,t.job_forfieldstatus,t.sectioncode"


 ' response.write sqlview1

  	Set List1 = Server.CreateObject("ADODB.Recordset")

    List1.Open sqlview1, connStr, 1
    ListCount = List1.RecordCount 

    Set caljson = new JSON  'JSON.asp가 필요함P
    Set rs = connStr.Execute(sqlview1)

    Dim List_jsonString1
    List_jsonString1 = caljson.toJSON("results1",rs, false)

    Dim jsonObjg1, outputObjg1
    set jsonObjg1 = new JSONobject
    set outputObjg1 = jsonObjg1.parse(List_jsonString1)  'json.objectClass.asp가 필요함(파싱하려면)


    Set oJSON = New aspJSON     'aspJSON.asp 필요함(for문)
    oJSON.loadJSON(List_jsonString1)

    Response.write List_jsonString1

end if
%>