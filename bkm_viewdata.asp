<% Response.Buffer = False %>
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


dim txtcari,fcba,bkmfrom,bkmto,idview,company,filtersection,filterrowstate
idview   = Request("idview")
'tampilkan data listing bkm'
if idview="1" then

    fcba = session("fcba")
    txtcari = Request("keyword")
    bkmfrom = Request("bkmfrom")
    bkmto   = Request("bkmto")
	pilihbkm   = Request("pilihbkm")
	section = Request("section")
	posting = Request("posting")

	sqlexec = "BEGIN " & _ 
              "bkmproses_upah_api.Update_hdr_fromto(to_date('" & bkmfrom & "','mm/dd/yyyy') , to_date('" & bkmto & "','mm/dd/yyyy') ,'" & fcba & "');" & _
			  "END; " 
    connStr.Execute(sqlexec)
    
    Set List = Server.CreateObject("ADODB.Recordset")

	if trim(section) <> "" then 
	    filtersection = " AND H.SECTIONCODE ='" & section & "' "
	end if
	if trim(posting) <> "" then 
	    filterrowstate = " AND H.ROWSTATE ='" & posting & "' "
	end if

    if pilihbkm= "504" then 

        sqlview = "Select h.BKMNO, h.BKMDATE,CASE WHEN document_classification = '504' THEN contractor_api.get_contractorname(h.contractorcode,h.fcba) ELSE h.SECTIONCODE END SECTIONCODE,h.GANGCODE,nvl(h.VEHICLECODE,'-') VEHICLECODE, h.TOTUPAH TOTALBKMUPAH,h.TOTBAHAN TOTALBKMMATERIAL,h.TOTVRA TOTALBKMVEHICLE,h.TOTWIP TOTALBKMWIP,h.DOCUMENT_CLASSIFICATION " & _
        "    ,authorizestructure_api.get_template_desc(h.DOCUMENT_CLASSIFICATION) dcname,decode(h.hujan,'0','-','YA') hujan,decode(h.POSTINGFLG,'1','YA','TIDAK') POSTINGFLG" & _
        "    ,h.TOTUPAH + h.TOTBAHAN + h.TOTVRA + h.TOTWIP TOTAL,h.createdate INPUTDATE, h.LASTUPDATE , h.RKBNO,h.rowstate,h.FCENTRY,h.FCEDIT  " & _	
        "    From   BKMHDR h WHERE h.BKMDATE between to_date('" & bkmfrom & "','mm/dd/yyyy') and to_date('" & bkmto & "','mm/dd/yyyy') and h.fcba='" & fcba & "' " & _ 
        "    AND document_classification Like '" & pilihbkm & "%' " & _ 
			filtersection  & _
			filterrowstate  & _
        "    ORDER BY h.BKMDATE,h.SECTIONCODE,h.GANGCODE,h.VEHICLECODE,h.DOCUMENT_CLASSIFICATION"

    end if

    if 	pilihbkm= "505" then 
 
        sqlview =  "Select h.BKMNO, h.BKMDATE,h.SECTIONCODE,h.GANGCODE,nvl(h.VEHICLECODE,'-') VEHICLECODE, h.TOTUPAH TOTALBKMUPAH,h.TOTBAHAN TOTALBKMMATERIAL,h.TOTVRA TOTALBKMVEHICLE,h.TOTWIP TOTALBKMWIP,h.DOCUMENT_CLASSIFICATION " & _
        "    ,authorizestructure_api.get_template_desc(h.DOCUMENT_CLASSIFICATION) dcname,decode(h.hujan,'0','-','YA') hujan,decode(h.POSTINGFLG,'1','YA','TIDAK') POSTINGFLG" & _
        "    ,h.TOTUPAH + h.TOTBAHAN + h.TOTVRA + h.TOTWIP TOTAL,h.createdate INPUTDATE, h.LASTUPDATE, h.RKBNO,h.rowstate,h.FCENTRY,h.FCEDIT " & _	
        "    From   BKMHDR h WHERE h.DOCUMENT_CLASSIFICATION='505' and h.BKMDATE between to_date('" & bkmfrom & "','mm/dd/yyyy') and to_date('" & bkmto & "','mm/dd/yyyy') and h.fcba='" & fcba & "' " & _ 
        "    AND document_classification Like '" & pilihbkm & "%' " & _ 
      		filtersection  & _
			filterrowstate  & _
        "    ORDER BY h.BKMDATE,h.SECTIONCODE,h.GANGCODE,h.VEHICLECODE,h.DOCUMENT_CLASSIFICATION "
	ELSE
		sqlview =  "Select h.BKMNO, h.BKMDATE,CASE WHEN document_classification = '504' THEN contractor_api.get_contractorname(h.contractorcode,h.fcba) ELSE h.SECTIONCODE END SECTIONCODE,h.GANGCODE,nvl(h.VEHICLECODE,'-') VEHICLECODE, h.TOTUPAH TOTALBKMUPAH,h.TOTBAHAN TOTALBKMMATERIAL,h.TOTVRA TOTALBKMVEHICLE,h.TOTWIP TOTALBKMWIP,h.DOCUMENT_CLASSIFICATION " & _
        "    ,authorizestructure_api.get_template_desc(h.DOCUMENT_CLASSIFICATION) dcname,decode(h.hujan,'0','-','YA') hujan,decode(h.POSTINGFLG,'1','YA','TIDAK') POSTINGFLG" & _
        "    ,h.TOTUPAH + h.TOTBAHAN + h.TOTVRA + h.TOTWIP TOTAL,h.createdate INPUTDATE, h.LASTUPDATE, h.RKBNO,h.rowstate,h.FCENTRY,h.FCEDIT " & _	
        "    From   BKMHDR h WHERE h.DOCUMENT_CLASSIFICATION <> '505' and h.BKMDATE between to_date('" & bkmfrom & "','mm/dd/yyyy') and to_date('" & bkmto & "','mm/dd/yyyy') and h.fcba='" & fcba & "' " & _ 
        "    AND document_classification Like '" & pilihbkm & "%' " & _ 
            filtersection  & _
			filterrowstate  & _
        "    ORDER BY h.BKMDATE,h.SECTIONCODE,h.GANGCODE,h.VEHICLECODE,h.DOCUMENT_CLASSIFICATION "
	
    end if

'    response.write sqlview

    List.Open sqlview, connStr, 1
    ListCount = List.RecordCount 

    Set caljson = new JSON  'JSON.aspê°€ í•„ìš”í•¨
    Set rs = connStr.Execute(sqlview)

    Dim List_jsonString
    List_jsonString = caljson.toJSON("results",rs, false)

    Dim jsonObj, outputObj
    set jsonObj = new JSONobject
    set outputObj = jsonObj.parse(List_jsonString)  'json.objectClass.aspê°€ í•„ìš”í•¨(íŒŒì‹±í•˜ë ¤ë©´)


    Set oJSON = New aspJSON     'aspJSON.asp í•„ìš”í•¨(forë¬¸)
    oJSON.loadJSON(List_jsonString)

    Response.write List_jsonString
end if

'tampilkan data gl'
if idview="2" then
    dim tahungl,bulangl,coagl,coaakhirgl,bulantogl,blntgl
    fcba = session("fcba")
    tahungl = Request("tahun")
    bulangl = Request("bulan")
	bulantogl = Request("bulanto")
	
    coagl   = Request("fcoa_ap")
	coaakhirgl   = Request("coaakhir")
	companygl = Request("company")
	fcbagl    = Request("fcba")
	sourcecode    = Request("sourcecode")
	
    Set List = Server.CreateObject("ADODB.Recordset")

	if trim(fcbagl) = "-" then 
		filter = " AND H.FCBA in (SELECT fccode from businessunit where fccompanycode='" & companygl & "' and isactive='TRUE') "
	else
	    filter = " AND H.FCBA ='" & fcbagl & "' "
	end if

    if trim(coagl) <> "--ALL--" OR trim(coaakhirgl) <> "--ALL--"  then
        filter2 = " AND C.FCCODE BETWEEN '" & coagl & "' AND '" & coaakhirgl & "'"
    end if

    if trim(sourcecode) <> ""  then
        filter3 = " AND J.SOURCECODE='" & sourcecode & "'"
    end if
 
    if bulangl > 12 then 
    sqlview =  "SELECT DOCNO,JOURNALNO,JOURNALDATE,COA,COANAME,OPENING_BALANCE,DEBET,CREDIT,ENDING_BALANCE,FCBA,SOURCECODE, " & _
      " ( COALESCE(SUM(DEBET + CREDIT) OVER (PARTITION BY FCBA,COA ORDER BY FCBA,COA,DOCNO,JOURNALDATE,JOURNALNO), 0)) AS saldo_akhir " & _
      " FROM " & _
      " ( " & _
      " Select '-' DOCNO, '-' JOURNALNO, TRUNC(TO_DATE('12'||'/01/'||'" & tahungl & "','MM/DD/YYYY'),'MM') JOURNALDATE, H.COA,'SALDO AWAL BLN '||" & bulangl-1 & "  COANAME, " & _
      " 0 OPENING_BALANCE,  Decode(Sign(H.OPENING_BALANCE) ,1,H.OPENING_BALANCE,0) DEBET,  Decode(Sign(H.OPENING_BALANCE) ,-1,H.OPENING_BALANCE,0) CREDIT, H.OPENING_BALANCE ENDING_BALANCE, H.FCBA,'-' SOURCECODE " & _
	  "	From  GLEDGERHDR H " & _
		"	 INNER JOIN JOBMASTER C ON (H.COA=C.FCCODE) " & _
		"	 WHERE H.PERIOD_YEAR='" & tahungl & "' AND H.PERIOD_MONTH = '" & bulangl & "'" & _
			filter & _
			filter2 & _ 
		"	 UNION ALL " & _
        " Select G.DOCNO, G.JOURNALNO, G.JOURNALDATE, G.COA,C.FCNAME COANAME, H.OPENING_BALANCE, " & _
		"	 G.DEBET, G.CREDIT, G.ENDING_BALANCE, G.FCBA , J.SOURCECODE " & _
		"	 From   GLEDGERDTL G " & _
		"	 INNER JOIN   GLEDGERHDR H ON (G.DOCNO=H.DOCNO AND G.FCBA=H.FCBA) " & _
		"	 INNER JOIN JOBMASTER C ON (G.COA=C.FCCODE) " & _
		"    INNER JOIN BKMJOURNALHDR J ON (G.JOURNALNO=J.DOCNO AND G.FCBA=J.FCBA) " & _
		"	 WHERE H.PERIOD_YEAR='" & tahungl & "' AND H.PERIOD_MONTH BETWEEN '" & bulangl & "' AND '" & bulantogl & "' " & _
			filter & _
			filter2 & _ 
       ") " & _
		" ORDER BY FCBA,COA,DOCNO,JOURNALDATE,JOURNALNO" 
	ELSE
		sqlview =  "SELECT DOCNO,JOURNALNO,JOURNALDATE,COA,COANAME,OPENING_BALANCE,DEBET,CREDIT,ENDING_BALANCE,FCBA,SOURCECODE, " & _
      " ( COALESCE(SUM(DEBET + CREDIT) OVER (PARTITION BY FCBA,COA ORDER BY FCBA,COA,DOCNO,JOURNALDATE,JOURNALNO), 0)) AS saldo_akhir " & _
      " FROM " & _
      " ( " & _
      " Select '-' DOCNO, '-' JOURNALNO, TRUNC(TO_DATE('" & bulangl & "'||'/01/'||'" & tahungl & "','MM/DD/YYYY'),'MM')-1 JOURNALDATE, H.COA,'SALDO AWAL BLN '||" & bulangl - 1 & "  COANAME, " & _
      " 0 OPENING_BALANCE,  Decode(Sign(H.OPENING_BALANCE) ,1,H.OPENING_BALANCE,0) DEBET,  Decode(Sign(H.OPENING_BALANCE) ,-1,H.OPENING_BALANCE,0) CREDIT, H.OPENING_BALANCE ENDING_BALANCE, H.FCBA,'-' SOURCECODE " & _
	  "	From  GLEDGERHDR H " & _
		"	 INNER JOIN JOBMASTER C ON (H.COA=C.FCCODE) " & _
		"	 WHERE H.PERIOD_YEAR='" & tahungl & "' AND H.PERIOD_MONTH = '" & bulangl & "'" & _
			filter & _
			filter2 & _ 
		"	 UNION ALL " & _
        " Select G.DOCNO, G.JOURNALNO, G.JOURNALDATE, G.COA,C.FCNAME COANAME, H.OPENING_BALANCE, " & _
		"	 G.DEBET, G.CREDIT, G.ENDING_BALANCE, G.FCBA , J.SOURCECODE " & _
		"	 From   GLEDGERDTL G " & _
		"	 INNER JOIN   GLEDGERHDR H ON (G.DOCNO=H.DOCNO AND G.FCBA=H.FCBA) " & _
		"	 INNER JOIN JOBMASTER C ON (G.COA=C.FCCODE) " & _
		"    INNER JOIN BKMJOURNALHDR J ON (G.JOURNALNO=J.DOCNO AND G.FCBA=J.FCBA) " & _
		"	 WHERE H.PERIOD_YEAR='" & tahungl & "' AND H.PERIOD_MONTH BETWEEN '" & bulangl & "' AND '" & bulantogl & "' " & _
			filter & _
			filter2 & _ 
			filter3 & _ 
       ") " & _
		" ORDER BY FCBA,COA,DOCNO,JOURNALDATE,JOURNALNO" 
	
	end if

  	' response.write sqlview

    List.Open sqlview, connStr, 1
    ListCount = List.RecordCount 

    Set caljson = new JSON  'JSON.aspê°€ í•„ìš”í•¨
    Set rs = connStr.Execute(sqlview)

    Dim List_jsonStringgl
    List_jsonStringgl = caljson.toJSON("results",rs, false)

    Dim jsonObjgl, outputObjgl
    set jsonObjgl = new JSONobject
    set outputObjgl = jsonObjgl.parse(List_jsonStringgl)  'json.objectClass.aspê°€ í•„ìš”í•¨(íŒŒì‹±í•˜ë ¤ë©´)


    Set oJSON = New aspJSON     'aspJSON.asp í•„ìš”í•¨(forë¬¸)
    oJSON.loadJSON(List_jsonStringgl)

    Response.write List_jsonStringgl
end if

'tampilkan data glhdr'
if idview="3" then
    dim tahunglhdr,bulanglhdr,coaglhdr,coaakhirglhdr,bulantoglhdr,tbl
    fcba = session("fcba")
    tahunglhdr = Request("tahun")
    bulanglhdr = Request("bulan")
	bulantoglhdr = Request("bulanto")	
    coaglhdr   = Request("coa")
	coaakhirglhdr   = Request("coaakhir")
	companyhdr = Request("company")
	fcbahdr    = Request("fcba")

	if (cint(bulan6)=13 or cint(bulan6)=14) then
		tbl = "FAC_TRIALBALANCE_SPT_ADT"
	ELSE
		tbl = "FAC_TRIALBALANCE"
	end if

    Set List = Server.CreateObject("ADODB.Recordset")

	if trim(fcbahdr) = "-" then 
		filter = " AND G.FCBA in (SELECT fccode from businessunit where fccompanycode='" & companyhdr & "' and isactive='TRUE') "
	else
	    filter = " AND G.FCBA ='" & fcbahdr & "' "
	end if

    if trim(coaglhdr) <> "--ALL--" OR trim(coaakhirglhdr) <> "--ALL--"  then
        filter2 = " AND FCCODE BETWEEN '" & coaglhdr & "' AND '" & coaakhirglhdr & "'"
    end if

    sqlview =  " SELECT SUM(OPENING_BALANCE) OPENING_BALANCE,SUM(DEBET) DEBET, SUM(CREDIT) CREDIT, SUM(OPENING_BALANCE)+ SUM(DEBET) + SUM(CREDIT) ENDING_BALANCE " & _ 
			   " FROM	" & _ 
	           " ( " & _ 
			   " Select SUM(G.OPENING_BALANCE) OPENING_BALANCE, 0 DEBET,0 CREDIT , 0 ENDING_BALANCE " & _ 
			   " From " & tbl & " G WHERE PERIOD_MONTH = '" & bulanglhdr & "' AND PERIOD_YEAR='" & tahunglhdr & "' " & _ 
			   filter & _ 
			   filter2 & _
			   " UNION ALL " & _
	           " SELECT 0 OPENING_BALANCE, SUM(G.DEBET) DEBET, SUM(G.CREDIT) CREDIT,0 ENDING_BALANCE " & _ 
			   " From  " & tbl & " G " & _
			   " WHERE PERIOD_MONTH BETWEEN '" & bulanglhdr & "' AND '" & bulantoglhdr & "' AND PERIOD_YEAR='" & tahunglhdr & "' " & _
			   filter & _ 
			   filter2 & _ 
			   " )"
			   '" GROUP BY PERIOD_MONTH,PERIOD_YEAR "

   ' response.write sqlview

    List.Open sqlview, connStr, 1
    ListCount = List.RecordCount 

    Set caljson = new JSON  'JSON.aspê°€ í•„ìš”í•¨
    Set rs = connStr.Execute(sqlview)

    Dim List_jsonStringglhdr
    List_jsonStringglhdr = caljson.toJSON("results",rs, false)

    Dim jsonObjglhdr, outputObjglhdr
    set jsonObjglhdr = new JSONobject
    set outputObjglhdr = jsonObjglhdr.parse(List_jsonStringglhdr)  'json.objectClass.aspê°€ í•„ìš”í•¨(íŒŒì‹±í•˜ë ¤ë©´)


    Set oJSON = New aspJSON     'aspJSON.asp í•„ìš”í•¨(forë¬¸)
    oJSON.loadJSON(List_jsonStringglhdr)

    Response.write List_jsonStringglhdr
end if

'tampilkan data lhp'
if idview="4" then
   
    fcba = session("fcba")
    bkmfrom = Request("bkmfrom")
    bkmto   = Request("bkmto")


    Set List = Server.CreateObject("ADODB.Recordset")


    sqlview =  "Select LHPNO,PRODDATE, TBSTERIMA, TBSSISA,TBSOLAH, CPOKIRIM, " & _ 
       " CPOCOT, CPOCST, CPOOT,  KERNELKIRIM, KERNELSILO, CANGKANGKIRIM, CANGKANGPRODUKSI,  " & _
       " CPOTANKISI, PKTANKISI, POTONGANTBS, ADJCPO " & _
       " From   PRODUCTIONTRN  " & _
       " WHERE PRODDATE BETWEEN to_date('" & bkmfrom & "','mm/dd/yyyy') and to_date('" & bkmto & "','mm/dd/yyyy') and fcba='" & fcba & "' " & _
       " order by proddate ASC"

    'response.write sqlview

    List.Open sqlview, connStr, 1
    ListCount = List.RecordCount 

    Set caljson = new JSON  'JSON.aspê°€ í•„ìš”í•¨
    Set rs = connStr.Execute(sqlview)

    Dim List_jsonStringgllhp
    List_jsonStringgllhp = caljson.toJSON("results",rs, false)

    Dim jsonObjgllhp, outputObjgllhp
    set jsonObjgllhp = new JSONobject
    set outputObjgllhp = jsonObjgllhp.parse(List_jsonStringgllhp)  'json.objectClass.aspê°€ í•„ìš”í•¨(íŒŒì‹±í•˜ë ¤ë©´)


    Set oJSON = New aspJSON     'aspJSON.asp í•„ìš”í•¨(forë¬¸)
    oJSON.loadJSON(List_jsonStringgllhp)

    Response.write List_jsonStringgllhp
end if

'tampilkan data journal'
if idview="5" then
    dim tahung5,bulang5,sourcegl5,filter,jrnfrom,jrnto,filterprd
	jrnfrom = Request("jrnfrom")
    jrnto   = Request("jrnto")
	tglprd  = Request("tglprd")
    tahun5 = Request("tahun")
    bulan5 = Request("bulan")
	source5= Request("sourcecode")
	groupaccount5= Request("groupaccount")
	company5 = Request("company")
	fcba5   = Request("fcba")
	posting = Request("posting")
    

    Set List = Server.CreateObject("ADODB.Recordset")
	
	if trim(fcba5) = "-" then 
		filter = " AND G.FCBA in (SELECT fccode from businessunit where fccompanycode='" & company5 & "' and isactive='TRUE') "
	else
	    filter = " AND G.FCBA ='" & fcba5 & "' "
	end if

	if trim(groupaccount5) = "-" then 
		filtergroupaccount5 = "  "
	else
	    filtergroupaccount5 = " AND C.jobgroupcode like '%" & groupaccount5 & "%' "
	end if

    if trim(tglprd) = "Tgl" THEN
		filterprd = " (G.DOCDATE between to_date('" & jrnfrom & "','mm/dd/yyyy') and to_date('" & jrnto & "','mm/dd/yyyy')) " 			   
	else
		filterprd = " Accountingyear='" & tahun5 & "' " & _ 
					" AND Accountingperiod ='" & bulan5 & "' " 		   
	end if

    sqlview =  "Select G.docno, G.docdate, nvl(G.remarks,'-') remarks, G.sourcecode, " & _ 
	           " round(Sum(NVL(H.amountdebit,0)),2) debet,round(Sum(NVL(H.amountkredit,0)),2) kredit , G.FCBA, C.jobgroupcode, D.fcname AS jobgroupname, Decode(G.POSTINGFLG,1,'YA','TIDAK') POSTINGFLG, G.Accountingperiod, G.Accountingyear " & _
			   " From  BKMJOURNALDTL H " & _
			   " INNER JOIN BKMJOURNALHDR G ON (G.DOCNO=H.DOCNO AND G.FCBA=H.FCBA) " & _
			   "LEFT JOIN JOBMASTER C ON H.JOBCODE = C.FCCODE " & _             
			   "LEFT JOIN JOBGROUPMASTER D ON C.JOBGROUPCODE = D.FCCODE " & _    
			   " WHERE " & _ 
			   filterprd & _ 
			   filter & _
			   " AND G.SOURCECODE Like '%" & source5 & "%'" & _
			   " AND G.POSTINGFLG like '%" & posting & "%' " & _
			   filtergroupaccount5 & _
			   " GROUP BY G.docno, G.docdate,G.remarks,G.sourcecode, G.FCBA, C.jobgroupcode, D.fcname, G.POSTINGFLG,G.Accountingperiod, G.Accountingyear " & _
			   " ORDER BY G.fcba,G.DOCDATE,G.sourcecode,G.docno " 


    ' response.write sqlview

    '---- Perbaikan performa: query dulu dijalankan dua kali (List.Open hanya untuk RecordCount yang
    '---- tidak dipakai) dan JSON hasilnya di-parse dua kali lagi (JSONobject & aspJSON) tanpa dipakai.
    '---- Dengan data ribuan baris halaman melewati batas waktu (ASP 0113). Output JSON tetap sama.
    Set caljson = new JSON
    Set rs = connStr.Execute(sqlview)

    Dim List_jsonString5
    List_jsonString5 = caljson.toJSON("results",rs, false)

    Response.write List_jsonString5
end if


'tampilkan data journal' TEST
if idview="6" then
   ' dim tahun6,bulan6,filter,company6,fcba6
   dim coaawal,coaakhir,filter2
    tahun6 = Request("tahun")
    bulan6 = Request("bulan")
	company6 = Request("company")
	fcba6 = Request("fcba")
	coaawal = Request("akunawal")
	coaakhir = Request("akunakhir")
	
	if (cint(bulan6)=13 or cint(bulan6)=14) then
		tbl = "FAC_TRIALBALANCE_SPT_ADT"
	ELSE
		tbl = "FAC_TRIALBALANCE"
	end if
    
  

    Set List = Server.CreateObject("ADODB.Recordset")

   
	
	if trim(company6) <> "" then 
		if trim(fcba6) = "-" then 
			filter = " AND FCBA in (SELECT fccode from businessunit where fccompanycode='" & company6 & "' and isactive='TRUE')"
		else
			filter = " AND FCBA ='" & fcba6 & "' "
		end if
  
	else
	       filter = " AND FCBA in (SELECT fccode from businessunit where isactive='TRUE')"
    End IF

     
    if trim(coaawal) <> "--ALL--" OR trim(coaakhir) <> "--ALL--"  then
                filter2 = " AND a.FCCODE BETWEEN '" & coaawal & "' AND '" & coaakhir & "'"
    end if
   

    

    sqlview =  "Select a.FCCODE,a.FCNAME,c.fcname as jobgroupname,sum(opening_balance) opening_balance ,sum(debet) debet,sum(credit) credit ,sum(opening_balance)+sum(debet)+sum(credit) ending_balance,  " & _
			   " jobgroupcode,job_forbalancesheetcode,job_forprofitandlosscode,job_forcashflowcode,job_forbudgetcode " & _
			   " From " & tbl & " a " & _
			   " LEFT JOIN jobmaster b ON (a.fccode = b.fccode ) " & _
			   " LEFT JOIN jobgroupmaster c ON (c.FCCODE = b.JOBGROUPCODE ) " & _
			   " WHERE period_year=" & tahun6 & " " & _ 
			   " AND period_month=" & bulan6 & " " & _ 
			   filter & _
               filter2 & _
			   " GROUP BY period_month,period_year,a.FCCODE,a.FCNAME,c.fcname, jobgroupcode,job_forbalancesheetcode,job_forprofitandlosscode,job_forcashflowcode,job_forbudgetcode  " & _ 
			   " ORDER BY a.fccode " 
 
'    response.write sqlview

    'List.Open sqlview, connStr, 1
    'ListCount = List.RecordCount 

    Set caljson = new JSON  'JSON.aspê°€ í•„ìš”í•¨
    Set rs = connStr.Execute(sqlview)

    Dim List_jsonString6
    List_jsonString6 = caljson.toJSON("results",rs, false)

    Dim jsonObj6, outputObj6
    set jsonObj6 = new JSONobject
    set outputObj6 = jsonObj6.parse(List_jsonString6)  'json.objectClass.aspê°€ í•„ìš”í•¨(íŒŒì‹±í•˜ë ¤ë©´)


    Set oJSON = New aspJSON     'aspJSON.asp í•„ìš”í•¨(forë¬¸)
    oJSON.loadJSON(List_jsonString6)

    Response.write List_jsonString6
end if

'tampilkan data tipe overtime hours'
if idview="7" then
  

    Set List = Server.CreateObject("ADODB.Recordset")
	

    sqlview =  "select distinct type id,fcname text" & _ 
    " from overtime_convertion " & _
    " where fcba='" & session("fcba") & "' " & _
    " order by type asc " 


    Set caljson = new JSON  'JSON.aspê°€ í•„ìš”í•¨
    Set rs = connStr.Execute(sqlview)

    Dim List_jsonString7
    List_jsonString7 = caljson.toJSON("results",rs, false)

    Dim jsonObj7, outputObj7
    set jsonObj7 = new JSONobject
    set outputObj7 = jsonObj7.parse(List_jsonString7)  'json.objectClass.aspê°€ í•„ìš”í•¨(íŒŒì‹±í•˜ë ¤ë©´)


    Set oJSON = New aspJSON     'aspJSON.asp í•„ìš”í•¨(forë¬¸)
    oJSON.loadJSON(List_jsonString7)

    Response.write List_jsonString7
end if

'tampilkan data rkb'
if idview="8" then
   ' dim tahun6,bulan6,filter,company6,fcba6
    tahun8 = Request("tahun")
    bulan8 = Request("bulan")
	company8 = Request("company")
	fcba8 = Request("fcba")
	type8 = Request("type")
	
    

    Set List = Server.CreateObject("ADODB.Recordset")
	
	if trim(fcba8) = "-" then 
		filter = " AND FCBA in (SELECT fccode from businessunit where fccompanycode='" & company8 & "' and isactive='TRUE') "
	else
	    filter = " AND FCBA ='" & fcba8 & "' "
	end if


    sqlview =  " Select FCBA,RKBNO,PERIOD_YEAR||'/'||PERIOD_MONTH period,SECTIONCODE," & _
	           " case document_classification " & _ 
			   " when '501' then 'Perawatan' " & _ 
			   " when '503' then 'Panen' " & _ 
			   " when '502' then 'Vehicle' " & _ 
			   " when '504' then 'Wip' End type, sum(TOTUPAH) TOTUPAH ,sum(TOTBAHAN) TOTBAHAN,SUM(TOTVRA) TOTVRA ,sum(TOTWIP) TOTWIP,ROWSTATE " & _
			   " From rkbhdr " & _
			   " WHERE period_year=" & tahun8 & " " & _ 
			   " AND period_month=" & bulan8 & " " & _ 
			   " AND document_classification like '" & trim(type8) & "%'" & _
			   filter & _
			   " GROUP BY FCBA,RKBNO,PERIOD_YEAR||'/'||PERIOD_MONTH,SECTIONCODE,document_classification,ROWSTATE " & _ 
			   " ORDER BY FCBA,period,document_classification,RKBNO " 

    'response.write sqlview

    'List.Open sqlview, connStr, 1
    'ListCount = List.RecordCount 

    Set caljson = new JSON  'JSON.aspê°€ í•„ìš”í•¨
    Set rs = connStr.Execute(sqlview)

    Dim List_jsonString8
    List_jsonString8 = caljson.toJSON("results",rs, false)

    Dim jsonObj8, outputObj8
    set jsonObj8 = new JSONobject
    set outputObj8 = jsonObj8.parse(List_jsonString8)  'json.objectClass.aspê°€ í•„ìš”í•¨(íŒŒì‹±í•˜ë ¤ë©´)


    Set oJSON = New aspJSON     'aspJSON.asp í•„ìš”í•¨(forë¬¸)
    oJSON.loadJSON(List_jsonString8)

    Response.write List_jsonString8
end if
'tampilkan data store'
if idview="9" then
   ' dim tahun6,bulan6,filter,company6,fcba6
    dt = Request("dt")
    company9 = Request("company")
	fcba9 = Request("fcba")
	type9 = Request("type")
	month9 = Request("month")
	year9 = Request("year")
	src9 = UCase(Request("src"))
	store = Request("storecode")
	sub9  = Request("sub")
	
    Set List = Server.CreateObject("ADODB.Recordset")
	

	filter = " AND YEAR=" & year9 & " AND MONTH=" & month9 & " AND FCBA ='" & fcba9 & "' AND stock_api.get_stockgroup(itemcode,fcba) Like '" & type9 & "%' " & _ 
		         " AND NVL(stock_api.get_StockSubGroup(itemcode,fcba),'-') Like '" & sub9 & "%' " 
				
	sqlview   =   " SELECT " & _
				  " fcba, " & _
				  " storecode, " & _
				  " itemcode, " & _
				  " stock_api.get_stockgroup(itemcode,fcba) StockGroupcode, " & _
				  " stockgroup_api.get_stockgrpname(stock_api.get_stockgroup(itemcode,fcba),fcba) stockgroupname, " & _
				  " nvl(stocksubgroup_api.get_stocksubgrpname(stock_api.get_StockSubGroup(itemcode,fcba)),'-') stocksubgroupname, " & _
				  " stock_api.get_ActiveMatrial(itemcode,fcba) ActMat , " & _
				  " stock_api.get_stockname(itemcode, fcba) fcname,  " & _
				  " stock_api.get_uom(itemcode, fcba) unitofmeasurement,  " & _
				  " Nvl(openingbalanceqty,0) openingbalanceqty, " & _
				  " Nvl(qtyreceived,0) qtyreceived, " & _
				  " Nvl(qtyreceivedtask,0) qtyreceivedtask, " & _
				  " Nvl(qtyreceivetaskdiv,0) qtyreceivetaskdiv, " & _
				  " Nvl(qtytransfered,0) qtytransfered, " & _
				  " Nvl(qtytransporttask,0) qtytransporttask, " & _
				  " Nvl(qtytransporttaskdiv,0) qtytransporttaskdiv, " & _
				  " Nvl(qtyreturned,0) qtyreturned, " & _
				  " Nvl(qtyissued,0) qtyissued, " & _
				  " Nvl(qtywriteoff,0) qtywriteoff, " & _
				  " Nvl(endingbalanceqty,0) endingbalance, " & _
				  " Round(Nvl(amtperunit,0),2) amtperunit, " & _
				  " Decode(storecode,'ST1',stock_api.get_stockminimum(itemcode, fcba),0) get_stockminimum  " & _
				  " FROM " & _
				  " stockledger " & _ 
				  " WHERE 1=1 " & _
				  filter & _ 
				  " AND ((itemcode like '%" & src9 & "%') or (stock_api.get_stockname(itemcode, fcba) like '%" & src9 & "%')) " & _ 
				  " AND stock_api.get_stockgroup(itemcode,fcba) IN (select distinct fccode from stockgroup where grouptype= 'STOCK') " & _
				  " AND storecode like '%" & store & "%' " & _
				  " Order by fcba,StockGroupname,itemcode,storecode "



    'response.write sqlview

    'List.Open sqlview, connStr, 1
    'ListCount = List.RecordCount 

    Set caljson = new JSON  'JSON.aspê°€ í•„ìš”í•¨
    Set rs = connStr.Execute(sqlview)

    Dim List_jsonString9
    List_jsonString9 = caljson.toJSON("results",rs, false)

    Dim jsonObj9, outputObj9
    set jsonObj9 = new JSONobject
    set outputObj9 = jsonObj9.parse(List_jsonString9)  'json.objectClass.aspê°€ í•„ìš”í•¨(íŒŒì‹±í•˜ë ¤ë©´)


    Set oJSON = New aspJSON     'aspJSON.asp í•„ìš”í•¨(forë¬¸)
    oJSON.loadJSON(List_jsonString9)

    Response.write List_jsonString9
end if
'tampilkan data hpp harian'
if idview="10" then
   
 	company10 = Request("company")
	fcba10 = Request("fcba")
	
    bkmfrom10 = Request("bkmfrom")
    bkmto10   = Request("bkmto")


    Set List = Server.CreateObject("ADODB.Recordset")
	
	sqllist =   " SELECT TO_CHAR(TO_DATE('" & bkmfrom10 & "','MM/DD/YYYY') - 1 + rownum,'MM/DD/YYYY') AS dt " & _
				" FROM all_objects " & _
				" WHERE TO_DATE('" & bkmfrom10 & "','MM/DD/YYYY') - 1 + rownum <= TO_DATE('" & bkmto10 & "', 'MM/DD/YYYY') "
			
			Set rslistdt= Server.CreateObject("ADODB.Recordset")
			rslistdt.CursorType = 3
			rslistdt.CursorLocation = 3
			rslistdt.Open sqllist, connstr
			if not rslistdt.eof then 
			sqlexec = "BEGIN " & vbCrLf 
			  Do While Not rslistdt.eof	
				sqlexec = sqlexec & "fac_hpp_harian_api.ProsesData(to_date('" & rslistdt("dt") & "','mm/dd/yyyy'),'" & fcba10 & "');"  & vbCrLf 
				rslistdt.MoveNext
			  loop
			End if
			sqlexec = sqlexec & "END; " 
			rslistdt.Close
			set rslistdt = Nothing			  


   ' sqlexec = "BEGIN " & _ 
    '          "fac_hpp_harian_api.ProcessHPP_Daily(to_date('" & bkmfrom10 & "','mm/dd/yyyy') , to_date('" & bkmto10 & "','mm/dd/yyyy') ,'" & fcba10 & "');" & _
	'		  "END; " 
    'response.write sqlexec
	connStr.Execute(sqlexec)

    sqlview =  " Select proddate,tbsterima_int,tbsterima_ext,tbsterima,tbsolah,tbssisa," & _
	   " cpo_prod_hi,cpo_kirim_hi,cpo_sisa_hi," & _
	   " pk_prod_hi,pk_kirim_hi,pk_sisa_hi," & _
	   " tbsterima_int_amnt,tbsterima_ext_amnt, prod_tbs_amnt,prod_olah_amnt,prod_total_amnt,prod_avr_amnt," & _ 
	   " hpp_cpo_amnt,hpp_cpo_rpkg," & _
	   " hpp_pk_amnt,hpp_pk_rpkg " & _
	   " From   fac_hpp_harian  " & _
       " WHERE PRODDATE BETWEEN to_date('" & bkmfrom10 & "','mm/dd/yyyy') and to_date('" & bkmto10 & "','mm/dd/yyyy') and fcba='" & fcba10 & "' " & _
       " order by proddate ASC "

    'response.write sqlview

    List.Open sqlview, connStr, 1
    ListCount = List.RecordCount 

    Set caljson = new JSON  'JSON.aspê°€ í•„ìš”í•¨
    Set rs = connStr.Execute(sqlview)

    Dim List_jsonString10
    List_jsonString10 = caljson.toJSON("results",rs, false)

    Dim jsonObj10, outputObj10
    set jsonObj10 = new JSONobject
    set outputObj10 = jsonObj10.parse(List_jsonString10)  'json.objectClass.aspê°€ í•„ìš”í•¨(íŒŒì‹±í•˜ë ¤ë©´)


    Set oJSON = New aspJSON     'aspJSON.asp í•„ìš”í•¨(forë¬¸)
    oJSON.loadJSON(List_jsonString10)

    Response.write List_jsonString10
end if
'tampilkan data hpp harian'
if idview="11" then
   
    tahun11 = Request("tahun")
    bulan11 = Request("bulan")
	company11 = Request("company")
	fcba11 = Request("fcba")
	type11 = Request("type")
	
	if trim(fcba9) = "-" then 
		filter = " AND a.FCBA in (SELECT fccode from businessunit where fccompanycode='" & company9 & "' and isactive='TRUE') AND a.fccode Like '" & type11 & "%' "
	else
	    filter = " AND a.FCBA ='" & fcba11 & "' AND a.fccode Like '" & type11 & "%' "
	end if
	
    sqlview =   " SELECT a.fcba, " & _
       " a.fccode, " & _
       " a.fcname, " & _
	   " a.CONTROLJOB accountcode,"  & _
       " Nvl(nopening, 0)                                    nOpening, " & _
       " Nvl(ndebet, 0)                                      nDebet, " & _
       " Nvl(ncredit, 0)                                     nCredit, " & _
       " Nvl(nopening, 0) + Nvl(ndebet, 0) - Nvl(ncredit, 0) AS nEnding " & _
	   " FROM   bank a " & _
		"	left join (SELECT fcba, " & _
		"						period_year, " & _
		"						period_month            period_month, " & _
		"						fccode," & _
		"						Nvl(opening_balance, 0) nOpening " & _
		"				FROM   fac_trialbalance " & _
		"				WHERE  period_year = " & tahun11 & " " & _
		"						AND period_month = " & bulan11 & ") b " & _
		"			ON ( a.fcba = b.fcba " & _
		"				AND Trim(a.accountcode) = Trim(b.fccode) ) " & _
		"	left join (SELECT fcba, " & _
		"						period_year, " & _
		"						period_month, " & _
		"						bankcode, " & _
		"						SUM(ndebet) nDebet " & _
		"				FROM   (SELECT fcba, " & _
		"								To_number(To_char(datecreated, 'YYYY')) " & _
		"								period_year, " & _
		"								To_number(To_char(datecreated, 'MM')) " & _
		"								period_month, " & _
		"								bankcode, " & _
		"								SUM(totalamount)                        nDebet " & _
		"						FROM   paymentvoucher a " & _
		"						WHERE  rowstateid NOT IN ( 'Rejected', 'Planned' ) " & _
		"								AND transtype = '01-RECEIVE' " & _
		"								AND To_number(To_char(datecreated, 'YYYY')) = " & tahun11 & " " & _
		"								AND To_number(To_char(datecreated, 'MM')) = " & bulan11 & " " & _
		"						GROUP  BY fcba, " & _
		"									To_number(To_char(datecreated, 'YYYY')), " & _
		"									To_number(To_char(datecreated, 'MM')), " & _
		"									bankcode " & _
		"						UNION ALL " & _
		"						SELECT chargeunit                              fcba, " & _
		"								To_number(To_char(datecreated, 'YYYY')) " & _
		"								period_year, " & _
		"								To_number(To_char(datecreated, 'MM')) " & _
		"								period_month, " & _
		"								chargecode                              bankcode, " & _
		"								SUM(Nvl(b.amount, 0))                  nDebet " & _
		"						FROM   paymentvoucher a " & _
		"								inner join paymentvoucherdetail b " & _
		"										ON ( a.vouchercode = b.vouchercode " & _
		"											AND a.fcba = b.fcba ) " & _
		"						WHERE  chargeunit IS NOT NULL " & _
		"								AND rowstateid NOT IN ( 'Rejected', 'Planned' ) " & _
		"								AND transtype = '02-PAYMENT' " & _
		"								AND To_number(To_char(datecreated, 'YYYY')) = " & tahun11 & " " & _
		"								AND To_number(To_char(datecreated, 'MM')) = " & bulan11 & " " & _
		"								AND chargejob IN (SELECT DISTINCT controljob " & _
		"												FROM   bank " & _
		"												WHERE  activation = 'Y') " & _
		"								AND EXISTS (SELECT 1 " & _
		"											FROM   bkmjournalhdr c " & _
		"											WHERE  Trim(a.vouchercode) || '-INT' = " & _
		"													Trim(c.docno)) " & _
		"						GROUP  BY chargeunit, " & _
		"									To_number(To_char(datecreated, 'YYYY')), " & _
		"									To_number(To_char(datecreated, 'MM')), " & _
		"									chargecode) x " & _
		"				GROUP  BY fcba, " & _
		"							period_year, " & _
		"							period_month, " & _
		"							bankcode) c " & _
		"			ON ( a.fcba = c.fcba " & _
		"				AND Trim(a.fccode) = Trim(c.bankcode) ) " & _
		"	left join (SELECT fcba, " & _
		"						To_number(To_char(datecreated, 'YYYY')) period_year, " & _
		"						To_number(To_char(datecreated, 'MM'))   period_month, " & _
		"						bankcode, " & _
		"						SUM(totalamount)                        nCredit " & _
		"				FROM   paymentvoucher a " & _
		"				WHERE  rowstateid NOT IN ( 'Rejected', 'Planned' ) " & _
		"						AND transtype = '02-PAYMENT' " & _
		"						AND To_number(To_char(datecreated, 'YYYY')) = " & tahun11 & " " & _
		"						AND To_number(To_char(datecreated, 'MM')) = " & bulan11 & " " & _
		"				GROUP  BY fcba, " & _
		"							To_number(To_char(datecreated, 'YYYY')), " & _
		"							To_number(To_char(datecreated, 'MM')), " & _
		"							bankcode) d " & _
		"			ON ( a.fcba = d.fcba " & _
		"				AND a.fccode = d.bankcode ) " & _
		" WHERE  a.activation = 'Y' " & _
		"	AND Upper(banktype) IN ( 'CASH', 'BANK' ) " & _ 
				filter & _
		" ORDER  BY a.fcba, " & _
		"		a.fccode "


    'response.write sqlview
    Set List = Server.CreateObject("ADODB.Recordset")
    List.Open sqlview, connStr, 1
    ListCount = List.RecordCount 

    Set caljson = new JSON  'JSON.aspê°€ í•„ìš”í•¨
    Set rs = connStr.Execute(sqlview)

    Dim List_jsonString11
    List_jsonString11 = caljson.toJSON("results",rs, false)

    Dim jsonObj11, outputObj11
    set jsonObj11 = new JSONobject
    set outputObj11 = jsonObj11.parse(List_jsonString11)  'json.objectClass.aspê°€ í•„ìš”í•¨(íŒŒì‹±í•˜ë ¤ë©´)


    Set oJSON = New aspJSON     'aspJSON.asp í•„ìš”í•¨(forë¬¸)
    oJSON.loadJSON(List_jsonString11)

    Response.write List_jsonString11
end if

'tampilkan data hpp bulanan'
if idview="12" then
   
    tahun12 = Request("hppperiod")
	company12 = Request("company")
	fcba12 = Request("fcba")
	
	sqlview =  	" SELECT a.fcba,a.description,a.accountingyear,a.accountingperiod,TO_CHAR(TO_DATE(accountingperiod, 'MM'), 'MONTH') period," & _
				" TBS_AWAL_QTY,round(TBS_AWAL_AMT,2) TBS_AWAL_AMT," & _
				" TBS_TERIMA_INT_QTY,round(TBS_TERIMA_INT_AMT,2) TBS_TERIMA_INT_AMT," & _
				" TBS_TERIMA_EXT_QTY,round(TBS_TERIMA_EXT_AMT,2) TBS_TERIMA_EXT_AMT," & _
				" TBS_TERIMA_QTY,round(TBS_TERIMA_AMT,2) TBS_TERIMA_AMT," & _
				" TBS_OLAH_QTY,round(TBS_OLAH_AMT,2) TBS_OLAH_AMT," & _
				" TBS_AKHIR_QTY,round(TBS_AKHIR_AMT,2) TBS_AKHIR_AMT," & _
				" COST_PROD_QTY,round(COST_PROD_AMT,2) COST_PROD_AMT," & _
				" COST_BTL_QTY,round(COST_BTL_AMT,2) COST_BTL_AMT," & _
				" COST_ASSET_QTY,round(COST_ASSET_AMT,2) COST_ASSET_AMT," & _
				" COST_TOT_QTY,round(COST_TOT_AMT,2)-round(COST_ASSET_AMT,2) COST_TOT_AMT ," & _
				" COST_HPP_QTY,round(COST_HPP_AMT,2) COST_HPP_AMT,round(COST_HPP_RPKG,2) COST_HPP_RPKG," & _
				" CPO_AWAL_QTY, round(CPO_AWAL_AMT,2) CPO_AWAL_AMT ," & _
				" CPO_PROSES_QTY,round(CPO_PROSES_AMT,2) CPO_PROSES_AMT," & _
				" CPO_JUAL_QTY,round(CPO_JUAL_AMT,2) CPO_JUAL_AMT," & _
				" CPO_BELI_QTY,round(CPO_BELI_AMT,2) CPO_BELI_AMT," & _
				" CPO_STOCK_QTY,round(CPO_STOCK_AMT,2) CPO_STOCK_AMT ," & _
				" CPO_TOT_QTY,round(CPO_TOT_AMT,2) CPO_TOT_AMT," & _
				" CPO_JALAN_AWL_QTY,round(CPO_JALAN_AWL_AMT,2) CPO_JALAN_AWL_AMT," & _
				" CPO_AKHIR_QTY,round(CPO_AKHIR_AMT,2) CPO_AKHIR_AMT," & _
				" CPO_HPP_QTY,round(CPO_HPP_AMT,2) CPO_HPP_AMT,round(CPO_HPP_RPKG,2) CPO_HPP_RPKG," & _
				" PK_AWAL_QTY,round(PK_AWAL_AMT,2) PK_AWAL_AMT," & _
				" PK_PROSES_QTY,round(PK_PROSES_AMT,2) PK_PROSES_AMT," & _
				" PK_JUAL_QTY,round(PK_JUAL_AMT,2) PK_JUAL_AMT," & _
				" PK_BELI_QTY,round(PK_BELI_AMT,2) PK_BELI_AMT," & _
				" PK_STOCK_QTY,round(PK_STOCK_AMT,2) PK_STOCK_AMT," & _
				" PK_TOT_QTY,round(PK_TOT_AMT,2) PK_TOT_AMT," & _
				" PK_JALAN_AWL_QTY,round(PK_JALAN_AWL_AMT,2) PK_JALAN_AWL_AMT ," & _
				" PK_AKHIR_QTY,round(PK_AKHIR_AMT,2) PK_AKHIR_AMT," & _
				" PK_HPP_QTY,round(PK_HPP_AMT,2) PK_HPP_AMT," & _
				" round(PK_HPP_RPKG,2) PK_HPP_RPKG," & _
				" CPO_KIRIM_QTY,PK_KIRIM_QTY," & _
				" CPO_JALAN_AHR_QTY,PK_JALAN_AHR_QTY, " & _
				" CPO_JALAN_AHR_AMT,PK_JALAN_AHR_AMT " & _
				" ,c.docno,c.docdate " & _
				" FROM accountingyearperiod a left JOIN " & _
				" fac_hpp_bulanan b ON (Trim(a.fcba)=Trim(b.fcba) AND a.accountingperiod=b.period_MONTH AND a.accountingyear=b.period_year) " & _
				" left JOIN (SELECT DISTINCT fcba,period_month,period_year,docno,docdate FROM FAC_HPP_BULANAN_JOURNAL) c ON " & _
				" (Trim(a.fcba)=Trim(c.fcba) AND a.accountingperiod=c.period_MONTH AND a.accountingyear=c.period_year) " & _
				" WHERE a.fcba='" & fcba12 & "' AND accountingyear=" & tahun12 & "  AND accountingperiod <= 12 " & _
				" ORDER BY accountingperiod "


    'response.write sqlview
    Set List = Server.CreateObject("ADODB.Recordset")
    List.Open sqlview, connStr, 1
    ListCount = List.RecordCount 

    Set caljson = new JSON  'JSON.aspê°€ í•„ìš”í•¨
    Set rs = connStr.Execute(sqlview)

    Dim List_jsonString12
    List_jsonString12 = caljson.toJSON("results",rs, false)

    Dim jsonObj12, outputObj12
    set jsonObj12 = new JSONobject
    set outputObj12 = jsonObj12.parse(List_jsonString12)  'json.objectClass.aspê°€ í•„ìš”í•¨(íŒŒì‹±í•˜ë ¤ë©´)


    Set oJSON = New aspJSON     'aspJSON.asp í•„ìš”í•¨(forë¬¸)
    oJSON.loadJSON(List_jsonString12)

    Response.write List_jsonString12
end if

'tampilkan data fixed asset'
if idview="13" then
   
    tahun13 = Request("tahun")
    bulan13 = Request("bulan")
	company13 = Request("company")
	fcba13 = Request("fcba")
	type13 = Request("type")
	
	if trim(fcba13) = "-" then 
		filter = " AND a.FCBA in (SELECT fccode from businessunit where fccompanycode='" & company13 & "' and isactive='TRUE') AND assetgroupcode Like '" & type13 & "%' "
	else
	    filter = " AND a.FCBA ='" & fcba13 & "' AND assetgroupcode Like '" & type13 & "%' "
	end if
	
    sqlview =  " SELECT a.fcba,a.fccode,fcname,transdate,assetgroupcode,assetgroupname,depreciationmonth,acquisitionvalue,additionalvalue," & _ 
	           " NVL(total,0) total, nvl(monthlydepreciation,0) amountdepreciation, nvl(a.totaldepreciation,0) totaldepreciation,nvl(a.bookvalue,0) bookvalue,decode(nvl(b.postingflg,1),0,'Tidak','Ya') posting " & _ 
	           " FROM fixedasset a LEFT JOIN Fixedasset_Depreciation_Ledger b ON (a.fcba=b.fcba AND a.fccode=b.fccode)" & _
			   " WHERE " & _
			   " (TO_CHAR(b.date_transaction,'YYYYMM')='" & tahun13 & "'||LPad('" & bulan13 & "',2,0) OR TO_CHAR(b.date_transaction,'YYYYMM') IS NULL) " & _
			   filter & _
			   " ORDER BY fcba,fccode "


    'response.write sqlview
    Set List = Server.CreateObject("ADODB.Recordset")
    List.Open sqlview, connStr, 1
    ListCount = List.RecordCount 

    Set caljson = new JSON  'JSON.aspê°€ í•„ìš”í•¨
    Set rs = connStr.Execute(sqlview)

    Dim List_jsonString13
    List_jsonString13 = caljson.toJSON("results",rs, false)

    Dim jsonObj13, outputObj13
    set jsonObj13 = new JSONobject
    set outputObj13 = jsonObj13.parse(List_jsonString13)  'json.objectClass.aspê°€ í•„ìš”í•¨(íŒŒì‹±í•˜ë ¤ë©´)


    Set oJSON = New aspJSON     'aspJSON.asp í•„ìš”í•¨(forë¬¸)
    oJSON.loadJSON(List_jsonString13)

    Response.write List_jsonString13
end if

'tampilkan data aresta'
if idview="14" then
   
	pcompany 	= Request("pcompany")
    pfcba 		= Request("pfcba")
    ptahun 		= Request("ptahun")
	pbulan 		= Request("pbulan")
	ptype 		= Request("ptype")
	statusfield = Request("statusfield")
	ownership   = Request("ownership")
	r1 			= Request("r1")
	sakun       = Request("sakun")

	
	if trim(pfcba) = "-" then 
		filterallfcba = " AND a.FCBA in (SELECT fccode from businessunit where fccompanycode='" & pcompany & "' and fccode not in('BRPS','CAES') and isactive='TRUE') AND a.YEAR= " & ptahun & " AND a.MONTH= " & pbulan & " "
	else
	    filterallfcba = " AND a.FCBA ='" & pfcba & "' AND a.YEAR= " & ptahun & " AND a.MONTH= " & pbulan & ""
	end if

	if  TRIM(ptype) <> "" then
		filterdivaresta	= " AND a.DIVISION ='" & ptype & "'"
	end if

    if  TRIM(statusfield) <> "" then
		filterstatusaresta	= " AND a.dx_block_status ='" & statusfield & "'"
	end if
    if  TRIM(ownership) <> "" then
		filterownership	= " AND a.ownership ='" & ownership & "'"
	end if

	if r1 = "2" then

    sqlview =  "SELECT " & _
	 " FCCODE, FCBA,DIVISION,TAHUNTANAM,TERRAINTYPE,dx_block_status as STATUS,TOTALSTAND,PRODUCTIVESTAND,OWNERSHIP,Nvl(LUAS_GROSS,2) LUAS_GROSS,Nvl(AREAL_LC,0) AREAL_LC, Nvl(HECTARAGEPLANTED,0) HECTARAGEPLANTED,Nvl(HECTARAGEHARVESTED,0) HECTARAGEHARVESTED, " & _
	 " Sum(perawatan) perawatan, " & _
	 " Sum(pemupukan) pemupukan, " & _
	 " Sum(panenha) panenha, " & _
	 " Sum(panenjjg) panenjjg, " & _
	 " Sum(panenkg) panenkg " & _
	 " FROM  " & _
	 " (  " & _
	 "		SELECT " & _
	 "		a.FCCODE, a.FCBA,a.DIVISION,TO_CHAR(a.PLANTINGDATE,'YYYY') TAHUNTANAM,a.TERRAINTYPE,a.dx_block_status,a.OWNERSHIP,round(a.HECTARAGEPLANTED,2) HECTARAGEPLANTED,round(a.LUAS_GROSS,2)LUAS_GROSS,a.AREAL_LC,  a.HECTARAGEHARVESTED,a.TOTALSTAND,a.PRODUCTIVESTAND , " & _
	 "  	b.perawatan, " & _
	 "		b.pemupukan, " & _
	 "		b.panenha,   " & _
	 "		b.panenjjg,  " & _
	 "		b.panenkg    " & _
	 "		FROM FIELD_HIST_TAB a " & _
	 "		LEFT JOIN " & _
     " 		( " & _
     " 		SELECT b.chargeunit,locationcode,To_Number(To_Char(bkmdate,'YYYY')) YEAR,To_Number(To_Char(bkmdate,'MM')) MONTH,job_forfieldstatus, " & _
     "		CASE  " & _
	 "		WHEN jobgroupcode IN " & _
	 "		( " & _
	 "		'1320100', " & _
	 "		'1320200', " & _
	 "		'1320300', " & _
	 "		'1320400', " & _
	 "		'1320500', " & _
	 "		'1320599', " & _
	 "		'1320600', " & _
	 "		'1320700', " & _
	 "		'1320900', " & _
	 "		'1321000', " & _
	 "		'1321100', " & _
	 "	    '1321200', " & _
	 "		'1321300', " & _
	 "		'1321400', " & _
	 "		'6500100', " & _
	 "		'6500200', " & _
	 "		'6500210', " & _
	 "		'6500300', " & _
	 "		'6500400', " & _
	 "		'6500500', " & _
	 "		'6500600', " & _
	 "		'6500700', " & _
	 "		'6500800', " & _
	 "		'6500900', " & _
	 "		'6501000') " & _
	 "		THEN  " & _
	 "		nvl(Output,0) " & _
	 "		ELSE " & _
	 "		0 " & _
	 "		END perawatan, " & _
	 "		CASE  " & _
	 "		WHEN jobgroupcode IN " & _
	 "		('1320800','6500300') " & _
	 "		THEN " & _
	 "		nvl(Output,0) " & _
	 "		ELSE " & _
	 "		0 " & _
	 "		END pemupukan, " & _
	 "		CASE  " & _
	 "		WHEN jobgroupcode in " & _
	 "		('1338900','6502100') " & _
	 "		THEN " & _
	 "		nvl(Output,0) " & _
	 "		ELSE " & _
	 "		0 " & _
	 "		END panenha, " & _
	 "		CASE  " & _
	 "		WHEN jobgroupcode in " & _
	 "		('1338900','6502100') " & _
	 "		THEN " & _
	 "		nvl(unit,0) " & _
	 "		ELSE " & _
	 "		0 " & _
	 "		END panenjjg, " & _
	 "		CASE  " & _
	 "		WHEN jobgroupcode in  " & _
	 "		('1338900','6502100') " & _
	 "		THEN " & _
	 "		nvl(bunch_weight,0) " & _
	 "		ELSE " & _
	 "		0 " & _
	 "		END panenkg " & _
	 "		FROM BKMUPAHDTL b "  & _ 
	 "		INNER JOIN BKMHDR c ON (b.fcba = c.fcba AND b.bkmno=c.bkmno) " & _
	 "		INNER JOIN JOBMASTER d ON (b.actcode = d.fccode)  " & _
	 "      WHERE 1=1 AND attendance <> 'KP' AND To_Number(To_Char(bkmdate,'YYYY')) = " & ptahun & " AND To_Number(To_Char(bkmdate,'MM')) = " & pbulan & "" & _   
	 "      AND d.fccode like '" & sakun & "%'" & _
	 "      UNION ALL " & _ 
	 " 		SELECT b.chargeunit,locationcode,To_Number(To_Char(wipdate,'YYYY')) YEAR,To_Number(To_Char(wipdate,'MM')) MONTH,job_forfieldstatus, " & _
     "		CASE  " & _
	 "		WHEN jobgroupcode IN " & _
	 "		( " & _
	 "		'1320100', " & _
	 "		'1320200', " & _
	 "		'1320300', " & _
	 "		'1320400', " & _
	 "		'1320500', " & _
	 "		'1320599', " & _
	 "		'1320600', " & _
	 "		'1320700', " & _
	 "		'1320900', " & _
	 "		'1321000', " & _
	 "		'1321100', " & _
	 "	    '1321200', " & _
	 "		'1321300', " & _
	 "		'1321400', " & _
	 "		'6500100', " & _
	 "		'6500200', " & _
	 "		'6500210', " & _
	 "		'6500300', " & _
	 "		'6500400', " & _
	 "		'6500500', " & _
	 "		'6500600', " & _
	 "		'6500700', " & _
	 "		'6500800', " & _
	 "		'6500900', " & _
	 "		'6501000') " & _
	 "		THEN  " & _
	 "		nvl(Output_final,0) " & _
	 "		ELSE " & _
	 "		0 " & _
	 "		END perawatan, " & _
	 "		CASE  " & _
	 "		WHEN jobgroupcode IN " & _
	 "		('1320800','6500300') " & _
	 "		THEN " & _
	 "		nvl(Output_final,0) " & _
	 "		ELSE " & _
	 "		0 " & _
	 "		END pemupukan, " & _
	 "		CASE  " & _
	 "		WHEN jobgroupcode in " & _
	 "		('1338900','6502100') " & _
	 "		THEN " & _
	 "		nvl(Output_final,0) " & _
	 "		ELSE " & _
	 "		0 " & _
	 "		END panenha, " & _
	 "		0 panenjjg, " & _
	 "		0 panenkg " & _
	 "		FROM workinprogressfinal_new b "  & _
	 "		INNER JOIN JOBMASTER d ON (b.jobcode = d.fccode)  " & _
	 "      WHERE 1=1 AND To_Number(To_Char(wipdate,'YYYY')) = " & ptahun & " AND To_Number(To_Char(wipdate,'MM')) = " & pbulan & "" & _   
	 "      AND d.fccode like '" & sakun & "%'" & _
	 " )b ON (a.fcba=b.chargeunit AND a.fccode= b.locationcode AND a.MONTH=b.MONTH AND a.YEAR=b.YEAR and a.dx_block_status=b.job_forfieldstatus)  WHERE 1=1  AND a.activation = 'Y'" 
	 sqlview	= sqlview & filterallfcba 
	 sqlview	= sqlview & filterdivaresta 
	 sqlview	= sqlview & filterstatusaresta 
	 sqlview	= sqlview & filterownership 
	 sqlview	= sqlview & "  ) t " 
	 sqlview	= sqlview & " GROUP BY FCCODE, FCBA,DIVISION,TAHUNTANAM,TERRAINTYPE,dx_block_status,OWNERSHIP,HECTARAGEPLANTED,LUAS_GROSS,AREAL_LC,  HECTARAGEHARVESTED,TOTALSTAND,PRODUCTIVESTAND"
	 sqlview	= sqlview & " ORDER BY FCBA,FCCODE,DIVISION ASC"
	else
		sqlview = "SELECT " & _
			" t.FCBA,'-' FCCODE,'-' division,'-' TAHUNTANAM,'-' TERRAINTYPE,t.dx_block_status,t.OWNERSHIP, " & _
			" SUM(Nvl(t.LUAS_GROSS,0))LUAS_GROSS, " & _
			" SUM(Nvl(t.AREAL_LC,0))AREAL_LC, " & _
			" SUM(t.TOTALSTAND)TOTALSTAND, " & _
			" SUM(t.PRODUCTIVESTAND) PRODUCTIVESTAND, " & _
			" SUM(Nvl(t.HECTARAGEPLANTED,0))HECTARAGEPLANTED, " & _
			" SUM(Nvl(t.HECTARAGEHARVESTED,0))HECTARAGEHARVESTED, " & _
			" Sum(t.perawatan) perawatan, " & _
			" Sum(t.pemupukan) pemupukan, " & _
			" Sum(t.panenha) panenha, " & _
			" Sum(t.panenjjg) panenjjg, " & _
			" Sum(t.panenkg) panenkg " & _
			" FROM  " & _
			 " (  " & _
			 "		SELECT " & _
			 "		a.FCCODE, a.FCBA,a.DIVISION,TO_CHAR(a.PLANTINGDATE,'YYYY') TAHUNTANAM,a.TERRAINTYPE,a.dx_block_status,a.OWNERSHIP,round(a.HECTARAGEPLANTED,2) HECTARAGEPLANTED,round(a.LUAS_GROSS,2)LUAS_GROSS,a.AREAL_LC,  a.HECTARAGEHARVESTED,a.TOTALSTAND,a.PRODUCTIVESTAND , " & _
			 "  	SUM(b.perawatan) perawatan, " & _
			 "		SUM(b.pemupukan) pemupukan, " & _
			 "		SUM(b.panenha) panenha,   " & _
			 "		SUM(b.panenjjg) panenjjg,  " & _
			 "		SUM(b.panenkg) panenkg   " & _
			 "		FROM FIELD_HIST_TAB a " & _
			 "		LEFT JOIN " & _
			 " 		( " & _
			 " 		SELECT b.chargeunit,locationcode,To_Number(To_Char(bkmdate,'YYYY')) YEAR,To_Number(To_Char(bkmdate,'MM')) MONTH,job_forfieldstatus, " & _
			 "		CASE  " & _
			 "		WHEN jobgroupcode IN " & _
			 "		( " & _
			 "		'1320100', " & _
			 "		'1320200', " & _
			 "		'1320300', " & _
			 "		'1320400', " & _
			 "		'1320500', " & _
			 "		'1320600', " & _
			 "		'1320700', " & _
			 "		'1320900', " & _
			 "		'1321000', " & _
			 "		'1321100', " & _
			 "	    '1321200', " & _
			 "		'1321300', " & _
			 "		'1321400', " & _
			 "		'6500100', " & _
			 "		'6500200', " & _
			 "		'6500210', " & _
			 "		'6500300', " & _
			 "		'6500400', " & _
			 "		'6500500', " & _
			 "		'6500600', " & _
			 "		'6500700', " & _
			 "		'6500800', " & _
			 "		'6500900', " & _
			 "		'6501000') " & _
			 "		THEN  " & _
			 "		nvl(Output,0) " & _
			 "		ELSE " & _
			 "		0 " & _
			 "		END perawatan, " & _
			 "		CASE  " & _
			 "		WHEN jobgroupcode IN " & _
			 "		('1320800','6500300') " & _
			 "		THEN " & _
			 "		nvl(Output,0) " & _
			 "		ELSE " & _
			 "		0 " & _
			 "		END pemupukan, " & _
			 "		CASE  " & _
			 "		WHEN jobgroupcode in " & _
			 "		('1338900','6502100') " & _
			 "		THEN " & _
			 "		nvl(Output,0) " & _
			 "		ELSE " & _
			 "		0 " & _
			 "		END panenha, " & _
			 "		CASE  " & _
			 "		WHEN jobgroupcode in " & _
			 "		('1338900','6502100') " & _
			 "		THEN " & _
			 "		nvl(unit,0) " & _
			 "		ELSE " & _
			 "		0 " & _
			 "		END panenjjg, " & _
			 "		CASE  " & _
			 "		WHEN jobgroupcode in  " & _
			 "		('1338900','6502100') " & _
			 "		THEN " & _
			 "		nvl(bunch_weight,0) " & _
			 "		ELSE " & _
			 "		0 " & _
			 "		END panenkg " & _
			 "		FROM BKMUPAHDTL b "  & _
			 "		INNER JOIN BKMHDR c ON (b.fcba = c.fcba AND b.bkmno=c.bkmno) " & _
			 "		INNER JOIN JOBMASTER d ON (b.actcode = d.fccode)  " & _
			 "      WHERE 1=1 AND attendance <> 'KP' AND To_Number(To_Char(bkmdate,'YYYY')) = " & ptahun & " AND To_Number(To_Char(bkmdate,'MM')) = " & pbulan & "" & _   
			 "      AND d.fccode like '" & sakun & "%'" & _
			 "      UNION ALL " & _
			 " 		SELECT b.chargeunit,locationcode,To_Number(To_Char(wipdate,'YYYY')) YEAR,To_Number(To_Char(wipdate,'MM')) MONTH,job_forfieldstatus, " & _
			 "		CASE  " & _
			 "		WHEN jobgroupcode IN " & _
			 "		( " & _
			 "		'1320100', " & _
			 "		'1320200', " & _
			 "		'1320300', " & _
			 "		'1320400', " & _
			 "		'1320500', " & _
			 "		'1320600', " & _
			 "		'1320700', " & _
			 "		'1320900', " & _
			 "		'1321000', " & _
			 "		'1321100', " & _
			 "	    '1321200', " & _
			 "		'1321300', " & _
			 "		'1321400', " & _
			 "		'6500100', " & _
			 "		'6500200', " & _
			 "		'6500210', " & _
			 "		'6500300', " & _
			 "		'6500400', " & _
			 "		'6500500', " & _
			 "		'6500600', " & _
			 "		'6500700', " & _
			 "		'6500800', " & _
			 "		'6500900', " & _
			 "		'6501000') " & _
			 "		THEN  " & _
			 "		nvl(Output_final,0) " & _
			 "		ELSE " & _
			 "		0 " & _
			 "		END perawatan, " & _
			 "		CASE  " & _
			 "		WHEN jobgroupcode IN " & _
			 "		('1320800','6500300') " & _
			 "		THEN " & _
			 "		nvl(Output_final,0) " & _
			 "		ELSE " & _
			 "		0 " & _
			 "		END pemupukan, " & _
			 "		CASE  " & _
			 "		WHEN jobgroupcode in " & _
			 "		('1338900','6502100') " & _
			 "		THEN " & _
			 "		nvl(Output_final,0) " & _
			 "		ELSE " & _
			 "		0 " & _
			 "		END panenha, " & _
			 "		0 panenjjg, " & _
			 "		0 panenkg " & _
			 "		FROM workinprogressfinal_new b "  & _
			 "		INNER JOIN JOBMASTER d ON (b.jobcode = d.fccode)  " & _
			 "      WHERE 1=1  AND To_Number(To_Char(wipdate,'YYYY')) = " & ptahun & " AND To_Number(To_Char(wipdate,'MM')) = " & pbulan & "" & _   
	 	     "      AND d.fccode like '" & sakun & "%'" & _
	         " )b ON (a.fcba=b.chargeunit AND a.fccode= b.locationcode AND a.MONTH=b.MONTH AND a.YEAR=b.YEAR and a.dx_block_status=b.job_forfieldstatus)  WHERE 1=1 AND a.activation = 'Y' " & _
					filterallfcba & _ 
					filterdivaresta & _
					filterstatusaresta & _
					filterownership & _
			"		GROUP BY a.FCCODE, a.FCBA,a.DIVISION,TO_CHAR(a.PLANTINGDATE,'YYYY'),a.TERRAINTYPE,a.dx_block_status,a.OWNERSHIP, " & _
			"       round(a.HECTARAGEPLANTED,2) ,round(a.LUAS_GROSS,2),a.AREAL_LC, a.HECTARAGEHARVESTED,a.TOTALSTAND,a.PRODUCTIVESTAND" & _
			"  ) t " & _
			"  GROUP BY t.FCBA,t.dx_block_status,t.OWNERSHIP " & _
			"  ORDER BY t.FCBA,t.dx_block_status,t.OWNERSHIP" 
	end if

 'response.write sqlview
    Set List = Server.CreateObject("ADODB.Recordset")
    List.Open sqlview, connStr, 1
    ListCount = List.RecordCount 

    Set caljson = new JSON  'JSON.aspê°€ í•„ìš”í•¨
    Set rs = connStr.Execute(sqlview)

    Dim List_jsonString14
    List_jsonString14 = caljson.toJSON("results",rs, false)

    Dim jsonObj14, outputObj14
    set jsonObj14 = new JSONobject
    set outputObj14 = jsonObj14.parse(List_jsonString14)  'json.objectClass.aspê°€ í•„ìš”í•¨(íŒŒì‹±í•˜ë ¤ë©´)


    Set oJSON = New aspJSON     'aspJSON.asp í•„ìš”í•¨(forë¬¸)
    oJSON.loadJSON(List_jsonString14)

    Response.write List_jsonString14
end if

'tampilkan data hrga'
if idview="15" then
   
	tahun = Request("tahun")
    bulan = Request("bulan")
    storecode = Request("storecode")
	itemcode = Request("itemcode")
    fcba = session("fcba")

    sqlview =  "SELECT AMTPERUNIT " & _
    " FROM STOCKLEDGER" & _
			   " WHERE " & _
			   " YEAR='" & tahun & "' AND MONTH='" & bulan & "' AND STORECODE='" & storecode & "' AND ITEMCODE='" & itemcode & "' and fcba='" & fcba & "'"


    'response.write sqlview
    Set List = Server.CreateObject("ADODB.Recordset")
    List.Open sqlview, connStr, 1
    ListCount = List.RecordCount 

    Set caljson = new JSON  'JSON.aspê°€ í•„ìš”í•¨
    Set rs = connStr.Execute(sqlview)

    Dim List_jsonString15
    List_jsonString15 = caljson.toJSON("results",rs, false)

    Dim jsonObj15, outputObj15
    set jsonObj15 = new JSONobject
    set outputObj15 = jsonObj15.parse(List_jsonString15)  'json.objectClass.aspê°€ í•„ìš”í•¨(íŒŒì‹±í•˜ë ¤ë©´)


    Set oJSON = New aspJSON     'aspJSON.asp í•„ìš”í•¨(forë¬¸)
    oJSON.loadJSON(List_jsonString15)

    Response.write List_jsonString15
end if
'tampilkan data journal'
if idview="16" then
    dim tahun16,bulan16,sourcel6,filter16,posting16
    tahun16 = Request("tahun")
    bulan16 = Request("bulan")
	sourcel6= Request("sourcecode")
	company16 = Request("company")
	fcba16   = Request("fcba")
	posting16 = Request("posting")
    

    Set List = Server.CreateObject("ADODB.Recordset")
	
	if trim(fcba16) = "-" then 
		filter16 = " AND a.FCBA in (SELECT fccode from businessunit where fccompanycode='" & company16 & "' and isactive='TRUE') "
	else
	    filter16 = " AND a.FCBA ='" & fcba16 & "' "
	end if

    sqlview =   " SELECT a.vouchercode docno,to_char(datecreated,'dd/mm/yyyy')docdate,to_char(datecreated,'mm/dd/yyyy')docdate2, nvl(referenceno,'-') remarks," & _
	            " Sum(CASE WHEN b.amount >= 0 THEN b.amount ELSE 0 END) debet, " & _
                " Sum(CASE WHEN b.amount < 0 THEN b.amount ELSE 0 END) kredit,a.fcba," & _
				" rowstateid,nvl(c.postingflg,'0') postingflg,a.transtype,a.bankcode,a.fcedit,a.fcentry " & _
				" FROM paymentvoucher a " & _
				" INNER JOIN paymentvoucherdetail b ON (a.fcba=b.fcba AND a.vouchercode=b.vouchercode) " & _
				" LEFT JOIN (SELECT DISTINCT journalno,fcba,postingflg FROM gledgerdtl) c ON  (a.fcba=b.fcba AND a.vouchercode=c.journalno ) " & _
				" where  " & _
				" Accountingyear="& tahun16 & " " & _ 
				" AND Accountingperiod="& bulan16 & " " & _
				" AND transtype='03-JOURNAL' " & _
				" AND nvl(c.postingflg,'0') like '%" & posting16 & "%' " & _
				filter16 & _ 
				" GROUP BY a.vouchercode,datecreated,referenceno,a.fcba,nvl(c.postingflg,'0'),rowstateid,a.transtype,a.bankcode,a.fcedit,a.fcentry " & _
				" ORDER BY a.fcba,docdate,docno " 
              
	List.Open sqlview, connStr, 1
    ListCount = List.RecordCount 

    Set caljson = new JSON  'JSON.aspê°€ í•„ìš”í•¨
    Set rs = connStr.Execute(sqlview)

    Dim List_jsonString16
    List_jsonString16 = caljson.toJSON("results",rs, false)

    Dim jsonObj16, outputObj16
    set jsonObj16 = new JSONobject
    set outputObj16 = jsonObj16.parse(List_jsonString16)  'json.objectClass.aspê°€ í•„ìš”í•¨(íŒŒì‹±í•˜ë ¤ë©´)


    Set oJSON = New aspJSON     'aspJSON.asp í•„ìš”í•¨(forë¬¸)
    oJSON.loadJSON(List_jsonString16)

    Response.write List_jsonString16
end if

'tampilkan data hpp harian ext'
if idview="17" then
   
    fcba17 = session("fcba")
    bkmfrom17 = Request("bkmfrom")
    bkmto17   = Request("bkmto")


    Set List = Server.CreateObject("ADODB.Recordset")
	
	sqllist =   " SELECT TO_CHAR(TO_DATE('" & bkmfrom17 & "','MM/DD/YYYY') - 1 + rownum,'MM/DD/YYYY') AS dt " & _
				" FROM all_objects " & _
				" WHERE TO_DATE('" & bkmfrom17 & "','MM/DD/YYYY') - 1 + rownum <= TO_DATE('" & bkmto17 & "', 'MM/DD/YYYY') "
			
			Set rslistdt= Server.CreateObject("ADODB.Recordset")
			rslistdt.CursorType = 3
			rslistdt.CursorLocation = 3
			rslistdt.Open sqllist, connstr
			if not rslistdt.eof then 
			sqlexec = "BEGIN " & vbCrLf 
			  Do While Not rslistdt.eof	
				sqlexec = sqlexec & "fac_hpp_harian_api.ProsesData(to_date('" & rslistdt("dt") & "','mm/dd/yyyy'),'" & fcba17 & "');"  & vbCrLf 
				rslistdt.MoveNext
			  loop
			End if
			sqlexec = sqlexec & "END; " 
			rslistdt.Close
			set rslistdt = Nothing			  


   ' sqlexec = "BEGIN " & _ 
    '          "fac_hpp_harian_api.ProcessHPP_Daily(to_date('" & bkmfrom10 & "','mm/dd/yyyy') , to_date('" & bkmto10 & "','mm/dd/yyyy') ,'" & fcba10 & "');" & _
	'		  "END; " 
    'response.write sqlexec
	connStr.Execute(sqlexec)

    sqlview =  " Select proddate,ext_tbs_terima_ext,ext_cpobeli,ext_cpoolah,ext_cpotot," & _
	   " ext_pkbeli,ext_pkolah,ext_pktot," & _
	   " ext_tbs_amnt,ext_trans_tbs_amnt,ext_trans_cpo_amnt,ext_trans_pk_amnt,ext_tottrans_amnt,(ext_trans_cpo_amnt+ext_trans_pk_amnt) ext_trans_prod_amnt," & _
	   " ext_cpobeli_amnt,ext_pkbeli_amnt,ext_totbeli_amnt,ext_cpoolah_amnt,ext_pkbolah_amnt,ext_totolah_amnt,ext_total_amnt," & _
	   " ext_hpp_cpo_rpkg,ext_hpp_pk_rpkg " & _
	   " From   fac_hpp_harian  " & _
       " WHERE PRODDATE BETWEEN to_date('" & bkmfrom17 & "','mm/dd/yyyy') and to_date('" & bkmto17 & "','mm/dd/yyyy') and fcba='" & fcba17 & "' " & _
       " order by proddate ASC "

    'response.write sqlview

    List.Open sqlview, connStr, 1
    ListCount = List.RecordCount 

    Set caljson = new JSON  'JSON.aspê°€ í•„ìš”í•¨
    Set rs = connStr.Execute(sqlview)

    Dim List_jsonString17
    List_jsonString17 = caljson.toJSON("results",rs, false)

    Dim jsonObj17, outputObj17
    set jsonObj17 = new JSONobject
    set outputObj17 = jsonObj17.parse(List_jsonString17)  'json.objectClass.aspê°€ í•„ìš”í•¨(íŒŒì‹±í•˜ë ¤ë©´)


    Set oJSON = New aspJSON     'aspJSON.asp í•„ìš”í•¨(forë¬¸)
    oJSON.loadJSON(List_jsonString17)

    Response.write List_jsonString17
end if

'tampilkan sales invoice'
if idview="18" then
    dim tglawal18,tglakhir18,posting18,fcba18
    tglawal18   = Request("tglawalsls")
    tglakhir18  = Request("tglakhirsls")
	fcba18      = Request("fcba")
    'fcba        = session("fcba")
	posting18   = Request("posting")
    

    Set List = Server.CreateObject("ADODB.Recordset")
	
    sqlview =   " SELECT sih.fcba,sih.invoice_by,sih.invoicecode,sih.transdate, " & _ 
    " sih.customercode,customer_api.get_description (sih.customercode, sih.fcba) customername, " & _
    " sah.productcode,sum(sid.quantity) quantity,sih.vat,AVG(sih.grandtotal)grandtotal, " & _
    " sih.remarks,sih.salescontractno,AVG(sih.amount_beforevat) amount_beforevat, " & _
    " AVG(sih.amountvat)amountvat,sih.VALID_FROM,sih.valid_until, " & _
    " sih.no_faktur,sih.expire_date,sih.fakturdate,sih.invoicedpno,j.postingflg  " & _
    " FROM salesinvoicedtl sid " & _
    " INNER JOIN salesinvoicehdr sih ON (sih.invoicecode = sid.invoicecode AND sih.fcba = sid.fcba) " & _
    " INNER JOIN salesagreementdet sah " & _
    " ON ( sah.fcba = sih.fcba AND sah.salescontractno = sih.salescontractno)  " & _
    " LEFT JOIN BKMJOURNALHDR j on (j.DOCNO=sih.INVOICECODE and j.fcba=sih.fcba) " & _
    " where sih.fcba='" & fcba18 & "' AND sih.TRANSDATE BETWEEN to_date('" & tglawal18 & "','mm/dd/yyyy') and to_date('" & tglakhir18 & "','mm/dd/yyyy') " & _
    " GROUP BY sih.fcba,sih.invoice_by,sih.invoicecode, " & _
    " sih.transdate, sih.customercode, customer_api.get_description (sih.customercode, sih.fcba) , " & _
    " sah.productcode, sih.vat, sih.remarks,sih.salescontractno, sih.VALID_FROM, " & _
    " sih.valid_until, sih.no_faktur,sih.expire_date,sih.fakturdate,sih.invoicedpno,j.postingflg"  

    
      'response.write sqlview
	List.Open sqlview, connStr, 1
    ListCount = List.RecordCount 

    Set caljson = new JSON  'JSON.aspê°€ í•„ìš”í•¨
    Set rs = connStr.Execute(sqlview)

    Dim List_jsonString18
    List_jsonString18 = caljson.toJSON("results",rs, false)

    Dim jsonObj18, outputObj18
    set jsonObj18 = new JSONobject
    set outputObj18 = jsonObj18.parse(List_jsonString18)  'json.objectClass.aspê°€ í•„ìš”í•¨(íŒŒì‹±í•˜ë ¤ë©´)


    Set oJSON = New aspJSON     'aspJSON.asp í•„ìš”í•¨(forë¬¸)
    oJSON.loadJSON(List_jsonString18)

    Response.write List_jsonString18
end if

'tampilkan Aging AP'
if idview="19" then
    dim tahun19,bulan19,filter19,posting19,filterbulan,regionap,filtercompany1,filtertypeag
    regionap 	= Request("region")
    tahun19 	= Request("tahun")
    bulan19 	= Request("bulan")
	company19 	= Request("company")
	fcba19   	= Request("fcba")
	fctype19   	= Request("fctype")
	fcoa19		= Request("fcoa_ap")
   
	if trim(company19)="-" THEN
		filtercompany1 = " WHERE isactive='TRUE' AND b.region ='" & regionap & "' "
	else
		if trim(fcba19) = "-" then 
			filtercompany1 = " WHERE isactive='TRUE' AND b.fccompanycode ='" & company19 & "' "
		else 
			filtercompany1 = " WHERE isactive='TRUE' AND b.fccode ='" & fcba19 & "' "
		end if
    end if
	
	sqlpt = "select b.fccode,b.fccompanyname from BUSINESSUNIT b" & _ 
	filtercompany1
	set rspt = connStr.Execute(sqlpt)

	if not rspt.eof then
			companyview = rspt("fccompanyname")
		While Not rspt.EOF
			fcbaloop = rspt("fccode")
			
			sqlexec = "BEGIN " & _ 
              "fac_generate_aging_ap_api.ProsesGenerate('" & tahun19 & "','" & bulan19 & "','" & fcbaloop & "','System');" & _
			  "END; " 
'Response.Write sqlexec 
    		connStr.Execute(sqlexec)
		rspt.MoveNext
		Wend
		rspt.close()
		rspt	=	null

	END IF

	
    Set List = Server.CreateObject("ADODB.Recordset")
	
	if trim(fctype19) <> "all" then 
		if trim(fctype19) = "0" then 
			filtertypeag = ""
		else
			filtertypeag = " AND a.fctype ='" & fctype19 & "' "
		end if
	end if

	if trim(company19)="-" THEN
		filtercompany1 = " WHERE b.isactive='TRUE' AND b.region ='" & regionap & "' "
	else
		if trim(fcba19) = "-" then 
			filtercompany1 = " WHERE b.isactive='TRUE' AND b.fccompanycode ='" & company19 & "' "
		else 
			filtercompany1 = " WHERE b.isactive='TRUE' AND b.fccode ='" & fcba19 & "' "
		end if
    end if

	if trim(bulan19) <> "all" then 
	    filterbulan = " AND a.period_month='" & bulan19 & "'"
	end if

	if trim(fcoa19) = "0" or trim(fcoa19) = "all" then 
		selectCoA = "c.fccode || ' - ' || c.fcname fccompanycode, a.fcba, a.fctype, d.fccode || ' - ' || d.FCNAME fcname, c.fccode, a.opening, a.debet, a.credit, a.ending "
		filterFCoA = " "
		innerJoin = " INNER JOIN SUPPLIER c ON a.fccode=c.fccode AND a.fcba=c.fcba INNER JOIN JOBMASTER d ON c.CONTROLJOB = d.FCCODE"
		groupbyCoA = " GROUP BY c.fccode || ' - ' || c.fcname, a.fcba, a.fctype, a.fctype, d.fccode || ' - ' || d.FCNAME, c.fccode, a.opening, a.debet, a.credit, a.ending"
	elseif trim(fcoa19) <> "0" And trim(fcoa19) <> "all" then
		selectCoA = "c.fccode || ' - ' || c.fcname fccompanycode, a.fcba, a.fctype, d.fccode || ' - ' || d.FCNAME fcname, c.fccode, a.opening, a.debet, a.credit, a.ending "
		filterFCoA = " AND c.CONTROLJOB='" & fcoa19 & "' AND a.fctype not in ('IPR') "
		innerJoin = " INNER JOIN SUPPLIER c ON a.fccode=c.fccode AND a.fcba=c.fcba INNER JOIN JOBMASTER d ON c.CONTROLJOB = d.FCCODE"
		groupbyCoA = " GROUP BY c.fccode || ' - ' || c.fcname, a.fcba, a.fctype, a.fctype, d.fccode || ' - ' || d.FCNAME, c.fccode, a.opening, a.debet, a.credit, a.ending"
	end if
	
	if trim(fcoa19) = "0" And trim(fctype19) = "0" then
		selectCoA = "b.FCCOMPANYCODE,a.fcba,a.fccode,a.fcname,a.fctype,a.opening,a.debet,a.credit,a.ending"
		filterFCoA = ""
		innerJoin = ""
		groupbyCoA = " GROUP BY b.FCCOMPANYCODE,a.fcba,a.fccode,a.fcname, a.fctype,a.opening,a.debet,a.credit,a.ending ORDER BY b.FCCOMPANYCODE,a.fcba,a.fctype,a.fcname"
	end if

	if trim(fcoa19) = "all" And trim(fctype19) = "0" then
		selectCoA = "c.fccode || ' - ' || c.fcname fccompanycode, a.fcba, a.fctype, d.fccode || ' - ' || d.FCNAME fcname, c.fccode, a.opening, a.debet, a.credit, a.ending "
		filterFCoA = " "
		innerJoin = " INNER JOIN SUPPLIER c ON a.fccode=c.fccode AND a.fcba=c.fcba INNER JOIN JOBMASTER d ON c.CONTROLJOB = d.FCCODE"
		groupbyCoA = " GROUP BY c.fccode || ' - ' || c.fcname, a.fcba, a.fctype, a.fctype, d.fccode || ' - ' || d.FCNAME, c.fccode, a.opening, a.debet, a.credit, a.ending"
	end if

	if trim(fcoa19) = "0" And trim(fctype19) = "all" then
		selectCoA = "b.FCCOMPANYCODE,a.fcba,a.fccode,a.fcname,a.fctype,a.opening,a.debet,a.credit,a.ending"
		filterFCoA = ""
		innerJoin = ""
		groupbyCoA = " GROUP BY b.FCCOMPANYCODE,a.fcba,a.fccode,a.fcname, a.fctype,a.opening,a.debet,a.credit,a.ending ORDER BY b.FCCOMPANYCODE,a.fcba,a.fctype,a.fcname"
	end if

	if trim(fcoa19) = "0" And Not(trim(fctype19) = "all" And trim(fctype19) = "0") then
		selectCoA = "b.FCCOMPANYCODE,a.fcba,a.fccode,a.fcname,a.fctype,a.opening,a.debet,a.credit,a.ending"
		filterFCoA = ""
		innerJoin = ""
		groupbyCoA = " GROUP BY b.FCCOMPANYCODE,a.fcba,a.fccode,a.fcname, a.fctype,a.opening,a.debet,a.credit,a.ending ORDER BY b.FCCOMPANYCODE,a.fcba,a.fctype,a.fcname"
	end if

	sqlview =   " SELECT " & _
			selectCoA &_
			" FROM fac_ap_aginghdr a " & _
			" INNER JOIN BUSINESSUNIT b ON a.fcba=b.fccode" & _
			innerJoin & _
			filtercompany1 & _ 
			filtertypeag & _
			" AND a.period_year="& tahun19 & " " & _ 
			filterbulan & _
			filterFCoA & _
			groupbyCoA

	'' response.write sqlview
	List.Open sqlview, connStr, 1
    ListCount = List.RecordCount 

    Set caljson = new JSON  'JSON.aspê°€ í•„ìš”í•¨
    Set rs = connStr.Execute(sqlview)

    Dim List_jsonString19
    List_jsonString19 = caljson.toJSON("results",rs, false)

    Dim jsonObj19, outputObj19
    set jsonObj19 = new JSONobject
    set outputObj19 = jsonObj19.parse(List_jsonString19)  'json.objectClass.aspê°€ í•„ìš”í•¨(íŒŒì‹±í•˜ë ¤ë©´)


    Set oJSON = New aspJSON     'aspJSON.asp í•„ìš”í•¨(forë¬¸)
    oJSON.loadJSON(List_jsonString19)

    Response.write List_jsonString19
end if


'tampilkan Laba Rugi Ebitda'
if idview="20" then
    dim tahun20,bulan20,filter20,posting20
    tahun20   = Request("tahun")
 	company20 = Request("company")
	fcba20    = Request("fcba")
  
     
	if trim(company20) <> "-" then 
			if trim(fcba20) = "-" then 
					sqllist =   " SELECT fccode as fcba " & _
						" FROM businessunit " & _
						" WHERE fccompanycode='" & company20 &"' and isactive='TRUE' " 
						
					
					Set rslistdt= Server.CreateObject("ADODB.Recordset")
					rslistdt.CursorType = 3
					rslistdt.CursorLocation = 3
					rslistdt.Open sqllist, connstr
					if not rslistdt.eof then 
					  sqlexec = "BEGIN " & vbCrLf 
					  Do While Not rslistdt.eof	
						sqlexec = sqlexec & "FAC_EBITDA_BULANAN_API.ProsesEbitda(12, " & tahun20 & ",'" & rslistdt("fcba") & "');"  & vbCrLf 
						rslistdt.MoveNext
					  loop
					End if
					sqlexec = sqlexec & "END; " 
					rslistdt.Close
					set rslistdt = Nothing			  

			else
					sqlexec = "BEGIN " & _ 
							  "FAC_EBITDA_BULANAN_API.ProsesEbitda(12," & tahun20 & ",'" & fcba20 & "');" & _
							  "END; " 
			end if
	else
			sqllist =   " SELECT fccode as fcba " & _
						" FROM businessunit " & _
						" WHERE isactive='TRUE' " 
						
					
					Set rslistdt= Server.CreateObject("ADODB.Recordset")
					rslistdt.CursorType = 3
					rslistdt.CursorLocation = 3
					rslistdt.Open sqllist, connstr
					if not rslistdt.eof then 
					  sqlexec = "BEGIN " & vbCrLf 
					  Do While Not rslistdt.eof	
						sqlexec = sqlexec & "FAC_EBITDA_BULANAN_API.ProsesEbitda(12, " & tahun20 & ",'" & rslistdt("fcba") & "');"  & vbCrLf 
						rslistdt.MoveNext
					  loop
					End if
					sqlexec = sqlexec & "END; " 
					rslistdt.Close
					set rslistdt = Nothing	
	
	end if
	connStr.Execute(sqlexec)
	
	
    Set List = Server.CreateObject("ADODB.Recordset")
	
	if trim(fcba20) = "-" then 
	sqlview =   " SELECT reportcode,year period,row_id,description,dispseq,rowtype," & _
				" round(sum(total_),0) total_, round(sum(b1_),0) b1_, round(sum(b2_),0) b2_,round(sum(b3_),0) b3_,round(sum(b4_),0) b4_,round(sum(b5_),0) b5_,round(sum(b6_),0) b6_, " & _ 
				" round(sum(b7_),0) b7_,round(sum(b8_),0) b8_,round(sum(b9_),0) b9_,round(sum(b10_),0) b10_,round(sum(b11_),0) b11_,round(sum(b12_),0) b12_ " & _
	            " FROM fac_ebitda_bulanan " & _
				" where  " & _
				" year="& tahun20 & _ 
				" AND FCBA in (SELECT fccode from businessunit where fccompanycode='" & company20 & "' and isactive='TRUE') " & _
				" GROUP BY reportcode,year,row_id,description,dispseq ,rowtype" & _
				" ORDER BY dispseq " 


	else
	sqlview =   " SELECT reportcode,year period,row_id,description,dispseq,rowtype," & _
				" round(sum(total_),0) total_, round(sum(b1_),0) b1_, round(sum(b2_),0) b2_,round(sum(b3_),0) b3_,round(sum(b4_),0) b4_,round(sum(b5_),0) b5_,round(sum(b6_),0) b6_, " & _ 
				" round(sum(b7_),0) b7_,round(sum(b8_),0) b8_,round(sum(b9_),0) b9_,round(sum(b10_),0) b10_,round(sum(b11_),0) b11_,round(sum(b12_),0) b12_ " & _
	            " FROM fac_ebitda_bulanan " & _
				" where  " & _
				" year="& tahun20 & _ 
				" AND FCBA ='" & fcba20 & "' " & _
				" GROUP BY reportcode,year,row_id,description,dispseq ,rowtype" & _
				" ORDER BY dispseq " 

	end if
	
              
	List.Open sqlview, connStr, 1
    ListCount = List.RecordCount 

    Set caljson = new JSON  'JSON.aspê°€ í•„ìš”í•¨
    Set rs = connStr.Execute(sqlview)

    Dim List_jsonString20
    List_jsonString20 = caljson.toJSON("results",rs, false)

    Dim jsonObj20, outputObj20
    set jsonObj20 = new JSONobject
    set outputObj20 = jsonObj20.parse(List_jsonString20)  'json.objectClass.aspê°€ í•„ìš”í•¨(íŒŒì‹±í•˜ë ¤ë©´)


    Set oJSON = New aspJSON     'aspJSON.asp í•„ìš”í•¨(forë¬¸)
    oJSON.loadJSON(List_jsonString20)

    Response.write List_jsonString20
end if


'tampilkan neraca'
if idview="21" then
    dim tahun21,bulan21,filter21,posting21
    tahun21   = Request("tahun")
 	company21 = Request("company")
	fcba21    = Request("fcba")
  
     
	if trim(company21) <> "-" then 
		if trim(fcba21) = "-" then 
				sqllist =   " SELECT fccode as fcba " & _
					" FROM businessunit " & _
					" WHERE fccompanycode='" & company21 &"' and isactive='TRUE' " 
					
				
				Set rslistdt= Server.CreateObject("ADODB.Recordset")
				rslistdt.CursorType = 3
				rslistdt.CursorLocation = 3
				rslistdt.Open sqllist, connstr
				if not rslistdt.eof then 
				  sqlexec = "BEGIN " & vbCrLf 
				  Do While Not rslistdt.eof	
					sqlexec = sqlexec & "fac_balancesheet_bulanan_api.ProsesNeraca(12, " & tahun21 & ",'" & rslistdt("fcba") & "');"  & vbCrLf 
					rslistdt.MoveNext
				  loop
				End if
				sqlexec = sqlexec & "END; " 
				rslistdt.Close
				set rslistdt = Nothing			  

		else
				sqlexec = "BEGIN " & _ 
						  "fac_balancesheet_bulanan_api.ProsesNeraca(12," & tahun21 & ",'" & fcba21 & "');" & _
						  "END; " 
		end if
	else
				sqllist =   " SELECT fccode as fcba " & _
							" FROM businessunit " & _
							" WHERE isactive='TRUE' Order by fccompanycode,fccode " 
									
				Set rslistdt= Server.CreateObject("ADODB.Recordset")
				rslistdt.CursorType = 3
				rslistdt.CursorLocation = 3
				rslistdt.Open sqllist, connstr
				if not rslistdt.eof then 
				  sqlexec = "BEGIN " & vbCrLf 
				  Do While Not rslistdt.eof	
					sqlexec = sqlexec & "fac_balancesheet_bulanan_api.ProsesNeraca(12, " & tahun21 & ",'" & rslistdt("fcba") & "');"  & vbCrLf 
					rslistdt.MoveNext
				  loop
				End if
				sqlexec = sqlexec & "END; " 
				rslistdt.Close
				set rslistdt = Nothing		

	end if
	connStr.Execute(sqlexec)
	
	
    Set List = Server.CreateObject("ADODB.Recordset")
	
	if trim(company21) <> "-" then 
		if trim(fcba21) = "-" then 
		sqlview =   " SELECT reportcode,year period,row_id,description,dispseq,rowtype," & _
					" round(sum(total_),0) total_, round(sum(b1_),0) b1_, round(sum(b2_),0) b2_,round(sum(b3_),0) b3_,round(sum(b4_),0) b4_,round(sum(b5_),0) b5_,round(sum(b6_),0) b6_, " & _ 
					" round(sum(b7_),0) b7_,round(sum(b8_),0) b8_,round(sum(b9_),0) b9_,round(sum(b10_),0) b10_,round(sum(b11_),0) b11_,round(sum(b12_),0) b12_, " & _
					" round(sum(b13_),0) b13_,round(sum(b14_),0) b14_,round(sum(b15_),0) b15_" & _
					" FROM fac_balancesheet_bulanan " & _
					" where  " & _
					" year="& tahun21 & _ 
					" AND FCBA in (SELECT fccode from businessunit where fccompanycode='" & company21 & "' and isactive='TRUE') " & _
					" GROUP BY reportcode,year,row_id,description,dispseq ,rowtype" & _
					" ORDER BY dispseq " 


		else
		sqlview =   " SELECT reportcode,year period,row_id,description,dispseq,rowtype," & _
					" round(sum(total_),0) total_, round(sum(b1_),0) b1_, round(sum(b2_),0) b2_,round(sum(b3_),0) b3_,round(sum(b4_),0) b4_,round(sum(b5_),0) b5_,round(sum(b6_),0) b6_, " & _ 
					" round(sum(b7_),0) b7_,round(sum(b8_),0) b8_,round(sum(b9_),0) b9_,round(sum(b10_),0) b10_,round(sum(b11_),0) b11_,round(sum(b12_),0) b12_, " & _
					" round(sum(b13_),0) b13_,round(sum(b14_),0) b14_,round(sum(b15_),0) b15_" & _
					" FROM fac_balancesheet_bulanan " & _
					" where  " & _
					" year="& tahun21 & _ 
					" AND FCBA ='" & fcba21 & "' " & _
					" GROUP BY reportcode,year,row_id,description,dispseq ,rowtype" & _
					" ORDER BY dispseq " 

		end if
	else
		sqlview =   " SELECT reportcode,year period,row_id,description,dispseq,rowtype," & _
					" round(sum(total_),0) total_, round(sum(b1_),0) b1_, round(sum(b2_),0) b2_,round(sum(b3_),0) b3_,round(sum(b4_),0) b4_,round(sum(b5_),0) b5_,round(sum(b6_),0) b6_, " & _ 
					" round(sum(b7_),0) b7_,round(sum(b8_),0) b8_,round(sum(b9_),0) b9_,round(sum(b10_),0) b10_,round(sum(b11_),0) b11_,round(sum(b12_),0) b12_, " & _
					" round(sum(b13_),0) b13_,round(sum(b14_),0) b14_,round(sum(b15_),0) b15_" & _
					" FROM fac_balancesheet_bulanan " & _
					" where  " & _
					" year="& tahun21 & _ 
					" AND FCBA in (SELECT fccode from businessunit where isactive='TRUE') " & _
					" GROUP BY reportcode,year,row_id,description,dispseq ,rowtype" & _
					" ORDER BY dispseq "
	
	end if
	
              
    'response.write sqlview
	List.Open sqlview, connStr, 1
    ListCount = List.RecordCount 

    Set caljson = new JSON  'JSON.aspê°€ í•„ìš”í•¨
    Set rs = connStr.Execute(sqlview)

    Dim List_jsonString21
    List_jsonString21 = caljson.toJSON("results",rs, false)

    Dim jsonObj21, outputObj21
    set jsonObj21 = new JSONobject
    set outputObj21 = jsonObj21.parse(List_jsonString21)  'json.objectClass.aspê°€ í•„ìš”í•¨(íŒŒì‹±í•˜ë ¤ë©´)


    Set oJSON = New aspJSON     'aspJSON.asp í•„ìš”í•¨(forë¬¸)
    oJSON.loadJSON(List_jsonString21)

    Response.write List_jsonString21
end if

'tampilkan arus kas '
if idview="22" then
    dim tahun22,bulan22,filter22,posting22
    tahun22   = Request("tahun")
 	company22 = Request("company")
	fcba22    = Request("fcba")
  
     
	if trim(company22) <> "-" then 
		if trim(fcba22) = "-" then 
				sqllist =   " SELECT fccode as fcba " & _
					" FROM businessunit " & _
					" WHERE fccompanycode='" & company22 &"' and isactive='TRUE' " 
					
				
				Set rslistdt= Server.CreateObject("ADODB.Recordset")
				rslistdt.CursorType = 3
				rslistdt.CursorLocation = 3
				rslistdt.Open sqllist, connstr
				if not rslistdt.eof then 
				  sqlexec = "BEGIN " & vbCrLf 
				  Do While Not rslistdt.eof	
					sqlexec = sqlexec & "fac_cashflow_bulanan_api.ProsesArusKas(12, " & tahun22 & ",'" & rslistdt("fcba") & "');"  & vbCrLf 
					rslistdt.MoveNext
				  loop
				End if
				sqlexec = sqlexec & "END; " 
				rslistdt.Close
				set rslistdt = Nothing			  

		else
				sqlexec = "BEGIN " & _ 
						  "fac_cashflow_bulanan_api.ProsesArusKas(12," & tahun22 & ",'" & fcba22 & "');" & _
						  "END; " 
		end if
	else
				sqllist =   " SELECT fccode as fcba " & _
							" FROM businessunit " & _
							" WHERE isactive='TRUE' Order by fccompanycode,fccode " 
									
				Set rslistdt= Server.CreateObject("ADODB.Recordset")
				rslistdt.CursorType = 3
				rslistdt.CursorLocation = 3
				rslistdt.Open sqllist, connstr
				if not rslistdt.eof then 
				  sqlexec = "BEGIN " & vbCrLf 
				  Do While Not rslistdt.eof	
					sqlexec = sqlexec & "fac_cashflow_bulanan_api.ProsesArusKas(12, " & tahun22 & ",'" & rslistdt("fcba") & "');"  & vbCrLf 
					rslistdt.MoveNext
				  loop
				End if
				sqlexec = sqlexec & "END; " 
				rslistdt.Close
				set rslistdt = Nothing		

	end if
	connStr.Execute(sqlexec)
	
	
    Set List = Server.CreateObject("ADODB.Recordset")
	
	if trim(company22) <> "-" then 
		if trim(fcba22) = "-" then 
		sqlview =   " SELECT reportcode,year period,row_id,description,dispseq,rowtype," & _
					" round(sum(total_),0) total_, round(sum(b1_),0) b1_, round(sum(b2_),0) b2_,round(sum(b3_),0) b3_,round(sum(b4_),0) b4_,round(sum(b5_),0) b5_,round(sum(b6_),0) b6_, " & _ 
					" round(sum(b7_),0) b7_,round(sum(b8_),0) b8_,round(sum(b9_),0) b9_,round(sum(b10_),0) b10_,round(sum(b11_),0) b11_,round(sum(b12_),0) b12_, " & _
					" round(sum(b13_),0) b13_,round(sum(b14_),0) b14_,round(sum(b15_),0) b15_" & _
					" FROM fac_cashflow_bulanan " & _
					" where  " & _
					" year="& tahun22 & _ 
					" AND FCBA in (SELECT fccode from businessunit where isactive='TRUE' and fccompanycode='" & company22 & "') " & _
					" GROUP BY reportcode,year,row_id,description,dispseq ,rowtype" & _
					" ORDER BY dispseq " 


		else
		sqlview =   " SELECT reportcode,year period,row_id,description,dispseq,rowtype," & _
					" round(sum(total_),0) total_, round(sum(b1_),0) b1_, round(sum(b2_),0) b2_,round(sum(b3_),0) b3_,round(sum(b4_),0) b4_,round(sum(b5_),0) b5_,round(sum(b6_),0) b6_, " & _ 
					" round(sum(b7_),0) b7_,round(sum(b8_),0) b8_,round(sum(b9_),0) b9_,round(sum(b10_),0) b10_,round(sum(b11_),0) b11_,round(sum(b12_),0) b12_, " & _
					" round(sum(b13_),0) b13_,round(sum(b14_),0) b14_,round(sum(b15_),0) b15_" & _
					" FROM fac_cashflow_bulanan " & _
					" where  " & _
					" year="& tahun22 & _ 
					" AND FCBA ='" & fcba22 & "' " & _
					" GROUP BY reportcode,year,row_id,description,dispseq ,rowtype" & _
					" ORDER BY dispseq " 

		end if
	else
		sqlview =   " SELECT reportcode,year period,row_id,description,dispseq,rowtype," & _
					" round(sum(total_),0) total_, round(sum(b1_),0) b1_, round(sum(b2_),0) b2_,round(sum(b3_),0) b3_,round(sum(b4_),0) b4_,round(sum(b5_),0) b5_,round(sum(b6_),0) b6_, " & _ 
					" round(sum(b7_),0) b7_,round(sum(b8_),0) b8_,round(sum(b9_),0) b9_,round(sum(b10_),0) b10_,round(sum(b11_),0) b11_,round(sum(b12_),0) b12_, " & _
					" round(sum(b13_),0) b13_,round(sum(b14_),0) b14_,round(sum(b15_),0) b15_" & _
					" FROM fac_cashflow_bulanan " & _
					" where  " & _
					" year="& tahun22 & _ 
					" AND FCBA in (SELECT fccode from businessunit where isactive='TRUE') " & _
					" GROUP BY reportcode,year,row_id,description,dispseq ,rowtype" & _
					" ORDER BY dispseq "
	
	end if
	
              
    'response.write sqlview
	List.Open sqlview, connStr, 1
    ListCount = List.RecordCount 

    Set caljson = new JSON  'JSON.aspê°€ í•„ìš”í•¨
    Set rs = connStr.Execute(sqlview)

    Dim List_jsonString22
    List_jsonString22 = caljson.toJSON("results",rs, false)

    Dim jsonObj22, outputObj22
    set jsonObj22 = new JSONobject
    set outputObj22 = jsonObj22.parse(List_jsonString22)  'json.objectClass.aspê°€ í•„ìš”í•¨(íŒŒì‹±í•˜ë ¤ë©´)


    Set oJSON = New aspJSON     'aspJSON.asp í•„ìš”í•¨(forë¬¸)
    oJSON.loadJSON(List_jsonString22)

    Response.write List_jsonString22
end if

'tampilkan Aging AR'
if idview="23" then
    dim tahun23,bulan23,filter23,posting23
    tahun23 = Request("tahun")
    bulan23 = Request("bulan")
	company23 = Request("company")
	fcba23   = Request("fcba")
	fctype23   = Request("fctype")
   
   	sqlexec = "BEGIN " & _ 
              "fac_generate_aging_ar_api.ProsesGenerate(" & tahun23 & "," & bulan23 & ",'" & fcba23 & "','System');" & _
			  "END; " 
    connStr.Execute(sqlexec)
	
    Set List = Server.CreateObject("ADODB.Recordset")
	
	if trim(fcba23) = "-" then 
		filter23 = " AND a.FCBA in (SELECT fccode from businessunit where fccompanycode='" & company23 & "' and isactive='TRUE') "
	else
	    filter23 = " AND a.FCBA ='" & fcba23 & "' "
	end if

    sqlview =   " SELECT fcba,fccode,fcname,fctype,opening,debet,credit,ending " & _
	            " FROM fac_ar_aginghdr a " & _
				" where  " & _
				" period_year="& tahun23 & " " & _ 
				" AND period_month="& bulan23 & " AND fctype='" & fctype23 & "' " & _
				filter23 & _ 
				" ORDER BY a.fcba,fctype,fcname " 
              
	'Response.write 	sqlview
	
	List.Open sqlview, connStr, 1
    ListCount = List.RecordCount 

    Set caljson = new JSON  'JSON.aspê°€ í•„ìš”í•¨
    Set rs = connStr.Execute(sqlview)

    Dim List_jsonString23
    List_jsonString23 = caljson.toJSON("results",rs, false)

    Dim jsonObj23, outputObj23
    set jsonObj23 = new JSONobject
    set outputObj23 = jsonObj23.parse(List_jsonString23)  'json.objectClass.aspê°€ í•„ìš”í•¨(íŒŒì‹±í•˜ë ¤ë©´)


    Set oJSON = New aspJSON     'aspJSON.asp í•„ìš”í•¨(forë¬¸)
    oJSON.loadJSON(List_jsonString23)

    Response.write List_jsonString23
end if

'tampilkan arus kas '
if idview="24" then
    dim tahun24,bulan24,filter24,posting24,region24
    region24   = Request("region")
    tahun24   = Request("tahun")
 	company24 = Request("company")
	fcba24    = Request("fcba")
	groupby   = Request("groupby")
  
	if trim(company24) <> "-" then 
		if trim(fcba24) = "-" then 
				sqllist =   " SELECT fccode as fcba " & _
					" FROM businessunit " & _
					" WHERE fccompanycode='" & company24 &"' and isactive='TRUE' " 
					
				
				Set rslistdt= Server.CreateObject("ADODB.Recordset")
				rslistdt.CursorType = 3
				rslistdt.CursorLocation = 3
				rslistdt.Open sqllist, connstr
				if not rslistdt.eof then 
				  sqlexec = "BEGIN " & vbCrLf 
				  Do While Not rslistdt.eof	
					sqlexec = sqlexec & "fac_trialbalance_bulanan_api.ProsesNeraca(" & groupby & ",12, " & tahun24 & ",'" & rslistdt("fcba") & "');"  & vbCrLf 
					rslistdt.MoveNext
				  loop
				End if
				sqlexec = sqlexec & "END; " 
				rslistdt.Close
				set rslistdt = Nothing			  

		else
				sqlexec = "BEGIN " & _ 
						  "fac_trialbalance_bulanan_api.ProsesNeraca(" & groupby & ",12," & tahun24 & ",'" & fcba24 & "');" & _
						  "END; " 
		end if
	else
				' Perbaikan: bila region dipilih, neraca hanya diproses untuk BA di region itu (yang memang ditampilkan)
				if trim(region24) <> "all" then
					sqllist =   " SELECT fccode as fcba " & _
							" FROM businessunit " & _
							" WHERE isactive='TRUE' AND region='" & Replace(trim(region24), "'", "''") & "' Order by fccompanycode,fccode " 
				else
				sqllist =   " SELECT fccode as fcba " & _
							" FROM businessunit " & _
							" WHERE isactive='TRUE' Order by fccompanycode,fccode " 
				end if
									
				Set rslistdt= Server.CreateObject("ADODB.Recordset")
				rslistdt.CursorType = 3
				rslistdt.CursorLocation = 3
				rslistdt.Open sqllist, connstr
				if not rslistdt.eof then 
				  sqlexec = "BEGIN " & vbCrLf 
				  Do While Not rslistdt.eof	
					sqlexec = sqlexec & "fac_trialbalance_bulanan_api.ProsesNeraca(" & groupby & ",12, " & tahun24 & ",'" & rslistdt("fcba") & "');"  & vbCrLf 
					rslistdt.MoveNext
				  loop
				End if
				sqlexec = sqlexec & "END; " 
				rslistdt.Close
				set rslistdt = Nothing		

	end if
	connStr.Execute(sqlexec)
	
	
    Set List = Server.CreateObject("ADODB.Recordset")

	' if region24<> "all" then 
	' end if
	if trim(region24) <> "all" then
	if trim(company24) <> "-" then 
		if trim(fcba24) = "-" then 
		sqlview =   " SELECT year period,acccode,description,rowtype," & _
					" round(sum(total_),0) total_, round(sum(b1_),0) b1_, round(sum(b2_),0) b2_,round(sum(b3_),0) b3_,round(sum(b4_),0) b4_,round(sum(b5_),0) b5_,round(sum(b6_),0) b6_, " & _ 
					" round(sum(b7_),0) b7_,round(sum(b8_),0) b8_,round(sum(b9_),0) b9_,round(sum(b10_),0) b10_,round(sum(b11_),0) b11_,round(sum(b12_),0) b12_, " & _
					" round(sum(b13_),0) b13_,round(sum(b14_),0) b14_,round(sum(b15_),0) b15_" & _
					" FROM fac_trialbalance_bulanan " & _
					" where  " & _
					" year="& tahun24 & _ 
					" AND FCBA in (SELECT fccode from businessunit where fccompanycode='" & company24 & "' and isactive='TRUE') " & _
					" AND rowtype='" & groupby & "'" & _
					" GROUP BY year,acccode,description,rowtype" & _
					" ORDER BY acccode " 


		else
		sqlview =   " SELECT year period,acccode,description,rowtype," & _
					" round(sum(total_),0) total_, round(sum(b1_),0) b1_, round(sum(b2_),0) b2_,round(sum(b3_),0) b3_,round(sum(b4_),0) b4_,round(sum(b5_),0) b5_,round(sum(b6_),0) b6_, " & _ 
					" round(sum(b7_),0) b7_,round(sum(b8_),0) b8_,round(sum(b9_),0) b9_,round(sum(b10_),0) b10_,round(sum(b11_),0) b11_,round(sum(b12_),0) b12_, " & _
					" round(sum(b13_),0) b13_,round(sum(b14_),0) b14_,round(sum(b15_),0) b15_" & _
					" FROM fac_trialbalance_bulanan " & _
					" where  " & _
					" year="& tahun24 & _ 
					" AND FCBA ='" & fcba24 & "' " & _
					" AND rowtype='" & groupby & "'" & _
					" GROUP BY year,acccode,description,rowtype" & _
					" ORDER BY acccode " 

		end if
	else
		sqlview =   " SELECT year period,acccode,description,rowtype," & _
					" round(sum(total_),0) total_, round(sum(b1_),0) b1_, round(sum(b2_),0) b2_,round(sum(b3_),0) b3_,round(sum(b4_),0) b4_,round(sum(b5_),0) b5_,round(sum(b6_),0) b6_, " & _ 
					" round(sum(b7_),0) b7_,round(sum(b8_),0) b8_,round(sum(b9_),0) b9_,round(sum(b10_),0) b10_,round(sum(b11_),0) b11_,round(sum(b12_),0) b12_, " & _
					" round(sum(b13_),0) b13_,round(sum(b14_),0) b14_,round(sum(b15_),0) b15_" & _
					" FROM fac_trialbalance_bulanan " & _
					" where  " & _
					" year="& tahun24 & _ 
					" AND FCBA in (SELECT fccode from businessunit where isactive='TRUE' AND region='" & region24 & "') " & _
					" AND rowtype='" & groupby & "'" & _
					" GROUP BY year,acccode,description,rowtype" & _
					" ORDER BY acccode "
	
	end if
	else 
		sqlview =   " SELECT year period,acccode,description,rowtype," & _
					" round(sum(total_),0) total_, round(sum(b1_),0) b1_, round(sum(b2_),0) b2_,round(sum(b3_),0) b3_,round(sum(b4_),0) b4_,round(sum(b5_),0) b5_,round(sum(b6_),0) b6_, " & _ 
					" round(sum(b7_),0) b7_,round(sum(b8_),0) b8_,round(sum(b9_),0) b9_,round(sum(b10_),0) b10_,round(sum(b11_),0) b11_,round(sum(b12_),0) b12_, " & _
					" round(sum(b13_),0) b13_,round(sum(b14_),0) b14_,round(sum(b15_),0) b15_" & _
					" FROM fac_trialbalance_bulanan " & _
					" where  " & _
					" year="& tahun24 & _ 
					" AND FCBA in (SELECT fccode from businessunit where isactive='TRUE') " & _
					" AND rowtype='" & groupby & "'" & _
					" GROUP BY year,acccode,description,rowtype" & _
					" ORDER BY acccode "
	end if
	
              
 '   response.write sqlview
	' Perbaikan: List.Open/RecordCount (query yang sama dijalankan dua kali, hasil tidak dipakai) dihapus

    Set caljson = new JSON  'JSON.aspê°€ í•„ìš”í•¨
    Set rs = connStr.Execute(sqlview)

    Dim List_jsonString24
    List_jsonString24 = caljson.toJSON("results",rs, false)

    Dim jsonObj24, outputObj24
    ' Perbaikan: JSON tidak lagi di-parse ulang dengan JSONobject/aspJSON (hasilnya tidak dipakai;
    ' memakan ~80 detik untuk ~950 baris). Output tetap List_jsonString24 yang sama.

    Response.write List_jsonString24
end if

'qry laporan pabrik
if idview="25" then
   
    fcba = Request("fcba")
    bkmfrom = Request("bkmfrom")
    bkmto   = Request("bkmto")


    Set List = Server.CreateObject("ADODB.Recordset")


    sqlview =  "SELECT a.proddate,a.cpoffa,f.prod_olah_amnt, " & _
	" ( SELECT Sum(prod_olah_amnt) FROM FAC_HPP_HARIAN b WHERE b.proddate <= a.proddate AND b.fcba=a.fcba " & _ 
	" and proddate BETWEEN TO_DATE('" & bkmfrom & "','MM/DD/YYYY') AND TO_DATE('" & bkmto & "','MM/DD/YYYY') ) prod_olah_amntsbi," & _
	" a.jamkerja jamkerjahi, " & _
	" (SELECT Sum(jamkerja) FROM productiontrn b WHERE b.proddate <= a.proddate AND b.fcba=a.fcba and proddate BETWEEN TO_DATE('" & bkmfrom & "','MM/DD/YYYY') AND TO_DATE('" & bkmto & "','MM/DD/YYYY') ) jamkerjasbi," & _
	" a.tbsterima tbsterimahi," & _
	" (SELECT Sum(tbsterima) FROM productiontrn b WHERE b.proddate <= a.proddate AND b.fcba=a.fcba and proddate BETWEEN TO_DATE('" & bkmfrom & "','MM/DD/YYYY') AND TO_DATE('" & bkmto & "','MM/DD/YYYY') ) tbsterimasbi," & _
	" a.tbssisa tbssisahi," & _
	" (SELECT Sum(tbssisa) FROM productiontrn b WHERE b.proddate <= a.proddate AND b.fcba=a.fcba and proddate BETWEEN TO_DATE('" & bkmfrom & "','MM/DD/YYYY') AND TO_DATE('" & bkmto & "','MM/DD/YYYY') ) tbssisasbi," & _
	" a.potongantbs potongantbshi," & _
	" (SELECT Sum(potongantbs) FROM productiontrn b WHERE b.proddate <= a.proddate AND b.fcba=a.fcba and proddate BETWEEN TO_DATE('" & bkmfrom & "','MM/DD/YYYY') AND TO_DATE('" & bkmto & "','MM/DD/YYYY') ) potongantbssbi," & _
	" a.cpokirim," & _
	" (SELECT Sum(cpokirim) FROM productiontrn b WHERE b.proddate <= a.proddate AND b.fcba=a.fcba and proddate BETWEEN TO_DATE('" & bkmfrom & "','MM/DD/YYYY') AND TO_DATE('" & bkmto & "','MM/DD/YYYY') ) cpokirimsbi," & _
	" a.adjcpo," & _
	" (SELECT Sum(adjcpo) FROM productiontrn b WHERE b.proddate <= a.proddate AND b.fcba=a.fcba and proddate BETWEEN TO_DATE('" & bkmfrom & "','MM/DD/YYYY') AND TO_DATE('" & bkmto & "','MM/DD/YYYY') ) adjcposbi," & _
	" a.cpocot," & _
	" (SELECT Sum(cpocot) FROM productiontrn b WHERE b.proddate <= a.proddate AND b.fcba=a.fcba and proddate BETWEEN TO_DATE('" & bkmfrom & "','MM/DD/YYYY') AND TO_DATE('" & bkmto & "','MM/DD/YYYY') ) cpocotsbi," & _
	" a.cpocst," & _
	" (SELECT Sum(cpocst) FROM productiontrn b WHERE b.proddate <= a.proddate AND b.fcba=a.fcba and proddate BETWEEN TO_DATE('" & bkmfrom & "','MM/DD/YYYY') AND TO_DATE('" & bkmto & "','MM/DD/YYYY') ) cpocstsbi," & _
	" a.cpoot," & _
	" (SELECT Sum(cpoot) FROM productiontrn b WHERE b.proddate <= a.proddate AND b.fcba=a.fcba and proddate BETWEEN TO_DATE('" & bkmfrom & "','MM/DD/YYYY') AND TO_DATE('" & bkmto & "','MM/DD/YYYY') ) cpootsbi," & _
	" a.cpotankisi," & _
	" (SELECT Sum(cpotankisi) FROM productiontrn b WHERE b.proddate <= a.proddate AND b.fcba=a.fcba and proddate BETWEEN TO_DATE('" & bkmfrom & "','MM/DD/YYYY') AND TO_DATE('" & bkmto & "','MM/DD/YYYY') ) cpotankisisbi," & _
	" a.kernelkirim," & _
	" (SELECT Sum(kernelkirim) FROM productiontrn b WHERE b.proddate <= a.proddate AND b.fcba=a.fcba and proddate BETWEEN TO_DATE('" & bkmfrom & "','MM/DD/YYYY') AND TO_DATE('" & bkmto & "','MM/DD/YYYY') ) kernelkirimsbi," & _
	" a.kernelsilo," & _
	" (SELECT Sum(kernelsilo) FROM productiontrn b WHERE b.proddate <= a.proddate AND b.fcba=a.fcba and proddate BETWEEN TO_DATE('" & bkmfrom & "','MM/DD/YYYY') AND TO_DATE('" & bkmto & "','MM/DD/YYYY') ) kernelsilosbi," & _
	" a.pktankisi," & _
	" (SELECT Sum(pktankisi) FROM productiontrn b WHERE b.proddate <= a.proddate AND b.fcba=a.fcba and proddate BETWEEN TO_DATE('" & bkmfrom & "','MM/DD/YYYY') AND TO_DATE('" & bkmto & "','MM/DD/YYYY') ) pktankisisbi " & _
	" FROM productiontrn a " & _
	" LEFT JOIN FAC_HPP_HARIAN f ON a.proddate=f.proddate and a.fcba=f.fcba " & _
	" WHERE a.fcba='" & fcba & "' " & _
	" AND a.proddate BETWEEN TO_DATE('" & bkmfrom & "','MM/DD/YYYY') AND TO_DATE('" & bkmto & "','MM/DD/YYYY')"

   'response.write sqlview

    List.Open sqlview, connStr, 1
    ListCount = List.RecordCount 

    Set caljson = new JSON  'JSON.aspê°€ í•„ìš”í•¨P
    Set rs = connStr.Execute(sqlview)

    Dim List_jsonStringlpb
    List_jsonStringlpb = caljson.toJSON("results",rs, false)

    Dim jsonObjgllpb, outputObjgllpb
    set jsonObjgllpb = new JSONobject
    set outputObjgllpb = jsonObjgllpb.parse(List_jsonStringlpb)  'json.objectClass.aspê°€ í•„ìš”í•¨(íŒŒì‹±í•˜ë ¤ë©´)


    Set oJSON = New aspJSON     'aspJSON.asp í•„ìš”í•¨(forë¬¸)
    oJSON.loadJSON(List_jsonStringlpb)

    Response.write List_jsonStringlpb
end if

'qry laporan harian divisi
if idview="26" then
   
    pcompany 	= Request("pcompany")
    fcba 		= Request("pfcba")
    bkmfrom 	= Request("bkmfrom")
    bkmto   	= Request("bkmto")
    divisi   	= Request("divisi")

	bulansplitfrom  = left(bkmfrom,2)
	tglsplitfrom    = mid(bkmfrom,4,2)
	thnsplitfrom    = right(bkmfrom,4)
	tglfrom     	= bulansplitfrom & "/" & tglsplitfrom & "/" & thnsplitfrom
	blnthn 			= bulansplitfrom&thnsplitfrom



    Set List = Server.CreateObject("ADODB.Recordset")

	if trim(divisi) <> "all" then
		filterdiv = " AND h.sectioncode='" & divisi & "'"
	end if


	if trim(fcba) = "-" then 
		filterlhd = " AND h.FCBA in (SELECT fccode from businessunit where fccompanycode='" & pcompany & "' and isactive='TRUE') "
	else
	    filterlhd = " AND h.FCBA ='" & fcba & "' "
	end if



 '   sqlview =  "SELECT h.fcba, d.actcode || ' - ' || j.fcname jobname,h.SECTIONCODE,d.LOCATIONCODE,to_char(f.plantingdate,'yyyy') tahuntanam,f.luas_gross,f.status, " & _
'	" SUM(d.MANDAYS) HK,SUM(d.OUTPUT) FISIK,SUM(d.UNIT) UNIT " & _
'	" FROM bkmhdr h " & _
'	" LEFT JOIN bkmupahdtl d on h.BKMNO = d.BKMNO and h.fcba=d.fcba " & _
'	" LEFT JOIN JOB j on j.fccode=d.ACTCODE and j.FCBA = d.FCBA " & _
'	" INNER JOIN FIELD f on f.fccode=d.LOCATIONCODE and f.fcba=d.fcba " & _
'	" WHERE h.bkmdate between to_date('" & bkmfrom & "','mm/dd/yyyy') AND to_date('" & bkmto & "','mm/dd/yyyy')  " & _
'		filterdiv & _ 
'		filterlhd & _ 
'	
'	" GROUP BY h.fcba, d.actcode || ' - ' || j.fcname,h.SECTIONCODE,d.LOCATIONCODE,f.plantingdate,f.luas_gross,f.status " & _
'	" order by  h.fcba,d.actcode || ' - ' || j.fcname asc"

sqlview = "SELECT t.fcba,t.actcode,t.fcname,t.sectioncode,t.locationcode,t.tahuntanam,t.luas_gross,t.dx_block_status AS STATUS, " & _
	" nvl(sum(t.fisik),0)fisik,nvl(sum(t.unit),0)unit,nvl(sum(t.hk),0) hk,nvl(sum(t.hkshi),0) hkshi,nvl(sum(t.hksbi),0) hksbi " & _
	" from( " & _
	"	SELECT h.fcba, d.actcode,j.fcname,h.SECTIONCODE,d.LOCATIONCODE, " & _
	"	to_char(f.plantingdate,'yyyy') tahuntanam, " & _
	"	f.luas_gross,f.dx_block_status, SUM(d.OUTPUT) FISIK,SUM(d.UNIT) UNIT,SUM(d.MANDAYS) HK,0 HKSHI,0 HKSBI " & _
	"	FROM bkmhdr h " & _
	"	LEFT JOIN bkmupahdtl d on h.BKMNO = d.BKMNO and h.fcba=d.fcba " & _
	"	LEFT JOIN JOB j on j.fccode=d.ACTCODE and j.FCBA = d.FCBA " & _
	"	INNER JOIN FIELD f on f.fccode=d.LOCATIONCODE and f.fcba=d.fcba " & _
	"	WHERE h.bkmdate=to_date('" & bkmfrom & "','mm/dd/yyyy') " & _
		filterdiv & _ 
		filterlhd & _ 
	"	GROUP BY h.fcba, d.actcode,j.fcname,h.SECTIONCODE,d.LOCATIONCODE, " & _
	"	f.plantingdate,f.luas_gross,f.dx_block_status " & _
	"	union " & _
	"	SELECT h.fcba, d.actcode,j.fcname,h.SECTIONCODE,d.LOCATIONCODE, " & _
	"	to_char(f.plantingdate,'yyyy') tahuntanam, " & _
	"	f.luas_gross,f.dx_block_status, SUM(d.OUTPUT) FISIK,SUM(d.UNIT) UNIT,0 HK,SUM(d.MANDAYS) HKSHI,0 HKSBI " & _
	"	FROM bkmhdr h " & _
	"	LEFT JOIN bkmupahdtl d on h.BKMNO = d.BKMNO and h.fcba=d.fcba " & _
	"	LEFT JOIN JOB j on j.fccode=d.ACTCODE and j.FCBA = d.FCBA " & _
	"	INNER JOIN FIELD f on f.fccode=d.LOCATIONCODE and f.fcba=d.fcba  " & _
	"	WHERE to_char(h.bkmdate,'mmyyyy')='" & blnthn & "' " & _
		filterdiv & _ 
		filterlhd & _ 
	"	GROUP BY h.fcba, d.actcode,j.fcname,h.SECTIONCODE,d.LOCATIONCODE, " & _
	"	f.plantingdate,f.luas_gross,f.dx_block_status " & _
	"	union " & _
	"	SELECT h.fcba, d.actcode,j.fcname,h.SECTIONCODE,d.LOCATIONCODE, " & _
	"	to_char(f.plantingdate,'yyyy') tahuntanam, " & _
	"	f.luas_gross,f.dx_block_status, SUM(d.OUTPUT) FISIK,SUM(d.UNIT) UNIT,0 HK,0 HKSHI,SUM(d.MANDAYS) HKSBI " & _
	"	FROM bkmhdr h " & _
	"	LEFT JOIN bkmupahdtl d on h.BKMNO = d.BKMNO and h.fcba=d.fcba " & _
	"	LEFT JOIN JOB j on j.fccode=d.ACTCODE and j.FCBA = d.FCBA " & _
	"	INNER JOIN FIELD f on f.fccode=d.LOCATIONCODE and f.fcba=d.fcba " & _
	"	WHERE to_char(h.bkmdate,'mm')<='" & bulansplitfrom & "' and to_char(h.bkmdate,'yyyy')='" & thnsplitfrom & "' " & _
		filterdiv & _ 
		filterlhd & _ 
	"	GROUP BY h.fcba, d.actcode,j.fcname,h.SECTIONCODE,d.LOCATIONCODE, " & _
	"	f.plantingdate,f.luas_gross,f.dx_block_status " & _
	")t " & _
	"GROUP BY  t.fcba,t.actcode,t.fcname,t.sectioncode,t.locationcode,t.tahuntanam,t.luas_gross,t.dx_block_status " & _
	"order by t.fcba,t.actcode,t.fcname asc"

 ' response.write sqlview

    List.Open sqlview, connStr, 1
    ListCount = List.RecordCount 

    Set caljson = new JSON  'JSON.aspê°€ í•„ìš”í•¨P
    Set rs = connStr.Execute(sqlview)

    Dim List_jsonStringlhd
    List_jsonStringlhd = caljson.toJSON("results",rs, false)

    Dim jsonObjgllhd, outputObjgllhd
    set jsonObjgllhd = new JSONobject
    set outputObjgllhd = jsonObjgllhd.parse(List_jsonStringlhd)  'json.objectClass.aspê°€ í•„ìš”í•¨(íŒŒì‹±í•˜ë ¤ë©´)


    Set oJSON = New aspJSON     'aspJSON.asp í•„ìš”í•¨(forë¬¸)
    oJSON.loadJSON(List_jsonStringlhd)

    Response.write List_jsonStringlhd
end if

'tampilkan arus kas '
if idview="27" then
   ' dim tahun24,bulan24,filter24,posting24,region24
    region24   = Request("region")
    tahun24   = Request("tahun")
 	company24 = Request("company")
	fcba24    = Request("fcba")
	groupby   = Request("groupby")
  
	' if trim(company24) <> "-" then 
	' 	if trim(fcba24) = "-" then 
	' 			sqllist =   " SELECT fccode as fcba " & _
	' 				" FROM businessunit " & _
	' 				" WHERE fccompanycode='" & company24 &"' and isactive='TRUE' " 
					
				
	' 			Set rslistdt= Server.CreateObject("ADODB.Recordset")
	' 			rslistdt.CursorType = 3
	' 			rslistdt.CursorLocation = 3
	' 			rslistdt.Open sqllist, connstr
	' 			if not rslistdt.eof then 
	' 			  sqlexec = "BEGIN " & vbCrLf 
	' 			  Do While Not rslistdt.eof	
	' 				sqlexec = sqlexec & "fac_trialbalance_bulanan_api.ProsesNeraca(" & groupby & ",12, " & tahun24 & ",'" & rslistdt("fcba") & "');"  & vbCrLf 
	' 				rslistdt.MoveNext
	' 			  loop
	' 			End if
	' 			sqlexec = sqlexec & "END; " 
	' 			rslistdt.Close
	' 			set rslistdt = Nothing			  

	' 	else
	' 			sqlexec = "BEGIN " & _ 
	' 					  "fac_trialbalance_bulanan_api.ProsesNeraca(" & groupby & ",12," & tahun24 & ",'" & fcba24 & "');" & _
	' 					  "END; " 
	' 	end if
	' else
	' 			sqllist =   " SELECT fccode as fcba " & _
	' 						" FROM businessunit " & _
	' 						" WHERE isactive='TRUE' Order by fccompanycode,fccode " 
									
	' 			Set rslistdt= Server.CreateObject("ADODB.Recordset")
	' 			rslistdt.CursorType = 3
	' 			rslistdt.CursorLocation = 3
	' 			rslistdt.Open sqllist, connstr
	' 			if not rslistdt.eof then 
	' 			  sqlexec = "BEGIN " & vbCrLf 
	' 			  Do While Not rslistdt.eof	
	' 				sqlexec = sqlexec & "fac_trialbalance_bulanan_api.ProsesNeraca(" & groupby & ",12, " & tahun24 & ",'" & rslistdt("fcba") & "');"  & vbCrLf 
	' 				rslistdt.MoveNext
	' 			  loop
	' 			End if
	' 			sqlexec = sqlexec & "END; " 
	' 			rslistdt.Close
	' 			set rslistdt = Nothing		

	' end if
	' connStr.Execute(sqlexec)
	
	
    Set List = Server.CreateObject("ADODB.Recordset")

	' if region24<> "all" then 
	' end if
	if trim(region24) <> "all" then
	if trim(company24) <> "-" then 
		if trim(fcba24) = "-" then 
		sqlview =   " SELECT year period,acccode,description,rowtype,0 b1qty_,0 b2qty_,0 b3qty_,0 b4qty_,0 b5qty_,0 b6qty_, " & _
					" 0 b7qty_,0 b8qty_,0 b9qty_,0 b10qty_,0 b11qty_,0 b12qty," & _
					" round(sum(total_),0) total_, round(sum(b1_),0) b1_, round(sum(b2_),0) b2_,round(sum(b3_),0) b3_,round(sum(b4_),0) b4_,round(sum(b5_),0) b5_,round(sum(b6_),0) b6_, " & _ 
					" round(sum(b7_),0) b7_,round(sum(b8_),0) b8_,round(sum(b9_),0) b9_,round(sum(b10_),0) b10_,round(sum(b11_),0) b11_,round(sum(b12_),0) b12_, " & _
					" round(sum(b13_),0) b13_,round(sum(b14_),0) b14_,round(sum(b15_),0) b15_" & _
					" FROM fac_trialbalance_bulanan " & _
					" where  " & _
					" year="& tahun24 & _ 
					" AND FCBA in (SELECT fccode from businessunit where fccompanycode='" & company24 & "' and isactive='TRUE') " & _
					" AND rowtype='" & groupby & "'" & _
					" GROUP BY year,acccode,description,rowtype" & _
					" ORDER BY acccode " 


		else
		sqlview =   " SELECT year period,acccode,description,rowtype, 0 b1qty_,0 b2qty_,0 b3qty_,0 b4qty_,0 b5qty_,0 b6qty_, " & _
					" 0 b7qty_,0 b8qty_,0 b9qty_,0 b10qty_,0 b11qty_,0 b12qty," & _
					" round(sum(total_),0) total_, round(sum(b1_),0) b1_, round(sum(b2_),0) b2_,round(sum(b3_),0) b3_,round(sum(b4_),0) b4_,round(sum(b5_),0) b5_,round(sum(b6_),0) b6_, " & _ 
					" round(sum(b7_),0) b7_,round(sum(b8_),0) b8_,round(sum(b9_),0) b9_,round(sum(b10_),0) b10_,round(sum(b11_),0) b11_,round(sum(b12_),0) b12_, " & _
					" round(sum(b13_),0) b13_,round(sum(b14_),0) b14_,round(sum(b15_),0) b15_" & _
					" FROM fac_trialbalance_bulanan " & _
					" where  " & _
					" year="& tahun24 & _ 
					" AND FCBA ='" & fcba24 & "' " & _
					" AND rowtype='" & groupby & "'" & _
					" GROUP BY year,acccode,description,rowtype" & _
					" ORDER BY acccode " 

		end if
	else
		sqlview =   " SELECT year period,acccode,description,rowtype, 0 b1qty_,0 b2qty_,0 b3qty_,0 b4qty_,0 b5qty_,0 b6qty_, " & _
					" 0 b7qty_,0 b8qty_,0 b9qty_,0 b10qty_,0 b11qty_,0 b12qty," & _
					" round(sum(total_),0) total_, round(sum(b1_),0) b1_, round(sum(b2_),0) b2_,round(sum(b3_),0) b3_,round(sum(b4_),0) b4_,round(sum(b5_),0) b5_,round(sum(b6_),0) b6_, " & _ 
					" round(sum(b7_),0) b7_,round(sum(b8_),0) b8_,round(sum(b9_),0) b9_,round(sum(b10_),0) b10_,round(sum(b11_),0) b11_,round(sum(b12_),0) b12_, " & _
					" round(sum(b13_),0) b13_,round(sum(b14_),0) b14_,round(sum(b15_),0) b15_" & _
					" FROM fac_trialbalance_bulanan " & _
					" where  " & _
					" year="& tahun24 & _ 
					" AND FCBA in (SELECT fccode from businessunit where isactive='TRUE' AND region='" & region24 & "') " & _
					" AND rowtype='" & groupby & "'" & _
					" GROUP BY year,acccode,description,rowtype" & _
					" ORDER BY acccode "
	
	end if
	else 
		sqlview =   " SELECT year period,acccode,description,rowtype, 0 b1qty_,0 b2qty_,0 b3qty_,0 b4qty_,0 b5qty_,0 b6qty_, " & _
					" 0 b7qty_,0 b8qty_,0 b9qty_,0 b10qty_,0 b11qty_,0 b12qty," & _
					" round(sum(total_),0) total_, round(sum(b1_),0) b1_, round(sum(b2_),0) b2_,round(sum(b3_),0) b3_,round(sum(b4_),0) b4_,round(sum(b5_),0) b5_,round(sum(b6_),0) b6_, " & _ 
					" round(sum(b7_),0) b7_,round(sum(b8_),0) b8_,round(sum(b9_),0) b9_,round(sum(b10_),0) b10_,round(sum(b11_),0) b11_,round(sum(b12_),0) b12_, " & _
					" round(sum(b13_),0) b13_,round(sum(b14_),0) b14_,round(sum(b15_),0) b15_" & _
					" FROM fac_trialbalance_bulanan " & _
					" where  " & _
					" year="& tahun24 & _ 
					" AND FCBA in (SELECT fccode from businessunit where isactive='TRUE') " & _
					" AND rowtype='" & groupby & "'" & _
					" GROUP BY year,acccode,description,rowtype" & _
					" ORDER BY acccode "
	end if
	
              
 '   response.write sqlview
	' Perbaikan: List.Open/RecordCount (query yang sama dijalankan dua kali, hasil tidak dipakai) dihapus

    Set caljson = new JSON  'JSON.aspê°€ í•„ìš”í•¨
    Set rs = connStr.Execute(sqlview)

   ' Dim List_jsonString24
    List_jsonString24 = caljson.toJSON("results",rs, false)

   ' Dim jsonObj24, outputObj24
    ' Perbaikan: JSON tidak lagi di-parse ulang dengan JSONobject/aspJSON (hasilnya tidak dipakai;
    ' memakan ~80 detik untuk ~950 baris). Output tetap List_jsonString24 yang sama.

    Response.write List_jsonString24
end if

'qry laporan harian divisi
if idview="28" then
   
    noref 		= Request("noref")
   
sqlview = "Select I.SIVCODE, I.STOCKCODE,S.FCNAME,S.UNITOFMEASUREMENT , I.QUANTITY,I.REMARKS,I.LINENOKEY " & _ 
" From   ISSUEDETAILS I " & _ 
" LEFT JOIN STOCK S ON I.STOCKCODE=S.FCCODE AND I.FCBA = S.FCBA " & _ 
" WHERE I.SIVCODE='" & noref & "'"

  'response.write sqlview

   Set List = Server.CreateObject("ADODB.Recordset")

    List.Open sqlview, connStr, 1
    ListCount = List.RecordCount 

    Set caljson = new JSON  'JSON.aspê°€ í•„ìš”í•¨P
    Set rs = connStr.Execute(sqlview)

    Dim List_jsonStringstreturn
    List_jsonStringstreturn = caljson.toJSON("results",rs, false)

    Dim jsonObjglstreturn, outputObjstreturn
    set jsonObjglstreturn = new JSONobject
    set outputObjstreturn = jsonObjglstreturn.parse(List_jsonStringstreturn)  'json.objectClass.aspê°€ í•„ìš”í•¨(íŒŒì‹±í•˜ë ¤ë©´)


    Set oJSON = New aspJSON     'aspJSON.asp í•„ìš”í•¨(forë¬¸)
    oJSON.loadJSON(List_jsonStringstreturn)

    Response.write List_jsonStringstreturn
end if

'qry rkb
if idview="29" then
   
    company 	= Request("company")
    nokontrak 	= Request("nokontrak")
    fcbarkb 	= Request("fcbarkb")
   
	sqlview = "Select D.JOBCODE,F.FCNAME, D.QTY, D.RATE, D.LINENO,H.CONTRACTORCODE,C.FCNAME CONTRACTORNAME " & _
	" From   CONTRACTAGREEMENTDETAIL D " & _
	" LEFT JOIN CONTRACTAGREEMENT H ON D.AGREEMENTCODE=H.AGREEMENTCODE AND H.FCBA=D.FCBA " & _ 
	" LEFT JOIN CONTRACTOR C ON H.CONTRACTORCODE=C.FCCODE AND H.FCBA=C.FCBA " & _
	" LEFT JOIN JOBMASTER F ON D.JOBCODE=F.FCCODE " & _
	" WHERE D.AGREEMENTCODE='" & nokontrak & "' AND D.FCBA='" & fcbarkb & "'"

  'response.write sqlview

   Set List = Server.CreateObject("ADODB.Recordset")

    List.Open sqlview, connStr, 1
    ListCount = List.RecordCount 

    Set caljson = new JSON  'JSON.aspê°€ í•„ìš”í•¨P
    Set rs = connStr.Execute(sqlview)

    Dim List_jsonStringrkb
    List_jsonStringrkb = caljson.toJSON("results",rs, false)

    Dim jsonObjglrkb, outputObjrkb
    set jsonObjglrkb = new JSONobject
    set outputObjrkb = jsonObjglrkb.parse(List_jsonStringrkb)  'json.objectClass.aspê°€ í•„ìš”í•¨(íŒŒì‹±í•˜ë ¤ë©´)


    Set oJSON = New aspJSON     'aspJSON.asp í•„ìš”í•¨(forë¬¸)
    oJSON.loadJSON(List_jsonStringrkb)

    Response.write List_jsonStringrkb
end if


'qry finmod
if idview="30" then
   
	region30 	= Request("region")     
    company30 	= Request("company")
    year30 	    = Request("tahun")
    
	if trim(region30) = "all" then
  sqlview = "	SELECT b.finmodcode,c.region,c.fccompanycode,a.jobgroupcode,jobgroupname,job_forfieldstatus,Sum(qty) qty,Sum(amount) amount,Decode(Sum(qty),0,0,Sum(amount)/Sum(qty)) price " & _
			"	FROM bud_finmod_dtl a " & _
			"	INNER JOIN bud_finmod_hdr b ON (a.finmodcode=b.finmodcode ) " & _
			"	INNER JOIN businessunit c ON (b.fccompanycode=c.fccompanycode) " & _
			"   WHERE b.period_year=" & year30 & "" & _
			"	GROUP BY  b.finmodcode,c.region,c.fccompanycode,a.jobgroupcode,jobgroupname,job_forfieldstatus " & _
			"   ORDER BY region,fccompanycode,jobgroupcode "
	ELSE
	sqlview = "	SELECT b.finmodcode,c.region,c.fccompanycode,a.jobgroupcode,jobgroupname,job_forfieldstatus,Sum(qty) qty,Sum(amount) amount,Decode(Sum(qty),0,0,Sum(amount)/Sum(qty)) price " & _
			"	FROM bud_finmod_dtl a " & _
			"	INNER JOIN bud_finmod_hdr b ON (a.finmodcode=b.finmodcode ) " & _
			"	INNER JOIN businessunit c ON (b.fccompanycode=c.fccompanycode) " & _
			"   WHERE b.period_year=" & year30 & " AND c.region = '" & region30 & "' AND c.fccompanycode like '" & company30 & "%' " & _
			"	GROUP BY b.finmodcode,c.region,c.fccompanycode,a.jobgroupcode,jobgroupname,job_forfieldstatus " & _
			"   ORDER BY region,fccompanycode,jobgroupcode "
	end if



  'response.write sqlview

   Set List = Server.CreateObject("ADODB.Recordset")

    List.Open sqlview, connStr, 1
    ListCount = List.RecordCount 

    Set caljson = new JSON  'JSON.aspê°€ í•„ìš”í•¨P
    Set rs = connStr.Execute(sqlview)

    Dim List_jsonStringfinmod
    List_jsonStringfinmod = caljson.toJSON("results",rs, false)

    Dim jsonObjglfinmod, outputObjfinmod
    set jsonObjglfinmod = new JSONobject
    set outputObjfinmod = jsonObjglfinmod.parse(List_jsonStringfinmod)  'json.objectClass.aspê°€ í•„ìš”í•¨(íŒŒì‹±í•˜ë ¤ë©´)


    Set oJSON = New aspJSON     'aspJSON.asp í•„ìš”í•¨(forë¬¸)
    oJSON.loadJSON(List_jsonStringfinmod)

    Response.write List_jsonStringfinmod
end if

'qry finmod
if idview="31" then
   
	src31 = UCase(Request("stock"))
	
    Set List = Server.CreateObject("ADODB.Recordset")
	sqlview   = " SELECT a.itemcode,c.fcname itemname,d.fccode suppliercode,c.unitofmeasurement,d.fcname suppliername,b.pocode,a.quantity,b.podate,round(a.unitprice,2) unitprice FROM purchaseorderdetails a " & _ 
				" INNER JOIN purchaseorder b ON (a.pocode=b.pocode AND a.fcba=b.fcba)" & _ 
				" INNER JOIN bankstock c ON (a.itemcode=c.fccode)" & _ 
				" INNER JOIN supplier d ON (b.suppliercode=d.fccode AND b.fcba=d.fcba)" & _ 
				" WHERE a.itemcode = '" & src31 & "'" & _ 
				" ORDER BY b.podate DESC" 
	
				
    Set caljson = new JSON  'JSON.aspê°€ í•„ìš”í•¨
    Set rs = connStr.Execute(sqlview)

    Dim List_jsonString31
    List_jsonString31 = caljson.toJSON("results",rs, false)

    Dim jsonObj31, outputObj31
    set jsonObj31 = new JSONobject
    set outputObj31 = jsonObj31.parse(List_jsonString31)  'json.objectClass.aspê°€ í•„ìš”í•¨(íŒŒì‹±í•˜ë ¤ë©´)


    Set oJSON = New aspJSON     'aspJSON.asp í•„ìš”í•¨(forë¬¸)
    oJSON.loadJSON(List_jsonString31)

    Response.write List_jsonString31
end if

'qry srch stock
if idview="32" then
   
	src32_Group	       = Request("src_groupcode")
	src32_SubGroup	   = Request("src_subgroupcode")
	src32_Cari	       = Request("src_stock")
	
	src32_src          = Request("src")
	src32_fcba         = Request("fcba")
	
	Set List  = Server.CreateObject("ADODB.Recordset")
	
	if src32_src = "BStk" then
	sqlview   = " SELECT a.fccode , a.fcname, unitofmeasurement  , stocksubgroupcode, NVL(b.fcname,'-') stocksubgroupname " & _ 
                " FROM bankstock a " & _ 
				" LEFT JOIN (select distinct fccode,fcname from stocksubgroup) b on (a.stocksubgroupcode =  b.fccode )" & _ 
				" WHERE isactive='Y' AND " & _
				" NVL(stockgroupcode,'-') like '" & src32_Group & "%' AND NVL(stocksubgroupcode,'-') like '" & src32_SubGroup & "%' AND (a.fccode like '" & trim(src32_Cari) & "%' or UPPER(a.fcname) like UPPER('%" & trim(src32_Cari) & "%')) " & _ 
				" ORDER BY a.fccode " 
	ELSE
	sqlview   = " SELECT a.fccode , a.fcname, unitofmeasurement  , stocksubgroupcode, NVL(b.fcname,'-') stocksubgroupname " & _ 
                " FROM stock a " & _ 
				" INNER JOIN (select distinct fccode,fcname,fcba from stocksubgroup where fcba='" & src32_fcba & "') b on (a.stocksubgroupcode =  b.fccode and a.fcba=b.fcba)" & _ 
				" WHERE isactive='Y' AND a.fcba='" & src32_fcba & "' AND " & _
				" NVL(stockgroupcode,'-') like '" & src32_Group & "%' AND NVL(stocksubgroupcode,'-') like '" & src32_SubGroup & "%' AND (a.fccode like '" & trim(src32_Cari) & "%' or UPPER(a.fcname) like UPPER('%" & trim(src32_Cari) & "%')) " & _ 
				" ORDER BY a.fccode " 
	end if
	
	'response.write 		sqlview
    Set caljson = new JSON  'JSON.aspê°€ í•„ìš”í•¨
    Set rs = connStr.Execute(sqlview)

    Dim List_jsonString32
    List_jsonString32 = caljson.toJSON("results",rs, false)

    Dim jsonObj32, outputObj32
    set jsonObj32 = new JSONobject
    set outputObj32 = jsonObj32.parse(List_jsonString32)  'json.objectClass.aspê°€ í•„ìš”í•¨(íŒŒì‹±í•˜ë ¤ë©´)


    Set oJSON = New aspJSON     'aspJSON.asp í•„ìš”í•¨(forë¬¸)
    oJSON.loadJSON(List_jsonString32)

    Response.write List_jsonString32
end if

'qry srch stock
if idview="33" then
   
	jobcode	       = Request("jobcode")
		
	Set List  = Server.CreateObject("ADODB.Recordset")

	sqlview   = " SELECT JOB_CATEGORY " & _ 
                " FROM JOBMASTER " & _ 
				" WHERE fccode='" & jobcode & "' " 

	
	'response.write 		sqlview
    Set caljson = new JSON  'JSON.aspê°€ í•„ìš”í•¨
    Set rs = connStr.Execute(sqlview)

    Dim List_jsonString33
    List_jsonString33 = caljson.toJSON("results",rs, false)

    Dim jsonObj33, outputObj33
    set jsonObj33 = new JSONobject
    set outputObj33 = jsonObj33.parse(List_jsonString33)  'json.objectClass.aspê°€ í•„ìš”í•¨(íŒŒì‹±í•˜ë ¤ë©´)


    Set oJSON = New aspJSON     'aspJSON.asp í•„ìš”í•¨(forë¬¸)
    oJSON.loadJSON(List_jsonString33)

    Response.write List_jsonString33
end if

'get no voucher
if idview="34" then
   
	txtbankcode	   = Request("txtbankcode")
	fcba	       = Request("fcba")
	transdate	   = Request("transdate")
	vtranstype	   = Request("vtranstype")

	if vtranstype="02-PAYMENT" then
		var_txtTable = "PAYMENTVOUCHER"
		var_tempx = "K"
	elseif vtranstype="01-RECEIVE" then
		var_txtTable = "RECEIVEVOUCHER"
		var_tempx = "M"
	else
		var_txtTable = "JOURNALVOUCHER"
		var_tempx = "J"
	end if
		
	

	sqlview   = " select Decode( Nvl(Max(cn.PREFIX),'-'), '-', '" & fcba & "'||'/'|| '" & var_tempx &"/' || '"& txtbankcode &"/' || TO_CHAR(TO_DATE('"& transdate &"','mm/dd/yyyy'),'YY') ||'/'|| LPAD(NVL(MAX(regexp_substr(det.VOUCHERCODE, '[^/]+', 1, 5)),0000)+0001,4,0), " & _
				" MAX(cn.PREFIX) || '"& var_tempx &"/' || '"& txtbankcode &"/' || TO_CHAR(TO_DATE('"& transdate &"','mm/dd/yyyy'),'YY') ||'/'|| LPAD(NVL(MAX(regexp_substr(det.VOUCHERCODE, '[^/]+', 1, 5)),0000)+0001,4,0)) DOCNUMBER " & _
				" from " & _
				" controlnumber cn " & _
				" left join paymentvoucher det on cn.fcba = det.fcba and cn.accountingyear = to_number(to_char(det.DATECREATED,'yyyy')) " & _
				" where cn.fcba ='" & fcba & "' " & _
				" and cn.LU_NAME ='"& var_txtTable &"' "&_
				" and cn.accountingyear =  TO_NUMBER(TO_CHAR(TO_DATE('"& transdate &"','mm/dd/yyyy'),'yyyy')) " & _
				" and cn.accountingperiod = TO_NUMBER(TO_CHAR(TO_DATE('"& transdate &"','mm/dd/yyyy'),'mm')) " & _
				" and LPAD(det.vouchercode,Length(rtrim(det.vouchercode))-4) = cn.PREFIX || '"& var_tempx &"/' || '"& txtbankcode &"/' || TO_CHAR(TO_DATE('"& transdate &"','mm/dd/yyyy'),'yy') ||'/' " 

	
	'response.write 		sqlview
	Set List  = Server.CreateObject("ADODB.Recordset")
    Set caljson = new JSON  'JSON.aspê°€ í•„ìš”í•¨
    Set rs = connStr.Execute(sqlview)

    Dim List_jsonString34
    List_jsonString34 = caljson.toJSON("results",rs, false)

    Dim jsonObj34, outputObj34
    set jsonObj34 = new JSONobject
    set outputObj34 = jsonObj34.parse(List_jsonString34)  'json.objectClass.aspê°€ í•„ìš”í•¨(íŒŒì‹±í•˜ë ¤ë©´)


    Set oJSON = New aspJSON     'aspJSON.asp í•„ìš”í•¨(forë¬¸)
    oJSON.loadJSON(List_jsonString34)

    Response.write List_jsonString34
end if

'Kartu Stock
if idview="35" then
    pfcba 			= Request("fcba")
	wrhouse 		= Request("wrhouse")
    pproductstock 	= Request("productstock")
    pdatefrom 	    = Request("datefrom")
    pdateto      	= Request("dateto")

	if trim(wrhouse) = "All" then
        addfilter = " AND a.STORECODE IN (select fccode from store where fcba = '"&pfcba&"' and activation = 'Y')"
		addfilter1 = " AND b.LOCATIONCODE IN (select fccode from store where fcba = '"&pfcba&"' and activation = 'Y')"
		addfilter2 = " AND b.STORECODE IN (select fccode from store where fcba = '"&pfcba&"' and activation = 'Y')"
		addfilter3 = " AND b.STORECODEFROM IN (select fccode from store where fcba = '"&pfcba&"' and activation = 'Y')"
	else
		addfilter = " AND a.STORECODE = '"&wrhouse&"'"
		addfilter1 = " AND b.LOCATIONCODE = '"&wrhouse&"'"
		addfilter2 = " AND b.STORECODE = '"&wrhouse&"'"
		addfilter3 = " AND b.STORECODEFROM = '"&wrhouse&"'"
    end if

	' Response.write Day(formattedDate)

	' CEK DATA TERLEBIH DAHULU
	sqlSelect = "select count(*) as ttl FROM STOCK_HISTORY " & _
				" where fcba = '"& pfcba &"' and itemcode = '"& pproductstock &"' " & _
				" AND TGLTRX BETWEEN '"& pdatefrom &"' " & _
				" AND '"& pdateto &"' "
	' Response.write sqlSelect
	Set cek = connStr.Execute(sqlSelect)
	
	' SUBSTR(a.fccode, 0, 1) = '1'

	' JIKA HASIL PENGECEKAN = 0 atau = 1, MAKA PROSES DELETE DAN INSERT KE TABLE STOCK_HISTORY
	If cek("ttl") = "0" Or cek("ttl") = "1" Then
		sqlDelete = " DELETE FROM STOCK_HISTORY " & _
					" where fcba = '"& pfcba &"' and itemcode = '"& pproductstock &"' " & _
					" AND TGLTRX BETWEEN '"& pdatefrom &"' " & _
					" AND '"& pdateto &"' "
		connStr.Execute(sqlDelete)

		sqlInsert = " INSERT INTO STOCK_HISTORY (" & _
						" TGLTRX, TYPE, ITEMCODE, ORIGIN, DEST, " & _
						" INQTY, OUTQTY, FCBA, STORE, NOTE, " & _
						" FCENTRY, STOCKAWAL, STOCKAKHIR, TRXCODE, COMPANYNAME, NOMINAL" & _
					" )" & _
					" SELECT " & _
					" TO_CHAR(a.SIVDATE, 'MM/DD/YYYY') tgl, 'OUT' TYPE, b.STOCKCODE, '-' origin, b.LOCATIONCODE dest," & _
					" 0 INQTY, b.QUANTITY OUTQTY, c.fccode, a.STORECODE STORE, 'ISSUE' modul," & _
					" b.FCENTRY, 0 STOCKAWAL, 0 STOCKAKHIR, " & _ 
					" a.SIVCODE, c.fccompanycode, case when b.amount = '0' then 0 else b.amount end amount" & _
					" from ISSUE a " & _
					" inner join ISSUEDETAILS b on a.SIVCODE = b.SIVCODE " & _
					" inner join businessunit c on a.FCBA = c.fccode " & _
					" where c.fccode = '"& pfcba &"' and b.stockcode = '"& pproductstock &"' " & _
					" AND a.SIVDATE BETWEEN TO_DATE('"& pdatefrom &"', 'MM/DD/YYYY')" & _
					" AND TO_DATE('"& pdateto &"', 'MM/DD/YYYY') AND a.ROWSTATE = 'Approved' "&addfilter&"" & _
					" UNION ALL " & _
					" SELECT " & _
					" TO_CHAR(a.RNDATE, 'MM/DD/YYYY') tgl, 'IN' TYPE, b.PURCHASEITEMCODE, b.LOCATIONCODE origin, '-' dest," & _
					" b.QUANTITY INQTY, 0 OUTQTY, c.fccode, a.STORECODE STORE, 'RECEIVED NOTE' modul," & _
					" b.FCENTRY, 0 STOCKAWAL, 0 STOCKAKHIR, " & _ 
					" a.RECEIVENOTECODE, c.fccompanycode, case when b.AMOUNTRECEIVED = '0' then 0 else b.AMOUNTRECEIVED end amount" & _
					" from RECEIVENOTE a " & _
					" inner join RECEIVENOTEDETAIL b on a.RECEIVENOTECODE = b.RECEIVENOTECODE " & _
					" inner join businessunit c on a.FCBA = c.fccode " & _
					" where c.fccode = '"& pfcba &"' " & _
					" and b.PURCHASEITEMCODE = '"& pproductstock &"' AND a.RNDATE BETWEEN TO_DATE('"& pdatefrom &"', 'MM/DD/YYYY') " & _
					" AND TO_DATE('"& pdateto &"', 'MM/DD/YYYY') "&addfilter&"" & _
					" UNION ALL " & _
					" SELECT " & _
					" TO_CHAR(a.ORDERDATE, 'MM/DD/YYYY') tgl, 'OUT' TYPE, b.stockcode, '-' origin, b.LOCATIONCODE dest," & _
					" 0 INQTY, b.QUANTITYTRANSFERED OUTQTY, c.fccode, b.STORECODE STORE, 'TRANSPORT TASK' modul," & _
					" b.FCENTRY, 0 STOCKAWAL, 0 STOCKAKHIR, a.ORDERNO, c.fccompanycode, case when b.amount = '0' then 0 else b.amount end amount" & _
					" from TRANSPORTTASK_TAB a " & _
					" inner join TRANSPORTTASKDTL_TAB b on a.ORDERNO = b.ORDERNO " & _
					" inner join businessunit c on a.FCBA = c.fccode " & _
					" where c.fccode = '"& pfcba &"' and b.stockcode = '"& pproductstock &"' " & _
					" AND a.ORDERDATE BETWEEN TO_DATE('"& pdatefrom &"', 'MM/DD/YYYY') " & _
					" AND TO_DATE('"& pdateto &"', 'MM/DD/YYYY') AND a.ROWSTATEID = 'Approved' "&addfilter2&"" & _
					" UNION ALL " & _
					" SELECT " & _
					" TO_CHAR(a.ORDERDATE, 'MM/DD/YYYY') tgl, 'IN' TYPE, b.stockcode, '-' origin, b.LOCATIONCODE dest," & _
					" b.QUANTITYRECEIVED INQTY, 0 OUTQTY, c.fccode, b.FCBASOURCE STORE, 'RECEIVED TASK' modul," & _
					" b.FCENTRY, 0 STOCKAWAL, 0 STOCKAKHIR, a.ORDERNO, c.fccompanycode, case when b.AMOUNTRECEIVED = '0' then 0 else b.AMOUNTRECEIVED end amount" & _
					" from RECEIVEDTASK_TAB a " & _
					" inner join RECEIVEDTASKDTL_TAB b on a.ORDERNO = b.ORDERNO " & _
					" inner join businessunit c on a.FCBA = c.fccode " & _
					" where c.fccode = '"& pfcba &"' and b.stockcode = '"& pproductstock &"' " & _
					" AND a.ORDERDATE BETWEEN TO_DATE('"& pdatefrom &"', 'MM/DD/YYYY') " & _
					" AND TO_DATE('"& pdateto &"', 'MM/DD/YYYY') "&addfilter1&"" & _
					" UNION ALL " & _
					" SELECT TO_CHAR(a.ORDERDATE, 'MM/DD/YYYY') tgl, " & _
					" 'OUT' TYPE, " & _
					" b.stockcode, " & _
					" b.STORECODEFROM origin, " & _
					" b.STORECODETO dest, " & _
					" 0 INQTY, " & _
					" b.QUANTITYTRANSFERED OUTQTY, " & _
					" c.fccode, " & _
					" b.STORECODEFROM STORE, " & _
					" 'TRANSPORT TASK DIV' modul, " & _
					" b.FCENTRY, " & _
					" 0 STOCKAWAL, " & _
					" 0 STOCKAKHIR, " & _
					" a.ORDERNO, " & _
					" c.fccompanycode, " & _
					" case when b.amount = 0 or b.amount = '' or b.amount is null then 0 else b.amount end amount " & _
					" FROM TRANSPORTTASKDIV_TAB a " & _
					" INNER JOIN TRANSPORTTASKDIVDTL_TAB b ON a.ORDERNO = b.ORDERNO " & _
					" INNER JOIN businessunit c ON a.FCBA = c.fccode " & _
					" WHERE c.fccode = '"& pfcba &"' " & _
					" AND b.stockcode = '"& pproductstock &"' " & _
					" AND a.ORDERDATE BETWEEN TO_DATE('"& pdatefrom &"', 'MM/DD/YYYY') AND TO_DATE('"& pdateto &"', 'MM/DD/YYYY') " & _
					addfilter3
		' Response.write sqlInsert
		Dim rowsAffected
		connStr.Execute sqlInsert, rowsAffected

		If rowsAffected > 0 Then
			' insert new line from transporttask divisi header & detail
			sqlInsertTrDiv	= " INSERT INTO STOCK_HISTORY (" & _
							" TGLTRX, TYPE, ITEMCODE, ORIGIN, DEST, " & _
							" INQTY, OUTQTY, FCBA, STORE, NOTE, " & _
							" FCENTRY, STOCKAWAL, STOCKAKHIR, TRXCODE, COMPANYNAME, NOMINAL" & _
							" )" & _
							" SELECT TO_CHAR(a.ORDERDATE, 'MM/DD/YYYY') tgl, " & _
							" 'IN' TYPE, " & _
							" b.stockcode, " & _
							" b.STORECODETO origin, " & _
							" b.STORECODEFROM dest, " & _
							" b.QUANTITYTRANSFERED INQTY, " & _
							" 0 OUTQTY, " & _
							" c.fccode, " & _
							" b.STORECODEFROM STORE, " & _
							" 'TRANSPORT TASK DIV' modul, " & _
							" b.FCENTRY, " & _
							" 0 STOCKAWAL, " & _
							" 0 STOCKAKHIR, " & _
							" a.ORDERNO || '-RCV'," & _
							" c.fccompanycode, " & _
							" case when b.amount = 0 or b.amount = '' or b.amount is null then 0 else b.amount end amount " & _
							" FROM TRANSPORTTASKDIV_TAB a " & _
							" INNER JOIN TRANSPORTTASKDIVDTL_TAB b ON a.ORDERNO = b.ORDERNO " & _
							" INNER JOIN businessunit c ON a.FCBA = c.fccode " & _
							" WHERE c.fccode = '"& pfcba &"' " & _
							" AND b.stockcode = '"& pproductstock &"' " & _
							" AND a.ORDERDATE BETWEEN TO_DATE('"& pdatefrom &"', 'MM/DD/YYYY') AND TO_DATE('"& pdateto &"', 'MM/DD/YYYY') " & _
							" AND b.STORECODETO != 'ST1'"
			Set insrt1 = connStr.Execute(sqlInsertTrDiv)
		end if
	end if

	Response.ContentType = "application/json"
	Response.Write "{""status"":""success"",""message"":""Data berhasil ditampilkan""}"
end if

'GET SALDO
if idView="36" Then
	pfcba 			= Request("fcba")
    pproductstock 	= Request("productstock")
	pdatefrom      	= Request("datefrom")
    pdateto      	= Request("dateto")
	wrhouse 		= Request("wrhouse")

	if trim(wrhouse) = "All" then
		addfilterSelect = " AND t.storecode IN (select fccode from store where fcba = '"&pfcba&"' and activation = 'Y') and storecode = 'ST1'"
	else
		addfilterSelect = " AND t.storecode = '"&wrhouse&"'"
    end if

	Dim formattedDate1
	formattedDate1 = Right("0" & Month(pdatefrom), 2) & "/" & Right("0" & Day(pdatefrom), 2) & "/" & Year(pdatefrom)

	if Day(formattedDate) = "1" then
		setWhereDate = ""
	else
		setWhereDate = ""
	end if
	
	sqlSaldoAwal = " SELECT t.storecode, " & _
						" store_api.get_fcname(t.storecode, t.fcba) storename, " & _
						" t.STOCKCODE, " & _
						" stock_api.get_stockname(t.STOCKCODE, t.fcba) stockname, " & _
						" stock_api.get_uom(t.STOCKCODE, t.fcba) uom, " & _
						" stock_api.get_StockSubGroup(t.STOCKCODE, t.fcba) StockSubGroup, " & _
						" sum(nvl(t.openingbalanceqty, 0)) AS openingbalanceqty, " & _
						" sum(nvl(t.receivenoteqty, 0)) AS receivenoteqty, " & _
						" sum(nvl(t.issueqty, 0)) AS issueqty, " & _
						" sum(nvl(t.quantitytransfered, 0)) AS quantitytransfered, " & _
						" sum(nvl(t.quantityreceived, 0)) AS quantityreceived, " & _
						" sum(nvl(t.qtywriteoff, 0)) AS qtywriteoff, " & _
						" (sum(nvl(t.openingbalanceqty, 0))+sum(nvl(t.receivenoteqty, 0))-sum(nvl(t.issueqty, 0))-sum(nvl(t.quantitytransfered, 0))+sum(nvl(t.quantityreceived, 0))+sum(nvl(t.qtywriteoff, 0))) AS qty, " & _
						" sum(nvl(t.openingbalanceamt, 0)) AS amt " & _
						" FROM " & _
						" (SELECT fcba, " & _
						" l.STORECODE, " & _
						" l.itemcode STOCKCODE, " & _
						" l.openingbalanceqty, " & _
						" l.openingbalanceamt, " & _
						" 0 AS receivenoteqty, " & _
						" 0 AS issueqty, " & _
						" 0 AS quantitytransfered, " & _
						" 0 AS quantityreceived, " & _
						" 0 AS qtywriteoff " & _
						" FROM stockledger l " & _
						" INNER JOIN bankstock k ON (l.itemcode=k.fccode) " & _
						" WHERE k.isactive='Y' " & _
						" AND l.YEAR=TO_CHAR(TO_DATE('"& pdatefrom &"', 'MM/DD/YYYY'), 'YYYY') " & _
						" AND l.MONTH=TO_CHAR(TO_DATE('"& pdatefrom &"', 'MM/DD/YYYY'), 'MM') " & _
						" AND l.fcba='"&pfcba&"' " & _
						" UNION ALL SELECT r.fcba, " & _
						"                   r.storecode, " & _
						"                   rd.purchaseitemcode STOCKCODE, " & _
						"                   0 AS openingbalanceqty, 0 AS openingbalanceamt," & _
						"                   rd.quantity AS receivenoteqty, " & _
						"                   0 AS issueqty, " & _
						"                   0 AS quantitytransfered, " & _
						"                   0 AS quantityreceived, " & _
						"                   0 AS qtywriteoff " & _
						"  FROM receivenote r, " & _
						"       receivenotedetail rd " & _
						"  WHERE r.receivenotecode=rd.receivenotecode " & _
						"    AND r.rndate between to_date('"&pdatefrom&"', 'mm/dd/yyyy') and to_date('"&pdateto&"', 'mm/dd/yyyy') " & _
						"    AND r.fcba='"&pfcba&"' " & _
						"  UNION ALL SELECT h.fcba, " & _
						"                   h.storecode, " & _
						"                   d.STOCKCODE, " & _
						"                   0 AS openingbalanceqty, 0 AS openingbalanceamt, " & _
						"                   0 AS receivenoteqty, " & _
						"                   d.quantity AS issueqty, " & _
						"                   0 AS quantitytransfered, " & _
						"                   0 AS quantityreceived, " & _
						"                   0 AS qtywriteoff " & _
						"  FROM ISSUE h, " & _
						"             issuedetails d " & _
						"  WHERE h.sivcode=d.sivcode " & _
						"    AND h.fcba='"&pfcba&"' " & _
						"    AND h.sivdate between to_date('"&pdatefrom&"', 'mm/dd/yyyy') and to_date('"&pdateto&"', 'mm/dd/yyyy') " & _
						" UNION ALL SELECT d.fcba, " & _
						"                  d.storecode, " & _
						"                  d.STOCKCODE, " & _
						"                  0 AS openingbalanceqty, 0 AS openingbalanceamt," & _
						"                  0 AS receivenoteqty, " & _
						"                  0 AS issueqty, " & _
						"                  d.quantitytransfered, " & _
						"                  0 AS quantityreceived, " & _
						"                  0 AS qtywriteoff " & _
						" FROM transporttaskdtl_tab d, " & _
						"      transporttask_tab h " & _
						" WHERE h.orderno=d.orderno " & _
						"   AND h.rowstateid NOT IN('Rejected') " & _
						"   AND h.fcba='"&pfcba&"' " & _
						"   AND h.orderdate between to_date('"&pdatefrom&"', 'mm/dd/yyyy') and to_date('"&pdateto&"', 'mm/dd/yyyy') " & _
						" UNION ALL SELECT d.fcba, " & _
						"                  d.locationcode, " & _
						"                  d.STOCKCODE, " & _
						"                  0 AS openingbalanceqty, 0 AS openingbalanceamt, " & _
						"                  0 AS receivenoteqty, " & _
						"                  0 AS issueqty, " & _
						"                  0 AS quantitytransfered, " & _
						"                  d.quantityreceived, " & _
						"                  0 AS qtywriteoff " & _
						" FROM receivedtask_tab h, " & _
						"      receivedtaskdtl_tab d " & _
						" WHERE h.orderno=d.orderno " & _
						"   AND h.orderdate between to_date('"&pdatefrom&"', 'mm/dd/yyyy') and to_date('"&pdateto&"', 'mm/dd/yyyy') " & _
						"   AND h.fcba='"&pfcba&"' " & _
						" UNION ALL SELECT h.fcba, " & _
						"                  h.STORECODE locationcode, " & _
						"                  d.stockcode, " & _
						"                  0 AS openingbalanceqty, 0 AS openingbalanceamt, " & _
						"                  0 AS receivenoteqty, " & _
						"                  0 AS issueqty, " & _
						"                  0 AS quantitytransfered, " & _
						"                  0 AS quantityreceived, " & _
						"                  d.quantity AS qtywriteoff " & _
						" FROM writeoff h, " & _
						"      writeoffdetails d " & _
						" WHERE h.writeoffcode=d.writeoffcode " & _
						"   AND h.wodate between to_date('"&pdatefrom&"', 'mm/dd/yyyy') and to_date('"&pdateto&"', 'mm/dd/yyyy') " & _
						"   AND h.fcba='"&pfcba&"' " & _
						" UNION ALL SELECT h.fcba, " & _
						"                  h.STORE locationcode, " & _
						"                  itemcode stockcode, " & _
						"                  0 AS openingbalanceqty, 0 AS openingbalanceamt, " & _
						"                  0 AS receivenoteqty, " & _
						"                  0 AS issueqty, " & _
						"                  0 AS quantitytransfered, " & _
						"                  qtyreturn AS quantityreceived, " & _
						"                  0 AS qtywriteoff " & _
						" FROM returnissuehdr h, " & _
						"      returnissuedtl d " & _
						" WHERE h.returnno=d.returnno " & _
						"   AND h.returndate between to_date('"&pdatefrom&"', 'mm/dd/yyyy') and to_date('"&pdateto&"', 'mm/dd/yyyy') " & _
						"      AND h.fcba = '"&pfcba&"')t " & _
						" WHERE t.STOCKCODE = '"&pproductstock&"' " & _
						addfilterSelect & _
						" GROUP BY fcba, " & _
						"          storecode, " & _
						"          stockcode " & _
						" ORDER BY storecode, " & _
						"          stockcode "

	' response.write sqlSaldoAwal
	Set rs = connStr.Execute(sqlSaldoAwal)

	'=== Convert Recordset ke JSON
	Set caljson = New JSON
	Dim dataJSON
	dataJSON = caljson.toJSON("opening", rs, False)

	'=== Bungkus ke object JSON baru
	Dim wrapper
	wrapper = "{""status"":""success"",""message"":""Data berhasil ditampilkan"",""results"":" & dataJSON & "}"

	Response.ContentType = "application/json"
	Response.Write wrapper
end if

if idView="37" Then
	pfcba 			= Request("fcba")
	wrhouse 		= Request("wrhouse")
    pproductstock 	= Request("productstock")
    pdatefrom 	    = Request("datefrom")
    pdateto      	= Request("dateto")

	if trim(wrhouse) = "All" then
        addfilter = " AND a.STORECODE IN (select fccode from store where fcba = '"&pfcba&"' and activation = 'Y')"
		addfilter1 = " AND b.LOCATIONCODE IN (select fccode from store where fcba = '"&pfcba&"' and activation = 'Y')"
		addfilter2 = " AND b.STORECODE IN (select fccode from store where fcba = '"&pfcba&"' and activation = 'Y')"
		addfilter3 = " AND b.STORECODEFROM IN (select fccode from store where fcba = '"&pfcba&"' and activation = 'Y')"
	else
		addfilter = " AND a.STORECODE = '"&wrhouse&"'"
		addfilter1 = " AND b.LOCATIONCODE = '"&wrhouse&"'"
		addfilter2 = " AND b.STORECODE = '"&wrhouse&"'"
		addfilter3 = " AND b.STORECODEFROM = '"&wrhouse&"'"
    end if

	Dim formattedDate
	formattedDate = Right("0" & Month(pdatefrom), 2) & "/" & Right("0" & Day(pdatefrom), 2) & "/" & Year(pdatefrom)
	' response.write Day(formattedDate)
	' PROSES PENGAMBILAN DATA KE TABLE STOCK_HISTORY DENGAN KONDISIONAL BY TANGGAL
	sqlView = " SELECT " & _
			" TO_CHAR(a.SIVDATE, 'MM/DD/YYYY') tgl, 'OUT' TYPE, 4 as urut, b.STOCKCODE, '-' origin, b.LOCATIONCODE dest," & _
			" 0 INQTY, b.QUANTITY OUTQTY, c.fccode, a.STORECODE STORE, 'ISSUE' modul," & _
			" b.FCENTRY, 0 STOCKAWAL, 0 STOCKAKHIR, " & _ 
			" a.SIVCODE, c.fccompanycode, case when b.amount = '0' then 0 else b.amount end amount" & _
			" from ISSUE a " & _
			" inner join ISSUEDETAILS b on a.SIVCODE = b.SIVCODE " & _
			" inner join businessunit c on a.FCBA = c.fccode " & _
			" where c.fccode = '"& pfcba &"' and b.stockcode = '"& pproductstock &"' " & _
			" AND a.SIVDATE BETWEEN TO_DATE('"& pdatefrom &"', 'MM/DD/YYYY')" & _
			" AND TO_DATE('"& pdateto &"', 'MM/DD/YYYY') AND a.ROWSTATE = 'Approved' "&addfilter&"" & _
			" UNION ALL " & _
			" SELECT " & _
			" TO_CHAR(a.RNDATE, 'MM/DD/YYYY') tgl, 'IN' TYPE, 1 as urut, b.PURCHASEITEMCODE, b.LOCATIONCODE origin, '-' dest," & _
			" b.QUANTITY INQTY, 0 OUTQTY, c.fccode, a.STORECODE STORE, 'RECEIVED NOTE' modul," & _
			" b.FCENTRY, 0 STOCKAWAL, 0 STOCKAKHIR, " & _ 
			" a.RECEIVENOTECODE, c.fccompanycode, case when b.AMOUNTRECEIVED = '0' then 0 else b.AMOUNTRECEIVED end amount" & _
			" from RECEIVENOTE a " & _
			" inner join RECEIVENOTEDETAIL b on a.RECEIVENOTECODE = b.RECEIVENOTECODE " & _
			" inner join businessunit c on a.FCBA = c.fccode " & _
			" where c.fccode = '"& pfcba &"' " & _
			" and b.PURCHASEITEMCODE = '"& pproductstock &"' AND a.RNDATE BETWEEN TO_DATE('"& pdatefrom &"', 'MM/DD/YYYY') " & _
			" AND TO_DATE('"& pdateto &"', 'MM/DD/YYYY') AND a.ROWSTATEID = 'Approved' "&addfilter&"" & _
			" UNION ALL " & _
			" SELECT " & _
			" TO_CHAR(a.ORDERDATE, 'MM/DD/YYYY') tgl, 'OUT' TYPE, 5 as urut, b.stockcode, '-' origin, b.LOCATIONCODE dest," & _
			" 0 INQTY, b.QUANTITYTRANSFERED OUTQTY, c.fccode, b.STORECODE STORE, 'TRANSPORT TASK' modul," & _
			" b.FCENTRY, 0 STOCKAWAL, 0 STOCKAKHIR, a.ORDERNO, c.fccompanycode, case when b.amount = '0' then 0 else b.amount end amount" & _
			" from TRANSPORTTASK_TAB a " & _
			" inner join TRANSPORTTASKDTL_TAB b on a.ORDERNO = b.ORDERNO " & _
			" inner join businessunit c on a.FCBA = c.fccode " & _
			" where c.fccode = '"& pfcba &"' and b.stockcode = '"& pproductstock &"' " & _
			" AND a.ORDERDATE BETWEEN TO_DATE('"& pdatefrom &"', 'MM/DD/YYYY') " & _
			" AND TO_DATE('"& pdateto &"', 'MM/DD/YYYY') AND a.ROWSTATEID = 'Approved' "&addfilter2&"" & _
			" UNION ALL " & _
			" SELECT " & _
			" TO_CHAR(a.ORDERDATE, 'MM/DD/YYYY') tgl, 'IN' TYPE, 2 as urut, b.stockcode, '-' origin, b.LOCATIONCODE dest," & _
			" b.QUANTITYRECEIVED INQTY, 0 OUTQTY, c.fccode, b.FCBASOURCE STORE, 'RECEIVED TASK' modul," & _
			" b.FCENTRY, 0 STOCKAWAL, 0 STOCKAKHIR, a.ORDERNO, c.fccompanycode, case when b.AMOUNTRECEIVED = '0' then 0 else b.AMOUNTRECEIVED end amount" & _
			" from RECEIVEDTASK_TAB a " & _
			" inner join RECEIVEDTASKDTL_TAB b on a.ORDERNO = b.ORDERNO " & _
			" inner join businessunit c on a.FCBA = c.fccode " & _
			" where c.fccode = '"& pfcba &"' and b.stockcode = '"& pproductstock &"' " & _
			" AND a.ORDERDATE BETWEEN TO_DATE('"& pdatefrom &"', 'MM/DD/YYYY') " & _
			" AND TO_DATE('"& pdateto &"', 'MM/DD/YYYY') AND a.ROWSTATE = 'Approved' "&addfilter1&"" & _
			" UNION ALL " & _
			" SELECT TO_CHAR(a.ORDERDATE, 'MM/DD/YYYY') tgl, " & _
			" 'OUT' TYPE, 3 as urut, " & _
			" b.stockcode, " & _
			" b.STORECODEFROM origin, " & _
			" b.STORECODETO dest, " & _
			" 0 INQTY, " & _
			" b.QUANTITYTRANSFERED OUTQTY, " & _
			" c.fccode, " & _
			" b.STORECODEFROM STORE, " & _
			" 'TRANSPORT TASK DIV' modul, " & _
			" b.FCENTRY, " & _
			" 0 STOCKAWAL, " & _
			" 0 STOCKAKHIR, " & _
			" a.ORDERNO, " & _
			" c.fccompanycode, " & _
			" case when b.amount = 0 or b.amount = '' or b.amount is null then 0 else b.amount end amount " & _
			" FROM TRANSPORTTASKDIV_TAB a " & _
			" INNER JOIN TRANSPORTTASKDIVDTL_TAB b ON a.ORDERNO = b.ORDERNO " & _
			" INNER JOIN businessunit c ON a.FCBA = c.fccode " & _
			" WHERE c.fccode = '"& pfcba &"' AND a.ROWSTATE = 'Approved' " & _
			" AND b.stockcode = '"& pproductstock &"' " & _
			" AND a.ORDERDATE BETWEEN TO_DATE('"& pdatefrom &"', 'MM/DD/YYYY') AND TO_DATE('"& pdateto &"', 'MM/DD/YYYY') " & _
			addfilter3 & _
			"ORDER BY tgl ASC"
	' sqlView = 	" select  " & _ 
	' 			" tgltrx, type, itemcode, origin, dest, inqty, outqty, fcba, " & _ 
	' 			" store, note, fcentry, stockawal, stockakhir, trxcode, nominal, companyname" & _
	' 			" from stock_history" & _
	' 			" where fcba = '"& pfcba &"' and itemcode = '"& pproductstock &"' " & _
	' 			" AND TO_DATE(tgltrx,'MM/DD/YYYY') between TO_DATE('"&pdatefrom&"','MM/DD/YYYY') and TO_DATE('"&pdateto&"','MM/DD/YYYY') " & _
	' 			addfilterSelect & _
	' 			" ORDER BY TO_DATE(tgltrx, 'MM/DD/YYYY') ASC "

	
	' response.write 		sqlview
	Set List  = Server.CreateObject("ADODB.Recordset")
	Set caljson = new JSON  'JSON.aspê°€ í•„ìš”í•¨
	Set rs = connStr.Execute(sqlview)

	Dim List_jsonString37
	List_jsonString37 = caljson.toJSON("results",rs, false)

	Dim jsonObj37, outputObj37
	set jsonObj37 = new JSONobject
	set outputObj37 = jsonObj37.parse(List_jsonString37)  'json.objectClass.aspê°€ í•„ìš”í•¨(íŒŒì‹±í•˜ë ¤ë©´)


	Set oJSON = New aspJSON     'aspJSON.asp í•„ìš”í•¨(forë¬¸)
	oJSON.loadJSON(List_jsonString37)

	Response.write List_jsonString37
end if
if idView="38" Then
	pfcba 			= Request("fcba")
	wrhouse 		= Request("wrhouse")
    pdatefrom 	    = Request("datefrom")
    pdateto      	= Request("dateto")

	if trim(wrhouse) = "All" then
		' addfilterSelect = " AND STORE IN (select fccode from store where fcba = '"&pfcba&"' and activation = 'Y') OR STORE = '-'"
		addfilterSelect = " AND l.storecode IN (select fccode from store where fcba = '"&pfcba&"' and activation = 'Y')"
		addfilterSelect1 = " AND rd.locationcode IN (select fccode from store where fcba = '"&pfcba&"' and activation = 'Y')"
		addfilterSelect2 = " AND d.locationcode IN (select fccode from store where fcba = '"&pfcba&"' and activation = 'Y')"
		addfilterSelect3 = " AND ard.locationcode IN (select fccode from store where fcba = '"&pfcba&"' and activation = 'Y')"
		selects = ", SUM(NVL(t.TERIMA_HI, 0)) - SUM(NVL(t.TERIMASD_HI, 0)) AS TERIMA_HI, SUM(NVL(t.TERIMASD_HI, 0)) AS TERIMASD_HI "
		selects1 = ", SUM(NVL(t.KELUAR_HI, 0)) - SUM(NVL(t.KELUARSD_HI, 0)) AS KELUAR_HI, SUM(NVL(t.KELUARSD_HI, 0)) AS KELUARSD_HI "
	else
		addfilterSelect = " AND l.storecode = '"&wrhouse&"'"
		addfilterSelect1 = " AND rd.locationcode = '"&wrhouse&"'"
		addfilterSelect2 = " AND d.locationcode = '"&wrhouse&"'"
		addfilterSelect3 = " AND ard.locationcode = '"&wrhouse&"'"
		selects = ", SUM(NVL(t.TERIMA_HI, 0)) AS TERIMA_HI, 0 AS TERIMASD_HI"
		selects1 = ", SUM(NVL(t.KELUAR_HI, 0)) AS KELUAR_HI, 0 AS KELUARSD_HI"
    end if

	Dim formattedDate2
	formattedDate2 = Right("0" & Month(pdatefrom), 2) & "/" & Right("0" & Day(pdatefrom), 2) & "/" & Year(pdatefrom)
	' response.write Day(formattedDate)
	' TEST PUSH KE GITHUB
	' PROSES PENGAMBILAN DATA KE TABLE STOCK_HISTORY DENGAN KONDISIONAL BY TANGGAL
	' if trim(wrhouse) = "All" then
	query = " SELECT t.fcba," & _
			" s.fcname stockname, " & _
			" t.STOCKCODE, " & _
			" t.STORECODE, " & _
			" s.UNITOFMEASUREMENT satuan, " & _
			" s.stockgroupcode || ' - ' || s.fcname stockgroup, " & _
			" (select sum(openingbalanceqty) from stockledger where itemcode in (t.STOCKCODE) and fcba = '"&pfcba&"' and year = TO_CHAR(TO_DATE('"&pdatefrom&"', 'MM/DD/YYYY'), 'YYYY') and month = TO_CHAR(TO_DATE('"&pdateto&"', 'MM/DD/YYYY'), 'MM')) openingbalanceqty " & _
			selects & _
			selects1 & _
			" FROM " & _
			" (SELECT r.fcba, " & _
			" '' AS STORECODE, " & _
			" rd.purchaseitemcode STOCKCODE, " & _
			" 0 AS openingbalanceqty, " & _
			" rd.QUANTITY AS TERIMA_HI, 0 AS TERIMASD_HI, " & _
			" 0 AS KELUAR_HI, 0 AS KELUARSD_HI " & _
			" FROM receivenote r, " & _
			" 		receivenotedetail rd " & _
			" WHERE r.receivenotecode=rd.receivenotecode " & _
			" 	AND r.rndate between to_date('"&pdatefrom&"', 'mm/dd/yyyy') AND to_date('"&pdateto&"', 'mm/dd/yyyy') " & _
			" 	AND r.fcba='"&pfcba&"' " & _
			addfilterSelect1 & _
			" UNION ALL SELECT ar.fcba, " & _
			" '' AS STORECODE, " & _
			" ard.stockcode, " & _
			" 0 AS openingbalanceqty, " & _
			" ard.QUANTITYRECEIVED AS TERIMA_HI, 0 AS TERIMASD_HI, " & _
			" 0 AS KELUAR_HI, 0 AS KELUARSD_HI " & _
			" FROM receivedtask_tab ar " & _
			" JOIN receivedtaskdtl_tab ard ON ar.orderno = ard.orderno " & _
			" WHERE ar.orderdate BETWEEN to_date('"&pdatefrom&"', 'mm/dd/yyyy') AND to_date('"&pdateto&"', 'mm/dd/yyyy') " & _
			" AND ar.fcba = '"&pfcba&"'" & _
			addfilterSelec3 & _
			" UNION ALL SELECT h.fcba, " & _
			" '' AS STORECODE, " & _
			" d.STOCKCODE, " & _
			" 0 AS openingbalanceqty, " & _
			" 0 AS TERIMA_HI, 0 AS TERIMASD_HI, " & _
			" d.QUANTITY AS KELUAR_HI, 0 AS KELUARSD_HI " & _
			" FROM ISSUE h, " & _
			" 			issuedetails d " & _
			" WHERE h.sivcode=d.sivcode " & _
			" 	AND h.fcba='"&pfcba&"' " & _
			" 	AND h.sivdate BETWEEN to_date('"&pdatefrom&"', 'mm/dd/yyyy') AND to_date('"&pdateto&"', 'mm/dd/yyyy') " & _
			addfilterSelect2 & _
			" UNION ALL SELECT d.fcba, " & _
			" '' AS STORECODE, " & _
			" d.STOCKCODE, " & _
			" 0 AS openingbalanceqty, " & _
			" 0 AS TERIMA_HI, 0 AS TERIMASD_HI, " & _
			" d.QUANTITYTRANSFERED AS KELUAR_HI, d.QUANTITYTRANSFERED AS KELUARSD_HI " & _
			" FROM transporttaskdtl_tab d, " & _
			" 		transporttask_tab h " & _
			" WHERE h.orderno=d.orderno " & _
			" 	AND h.rowstateid NOT IN('Rejected') " & _
			" 	AND h.fcba='"&pfcba&"' " & _
			" 	AND h.orderdate BETWEEN to_date('"&pdatefrom&"', 'mm/dd/yyyy') AND to_date('"&pdateto&"', 'mm/dd/yyyy') " & _ 
			" UNION ALL SELECT a.fcba, " & _
			" '' AS STORECODE, " & _
			" b.stockcode, " & _
			" 0 AS openingbalanceqty, " & _
			" 0 AS TERIMA_HI, " & _
			" b.QUANTITYRECEIVED AS TERIMASD_HI, " & _
			" 0 AS KELUAR_HI, 0 AS KELUARSD_HI " & _
			" FROM receivedtask_tab a " & _
			" JOIN receivedtaskdtl_tab b ON a.orderno = b.orderno " & _
			" WHERE a.orderdate BETWEEN to_date('"&pdatefrom&"', 'mm/dd/yyyy') AND to_date('"&pdateto&"', 'mm/dd/yyyy') " & _
			" AND a.fcba = '"&pfcba&"' and b.LOCATIONCODE != 'ST1' " & _
			addTable & _
			" )t " & _
			" LEFT JOIN stock s ON s.fccode = t.stockcode " & _
			" GROUP BY t.fcba, t.storecode, t.stockcode, s.fcname, s.UNITOFMEASUREMENT, s.stockgroupcode || ' - ' || s.fcname " & _
			" ORDER BY t.storecode, t.stockcode" & _
			" FETCH FIRST 100 ROWS ONLY"
    ' end if

	
	' response.write 		query
	Set List  = Server.CreateObject("ADODB.Recordset")
	Set caljson = new JSON  'JSON.aspê°€ í•„ìš”í•¨
	Set rs = connStr.Execute(query)

	Dim List_jsonString38
	List_jsonString38 = caljson.toJSON("results",rs, false)

	Dim jsonObj38, outputObj38
	set jsonObj38 = new JSONobject
	set outputObj38 = jsonObj38.parse(List_jsonString38)  'json.objectClass.aspê°€ í•„ìš”í•¨(íŒŒì‹±í•˜ë ¤ë©´)


	Set oJSON = New aspJSON     'aspJSON.asp í•„ìš”í•¨(forë¬¸)
	oJSON.loadJSON(List_jsonString38)

	Response.write List_jsonString38
end if

' get hpp kebun
if idView="39" Then
    company      	= Request("company")
	fcba 			= Request("fcba")
    hppfrom 	    = Request("bkmfrom")
    hppto      		= Request("bkmto")

	if trim(fcba) = "-" then
		filter = " AND NVL(s.fcba, b.fcba) IN ( " & _
        " SELECT fccode " & _
        " FROM businessunit " & _
        " WHERE fccompanycode = '" & company & "'  " & _
        " AND isactive = 'TRUE'  AND fctype='E' " & _
        " ) "
	else
	    filter = " AND s.fcba ='" & fcba & "' "
	end if

	qq = " WITH shipment AS ( " & _
			" SELECT " & _
			"   b.fccompanycode, " & _
        	"	ssh.fcba," & _
			"	ssd.postingdate tglkirim, " & _
			"	SUM(ssd.quantity) quantitysite, " & _
			"	SUM(ssd.acceptquantity) quantitypks " & _
			"		FROM salesshipmentdtl ssd " & _
			"		INNER JOIN salesshipmenthdr ssh " & _
			"			ON (ssd.salesshipmentno = ssh.salesshipmentno AND ssd.fcba = ssh.fcba) " & _
			"		INNER JOIN transporter tr  " & _
			"			ON (ssd.transporter = tr.fccode AND ssd.fcba = tr.fcba) " & _
			"		 INNER JOIN BUSINESSUNIT b ON (b.FCCODE=ssh.fcba) " & _
			"		WHERE  ssd.POSTINGDATE BETWEEN TO_DATE('" & hppfrom & "','mm/dd/yyyy')  " & _
			"								AND TO_DATE('" & hppto & "','mm/dd/yyyy') " & _
			"		GROUP BY b.fccompanycode, ssh.fcba,ssd.postingdate " & _
			"	), " & _
			" lhp AS ( " & _
			"	SELECT b.fccompanycode, " & _
			"		la.unit AS fcba, " & _
			"		lh.docketdate AS tglkirim, " & _
			"		SUM(ld.ffbweight) AS quantitysite, " & _
			"		SUM(ld.ffbweight) AS quantitypks " & _
			"	FROM lhptikethdr lh " & _
			"	INNER JOIN lhpffbwbtrn ld " & _
			"		ON lh.docketid = ld.docketid " & _
			"	AND lh.fcba = ld.fcba " & _
			"	INNER JOIN lhpAreamst la " & _
			"		ON la.areaid = ld.areaid " & _
			"	AND la.fcba = ld.fcba " & _
			"	INNER JOIN businessunit b " & _
			"		ON b.fccode = la.unit " & _
			"	WHERE lh.docketdate BETWEEN TO_DATE('" & hppfrom & "','mm/dd/yyyy') " & _
			"							AND TO_DATE('" & hppto & "','mm/dd/yyyy') " & _
			"	AND lh.productid = 2 " & _
			"	GROUP BY b.fccompanycode, la.unit, lh.docketdate " & _
			" ), " & _
			" shipment_final AS ( " & _
			"	SELECT fccompanycode, fcba, tglkirim, quantitysite, quantitypks, 1 AS priority " & _
			" FROM shipment " & _
			"	UNION ALL " & _
			"	SELECT fccompanycode, fcba, tglkirim, quantitysite, quantitypks, 2 AS priority " & _
			" FROM lhp " & _
			" ),"  & _
			" shipment_dedup AS ( "  & _
			"  SELECT fccompanycode, fcba, tglkirim, quantitysite, quantitypks "  & _
			" FROM ( "  & _
			"	SELECT sf.*,  "  & _
			"		ROW_NUMBER() OVER (PARTITION BY fccompanycode, fcba, tglkirim ORDER BY priority) AS rn "  & _
			"	FROM shipment_final sf "  & _
			") x "  & _
			" WHERE rn = 1 "  & _   
			" ), " & _
			"	biaya AS ( " & _
			"		SELECT  " & _
			"			b.fccompanycode, f.fcba,f.journaldate tglkirim, " & _
			"			SUM(CASE WHEN j.JOB_FORBUDGETCODE = '101010' THEN f.debet ELSE 0 END) perawatantm, " & _
			"			SUM(CASE WHEN j.JOB_FORBUDGETCODE = '101020' THEN f.debet ELSE 0 END) pemupukantm, " & _
			"			SUM(CASE WHEN j.JOB_FORBUDGETCODE = '101030' THEN f.debet ELSE 0 END) panendantransport, " & _
			"			SUM(CASE WHEN j.JOB_FORBUDGETCODE = '101040' THEN f.debet ELSE 0 END) sarpastm, " & _
			"			SUM(CASE WHEN j.JOB_FORBUDGETCODE = '101050' THEN f.debet+f.credit ELSE 0 END) umum " & _
			"		FROM GLEDGERDTL f " & _
			"		INNER JOIN jobmaster j " & _
			"			ON f.coa = j.fccode " & _
			"		INNER JOIN BUSINESSUNIT b ON (b.FCCODE=f.fcba) " & _
			"		WHERE journaldate BETWEEN TO_DATE('" & hppfrom & "','mm/dd/yyyy')  " & _
			"								AND TO_DATE('" & hppto & "','mm/dd/yyyy') " & _
			"		AND f.SOURCECODE IN('OPR-BKM','MAM-SIS','FAC-VRA','OPR-WIP','FAC-BNK','FAC-ISP','MAM-SRN','FAC-JRV','FAC-WSI') " & _
			"		GROUP BY b.fccompanycode, f.fcba,f.journaldate " & _
			"	) " & _
			"	SELECT  " & _
			"       b.fccompanycode," & _
    		"		b.fcba," & _
			"		NVL(s.tglkirim, b.tglkirim) tglkirim,"  & _
			"		NVL(s.quantitysite,0) quantitysite, " & _
			"		NVL(s.quantitypks,0) quantitypks, " & _
			"		NVL(s.quantitysite,0) - NVL(s.quantitypks,0) AS varian," & _
			"		NVL(b.perawatantm,0) perawatantm, " & _
			"		NVL(b.pemupukantm,0) pemupukantm, " & _
			"		NVL(b.panendantransport,0) panendantransport, " & _
			"		NVL(b.sarpastm,0) sarpastm, " & _
			"		NVL(b.umum,0) umum, " & _
			"		(NVL(b.perawatantm,0) " & _
			"		+ NVL(b.pemupukantm,0) " & _
			"		+ NVL(b.panendantransport,0) " & _
			"		+ NVL(b.sarpastm,0) " & _
			"		+ NVL(b.umum,0)) total " & _
			"	FROM shipment_dedup s " & _
			"	FULL OUTER JOIN biaya b " & _
			"	ON s.tglkirim = b.tglkirim AND s.fcba = b.fcba " & _
			"   WHERE NVL(s.fccompanycode, b.fccompanycode) = '" & company & "' " & _
			filter & _
			"	ORDER BY NVL(s.fcba, b.fcba),NVL(s.tglkirim, b.tglkirim) ASC NULLS FIRST"
		''	response.write qq


	Set List  = Server.CreateObject("ADODB.Recordset")
	Set caljson = new JSON  'JSON.aspê°€ í•„ìš”í•¨
	Set rs = connStr.Execute(qq)

	Dim List_jsonString39
	List_jsonString39 = caljson.toJSON("results",rs, false)

	Dim jsonObj39, outputObj39
	set jsonObj39 = new JSONobject
	set outputObj39 = jsonObj39.parse(List_jsonString39)  'json.objectClass.aspê°€ í•„ìš”í•¨(íŒŒì‹±í•˜ë ¤ë©´)


	Set oJSON = New aspJSON     'aspJSON.asp í•„ìš”í•¨(forë¬¸)
	oJSON.loadJSON(List_jsonString39)

	Response.write List_jsonString39
end if

'QUERY GET GL DETAIL UNTUK EXPORT EXCEL
if idview="40" then
   	dim tahungl2, bulangl2, coagl2, coaakhirgl2, bulantogl2, blntgl2
    ' fcba 		= session("fcba")
    tahungl2 	= Request("tahun")
    bulangl2 	= Request("bulan")
	bulantogl2 	= Request("bulanto")
	
    coagl2   	= Request("fcoa_ap")
	coaakhirgl2  = Request("coaakhir")
	companygl2	= Request("company")
	fcbagl2    	= Request("fcba")
	sourcecode2  = Request("sourcecode")
	
    Set List = Server.CreateObject("ADODB.Recordset")

	' sqlview = "select " & _
	' 		" a.DOCDATE, " & _
	' 		" b.source, " & _
	' 		" b.BKMNO, " & _
	' 		" a.DOCNO, " & _
	' 		" b.DESCRIPTION, " & _
	' 		" C.fccode, " & _
	' 		" b.locationcode, " & _
	' 		" case " & _
	' 		"		when SUBSTR(b.locationcode, 1, 3) = 'SPD' then (select fcname from supplier where fccode = b.locationcode OFFSET 1 ROWS FETCH NEXT 1 ROWS ONLY) " & _
	' 		"		else (select fcname from gcmaster where fccode = b.locationcode OFFSET 1 ROWS FETCH NEXT 1 ROWS ONLY) " & _
	' 		" end locationname," & _
	' 		" MAX(c.fccode || '-' || c.fcname) AS jobcode, " & _
	' 		" SUM(b.amountdebit) AS amountdebit, " & _
	' 		" SUM(b.amountkredit) AS amountkredit, " & _
	' 		" MAX(jobmaster_api.get_fcname(b.CHARGEJOB) || '-' || b.CHARGECODE) AS CHARGENAME, " & _
	' 		" b.CHARGEJOB " & _
	' 		" from BKMJOURNALHDR a " & _
	' 		" left join BKMJOURNALDTL b on a.DOCNO = b.DOCNO " & _
	' 		" INNER JOIN JOBMASTER C ON c.fccode = b.JOBCODE " & _
	' 		" where b.bkmno in  " & _
	' 		" 	( select b.JOURNALNO from GLEDGERHDR a " & _
	' 		" 	left join GLEDGERDTL b on a.DOCNO = b.DOCNO  " & _
	' 		" 	where a.period_month between '"&bulangl2&"' and '"&bulantogl2&"' and a.period_year = "&tahungl2&" and a.fcba = '"&fcbagl2&"' and b.COA between '"&coagl2&"' and '"&coaakhirgl2&"' " & _
	' 		" 	group by b.JOURNALNO ) " & _
	' 		" GROUP BY a.DOCDATE, b.source, b.BKMNO, a.DOCNO, b.DESCRIPTION, C.fccode, b.locationcode, b.CHARGEJOB"

	sqlview = " select "& _
				" h.DOCDATE,"& _
				" d.SOURCE, "& _
				" d.BKMNO,"& _
				" ("& _
				" select gl.DOCNO from GLEDGERHDR gl"& _
				" left join GLEDGERDTL gd on gl.DOCNO = gd.DOCNO"& _
				" where gd.JOURNALNO in (d.bkmno) group by gl.DOCNO"& _
				" FETCH FIRST 1 ROWS ONLY"& _
				" ) as DOCNO,"& _
				" d.description, "& _
				" j.fccode,"& _
				" d.locationcode, "& _
				" case"& _
				" when SUBSTR(d.locationcode, 1, 3) = 'SPD' then (select fcname from supplier where fccode = d.locationcode OFFSET 1 ROWS FETCH NEXT 1 ROWS ONLY)"& _
				" else (select fcname from gcmaster where fccode = d.locationcode OFFSET 1 ROWS FETCH NEXT 1 ROWS ONLY)"& _
				" end locationname,"& _
				" j.fccode || '-' ||  j.fcname as jobcode,"& _
				" sum(d.amountdebit) amountdebit,"& _
				" sum(d.amountkredit)amountkredit,"& _
				" d.CHARGEJOB, "& _
				" jobmaster_api.get_fcname(d.CHARGEJOB)||'-'||d.CHARGECODE  CHARGENAME"& _
				" from BKMJOURNALHDR h"& _
				" inner join BKMJOURNALDTL d on  (h.DOCNO=d.DOCNO and h.fcba=d.fcba)"& _
				" inner join jobmaster j on (j.fccode=d.JOBCODE)"& _
				" where d.BKMNO in ("& _
				" SELECT"& _
				" JOURNALNO"& _
				" FROM"& _
				" (SELECT '-' DOCNO,"& _
				" '-' JOURNALNO,"& _
				" TRUNC(TO_DATE('8'||'/01/'||'2025', 'MM/DD/YYYY'), 'MM')-1 JOURNALDATE,"& _
				" H.COA,"& _
				" 'SALDO AWAL BLN '||7 COANAME,"& _
				" 0 OPENING_BALANCE,"& _
				" Decode(Sign(H.OPENING_BALANCE), 1, H.OPENING_BALANCE, 0) DEBET,"& _
				" Decode(Sign(H.OPENING_BALANCE), -1, H.OPENING_BALANCE, 0) CREDIT,"& _
				" H.OPENING_BALANCE ENDING_BALANCE,"& _
				" H.FCBA,"& _
				" '-' SOURCECODE"& _
				" FROM GLEDGERHDR H"& _
				" INNER JOIN JOBMASTER C ON (H.COA=C.FCCODE)"& _
				" WHERE H.PERIOD_YEAR= "&tahungl2&""& _
				" AND H.PERIOD_MONTH = '"&bulangl2&"'"& _
				" AND H.FCBA = '"&fcbagl2&"' "& _
				" AND C.FCCODE BETWEEN'"&coagl2&"' AND '"&coaakhirgl2&"'"& _
				" UNION ALL SELECT G.DOCNO, G.JOURNALNO, G.JOURNALDATE, G.COA, C.FCNAME COANAME, H.OPENING_BALANCE, G.DEBET, G.CREDIT, G.ENDING_BALANCE, G.FCBA, J.SOURCECODE " & _
				" FROM GLEDGERDTL G " & _
				" INNER JOIN GLEDGERHDR H ON (G.DOCNO=H.DOCNO " & _
				" AND G.FCBA=H.FCBA) " & _
				" INNER JOIN JOBMASTER C ON (G.COA=C.FCCODE) " & _
				" INNER JOIN BKMJOURNALHDR J ON (G.JOURNALNO=J.DOCNO " & _
				" AND G.FCBA=J.FCBA) " & _
				" WHERE H.PERIOD_YEAR= "&tahungl2&" " & _
				" AND H.PERIOD_MONTH BETWEEN '"&bulangl2&"' AND '"&bulantogl2&"' " & _
				" AND H.FCBA = '"&fcbagl2&"' " & _
				" AND C.FCCODE BETWEEN '"&coagl2&"' AND '"&coaakhirgl2&"') " & _
				" WHERE JOURNALNO != '-' " & _
				" ) " & _
				" group by  " & _
				" h.DOCDATE, " & _
				" d.SOURCE, " & _
				" h.fcba, " & _
				" d.BKMNO, " & _
				" d.description, " & _
				" j.fccode, " & _
				" d.locationcode, " & _
				" j.fccode || '-' ||  j.fcname, " & _
				" lineno, " & _
				" d.CHARGEJOB, " & _
				" d.CHARGECODE  " & _
				" order by bkmno "
	' response.write sqlview
    List.Open sqlview, connStr, 1
    ListCount = List.RecordCount 

    Set caljson = new JSON  'JSON.aspê°€ í•„ìš”í•¨
    Set rs = connStr.Execute(sqlview)

    Dim List_jsonStringgl40
    List_jsonStringgl40 = caljson.toJSON("results",rs, false)

    Dim jsonObjgl40, outputObjgl40
    set jsonObjgl40 = new JSONobject
    set outputObjgl40 = jsonObjgl40.parse(List_jsonStringgl40)  'json.objectClass.aspê°€ í•„ìš”í•¨(íŒŒì‹±í•˜ë ¤ë©´)


    Set oJSON = New aspJSON     'aspJSON.asp í•„ìš”í•¨(forë¬¸)
    oJSON.loadJSON(List_jsonStringgl40)

    Response.write List_jsonStringgl40
end if

if idview="41" then
	Dim pfcCode, pfcba
	pfcCode = Request("fccode")
	pfcba   = Request("fcba")

	Function AddLeadingZero(Value)
		If Value < 10 Then
			AddLeadingZero = "0" & Value
		Else
			AddLeadingZero = Value
		End If
	End Function

	' Dim currentDate, dayPart, monthPart, yearPart, formattedDate
	currentDate = Now()
	dayPart = AddLeadingZero(Day(currentDate))
	monthPart = AddLeadingZero(Month(currentDate))
	yearPart = Year(currentDate)
	' Gunakan format mm/dd/yyyy agar cocok dengan to_date di Oracle
	formattedDate = monthPart & "/" & dayPart & "/" & yearPart

	' Buat query update
	Dim updt
	updt = "UPDATE receivenote SET " & _
		"RNDATE = to_date('" & formattedDate & "','mm/dd/yyyy'), " & _
		"ROWSTATEID = 'Canceled', " & _
		"LASTUPDATE = sysdate, " & _
		"FCEDIT = '" & Session("CodeName") & "' " & _
		"WHERE receivenotecode = '" & pfcCode & "' " & _
		"AND fcba = '" & pfcba & "'"

	' Jalankan query update
	connStr.Execute updt

	' Get data GL dan Delete
	delGL1 = "delete gledgerhdr where docno in (select docno from gledgerdtl where journalno = '" & pfcCode & "')"
	delGL2 = "delete gledgerdtl where journalno = '" & pfcCode & "'"

	connStr.Execute delGL1
	connStr.Execute delGL2

	' Delete bkmjournalhdr dan bkmjournaldtl
	deljurnal1 = "delete from bkmjournalhdr where docno = '" & pfcCode & "'"
	deljurnal2 = "delete from bkmjournaldtl where docno = '" & pfcCode & "'"

	connStr.Execute deljurnal1
	connStr.Execute deljurnal2

	' Get data GL dan Delete
	updtrcv = "update purchaseorderdetails set received = 0 where pocode in (select pocode from receivenote where receivenotecode = '" & pfcCode & "')"

	connStr.Execute updtrcv

	' updtstkledger = "update stockledger set qtyreceived = 0 where itemcode = '' and storecode = '' and fcba = '"&pfcba&"' and year = to_char(sysdate, 'YYYY')"

	' connStr.Execute updtrcv

	' Balikin QTY yang di stock ledger sesuai dengan item dan qty yang didelete

	' Buat respon JSON sederhana
	Response.ContentType = "application/json"
	Response.Write "{""status"":""OK"",""message"":""Data berhasil dibatalkan""}"
end if

if idview="42" then
	Dim ppocode
	ppocode = Request("pocode")

	' Query count data
	getdata = "SELECT COUNT(*) as jml FROM receivenote WHERE rowstateid = 'Planned' and pocode = '" & ppocode & "'"

	Set rs = connStr.Execute(getdata)

	Dim countResult
	If Not rs.EOF Then
		countResult = rs(0)
	Else
		countResult = 0
	End If

	rs.Close
	Set rs = Nothing

	' --- Tambahan validasi tipe ---
	If IsNull(countResult) Or Trim(countResult) = "" Then
		countResult = 0
	Else
		countResult = CLng(countResult)
	End If
	' -------------------------------

	If countResult > 0 Then
		Response.ContentType = "application/json"
		Response.Write "{""status"":""ERROR"",""message"":""No. PO " & ppocode & " sudah pernah dilakukan GRN sebelumnya dan belum dilakukan approval!"",""count"":" & countResult & "}"
		Response.End
	Else
		Response.ContentType = "application/json"
		Response.Write "{""status"":""OK"",""message"":""No. PO " & ppocode & " belum pernah dilakukan GRN."",""count"":" & countResult & "}"
	End If
end if
if idView="43" Then
	Dim pcompany,companyitem,datefrom,dateto,q,qq,no,rsdata
    pcompany  		= Request("companyitem")
    companyitem 	= Request("fcba")
    datefrom  		= Request("datefrom")
    dateto    		= Request("dateto")
    groupbarang    	= Request("groupbarang")
	querycode		= Request("querycode")

    q = "SELECT fccode, fccompanycode, FCNAME, FCCOMPANYNAME, FCMANAGER, FCKTU, FCCROP, FCTYPE FROM BusinessUnit WHERE fcCode='" & fcba & "' "
    set rsCat = connStr.Execute(q)
    if not rsCat.eof then
        CompanyName = rsCat("FCCOMPANYNAME")
        fcbaname    = rsCat("FCNAME")
    end if

	if trim(companyitem) = "-" then 
		filter = " AND G.FCBA in (SELECT fccode from businessunit where fccompanycode='" & pcompany & "' and isactive='TRUE') "
	else
	    filter = " AND G.FCBA ='" & companyitem & "' "
	end if

	if trim(groupbarang) <> "" then 
	    groupfilter = " AND GR.FCCODE ='" & groupbarang & "' "
	end if

	if trim(querycode) = 1 then 
		selectField = "TRIM(B.FCCOMPANYCODE) FCCOMPANYCODE, TRIM(H.FCBA) FCBA, TRIM(GR.FCNAME) AS GROUPNAME, TRIM(TO_CHAR(G.DOCDATE, 'DD/MM/YYYY')) AS TGLTRX, TRIM(H.ITEMCODE) ITEMCODE, TRIM(H.ITEMNAME) ITEMNAME, TRIM(S.UNITOFMEASUREMENT) SATUAN, TRIM(H.LOCATIONCODE) LOCATIONCODE, round(Sum(NVL(H.amountdebit, 0)), 2) amount, sum(H.UNIT) UNIT"
		GroupField = "GROUP BY TRIM(B.FCCOMPANYCODE), TRIM(H.FCBA), TRIM(GR.FCNAME), TRIM(TO_CHAR(G.DOCDATE, 'DD/MM/YYYY')), TRIM(H.ITEMCODE), TRIM(H.ITEMNAME), TRIM(S.UNITOFMEASUREMENT), TRIM(H.LOCATIONCODE)"
		OrderByField = " ORDER BY TRIM(H.FCBA) ASC"
	else
		selectField = " TRIM(B.FCCOMPANYCODE) FCCOMPANYCODE, TRIM(H.FCBA) FCBA, TRIM(H.JOBCODE) JOBCODE, TRIM(J.FCNAME) AS JOBNAME, TRIM(GR.FCNAME) AS GROUPNAME, TRIM(TO_CHAR(G.DOCDATE, 'DD/MM/YYYY')) AS TGLTRX, TRIM(H.ITEMCODE) ITEMCODE, TRIM(H.ITEMNAME) ITEMNAME, TRIM(S.UNITOFMEASUREMENT) SATUAN, TRIM(H.LOCATIONCODE) LOCATIONCODE,round(Sum(NVL(H.amountdebit,0)),2) amount, sum(H.UNIT) unit"
		GroupField = " GROUP BY TRIM(B.FCCOMPANYCODE), TRIM(H.FCBA), TRIM(H.JOBCODE), TRIM(J.FCNAME), TRIM(GR.FCNAME), TRIM(TO_CHAR(G.DOCDATE, 'DD/MM/YYYY')), TRIM(H.ITEMCODE), TRIM(H.ITEMNAME), TRIM(S.UNITOFMEASUREMENT), TRIM(H.LOCATIONCODE)"
		OrderByField = " ORDER BY TRIM(H.FCBA) ASC"
	end if

	query = " Select " & selectField & _
			" From  BKMJOURNALDTL H  " & _
			" INNER JOIN BKMJOURNALHDR G ON (G.DOCNO=H.DOCNO AND G.FCBA=H.FCBA) " & _
			" LEFT JOIN STOCK S ON (S.FCCODE=H.ITEMCODE AND S.FCBA=H.FCBA) " & _
			" LEFT JOIN STOCKGROUP GR ON (S.STOCKGROUPCODE=GR.FCCODE AND S.FCBA=GR.FCBA) " & _
			" INNER JOIN JOBMASTER J ON J.FCCODE=H.JOBCODE " & _
			" INNER JOIN BUSINESSUNIT B ON B.FCCODE=H.FCBA " & _
			" WHERE (G.DOCDATE between to_date('" & datefrom & "','mm/dd/yyyy') and to_date('" & dateto & "','mm/dd/yyyy')) AND H.ITEMCODE IS NOT NULL " & _
			filter  & _
			groupfilter  & _
			" AND G.POSTINGFLG = 1  " & _
			" AND H.amountdebit > 0 " & _
			" AND H.UNIT > 0 " & _
			GroupField & _
			OrderByField
    ' end if

	
	' response.write 		query
	Set List  = Server.CreateObject("ADODB.Recordset")
	Set caljson = new JSON  'JSON.aspê°€ í•„ìš”í•¨
	Set rs = connStr.Execute(query)

	Dim List_jsonString43
	List_jsonString43 = caljson.toJSON("results",rs, false)

	Dim jsonObj43, outputObj43
	set jsonObj43 = new JSONobject
	set outputObj43 = jsonObj43.parse(List_jsonString43)  'json.objectClass.aspê°€ í•„ìš”í•¨(íŒŒì‹±í•˜ë ¤ë©´)


	Set oJSON = New aspJSON     'aspJSON.asp í•„ìš”í•¨(forë¬¸)
	oJSON.loadJSON(List_jsonString43)

	Response.write List_jsonString43
end if

if idView = "44" then
	
	sqlPupuk =  " SELECT " & _
                "     T.fccompanycode, " & _
                "     T.UNITOFMEASUREMENT, " & _
                "     T.STOCKCODE,  " & _
                "     T.FCNAME,  " & _
                "     (sum(nvl(t.openingbalanceqty,0))+sum(nvl(t.receivenoteqty,0))-sum(nvl(t.issueqty,0))-sum(nvl(t.quantitytransfered,0))+sum(nvl(t.quantityreceived,0))+sum(nvl(t.qtywriteoff,0))) as qty, " & _
                "     SUM(T.openingbalanceqty) openingbalanceqty, " & _
                "     SUM(T.receivenoteqty) receivenoteqty, " & _
                "     SUM(T.issueqty) issueqty, " & _
                "     SUM(T.quantitytransfered) quantitytransfered, " & _
                "     SUM(T.quantityreceived) quantityreceived, " & _
                "     SUM(T.qtywriteoff) qtywriteoff " & _
                " FROM " & _
                " (SELECT  " & _
                "     bu.fccompanycode, " & _
                "     CASE WHEN b.UNITOFMEASUREMENT = 'KG' THEN 'TON' ELSE b.UNITOFMEASUREMENT END AS UNITOFMEASUREMENT, " & _
                "     l.itemcode STOCKCODE, b.FCNAME, l.openingbalanceqty, " & _
                "     l.qtyreceived AS receivenoteqty, " & _
                "     0 AS issueqty, " & _
                "     0 AS quantitytransfered, " & _
                "     0 AS quantityreceived, " & _
                "     0 AS qtywriteoff " & _
                "     FROM stockledger l " & _
                "     INNER JOIN bankstock k ON (l.itemcode=k.fccode) " & _
                "     INNER JOIN businessunit bu ON bu.fccode = l.fcba " & _
                "     INNER JOIN stock b ON l.itemcode = b.FCCODE " & _
                "     WHERE k.isactive='Y' " & _
                "     AND l.YEAR='2026' " & _
                "     AND l.MONTH='1' " & _
                "     AND b.stockgroupcode IN ('103', '305', '102', '301') " & _
                "     AND bu.fccompanycode in ('PT.ARJ', 'PT.BRS', 'PT.GAS', 'PT.KSR', 'PT.BIA', 'PT.BRP', 'PT.KAA', 'PT.KSR', 'PT.ESR') " & _
                "     group by  " & _
                "     bu.fccompanycode,  " & _
                "     CASE WHEN b.UNITOFMEASUREMENT = 'KG' THEN 'TON' ELSE b.UNITOFMEASUREMENT END, " & _
                "     l.itemcode, b.FCNAME, l.openingbalanceqty, l.qtyreceived " & _
                "     UNION ALL " & _
                "     SELECT " & _
                "     bu.fccompanycode, " & _
                "     CASE WHEN b.UNITOFMEASUREMENT = 'KG' THEN 'TON' ELSE b.UNITOFMEASUREMENT END AS UNITOFMEASUREMENT, " & _
                "     rd.purchaseitemcode STOCKCODE, b.FCNAME, " & _
                "     0 AS openingbalanceqty, " & _
                "     rd.quantity AS receivenoteqty, " & _
                "     0 AS issueqty, " & _
                "     0 AS quantitytransfered, " & _
                "     0 AS quantityreceived, " & _
                "     0 AS qtywriteoff " & _
                "     FROM receivenote r " & _
                "     INNER JOIN receivenotedetail rd on r.RECEIVENOTECODE = rd.RECEIVENOTECODE " & _
                "     INNER JOIN businessunit bu ON bu.fccode = r.fcba " & _
                "     INNER JOIN stock b ON rd.purchaseitemcode = b.FCCODE " & _
                "     WHERE r.receivenotecode=rd.receivenotecode " & _
                "     AND r.rndate BETWEEN to_date('01/01/2026', 'mm/dd/yyyy') AND to_date('01/08/2026', 'mm/dd/yyyy') " & _
                "     AND b.stockgroupcode IN ('103', '305', '102', '301') " & _
                "     AND bu.fccompanycode in ('PT.ARJ', 'PT.BRS', 'PT.GAS', 'PT.KSR', 'PT.BIA', 'PT.BRP', 'PT.KAA', 'PT.KSR', 'PT.ESR') " & _
                "     group by  " & _
                "     bu.fccompanycode, " & _
                "     CASE WHEN b.UNITOFMEASUREMENT = 'KG' THEN 'TON' ELSE b.UNITOFMEASUREMENT END, " & _
                "     rd.purchaseitemcode, b.FCNAME, rd.quantity  " & _
                "     UNION ALL " & _
                "     SELECT bu.fccompanycode, " & _
                "     CASE WHEN b.UNITOFMEASUREMENT = 'KG' THEN 'TON' ELSE b.UNITOFMEASUREMENT END AS UNITOFMEASUREMENT, " & _
                "     d.STOCKCODE, b.FCNAME, " & _
                "     0 AS openingbalanceqty, " & _
                "     0 AS receivenoteqty, " & _
                "     d.quantity AS issueqty, " & _
                "     0 AS quantitytransfered, " & _
                "     0 AS quantityreceived, " & _
                "     0 AS qtywriteoff " & _
                "     FROM ISSUE h " & _
                "     INNER JOIN issuedetails d on h.SIVCODE = d.SIVCODE " & _
                "     INNER JOIN businessunit bu ON bu.fccode = h.fcba " & _
                "     INNER JOIN stock b ON d.STOCKCODE = b.FCCODE " & _
                "     WHERE h.sivcode=d.sivcode " & _
                "     AND h.sivdate BETWEEN to_date('01/01/2026', 'mm/dd/yyyy') AND to_date('01/08/2026', 'mm/dd/yyyy') " & _
                "     AND b.stockgroupcode IN ('103', '305', '102', '301') " & _
                "     AND bu.fccompanycode in ('PT.ARJ', 'PT.BRS', 'PT.GAS', 'PT.KSR', 'PT.BIA', 'PT.BRP', 'PT.KAA', 'PT.KSR', 'PT.ESR') " & _
                "     GROUP BY  " & _
                "     bu.fccompanycode, " & _
                "     CASE WHEN b.UNITOFMEASUREMENT = 'KG' THEN 'TON' ELSE b.UNITOFMEASUREMENT END, " & _
                "     d.STOCKCODE, b.FCNAME, d.quantity " & _
                "     UNION ALL  " & _
                "     SELECT bu.fccompanycode, " & _
                "     CASE WHEN b.UNITOFMEASUREMENT = 'KG' THEN 'TON' ELSE b.UNITOFMEASUREMENT END AS UNITOFMEASUREMENT, " & _
                "     d.STOCKCODE, b.FCNAME, " & _
                "     0 AS openingbalanceqty, " & _
                "     0 AS receivenoteqty, " & _
                "     0 AS issueqty, " & _
                "     d.quantitytransfered, " & _
                "     0 AS quantityreceived, " & _
                "     0 AS qtywriteoff " & _
                "     FROM transporttaskdtl_tab d " & _
                "     INNER JOIN transporttask_tab h on d.ORDERNO = h.ORDERNO " & _
                "     INNER JOIN businessunit bu ON bu.fccode = h.fcba " & _
                "     INNER JOIN stock b ON d.STOCKCODE = b.FCCODE " & _
                "     WHERE h.orderno=d.orderno " & _
                "     AND h.rowstateid NOT IN('Rejected') " & _
                "     AND h.orderdate BETWEEN to_date('01/01/2026', 'mm/dd/yyyy') AND to_date('01/08/2026', 'mm/dd/yyyy') " & _
                "     AND b.stockgroupcode IN ('103', '305', '102', '301') " & _
                "     AND bu.fccompanycode in ('PT.ARJ', 'PT.BRS', 'PT.GAS', 'PT.KSR', 'PT.BIA', 'PT.BRP', 'PT.KAA', 'PT.KSR', 'PT.ESR') " & _
                "     GROUP BY  " & _
                "     bu.fccompanycode, " & _
                "     CASE WHEN b.UNITOFMEASUREMENT = 'KG' THEN 'TON' ELSE b.UNITOFMEASUREMENT END, " & _
                "     d.STOCKCODE, b.FCNAME, d.quantitytransfered  " & _
                "     UNION ALL  " & _
                "     SELECT  " & _
                "     bu.fccompanycode, " & _
                "     d.STOCKCODE, " & _
                "     CASE WHEN b.UNITOFMEASUREMENT = 'KG' THEN 'TON' ELSE b.UNITOFMEASUREMENT END AS UNITOFMEASUREMENT, " & _
                "     b.FCNAME, " & _
                "     0 AS openingbalanceqty, " & _
                "     0 AS receivenoteqty, " & _
                "     0 AS issueqty, " & _
                "     0 AS quantitytransfered, " & _
                "     d.quantityreceived, " & _
                "     0 AS qtywriteoff " & _
                "     FROM receivedtask_tab h " & _
                "     INNER JOIN receivedtaskdtl_tab d on h.ORDERNO = d.ORDERNO " & _
                "     INNER JOIN businessunit bu ON bu.fccode = h.fcba " & _
                "     INNER JOIN stock b ON d.STOCKCODE = b.FCCODE " & _
                "     WHERE h.orderno=d.orderno " & _
                "     AND h.orderdate BETWEEN to_date('01/01/2026', 'mm/dd/yyyy') AND to_date('01/08/2026', 'mm/dd/yyyy')  " & _
                "     AND b.stockgroupcode IN ('103', '305', '102', '301') " & _
                "     AND bu.fccompanycode in ('PT.ARJ', 'PT.BRS', 'PT.GAS', 'PT.KSR', 'PT.BIA', 'PT.BRP', 'PT.KAA', 'PT.KSR', 'PT.ESR') " & _
                "     GROUP BY  " & _
                "     bu.fccompanycode, " & _
                "     CASE WHEN b.UNITOFMEASUREMENT = 'KG' THEN 'TON' ELSE b.UNITOFMEASUREMENT END, " & _
                "     d.STOCKCODE, b.FCNAME, d.quantityreceived  " & _
                "     UNION ALL " & _
                "     SELECT  " & _
                "     bu.fccompanycode, " & _
                "     d.STOCKCODE, " & _
                "     CASE WHEN b.UNITOFMEASUREMENT = 'KG' THEN 'TON' ELSE b.UNITOFMEASUREMENT END AS UNITOFMEASUREMENT, " & _
                "     b.FCNAME, " & _
                "     0 AS openingbalanceqty, " & _
                "     0 AS receivenoteqty, " & _
                "     0 AS issueqty, " & _
                "     0 AS quantitytransfered, " & _
                "     0 AS quantityreceived, " & _
                "     d.quantity AS qtywriteoff " & _
                "     FROM writeoff h " & _
                "     INNER JOIN writeoffdetails d on h.WRITEOFFCODE = d.WRITEOFFCODE  " & _
                "     INNER JOIN businessunit bu ON bu.fccode = h.fcba " & _
                "     INNER JOIN stock b ON d.STOCKCODE = b.FCCODE " & _
                "     WHERE h.writeoffcode=d.writeoffcode " & _
                "     AND h.wodate BETWEEN to_date('01/01/2026', 'mm/dd/yyyy') AND to_date('01/08/2026', 'mm/dd/yyyy')  " & _
                "     AND b.stockgroupcode IN ('103', '305', '102', '301') " & _
                "     AND bu.fccompanycode in ('PT.ARJ', 'PT.BRS', 'PT.GAS', 'PT.KSR', 'PT.BIA', 'PT.BRP', 'PT.KAA', 'PT.KSR', 'PT.ESR') " & _
                "     GROUP BY  " & _
                "     bu.fccompanycode, " & _
                "     CASE WHEN b.UNITOFMEASUREMENT = 'KG' THEN 'TON' ELSE b.UNITOFMEASUREMENT END, " & _
                "     d.STOCKCODE, b.FCNAME, d.quantity  " & _
                "     UNION ALL  " & _
                "     SELECT  " & _
                "     bu.fccompanycode, " & _
                "     d.itemcode STOCKCODE, " & _
                "     CASE WHEN b.UNITOFMEASUREMENT = 'KG' THEN 'TON' ELSE b.UNITOFMEASUREMENT END AS UNITOFMEASUREMENT, " & _
                "     b.FCNAME, " & _
                "     0 AS openingbalanceqty, " & _
                "     0 AS receivenoteqty, " & _
                "     0 AS issueqty, " & _
                "     0 AS quantitytransfered, " & _
                "     d.qtyreturn AS quantityreceived, " & _
                "     0 AS qtywriteoff " & _
                "     FROM returnissuehdr h " & _
                "     INNER JOIN returnissuedtl d on h.RETURNNO = d.RETURNNO " & _
                "     INNER JOIN businessunit bu ON bu.fccode = h.fcba " & _
                "     INNER JOIN stock b ON d.itemcode = b.FCCODE " & _
                "     WHERE h.returnno=d.returnno " & _
                "     AND h.returndate BETWEEN to_date('01/01/2026', 'mm/dd/yyyy') AND to_date('01/08/2026', 'mm/dd/yyyy')  " & _
                "     AND b.stockgroupcode IN ('103', '305', '102', '301') " & _
                "     AND bu.fccompanycode in ('PT.ARJ', 'PT.BRS', 'PT.GAS', 'PT.KSR', 'PT.BIA', 'PT.BRP', 'PT.KAA', 'PT.KSR', 'PT.ESR') " & _
                "     GROUP BY  " & _
                "     bu.fccompanycode, " & _
                "     CASE WHEN b.UNITOFMEASUREMENT = 'KG' THEN 'TON' ELSE b.UNITOFMEASUREMENT END, " & _
                "     d.itemcode, b.FCNAME, d.qtyreturn) T " & _
                " GROUP BY  " & _
                " T.fccompanycode, " & _
                " T.UNITOFMEASUREMENT, " & _
                " T.STOCKCODE,  " & _
                " T.FCNAME "

	Set List  = Server.CreateObject("ADODB.Recordset")
	Set caljson = new JSON  'JSON.aspê°€ í•„ìš”í•¨
	Set rs = connStr.Execute(sqlPupuk)

	Dim List_jsonString44
	List_jsonString44 = caljson.toJSON("results",rs, false)

	Dim jsonObj44, outputObj44
	set jsonObj44 = new JSONobject
	set outputObj44 = jsonObj44.parse(List_jsonString44)  'json.objectClass.aspê°€ í•„ìš”í•¨(íŒŒì‹±í•˜ë ¤ë©´)


	Set oJSON = New aspJSON     'aspJSON.asp í•„ìš”í•¨(forë¬¸)
	oJSON.loadJSON(List_jsonString44)

	Response.write List_jsonString44

end if

if idView="45" Then
	Dim tahun44,bulan44,fcba44, kolomid,kolomname
    tahun44 = Request("tahun")
    bulan44 = Request("bulan")
    fcba44  = session("fcba")
	kolomid = Request("kolomid")

	if kolomid = 1 then 
		kolomname = "aresta_closing"
	end if
	if kolomid = 2 then 
		kolomname = "kasbank_closing"
	end if
	if kolomid = 3 then 
		kolomname = "stock_closing"
	end if
	if kolomid = 4 then 
		kolomname = "wip_closing"
	end if
	if kolomid = 5 then 
		kolomname = "wsvra_closing"
	end if
	if kolomid = 6 then 
		kolomname = "payroll_closing"
	end if
	if kolomid = 7 then 
		kolomname = "unit_closing"
	end if
   
	query = " SELECT " & kolomname & " AS closingStatus " & _ 
	" FROM EMPALLDEDSTATUS " & _
	" WHERE YEAR='" & tahun44 & "' AND MONTH='" & bulan44 & "' AND FCBA='" & fcba44 & "'"

	
	'' response.write 		query
	Set List  = Server.CreateObject("ADODB.Recordset")
	Set caljson = new JSON  'JSON.aspê°€ í•„ìš”í•¨
	Set rs = connStr.Execute(query)

	Dim List_jsonString45
	List_jsonString45 = caljson.toJSON("results",rs, false)

	Dim jsonObj45, outputObj45
	set jsonObj45 = new JSONobject
	set outputObj45 = jsonObj45.parse(List_jsonString45)  


	Set oJSON = New aspJSON     
	oJSON.loadJSON(List_jsonString45)

	Response.write List_jsonString45
end if

' Get HPP BIBIT
if idView="46" Then
    company46      	= Request("company")
	fcba46 			= Request("fcba")
    hppfrom46	    = Request("bkmfrom")
    hppto46   		= Request("bkmto")

	if trim(fcba46) = "-" then
		wheresql = " NVL(a.fcba, a.fcba) IN ( " & _
        " SELECT fccode " & _
        " FROM businessunit " & _
        " WHERE fccompanycode = '" & company46 & "'  " & _
        " AND isactive = 'TRUE'  AND fctype='E' " & _
        " ) "
	else
	    wheresql = " a.fcba ='" & fcba46 & "' "
	end if

	Set List = Server.CreateObject("ADODB.Recordset")

	sqlview = "WITH jobmaster_group AS ( " & _
				"    SELECT " & _
				"        fccode, " & _
				"        jobgroupcode, " & _
				"        MAX(CASE WHEN jobgroupcode = '1310200' THEN 1 ELSE 0 END) AS is_group_1310200, " & _
				"        MAX(CASE WHEN jobgroupcode = '1310300' THEN 1 ELSE 0 END) AS is_group_1310300 " & _
				"    FROM jobmaster " & _
				"    WHERE jobgroupcode IN ('1310200', '1310300') " & _
				"    GROUP BY fccode, jobgroupcode " & _
				"), " & _
				"persiapan_codes AS ( " & _
				"    SELECT fccode " & _
				"    FROM jobmaster " & _
				"    WHERE fccode IN ( " & _
				"        '1310101','1310102','1310103','1310104','1310105','1310106', " & _
				"        '1310107','1310108','1310109','1310110','1310111','1310112','1310113' " & _
				"    ) " & _
				") " & _
				"SELECT " & _
				"    d.FCCOMPANYCODE AS company, " & _
				"    a.FCBA, " & _
				"    a.DOCDATE, " & _
				"    SUM(CASE WHEN b.JOBCODE = '1310203' THEN b.UNIT ELSE 0 END) AS kecambah, " & _
				"    SUM(CASE WHEN b.JOBCODE = '1310205' THEN b.UNIT ELSE 0 END) AS tanam_kecambah, " & _
				"    SUM(CASE WHEN b.JOBCODE = '1310214' THEN b.UNIT ELSE 0 END) AS pre_nursery, " & _
				"    SUM(CASE WHEN b.JOBCODE = '1310305' THEN b.UNIT ELSE 0 END) AS main_nursery, " & _
				"    SUM(CASE WHEN b.JOBCODE = '1310401' THEN b.UNIT ELSE 0 END) AS transfer, " & _
				"    SUM(CASE WHEN b.JOBCODE = '1310401' THEN b.AMOUNTDEBIT + b.AMOUNTKREDIT ELSE 0 END) AS alokasi, " & _
				"    SUM(CASE WHEN pc.fccode IS NOT NULL THEN b.AMOUNTDEBIT + b.AMOUNTKREDIT ELSE 0 END) AS persiapan, " & _
				"    SUM(CASE WHEN jg.is_group_1310200 = 1 THEN b.AMOUNTDEBIT + b.AMOUNTKREDIT ELSE 0 END) AS prenursery2, " & _
				"    SUM(CASE WHEN jg.is_group_1310300 = 1 THEN b.AMOUNTDEBIT + b.AMOUNTKREDIT ELSE 0 END) AS main_nursery2 " & _
				"FROM BKMJOURNALHDR a " & _
				"JOIN BUSINESSUNIT d " & _
				"    ON a.FCBA = d.FCCODE " & _
				"LEFT JOIN BKMJOURNALDTL b " & _
				"    ON a.DOCNO = b.DOCNO " & _
				"LEFT JOIN jobmaster_group jg " & _
				"    ON b.JOBCODE = jg.fccode " & _
				"LEFT JOIN persiapan_codes pc " & _
				"    ON b.JOBCODE = pc.fccode " & _
				"WHERE " & _
				wheresql & _
				" AND a.DOCDATE BETWEEN to_date('" & hppfrom46 & "', 'mm/dd/yyyy') AND to_date('" & hppto46 & "', 'mm/dd/yyyy') " & _
                " AND a.ACCOUNTINGPERIOD <= 12 " & _
				"GROUP BY " & _
				"    d.FCCOMPANYCODE, a.FCBA, a.DOCDATE " & _
				"HAVING " & _
				"    SUM(CASE WHEN b.JOBCODE = '1310203' THEN b.UNIT ELSE 0 END) <> 0 " & _
				"    OR SUM(CASE WHEN b.JOBCODE = '1310205' THEN b.UNIT ELSE 0 END) <> 0 " & _
				"    OR SUM(CASE WHEN b.JOBCODE = '1310214' THEN b.UNIT ELSE 0 END) <> 0 " & _
				"    OR SUM(CASE WHEN b.JOBCODE = '1310305' THEN b.UNIT ELSE 0 END) <> 0 " & _
				"    OR SUM(CASE WHEN b.JOBCODE = '1310401' THEN b.UNIT ELSE 0 END) <> 0 " & _
				"    OR SUM(CASE WHEN b.JOBCODE = '1310401' THEN b.AMOUNTDEBIT + b.AMOUNTKREDIT ELSE 0 END) <> 0 " & _
				"    OR SUM(CASE WHEN pc.fccode IS NOT NULL THEN b.AMOUNTDEBIT + b.AMOUNTKREDIT ELSE 0 END) <> 0 " & _
				"    OR SUM(CASE WHEN jg.is_group_1310200 = 1 THEN b.AMOUNTDEBIT + b.AMOUNTKREDIT ELSE 0 END) <> 0 " & _
				"    OR SUM(CASE WHEN jg.is_group_1310300 = 1 THEN b.AMOUNTDEBIT + b.AMOUNTKREDIT ELSE 0 END) <> 0 " & _
				"ORDER BY a.DOCDATE ASC"

    ' response.write sqlview

    List.Open sqlview, connStr, 1
    ListCount = List.RecordCount 

    Set caljson = new JSON  'JSON.aspê°€ í•„ìš”í•¨
    Set rs = connStr.Execute(sqlview)

    Dim List_jsonString46
    List_jsonString46 = caljson.toJSON("results",rs, false)

    Dim jsonObj46, outputObj46
    set jsonObj46 = new JSONobject
    set outputObj46 = jsonObj46.parse(List_jsonString46)  'json.objectClass.aspê°€ í•„ìš”í•¨(íŒŒì‹±í•˜ë ¤ë©´)


    Set oJSON = New aspJSON     'aspJSON.asp í•„ìš”í•¨(forë¬¸)
    oJSON.loadJSON(List_jsonString46)

    Response.write List_jsonString46
end if

' Get HPP BIBIT REPORT ALL
if idView="47" Then
	company47   = Request("company")
	fcba47      = Request("fcba")
	tahun47   	= Request("tahun")
	bulan47     = Request("bulan")
	from47		= Request("from")
	to47		= Request("to")
	batch47		= Request("batch")

	if trim(bulan47) = "0" then
		whereMonth = ""
	else
		whereMonth = " AND a.ACCOUNTINGPERIOD = "& bulan47 &" AND a.ACCOUNTINGYEAR = "& tahun47 &" "
	end if

	if trim(from47) = "0" then
		whereDate = ""
	else
		whereDate = " AND a.DOCDATE BETWEEN to_date('" & from47 & "', 'mm/dd/yyyy') AND to_date('" & to47 & "', 'mm/dd/yyyy') AND a.ACCOUNTINGPERIOD <= 12 "
	end if

	If trim(batch47) = "-" Then
		wherebatch = " "
	Else
		wherebatch = " AND b.LOCATIONCODE ='" & batch47 & "' "
	End If

	' FIX 1: fcba46 -> fcba47
	If trim(fcba47) = "-" Then
		wheresql47 = " a.fcba IN ( " & _
					" SELECT fccode " & _
					" FROM businessunit " & _
					" WHERE fccompanycode = '" & company47 & "' " & _
					" AND isactive = 'TRUE' AND fctype='E' " & _
					" ) "
	Else
		wheresql47 = " a.fcba ='" & fcba47 & "' "
	End If

	sqlview47 = "SELECT " & _
				" b.FCBA, " & _
				" MIN(b.LOCATIONTYPE) AS LOCATIONTYPE, " & _
				" TRIM(b.LOCATIONCODE) AS LOCATIONCODE, MAX(n.DATEPLANTED) AS DATEPLANTED, " & _
				" COALESCE(MIN(n.NURSERYCODE), TRIM(b.LOCATIONCODE)) AS NURSERYCODE, " & _
				" MIN(c.FCNAME) AS FCNAME, " & _
				" SUM(CASE WHEN b.JOBCODE = '1310203' THEN b.UNIT ELSE 0 END) AS KECAMBAH, " & _
				" SUM(CASE WHEN b.JOBCODE = '1310205' THEN b.UNIT ELSE 0 END) AS TANAM_KECAMBAH, " & _
				" SUM(CASE WHEN b.JOBCODE = '1310214' THEN b.UNIT ELSE 0 END) AS PRE_NURSERY, " & _
				" SUM(CASE WHEN b.JOBCODE = '1310305' THEN b.UNIT ELSE 0 END) AS MAIN_NURSERY, " & _
				" SUM(CASE WHEN b.JOBCODE = '1310401' THEN b.UNIT ELSE 0 END) AS TRANSFER, " & _
				" SUM(CASE WHEN b.JOBCODE = '1310401' THEN b.AMOUNTDEBIT + b.AMOUNTKREDIT ELSE 0 END) AS ALOKASI, " & _
				" 0 AS SALDOAWAL, " & _
				" SUM(b.AMOUNTDEBIT) AS DEBET, " & _
				" SUM(b.AMOUNTKREDIT) AS KREDIT, " & _
				" 0 AS SALDOAKHIR, " & _
				" 0 AS STOK, " & _
				" 0 AS HPP, " & _
				" MIN(n.DESCRIPTION) AS DESCRIPTION " & _
				" FROM BKMJOURNALHDR a " & _
				" LEFT JOIN BKMJOURNALDTL b ON a.DOCNO = b.DOCNO " & _
				" LEFT JOIN JOBMASTER c ON b.JOBCODE = c.FCCODE " & _
				" LEFT JOIN ( " & _
				"		SELECT  " & _
				"			TRIM(NURSERYCODE) AS NURSERYCODE, " & _
				"			MAX(DATEPLANTED) AS DATEPLANTED, " & _
				"			MIN(DESCRIPTION) AS DESCRIPTION " & _
				"		FROM NURSERY WHERE DATEPLANTED IS NOT NULL " & _
				"		GROUP BY TRIM(NURSERYCODE) " & _
				"	) n ON b.LOCATIONCODE = n.NURSERYCODE " & _
				" WHERE " & wheresql47 & _
				wherebatch & _
				" AND c.fccode IN ('1310203', '1310205', '1310214', '1310305', '1310401')" & _
				whereMonth & _
				whereDate & _
				" GROUP BY b.FCBA, b.LOCATIONCODE ORDER BY NURSERYCODE ASC"
	' response.write sqlview47
	' FIX 2: Gunakan hanya satu cara buka recordset, pakai Execute saja
	' FIX 3: sqlview -> sqlview47
	Set rs = connStr.Execute(sqlview47)

	Set caljson47 = New JSON
	' FIX 4: List_jsonString46 -> List_jsonString47
	Dim List_jsonString47
	List_jsonString47 = caljson47.toJSON("results", rs, false)

	Dim jsonObj47, outputObj47
	Set jsonObj47 = New JSONobject
	Set outputObj47 = jsonObj47.parse(List_jsonString47)

	Set oJSONs = New aspJSON
	oJSONs.loadJSON(List_jsonString47)

	Response.Write List_jsonString47

	rs.Close
	Set rs = Nothing
end if

%>
