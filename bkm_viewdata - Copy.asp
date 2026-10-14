
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


dim txtcari,fcba,bkmfrom,bkmto,idview
idview   = Request("idview")
'tampilkan data listing bkm'
if idview="1" then

    fcba = session("fcba")
    txtcari = Request("keyword")
    bkmfrom = Request("bkmfrom")
    bkmto   = Request("bkmto")
	pilihbkm   = Request("pilihbkm")

	sqlexec = "BEGIN " & _ 
              "bkmproses_upah_api.Update_hdr_fromto(to_date('" & bkmfrom & "','mm/dd/yyyy') , to_date('" & bkmto & "','mm/dd/yyyy') ,'" & fcba & "');" & _
			  "END; " 
    'connStr.Execute(sqlexec)
    
    Set List = Server.CreateObject("ADODB.Recordset")


    sqlview =  "Select h.BKMNO, h.BKMDATE CREATEDATE,h.SECTIONCODE,h.GANGCODE,nvl(h.VEHICLECODE,'-') VEHICLECODE, h.TOTUPAH TOTALBKMUPAH,h.TOTBAHAN TOTALBKMMATERIAL,h.TOTVRA TOTALBKMVEHICLE,h.TOTWIP TOTALBKMWIP,h.DOCUMENT_CLASSIFICATION " & _
	"    ,authorizestructure_api.get_template_desc(h.DOCUMENT_CLASSIFICATION) dcname,decode(h.hujan,'0','-','YA') hujan,decode(h.POSTINGFLG,'1','YA','TIDAK') POSTINGFLG" & _
	"    ,h.TOTUPAH + h.TOTBAHAN + h.TOTVRA + h.TOTWIP TOTAL " & _	
    "    From   BKMHDR h WHERE h.BKMDATE between to_date('" & bkmfrom & "','mm/dd/yyyy') and to_date('" & bkmto & "','mm/dd/yyyy') and h.fcba='" & fcba & "' " & _ 
    "    AND document_classification Like '" & pilihbkm & "%' " & _ 
	"    ORDER BY h.BKMDATE,h.SECTIONCODE,h.GANGCODE,h.VEHICLECODE,h.DOCUMENT_CLASSIFICATION "

  ''  response.write sqlview

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
end if

'tampilkan data gl'
if idview="2" then
    dim tahungl,bulangl,coagl
    fcba = session("fcba")
    tahungl = Request("tahun")
    bulangl = Request("bulan")
    coagl   = Request("coa")
	companygl = Request("company")
	fcbagl    = Request("fcba")


    Set List = Server.CreateObject("ADODB.Recordset")

	if trim(fcbagl) = "-" then 
		filter = " AND G.FCBA in (SELECT fccode from businessunit where fccompanycode='" & companygl & "') "
	else
	    filter = " AND G.FCBA ='" & fcbagl & "' "
	end if


    sqlview =  "Select G.DOCNO, G.JOURNALNO, G.JOURNALDATE, G.COA,G.COA || ' - ' ||JOB_API.get_jobcode_name(COA,G.FCBA) COANAME, G.OPENING_BALANCE, " & _
			" G.DEBET, G.CREDIT, G.ENDING_BALANCE, G.CREATEDATE, G.LASTUPDATE, " & _
			" G.FCENTRY, G.FCEDIT, G.FCIP, G.FCBA " & _
			" From   GLEDGERDTL G " & _
			" INNER JOIN JOB C ON (G.COA=C.FCCODE AND G.FCBA=C.FCBA) " & _
			" WHERE TO_NUMBER(TO_CHAR(G.JOURNALDATE,'YYYY'))='" & tahungl & "' AND TO_NUMBER(TO_CHAR(G.JOURNALDATE,'MM'))='" & bulangl & "' AND G.COA='" & coagl & "' " & _
			filter & _
			" ORDER BY G.fcba,G.JOURNALDATE,G.JOURNALNO " 

    'response.write sqlview

    List.Open sqlview, connStr, 1
    ListCount = List.RecordCount 

    Set caljson = new JSON  'JSON.asp가 필요함
    Set rs = connStr.Execute(sqlview)

    Dim List_jsonStringgl
    List_jsonStringgl = caljson.toJSON("results",rs, false)

    Dim jsonObjgl, outputObjgl
    set jsonObjgl = new JSONobject
    set outputObjgl = jsonObjgl.parse(List_jsonStringgl)  'json.objectClass.asp가 필요함(파싱하려면)


    Set oJSON = New aspJSON     'aspJSON.asp 필요함(for문)
    oJSON.loadJSON(List_jsonStringgl)

    Response.write List_jsonStringgl
end if

'tampilkan data glhdr'
if idview="3" then
    dim tahunglhdr,bulanglhdr,coaglhdr
    fcba = session("fcba")
    tahunglhdr = Request("tahun")
    bulanglhdr = Request("bulan")
    coaglhdr   = Request("coa")
	companyhdr = Request("company")
	fcbahdr    = Request("fcba")



    Set List = Server.CreateObject("ADODB.Recordset")

	if trim(fcbahdr) = "-" then 
		filter = " AND G.FCBA in (SELECT fccode from businessunit where fccompanycode='" & companyhdr & "') "
	else
	    filter = " AND G.FCBA ='" & fcbahdr & "' "
	end if


    sqlview =  "Select SUM(G.OPENING_BALANCE) OPENING_BALANCE, " & _
			   " SUM(G.DEBET) DEBET, SUM(G.CREDIT) CREDIT,SUM(G.OPENING_BALANCE)+ SUM(G.DEBET) + SUM(G.CREDIT)  ENDING_BALANCE " & _ 
			   " From  GLEDGERHDR G " & _
			   " WHERE PERIOD_MONTH='" & bulanglhdr & "' AND PERIOD_YEAR='" & tahunglhdr & "' AND G.COA='" & coaglhdr & "'" & _
			   filter & _ 
			   " GROUP BY PERIOD_MONTH,PERIOD_YEAR,COA "

    'response.write sqlview

    List.Open sqlview, connStr, 1
    ListCount = List.RecordCount 

    Set caljson = new JSON  'JSON.asp가 필요함
    Set rs = connStr.Execute(sqlview)

    Dim List_jsonStringglhdr
    List_jsonStringglhdr = caljson.toJSON("results",rs, false)

    Dim jsonObjglhdr, outputObjglhdr
    set jsonObjglhdr = new JSONobject
    set outputObjglhdr = jsonObjglhdr.parse(List_jsonStringglhdr)  'json.objectClass.asp가 필요함(파싱하려면)


    Set oJSON = New aspJSON     'aspJSON.asp 필요함(for문)
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

    Set caljson = new JSON  'JSON.asp가 필요함
    Set rs = connStr.Execute(sqlview)

    Dim List_jsonStringgllhp
    List_jsonStringgllhp = caljson.toJSON("results",rs, false)

    Dim jsonObjgllhp, outputObjgllhp
    set jsonObjgllhp = new JSONobject
    set outputObjgllhp = jsonObjgllhp.parse(List_jsonStringgllhp)  'json.objectClass.asp가 필요함(파싱하려면)


    Set oJSON = New aspJSON     'aspJSON.asp 필요함(for문)
    oJSON.loadJSON(List_jsonStringgllhp)

    Response.write List_jsonStringgllhp
end if

'tampilkan data journal'
if idview="5" then
    dim tahung5,bulang5,sourcegl5,filter
    tahun5 = Request("tahun")
    bulan5 = Request("bulan")
	source5= Request("sourcecode")
	company5 = Request("company")
	fcba5 = Request("fcba")
    

    Set List = Server.CreateObject("ADODB.Recordset")
	
	if trim(fcba5) = "-" then 
		filter = " AND G.FCBA in (SELECT fccode from businessunit where fccompanycode='" & company5 & "') "
	else
	    filter = " AND G.FCBA ='" & fcba5 & "' "
	end if


    sqlview =  "Select G.docno, G.docdate, nvl(G.remarks,'-') remarks, G.sourcecode, NVL(G.totamount,0) totamount, G.FCBA, G.POSTINGFLG " & _
			   " From   BKMJOURNALHDR G " & _
			   " WHERE TO_NUMBER(TO_CHAR(G.DOCDATE,'YYYY'))='" & tahun5 & "' " & _ 
			   " AND TO_NUMBER(TO_CHAR(G.DOCDATE,'MM'))='" & bulan5 & "' " & _ 
			   filter & _
			   " AND G.SOURCECODE Like '%" & source5 & "%'" & _
			   " ORDER BY G.fcba,G.DOCDATE,G.sourcecode,G.docno " 

    'response.write sqlview

    List.Open sqlview, connStr, 1
    ListCount = List.RecordCount 

    Set caljson = new JSON  'JSON.asp가 필요함
    Set rs = connStr.Execute(sqlview)

    Dim List_jsonString5
    List_jsonString5 = caljson.toJSON("results",rs, false)

    Dim jsonObj5, outputObj5
    set jsonObj5 = new JSONobject
    set outputObj5 = jsonObj5.parse(List_jsonString5)  'json.objectClass.asp가 필요함(파싱하려면)


    Set oJSON = New aspJSON     'aspJSON.asp 필요함(for문)
    oJSON.loadJSON(List_jsonString5)

    Response.write List_jsonString5
end if


'tampilkan data journal'
if idview="6" then
   ' dim tahun6,bulan6,filter,company6,fcba6
    tahun6 = Request("tahun")
    bulan6 = Request("bulan")
	company6 = Request("company")
	fcba6 = Request("fcba")
    

    Set List = Server.CreateObject("ADODB.Recordset")
	
	if trim(fcba6) = "-" then 
		filter = " AND FCBA in (SELECT fccode from businessunit where fccompanycode='" & company6 & "') "
	else
	    filter = " AND FCBA ='" & fcba6 & "' "
	end if


    sqlview =  "Select FCCODE,FCNAME,sum(opening_balance) opening_balance ,sum(debet) debet,sum(credit) credit ,sum(opening_balance)+sum(debet)+sum(credit) ending_balance  " & _
			   " From FAC_TRIALBALANCE  " & _
			   " WHERE period_year=" & tahun6 & " " & _ 
			   " AND period_month=" & bulan6 & " " & _ 
			   filter & _
			   " GROUP BY period_month,period_year,FCCODE,FCNAME " & _ 
			   " ORDER BY fccode " 

    'response.write sqlview

    'List.Open sqlview, connStr, 1
    'ListCount = List.RecordCount 

    Set caljson = new JSON  'JSON.asp가 필요함
    Set rs = connStr.Execute(sqlview)

    Dim List_jsonString6
    List_jsonString6 = caljson.toJSON("results",rs, false)

    Dim jsonObj6, outputObj6
    set jsonObj6 = new JSONobject
    set outputObj6 = jsonObj6.parse(List_jsonString6)  'json.objectClass.asp가 필요함(파싱하려면)


    Set oJSON = New aspJSON     'aspJSON.asp 필요함(for문)
    oJSON.loadJSON(List_jsonString6)

    Response.write List_jsonString6
end if

'tampilkan data tipe overtime hours'
if idview="7" then
  

    Set List = Server.CreateObject("ADODB.Recordset")
	

    sqlview =  "select distinct type id,fcname text" & _ 
    " from overtime_convertion " & _
    " where fcba='ARM' " & _
    " order by type asc " 


    Set caljson = new JSON  'JSON.asp가 필요함
    Set rs = connStr.Execute(sqlview)

    Dim List_jsonString7
    List_jsonString7 = caljson.toJSON("results",rs, false)

    Dim jsonObj7, outputObj7
    set jsonObj7 = new JSONobject
    set outputObj7 = jsonObj7.parse(List_jsonString7)  'json.objectClass.asp가 필요함(파싱하려면)


    Set oJSON = New aspJSON     'aspJSON.asp 필요함(for문)
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
		filter = " AND FCBA in (SELECT fccode from businessunit where fccompanycode='" & company8 & "') "
	else
	    filter = " AND FCBA ='" & fcba8 & "' "
	end if


    sqlview =  " Select FCBA,RKBNO,PERIOD_YEAR||'/'||PERIOD_MONTH period,SECTIONCODE," & _
	           " case document_classification " & _ 
			   " when '501' then 'Perawatan' " & _ 
			   " when '503' then 'Panen' " & _ 
			   " when '502' then 'Vehicle' " & _ 
			   " when '504' then 'Wip' End type, sum(TOTUPAH) TOTUPAH ,sum(TOTBAHAN) TOTBAHAN,SUM(TOTVRA) TOTVRA ,sum(TOTWIP) TOTWIP " & _
			   " From rkbhdr " & _
			   " WHERE period_year=" & tahun8 & " " & _ 
			   " AND period_month=" & bulan8 & " " & _ 
			   " AND document_classification like '" & trim(type8) & "%'" & _
			   filter & _
			   " GROUP BY FCBA,RKBNO,PERIOD_YEAR||'/'||PERIOD_MONTH,SECTIONCODE,document_classification " & _ 
			   " ORDER BY FCBA,period,document_classification,RKBNO " 

    'response.write sqlview

    'List.Open sqlview, connStr, 1
    'ListCount = List.RecordCount 

    Set caljson = new JSON  'JSON.asp가 필요함
    Set rs = connStr.Execute(sqlview)

    Dim List_jsonString8
    List_jsonString8 = caljson.toJSON("results",rs, false)

    Dim jsonObj8, outputObj8
    set jsonObj8 = new JSONobject
    set outputObj8 = jsonObj8.parse(List_jsonString8)  'json.objectClass.asp가 필요함(파싱하려면)


    Set oJSON = New aspJSON     'aspJSON.asp 필요함(for문)
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
	
    

    Set List = Server.CreateObject("ADODB.Recordset")
	
	if trim(fcba9) = "-" then 
		filter = " AND FCBA in (SELECT fccode from businessunit where fccompanycode='" & company9 & "') AND StockGroupcode Like '" & type9 & "%' "
	else
	    filter = " AND FCBA ='" & fcba9 & "' AND StockGroupcode Like '" & type9 & "%' "
	end if


    sqlview =   " select " & _
				" fcba , StockGroupname ,ActMat, stockcode, fcname, unitofmeasurement , " & _
				" SUM (opening_balance_day) opening_balance_day, " & _
				" SUM (received) received, " & _
				" SUM (issued) issued, " & _
				" SUM (adjustment) adjustment, " & _
				" SUM (closing_balance_day) closing_balance_day, " & _
				" stock_api.get_stockminimum(stockcode, fcba) get_stockminimum " & _
				" FROM ( " & _
				" SELECT " & _
				" USER_COMPANY_API.GetGroupPlantationCode('IT_HO1',a.fcba) GroupPlantation, " & _
				" USER_COMPANY_API.GetPlantationCode('IT_HO1',a.fcba) Plantation, " & _
				" USER_COMPANY_API.GetGroupRegionCode('IT_HO1',a.fcba) groupregion , " & _
				" USER_COMPANY_API.GetRegionCode('IT_HO1',a.fcba) region, " & _
				" USER_COMPANY_API.GETCOMPANYCODE('IT_HO1',a.fcba) company, " & _
				" BUSINESSUNIT_API.get_OrderData(a.fcba) norut, " & _
				" a.fcba, " & _
				" stock_api.get_stockgroup(a.stockcode,a.fcba) StockGroupcode, " & _
				" stockgroup_api.get_stockgrpname(stock_api.get_stockgroup(a.stockcode,a.fcba),a.fcba) stockgroupname, " & _
				" stock_api.get_ActiveMatrial(a.stockcode,a.fcba) ActMat , " & _
				" a.stockcode, " & _
				" stock_api.get_stockname(a.stockcode, a.fcba) fcname, " & _
				" stock_api.get_uom(a.stockcode, a.fcba) unitofmeasurement, " & _
				" SUM (a.opening_balance_day) opening_balance_day, " & _
				" SUM (a.received) received, " & _
				" SUM (issued) issued, " & _
				" SUM (a.adjustment) adjustment, " & _
				" SUM (a.opening_balance_day) + SUM (a.received) - SUM (a.issued) + SUM (a.adjustment) closing_balance_day, " & _
				" abs(SUM(a.opening_balance_day)) + abs(SUM(a.received)) + abs(SUM(a.issued)) nominal " & _
				" FROM " & _
				" 	  ( SELECT fcba, itemcode stockcode, openingbalanceqty opening_balance_day, " & _
				" 		0 received, 0 issued, 0 adjustment " & _
				" 		FROM stockledger WHERE year = " & Year(dt) & " AND month = " & Month(dt) & "" & _
				" 		UNION ALL " & _
				" 		SELECT fcba, stockcode, 0 opening_balance_day, " & _
				" 		quantity received, 0 issued, 0 adjustment " & _
				" 		FROM stocktranshist " & _
				" 		WHERE transtype IN ('RCVDIRECT', 'RECEIVED', 'MANUAL RECEIVED') " & _
				" 		AND transdate <= TO_DATE ('" & dt & "', 'MM/DD/YYYY') " & _
				" 		AND to_char(transdate, 'YYYYMM') = to_Char(TO_DATE ('" & dt & "', 'MM/DD/YYYY'), 'YYYYMM') " & _
				" 		UNION ALL " & _
				" 		SELECT fcba, stockcode, 0 opening_balance_day, 0 received, quantity * -1 issued, 0 adjustment " & _
				" 		FROM stocktranshist " & _
				" 		WHERE transtype IN ('ST_ISSUED', 'ST_TransportTask') AND transdate <= TO_DATE ('" & dt & "', 'MM/DD/YYYY') " & _
				" 		AND to_char(transdate, 'YYYYMM') = to_Char(TO_DATE ('" & dt & "', 'MM/DD/YYYY'), 'YYYYMM') " & _
				" 		UNION ALL SELECT fcba, stockcode, 0 opening_balance_day, 0 received, 0 issued, quantity adjustment " & _
				" 		FROM stocktranshist WHERE transtype = 'ST_WRITEOFF' " & _
				" 		 AND transdate <= TO_DATE ('" & dt & "', 'MM/DD/YYYY') " & _
				" 		 AND to_char(transdate, 'YYYYMM') = to_Char(TO_DATE ('" & dt & "', 'MM/DD/YYYY'), 'YYYYMM') " & _
				" 	  ) a " & _
				" GROUP BY a.fcba, a.stockcode  " & _
				" ) " & _
				" where company is not null " & _
				filter & _ 
				" AND STOCKGROUPCODE IN (select distinct fccode from stockgroup where grouptype= 'STOCK') " & _
				" AND nominal > 0 " & _
				" group by fcba , StockGroupname ,ActMat, stockcode, fcname, unitofmeasurement " & _
				" Order by StockGroupname,stockcode,fcba " 


    'response.write sqlview

    'List.Open sqlview, connStr, 1
    'ListCount = List.RecordCount 

    Set caljson = new JSON  'JSON.asp가 필요함
    Set rs = connStr.Execute(sqlview)

    Dim List_jsonString9
    List_jsonString9 = caljson.toJSON("results",rs, false)

    Dim jsonObj9, outputObj9
    set jsonObj9 = new JSONobject
    set outputObj9 = jsonObj9.parse(List_jsonString9)  'json.objectClass.asp가 필요함(파싱하려면)


    Set oJSON = New aspJSON     'aspJSON.asp 필요함(for문)
    oJSON.loadJSON(List_jsonString9)

    Response.write List_jsonString9
end if
'tampilkan data hpp harian'
if idview="10" then
   
    fcba10 = session("fcba")
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
	'connStr.Execute(sqlexec)

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

    Set caljson = new JSON  'JSON.asp가 필요함
    Set rs = connStr.Execute(sqlview)

    Dim List_jsonString10
    List_jsonString10 = caljson.toJSON("results",rs, false)

    Dim jsonObj10, outputObj10
    set jsonObj10 = new JSONobject
    set outputObj10 = jsonObj10.parse(List_jsonString10)  'json.objectClass.asp가 필요함(파싱하려면)


    Set oJSON = New aspJSON     'aspJSON.asp 필요함(for문)
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
		filter = " AND a.FCBA in (SELECT fccode from businessunit where fccompanycode='" & company9 & "') AND a.fccode Like '" & type11 & "%' "
	else
	    filter = " AND a.FCBA ='" & fcba11 & "' AND a.fccode Like '" & type11 & "%' "
	end if
	
    sqlview =  " SELECT a.fcba,a.fccode,a.fcname,a.accountcode,Nvl(nOpening,0) nOpening,Nvl(nDebet,0) nDebet,Nvl(nCredit,0) nCredit,Nvl(nOpening,0)+Nvl(nDebet,0)-Nvl(nCredit,0) AS nEnding " & _
				" FROM BANK a " & _
				" LEFT JOIN (SELECT fcba,period_year,period_month period_month,fccode,Nvl(opening_balance,0) nOpening " & _
				"            FROM trialbalance " & _
				"            WHERE period_year=" & tahun11 & " AND period_month=" & bulan11 & ") b ON (a.fcba=b.fcba AND a.accountcode=b.fccode) " & _
				" LEFT JOIN (SELECT fcba,TO_NUMBER(TO_CHAR(datecreated,'YYYY')) period_year,TO_NUMBER(TO_CHAR(datecreated,'MM')) period_month,bankcode,Sum(totalamount) nDebet " & _
				"            FROM paymentvoucher a" & _
				"            WHERE transtype='01-RECEIVE' AND TO_NUMBER(TO_CHAR(datecreated,'YYYY'))=" & tahun11 & " AND TO_NUMBER(TO_CHAR(datecreated,'MM'))=" & bulan11 & " " & _
				"            GROUP BY fcba,TO_NUMBER(TO_CHAR(datecreated,'YYYY')) ,TO_NUMBER(TO_CHAR(datecreated,'MM')),bankcode) c " & _
				"            ON (a.fcba=c.fcba AND Trim(a.fccode)=Trim(c.bankcode) AND b.period_month=c.period_month AND b.period_year=c.period_year) " & _
				" LEFT JOIN (SELECT fcba,TO_NUMBER(TO_CHAR(datecreated,'YYYY')) period_year,TO_NUMBER(TO_CHAR(datecreated,'MM')) period_month,bankcode,Sum(totalamount) nCredit " & _
				"            FROM paymentvoucher a" & _
				"            WHERE transtype='02-PAYMENT' AND TO_NUMBER(TO_CHAR(datecreated,'YYYY'))=" & tahun11 & " AND TO_NUMBER(TO_CHAR(datecreated,'MM'))=" & bulan11 & " " & _
				"            GROUP BY fcba,TO_NUMBER(TO_CHAR(datecreated,'YYYY')) ,TO_NUMBER(TO_CHAR(datecreated,'MM')),bankcode) d " & _
				"            ON (a.fcba=d.fcba AND a.fccode=d.bankcode AND b.period_month=d.period_month AND b.period_year=d.period_year) " & _
				" WHERE " & _
				" a.Activation = 'Y' AND UPPER(banktype) IN ('CASH', 'BANK') " & _
				filter & _
				" ORDER BY a.fcba,a.fccode "


    'response.write sqlview
    Set List = Server.CreateObject("ADODB.Recordset")
    List.Open sqlview, connStr, 1
    ListCount = List.RecordCount 

    Set caljson = new JSON  'JSON.asp가 필요함
    Set rs = connStr.Execute(sqlview)

    Dim List_jsonString11
    List_jsonString11 = caljson.toJSON("results",rs, false)

    Dim jsonObj11, outputObj11
    set jsonObj11 = new JSONobject
    set outputObj11 = jsonObj11.parse(List_jsonString11)  'json.objectClass.asp가 필요함(파싱하려면)


    Set oJSON = New aspJSON     'aspJSON.asp 필요함(for문)
    oJSON.loadJSON(List_jsonString11)

    Response.write List_jsonString11
end if

'tampilkan data hpp harian'
if idview="12" then
   
    tahun12 = Request("hppperiod")
    fcba12  = Session("fcba")
	
	sqlview =  	" SELECT a.fcba,a.description,a.accountingyear,a.accountingperiod,TO_CHAR(TO_DATE(accountingperiod, 'MM'), 'MONTH') period," & _
				" tbs_terima_qty,tbs_olah_qty," & _
				" cpo_beli_qty,cpo_proses_qty," & _
				" pk_beli_qty,pk_proses_qty," & _
				" cpo_awal_amt,pk_awal_amt," & _
				" cost_prod_amt,cost_btl_amt,cost_asset_amt," & _
				" cpo_akhir_amt,pk_akhir_amt," & _
				" cpo_hpp_amt,pk_hpp_amt " & _
				" FROM accountingyearperiod a left JOIN " & _
				" fac_hpp_bulanan b ON (Trim(a.fcba)=Trim(b.fcba) AND a.accountingperiod=b.period_MONTH AND a.accountingyear=b.period_year) " & _
				" WHERE a.fcba='" & fcba12 & "' AND accountingyear=" & tahun12 & "  AND accountingperiod <= 12 " & _
				" ORDER BY To_Number(accountingperiod) "


    'response.write sqlview
    Set List = Server.CreateObject("ADODB.Recordset")
    List.Open sqlview, connStr, 1
    ListCount = List.RecordCount 

    Set caljson = new JSON  'JSON.asp가 필요함
    Set rs = connStr.Execute(sqlview)

    Dim List_jsonString12
    List_jsonString12 = caljson.toJSON("results",rs, false)

    Dim jsonObj12, outputObj12
    set jsonObj12 = new JSONobject
    set outputObj12 = jsonObj12.parse(List_jsonString12)  'json.objectClass.asp가 필요함(파싱하려면)


    Set oJSON = New aspJSON     'aspJSON.asp 필요함(for문)
    oJSON.loadJSON(List_jsonString12)

    Response.write List_jsonString12
end if
%>