
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


dim txtcari,pocode, company,fcba,tglHariini ,validasi
company    = Request("company")
fcba       = Request("fcba")
validasi   = Request("validasi")
tglHariini = textDate(date)


Set List = Server.CreateObject("ADODB.Recordset")

if validasi = 0 then 
  
sqlview = 			"   SELECT fccompanycode,fcba_name,id,kodelink," & _
			"   NVL(HMin1_point,0) HMin1_point,NVL(HMin2_point,0) HMin2_point,NVL(HMin3_point,0) HMin3_point,NVL(HMin4_point,0) HMin4_point " & _
			"	FROM " & _
			"	(" & _
			"	SELECT fccompanycode,fcba,businessunit_api.get_fcname(fcba) fcba_name,id,kodelink," & _
			"	ROW_NUMBER() OVER (PARTITION BY fccompanycode, fcba, id,kodelink ORDER BY TANGGAL DESC) AS TANGGAL_ID," & _
			"	NVL(Decode(Sum(basis),0,0,Round(Sum(hitung)/Sum(basis),2)*100),0) point" & _
			"	FROM" & _
			"	(" & _
			"	SELECT fccompanycode, fcba, id,kodelink, TANGGAL," & _
			"	COUNT(kodelink) OVER ( PARTITION BY fccompanycode,fcba,id,kodelink,tanggal) Basis," & _
			"	0 Hitung" & _
			"	FROM" & _
			"	(" & _
			"		/* Set PT,BU,KodeLink dan Tanggal */" & _
			"		SELECT b.fccompanycode, b.fccode fcba, c.id,c.kodelink, TANGGAL FROM" & _
			"		(" & _
			"		SELECT (TO_DATE('" & tglHariini & "', 'MM/DD/YYYY') - LEVEL ) AS tanggal" & _
			"		FROM dual" & _
			"		CONNECT BY LEVEL <= 4" & _
			"		) a" & _
			"		LEFT JOIN businessunit b ON (1=1)" & _
			"		LEFT JOIN" & _
			"		(" & _
			"		/* Kode link */" & _
			"		/* employee   ambil gangcode */" & _
			"		/* vehicle/ws ambil vehiclecode */" & _
			"		/* wip        ambil nomor kontrak */" & _
			"		SELECT fcba,'BKM' id,gangcode AS Kodelink FROM employee WHERE dateterminate IS NULL" & _
			"		UNION" & _
			"		SELECT fcba,'VRA' id,fccode AS Kodelink FROM vehicle  WHERE activation='Y'" & _
			"		UNION" & _
			"		SELECT fcba,'WSI' id,fccode AS Kodelink FROM workshop WHERE activation='Y'" & _
			"		UNION" & _
			"		SELECT fcba,'WIP' id,agreementcode AS Kodelink FROM contractagreement  WHERE  TO_DATE('" & tglHariini & "', 'MM/DD/YYYY') BETWEEN startdate AND finishdate" & _
			"		UNION" & _
			"		SELECT fcba,'SIS' id,fccode AS Kodelink FROM store  WHERE  iscentral='Y'" & _
			"		UNION" & _
			"		SELECT fcba,'BNK' id,fccode AS Kodelink FROM bank  WHERE   banktype='CASH'" & _
			"		)c ON (b.fccode=c.fcba)" & _
			"		WHERE b.isactive='TRUE' AND b.fccrop='OP'" & _
			"		AND ((b.fctype <> 'HO' AND id IN ('BKM','VRA','WSI','SIS','BNK','WIP')) or (b.fctype='HO' AND id IN ('WIP')))" & _
			"	) mst" & _
			"	/* Transaksi */" & _
			"	UNION ALL" & _
			"	SELECT DISTINCT bu.fccompanycode, bkm.fcba,'BKM' id, gangcode kodelink, bkmdate tanggal,0 basis, 1 Hitung" & _
			"	FROM  bkmhdr bkm INNER JOIN businessunit bu ON (bkm.fcba=bu.fccode) WHERE  document_classification IN  ('501','502','503')" & _
			"	AND bkmdate BETWEEN (TO_DATE ('" & tglHariini & "', 'MM/DD/YYYY')- 4)  AND  TO_DATE ('" & tglHariini & "', 'MM/DD/YYYY')-1" & _
			"	UNION ALL" & _
			"	SELECT DISTINCT bu.fccompanycode, bkm.fcba,Decode(bkm.vehiclecode,'WS1','WSI','VRA') id, bkm.vehiclecode kodelink, bkmdate tanggal,0 basis, 1 Hitung" & _
			"	FROM  bkmhdr bkm" & _
			"	INNER JOIN bkmvehicledtl bkmd ON (bkm.bkmno=bkmd.bkmno AND bkm.fcba=bkmd.fcba)" & _
			"	INNER JOIN businessunit bu ON (bkm.fcba=bu.fccode) WHERE  document_classification IN  ('502')" & _
			"	AND bkmdate BETWEEN (TO_DATE ('" & tglHariini & "', 'MM/DD/YYYY')- 4)  AND  TO_DATE ('" & tglHariini & "', 'MM/DD/YYYY')-1" & _
			"	UNION ALL" & _
			"	SELECT DISTINCT bu.fccompanycode, bkm.fcba,'WIP' id, bkmd.agreementcode kodelink, bkmdate tanggal,0 basis, 1 Hitung" & _
			"	FROM  bkmhdr bkm" & _
			"	INNER JOIN bkmwipdtl bkmd ON (bkm.bkmno=bkmd.bkmno AND bkm.fcba=bkmd.fcba)" & _
			"	INNER JOIN businessunit bu ON (bkm.fcba=bu.fccode) WHERE  document_classification IN  ('504')" & _
			"	AND bkmdate BETWEEN (TO_DATE ('" & tglHariini & "', 'MM/DD/YYYY')- 4)  AND  TO_DATE ('" & tglHariini & "', 'MM/DD/YYYY')-1" & _
			"	UNION ALL" & _
			"	SELECT DISTINCT bu.fccompanycode, iss.fcba,'SIS' id, storecode kodelink, sivdate tanggal,0 basis, 1 Hitung" & _
			"	FROM issue iss INNER JOIN businessunit bu ON (iss.fcba=bu.fccode)" & _
			"	WHERE iss.storecode='ST1'" & _
			"	AND sivdate BETWEEN (TO_DATE ('" & tglHariini & "', 'MM/DD/YYYY')- 4)  AND  TO_DATE ('" & tglHariini & "', 'MM/DD/YYYY')-1" & _
			"	UNION ALL" & _
			"	SELECT DISTINCT bu.fccompanycode, bnk.fcba,'BNK' id, bankcode kodelink, datecreated tanggal,0 basis, 1 Hitung" & _
			"	FROM paymentvoucher bnk INNER JOIN businessunit bu ON (bnk.fcba=bu.fccode)" & _
			"	WHERE " & _
			"	datecreated BETWEEN (TO_DATE ('" & tglHariini & "', 'MM/DD/YYYY')- 4)  AND  TO_DATE ('" & tglHariini & "', 'MM/DD/YYYY')-1" & _
			"	)" & _
			"	WHERE fccompanycode like '" & company & "%'" & _
			"	GROUP BY  fccompanycode,fcba,id,kodelink,TANGGAL" & _
			"	)" & _
			"	pivot" & _
			"	(" & _
			"		MAX(nvl(point,0)) AS point" & _
			"		FOR TANGGAL_ID IN (1 AS HMin1, 2 HMin2, 3 HMin3, 4 HMin4 "  & _
			"                          ) " & _
			"	)" & _
			"	ORDER BY  fccompanycode,fcba,id,kodelink"

end if

if validasi = 1 then 
sqlview =   "   SELECT companycode,NVL(HMin1_point,0) HMin1_point,NVL(HMin2_point,0) HMin2_point,NVL(HMin3_point,0) HMin3_point,NVL(HMin4_point,0) HMin4_point " & _
			"	FROM" & _
			"	(" & _
			"	SELECT companycode," & _
			"	ROW_NUMBER() OVER (PARTITION BY companycode ORDER BY TANGGAL DESC) AS TANGGAL_ID," & _
			"	NVL(Decode(Sum(basis),0,0,Round(Sum(hitung)/Sum(basis),2)*100),0) point" & _
			"	FROM" & _
			"	(" & _
			"	SELECT companycode, id, TANGGAL," & _
			"	COUNT(id) OVER ( PARTITION BY companycode,id,tanggal) Basis," & _
			"	0 Hitung" & _
			"	FROM" & _
			"	(" & _
			"		SELECT b.companycode, c.id, TANGGAL FROM" & _
			"		(" & _
			"		SELECT (TO_DATE('" & tglHariini & "', 'MM/DD/YYYY') - LEVEL ) AS tanggal" & _
			"		FROM dual" & _
			"		CONNECT BY LEVEL <= 4" & _
			"		) a" & _
			"		LEFT JOIN company b ON (1=1)" & _
			"		INNER JOIN" & _
			"		(" & _
			"		SELECT DISTINCT fccompanycode companycode,'BKM' id FROM businessunit WHERE isactive='TRUE' AND fccrop='OP'  AND fctype <> 'HO'" & _
			"		UNION" & _
			"		SELECT DISTINCT fccompanycode companycode,'WIP' id FROM businessunit WHERE isactive='TRUE' AND fccrop='OP'  AND fctype <> 'HO'" & _
			"		UNION" & _
			"		SELECT DISTINCT fccompanycode companycode,'BNK' id FROM businessunit WHERE isactive='TRUE' AND fccrop='OP' AND fctype <> 'HO'" & _
			"		UNION" & _
			"		SELECT DISTINCT fccompanycode companycode,'WSI' id FROM businessunit WHERE isactive='TRUE' AND fccrop='OP'  AND fctype <> 'HO'" & _
			"		UNION" & _
			"		SELECT DISTINCT fccompanycode companycode,'VRA' id FROM businessunit WHERE isactive='TRUE' AND fccrop='OP' AND fctype <> 'HO'" & _
			"		UNION" & _
			"		SELECT DISTINCT fccompanycode companycode,'SIS' id FROM businessunit WHERE isactive='TRUE' AND fccrop='OP' AND fctype <> 'HO'" & _
			"		)c ON (b.companycode=c.companycode)" & _
			"	) mst" & _
			"	UNION ALL" & _
			"	SELECT DISTINCT bu.fccompanycode companycode,'BKM' id, bkmdate tanggal,0 basis, 1 Hitung" & _
			"	FROM  bkmhdr bkm INNER JOIN businessunit bu ON (bkm.fcba=bu.fccode) WHERE  document_classification IN  ('501','502','503')" & _
			"	AND bkmdate BETWEEN (TO_DATE ('" & tglHariini & "', 'MM/DD/YYYY')- 4)  AND  TO_DATE ('" & tglHariini & "', 'MM/DD/YYYY')-1" & _
			"	UNION ALL" & _
			"	SELECT DISTINCT bu.fccompanycode companycode,Decode(bkm.vehiclecode,'WS1','WSI','VRA') id, bkmdate tanggal,0 basis, 1 Hitung" & _
			"	FROM  bkmhdr bkm" & _
			"	INNER JOIN bkmvehicledtl bkmd ON (bkm.bkmno=bkmd.bkmno AND bkm.fcba=bkmd.fcba)" & _
			"	INNER JOIN businessunit bu ON (bkm.fcba=bu.fccode) WHERE  document_classification IN  ('502')" & _
			"	AND bkmdate BETWEEN (TO_DATE ('" & tglHariini & "', 'MM/DD/YYYY')- 4)  AND  TO_DATE ('" & tglHariini & "', 'MM/DD/YYYY')-1" & _
			"	UNION ALL" & _
			"	SELECT DISTINCT bu.fccompanycode companycode,'WIP' id,  bkmdate tanggal,0 basis, 1 Hitung" & _
			"	FROM  bkmhdr bkm" & _
			"	INNER JOIN bkmwipdtl bkmd ON (bkm.bkmno=bkmd.bkmno AND bkm.fcba=bkmd.fcba)" & _
			"	INNER JOIN businessunit bu ON (bkm.fcba=bu.fccode) WHERE  document_classification IN  ('504')" & _
			"	AND bkmdate BETWEEN (TO_DATE ('" & tglHariini & "', 'MM/DD/YYYY')- 4)  AND  TO_DATE ('" & tglHariini & "', 'MM/DD/YYYY')-1" & _
			"	UNION ALL" & _
			"	SELECT DISTINCT bu.fccompanycode companycode,'SIS' id, sivdate tanggal,0 basis, 1 Hitung" & _
			"	FROM issue iss INNER JOIN businessunit bu ON (iss.fcba=bu.fccode)" & _
			"	WHERE iss.storecode='ST1'" & _
			"	AND sivdate BETWEEN (TO_DATE ('" & tglHariini & "', 'MM/DD/YYYY')- 4)  AND  TO_DATE ('" & tglHariini & "', 'MM/DD/YYYY')-1" & _
			"	UNION ALL" & _
			"	SELECT DISTINCT bu.fccompanycode companycode,'BNK' id, datecreated tanggal,0 basis, 1 Hitung" & _
			"	FROM paymentvoucher bnk INNER JOIN businessunit bu ON (bnk.fcba=bu.fccode)" & _
			"	WHERE bnk.bankcode IN ('K01','K02')" & _
			"	AND datecreated BETWEEN (TO_DATE ('" & tglHariini & "', 'MM/DD/YYYY')- 4)  AND  TO_DATE ('" & tglHariini & "', 'MM/DD/YYYY')-1" & _
			"	)" & _
			"	GROUP BY  companycode,TANGGAL" & _
			"	)" & _
			"	pivot" & _
			"	(" & _
			"		MAX(point) AS point" & _
			"		FOR TANGGAL_ID IN (1 AS HMin1, 2 HMin2, 3 HMin3, 4 HMin4) " & _
			"	)" & _
			"  ORDER BY  companycode "
end if


			'Response.write sqlview
			
     List.Open sqlview, connStr, 1
     ListCount = List.RecordCount 

     Set caljson = new JSON  
     Set rs = connStr.Execute(sqlview)

     Dim List_jsonString
     List_jsonString = caljson.toJSON("dataview",rs, false)

     Dim jsonObj, outputObj
     set jsonObj = new JSONobject
     set outputObj = jsonObj.parse(List_jsonString)  


     Set oJSON = New aspJSON    
     oJSON.loadJSON(List_jsonString)

     Response.write List_jsonString	


%>