<%validatesession%>
<!--#include file="connDb1.Inc"-->
<%
	'POST DATA VALUE'
	fcba 		= Session("fcba")
	txttgl 		= Request("txttgl") 
	jamkerja 	= Request("jamkerja") 
	terimatbs 	= Request("terimatbs") 
	sisatbs 	= Request("sisatbs") 
	kirimcpo 	= Request("kirimcpo") 
	cotcpo 		= Request("cotcpo") 
	cstcpo 		= Request("cstcpo") 
	otcpo 		= Request("otcpo") 
	ffacpo 		= Request("ffacpo") 
	aircpo 		= Request("aircpo") 
	kotorancpo 	= Request("kotorancpo") 
	dobicpo 	= Request("dobicpo") 
	kirimkernel = Request("kirimkernel") 
	silokernel  = Request("silokernel") 
	kirimckg    = Request("kirimckg") 
	produksickg = Request("produksickg") 
	shift1 		= Request("shift1") 
	hk1 		= Request("hk1") 
	cuti1 		= Request("cuti1") 
	sakit1 		= Request("sakit1") 
	izin1 		= Request("izin1") 
	mangkir1 	= Request("mangkir1") 
	shift2 		= Request("shift2") 
	hk2 		= Request("hk2") 
	cuti2 		= Request("cuti2") 
	sakit2 		= Request("sakit2") 
	izin2 		= Request("izin2") 
	mangkir2 	= Request("mangkir2") 
	bulkingsilo = Request("bulkingsilo") 
	isibulking 	= Request("isibulking") 
	ffabulking 	= Request("ffabulking") 
	airbulking 	= Request("airbulking") 
	kotoranbulking 	= Request("kotoranbulking") 
	TankID1 	= Request("TankID1") 
	isicpo1 	= Request("isicpo1") 
	ffacpo1 	= Request("ffacpo1") 
	aircpo1 	= Request("aircpo1") 
	kotorancpo1 = Request("kotorancpo1") 
	TankID2 	= Request("TankID2") 
	isicpo2 	= Request("isicpo2") 
	ffacpo2 	= Request("ffacpo2") 
	aircpo2 	= Request("aircpo2") 
	kotorancpo2 = Request("kotorancpo2") 
	TankID3 	= Request("TankID3") 
	isicpo3 	= Request("isicpo3") 
	ffacpo3 	= Request("ffacpo3") 
	aircpo3 	= Request("aircpo3") 
	kotorancpo3 = Request("kotorancpo3") 


	'POST VALUE PRODUCTIONHK'
	shift1 		= Request("shift1") 
	hk1 		= Request("hk1") 
	cuti1 		= Request("cuti1") 
	sakit1 		= Request("sakit1") 
	izin1 		= Request("izin1") 
	mangkir1 	= Request("mangkir1") 
	shift2 		= Request("shift2") 
	hk2 		= Request("hk2") 
	cuti2 		= Request("cuti2") 
	sakit2 		= Request("sakit2") 
	izin2 		= Request("izin2") 
	mangkir2 	= Request("mangkir2") 
	pottbskeluar 	= Request("pottbskeluar") 
	adjcpo 	= Request("adjcpo") 

	if pottbskeluar = "" then 
		pottbskeluar = 0
	end if
	if adjcpo = "" then 
		adjcpo = 0
	end if
	if hk1 = "" then 
		hk1 = 0
	end if
	if cuti1 = "" then 
		cuti1 = 0
	end if
	if sakit1 = "" then 
		sakit1 = 0
	end if
	if izin1 = "" then 
		izin1 = 0
	end if
	if mangkir1 = "" then 
		mangkir1 = 0
	end if
	if hk2 = "" then 
		hk2 = 0
	end if
	if cuti2 = "" then 
		cuti2 = 0
	end if
	if sakit2 = "" then 
		sakit2 = 0
	end if
	if izin2 = "" then 
		izin2 = 0
	end if
	if mangkir2 = "" then 
		mangkir2 = 0
	end if

	'post value bulkingsilo'
	bulkingsilo 	= Request("bulkingsilo") 
	isibulking 		= Request("isibulking") 
	ffabulking 		= Request("ffabulking") 
	airbulking 		= Request("airbulking") 
	kotoranbulking 	= Request("kotoranbulking") 

	'penmbahan baru'
	totalcpo	= Request("totalcpo") 
	air			= Request("air")
	caco3		= Request("caco3") 
	keterangan	= Request("ket")
	jammulai	= Request("jammulai")
	jamselesai	= Request("jamselesai")
	
	'delete dulu productiontrn,productiontankcpo,productionbulkingsilo,productionhk'
	sqldelete = "delete from productiontrn where proddate=TO_DATE('" & txttgl & "', 'mm/dd/yyyy')"
	connStr.execute(sqldelete)

	sqldelete = "delete from productiontankcpo where proddate=TO_DATE('" & txttgl & "', 'mm/dd/yyyy')"
	connStr.execute(sqldelete)

	sqldelete = "delete from productionbulkingsilo where proddate=TO_DATE('" & txttgl & "', 'mm/dd/yyyy')"
	connStr.execute(sqldelete)

	sqldelete = "delete from productionhk where proddate=TO_DATE('" & txttgl & "', 'mm/dd/yyyy')"
	connStr.execute(sqldelete)

	'buat nomor lhp'
	dim bulansplit,tglsplit,thnsplit,sqlnomor,rsnomor,tglprev,lhpno
	bulansplit  = left(txttgl,2)
	tglsplit    = mid(txttgl,4,2)
	thnsplit    = right(txttgl,4)
	
	sqlnomor = "SELECT SUBSTR(lhpno,16,5) AS LastNo " & _
	" FROM productiontrn " & _
	" WHERE SUBSTR(lhpno,9,6)='" & thnsplit&bulansplit & "' " & _
	" AND FCBA='" & fcba & "' ORDER BY SUBSTR(lhpno,16,5) DESC FETCH FIRST 1 ROWS ONLY"
	Set rsnomor = Server.CreateObject("ADODB.Recordset")
	rsnomor.CursorType = 3
	rsnomor.CursorLocation = 3
	rsnomor.Open sqlnomor, connstr
	if rsnomor.eof then
			tglprev 		= "0000"
			currentnomor 	= tglprev + 1
			currentnomor 	= "0000" & currentnomor
			currentnomorfix = right(currentnomor,5)
			lhpno			=	fcba&"/"&"LHP"&"/"&thnsplit&bulansplit&"/"&currentnomorfix
			'Response.Write "ok"
	else 
			tglprev 		= rsnomor("LastNo")
			currentnomor 	= tglprev + 1
			currentnomor 	= "0000" & currentnomor
			currentnomorfix = right(currentnomor,5)
			lhpno			= fcba&"/"&"LHP"&"/"&thnsplit&bulansplit&"/"&currentnomorfix
	end if
	rsnomor.close 
	set rsnomor = nothing

	'insert data ke productiontrn'
	sqlsimpan = "insert into PRODUCTIONTRN (PRODDATE, JAMKERJA, TBSTERIMA, TBSSISA, CPOKIRIM, " & _ 
	" CPOCOT, CPOCST, CPOOT, CPOFFA, CPOAIR, CPOKOTORAN, KERNELKIRIM, KERNELSILO, " & _
	" CANGKANGKIRIM, CANGKANGPRODUKSI, USERID, TRANSDATE,FCBA,FCIP,TBSOLAH,CPODOBI,CPOTANKISI, " & _
	" PKTANKISI,JAMMULAI,JAMSELESAI,KETERANGAN,AIR,CACO3,POTONGANTBS,ADJCPO,LHPNO) " & _ 
	" VALUES (TO_DATE('" & txttgl & "', 'mm/dd/yyyy'),'" & jamkerja & "','" & terimatbs & "', " & _ 
	" '" & sisatbs & "','" & kirimcpo & "','" & cotcpo & "','" & cstcpo & "', " & _ 
	" '" & otcpo & "','" & ffacpo & "','" & aircpo & "','" & kotorancpo & "', " & _ 
	" '" & kirimkernel & "','" & silokernel & "','" & kirimckg & "','" & produksickg & "', " & _ 
	" '" & Session("CodeName") & "',to_date('" & txttgl & "','mm/dd/yyyy hh24:mi:ss'),'" & fcba & "', " & _
	" '" & Request.ServerVariables("REMOTE_ADDR") & "','0','" & dobicpo & "', " & _
	" '" & totalcpo & "','" & isibulking & "','" & jammulai & "','" & jamselesai & "', " & _
	" '" & keterangan & "','" & air & "','" & caco3 & "','" & pottbskeluar & "','" & adjcpo & "','" & lhpno & "')"
	'response.write sqlsimpan
	connStr.execute(sqlsimpan)

	'insert to productionhk'
	sqlsimpanhk = "insert into PRODUCTIONHK(PRODDATE, SHIFT, HK, CUTI, SAKIT,IZIN, MANGKIR,FCBA) " & _ 
	" VALUES(TO_DATE('" & txttgl & "', 'mm/dd/yyyy'),'" & shift1 & "','" & hk1 & "', " & _ 
	" '" & cuti1 & "','" & sakit1 & "','" & izin1 & "','" & mangkir1 & "','" & fcba & "') "
	connStr.execute(sqlsimpanhk)

	sqlsimpanhk = "insert into PRODUCTIONHK(PRODDATE, SHIFT, HK, CUTI, SAKIT,IZIN, MANGKIR,FCBA) " & _ 
	" VALUES(TO_DATE('" & txttgl & "', 'mm/dd/yyyy'),'" & shift2 & "','" & hk2 & "', " & _ 
	" '" & cuti2 & "','" & sakit2 & "','" & izin2 & "','" & mangkir2 & "','" & fcba & "') "
	connStr.execute(sqlsimpanhk)

	'INSERT TO PRODUCTION'
	sqlbulking = "insert into PRODUCTIONBULKINGSILO( PRODDATE, SILOID, BULKINGSILOISI, BULKINGSILOFFA, BULKINGSILOAIR, BULKINGSILOKOTORAN,FCBA,LHPNO) " & _ 
	" VALUES(TO_DATE('" & txttgl & "', 'mm/dd/yyyy'),'" & bulkingsilo & "','" & isibulking & "', " & _ 
	" '" & ffabulking & "','" & airbulking & "','" & kotoranbulking & "','" & fcba & "','" & lhpno & "')"
	connStr.execute(sqlbulking)

	'INSERT TO TANK CPO'
	sqltankcpo = "insert into PRODUCTIONTANKCPO( PRODDATE, TANKID, CPOISITANK, FFATANKCPO, AIRTANKCPO, KOTORANTANKCPO,FCBA,LHPNO) " & _ 
	" VALUES(TO_DATE('" & txttgl & "', 'mm/dd/yyyy'),'" & TankID1 & "','" & isicpo1 & "', " & _ 
	" '" & ffacpo1 & "','" & aircpo1 & "','" & kotorancpo1 & "','" & fcba & "','" & lhpno & "')"
	connStr.execute(sqltankcpo)

	sqltankcpo = "insert into PRODUCTIONTANKCPO( PRODDATE, TANKID, CPOISITANK, FFATANKCPO, AIRTANKCPO, KOTORANTANKCPO,FCBA,LHPNO) " & _ 
	" VALUES(TO_DATE('" & txttgl & "', 'mm/dd/yyyy'),'" & TankID2 & "','" & isicpo2 & "', " & _ 
	" '" & ffacpo2 & "','" & aircpo2 & "','" & kotorancpo2 & "','" & fcba & "','" & lhpno & "')"
	connStr.execute(sqltankcpo)

	sqltankcpo = "insert into PRODUCTIONTANKCPO( PRODDATE, TANKID, CPOISITANK, FFATANKCPO, AIRTANKCPO, KOTORANTANKCPO,FCBA,LHPNO) " & _ 
	" VALUES(TO_DATE('" & txttgl & "', 'mm/dd/yyyy'),'" & TankID3 & "','" & isicpo3 & "', " & _ 
	" '" & ffacpo3 & "','" & aircpo3 & "','" & kotorancpo3 & "','" & fcba & "','" & lhpno & "')"
	connStr.execute(sqltankcpo)
	
	sqllhpharian = "BEGIN fac_hpp_harian_api.ProcessHPP_Pabrik(TO_DATE('" & txttgl & "', 'mm/dd/yyyy'),'" & fcba & "');END;"
	connStr.execute(sqllhpharian)
	
	response.redirect("lhp.asp")
%>