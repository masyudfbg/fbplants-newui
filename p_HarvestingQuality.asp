<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>ePLaS Reporting  - Harvesting Pinalty</title>
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
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

v_txtRECEPTIONDATE_From= trim(request("FDDATE_From"))
v_txtRECEPTIONDATE_To= trim(request("FDDATE_To"))
v_txtEMPCODE=trim(request("EMPLOYEECODE"))
v_cField=trim(REQUEST("cField"))

if  len(trim(v_cField)) = 0 then
	txtSQL = 	" SELECT " & _
							" NVL(SUM(UNDER_RIPE),0) BUAHMENTAH," & _
							" NVL(SUM(OVER_RIPE),0) BUAHTERLALUMATANG," & _
							" NVL(SUM(ABNORMAL),0) ABNORMAL," & _
							" NVL(SUM(LONG_STALK),0) TANGKAIPANJANG," & _
							" NVL(SUM(EATEN_BY_RAT),0) BUAHDIMAKANTIKUS," & _
							" NVL(SUM(UNHARVEST_FFB),0) BUAHTIDAKDIPANEN," & _
							" NVL(SUM(UNCOLLECT_LF_CIRCLE),0) BRONDOLANDIPIRINGAN," & _
							" NVL(SUM(UNCOLLECT_LF_PIECE),0) BUAHTINGGALDIPIRINGAN," & _
							" NVL(SUM(UNARRANGE_FFB),0) BUAHTIDAKDISUSUNRAPI," & _
							" NVL(SUM(UNPRUNE_FROND),0) PELEPAHSENGKLEH," & _
							" NVL(SUM(QE_1),0) PELEPAHTIDAKDISUSUNRAPI," & _
							" NVL(SUM(QE_2),0) BUAHMATAHARI," & _
							" NVL(SUM(QE_3),0) BUAHBUSUK," & _
							" NVL(SUM(QE_4),0) QE_1," & _
							" NVL(SUM(QE_5),0) QE_2," & _
							" NVL(SUM(QE_6),0) QE_3," & _
							" NVL(SUM(QE_7),0) QE_4," & _
							" NVL(SUM(QE_8),0) QE_5," & _
							" NVL(SUM(QE_9),0) QE_6," & _
							" NVL(SUM(QE_10),0) QE_7" & _
				" FROM HARVESTINGQUALITY" & _
				" WHERE FCBA='" & Session("fcBA") & "'"
else
	IF instr(v_cField,"EMPCODE") > 0 THEN
		addname	=	", EMPLOYEE_API.get_EmployeeName(fcba, EMPCODE) empname"
		groupbyname = ", EMPLOYEE_API.get_EmployeeName(fcba, EMPCODE)"
	ELSE
		addname	=	""
		groupbyname = ""
	END IF
			   txtSQL = " SELECT " & v_cField & addname & "," & _
										" NVL(SUM(UNDER_RIPE),0) BUAHMENTAH," & _
										" NVL(SUM(OVER_RIPE),0) BUAHTERLALUMATANG," & _
										" NVL(SUM(ABNORMAL),0) ABNORMAL," & _
										" NVL(SUM(LONG_STALK),0) TANGKAIPANJANG," & _
										" NVL(SUM(EATEN_BY_RAT),0) BUAHDIMAKANTIKUS," & _
										" NVL(SUM(UNHARVEST_FFB),0) BUAHTIDAKDIPANEN," & _
										" NVL(SUM(UNCOLLECT_LF_CIRCLE),0) BRONDOLANDIPIRINGAN," & _
										" NVL(SUM(UNCOLLECT_LF_PIECE),0) BUAHTINGGALDIPIRINGAN," & _
										" NVL(SUM(UNARRANGE_FFB),0) BUAHTIDAKDISUSUNRAPI," & _
										" NVL(SUM(UNPRUNE_FROND),0) PELEPAHSENGKLEH," & _
										" NVL(SUM(QE_1),0) PELEPAHTIDAKDISUSUNRAPI," & _
										" NVL(SUM(QE_2),0) BUAHMATAHARI," & _
										" NVL(SUM(QE_3),0) BUAHBUSUK," & _
										" NVL(SUM(QE_4),0) QE_1," & _
										" NVL(SUM(QE_5),0) QE_2," & _
										" NVL(SUM(QE_6),0) QE_3," & _
										" NVL(SUM(QE_7),0) QE_4," & _
										" NVL(SUM(QE_8),0) QE_5," & _
										" NVL(SUM(QE_9),0) QE_6," & _
										" NVL(SUM(QE_10),0) QE_7" & _
						" FROM HARVESTINGQUALITY" & _
						" WHERE FCBA = '" & Session("fcBA") & "'"
end if

if  len(trim(v_txtRECEPTIONDATE_From)) > 0 then
    txtSQL = txtSQL & " AND FDDATE >= TO_DATE('" & v_txtRECEPTIONDATE_From & "','MM/DD/YYYY')"
end if
if  len(trim(v_txtRECEPTIONDATE_To)) > 0 then
    txtSQL =txtSQL &  " AND FDDATE <= TO_DATE('" & v_txtRECEPTIONDATE_To & "','MM/DD/YYYY')"
end if
if  len(trim(v_txtEMPCODE)) > 0 then
    txtSQL = txtSQL & " AND EMPCODE = '" & v_txtEMPCODE & "'"
end if

if  len(trim(v_cField)) > 0 then
	txtsql=(txtsql) & " GROUP BY " & v_cField & groupbyname
end if

' response.Write(txtsql)
response.Flush()
%>
<%

Dim Kolom

Kolom	= 0

if instr(v_cField,"FDDATE") > 0 then
	Kolom	=	Kolom + 1
End if
if instr(v_cField,"EMPCODE") > 0 then
	Kolom	=	Kolom + 2
End if
if instr(v_cField,"FIELDCODE") > 0 then
	Kolom	=	Kolom + 1
End if

%>
	<body>
		<div id="wrapper">
        <!--#include file="print_excel.asp"-->
        <table border="0" width="100%" id="product-table" class="xthf-site">
         <thead>
         <tr>
         <th colspan="<%=Kolom + 21%>" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=Session("COMPANY")%></h4>
					<h5><%=Session("BA")%></h5>
				</div>
				<div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h2>HARVESTING PINALTY</h2>
				</div>
			</div>
</th>
			</tr>
						<!--<div class="printhide">-->
						<tr>
							<td colspan="14" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></td>
				</tr>
				<!--</div>-->
		<tr>
		<%	
		Set rsCat = connStr.Execute(txtSQL)
		if Not rsCat.eof then
		   nCountofField=rsCat.fields.count-1
		end if
		NO=1%>
		</tr>
		<tr>
	    <th style="color:#fff;">No</th>
			
        <%if instr(v_cField,"FDDATE") > 0 then%>
			<th style="color:#fff;">Tanggal</th>
		<%end if%>
        <%if instr(v_cField,"EMPCODE") > 0 then%>
			<th style="color:#fff;">NIK</th>
			<th style="color:#fff;">Nama</th>
		<%end if%>
        <%if instr(v_cField,"FIELDCODE") > 0 then%>
			<th style="color:#fff;">Blok</th>	
		<%end if%>
		<th style="color:#fff;">Buah Mentah</th>
		<th style="color:#fff;">Buah Terlalu Matang</th>
		<th style="color:#fff;">Abnormal</th>
		<th style="color:#fff;">Tangkai Panjang</th>
		<th style="color:#fff;">Buah Dimakan Tikus</th>
		<th style="color:#fff;">Buah Tinggal</th>
		<th style="color:#fff;">Brondolan Dipiringan</th>
		<th style="color:#fff;">Buah Tinggal Dipiringan</th>
		<th style="color:#fff;">Buah Tidak Disusun Rapi di TPH</th>
		<th style="color:#fff;">Pelepah Sengkleh</th>
		<th style="color:#fff;">Pelepah Tidak Disusun Rapi</th>
		<th style="color:#fff;">Buah Matahari</th>
		<th style="color:#fff;">Buah Busuk</th>
		<th style="color:#fff;">Buah Mentah di Peram</th>
		<th style="color:#fff;">Over Pruning</th>
		<th style="color:#fff;">Brondolan Tidak di Alas</th>
		<th style="color:#fff;">Brondolan Kotor/ Sampah</th>
		<th style="color:#fff;">Buah di Belah</th>
		<th style="color:#fff;">Buah Tidak di Antrikan/ di Tulis</th>
		<th style="color:#fff;">Buah Tidak Diletakan Di TPH</th>

	</tr>
    </thead>
    <tbody>
    <%
    Set rsCat = Server.CreateObject("ADODB.Recordset")
   rsCat.CursorType = 3
   rsCat.CursorLocation = 3
   rsCat.Open txtSql, connstr
   if Not rsCat.eof then
	nCountofField=rsCat.fields.count-1
	no=1
	nQty=0
	nAmount=0
	
	Dim nNIK
	Dim nTANGGAL
	Dim nBLOK
	Dim nBUAHMENTAH
	Dim nBUAHTERLALUMATANG
	Dim nABNORMAL
	Dim nTANGKAIPANJANG
	Dim nBUAHDIMAKANTIKUS
	Dim nBUAHTIDAKDIPANEN
	Dim nBRONDOLANDIPIRINGAN
	Dim nBUAHTINGGALDIPIRINGAN
	Dim nBUAHTIDAKDISUSUNRAPI
	Dim nPELEPAHSENGKLEH
	Dim nPELEPAHTIDAKDISUSUNRAPI
	Dim nBUAHMATAHARI
	Dim nBUAHBUSUK
	Dim nQE_1
	Dim nQE_2
	Dim nQE_3
	Dim nQE_4
	Dim nQE_5
	Dim nQE_6
	Dim nQE_7

	nNIK = 0
	nTANGGAL = 0
	nBLOK = 0
	nBUAHMENTAH = 0
	nBUAHTERLALUMATANG = 0
	nABNORMAL = 0
	nTANGKAIPANJANG = 0
	nBUAHDIMAKANTIKUS = 0
	nBUAHTIDAKDIPANEN = 0
	nBRONDOLANDIPIRINGAN = 0
	nBUAHTINGGALDIPIRINGAN = 0
	nBUAHTIDAKDISUSUNRAPI = 0
	nPELEPAHSENGKLEH = 0
	nPELEPAHTIDAKDISUSUNRAPI = 0
	nBUAHMATAHARI = 0
	nBUAHBUSUK = 0
	nQE_1 = 0
	nQE_2 = 0
	nQE_3 = 0
	nQE_4 = 0
	nQE_5 = 0
	nQE_6 = 0
	nQE_7 = 0

	While Not rsCat.EOF%>
	  		<tr>
			<td align="Right" style="<%=cmdstyleLeft%>"><font face="Verdana" size="1"><%=NO%></td>
        <%if instr(v_cField,"FDDATE") > 0 then%>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("FDDATE")%></td>
		<%end if%>
        <%if instr(v_cField,"EMPCODE") > 0 then%>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("EMPCODE")%></td>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("EMPNAME")%></td>
		<%end if%>
        <%if instr(v_cField,"FIELDCODE") > 0 then%>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("FIELDCODE")%></td>
		<%end if%>			
			<%IF CDBL(rscat("BUAHMENTAH")) = 0 THEN%><td align="Right"></td><%ELSE%> <td align="Right"><%=FormatNumber(rscat("BUAHMENTAH"), 2)%></td><%END IF%>
			<%IF CDBL(rscat("BUAHTERLALUMATANG")) = 0 THEN%><td align="Right"></td><%ELSE%> <td align="Right"><%=FormatNumber(rscat("BUAHTERLALUMATANG"), 2)%></td><%END IF%>
			<%IF CDBL(rscat("ABNORMAL")) = 0 THEN%><td align="Right"></td><%ELSE%> <td align="Right"><%=FormatNumber(rscat("ABNORMAL"), 2)%></td><%END IF%>
			<%IF CDBL(rscat("TANGKAIPANJANG")) = 0 THEN%><td align="Right"></td><%ELSE%> <td align="Right"><%=FormatNumber(rscat("TANGKAIPANJANG"), 2)%></td><%END IF%>
			<%IF CDBL(rscat("BUAHDIMAKANTIKUS")) = 0 THEN%><td align="Right"></td><%ELSE%> <td align="Right"><%=FormatNumber(rscat("BUAHDIMAKANTIKUS"), 2)%></td><%END IF%>
			<%IF CDBL(rscat("BUAHTIDAKDIPANEN")) = 0 THEN%><td align="Right"></td><%ELSE%> <td align="Right"><%=FormatNumber(rscat("BUAHTIDAKDIPANEN"), 2)%></td><%END IF%>
			<%IF CDBL(rscat("BRONDOLANDIPIRINGAN")) = 0 THEN%><td align="Right"></td><%ELSE%> <td align="Right"><%=FormatNumber(rscat("BRONDOLANDIPIRINGAN"), 2)%></td><%END IF%>
			<%IF CDBL(rscat("BUAHTINGGALDIPIRINGAN")) = 0 THEN%><td align="Right"></td><%ELSE%> <td align="Right"><%=FormatNumber(rscat("BUAHTINGGALDIPIRINGAN"), 2)%></td><%END IF%>
			<%IF CDBL(rscat("BUAHTIDAKDISUSUNRAPI")) = 0 THEN%><td align="Right"></td><%ELSE%> <td align="Right"><%=FormatNumber(rscat("BUAHTIDAKDISUSUNRAPI"), 2)%></td><%END IF%>
			<%IF CDBL(rscat("PELEPAHSENGKLEH")) = 0 THEN%><td align="Right"></td><%ELSE%> <td align="Right"><%=FormatNumber(rscat("PELEPAHSENGKLEH"), 2)%></td><%END IF%>
			<%IF CDBL(rscat("PELEPAHTIDAKDISUSUNRAPI")) = 0 THEN%><td align="Right"></td><%ELSE%> <td align="Right"><%=FormatNumber(rscat("PELEPAHTIDAKDISUSUNRAPI"), 2)%></td><%END IF%>
			<%IF CDBL(rscat("BUAHMATAHARI")) = 0 THEN%><td align="Right"></td><%ELSE%> <td align="Right"><%=FormatNumber(rscat("BUAHMATAHARI"), 2)%></td><%END IF%>
			<%IF CDBL(rscat("BUAHBUSUK")) = 0 THEN%><td align="Right"></td><%ELSE%> <td align="Right"><%=FormatNumber(rscat("BUAHBUSUK"), 2)%></td><%END IF%>
			<%IF CDBL(rscat("QE_1")) = 0 THEN%><td align="Right"></td><%ELSE%> <td align="Right"><%=FormatNumber(rscat("QE_1"), 2)%></td><%END IF%>
			<%IF CDBL(rscat("QE_2")) = 0 THEN%><td align="Right"></td><%ELSE%> <td align="Right"><%=FormatNumber(rscat("QE_2"), 2)%></td><%END IF%>
			<%IF CDBL(rscat("QE_3")) = 0 THEN%><td align="Right"></td><%ELSE%> <td align="Right"><%=FormatNumber(rscat("QE_3"), 2)%></td><%END IF%>
			<%IF CDBL(rscat("QE_4")) = 0 THEN%><td align="Right"></td><%ELSE%> <td align="Right"><%=FormatNumber(rscat("QE_4"), 2)%></td><%END IF%>
			<%IF CDBL(rscat("QE_5")) = 0 THEN%><td align="Right"></td><%ELSE%> <td align="Right"><%=FormatNumber(rscat("QE_5"), 2)%></td><%END IF%>
			<%IF CDBL(rscat("QE_6")) = 0 THEN%><td align="Right"></td><%ELSE%> <td align="Right"><%=FormatNumber(rscat("QE_6"), 2)%></td><%END IF%>
			<%IF CDBL(rscat("QE_7")) = 0 THEN%><td align="Right"></td><%ELSE%> <td align="Right"><%=FormatNumber(rscat("QE_7"), 2)%></td><%END IF%>
			</tr>
		  <%
		no = no + 1
		
		nBUAHMENTAH = nBUAHMENTAH + Cdbl(rscat("BUAHMENTAH"))
		nBUAHTERLALUMATANG = nBUAHTERLALUMATANG + Cdbl(rscat("BUAHTERLALUMATANG"))
		nABNORMAL = nABNORMAL + Cdbl(rscat("ABNORMAL"))
		nTANGKAIPANJANG = nTANGKAIPANJANG + Cdbl(rscat("TANGKAIPANJANG"))
		nBUAHDIMAKANTIKUS = nBUAHDIMAKANTIKUS + Cdbl(rscat("BUAHDIMAKANTIKUS"))
		nBUAHTIDAKDIPANEN = nBUAHTIDAKDIPANEN + Cdbl(rscat("BUAHTIDAKDIPANEN"))
		nBRONDOLANDIPIRINGAN = nBRONDOLANDIPIRINGAN + Cdbl(rscat("BRONDOLANDIPIRINGAN"))
		nBUAHTINGGALDIPIRINGAN = nBUAHTINGGALDIPIRINGAN + Cdbl(rscat("BUAHTINGGALDIPIRINGAN"))
		nBUAHTIDAKDISUSUNRAPI = nBUAHTIDAKDISUSUNRAPI + Cdbl(rscat("BUAHTIDAKDISUSUNRAPI"))
		nPELEPAHSENGKLEH = nPELEPAHSENGKLEH + Cdbl(rscat("PELEPAHSENGKLEH"))
		nPELEPAHTIDAKDISUSUNRAPI = nPELEPAHTIDAKDISUSUNRAPI + Cdbl(rscat("PELEPAHTIDAKDISUSUNRAPI"))
		nBUAHMATAHARI = nBUAHMATAHARI + Cdbl(rscat("BUAHMATAHARI"))
		nBUAHBUSUK = nBUAHBUSUK + Cdbl(rscat("BUAHBUSUK"))
		nQE_1 = nQE_1 + Cdbl(rscat("QE_1"))
		nQE_2 = nQE_2 + Cdbl(rscat("QE_2"))
		nQE_3 = nQE_3 + Cdbl(rscat("QE_3"))
		nQE_4 = nQE_4 + Cdbl(rscat("QE_4"))
		nQE_5 = nQE_5 + Cdbl(rscat("QE_5"))
		nQE_6 = nQE_6 + Cdbl(rscat("QE_6"))
		nQE_7 = nQE_7 + Cdbl(rscat("QE_7"))

		  rsCat.MoveNext
	  Wend
%>

		<tr style="background-color:#F4F4A0;">
			<td align="center" colspan="<%=Kolom + 1%>" ><strong>Total</strong></td>
			<%IF nBUAHMENTAH = 0 THEN%><td align="Right"><strong></strong></td><%ELSE%> <td align="Right"><strong><%=FormatNumber(nBUAHMENTAH, 2)%></strong></td><%END IF%>
			<%IF nBUAHTERLALUMATANG = 0 THEN%><td align="Right"><strong></strong></td><%ELSE%> <td align="Right"><strong><%=FormatNumber(nBUAHTERLALUMATANG, 2)%></strong></td><%END IF%>
			<%IF nABNORMAL = 0 THEN%><td align="Right"><strong></strong></td><%ELSE%> <td align="Right"><strong><%=FormatNumber(nABNORMAL, 2)%></strong></td><%END IF%>
			<%IF nTANGKAIPANJANG = 0 THEN%><td align="Right"><strong></strong></td><%ELSE%> <td align="Right"><strong><%=FormatNumber(nTANGKAIPANJANG, 2)%></strong></td><%END IF%>
			<%IF nBUAHDIMAKANTIKUS = 0 THEN%><td align="Right"><strong></strong></td><%ELSE%> <td align="Right"><strong><%=FormatNumber(nBUAHDIMAKANTIKUS, 2)%></strong></td><%END IF%>
			<%IF nBUAHTIDAKDIPANEN = 0 THEN%><td align="Right"><strong></strong></td><%ELSE%> <td align="Right"><strong><%=FormatNumber(nBUAHTIDAKDIPANEN, 2)%></strong></td><%END IF%>
			<%IF nBRONDOLANDIPIRINGAN = 0 THEN%><td align="Right"><strong></strong></td><%ELSE%> <td align="Right"><strong><%=FormatNumber(nBRONDOLANDIPIRINGAN, 2)%></strong></td><%END IF%>
			<%IF nBUAHTINGGALDIPIRINGAN = 0 THEN%><td align="Right"><strong></strong></td><%ELSE%> <td align="Right"><strong><%=FormatNumber(nBUAHTINGGALDIPIRINGAN, 2)%></strong></td><%END IF%>
			<%IF nBUAHTIDAKDISUSUNRAPI = 0 THEN%><td align="Right"><strong></strong></td><%ELSE%> <td align="Right"><strong><%=FormatNumber(nBUAHTIDAKDISUSUNRAPI, 2)%></strong></td><%END IF%>
			<%IF nPELEPAHSENGKLEH = 0 THEN%><td align="Right"><strong></strong></td><%ELSE%> <td align="Right"><strong><%=FormatNumber(nPELEPAHSENGKLEH, 2)%></strong></td><%END IF%>
			<%IF nPELEPAHTIDAKDISUSUNRAPI = 0 THEN%><td align="Right"><strong></strong></td><%ELSE%> <td align="Right"><strong><%=FormatNumber(nPELEPAHTIDAKDISUSUNRAPI, 2)%></strong></td><%END IF%>
			<%IF nBUAHMATAHARI = 0 THEN%><td align="Right"><strong></strong></td><%ELSE%> <td align="Right"><strong><%=FormatNumber(nBUAHMATAHARI, 2)%></strong></td><%END IF%>
			<%IF nBUAHBUSUK = 0 THEN%><td align="Right"><strong></strong></td><%ELSE%> <td align="Right"><strong><%=FormatNumber(nBUAHBUSUK, 2)%></strong></td><%END IF%>
			<%IF nQE_1 = 0 THEN%><td align="Right"><strong></strong></td><%ELSE%> <td align="Right"><strong><%=FormatNumber(nQE_1, 2)%></strong></td><%END IF%>
			<%IF nQE_2 = 0 THEN%><td align="Right"><strong></strong></td><%ELSE%> <td align="Right"><strong><%=FormatNumber(nQE_2, 2)%></strong></td><%END IF%>
			<%IF nQE_3 = 0 THEN%><td align="Right"><strong></strong></td><%ELSE%> <td align="Right"><strong><%=FormatNumber(nQE_3, 2)%></strong></td><%END IF%>
			<%IF nQE_4 = 0 THEN%><td align="Right"><strong></strong></td><%ELSE%> <td align="Right"><strong><%=FormatNumber(nQE_4, 2)%></strong></td><%END IF%>
			<%IF nQE_5 = 0 THEN%><td align="Right"><strong></strong></td><%ELSE%> <td align="Right"><strong><%=FormatNumber(nQE_5, 2)%></strong></td><%END IF%>
			<%IF nQE_6 = 0 THEN%><td align="Right"><strong></strong></td><%ELSE%> <td align="Right"><strong><%=FormatNumber(nQE_6, 2)%></strong></td><%END IF%>
			<%IF nQE_7 = 0 THEN%><td align="Right"><strong></strong></td><%ELSE%> <td align="Right"><strong><%=FormatNumber(nQE_7, 2)%></strong></td><%END IF%>

		</tr>
   <%
   else
      response.Write("Data is empty")	  
   end if 
   rsCat.Close
   set rsCat = Nothing%>
   </tbody>
   </TABLE>
<!--#include file="Bottom_Report.asp"-->
</div>
</body>
</html>