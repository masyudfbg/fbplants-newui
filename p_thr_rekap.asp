<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>ePlaS Reporting - Tunjangan Hari Raya</title>
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
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>
</head>
<%

dim tglawal,tglakhir,company,fcba,jenisthr,religion,jenisreport,tglhariraya
Dim txtxbln,txttgl,txtthn


Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

tglawal		=trim(request("tglawal"))
tglakhir	=trim(request("tglakhir"))
company		=trim(request("company"))
fcba		=trim(request("fcba"))
jenisthr	=trim(request("jenisthr"))
religion	=trim(request("religion"))
jenisreport	=trim(request("jenisreport"))
tglhariraya	=trim(request("tglhariraya"))

businessunit= trim(request("fcba"))

If religion <> "" Then
    ' Pecah string menjadi array berdasarkan koma
    religionArray = Split(religion, ",")
    
    ' Ubah setiap elemen menjadi format 'value'
    Dim i
    For i = 0 To UBound(religionArray)
        religionArray(i) = "'" & Trim(religionArray(i)) & "'"
    Next

    ' Gabungkan kembali menjadi string SQL
    formattedReligion = Join(religionArray, ", ")

Else
   Response.Write "Error: Parameter 'religion' tidak ditemukan!"
End If

	
Dim startDate, endDate, startMonth, startYear, endMonth, endYear,formatAwal,formatAkhir
Dim result1,result2,result3,result4,result5,result6,result7,result8,result9,result10,result11,result12,result13
Dim resultt1,resultt2,resultt3,resultt4,resultt5,resultt6,resultt7,resultt8,resultt9,resultt10,resultt11,resultt12,resultt13
Dim filtercomp,filtertipethr
Dim hsl_bln1,hsl_bln2,hsl_bln3,hsl_bln4,hsl_bln5,hsl_bln6,hsl_bln7,hsl_bln8,hsl_bln9,hsl_bln10,hsl_bln11,hsl_bln12,hsl_bln13,hsl_totalhk,hsl_ttlamount

Dim hsl_status,hsl_presentrate,hsl_tglawal,hsl_tglakhir,hsl_selisihbln,tgl_gabung


' Tanggal awal dan akhir
startDate = CDate(tglawal)
endDate = CDate(tglhariraya)

' Ambil bulan dan tahun awal dan akhir
startMonth = Month(startDate)
startYear = Year(startDate)
endMonth = Month(endDate)
endYear = Year(endDate)



' Hasil awal kosong
result = ""

' Buat daftar bulan secara manual berdasarkan rentang
result1 = result1 & textMonth(startMonth) & " " & startYear & "<br>"
resultt1 = resultt1 & (startMonth) & "" & startYear
IF  startMonth = 12 THEN
	startMonth = 1
	startYear  = startYear + 1
ELSE
	startMonth = startMonth + 1
	startYear = startYear 
END IF


result2 = result2 & textMonth(startMonth) & " " & startYear & "<br>"
resultt2 = resultt2 & (startMonth) & "" & startYear
IF  startMonth = 12 THEN
	startMonth = 1
	startYear  = startYear + 1
ELSE
	startMonth = startMonth + 1
	startYear = startYear 
END IF
result3 = result3 & textMonth(startMonth) & " " & startYear & "<br>"
resultt3 = resultt3 & (startMonth) & "" & startYear
IF  startMonth = 12 THEN
	startMonth = 1
	startYear  = startYear + 1
ELSE
	startMonth = startMonth + 1
	startYear = startYear 
END IF
result4 = result4 & textMonth(startMonth) & " " & startYear & "<br>"
resultt4 = resultt4 & (startMonth) & "" & startYear
IF  startMonth = 12 THEN
	startMonth = 1
	startYear  = startYear + 1
ELSE
	startMonth = startMonth + 1
	startYear = startYear 
END IF
result5 = result5 & textMonth(startMonth) & " " & startYear & "<br>"
resultt5 = resultt5 & (startMonth) & "" & startYear
IF  startMonth = 12 THEN
	startMonth = 1
	startYear  = startYear + 1
ELSE
	startMonth = startMonth + 1
	startYear = startYear 
END IF
result6 = result6 & textMonth(startMonth) & " " & startYear & "<br>"
resultt6 = resultt6 & (startMonth) & "" & startYear
IF  startMonth = 12 THEN
	startMonth = 1
	startYear  = startYear + 1
ELSE
	startMonth = startMonth + 1
	startYear = startYear 
END IF
result7 = result7 & textMonth(startMonth) & " " & startYear & "<br>"
resultt7 = resultt7 & (startMonth) & "" & startYear
IF  startMonth = 12 THEN
	startMonth = 1
	startYear  = startYear + 1
ELSE
	startMonth = startMonth + 1
	startYear = startYear 
END IF
result8 = result8 & textMonth(startMonth) & " " & startYear & "<br>"
resultt8 = resultt8 & (startMonth) & "" & startYear
IF  startMonth = 12 THEN
	startMonth = 1
	startYear  = startYear + 1
ELSE
	startMonth = startMonth + 1
	startYear = startYear 
END IF
result9 = result9 & textMonth(startMonth) & " " & startYear & "<br>"
resultt9 = resultt9 & (startMonth) & "" & startYear
IF  startMonth = 12 THEN
	startMonth = 1
	startYear  = startYear + 1
ELSE
	startMonth = startMonth + 1
	startYear = startYear 
END IF
result10 = result10 & textMonth(startMonth) & " " & startYear & "<br>"
resultt10 = resultt10 & (startMonth) & "" & startYear
IF  startMonth = 12 THEN
	startMonth = 1
	startYear  = startYear + 1
ELSE
	startMonth = startMonth + 1
	startYear = startYear 
END IF
result11 = result11 & textMonth(startMonth) & " " & startYear & "<br>"
resultt11 = resultt11 & (startMonth) & "" & startYear
IF  startMonth = 12 THEN
	startMonth = 1
	startYear  = startYear + 1
ELSE
	startMonth = startMonth + 1
	startYear = startYear 
END IF
result12 = result12 & textMonth(startMonth) & " " & startYear & "<br>"
resultt12 = resultt12 & (startMonth) & "" & startYear
IF  startMonth = 12 THEN
	startMonth = 1
	startYear  = startYear + 1
ELSE
	startMonth = startMonth + 1
	startYear = startYear 
END IF
result13 = result13 & textMonth(startMonth) & " " & startYear & "<br>"
resultt13 = resultt13 & (startMonth) & "" & startYear
' Tambahkan logika untuk setiap bulan lainnya secara manual...




q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & businessunit & "' "
	' response.write q
set rsCat = connStr.Execute(q)
if not rsCat.eof then
   fcbaname = rsCat("fcName")
   CompanyName=rsCat("FCCOMPANYNAME")
   MgrName = rsCat("FCMANAGER")
   KTUName = rsCat("FCKTU")
   CropName=rsCat("FCCROP")
   FcbaType=rsCat("fcType")
end if

if trim(fcba) = "-" then 
	filtercomp = " AND bb.isactive='TRUE' AND bb.fccompanycode ='" & company & "' "
else 
	filtercomp = " AND bb.isactive='TRUE' AND bb.fccode ='" & fcba & "' "
end if

if trim(jenisthr) = "1" then 
	filtertipethr = " AND c.RELIGION in(" & formattedReligion & ") "
else 
	filtertipethr = " AND c.RELIGION in(" & formattedReligion & ") "
end if

txtSQL =  "SELECT fccompanycode, fcba,fccode,fcname,to_char(dateofjoin,'dd/mm/yyyy')dateofjoin,to_char(dateterminate,'dd/mm/yyyy')dateterminate ,presentgrade,emp_type,presentrate, " & _
" SUM(bulan_1) bulan_1," & _
" SUM(bulan_2) bulan_2," & _
" SUM(bulan_3) bulan_3," & _
" SUM(bulan_4) bulan_4," & _
" SUM(bulan_5) bulan_5," & _
" SUM(bulan_6) bulan_6," & _
" SUM(bulan_7) bulan_7," & _
" SUM(bulan_8) bulan_8," & _
" SUM(bulan_9) bulan_9," & _
" SUM(bulan_10) bulan_10," & _
" SUM(bulan_11) bulan_11," & _
" SUM(bulan_12) bulan_12," & _
" SUM(bulan_13) bulan_13," & _
" SUM(bulan_1) + " & _
" SUM(bulan_2) + " & _
" SUM(bulan_3) + " & _
" SUM(bulan_4) + " & _
" SUM(bulan_5) + " & _
" SUM(bulan_6) + " & _
" SUM(bulan_7) + " & _
" SUM(bulan_8) + " & _
" SUM(bulan_9) + " & _
" SUM(bulan_10) + " & _
" SUM(bulan_11) + " & _
" SUM(bulan_12) + " & _
" SUM(bulan_13) TotHariKerja " & _
" FROM " & _
" ( " & _
" SELECT fccompanycode,fcba,fccode,fcname,dateofjoin,dateterminate ,presentgrade,emp_type,presentrate, " & _
" SUM(DECODE(TO_NUMBER(TO_CHAR(bkmdate, 'mmyyyy')), '" & resultt1 & "', Mandays, 0)) bulan_1, " & _
" SUM(DECODE(TO_NUMBER(TO_CHAR(bkmdate, 'mmyyyy')), '" & resultt2 & "', Mandays, 0)) bulan_2, " & _
" SUM(DECODE(TO_NUMBER(TO_CHAR(bkmdate, 'mmyyyy')), '" & resultt3 & "', Mandays, 0)) bulan_3, " & _
" SUM(DECODE(TO_NUMBER(TO_CHAR(bkmdate, 'mmyyyy')), '" & resultt4 & "', Mandays, 0)) bulan_4, " & _
" SUM(DECODE(TO_NUMBER(TO_CHAR(bkmdate, 'mmyyyy')), '" & resultt5 & "', Mandays, 0)) bulan_5, " & _
" SUM(DECODE(TO_NUMBER(TO_CHAR(bkmdate, 'mmyyyy')), '" & resultt6 & "', Mandays, 0)) bulan_6, " & _
" SUM(DECODE(TO_NUMBER(TO_CHAR(bkmdate, 'mmyyyy')), '" & resultt7 & "', Mandays, 0)) bulan_7, " & _
" SUM(DECODE(TO_NUMBER(TO_CHAR(bkmdate, 'mmyyyy')), '" & resultt8 & "', Mandays, 0)) bulan_8, " & _
" SUM(DECODE(TO_NUMBER(TO_CHAR(bkmdate, 'mmyyyy')), '" & resultt9 & "', Mandays, 0)) bulan_9, " & _
" SUM(DECODE(TO_NUMBER(TO_CHAR(bkmdate, 'mmyyyy')), '" & resultt10 & "', Mandays, 0)) bulan_10, " & _
" SUM(DECODE(TO_NUMBER(TO_CHAR(bkmdate, 'mmyyyy')), '" & resultt11 & "', Mandays, 0)) bulan_11, " & _
" SUM(DECODE(TO_NUMBER(TO_CHAR(bkmdate, 'mmyyyy')), '" & resultt12 & "', Mandays, 0)) bulan_12, " & _
" SUM(DECODE(TO_NUMBER(TO_CHAR(bkmdate, 'mmyyyy')), '" & resultt13 & "', Mandays, 0)) bulan_13 " & _
" FROM " & _
"    ( " & _
"    SELECT bb.fccompanycode,c.fcba,c.fccode,c.fcname,c.dateofjoin,c.dateterminate ,c.presentgrade,c.emp_type,bkmdate,c.presentrate, " & _
"    count(DISTINCT bkmdate) Mandays " & _
"    FROM bkmupahdtl a " & _
"    INNER JOIN bkmhdr b   ON (a.fcba=b.fcba AND a.bkmno=b.bkmno) " & _
"    INNER JOIN employee c ON (a.employeecode = c.fccode AND a.fcba=c.fcba) " & _
"	 INNER JOIN businessunit bb ON (bb.fccode=c.fcba) " & _
"    WHERE " & _
"    a.attendance in('KJ','KK','DL','WS','WH','KB','TG') AND bkmdate BETWEEN to_date('" & tglawal & "','mm/dd/yyyy') AND to_date('" & tglakhir & "','mm/dd/yyyy')  " & _
"	AND (c.dateterminate IS NULL OR c.dateterminate >= to_date('" & tglhariraya & "','mm/dd/yyyy') - 31) " & _
filtercomp  & _
filtertipethr  & _
"    GROUP BY bb.fccompanycode,c.fcba,c.fccode,c.fcname,c.dateofjoin,c.dateterminate,c.presentgrade,c.emp_type,bkmdate,c.presentrate " & _
"    ) " & _
" GROUP BY fccompanycode,fcba,fccode,fcname,dateofjoin,dateterminate,presentgrade,emp_type,presentrate " & _
" ) " & _ 
" WHERE 1=1 " & _
" GROUP BY  fccompanycode,fcba,fccode,fcname,dateofjoin,dateterminate ,presentgrade,emp_type,presentrate " & _
" ORDER BY fcba, CASE presentgrade    WHEN 'KBT' THEN 1    WHEN 'KHT' THEN 2    WHEN 'KHL' THEN 3 END,fcname"  

 'response.write(txtSQL)

txtSelect = txtSelect & " From " & FormatDateTime(tglawal,1) & " To " &  FormatDateTime(tglakhir,1)

if len(trim(txtSelect)) > 0 then
   txtSelect = "Selected  by : " & txtSelect
else
   txtSelect = "Selected by All Data"
end if

'response.Write(txtsql)
response.Flush()


%>
	<body> 
		<div id="wrapper">
        <!--#include file="print_excel.asp"-->
        <table border="0" width="100%" id="product-table" class="xthf-site">
         <thead>
         <tr>
         <th colspan="24" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=CompanyName%></h4>
					<h5><%=fcbaname%></h5>
				</div>
				<div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h2>REKAP TUNJANGAN HARI RAYA (THR)</h2>
				</div>
			</div>
				</th>
			</tr>
						<!--<div class="printhide">-->
						<tr>
							<th colspan="24" style=" background-color:#FFFFFF; color:#000000; border:none"><%=txtSelect%></td>
						</tr>
						<!--</div>-->
							<tr>
									<%	
									
									Set rsCat = Server.CreateObject("ADODB.Recordset")
									rsCat.CursorType = 3
									rsCat.CursorLocation = 3
									rsCat.Open txtSQL, connstr

									
									if Not rsCat.eof then
									   nCountofField=rsCat.fields.count-1
									end if
									NO=1%>
							</tr>

							<tr>
								<th rowspan="2">Company</th>
								<th rowspan="2">Unit</th>
								<th rowspan="2">Employee Code</th>
								<th rowspan="2">Employee Name</th>
								<th rowspan="2">Date Of Join</th>
								<th rowspan="2">Date Terminate</th>
								<th rowspan="2">Present Grade</th>
								<th rowspan="2">Employe Type</th>
								<th rowspan="2">Presen Rate</th>
								<th rowspan="2">Masa Kerja</th>
								<th rowspan="2">Amount</th>
							</tr>
						
                            </thead>
                            <tbody>
								<%
								Dim v_bln1,v_bln2,v_bln3,v_bln4,v_bln5,v_bln6,v_bln7,v_bln8,v_bln9,v_bln10,v_bln11,v_bln12,v_bln13,v_totalhk,v_ttlamount
								DIM tahun,bulan,hslmasakerja
								v_bln1		= 0
								v_bln2 		= 0
								v_bln3		= 0
								v_bln4 		= 0
								v_bln5		= 0
								v_bln6 		= 0
								v_bln7		= 0
								v_bln8 		= 0
								v_bln9		= 0
								v_bln10 	= 0
								v_bln11		= 0
								v_bln12 	= 0
								v_bln13 	= 0
								v_totalhk 	= 0
								v_ttlamount 	= 0
								
								While Not rsCat.EOF

									hsl_bln1 = rscat("bulan_1")
									hsl_bln2 = rscat("bulan_2")
									hsl_bln3 = rscat("bulan_3")
									hsl_bln4 = rscat("bulan_4")
									hsl_bln5 = rscat("bulan_5")
									hsl_bln6 = rscat("bulan_6")
									hsl_bln7 = rscat("bulan_7")
									hsl_bln8 = rscat("bulan_8")
									hsl_bln9 = rscat("bulan_9")
									hsl_bln10= rscat("bulan_10")
									hsl_bln11=rscat("bulan_11")
									hsl_bln12=rscat("bulan_12")
									hsl_bln13=rscat("bulan_13")
									hsl_status	= rscat("presentgrade")
									hsl_presentrate=rscat("presentrate")
									tgl_gabung	= rscat("dateofjoin")


									if hsl_status="KHL" AND (hsl_bln1) > "20" THEN
										hsl_bln1 = "20"
									ELSE 
										hsl_bln1 = hsl_bln1
									END IF
									if hsl_status="KHL" AND (hsl_bln2) > "20" THEN
										hsl_bln2 = "20"
									ELSE 
										hsl_bln2 = hsl_bln2
									END IF
									if hsl_status="KHL" AND (hsl_bln3) > "20" THEN
										hsl_bln3 = "20"
									ELSE 
										hsl_bln3 = hsl_bln3
									END IF
									if hsl_status="KHL" AND (hsl_bln4) > "20" THEN
										hsl_bln4 = "20"
									ELSE 
										hsl_bln4 = hsl_bln4
									END IF
									if hsl_status="KHL" AND (hsl_bln5) > "20" THEN
										hsl_bln5 = "20"
									ELSE 
										hsl_bln5 = hsl_bln5
									END IF
									if hsl_status="KHL" AND (hsl_bln6) > "20" THEN
										hsl_bln6 = "20"
									ELSE 
										hsl_bln6 = hsl_bln6
									END IF
									if hsl_status="KHL" AND (hsl_bln7) > "20" THEN
										hsl_bln7 = "20"
									ELSE 
										hsl_bln7 = hsl_bln7
									END IF
									if hsl_status="KHL" AND (hsl_bln8) > "20" THEN
										hsl_bln8 = "20"
									ELSE 
										hsl_bln8 = hsl_bln8
									END IF
									if hsl_status="KHL" AND (hsl_bln9) > "20" THEN
										hsl_bln9 = "20"
									ELSE 
										hsl_bln9 = hsl_bln9
									END IF
									if hsl_status="KHL" AND (hsl_bln10) > "20" THEN
										hsl_bln10 = "20"
									ELSE 
										hsl_bln10 = hsl_bln10
									END IF
									if hsl_status="KHL" AND (hsl_bln11) > "20" THEN
										hsl_bln11 = "20"
									ELSE 
										hsl_bln11 = hsl_bln11
									END IF
									if hsl_status="KHL" AND (hsl_bln12) > "20" THEN
										hsl_bln12 = "20"
									ELSE 
										hsl_bln12 = hsl_bln12
									END IF
									if hsl_status="KHL" AND (hsl_bln13) > "20" THEN
										hsl_bln13 = "20"
									ELSE 
										hsl_bln13 = hsl_bln13
									END IF

									hsl_selisihbln	= DateDiff("m", tgl_gabung, endDate)
							'		Response.write hsl_selisihbln

									hsl_totalhk = cdbl(hsl_bln1)+cdbl(hsl_bln2)+cdbl(hsl_bln3)+cdbl(hsl_bln4)+cdbl(hsl_bln5)+cdbl(hsl_bln6)+cdbl(hsl_bln7)+cdbl(hsl_bln8)+cdbl(hsl_bln9)+cdbl(hsl_bln10)+cdbl(hsl_bln11)+cdbl(hsl_bln12)+cdbl(hsl_bln13)
									
							
									If hsl_status="KHT" AND cdbl(hsl_selisihbln) >= 12 Then
										hsl_ttlamount = hsl_presentrate 
									Elseif hsl_status="KHT" AND cdbl(hsl_selisihbln) < 12 Then
										hsl_ttlamount = cdbl(hsl_selisihbln/12)*cdbl(hsl_presentrate) 
									elseIf hsl_status="KBT" AND (hsl_selisihbln) >= 12 Then
										hsl_ttlamount = hsl_presentrate 
										
									Elseif hsl_status="KBT" AND cdbl(hsl_selisihbln) < 12 Then
										hsl_ttlamount = cdbl(hsl_selisihbln/12)*cdbl(hsl_presentrate) 
									elseif hsl_status="KHL" then
										hsl_ttlamount = (cdbl(hsl_totalhk)/240)*(cdbl(hsl_presentrate)*25)	
									End If

									
									tahun = Int(hsl_selisihbln) / 12
									bulan = hsl_selisihbln Mod 12 ' hitung sisa bulan
									tahun = Int(tahun)
								'	bulan = Int(bulan)
									hslmasakerja = tahun & " Tahun " & bulan & " Bulan "
									

									
								%>


									<tr>
									</td>
										<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fccompanycode")%> </td>
										<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fcba")%> </td>
										<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fccode")%> </td>
										<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fcname")%> </td>
										<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=(rscat("dateofjoin"))%></td>
										<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("dateterminate")%></td>
										<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("presentgrade")%></td>
										<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("emp_type")%></td>
										<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(rscat("presentrate"),0)%></td>
										<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=hslmasakerja%></td>
										<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(hsl_ttlamount,0)%></td>
							
									</tr>
								  <%
									
									v_totalhk 	= cdbl(v_totalhk) + cdbl(hsl_totalhk)
									v_ttlamount = cdbl(v_ttlamount) + cdbl(hsl_ttlamount)
								  rsCat.MoveNext
								  'response.Flush()
								  no=no+1
								Wend
								rsCat.close()
								rsCat=null%>
								<input  type="hidden" value="<%=txtTable%>" name="txtTable">
								<input  type="hidden" value="<%=txtModule%>" name="txtModule">
								</form>
								</tbody>
								<tr bgcolor="#F4F4A0" style="font-weight: bold;">
									<td colspan="9" align="center">TOTAL</td>
									<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(v_totalhk)),0)%></td>
									<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(v_ttlamount)),0)%></td>
							
								</tr>
						</table>
						<!--#include file="Bottom_Report.asp"-->
		</DIV>
	</body>	
</HTML>