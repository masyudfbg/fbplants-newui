<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>Reporting - Account Report</title>
</head>
<%
Dim region,company,fcba,actfrom,actto,bulan,tahun,costelemen,sectioncode,t1,r1,statustm
dim filterdiv,filtersource,filterfcba,filterprd,filterprdlalu,filtercostelemen,filterprdsbi,filtertmtbm
dim prev_bulanfrom,prev_bulanto
dim bulansplitfrom,tglsplitfrom,thnsplitfrom,tglfrom,tglto,tglawalsbi
dim bulansplitto,tglsplitto,thnsplitto,xbln1,xthn1
dim rkbhablnlalu
region 		= Request("region")
company 	= Request("company")
fcba 		= Request("fcba")
actfrom   	= Request("actfrom")
actto   	= Request("actto")
bulan   	= Request("bulan")
tahun   	= Request("tahun")
costelemen  = Request("costelemen")
sectioncode	= Request("section")
t1   		= Request("t1")
r1   		= Request("r1")
statustm	= Request("statustm")
if t1 = "Tgl" then 
	periodeawal = actfrom 
	periodeakhir= actto
else 
	periodeawal = bulan 
	periodeakhir= tahun
end if
%>

<script src="/distexcel/tableToExcel.js"></script>
<script type="text/javascript">

	function getExport()
	{
		var region 		= '<%=region%>';
		var company 	= '<%=company%>';
		var fcba 		= '<%=fcba%>';
		var actfrom   	= '<%=actfrom%>';
		var actto   	= '<%=actto%>';
		var bulan   	= '<%=bulan%>';
		var tahun   	= '<%=tahun%>';
		var costelemen  = '<%=costelemen%>';
		var sectioncode	= '<%=section%>';
		var t1   		= '<%=t1%>';
		var r1   		= '<%=r1%>';
		var statustm  	= '<%=statustm%>';
		if(t1=="Tgl"){
			var nm = 'ACTREPORTDETAILS_' + region + '_' + fcba + '_'+actfrom+actto+".xlsx";  
		}else{
			var nm = 'ACTREPORTDETAILS_' + region + '_' + fcba + '_'+bulan+tahun+".xlsx";  
		}
	
		TableToExcel.convert(document.getElementById("product-table"), {
            name: nm,
            sheet: {
            name: "ACTREPORTDETAILS"
            }
          });
		
	}
</script>	
<script language="JavaScript" type="text/javascript" >
	document.getElementById('tableview').style.display = "none";
</script>
<%


	'rubah format tgl awal jd dd/mm/yyyy
	bulansplitfrom  = left(actfrom,2)
    tglsplitfrom    = mid(actfrom,4,2)
    thnsplitfrom    = right(actfrom,4)
    tglfrom     = tglsplitfrom & "/" & bulansplitfrom & "/" & thnsplitfrom

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


    sqlview =  "select t.region,t.fccompanycode,t.source,t.fccode,j.job_forfieldstatus fieldstatus,t.fccode jobcode,t.jobname, " & _
	" t.sectioncode, t.fieldcode,t.fieldarea,t.fieldpokok, " & _
	" round(NVL(sum(t.hkblnlalu),0),0)hkblnlalu,round(DECODE(NVL(sum(t.fisikblnlalu),0),0,0),2)fisikblnlalu,round(NVL(sum(t.unitblnlalu),0),0)unitblnlalu,round(NVL(sum(t.amountblnlalu),0),0)amountblnlalu, " & _
	" round(sum(t.hkblnini),0)hkblnini,round(nvl(sum(t.fisikblnini),0),2)fisikblnini,round(sum(t.unitblnini),0)unitblnini,round(sum(t.amountblnini),0)amountblnini, " & _
	" round(sum(t.hksbi),0)hksbi,round(decode(sum(t.fisiksbi),0,0),2)fisiksbi,round(sum(t.unitsbi),0)unitsbi,round(sum(t.amountsbi),0)amountsbi, " & _
	" round(NVL(SUM(t.rkb_hkblnlalu),0),0)rkb_hkblnlalu,round(NVL(SUM(t.rkb_outputblnlalu),0),0)rkb_outputblnlalu,round(NVL(SUM(t.rkb_amountblnlalu),0),0)rkb_amountblnlalu, " & _
	" round(NVL(SUM(t.rkb_hkblnini),0),0)rkb_hkblnini,round(NVL(SUM(t.rkb_outputblnini),0),0)rkb_outputblnini,round(NVL(SUM(t.rkb_amountblnini),0),0)rkb_amountblnini, " & _
	" round(NVL(SUM(t.rkb_hksbi),0),0)rkb_hksbi,round(NVL(SUM(t.rkb_outputsbi),0),0)rkb_outputsbi,round(NVL(SUM(t.rkb_amountsbi),0),0)rkb_amountsbi, " & _
	" round(DECODE(NVL(SUM(t.rkb_outputblnlalu),0),0,0,(SUM(t.rkb_outputblnlalu)) / round(NVL(SUM(t.rkb_amountblnlalu),0),0)),0)rkb_rphablnlalu, " & _
	" round(DECODE(NVL(SUM(t.rkb_outputsbi),0),0,0,(SUM(t.rkb_outputsbi)) / round(NVL(SUM(t.rkb_amountsbi),0),0)),0)rkb_rphasbi " & _
	" from( " & _
			" SELECT b.region,b.fccompanycode,b.fccode,o.fccode jobcode, o.fccode || ' - ' || j.fcname jobname, " & _
			" o.source,o.sectioncode,o.fieldcode,o.fieldarea,o.fieldpokok, " & _
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
			" group by  b.region,b.fccompanycode,b.fccode,o.fccode, o.fccode || ' - ' || j.fcname, " & _
			"  o.source,o.sectioncode,o.fieldcode,o.fieldarea,o.fieldpokok " & _
			"  union all " & _
			" SELECT b.region,b.fccompanycode,b.fccode,o.fccode jobcode, o.fccode || ' - ' || j.fcname jobname, " & _
			"  o.source,o.sectioncode,o.fieldcode,o.fieldarea,o.fieldpokok, " & _
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
			" group by  b.region,b.fccompanycode,b.fccode,o.fccode, o.fccode || ' - ' || j.fcname, " & _
			"  o.source,o.sectioncode,o.fieldcode,o.fieldarea,o.fieldpokok " & _
			"  union all " & _
			" SELECT b.region,b.fccompanycode,b.fccode,o.fccode jobcode, o.fccode || ' - ' || j.fcname jobname, " & _
			"  o.source,o.sectioncode,o.fieldcode,o.fieldarea,o.fieldpokok, " & _
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
			" group by  b.region,b.fccompanycode,b.fccode,o.fccode, o.fccode || ' - ' || j.fcname, " & _
			"  o.source,o.sectioncode,o.fieldcode,o.fieldarea,o.fieldpokok " & _
	" )t LEFT JOIN jobmaster j ON t.jobcode=j.fccode" & _
	filtercompany & _
	filtercostelemen & _
	filterdiv & _
	filtertmtbm & _

	" group by t.region,t.fccompanycode,t.fccode,j.job_forfieldstatus,t.jobname,t.source,t.sectioncode,t.fieldcode,t.fieldarea,t.fieldpokok " & _
	" order by t.region,t.fccompanycode,t.fccode,t.jobname"
	
	'	 response.write sqlview
Set rsCat = Server.CreateObject("ADODB.Recordset")
			rsCat.CursorType = 3
			rsCat.CursorLocation = 3
			rsCat.Open sqlview, connstr
%>

<body>
    <span id="loadingjudul1">Proses export data...</span>
    <div id="tableview" style="display:none;">
   	<table border="0" width="100%" id="product-table" >
			<thead>
				<tr>
					<th align="left" colspan="23" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;">ACT REPORT LISTING  <%=regional%> PT. : <%=company%> Periode : <%=periodeawal%> / <%=periodeakhir%></th>		
				</tr>
				<tr>
						<th style="color:#fff;">Region</th>
						<th style="color:#fff;">Company</th>
						<th style="color:#fff;">FCBA</th>
						<th style="color:#fff;">Pekerjaan</th>
						<th style="color:#fff;">Source</th>
						<th style="color:#fff;">Afdeling</th>
						<th style="color:#fff;">Blok</th>
						<th style="color:#fff;">Status</th>
						<th style="color:#fff;">Luas</th>
						<th style="color:#fff;">Pokok</th>
						<th style="color:#fff;">Fisik</th>
						<th style="color:#fff;">Rupiah</th>
						<th style="color:#fff;">Rp/Ha</th>
						<th style="color:#fff;">Fisik</th>
						<th style="color:#fff;">Rupiah</th>
						<th style="color:#fff;">Rp/Ha</th>
						<th style="color:#fff;">%</th>
						<th style="color:#fff;">Fisik</th>
						<th style="color:#fff;">Rupiah</th>
						<th style="color:#fff;">Rp/Ha</th>
						<th style="color:#fff;">Fisik</th>
						<th style="color:#fff;">Rupiah</th>
						<th style="color:#fff;">Rp/Ha</th>
						<th style="color:#fff;">%</th>
						<th style="color:#fff;">Fisik</th>
						<th style="color:#fff;">Rupiah</th>
						<th style="color:#fff;">Rp/Ha</th>
						<th style="color:#fff;">Fisik</th>
						<th style="color:#fff;">Rupiah</th>
						<th style="color:#fff;">Rp/Ha</th>
						<th style="color:#fff;">%</th>
				</tr>
		</thead>
<tbody>
	<%
		While Not rsCat.EOF
			fisikblnlalu 		= rscat("fisikblnlalu")
			amountblnlalu 		= rscat("amountblnlalu")
			rkb_outputblnlalu 	= rscat("rkb_outputblnlalu")
			rkb_amountblnlalu 	= rscat("rkb_amountblnlalu")
			fisikblnini 		= rscat("fisikblnini")
			amountblnini 		= rscat("amountblnini")
			rkb_outputblnini 	= rscat("rkb_outputblnini")
			rkb_amountblnini 	= rscat("rkb_amountblnini")
			fisiksbi 			= rscat("fisiksbi")
			amountsbi 			= rscat("amountsbi")
			rkb_outputsbi 		= rscat("rkb_outputsbi")
			rkb_amountsbi 		= rscat("rkb_amountsbi")
			' if fisikblnlalu = "0" then
			' 		rphablnlalu	= 0
			' else
			' 		rphablnlalu = cdbl(amountblnlalu)/cdbl(fisikblnlalu)
			' end if
			' if fisikblnlalu = "0" then
			' 		rkbhablnlalu	= 0
			' else
			' 		rkbhablnlalu = cdbl(rkb_amountblnlalu)/cdbl(rkb_outputblnlalu)
			' end if
			if fisikblnini = "0" then
					rpbulanini	= 0
			else
					rpbulanini = cdbl(amountblnini)/cdbl(fisikblnini)
			end if
			if rkb_outputblnini = "0" then
					rkbhabulanini	= 0
			else
					rkbhabulanini = cdbl(rkb_amountblnini)/cdbl(rkb_outputblnini)
			end if
			' if fisiksbi = "0" then
			' 		rphasbi	= 0
			' else
			' 		rphasbi = cdbl(amountsbi)/cdbl(fisiksbi)
			' end if
			' if rkb_outputsbi = "0" then
			' 		rkbhasbi	= 0
			' else
			' 		rkbhasbi = cdbl(rkb_amountsbi)/cdbl(rkb_outputsbi)
			' end if
	%>
		<tr>			
			<td align="Left" ><%=rscat("region")%></td>			
			<td align="Left" ><%=rscat("fccompanycode")%></td>			
			<td align="Left" ><%=rscat("fccode")%></td>				
			<td align="left" ><%=rscat("jobname")%></td>					
			<td align="Left" ><%=rscat("source")%></td>				
			<td align="Left" ><%=rscat("sectioncode")%></td>								
			<td align="Left" ><%=rscat("fieldcode")%></td>				
			<td align="Left" ><%=rscat("fieldstatus")%></td>				
			<td align="Left" ><%=rscat("fieldarea")%></td>				
			<td align="Left" ><%=rscat("fieldpokok")%></td>				
			<td data-t="n" align="Right" ><%=FbpNum2((fisikblnlalu))%></td>		
			<td data-t="n" align="Right" ><%=FbpNum2((amountblnlalu))%></td>
			<td data-t="n" align="Right" ><%=FbpNum2((rphablnlalu))%></td>				
			<td data-t="n" align="Right" ><%=FbpNum2((rkb_outputblnlalu))%></td>				
			<td data-t="n" align="Right" ><%=FbpNum2((rkb_amountblnlalu))%></td>				
			<td data-t="n" align="Right" ><%=FbpNum2(rscat("rkb_rphablnlalu"))%></td>				
			<td data-t="n" align="Right" >0</td>				
			<td data-t="n" align="Right" ><%=FbpNum2((fisikblnini))%></td>		
			<td data-t="n" align="Right" ><%=FbpNum2((amountblnini))%></td>
			<td data-t="n" align="Right" ><%=(rpbulanini)%></td>
			<td data-t="n" align="Right" ><%=FbpNum2((rkb_outputblnini))%></td>				
			<td data-t="n" align="Right" ><%=FbpNum2((rkb_amountblnini))%></td>				
			<td data-t="n" align="Right" ><%=(rkbhabulanini)%></td>
			<td data-t="n" align="Right" >0</td>
			<td data-t="n" align="Right" ><%=FbpNum2((fisiksbi))%></td>		
			<td data-t="n" align="Right" ><%=FbpNum2((amountsbi))%></td>
			<td data-t="n" align="Right" ><%=FbpNum2((rphasbi))%></td>
			<td data-t="n" align="Right" ><%=FbpNum2((rkb_outputsbi))%></td>				
			<td data-t="n" align="Right" ><%=FbpNum2((rkb_amountsbi))%></td>				
			<td data-t="n" align="Right" ><%=FbpNum2(rscat("rkb_rphasbi"))%></td>
		</tr>
	  <%
	  rsCat.MoveNext
	Wend
	rsCat.close()
	rsCat=null
	Response.Write("<script language='javascript'> getExport();</script>")
%>
</tbody>

</table>
</div>

</BODY>
</HTML>
