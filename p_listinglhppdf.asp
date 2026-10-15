<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<script language="JavaScript" src='/js/x.js'></script>
<script language="JavaScript" src='/js/xtableheaderfixed.js'></script>
<title>Reporting - Listing LHP</title>
</head>
<%
Dim v_from,v_to,v_company,v_fcba,v_fctype,q
v_from      = trim(REQUEST("from"))
v_to      	= trim(REQUEST("to"))
v_fcba      = Session("fcBA")

%>


<script language="JavaScript">
xAddEventListener(window, 'load',
  function() {
    new xTableHeaderFixed('xthf-site', window);
    new xTableHeaderFixed('xthf-green', window);
    new xTableHeaderFixed('xthf-blue', window);
  }, false
);
 window.print();
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

	dim tglmundur,arr_field_eof,arr_field
	tglmundur = DateAdd("d",-1, v_from)
	'response.write tglmundur

	
	
arr_field_eof = true
	q = " Select LHPNO,PRODDATE, TBSTERIMA, TBSSISA,TBSOLAH, CPOKIRIM, " & _ 
       " CPOCOT, CPOCST, CPOOT,  KERNELKIRIM, KERNELSILO, CANGKANGKIRIM, CANGKANGPRODUKSI,  " & _
       " CPOTANKISI, PKTANKISI, POTONGANTBS, ADJCPO " & _
       " From   PRODUCTIONTRN  " & _
       " WHERE PRODDATE BETWEEN to_date('" & tglmundur & "','mm/dd/yyyy') and to_date('" & v_to & "','mm/dd/yyyy') and fcba='" & v_fcba & "' " & _
       " order by proddate ASC"

set rsCat = server.CreateObject("Adodb.recordset")
rsCat.Open q, connstr
if Not rsCat.Eof then
	arr_field = rsCat.getrows()
	arr_field_eof = false
End if
rsCat.Close
Set rsCat = Nothing
		
	'	 response.write q
'Set rsCat = Server.CreateObject("ADODB.Recordset")
'			rsCat.CursorType = 3
'			rsCat.CursorLocation = 3
'			rsCat.Open q, connstr
%>

<body>
    <span id="loadingjudul1">Proses export data...</span>
   
   	<table border="0" width="100%" id="product-table" >
			<thead>
				<tr>
					<th align="left" colspan="23" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;">LHP LISTING Unit : <%=v_fcba%> Periode : <%=v_from%> / <%=v_to%></th>		
				</tr>
				
		<tr>
			<th rowspan='2'>LHP No</th>
			<th rowspan='2'>Date Created</th>
			<th colspan='3' align="center">TBS</th>
			<th colspan='4' align="center">CPO</th>
			<th colspan='3' align="center">Kernel</th>
		</tr>
		<tr>
			<th>Terima</th>
			<th>Olah</th>
			<th>Sisa</th>
			<th>Produksi</th>
			<th>Kirim</th>
			<th>Adjusment</th>
			<th>Sisa</th>
			<th>Produksi</th>
			<th>Kirim</th>
			<th>Sisa</th>
		</tr>

		</thead>
<tbody>
	<%
		dim no,lhpno,tgl,tbssisahi,potongantbshi,cpokirim,adjcpo,cpocothi
		dim cpocsthi,cpoothi,cpotankisihi,totalcpohi,cposisa
		dim kernelkirimhi,kernelsilohi,pktankisihi,totalpkhi,pksisa
		dim sisakemarin,tersedia,tbsolah,cpocotkemarin,cpocstkemarin,cpootkemarin,cpotankisikemarin
		dim totalcpokemarin,cpohi,kernelsilokemarin,pktankisikemarin,totalpkkemarin,produksipkhi
		dim ttltbsterima,ttltbsolah,ttlcpoproduksi,ttlcpokirim,ttladjusment,ttlpkproduksi,ttlpkkirim
		no = 1
		ttltbsterima = 0
		ttltbsolah = 0
		ttlcpoproduksi = 0
		ttlcpokirim = 0
		ttladjusment = 0
		ttlpkproduksi = 0
		ttlpkkirim = 0
		'While Not rsCat.EOF
		if not arr_field_eof then
				for xi=0 to ubound(arr_field,2)
					
					lhpno 			= arr_field(0, xi)
					tgl 			= arr_field(1, xi)
					tbsterima 		= arr_field(2, xi)
					tbssisahi 		= arr_field(3, xi)
					cpokirim  		= arr_field(5, xi)
					cpocothi 		= arr_field(6, xi)
					cpocsthi 		= arr_field(7, xi)
					cpoothi 		= arr_field(8, xi)
					kernelkirimhi	= arr_field(9, xi)
					kernelsilohi	= arr_field(10, xi)
					cpotankisihi 	= arr_field(13, xi)
					pktankisihi		= arr_field(14, xi)
					potongantbshi 	= arr_field(15, xi)
					adjcpo  		= arr_field(16, xi)


					totalcpohi		= cdbl(cpokirim)+(cdbl(cpocothi)+cdbl(cpocsthi)+cdbl(cpoothi)+cdbl(cpotankisihi))

				
					totalpkhi		= cdbl(kernelkirimhi)+(cdbl(kernelsilohi)+cdbl(pktankisihi))

					'GET Sisa stock CPO
							
					cposisa 		= cdbl(cpotankisihi)+cdbl(cpocothi)+cdbl(cpocsthi)+cdbl(cpoothi)-cdbl(adjcpo)

					'GET Sisa stock PK
					pksisa 			= cdbl(pktankisihi)+cdbl(kernelsilohi)
								
					
					if xi > 0 then
						sisakemarin	=  arr_field(3, xi-1)
						tersedia   	= cdbl(tbsterima) + cdbl(sisakemarin)
						tbsolah    	= cdbl(tersedia) - cdbl(tbssisahi)-cdbl(potongantbshi)

						cpocotkemarin 		= arr_field(6, xi-1)
						cpocstkemarin 		= arr_field(7, xi-1)
						cpootkemarin 		= arr_field(8, xi-1)
						cpotankisikemarin 	= arr_field(13, xi-1)
						totalcpokemarin		= cdbl(cpocotkemarin)+cdbl(cpocstkemarin)+cdbl(cpootkemarin)+cdbl(cpotankisikemarin)
						if tbsolah>0 then
							cpohi				= cdbl(totalcpohi)-cdbl(totalcpokemarin)
						else
							cpohi = 0
						end if

						
						kernelsilokemarin	= arr_field(10, xi-1)
						pktankisikemarin	= arr_field(14, xi-1)
						totalpkkemarin		= cdbl(kernelsilokemarin)+cdbl(pktankisikemarin)
						produksipkhi		= cdbl(totalpkhi)-cdbl(totalpkkemarin)

						ttltbsterima = cdbl(ttltbsterima)  + cdbl(tbsterima)
						ttltbsolah = cdbl(ttltbsolah) + cdbl(tbsolah)
						ttlcpoproduksi = cdbl(ttlcpoproduksi) + cdbl(tbssisahi)
						ttlcpokirim = cdbl(ttlcpokirim) + cdbl(cpohi)
						ttladjusment = cdbl(ttladjusment) + cdbl(adjcpo)
						ttlpkproduksi = cdbl(ttlpkproduksi) + cdbl(produksipkhi)
						ttlpkkirim = cdbl(ttlpkkirim) + cdbl(kernelkirimhi)


			%>
					<tr>			
						<td align="Left" ><%=lhpno%></td>			
						<td align="Left" ><%=tgl%></td>	
						<td align="right"><%=FormatNumber(to_nol(tbsterima),2)%></td>	
						<td align="right"><%=FormatNumber(to_nol(tbsolah),2)%></td>	
						<td align="right"><%=FormatNumber(to_nol(tbssisahi),2)%></td>	
						<td align="right"><%=FormatNumber(to_nol(cpohi),2)%></td>	
						<td align="right"><%=FormatNumber(to_nol(cpokirim),2)%></td>	
						<td align="right"><%=FormatNumber(to_nol(adjcpo),2)%></td>	
						<td align="right"><%=FormatNumber(to_nol(cposisa),2)%></td>	
						<td align="right"><%=FormatNumber(to_nol(produksipkhi),2)%></td>	
						<td align="right"><%=FormatNumber(to_nol(kernelkirimhi),2)%></td>	
						<td align="right"><%=FormatNumber(to_nol(pksisa),2)%></td>	
					</tr>
					
				
				
	  <%			end if
	  			next
			
			end if

			
	 
	
%>

<tr>
						<td colspan=2 align="right"><strong>Total</strong></td>
						<td align="right"><strong><%=formatnumber(ttltbsterima,2)%></strong></td>
						<td align="right"><strong><%=formatnumber(ttltbsolah,2)%></strong></td>
						<td align="right"></td>
						<td align="right"><strong><%=formatnumber(ttlcpoproduksi,2)%></strong></td>
						<td align="right"><strong><%=formatnumber(ttlcpokirim,2)%></strong></td>
						<td align="right"><strong><%=formatnumber(ttladjusment,2)%></strong></td>
						<td align="right"></td>
						<td align="right"><strong><%=formatnumber(ttlpkproduksi,2)%></strong></td>
						<td align="right"><strong><%=formatnumber(ttlpkkirim,2)%></strong></td>
					</tr>
</tbody>
</table>

</BODY>
</HTML>
