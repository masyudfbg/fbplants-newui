<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>Reporting - Laporan Harian Divisi</title>
</head>
<%
Dim v_from,v_to,v_company,v_fcba,v_divisi,filter,filterdiv,filterlhd
v_from       = trim(REQUEST("bkmfrom"))
v_to         = trim(REQUEST("bkmto"))
v_company    = trim(REQUEST("pcompany"))
v_fcba       = trim(REQUEST("pfcba"))
v_divisi 	 = trim(REQUEST("divisi"))	
%>

<script language="JavaScript" type="text/javascript"">
window.onload = function(){  
	window.print();
}  
 
 </script>

<%

	if trim(v_divisi) <> "all" then
		filterdiv = " AND h.sectioncode='" & v_divisi & "'"
	end if


	if trim(v_fcba) = "-" then 
		filterlhd = " AND h.FCBA in (SELECT fccode from businessunit where fccompanycode='" & v_company & "') "
	else
	    filterlhd = " AND h.FCBA ='" & v_fcba & "' "
	end if

	q = " SELECT h.fcba, d.actcode || ' - ' || j.fcname jobname,h.SECTIONCODE,d.LOCATIONCODE,f.plantingdate tahuntanam,f.luas_gross,f.status, " & _
	" nvl(SUM(d.MANDAYS),0) HK,nvl(SUM(d.OUTPUT),0) FISIK,nvl(SUM(d.UNIT),0) UNIT " & _
	" FROM bkmhdr h " & _
	" LEFT JOIN bkmupahdtl d on h.BKMNO = d.BKMNO and h.fcba=d.fcba " & _
	" LEFT JOIN JOB j on j.fccode=d.ACTCODE and j.FCBA = d.FCBA " & _
	" INNER JOIN FIELD f on f.fccode=d.LOCATIONCODE and f.fcba=d.fcba " & _
	" WHERE h.bkmdate between to_date('" & v_from & "','dd/mm/yyyy') AND to_date('" & v_to & "','dd/mm/yyyy') " & _
		filterdiv & _ 
		filterlhd & _ 
	" GROUP BY h.fcba, d.actcode || ' - ' || j.fcname,h.SECTIONCODE,d.LOCATIONCODE,f.plantingdate,f.luas_gross,f.status " & _
	" order by  d.actcode || ' - ' || j.fcname asc"
		
		' response.write q
Set rsCat = Server.CreateObject("ADODB.Recordset")
			rsCat.CursorType = 3
			rsCat.CursorLocation = 3
			rsCat.Open q, connstr

			'set rsview = connStr.Execute(sqlview)
%>
<div id="wrapper">
<body>
  
   	<table border="0" width="100%" id="product-table" >
			<thead>
				<tr>
					<th align="left" colspan="23" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;">LHD LISTING  <%=v_company%> Unit : <%=v_fcba%> Periode : <%=v_from%> / <%=v_to%></th>		
				</tr>
				
		 <th style="color:#fff;">FCBA</th>
		 <th style="color:#fff;">Activity</th> 
		 <th style="color:#fff;">Section</th> 
		 <th style="color:#fff;">Blok</th> 
		 <th style="color:#fff;">TT</th> 
		 <th style="color:#fff;">Luas</th> 
		 <th style="color:#fff;">Status</th> 
		 <th style="color:#fff;">HK</th> 
		 <th style="color:#fff;">Ha</th> 
		 <th style="color:#fff;">Qty</th> 		
	</tr>
		</thead>
<tbody>
	<%
		dim ttlhk,ttlha,ttlqty 
		ttlhk = 0
		ttlha = 0
		ttlqty = 0
		While Not rsCat.EOF
	%>
		<tr>			
			<td align="Left" ><%=rscat("FCBA")%></td>			
			<td align="Left" ><%=rscat("jobname")%></td>			
			<td align="Left" ><%=rscat("sectioncode")%></td>				
			<td align="Left" ><%=rscat("locationcode")%></td>				
			<td align="left" ><%=rscat("tahuntanam")%></td>		
			<td align="left" ><%=rscat("luas_gross")%></td>		
			<td align="left" ><%=rscat("status")%></td>
			<td data-t="n" align="Right" ><%=rscat("hk")%></td>				
			<td data-t="n" align="Right" ><%=rscat("fisik")%></td>				
			<td data-t="n" align="Right" ><%=rscat("unit")%></td>				
			
		</tr>
	  <%
	  	ttlhk = ttlhk + cdbl(rscat("hk"))
		ttlha = ttlha + cdbl(rscat("fisik"))
		ttlqty = ttlqty + cdbl(rscat("unit"))
	  rsCat.MoveNext
	Wend
	rsCat.close()
	rsCat=null
	
%>
	<tr>
		<td colspan="7" align="right">Total</td>
		<td align="right"><%=formatnumber(ttlhk)%></td>
		<td align="right"><%=formatnumber(ttlha)%></td>
		<td align="right"><%=formatnumber(ttlqty)%></td>
	</tr>
</tbody>

</table>
</div>

</BODY>
</div>
</HTML>
