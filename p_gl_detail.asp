<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>Reporting - Detail Listing Journal Group</title>
</head>
<%
Dim bulan,bulanto,tahun,companygl,fcbagl,coa,coaakhir,filter 
dim jrnfrom,jrnto,tglprd,posting,filterposting
dim totOpening, totDebet, totCredit, totEnding


bulan        = trim(REQUEST("bulan"))
bulanto      = trim(REQUEST("bulanto"))
tahun        = trim(REQUEST("tahun"))
companygl    = trim(REQUEST("pcompanygl"))
fcbagl       = trim(REQUEST("pfcbagl"))
coaakhir 	 = trim(REQUEST("coaakhir"))	
coa 		 = trim(REQUEST("coa"))	


%>

<script src="/distexcel/tableToExcel.js"></script>
<script type="text/javascript">

	function getExport()
	{
		var bulan 	= '<%= bulan %>';
		var tahun 	= '<%= tahun %>';
		var pcompanygl = '<%= companygl %>';
		var pfcbagl    = '<%= fcbagl %>';
		var nm = 'Journal_' + pcompanygl + '_' + pfcbagl + '_'+tahun+bulan+".xlsx";  
	//alert(nm);
	 
		TableToExcel.convert(document.getElementById("product-table"), {
            name: nm,
            sheet: {
            name: "Journal"
            }
          });
		//$('table').tblToExcel();
		
	}
</script>	
<script language="JavaScript" type="text/javascript" >
	document.getElementById('tableview').style.display = "none";
</script>
<%



    if trim(fcbagl) = "-" then 
		filter = " AND H.FCBA in (SELECT fccode from businessunit where fccompanycode='" & companygl & "' and isactive='TRUE') "
	else
	    filter = " AND H.FCBA ='" & fcbagl & "' "
	end if

    if trim(coa) <> "--ALL--" OR trim(coaakhirgl) <> "--ALL--"  then
        filter2 = " AND C.FCCODE BETWEEN '" & coa & "' AND '" & coaakhir & "'"
    end if

   
			
q = " WITH base_data AS ( " & _ 
    " SELECT '-'                                               DOCNO, " & _ 
    "       '-'                                               JOURNALNO, " & _ 
    "       Trunc(To_date('" & bulan & "' " & _ 
    "                     || '/01/' " & _ 
    "                     || '" & tahun & "', 'MM/DD/YYYY'), 'MM') - 1 JOURNALDATE, " & _ 
    "       '-'                                               SOURCECODE, " & _ 
    "       H.coa, " & _ 
    "       'SALDO AWAL BLN ' " & _ 
    "       || 4                                              COANAME, " & _ 
    "       0                                                 opening_balance, " & _ 
    "       Decode(Sign(H.opening_balance) " & _ 
    "      , 1, H.opening_balance, " & _ 
    "         0)                debet, " & _ 
    "       Decode(Sign(H.opening_balance), -1, H.opening_balance, " & _ 
    "                                       0)                credit, " & _ 
    "       H.opening_balance                                 ending_balance, " & _ 
    "       H.fcba, " & _ 
    "       0                                                 IS_LAWAN " & _ 
    " FROM   gledgerhdr H " & _ 
    "       INNER JOIN jobmaster C " & _ 
    "               ON ( H.coa = C.fccode ) " & _ 
    " WHERE  H.period_year = '" & tahun & "' " & _ 
    "       AND H.period_month= '" & bulan & "' " & _ 
  			filter & _
			filter2 & _ 
    " UNION ALL " & _ 
    " SELECT G.docno, " & _ 
    "       G.journalno, " & _ 
    "       G.journaldate, " & _ 
    "       J.sourcecode, " & _ 
    "       G.coa, " & _ 
    "       C.fcname COANAME, " & _ 
    "       H.opening_balance, " & _ 
    "       G.debet, " & _ 
    "       G.credit, " & _ 
    "       G.ending_balance, " & _ 
    "       G.fcba, " & _ 
    "       0       IS_LAWAN " & _ 
    " FROM   gledgerdtl G " & _ 
    "       INNER JOIN gledgerhdr H " & _ 
    "               ON ( G.docno = H.docno " & _ 
    "                    AND G.fcba = H.fcba ) " & _ 
    "       INNER JOIN jobmaster C " & _ 
    "               ON ( G.coa = C.fccode ) " & _ 
    "       INNER JOIN bkmjournalhdr J " & _ 
    "               ON ( G.journalno = J.docno " & _ 
    "                    AND G.fcba = J.fcba ) " & _ 
    " WHERE  H.period_year = '" & tahun & "'  " & _ 
    "      AND H.period_month BETWEEN '" & bulan & "' AND '" & bulanto & "' " & _ 
   			filter & _
			filter2 & _ 
    " UNION ALL " & _  
    " SELECT G.docno, " & _ 
    "       G.journalno, " & _ 
    "       G.journaldate, " & _ 
    "       J.sourcecode, " & _ 
    "       LWC.coa, " & _ 
    "       LWC.fcname  COANAME, " & _ 
    "      H.opening_balance, " & _ 
    "       G.debet, " & _ 
    "       G.credit, " & _ 
    "       G.ending_balance, " & _ 
    "       G.fcba, " & _ 
    "       1           IS_LAWAN " & _ 
    "FROM   gledgerdtl G " & _ 
    "       INNER JOIN gledgerhdr H " & _ 
    "               ON ( G.docno = H.docno " & _ 
    "                    AND G.fcba = H.fcba ) " & _ 
    "       INNER JOIN jobmaster C " & _ 
    "               ON ( G.coa = C.fccode ) " & _ 
    "       INNER JOIN bkmjournalhdr J " & _ 
    "               ON ( G.journalno = J.docno " & _ 
    "                    AND G.fcba = J.fcba ) " & _ 
    "       INNER JOIN ( " & _ 
    "            SELECT G2.journalno, " & _ 
    "                   G2.fcba, " & _ 
    "                   G2.coa, " & _ 
    "                   C2.fcname " & _ 
    "            FROM   gledgerdtl G2 " & _ 
    "                   LEFT JOIN jobmaster C2 " & _ 
    "                          ON ( G2.coa = C2.fccode ) " & _ 
    "            WHERE  G2.coa <> '" & coa & "'  " & _ 
    "       ) LWC " & _ 
    "               ON ( G.journalno = LWC.journalno " & _ 
    "                    AND G.fcba = LWC.fcba ) " & _ 
    " WHERE  H.period_year = '" & tahun & "' " & _ 
    "       AND H.period_month BETWEEN '" & bulan & "' AND '" & bulanto & "' " & _ 
    		filter & _
			filter2 & _ 
 " ) " & _ 
" SELECT docno, " & _ 
    "   journalno, " & _ 
    "   journaldate, " & _ 
    "   sourcecode, " & _ 
    "   coa, " & _ 
    "   coaname, " & _ 
    "   opening_balance, " & _ 
    "   debet, " & _ 
    "   credit, " & _ 
    "   ending_balance, " & _ 
    "   fcba " & _ 
" FROM   base_data " & _ 
" ORDER  BY fcba, " & _ 
    "     docno, " & _ 
    "      journaldate, " & _ 
    "      journalno, " & _ 
    "      is_lawan"  
		
	'	 response.write q  
Set rsCat = Server.CreateObject("ADODB.Recordset")
			rsCat.CursorType = 3
			rsCat.CursorLocation = 3
			rsCat.Open q, connstr
%>

<body>
    <span id="loadingjudul1">Proses export data...</span>
    <div id="tableview" style="display:none;">
   	<table border="0" width="100%" id="product-table" >
			<thead>
				<tr>
					<th align="left" colspan="23" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;">JOURNAL LISTING  <%=v_company%> Unit : <%=v_fcba%> Periode : <%=v_tahun%> / <%=v_bulan%></th>		
				</tr>
				
		 <th style="color:#fff;">FCBA</th>	
		 <th style="color:#fff;">Tanggal</th>
		 <th style="color:#fff;">Source</th>		
		 <th style="color:#fff;">No.Journal</th> 
		 <th style="color:#fff;">Job</th> 
		 <th style="color:#fff;">Job Name</th> 
		 <th style="color:#fff;">Opening</th> 
		 <th style="color:#fff;">Debit</th> 
		 <th style="color:#fff;">Kredit</th> 
		 <th style="color:#fff;">Ending</th> 		
		 
	</tr>
		</thead>
<tbody>
	<%
        totOpening = 0
        totDebet   = 0
        totCredit  = 0
        totEnding  = 0
		While Not rsCat.EOF
	%>
		<tr>	
			<td align="Left" ><%=rscat("fcba")%></td>			
			<td align="Left" ><%=rscat("JOURNALDATE")%></td>			
			<td align="Left" ><%=rscat("SOURCECODE")%></td>			
			<td align="Left" ><%=rscat("JOURNALNO")%></td>			
			<td align="Left" ><%=rscat("coa")%></td>			
			<td align="Left" ><%=rscat("COANAME")%></td>				
			<td style="text-align: right;" ><%=formatnumber(rscat("opening_balance"),0)%></td>				
			<td style="text-align: right;" ><%=formatnumber(rscat("debet"),0)%></td>				
			<td style="text-align: right;" ><%=formatnumber(rscat("credit"),0)%></td>	
			<td style="text-align: right;" ><%=formatnumber(rscat("ending_balance"),0)%></td>	
							
			
		</tr>
        
	  <%
        totOpening = totOpening + cdbl(rscat("opening_balance"))
        totDebet   = totDebet   + cdbl(rscat("debet"))
        totCredit  = totCredit  + cdbl(rscat("credit"))
        totEnding  = totEnding  + cdbl(rscat("ending_balance"))


	  rsCat.MoveNext
	Wend
	rsCat.close()
	rsCat=null%>
        <tr style="text-align: right; background-color:#f0f0f0;">
            <td align="Left" colspan="6">TOTAL</td>
            <td style="text-align: right;"><%=formatnumber(totOpening,0)%></td>
            <td style="text-align: right;"><%=formatnumber(totDebet,0)%></td>
            <td style="text-align: right;"><%=formatnumber(totCredit,0)%></td>
            <td style="text-align: right;"><%=formatnumber(totEnding,0)%></td>
        </tr>

</tbody>
</table>
</div>

</BODY>
<%
	Response.Write("<script language='javascript'> getExport();</script>")
    
%>
</HTML>
