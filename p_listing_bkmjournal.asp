<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>Reporting - Detail Listing Journal</title>
</head>
<%
Dim v_bulan,v_tahun,v_company,v_fcba,v_sourcecode,v_groupaccount,filter 
dim jrnfrom,jrnto,tglprd,posting
v_bulan      = trim(REQUEST("bulan"))
v_tahun      = trim(REQUEST("tahun"))
v_company    = trim(REQUEST("company"))
v_fcba       = trim(REQUEST("fcba"))
v_sourcecode = trim(REQUEST("sourcecode"))	
v_groupaccount = trim(REQUEST("groupaccount"))	

jrnfrom = Request("jrnfrom")
jrnto   = Request("jrnto")
tglprd  = Request("tglprd")
posting = Request("posting")
%>

<script src="/distexcel/tableToExcel.js"></script>
<script type="text/javascript">

	function getExport()
	{
		var bulan 	= '<%= v_bulan %>';
		var tahun 	= '<%= v_tahun %>';
		var pcompanygl = '<%= v_company %>';
		var pfcbagl    = '<%= v_fcba %>';
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
<%
if trim(v_fcba) = "-" then  
filter = "  a.fcba IN (SELECT fccode FROM businessunit WHERE fccompanycode='" & v_company & "') " 
else
filter = "  a.fcba = '" & v_fcba & "'" 
end if

if trim(v_groupaccount) = "-" then  
filterv_groupaccount = "  " 
else
filterv_groupaccount = " AND C.jobgroupcode like '%" & v_groupaccount & "%' "
end if

  if trim(tglprd) = "Tgl" THEN
		filterprd = " and b.DOCDATE between to_date('" & jrnfrom & "','mm/dd/yyyy') and to_date('" & jrnto & "','mm/dd/yyyy') " 			   
	else
		filterprd = " AND Accountingyear='" & v_tahun & "' " & _ 
					" AND Accountingperiod ='" & v_bulan & "' " 		   
	end if

			
q = " SELECT b.accountingperiod,b.accountingyear,b.fcba,b.docno,to_char(b.docdate,'MM/DD/YYYY') docdate,b.sourcecode,a.source,a.bkmno,v.fcname vehiclename,a.description," & _ 
    " a.jobcode,c.fcname jobname, a.locationtype,a.locationcode, amountdebit,amountkredit,a.hk,a.unit,a.output, " & _
	" CASE WHEN b.sourcecode = 'MAM-SIS' THEN e.UNITOFMEASUREMENT WHEN b.sourcecode = 'FAC-VRA' THEN vg.units ELSE c.uom_unit  END AS uom, " & _ 
	" a.itemcode,a.itemname, " & _
	" c.jobgroupcode,d.fcname jobgroupname, " & _ 
	" c.job_forbudgetcode,c.job_forbudgetname " & _
	" FROM bkmjournaldtl a INNER JOIN bkmjournalhdr b ON (a.docno=b.docno AND a.fcba=b.fcba) " & _
	" INNER JOIN jobmaster c ON (a.jobcode=c.fccode) " & _
	" LEFT JOIN jobgroupmaster d ON (c.jobgroupcode=d.fccode) " & _ 
	" LEFT JOIN bankstock e ON (a.itemcode = e.fccode) " & _ 
	" LEFT JOIN vehicle v on (v.fccode=a.bkmno and v.fcba=a.fcba)" & _
	" LEFT JOIN vehiclegroup vg on (vg.VEHICLEGROUPCODE =v.VEHICLEGROUPCODE  and vg.fcba=v.fcba) " & _
	" WHERE " & _ 
	filter & _
	filterprd & _
	" AND SOURCECODE Like '%" & v_sourcecode & "%'" & _
	filterv_groupaccount & _
	" ORDER BY b.fcba,b.sourcecode,b.docdate,b.docno "
		
		'' response.write q dfsd
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
				
		 <th style="color:#fff;">Acc.Period</th>
		 <th style="color:#fff;">Acc.Year</th>		
		 <th style="color:#fff;">FCBA</th>
		 <th style="color:#fff;">No.Journal</th> 
		 <th style="color:#fff;">Tanggal</th> 
		 <th style="color:#fff;">SourceCode</th> 
		 <th style="color:#fff;">Source</th> 
		 <th style="color:#fff;">No.Reference</th> 
		 <th style="color:#fff;">Vehicle Name</th> 
		 <th style="color:#fff;">Deskripsi</th> 
		 <th style="color:#fff;">Kode Barang</th>
		 <th style="color:#fff;">Nama Barang</th>
		 <th style="color:#fff;">Kode Akun</th> 
		 <th style="color:#fff;">Nama Akun</th> 
		 <th style="color:#fff;">Target Type</th> 
		 <th style="color:#fff;">Target Code</th> 
		 <th style="color:#fff;">Group Akun</th> 
		 <th style="color:#fff;">Group Nama</th> 
		 <th style="color:#fff;">Budget Akun</th> 
		 <th style="color:#fff;">Budget Name</th> 
		 <th style="color:#fff;">Debit</th> 
		 <th style="color:#fff;">Kredit</th> 
		 <th style="color:#fff;">Hk</th> 
		 <th style="color:#fff;">Unit</th> 
		 <th style="color:#fff;">UOM</th> 
		 <th style="color:#fff;">Output(Ha)</th> 
		 
	</tr>
		</thead>
<tbody>
	<%
		While Not rsCat.EOF
	%>
		<tr>	
			<td align="Left" ><%=rscat("accountingperiod")%></td>			
			<td align="Left" ><%=rscat("accountingyear")%></td>			
			<td align="Left" ><%=rscat("fcba")%></td>			
			<td align="Left" ><%=rscat("docno")%></td>			
			<td data-t="d" align="Left" ><%=rscat("docdate")%></td>				
			<td align="Left" ><%=rscat("sourcecode")%></td>				
			<td align="Left" ><%=rscat("source")%></td>		
			<td align="Left" ><%=rscat("bkmno")%></td>	
			<td align="Left" ><%=rscat("vehiclename")%></td>		
			<td align="Left" ><%=rscat("description")%></td>
			<td align="Left" ><%=rscat("itemcode")%></td>
			<td align="Left" ><%=rscat("itemname")%></td>
			<td align="Left" ><%=rscat("jobcode")%></td>				
			<td align="Left" ><%=rscat("jobname")%></td>	
			<td align="Left" ><%=rscat("locationtype")%></td>			
			<td align="Left" ><%=rscat("locationcode")%></td>	

			<td align="Left" ><%=rscat("jobgroupcode")%></td>				
			<td align="Left" ><%=rscat("jobgroupname")%></td>	
			<td align="Left" ><%=rscat("job_forbudgetcode")%></td>				
			<td align="Left" ><%=rscat("job_forbudgetname")%></td>	
			<td data-t="n" align="Right" ><%=FbpNum2(rscat("amountdebit"))%></td>		
			<td data-t="n" align="Right" ><%=FbpNum2(rscat("amountkredit"))%></td>	
			<td data-t="n" align="Right" ><%=FbpNum2(rscat("hk"))%></td>	
			<td data-t="n" align="Right" ><%=FbpNum2(rscat("unit"))%></td>	
			<td align="Left" ><%=rscat("uom")%></td>	
			<td data-t="n" align="Right" ><%=FbpNum2(rscat("output"))%></td>	
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
