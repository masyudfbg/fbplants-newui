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
Dim v_bulan,v_tahun,v_company,v_fcba,v_sourcecode,filter 
dim jrnfrom,jrnto,tglprd,posting,filterposting
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
<script language="JavaScript" type="text/javascript" >
	document.getElementById('tableview').style.display = "none";
</script>
<%
if trim(v_fcba) = "-" then  
filter = "  a.fcba IN (SELECT fccode FROM businessunit WHERE fccompanycode='" & v_company & "' and isactive='TRUE') " 
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

	if posting<> "" then
		filterposting = " AND b.postingflg like '%" & posting & "%' " 
	end if

			
q = " SELECT " & _ 
	" b.ACCOUNTINGYEAR, " & _ 
    " b.ACCOUNTINGPERIOD,"  & _ 
    " bs.fccompanycode,"  & _ 
    " b.fcba,"  & _ 
    " c.jobgroupcode," & _ 
    " d.fcname AS jobgroupname," & _ 
    " b.SOURCECODE," & _ 
    " a.source," & _ 
    " c.JOB_FORBUDGETCODE," & _ 
    " c.JOB_FORBUDGETNAME," & _ 
    " MAX(CASE " & _ 
    "    WHEN b.sourcecode = 'MAM-SIS' THEN e.UNITOFMEASUREMENT " & _ 
    "    ELSE c.uom_unit " & _ 
    " END) AS uom, " & _ 
    " SUM(a.amountdebit) AS total_debit, " & _ 
    " SUM(a.amountkredit) AS total_kredit, " & _ 
    " SUM(a.hk) AS total_hk, " & _ 
    " SUM(a.unit) AS total_unit, " & _ 
    " SUM(a.output) AS total_output, " & _ 
    " a.fcba || '/' || b.sourcecode || '/' || b.accountingyear || '/' ||  " & _ 
    " ROW_NUMBER() OVER ( " & _ 
    "    PARTITION BY a.fcba, b.sourcecode, b.accountingyear " & _ 
    "    ORDER BY c.jobgroupcode " & _ 
    " ) AS nomor_jurnal " & _ 
    " FROM " & _ 
    " bkmjournaldtl a " & _ 
    " JOIN bkmjournalhdr b ON a.docno = b.docno AND a.fcba = b.fcba " & _ 
    " JOIN jobmaster c ON a.jobcode = c.fccode " & _ 
    " LEFT JOIN jobgroupmaster d ON c.jobgroupcode = d.fccode " & _ 
    " LEFT JOIN bankstock e ON a.itemcode = e.fccode " & _ 
    " LEFT JOIN businessunit bs ON b.fcba = bs.fccode " & _ 
    " WHERE " & _ 
	filter & _
	filterprd & _
	filterposting & _
	" AND SOURCECODE Like '%" & v_sourcecode & "%' " & _
	filterv_groupaccount & _
    " GROUP BY " & _ 
	" b.ACCOUNTINGYEAR, " & _ 
    " b.ACCOUNTINGPERIOD,"  & _ 
    " bs.fccompanycode,"  & _ 
    " b.fcba,"  & _ 
    " c.jobgroupcode, " & _ 
    " d.fcname, " & _ 
    " b.SOURCECODE, " & _ 
    " a.source, " & _ 
	" c.JOB_FORBUDGETCODE," & _ 
    " c.JOB_FORBUDGETNAME," & _ 
	" a.fcba, " & _ 
    " b.accountingyear " & _ 
	" ORDER BY " & _ 
    " c.jobgroupcode" 
		
		 'response.write q  
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
				
		 <th style="color:#fff;">Company</th>		
		 <th style="color:#fff;">Acc.Period</th>
		 <th style="color:#fff;">Acc.Year</th>		
		 <th style="color:#fff;">FCBA</th>
		 <th style="color:#fff;">No.Journal</th> 
		 <th style="color:#fff;">SourceCode</th> 
		  <th style="color:#fff;">Source</th> 
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
			<td align="Left" ><%=rscat("fccompanycode")%></td>			
			<td align="Left" ><%=rscat("accountingperiod")%></td>			
			<td align="Left" ><%=rscat("accountingyear")%></td>			
			<td align="Left" ><%=rscat("fcba")%></td>			
			<td align="Left" ><%=rscat("nomor_jurnal")%></td>			
			<td align="Left" ><%=rscat("sourcecode")%></td>				
			<td align="Left" ><%=rscat("source")%></td>				
			<td align="Left" ><%=rscat("jobgroupcode")%></td>				
			<td align="Left" ><%=rscat("jobgroupname")%></td>	
			<td align="Left" ><%=rscat("JOB_FORBUDGETCODE")%></td>	
			<td align="Left" ><%=rscat("JOB_FORBUDGETNAME")%></td>	
			<td data-t="n" align="Right" ><%=rscat("total_debit")%></td>		
			<td data-t="n" align="Right" ><%=rscat("total_kredit")%></td>
			<td data-t="n" align="Right" ><%=rscat("total_hk")%></td>	
			<td data-t="n" align="Right" ><%=rscat("total_unit")%></td>	
			<td align="Left" ><%=rscat("uom")%></td>				
			<td data-t="n" align="Right" ><%=rscat("total_output")%></td>
				
				
			
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
