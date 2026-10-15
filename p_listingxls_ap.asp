<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>Reporting - Listing AP</title>
</head>
<%
Dim v_bulan,v_tahun,v_company,v_fcba,v_fctype, v_coa,filter
v_bulan      = trim(REQUEST("bulan"))
v_tahun      = trim(REQUEST("tahun"))
v_company    = trim(REQUEST("pcompany"))
v_fcba       = trim(REQUEST("pfcba"))
v_fctype 	 = trim(REQUEST("fctype"))
v_coa 		 = trim(REQUEST("coa"))
%>

<script src="/distexcel/tableToExcel.js"></script>
<script type="text/javascript">

	function getExport()
	{
		var bulan 	= '<%= v_bulan %>';
		var tahun 	= '<%= v_tahun %>';
		var pcompanygl = '<%= v_company %>';
		var fctype    = '<%= v_fctype %>';
		var nm = 'AP_' + pcompanygl + '_' + fctype + '_'+tahun+bulan+".xlsx";  
	//alert(nm);
	 
		TableToExcel.convert(document.getElementById("product-table"), {
            name: nm,
            sheet: {
            name: "AP"
            }
          });
		//$('table').tblToExcel();
		
	}
</script>	
<script language="JavaScript" type="text/javascript" >
	document.getElementById('tableview').style.display = "none";
</script>
<%
	if Not trim(v_coa) = 0 Then
		joinCoa = " left join supplier b on a.FCCODE = b.FCCODE" 
		filterCoa = "b.CONTROLJOB = '"& v_coa &"' AND "
	else
		joinCoa = " "
		filterCoa = "a.FCTYPE='" & v_fctype & "' AND "
	End if
	q = " SELECT a.FCBA, a.FCCODE, a.FCNAME, " & _ 
        " a.FCTYPE, a.OPENING, a.DEBET, a.CREDIT, a.ENDING " & _ 
		" FROM FAC_AP_AGINGHDR a" & _ 
		joinCoa & _
		" WHERE " & _
		filterCoa & _
		" a.PERIOD_MONTH='" & v_bulan & "' and a.PERIOD_YEAR='" & v_tahun & "' AND a.FCBA='" & v_fcba & "'"
		
		' response.write q
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
					<th align="left" colspan="23" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;">AP LISTING  <%=v_company%> Unit : <%=v_fcba%> Periode : <%=v_tahun%> / <%=v_bulan%></th>		
				</tr>
				
		 <th style="color:#fff;">FCBA</th>
		 <th style="color:#fff;">Kode</th> 
		 <th style="color:#fff;">Nama</th> 
		 <th style="color:#fff;">Type</th> 
		 <th style="color:#fff;">Saldo Awal</th> 
		 <th style="color:#fff;">Debet</th> 
		 <th style="color:#fff;">Kredit</th> 
		 <th style="color:#fff;">Saldo Akhir</th> 
		
	</tr>
		</thead>
<tbody>
	<%
		While Not rsCat.EOF
	%>
		<tr>			
			<td align="Left" ><%=rscat("FCBA")%></td>			
			<td align="Left" ><%=rscat("FCCODE")%></td>			
			<td align="Left" ><%=rscat("FCNAME")%></td>				
			<td align="Left" ><%=rscat("FCTYPE")%></td>				
			<td data-t="n" align="Right" ><%=FbpNum2(rscat("OPENING"))%></td>		
			<td data-t="n" align="Right" ><%=FbpNum2(rscat("DEBET"))%></td>		
			<td data-t="n" align="Right" ><%=FbpNum2(rscat("CREDIT"))%></td>
			<td data-t="n" align="Right" ><%=FbpNum2(rscat("ENDING"))%></td>				
			
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
