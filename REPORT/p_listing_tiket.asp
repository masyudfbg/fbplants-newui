<%@Language=VBScript%>
<% ValidateSession %>
<!--#include virtual="/ConndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
Dim ScriptName, v_txtDateFrom, v_txtDateTo, v_txtsupllier,txtproduct

'v_txtfcba 			  = session("fcba")

v_txtDateFrom = trim(Request("txtDateFrom"))
v_txtDateTo = trim(Request("txtDateTo"))
txtfcba = trim(Request("txtfcba"))
txtproduct = Request("produkmill")


 if "" & v_txtDateFrom = "" then v_txtDateFrom = textDate(date)
 if "" & v_txtDateTo = "" then v_txtDateTo = textDate(date)

Dim sqlcpo,rscpo 


%>
<html>

<head>
<title>ePlaS Reporting - Listing Tiket</title>
	<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
	<script language="JavaScript" src='/js/x.js'></script>
	<script language="JavaScript" src='/js/xtableheaderfixed.js'></script>
	<script language="JavaScript">
		xAddEventListener(window, 'load',
		  function() {
			new xTableHeaderFixed('xthf-site', window);
		  }, false
		);
	</script>
	<link rel="stylesheet" href="/css/jquery-ui.css" />
    <script src="/js/jquery-1.9.1.js"></script>
	<script src="/js/jquery-ui.js"></script>
	<link href="/js/select2/select2.min.css" rel="stylesheet" type="text/css"/>
	<script src="/js/select2/select2.min.js" type="text/javascript"></script>
    <script>
		$(function() {
			$( ".datepicker" ).datepicker();
			//getPT();
		});
	</script>
	<script language="JavaScript" src="/js/date-picker.js"></script> 
	<script language="JavaScript">
		
		function getFCBA()
		{
			var company = $("#company").val();
			showFCBA(company);
			

		}

		function showFCBA(company) {
				$('#txtfcba').select2({
					ajax: {
						delay: 1000,
						url : '/getlistfcbaajax.asp?fccompanycode=' + company,
						dataType: 'json',
						type: "POST",
						allowClear: true,
						data: function (params) {
								var query = {
												keyword: params.term
											}
								return query;
						}
					}
				});
			}
	</script>

	<style type="text/css">
		@media print
		{
		.printhide{display:none;}
		}

		.lebarselect{
				width:140px;
		}
	</style>
</head>

<body>
	<div id="wrapper">
		<div id="header-wrapper">
			<div class="printhide">
				<!--#include file="../print_excel.asp"-->
				<!--#include file="menutop.asp"-->
			</div>
		<form name="frmEntry" method="post" action="<%=ScriptName%>">
		<div class="printhide">
		
		<table id="form-table" class="centered-margin">
			<tr>
				<form name="frmEntry" method="post" action="<%=ScriptName%>">
				<td><font face="verdana" size="1">PT.</font></td>
				<td>
					<select name="company" id="company" style="width: 100px;" class="lebarselect" onchange="getFCBA()">
					<option value=""></option>
					<% 	
							Dim q,rsCompany,company
							q = " SELECT companycode, companyname " & _
							 	" FROM company " & _
							    " order by companyname"
							Set rsCompany = Server.CreateObject("ADODB.Recordset")
							rsCompany.CursorType = 3
							rsCompany.CursorLocation = 3
							rsCompany.Open q, connstr
							if not rsCompany.eof then%>
							<%
							Do While Not rsCompany.eof
							%>
								<option value="<%=rsCompany("companycode")%>" 
									<%if cstr(""&company)=cstr(""&rsCompany("companycode")) then %> selected 
									<%End If%>><%=rsCompany("companycode")%>
								</option> 
							<%rsCompany.movenext
							loop
							End if
							rsCompany.Close
							set rsCompany = Nothing
							q=""
						%>
				</select>
				</td>
				<td><font face="verdana" size="1">FCBA</font></td>
				<td><select name="txtfcba" id="txtfcba" class="lebarselect"/>
						<option value="all">Pilih FCBA--</option>
					</select>
				</td>
				<td>

					<font face="verdana" size="1">Product</font> &nbsp;&nbsp;<select name="produkmill" id="produk" style="width: 100px;" class="lebarselect">
					<option value=""></option>
					<% 	
							Dim qpro,rspro,produk
							qpro = " SELECT PRODUCTID, PRODUCTNAME " & _
							 	" FROM LHPPRODUCT " & _
							    " order by PRODUCTNAME"
							Set rspro = Server.CreateObject("ADODB.Recordset")
							rspro.CursorType = 3
							rspro.CursorLocation = 3
							rspro.Open qpro, connstr
							if not rspro.eof then%>
							<%
							Do While Not rspro.eof
							%>
								<option value="<%=rspro("PRODUCTID")%>" 
									<%if cstr(""&produkmill)=cstr(""&rspro("PRODUCTNAME")) then %> selected 
									<%End If%>><%=rspro("PRODUCTNAME")%>
								</option> 
							<%rspro.movenext
							loop
							End if
							rspro.Close
							set rspro = Nothing
							qpro=""
						%>
				</select>
				</td>
				<td><font face="verdana" size="1">Tgl Tiket</font></td>
				<td align="center">
					<input class="datepicker" type="text" maxlength="10" name="txtDateFrom" size="10" value="<%=v_txtDateFrom%>" title="Type Date mm/dd/yyyy" >
				</td>
				<td><font face="verdana" size="1">s/d</font></td>
				<td align="center">
					<input class="datepicker" type="text" maxlength="10" name="txtDateTo" size="10" value="<%=v_txtDateTo%>" title="Type Date mm/dd/yyyy" >
				</td>
				
				<td><input type="submit" value="View" /></td>
				</form>
			</tr>
		</table>
		</div>
		</form>
	</div>

	<%
		IF txtproduct="2" then
	%>
	<table border="0" width="100%" id="product-table" class="xthf-site">
		<thead>
			<tr>
				<th colspan="<%=add_newspan+13%>" style="background-color:#FFFFFF; border:none;">
					<div id="header-wrapper">
					<div id="head_judul" style="margin-bottom:-30px; position:static; text-align:left;">
				
					</div> 
					<br>
					<br>
	 
					<div id="subheader" class="centered">
						<h2>LISTING TIKET</h2>		
						<% if (v_txtdatefrom <> "" or v_txtdateto <> "") and v_txtdatefrom = v_txtdateto then%>
							<h5>TANGGAL: <%=v_txtdateto%></h5>
						<% else %>
							<h5>TANGGAL: <%=v_txtdatefrom%> - <%=v_txtdateto%></h5>
						<% end if %>
					</div>
					<br>
					</div>
				</th>
			</tr>
			<br>
			<tr>
				<th  style="color:#fff;">ESTATE/CUSTOMER</th>
				<th  style="color:#fff;">TANGGAL</th>
				<th  style="color:#fff;">NO.TIKET</th>
				<th  style="color:#fff;">SPBNO</th>
				<th  style="color:#fff;">PRODUCT</th>
				<th  style="color:#fff;">TIMEIN</th>
				<th  style="color:#fff;">TIMEOUT</th>
				<th  style="color:#fff;">TRUCT</th>
				<th  style="color:#fff;">BRUTO</th>
				<th  style="color:#fff;">TARA</th>
				<th  style="color:#fff;">NETTO1</th>
				<th  style="color:#fff;">SORTASI</th>
				<th  style="color:#fff;">NETTO2</th>
			</tr>
	</thead>
	<tbody>
	<%

	Dim rsBeliTBSall, strBeliTBSall,netto1
		Dim ttlbruto,ttltara,ttlnetto,product

	if txtfcba = "ARM" THEN

		strBeliTBSall = " select a.areaname,to_char(hd.docketdate,'dd/mm/yyyy')docketdate,hd.docketid,hd.spbno,"& _ 
		" hd.timein,hd.timeout,hd.noreg,dt.weightin,dt.weightout,dt.weightin-dt.weightout AS netto1,dt.sortasiwgt,dt.ffbweight,hd.fcba " & _
		" from LHPTIKETHDR hd  " & _
		" left join LHPFFBWBTRN dt on hd.docketid=dt.docketid  " & _
		" left join LHPAREAMST a on dt.areaid=a.areaid AND a.fcba=dt.fcba " & _
		" where hd.DOCKETDATE between to_date('" & v_txtDateFrom & "','mm/dd/yyyy') and to_date('" & v_txtDateTo & "', 'mm/dd/yyyy') " & _
		" and hd.fcba=dt.fcba and hd.fcba='" & txtfcba & "' " & _
		" order by hd.docketdate,a.areaname asc"

	ELSE 
		strBeliTBSall = " select hd.fcba areaname,to_char(hd.docketdate,'dd/mm/yyyy')docketdate,hd.docketid,hd.spbno,"& _ 
		" hd.timein,hd.timeout,hd.noreg,nvl(hd.weightin,0)weightin,nvl(hd.weightout,0)weightout,nvl(hd.weightin,0)-nvl(hd.weightout,0) AS netto1,nvl(hd.sortasiwgt,0)sortasiwgt,nvl(hd.weightin,0)-nvl(hd.weightout,0) ffbweight,hd.fcba " & _
		" from LHPTIKETHDR hd  " & _
		" where hd.DOCKETDATE between to_date('" & v_txtDateFrom & "','mm/dd/yyyy') and to_date('" & v_txtDateTo & "', 'mm/dd/yyyy') " & _
		" and hd.fcba='" & txtfcba & "' " & _
		" order by hd.docketdate,hd.fcba asc"
	END IF

		
  ' response.write strBeliTBSall	
	Set rsBeliTBSall = Server.CreateObject("ADODB.Recordset")
	rsBeliTBSall.Open strBeliTBSall, connstr, adOpenStatic
   	if Not rsBeliTBSall.eof then
	    Dim gradingqty,netto2,totalrp,weightin,weightout,pricetbs
		Dim ttlkg1,ttlgrading, ttlkg,ttlgrand,granttl
		Dim fcba
		
		ttlkg1			=	0
		ttlnetto1			=	0
		ttlgrading 		=	0
		ttlkg    		=	0
		ttlgrand 		=	0
		while Not rsBeliTBSall.eof 
			
			
		%>
		<tr>
			<td align="left"><%=rsBeliTBSall("areaname")%></td>
			<td align="center"><%=rsBeliTBSall("docketdate")%></td>
			<td align="center"><%=rsBeliTBSall("docketid")%></td>
			<td align="center"><%=rsBeliTBSall("spbno")%></td>
			<td align="center">TBS</td>
			<td align="center"><%=rsBeliTBSall("timein")%></td>
			<td align="center"><%=rsBeliTBSall("timeout")%></td>
			<td align="center"><%=rsBeliTBSall("noreg")%></td>

			<%
			if txtfcba="BRE" OR txtfcba="CDE" then 
				dim netto1bre
				netto1bre = cdbl(rsBeliTBSall("weightout"))-cdbl(rsBeliTBSall("weightin"))
			%>
				<td align="right"><%=FormatNumber(rsBeliTBSall("weightout"),0)%></td>
				<td align="right"><%=FormatNumber(rsBeliTBSall("weightin"),0)%></td>
				<td align="right"><%=FormatNumber(netto1bre,0)%></td>
				<td align="right">0</td>
				<td align="right"><%=FormatNumber(netto1bre,0)%></td>
			<%else%>
				<td align="right"><%=FormatNumber(rsBeliTBSall("weightin"),0)%></td>
				<td align="right"><%=FormatNumber(rsBeliTBSall("weightout"),0)%></td>
				<td align="right"><%=FormatNumber(rsBeliTBSall("netto1"),0)%></td>
				<td align="right"><%=FormatNumber(rsBeliTBSall("sortasiwgt"),0)%></td>
				<td align="right"><%=FormatNumber(rsBeliTBSall("ffbweight"),0)%></td>
			<%end if%>
			
		
		  </tr>	
		      
		<%
			if txtfcba="BRE" OR txtfcba="CDE" then 
            	ttlkg			=	ttlkg + cdbl(rsBeliTBSall("weightout"))
            	ttlkg1			=	ttlkg1 + cdbl(rsBeliTBSall("weightin"))
				ttlnetto1			=	ttlnetto1 + cdbl(rsBeliTBSall("netto1"))*-1
				ttlgrand			=	ttlgrand + cdbl(rsBeliTBSall("netto1"))*-1
			else
            	ttlkg			=	ttlkg + cdbl(rsBeliTBSall("weightin"))
            	ttlkg1			=	ttlkg1 + cdbl(rsBeliTBSall("weightout"))
				ttlnetto1			=	ttlnetto1 + cdbl(rsBeliTBSall("netto1"))
				ttlgrand			=	ttlgrand + cdbl(rsBeliTBSall("ffbweight"))
			end if
           
            
            ttlgrading			=	ttlgrading + cdbl(rsBeliTBSall("sortasiwgt"))
            
          
           
			rsBeliTBSall.MoveNext
		wend
		%>
		  <tr bgcolor="#F4F4A0">
			<td align="center" colspan="<%=add_newspan+8%>"><strong>TOTAL</td>
			<td align="right"><strong><%=FormatNumber(ttlkg,0)%></strong></td>
			<td align="right"><strong><%=FormatNumber(ttlkg1,0)%></strong></td>
			<td align="right"><strong><%=FormatNumber(ttlnetto1,0)%></strong></td>
			<td align="right"><strong><%=FormatNumber(ttlgrading,0)%></strong></td>
			<td align="right"><strong><%=FormatNumber(ttlgrand,0)%></strong></td>
			
		  </tr>	
		<%
	end if

		rsBeliTBSall.Close
		set rsBeliTBSall = Nothing
		%>
	
	</table>

	<% 
	
	
	else 
		dim sqltbl,sqlcpo1

		v_txtDateFrom = trim(Request("txtDateFrom"))
		v_txtDateTo = trim(Request("txtDateTo"))
		txtfcba = trim(Request("txtfcba"))
		txtproduct = Request("produkmill")

		if txtproduct<> "" then


		if txtproduct="1" then
			sqltbl = "LHPTIKETCPO"
		end if
		if txtproduct="3" then
			sqltbl = "LHPTIKETKERNEL"
		end if
		if txtproduct="5" then
			sqltbl = "LHPTIKETCKG"
		end if

	

	
		sqlcpo1 = "select to_char(h.docketdate,'dd/mm/yyyy')docketdate,dom.doorderid,cr.ContractNo,h.docketid,h.PRODUCTID ,h.TimeIn,h.TimeOut,cst.CustomerName,h.NOREG ,c.ContractorName, " & _
		" h.WEIGHTOUT, h.WEIGHTIN,h.WEIGHTOUT-h.WEIGHTIN as Netto " & _
		" from LHPTIKETHDR h  " & _
		" LEFT JOIN " & sqltbl & " d ON h.docketid=d.docketid AND h.fcba=d.fcba " & _
		" LEFT JOIN LHPDOMST dom ON dom.doid=d.doid and dom.fcba=d.fcba " & _
		" LEFT JOIN LHPCONTRACTMST cr ON dom.ContractID=cr.ContractID and h.fcba=cr.fcba " & _
		" LEFT JOIN LHPCUSTOMERMST cst ON cst.CUSTOMERID=cr.CustomerID and cst.fcba=cr.fcba " & _
		" LEFT JOIN LHPCONTRACTORMST c ON d.ContractorID = c.ContractorID and d.fcba=c.fcba " & _
		" where h.docketdate between to_Date('" & v_txtDateFrom & "','mm/dd/yyyy') and to_Date('" & v_txtDateTo & "','mm/dd/yyyy') " & _
		" AND h.PRODUCTID='" & txtproduct & "' ORDER BY h.DocketDate, h.DocketID"
		Set rscpo = Server.CreateObject("ADODB.Recordset")
		rscpo.Open sqlcpo1, connstr,adOpenStatic

		
	
		
%>
		   	<table border="0" width="100%" id="product-table" class="xthf-site">
		<thead>
			<tr>
				<th colspan="<%=add_newspan+13%>" style="background-color:#FFFFFF; border:none;">
					<div id="header-wrapper">
					<div id="head_judul" style="margin-bottom:-30px; position:static; text-align:left;">
				
					</div> 
					<br>
					<br>
	 
					<div id="subheader" class="centered">
						<h2>LISTING TIKET</h2>		
						<% if (v_txtdatefrom <> "" or v_txtdateto <> "") and v_txtdatefrom = v_txtdateto then%>
							<h5>TANGGAL: <%=v_txtdateto%></h5>
						<% else %>
							<h5>TANGGAL: <%=v_txtdatefrom%> - <%=v_txtdateto%></h5>
						<% end if %>
					</div>
					<br>
					</div>
				</th>
			</tr>
			<br>
			<tr>
				<th  style="color:#fff;">TANGGAL</th>
				<th  style="color:#fff;">DO</th>
				<th  style="color:#fff;">KONTRAK</th>
				<th  style="color:#fff;">NO.TIKET</th>
				<th  style="color:#fff;">PRODUCT</th>
				<th  style="color:#fff;">TIMEIN</th>
				<th  style="color:#fff;">TIMEOUT</th>
				<th  style="color:#fff;">ESTATE/CUSTOMER</th>
				<th  style="color:#fff;">CONTRACTOR</th>
				<th  style="color:#fff;">TRUCK</th>
				<th  style="color:#fff;">BRUTO</th>
				<th  style="color:#fff;">TARA</th>
				<th  style="color:#fff;">NETTO</th>
			</tr>
			</thead>

			<%
				
				ttlbruto		=	0
            	ttltara			=	0
            	ttlnetto		=	0

			
				while Not rscpo.eof 
					product = rscpo("PRODUCTID")
					if product="1" then 
						product="CPO"
					elseif product="3" then
						product="KERNEL"
					elseif product="4" then
						product="AMPAS"
					elseif product="5" then
						product="CANGKANG"
					elseif product="6" then
						product="KOMPOS"
					end if
			%>

			<tr>
				<td  align="left"><%=rscpo("docketdate")%></td>
				<td  align="left"><%=rscpo("doorderid")%></td>
				<td  align="left"><%=rscpo("ContractNo")%></td>
				<td  align="left"><%=rscpo("docketid")%></td>
				<td  align="left"><%=product%></td>
				<td  align="left"><%=rscpo("TimeIn")%></td>
				<td  align="left"><%=rscpo("TimeOut")%></td>
				<td  align="left"><%=rscpo("CustomerName")%></td>
				<td  align="left"><%=rscpo("ContractorName")%></td>
				<td  align="left"><%=rscpo("NOREG")%></td>
				<td  align="right"><%=FormatNumber(rscpo("WEIGHTOUT"),0)%></td>
				<td  align="right"><%=FormatNumber(rscpo("WEIGHTIN"),0)%></td>
				<td  align="right"><%=FormatNumber(rscpo("Netto"),0)%></td>
				
			</tr>
		<%	

			ttlbruto			=	ttlbruto + cdbl(rscpo("WEIGHTOUT"))
            ttltara				=	ttltara + cdbl(rscpo("WEIGHTIN"))
            ttlnetto			=	ttlnetto + cdbl(rscpo("Netto"))
           
		rscpo.MoveNext
			
		wend
		%>
		 <tr bgcolor="#F4F4A0">
			<td align="center" colspan="<%=add_newspan+10%>"><strong>TOTAL</td>
			<td align="right"><strong><%=FormatNumber(ttlbruto,0)%></strong></td>
			<td align="right"><strong><%=FormatNumber(ttltara,0)%></strong></td>
			<td align="right"><strong><%=FormatNumber(ttlnetto,0)%></strong></td>
			
		  </tr>	
	
		
	</table>
	<%end if
	end if%>
	</tbody>

<!--#include file="bottom.Asp"-->