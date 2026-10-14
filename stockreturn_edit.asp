<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include file="qtyOnHandFunction.inc"-->
<%
dim businessunit,fcbaname,CompanyName,q,sqlhdr,simpan
dim returnissueno,tglreturn,store,noref,remarkhdr
dim nQtdFields,nCount
dim itemcode,itemname,uom,qtyissue,qtyreturn,remarkdtl,linenokey,noreturn
dim sqlnomor,rsnomor,tglprev,currentnomor,currentnomorfix
dim bulansplit,tglsplit,thnsplit
dim sqledit,rsedit
dim sqldtl,rsdtl,sqldelhdr,sqldeldtl, totalrows,fcCode
dim fcba,strQueryStatus,rsReceiveDetail,originalQuantity,txtItemCodeDelete,txtLineNoKey,storecode
	DIM qrylessstockonhand,qrylessstockledger,qrynewtockonhand,qrynewtockledger
	dim ScriptName,modulfile,executeQry,rsStockLedger

returnissueno	= trim(request("returnno"))
tglreturn 		= trim(request("returndate"))
businessunit    = trim(request("fcba"))
simpan		    = trim(request("simpan"))


q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & businessunit & "' "
	'response.write q
set rsCat = connStr.Execute(q)
if not rsCat.eof then
  
   fcbaname = rsCat("fcName")
   CompanyName=rsCat("FCCOMPANYNAME")
 
end if

'GET COUNT DATA DETAIL
sqldtl = "SELECT COUNT(D.ITEMCODE) AS TOTALROWS " & _
" From   RETURNISSUEDTL D " & _
" LEFT JOIN STOCK S ON D.ITEMCODE=S.FCCODE AND D.FCBA=S.FCBA" & _
" where D.RETURNNO='" & returnissueno & "'"
'Response.write sqldtl
set rsdtl = connStr.Execute(sqldtl)
if not rsdtl.eof then
   totalrows = rsdtl("TOTALROWS")
   
end if

'GET DATA RETURN TRNSAKSI
sqledit = "Select RETURNNO, to_char(RETURNDATE,'mm/dd/yyyy')RETURNDATE, STORE, NOREF, REMARK, " & _ 
" FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, " & _
" LASTTIME, POSTINGFLG " & _
" From   RETURNISSUEHDR " & _
" where RETURNNO='" & returnissueno & "'"
set rsedit = connStr.Execute(sqledit)
if not rsedit.eof then
   tglreturn 		= rsedit("RETURNDATE")
   fcba				= rsedit("fcba")
   noref			= rsedit("NOREF")
   remarkhdr		= rsedit("REMARK")
end if

'GET DATA RETURN TRNSAKSI DETAILS
sqldtl = "SELECT D.ITEMCODE,S.FCNAME, D.UOM, D.QTYISSUE, D.QTYRETURN, D.REMARKDTL, " & _
" D.RETURNNO, D.LINENOKEY, D.FCBA " & _
" From   RETURNISSUEDTL D " & _
" LEFT JOIN STOCK S ON D.ITEMCODE=S.FCCODE AND D.FCBA=S.FCBA" & _
" where D.RETURNNO='" & returnissueno & "'"
'Response.write sqldtl
set rsdtl = connStr.Execute(sqldtl)

'PROSES SIMPAN DATA
if simpan = "SIMPAN" THEN 
	noreturn	= trim(request("noreturn"))
	
	tglreturn	= trim(request("tglreturn"))
	store		= trim(request("store"))
	noref		= trim(request("noref"))
	remarkhdr	= trim(request("remarkhdr"))
	fcba		= trim(request("fcba"))

	
	'delete dulu tabel returnhdr
	sqldelhdr = "delete from RETURNISSUEHDR where returnno='" & noreturn & "'"
	ExecuteAndLog "StockReturn.asp", sqldelhdr
	'connStr.execute(sqldelhdr)


	'simpan ke tabel hdr return
	sqlhdr = "insert into RETURNISSUEHDR(RETURNNO, RETURNDATE, STORE, NOREF, REMARK, " & _
	" FCENTRY,FCEDIT,FCIP,FCBA,LASTUPDATE,LASTTIME) " & _
	" values('" & noreturn & "',TO_DATE('" & tglreturn & "', 'mm/dd/yyyy'),'" & store & "','" & noref & "','" & remarkhdr & "', " & _ 
	" '" & Session("CodeName") & "', " & _ 
	" '" & Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") & "','" & fcba & "',SYSDATE,SYSDATE)"

	ExecuteAndLog "StockReturn.asp", sqlhdr
	'connstr.BeginTrans
	'Response.Write sqlhdr
	'connStr.execute(sqlhdr)

	'looping di tabel detail
	nQtdFields 	= request.form("jumldata")
	nCount 		= 1

	do while cint(nCount) <= cint(nQtdFields) 
		itemcode 		= request.form("itemcode" & nCount)
	'	fcname 			= request.form("fcname" & nCount)
		uom 			= request.form("uom" & nCount)
		qtyissue 		= request.form("qtyissue" & nCount)
		qtyreturn 		= request.form("qtyreturn" & nCount)
		remarkdtl 			= request.form("remark" & nCount)
		linenokey 			= request.form("linenokey" & nCount)

		'ambil qty original di returndetail
		q = "Select ITEMCODE, QTYRETURN, LINENOKEY from RETURNISSUEDTL where RETURNNO='" & noreturn & _
		"' and fcBA='" & fcba & "' AND itemcode='" & itemcode & "' AND linenokey='" & linenokey & "'"

		Set rsReceiveDetail = Server.CreateObject("ADODB.Recordset")
		rsReceiveDetail.Open q, connstr
		
				
		IF Not rsReceiveDetail.eof then
			originalQuantity = rsReceiveDetail("QTYRETURN")
			txtItemCodeDelete = rsReceiveDetail("ITEMCODE")
			txtLineNoKey = rsReceiveDetail("LINENOKEY")
		End if
		rsReceiveDetail.Close
		set rsReceiveDetail = Nothing		

		'function untuk mengembalikan qty ketika ada proses delete atau update
	
		qrylessstockonhand	= cmdUpdateStockOnHand(store,txtItemCodeDelete,originalQuantity,"RT","Update")
		qrylessstockledger 	= cmdUpdateStockLedger(store,txtItemCodeDelete,tglreturn,originalQuantity,"RT","Update")

		'delete dulu tabel returndtl
		sqldeldtl = "delete from RETURNISSUEDTL where returnno='" & noreturn & _ 
		"' and fcBA='" & fcba & "' AND itemcode='" & itemcode & "' AND linenokey='" & linenokey & "'"

		ExecuteAndLog "StockReturn.asp", sqldeldtl
		'connStr.execute(sqldeldtl)

		if qtyreturn > 0 then
			sqldtl = "insert into RETURNISSUEDTL(RETURNNO, ITEMCODE, UOM, QTYISSUE, QTYRETURN,REMARKDTL,LINENOKEY,FCBA) " & _
			" values('" & noreturn & "','" & itemcode & "','" & uom & "','" & qtyissue & "', " & _ 
			" '" & qtyreturn & "','" & remarkdtl & "','" & linenokey & "','" & fcba & "')"
			ExecuteAndLog "StockReturn.asp", sqldtl
		'	Response.Write sqldtl
			'connstr.BeginTrans
			'connStr.execute(sqldtl)

			'insert kembali ke tabel stockon hand
			qrynewtockonhand	= UpdateNewStockOnHand(store,txtItemCodeDelete,qtyreturn,"RT","Update")
			qrynewtockledger 	= UpdateNewStockLedger(store,txtItemCodeDelete,tglreturn,qtyreturn,"RT","Update")

		end if

		

	 nCount 	= nCount + 1
	loop
	response.Redirect("stockreturn_edit.asp?returnno="&noreturn)
	'response.redirect("bkm_traksi_edit.asp?bkmno="&bkmno)
END IF

if simpan="DELETE" then
    noreturn	= trim(request("noreturn"))	
	store		= trim(request("store"))
	tglreturn 	= trim(request("returndate"))

	q = "Select ITEMCODE, QTYRETURN, LINENOKEY from RETURNISSUEDTL where RETURNNO='" & noreturn & _
		"' and fcBA='" & Session("fcBA") & "'"
'	Response.Write q
'	Response.End

	Set rsReceiveDetail = Server.CreateObject("ADODB.Recordset")
	rsReceiveDetail.Open q, connstr

	While Not rsReceiveDetail.EOF

		originalQuantity 	= rsReceiveDetail("QTYRETURN")
		txtItemCodeDelete 	= rsReceiveDetail("ITEMCODE")
		txtLineNoKey 		= rsReceiveDetail("LINENOKEY")

		'function untuk mengembalikan qty ketika ada proses delete
		qrylessstockonhand	= cmdUpdateStockOnHand(store,txtItemCodeDelete,originalQuantity,"RT","Delete")
		qrylessstockledger 	= cmdUpdateStockLedger(store,txtItemCodeDelete,tglreturn,originalQuantity,"RT","Delete")

		'Response.Write qrylessstockledger
		'Response.end

	rsReceiveDetail.MoveNext	 
	Wend
	
	rsReceiveDetail.Close
	set rsReceiveDetail = Nothing		

	
	'delete dulu tabel returndtl
	sqldeldtl = "delete from RETURNISSUEDTL where returnno='" & noreturn & "'"
	ExecuteAndLog "StockReturn.asp", sqldeldtl
	'connStr.execute(sqldeldtl)

	'delete dulu tabel returnhdr
	sqldelhdr = "delete from RETURNISSUEHDR where returnno='" & noreturn & "'"
	ExecuteAndLog "StockReturn.asp", sqldelhdr
	'connStr.execute(sqldelhdr)

	
end if

if simpan="POSTING" then
  					strQueryStatus = "UPDATE RETURNISSUEHDR SET POSTINGFLG='1',lastupdate=sysdate,fcedit='" & session("CodeName") & "' WHERE fcba = '" & fcba & "' AND returnno = '" & returnissueno & "'"
					'response.write strQueryStatus & "<br>"
					connStr.execute(strQueryStatus)
					
					strQueryStatus = "BEGIN fac_journal_mam_api.MAM_SRI_DOCNO('" & returnissueno & "', TO_NUMBER(TO_CHAR(TO_DATE('" & tglreturn  & "','MM/DD/YYYY'),'YYYY')), " & _ 
			                         "TO_NUMBER(TO_CHAR(TO_DATE('" & tglreturn  & "','MM/DD/YYYY'),'MM')) , '" & fcba & "','System'); " & _ 
									 "END;"			
					'response.write strQueryStatus & "<br>"
					connStr.execute(strQueryStatus)
					
					
					strQueryStatus = "update BKMJOURNALHDR set postingflg=1,lastupdate=sysdate,fcedit='" & session("CodeName") & "' where fcba='" & fcba & "' AND docno='" & returnissueno & "'"
					'response.write strQueryStatus & "<br>"
					connStr.execute(strQueryStatus)
	
					strQueryStatus = "BEGIN fac_closing_accounting_api.ProsesClosing(TO_NUMBER(TO_CHAR(TO_DATE('" & tglreturn  & "','MM/DD/YYYY'),'YYYY')), " & _ 
			                         "TO_NUMBER(TO_CHAR(TO_DATE('" & tglreturn  & "','MM/DD/YYYY'),'MM')) , '" & fcba & "','System'); " & _ 
									 "END;"			
					'response.write strQueryStatus	& "<br>"			 
					connStr.execute(strQueryStatus)

	
end if
%>

<html>
	<head>
		<title>FBPlants - Stock Return Edit</title>
		<script type="text/javascript" language="JavaScript1.2" src="stlib.js" ></script>
		<link rel="StyleSheet" href="dtree.css" type="text/css" />
		<link rel="StyleSheet" href="style-content.css" type="text/css" />
		<link rel="stylesheet" href="css/jquery-ui.css" />
		<!-- HADIR/ font ICON -->
		<link rel="stylesheet" href="css/font/font-awesome.css" />
		<!-- END:HADIR -->
		<script type="text/javascript" src="dtree.js"></script>
		<script src="/js/jquery-1.9.1.js"></script>
		<script src="/js/jquery-ui.js"></script>
		<link rel="stylesheet" href="css/jquery-ui.css" />
		<link rel="StyleSheet" href="style-content.css" type="text/css" />

		<link href="/js/spry/widgets/textfieldvalidation/SpryValidationTextField.css" rel="stylesheet" type="text/css"/>
		<script src="/js/spry/widgets/textfieldvalidation/SpryValidationTextField.js" type="text/javascript"></script>
		
		<link href="/js/spry/widgets/selectvalidation/SpryValidationSelect.css" rel="stylesheet" type="text/css"/>
		<script src="/js/spry/widgets/selectvalidation/SpryValidationSelect.js" type="text/javascript"></script>

	</head>

<style>
	.posisikirijudul{
		position: relative;
		left: 0px;
	}
	.posisikiri{
		position: relative;
		left: -15px;
	}
	.lebarsection{
		width: 220px;
	}
	.lebargangcode{
		width: 200px;
	}
	.lebarhk{
		width: 45px;
	}
	.lebarabsen{
		width: 45px;
	}
	.lebarlokasi{
		width: 50px;
	}
	.lebartc{
		width: 60px;
	}
	.lebarovtime{
		width: 70px;
	}
	.lebarrate{
		width: 60px;
	}
	.lebaractivity{
		width: 110px;
	}
	.lebarallow{
		width: 90px;
	}
	.lebarupah{
		width: 120px;
	}
	.warnagaris{
		border-color: #FFA800;
	}
	.tebal{
		font-weight: bold;
	}
	.ratakiri{
		text-align: left;
	}
	.ratakanan{
		text-align: right;
	}
	.ratatengah{
		text-align: center;
	}
	.lebaritembarang{
		width: 400px;
	}
	.lebaractbarang{
		width: 300px;
	}
	.lebarvehicle{
		width: 200px;
	}
	.lebartime{
		width: 90px;
	}

	.lebarspkno{
		width: 250px;
	}
	
</style>

<script>

</script>

<body>

<!--#include file="top.jsp"-->
<div style=" width:21.7%; float:left;margin-left:0.3%; margin-top:0.3%; margin-bottom:0.3%;">
 <!--#include file="menu.asp"-->
</div>
<div class="judul_head" valign="top" bgcolor="#fff" style=" width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;"> 
 	<div class="judul">
	   Module : <%=request("txtModule")%> / <%=request("txtTable")%>
	</div>
<div class="box_both" style="overflow:scroll;">  
		
<form name="frmEntry" method="post" action="stockreturn_edit.asp">
	<input type="hidden" name="jumldata" id="jumldata" value="<%=totalrows%>" />
	<input type="hidden" name="proses" id="proses" value="1" />
	<table class="data posisikiri" border="0" width="100%">			
		<tr>
			<td align="left" >Order No</td>
			<td align="left">
				<input type="text" readonly name="noreturn" id="noreturn" value="<%=returnissueno%>"   class="lebarsection">
				<input type="hidden" readonly name="fcba" id="fcba" value="<%=fcba%>"   class="lebarsection">
		</tr>
		<tr>	
			</td>
			<td align="left" >Date.Return</td>
			<td align="left">
				<input type="text" class="datepicker" maxlength="15" name="tglreturn"  id="tglreturn" size="15" value="<%=tglreturn%>" title="Tanggal Return">  <font color="#FF0000">&nbsp;*</font></td>
			</td>
		</tr>
		<tr>	
			<td align="left" >No. Issue</td>
			<td align="left">
				<input type="text" readonly name="noref" id="noref" value="<%=noref%>"   class="lebarsection">
			
			</td>
			
		</tr>
		<tr>
			<td align="left" >Store</td>
			<td align="left">
				<select name="store" id="store" class="lebargangcode" style=height:24px>
					<option value="ST1">ST1 - GUDANG KEBUN</option> 
					
				</select>
			</td>
		</tr>
		<tr>
			<td align="left" >Remark</td>
			<td align="left">
				<textarea align='left' name="remarkhdr" id="remarkhdr" rows="6" cols="40"><%=remarkhdr%>
				</textarea>
			</td>
			
		</tr>

	</table>
	<br>
	<a href="?returnno=<%=returnissueno%>&returndate=<%=tglreturn%>&fcba=<%=businessunit%>&simpan=POSTING" onClick="return confirm('Are you sure to Posting Journal this data :  <%=noreturn%> ?');">Posting</a>

	<hr />
		
	<table align=center border="0" class='data ' cellpadding="0" cellspacing="0" width='100%' id='tableall'>
		<tr>
			<td colspan="">
					<table class='data' id='table2'>
						<tr id="trheader2">
							<th>Item Code</th>
							<th>Item Name</th>
							<th>UOM</th>
							<th>Qty Issue</th>
							<th>Qty Return</th>
							<th colspan="1">Remark</th>
						</tr>
						<%
							dim no 
							no = 1
							While Not rsdtl.EOF
						%>
						<tr id=''>
							<td align="left">
								<%=rsdtl("itemcode")%>
								<input type="hidden" class="ratatengah" name="itemcode<%=no%>" id="itemcode" value="<%=rsdtl("itemcode")%>"></input>
							</td>
							<td align="left">
								<%=rsdtl("fcname")%>
								
							</td>
							<td align="left">
								<%=rsdtl("uom")%>
								<input type="hidden"  name="uom<%=no%>" id="uom<%=no%>" value="<%=rsdtl("uom")%>"></input>
							</td>
							<td align="center">
								<%=rsdtl("qtyissue")%>
								<input type="hidden"  name="qtyissue<%=no%>" id="qtyissue" value="<%=rsdtl("qtyissue")%>"></input>
							</td>
							<td style=width:40px>
								
								<input type="text" class="ratatengah" name="qtyreturn<%=no%>" id="qtyreturn" onkeyup="hanyaAngka(this);" value="<%=rsdtl("qtyreturn")%>"></input>
							</td>
							<td align="left">
								<%=rsdtl("remarkdtl")%>
								<input type="hidden"  name="remark<%=no%>" id="remark" value="<%=rsdtl("remarkdtl")%>"></input>
								<input type="hidden"  name="linenokey<%=no%>" id="linenokey" value="<%=rsdtl("linenokey")%>"></input>
							</td>
						</tr>
						<%
							 rsdtl.MoveNext	 
							 no = no + 1
							Wend
						%>
						<tr>
							<td colspan='6' align=center>
							     <input Type="submit"  name="simpan" value="SIMPAN">
								&nbsp;&nbsp;
								<input type=button onclick="window.close();" name="btncancel" value="CLOSE">
								&nbsp;&nbsp;
								<input type=submit name="simpan" value="DELETE">
							</td>
						
						</tr>
						
					</table>
				
			</td>
		</tr>
		
	</table>

</form>
</div>
</div></div>
<!--#include file="bottom.jsp"-->
</body>
</html>
<link href="/js/select2/select2.min.css" rel="stylesheet" type="text/css"/>
<script src="/js/select2/select2.min.js" type="text/javascript"></script>

<script type="text/javascript">


	function hanyaAngka(evt) {
		evt.value = evt.value.replace(/[^0-9,.-]/g,"");
	}

$(function() {	
	$("#tglreturn").datepicker();
});



function numformat(val) {
	var nStr = val.toString();
	nStr = nStr.replace(/\,/g, "");
	var x = nStr.split('.');
	var x1 = x[0];
	var x2 = x.length > 1 ? '.' + x[1] : '';
	var rgx = /(\d+)(\d{3})/;
	while (rgx.test(x1)) {
		x1 = x1.replace(rgx, '$1' + '.' + '$2');
	}
	return x1 + x2;
}

</script>