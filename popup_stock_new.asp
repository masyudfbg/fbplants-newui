<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->  
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>ePlaS - Pilih Stock untuk Issue</title>
<link rel="stylesheet" href="css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="css/jquery-ui.css" />
<link rel="stylesheet" href="/css/fbp-shell.css?v=20261009d" type="text/css" />
<link rel="stylesheet" href="/css/fbp-table.css?v=20261010c" type="text/css" />
<script type="text/javascript" src="/js/fbp-shell.js?v=20261014b"></script>
<script type="text/javascript" src="/js/fbp-table.js?v=20261009c"></script>
</head>
<style type="text/css">
/* Popup pilih stock: kop, filter, tabel ala DataTables */
body { background: #f3f6f5; }
#wrapper { padding: 16px 18px; }
.ps-head {
	display: flex; flex-wrap: wrap; align-items: flex-start; justify-content: space-between; gap: 12px;
	padding: 14px 18px; background: #fff; border: 1px solid #e1e7e5; border-radius: 14px; border-left: 5px solid #1B7863;
}
.ps-company { font-size: 12px; font-weight: 700; color: #5d6d69; text-transform: uppercase; letter-spacing: .6px; }
.ps-title { margin: 4px 0 0; font-size: 20px; font-weight: 700; color: #12584a; }
.ps-sub { margin-top: 2px; font-size: 12.5px; color: #5d6d69; }
.ps-chip { display: inline-flex; gap: 6px; padding: 3px 10px; font-size: 12px; background: #eef5f3; border: 1px solid #d3e6e0; border-radius: 999px; }
.ps-chip b { font-weight: 600; color: #5d6d69; }
.ps-filter {
	display: grid; grid-template-columns: minmax(200px, 1fr) minmax(200px, 1.4fr) auto; gap: 12px; align-items: end;
	margin: 12px 0; padding: 14px 16px; background: #fff; border: 1px solid #e1e7e5; border-radius: 14px;
}
.ps-filter label { display: block; margin: 0 0 5px; font-size: 12px; font-weight: 600; color: #5d6d69; }
.ps-filter select, .ps-filter input[type="text"] {
	box-sizing: border-box; width: 100%; height: 38px; margin: 0; padding: 0 12px;
	font: inherit; font-size: 13px; color: #17262a; background-color: #fff; border: 1px solid #d5dfdc; border-radius: 8px; outline: none;
}
.ps-filter select {
	padding-right: 32px; -webkit-appearance: none; appearance: none; cursor: pointer;
	background: #fff url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='8' viewBox='0 0 12 8'%3E%3Cpath d='M1 1.5l5 5 5-5' fill='none' stroke='%235d6d69' stroke-width='1.6' stroke-linecap='round' stroke-linejoin='round'/%3E%3C/svg%3E") no-repeat right 12px center;
}
.ps-filter select:focus, .ps-filter input[type="text"]:focus { border-color: #1B7863; box-shadow: 0 0 0 3px rgba(27, 120, 99, .18); }
.ps-btn { height: 38px; padding: 0 20px; font: inherit; font-size: 13px; font-weight: 600; color: #fff; background: #1B7863; border: 0; border-radius: 8px; cursor: pointer; }
.ps-btn:hover { background: #12584a; }
.ps-panel { padding: 12px 14px 4px; background: #fff; border: 1px solid #e1e7e5; border-radius: 14px; }
.ps-panel .fbt-tools { margin-bottom: 10px; }
.ps-scroll { overflow: auto; max-height: calc(100vh - 330px); min-height: 160px; border: 1px solid #e1e7e5; border-radius: 10px; }
#product-table { margin: 0; border-collapse: separate; border-spacing: 0; min-width: 100%; }
#product-table tr.fbt-toolbar-row { display: none; }
#product-table thead tr:last-child th { position: sticky; top: 0; z-index: 2; white-space: nowrap; background: #1B7863 !important; border-width: 0 1px 0 0; border-color: #2f8a75; }
#product-table td { border-width: 0 0 1px 0 !important; border-color: #e6ecea !important; vertical-align: middle; white-space: nowrap; }
#product-table td.ps-name { white-space: normal; min-width: 240px; }
#product-table td.ps-code { font-weight: 600; color: #12584a; }
#product-table td.ps-qty { text-align: right; font-weight: 700; font-variant-numeric: tabular-nums; }
#product-table td.ps-qty.zero { color: #b5c0bd; font-weight: 400; }
#product-table td.ps-qty.neg { color: #a8322a; }
#product-table tbody tr { cursor: pointer; }
#product-table tbody tr.ps-disabled { cursor: default; }
#product-table tbody tr.ps-disabled > td { color: #8a9a96; }
.ps-na { color: #b5c0bd; }
#product-table tbody tr:hover > td { background: #eef6f3 !important; }
.ps-select { height: 30px; padding: 0 14px; font: inherit; font-size: 12px; font-weight: 600; color: #1B7863; background: #fff; border: 1px solid #1B7863; border-radius: 7px; cursor: pointer; }
.ps-select:hover { color: #fff; background: #1B7863; }
.ps-empty { padding: 28px !important; text-align: center; color: #5d6d69; white-space: normal !important; cursor: default; }
@media (max-width: 640px) { .ps-filter { grid-template-columns: 1fr; } }
</style>
<body class="rpt-form">
<%
  Dim tglbulanprev,bulanprevsplit,tglprevsplit,thnprevsplit,v_txthi,strStockOnHand,rsGrnInv
  Dim openingqty,receivenote,issue,transportask,receivetask,qtyonhand
  Dim bulannow,tahunnow

  Function IsPostBack()
     If Request.ServerVariables("Request_Method") = "POST" Then
          IsPostBack = True
     Else
          IsPostBack = False
     End If
   End Function

  '---- Escape tanda kutip untuk literal string SQL / encode output HTML
  Function SqlStr(v)
    SqlStr = Replace(Trim(v & ""), "'", "''")
  End Function
  Function H(v)
    If IsNull(v) Then H = "" Else H = Server.HTMLEncode(v & "")
  End Function

  v_txtFcba = Request("fcba") 
  varDefaultBa = Session("fcBA")   
  If (v_txtFcba = "" Or Len(Trim(v_txtFcba)) = 0) Then
    v_txtFcba = varDefaultBa
  End If  
    
  q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & SqlStr(v_txtfcba) & "' "
  set rsCat = connStr.Execute(q)
  if not rsCat.eof then
     fcbaname = rsCat("fcName")
     CompanyName=rsCat("FCCOMPANYNAME")
     MgrName = rsCat("FCMANAGER")
     KTUName = rsCat("FCKTU")
     CropName=rsCat("FCCROP")
     FcbaType=rsCat("fcType")
  else
      nSpanTotal = nSpanTotal - 1 
  end if

  
  v_txthi         = date()
  IssueDate       = Trim(Request("IssueDate"))
  '---- Tanggal kosong / tidak valid (MM/DD/YYYY) -> hari ini, agar month()/to_date tidak error
  If Len(IssueDate) <> 10 Or Not IsDate(IssueDate) Then IssueDate = Right("0" & Month(Date), 2) & "/" & Right("0" & Day(Date), 2) & "/" & Year(Date)
  tglbulanprev    = IssueDate
  bulanprevsplit  = month(tglbulanprev)
  tglprevsplit    = day(tglbulanprev)
  thnprevsplit    = year(tglbulanprev)
  
  bulannow        = left(IssueDate,2)
  tahunnow        = right(IssueDate,4)
  tglawal 		  = bulannow & "/" & "01" & "/" & tahunnow
  
  v_SubGroup	   = Request("subgroupcode")
  v_Cari	       = Request("txtcari")
  
  
  'Response.Write v_SubGroup & "-" & v_Cari

if cint(tahunnow) < 2023 then
strStockOnHand		=	"select t.storecode,store_api.get_fcname(t.storecode, t.fcba) storename,t.STOCKCODE,stock_api.get_stockname(t.STOCKCODE, t.fcba) stockname,stock_api.get_uom(t.STOCKCODE, t.fcba) uom,stock_api.get_StockSubGroup(t.STOCKCODE,t.fcba) StockSubGroup, " & _ 
" sum(nvl(t.openingbalanceqty,0)) as openingbalanceqty,sum(nvl(t.receivenoteqty,0)) as receivenoteqty, " & _
" sum(nvl(t.issueqty,0)) as issueqty,sum(nvl(t.quantitytransfered,0)) as quantitytransfered,sum(nvl(t.quantityreceived,0)) as quantityreceived,sum(nvl(t.qtywriteoff,0)) as qtywriteoff, " & _
" (sum(nvl(t.openingbalanceqty,0))+sum(nvl(t.receivenoteqty,0))-sum(nvl(t.issueqty,0))-sum(nvl(t.quantitytransfered,0))+sum(nvl(t.quantityreceived,0))+sum(nvl(t.qtywriteoff,0))) as qty " & _
" from( " & _
    " select fcba,l.STORECODE,l.itemcode STOCKCODE, " & _
    " l.openingbalanceqty, " & _
    " 0 as receivenoteqty,0 as issueqty,0 as quantitytransfered,0 as quantityreceived,0 as qtywriteoff " & _
    " from stockledger l " & _
    " where l.YEAR='" & thnprevsplit & "' and l.MONTH='" & bulanprevsplit & "' and l.fcba='" & SqlStr(v_txtfcba) & "' " & _
    "      union all " & _
    " select r.fcba,r.storecode,rd.purchaseitemcode STOCKCODE, " & _
    " 0 as openingbalanceqty,rd.quantity as receivenoteqty,0 as issueqty,0 as quantitytransfered,0 as quantityreceived,0 as qtywriteoff " & _
    " from receivenote r,receivenotedetail rd " & _
    " where r.receivenotecode=rd.receivenotecode and r.rndate between to_date('" & tglawal & "','mm/dd/yyyy') and to_date('" & IssueDate & "','mm/dd/yyyy') " & _
    " and r.fcba='" & SqlStr(v_txtfcba) & "' " & _
    "      union all " & _
    " select h.fcba,h.storecode,d.STOCKCODE," & _
    " 0 as openingbalanceqty,0 as receivenoteqty,d.quantity as issueqty,0 as quantitytransfered, " & _
    " 0 as quantityreceived,0 as qtywriteoff " & _
    " from issue h,issuedetails d " & _
    " where h.sivcode=d.sivcode and h.fcba='" & SqlStr(v_txtfcba) & "' and h.sivdate between to_date('" & tglawal & "','mm/dd/yyyy') and to_date('" & IssueDate & "','mm/dd/yyyy') " & _
    "      union all " & _
    " select d.fcba,d.storecode,d.STOCKCODE, " & _
    " 0 as openingbalanceqty,0 as receivenoteqty,0 as issueqty,d.quantitytransfered,0 as quantityreceived,0 as qtywriteoff " & _
    " from transporttaskdtl_tab d,transporttask_tab h " & _
    " where h.orderno=d.orderno and h.rowstateid not in('Rejected') and h.fcba='" & SqlStr(v_txtfcba) & "' AND h.orderdate between to_date('" & tglawal & "','mm/dd/yyyy') and to_date('" & IssueDate & "','mm/dd/yyyy') " & _
    "     union all " & _
    " select d.fcba,d.locationcode,d.STOCKCODE, " & _
    " 0 as openingbalanceqty,0 as receivenoteqty,0 as issueqty,0 as quantitytransfered,d.quantityreceived,0 as qtywriteoff  " & _
    " from receivedtask_tab h,receivedtaskdtl_tab d " & _
    " where h.orderno=d.orderno and h.orderdate between to_date('" & tglawal & "','mm/dd/yyyy') and to_date('" & IssueDate & "','mm/dd/yyyy')  and h.fcba='" & SqlStr(v_txtfcba) & "' " & _
    "  UNION ALL " & _
    " select h.fcba,h.STORECODE locationcode,d.stockcode, " & _
    " 0 as openingbalanceqty,0 as receivenoteqty,0 as issueqty,0 as quantitytransfered, " & _
    " 0 as quantityreceived,d.quantity as qtywriteoff  " & _
    " from writeoff h,writeoffdetails d " & _
    " where h.writeoffcode=d.writeoffcode and h.wodate between to_date('" & tglawal & "','mm/dd/yyyy') and to_date('" & IssueDate & "','mm/dd/yyyy') and h.fcba='" & SqlStr(v_txtfcba) & "' " & _ 	
	"  UNION ALL " & _
    " select h.fcba,h.STORE locationcode,itemcode stockcode, " & _
    " 0 as openingbalanceqty,0 as receivenoteqty,0 as issueqty,0 as quantitytransfered, " & _
    " qtyreturn as quantityreceived,0 as qtywriteoff  " & _
    " from returnissuehdr h,returnissuedtl d " & _
    " where h.returnno=d.returnno and h.returndate between to_date('" & tglawal & "','mm/dd/yyyy') and to_date('" & IssueDate & "','mm/dd/yyyy') and h.fcba='" & SqlStr(v_txtfcba) & "' " & _ 
" )t " & _
" WHERE NVL(stock_api.get_StockSubGroup(t.STOCKCODE,t.fcba),'-') like '" & SqlStr(v_SubGroup) & "%' AND (STOCKCODE like '" & SqlStr(v_Cari) & "%' or stock_api.get_stockname(t.STOCKCODE, t.fcba) like '" & SqlStr(v_Cari) & "%') " & _  
" group by fcba,storecode,stockcode order by storecode,stockcode" 
Else
strStockOnHand		=	"select t.storecode,store_api.get_fcname(t.storecode, t.fcba) storename,t.STOCKCODE,stock_api.get_stockname(t.STOCKCODE, t.fcba) stockname,stock_api.get_uom(t.STOCKCODE, t.fcba) uom,stock_api.get_StockSubGroup(t.STOCKCODE,t.fcba) StockSubGroup, " & _ 
" sum(nvl(t.openingbalanceqty,0)) as openingbalanceqty,sum(nvl(t.receivenoteqty,0)) as receivenoteqty, " & _
" sum(nvl(t.issueqty,0)) as issueqty,sum(nvl(t.quantitytransfered,0)) as quantitytransfered,sum(nvl(t.quantityreceived,0)) as quantityreceived,sum(nvl(t.qtywriteoff,0)) as qtywriteoff, " & _
" (sum(nvl(t.openingbalanceqty,0))+sum(nvl(t.receivenoteqty,0))-sum(nvl(t.issueqty,0))-sum(nvl(t.quantitytransfered,0))+sum(nvl(t.quantityreceived,0))+sum(nvl(t.qtywriteoff,0))) as qty " & _
" from( " & _
    " select fcba,l.STORECODE,l.itemcode STOCKCODE, " & _
    " l.openingbalanceqty, " & _
    " 0 as receivenoteqty,0 as issueqty,0 as quantitytransfered,0 as quantityreceived,0 as qtywriteoff " & _
    " from stockledger l inner join bankstock k on (l.itemcode=k.fccode) " & _
    " where k.isactive='Y' AND l.YEAR='" & thnprevsplit & "' and l.MONTH='" & bulanprevsplit & "' and l.fcba='" & SqlStr(v_txtfcba) & "' " & _
    "      union all " & _
    " select r.fcba,r.storecode,rd.purchaseitemcode STOCKCODE, " & _
    " 0 as openingbalanceqty,rd.quantity as receivenoteqty,0 as issueqty,0 as quantitytransfered,0 as quantityreceived,0 as qtywriteoff " & _
    " from receivenote r,receivenotedetail rd " & _
    " where r.receivenotecode=rd.receivenotecode and r.rndate between to_date('" & tglawal & "','mm/dd/yyyy') and to_date('" & IssueDate & "','mm/dd/yyyy') " & _
    " and r.fcba='" & SqlStr(v_txtfcba) & "' " & _
    "      union all " & _
    " select h.fcba,h.storecode,d.STOCKCODE," & _
    " 0 as openingbalanceqty,0 as receivenoteqty,d.quantity as issueqty,0 as quantitytransfered, " & _
    " 0 as quantityreceived,0 as qtywriteoff " & _
    " from issue h,issuedetails d " & _
    " where h.sivcode=d.sivcode and h.fcba='" & SqlStr(v_txtfcba) & "' and h.sivdate between to_date('" & tglawal & "','mm/dd/yyyy') and to_date('" & IssueDate & "','mm/dd/yyyy') " & _
    "      union all " & _
    " select d.fcba,d.storecode,d.STOCKCODE, " & _
    " 0 as openingbalanceqty,0 as receivenoteqty,0 as issueqty,d.quantitytransfered,0 as quantityreceived,0 as qtywriteoff " & _
    " from transporttaskdtl_tab d,transporttask_tab h " & _
    " where h.orderno=d.orderno and h.rowstateid not in('Rejected') and h.fcba='" & SqlStr(v_txtfcba) & "' AND h.orderdate between to_date('" & tglawal & "','mm/dd/yyyy') and to_date('" & IssueDate & "','mm/dd/yyyy') " & _
    "     union all " & _
    " select d.fcba,d.locationcode,d.STOCKCODE, " & _
    " 0 as openingbalanceqty,0 as receivenoteqty,0 as issueqty,0 as quantitytransfered,d.quantityreceived,0 as qtywriteoff  " & _
    " from receivedtask_tab h,receivedtaskdtl_tab d " & _
    " where h.orderno=d.orderno and h.orderdate between to_date('" & tglawal & "','mm/dd/yyyy') and to_date('" & IssueDate & "','mm/dd/yyyy')  and h.fcba='" & SqlStr(v_txtfcba) & "' " & _
    "  UNION ALL " & _
    " select h.fcba,h.STORECODE locationcode,d.stockcode, " & _
    " 0 as openingbalanceqty,0 as receivenoteqty,0 as issueqty,0 as quantitytransfered, " & _
    " 0 as quantityreceived,d.quantity as qtywriteoff  " & _
    " from writeoff h,writeoffdetails d " & _
    " where h.writeoffcode=d.writeoffcode and h.wodate between to_date('" & tglawal & "','mm/dd/yyyy') and to_date('" & IssueDate & "','mm/dd/yyyy') and h.fcba='" & SqlStr(v_txtfcba) & "' " & _ 	
	"  UNION ALL " & _
    " select h.fcba,h.STORE locationcode,itemcode stockcode, " & _
    " 0 as openingbalanceqty,0 as receivenoteqty,0 as issueqty,0 as quantitytransfered, " & _
    " qtyreturn as quantityreceived,0 as qtywriteoff  " & _
    " from returnissuehdr h,returnissuedtl d " & _
    " where h.returnno=d.returnno and h.returndate between to_date('" & tglawal & "','mm/dd/yyyy') and to_date('" & IssueDate & "','mm/dd/yyyy') and h.fcba='" & SqlStr(v_txtfcba) & "' " & _ 
" )t " & _
" WHERE NVL(stock_api.get_StockSubGroup(t.STOCKCODE,t.fcba),'-') like '" & SqlStr(v_SubGroup) & "%' AND (STOCKCODE like '" & SqlStr(v_Cari) & "%' or UPPER(stock_api.get_stockname(t.STOCKCODE, t.fcba)) like UPPER('%" & SqlStr(v_Cari) & "%')) " & _  
" group by fcba,storecode,stockcode order by storecode,stockcode" 

end If
' response.write strStockOnHand
%>
<div id="wrapper">
<form name="frmAR" method="post" action="">
	<div class="ps-head">
		<div>
			<div class="ps-company">PT. <%=H(CompanyName)%></div>
			<h2 class="ps-title">Pilih Stock untuk Issue</h2>
			<div class="ps-sub">Stock on hand <%=H(FcbaName)%> &middot; klik <b>Select</b> atau klik dua kali baris untuk memilih (hanya stock dengan qty &gt; 0)</div>
		</div>
		<span class="ps-chip"><b>Per tanggal</b><%=H(IssueDate)%></span>
	</div>

	<div class="ps-filter">
		<div>
			<label for="subgroupcode">Sub Group</label>
			<select name="subgroupcode" id="subgroupcode">
				<option value="">-- Semua Sub Group --</option>
				<%
				Dim q,rsComboAFD
				q = "select fccode as id, fcname as text from stocksubgroup where fcba='" & SqlStr(v_txtfcba) & "'" & _
					"order by fccode asc"
				Set rsComboAFD = Server.CreateObject("ADODB.Recordset")
				rsComboAFD.CursorType = 3
				rsComboAFD.CursorLocation = 3
				rsComboAFD.Open q, connstr
				Do While Not rsComboAFD.eof
				%>
				<option value="<%=H(rsComboAFD("id"))%>"<%if rsComboAFD("id") & "" = v_SubGroup & "" then%> selected<%end if%>><%=H(rsComboAFD("id"))%> - <%=H(rsComboAFD("text"))%></option>
				<%
					rsComboAFD.movenext
				loop
				rsComboAFD.Close
				set rsComboAFD = Nothing
				q=""
				%>
			</select>
		</div>
		<div>
			<label for="txtcari">Kode / nama stock</label>
			<input type="text" maxlength="20" name="txtcari" id="txtcari" value="<%=H(v_Cari)%>" placeholder="Ketik kode atau nama stock" autocomplete="off">
		</div>
		<div>
			<input type="submit" name="btnview" id="btnview" value="Tampilkan" class="ps-btn">
		</div>
	</div>

	<div class="ps-panel">
		<div id="ps-tools"></div>
		<div class="ps-scroll">
		<table border="0" width="100%" id="product-table">
			<thead>
				<tr>
					<th data-nosort>Action</th>
					<th>Store Code</th>
					<th>Stock Code</th>
					<th>Stock Name</th>
					<th>UOM</th>
					<th>Quantity</th>
				</tr>
			</thead>
			<tbody>
<%
if IsPostBack then

	'response.write strStockOnHand

	Set rsGrnInv = Server.CreateObject("ADODB.Recordset")
	rsGrnInv.Open strStockOnHand, connstr, adOpenStatic

	if rsGrnInv.BOF and rsGrnInv.EOF then
%>
				<tr><td colspan="6" class="ps-empty">Tidak ada stock untuk filter ini.</td></tr>
<%	end if

	dim no, qtyCls
	no = 1
	while not rsGrnInv.eof

		'---- qty dari query (kolom qty), sama seperti tampilan lama
		qtyonhand   = rsGrnInv("qty")
		qtyCls = ""
		if cdbl(to_nol(qtyonhand)) = 0 then qtyCls = " zero"
		if cdbl(to_nol(qtyonhand)) < 0 then qtyCls = " neg"
%>
				<%'---- stock qty <= 0 tidak bisa dipilih (sama seperti tampilan lama)%>
				<tr<%if cdbl(to_nol(qtyonhand)) > 0 then%> ondblclick="getSelected(<%=no%>)"<%else%> class="ps-disabled"<%end if%>>
					<td align="center">
						<%if cdbl(to_nol(qtyonhand)) > 0 then%>
						<!-- type="button": tombol di dalam form tidak boleh ikut submit -->
						<button type="button" class="ps-select" onclick="getSelected(<%=no%>)">Select</button>
						<%else%><span class="ps-na" title="Qty tidak tersedia">-</span><%end if%>
					</td>
					<td><%=H(rsGrnInv("storecode"))%></td>
					<td class="ps-code"><%=H(rsGrnInv("stockcode"))%>
						<input type="hidden" id="itemcode<%=no%>" value="<%=H(rsGrnInv("stockcode"))%>">
					</td>
					<td class="ps-name"><%=H(rsGrnInv("stockname"))%>
						<input type="hidden" id="StockName<%=no%>" value="<%=H(rsGrnInv("stockname"))%>">
					</td>
					<td align="center"><%=H(rsGrnInv("uom"))%></td>
					<td class="ps-qty<%=qtyCls%>"><%=formatnumber(cdbl(to_nol(qtyonhand)),2)%>
						<input type="hidden" id="stockitem<%=no%>" value="<%=formatnumber(cdbl(to_nol(qtyonhand)),2)%>">
						<input type="hidden" id="StokBarang<%=no%>" value="<%=formatnumber(cdbl(to_nol(qtyonhand)),2)%>">
					</td>
				</tr>
<%
		no = no + 1
		rsGrnInv.MoveNext
	Wend
	rsGrnInv.Close
	set rsGrnInv = Nothing
else
%>
				<tr><td colspan="6" class="ps-empty">Pilih Sub Group dan/atau ketik kode / nama stock, lalu klik <b>Tampilkan</b>.</td></tr>
<%
end if
%>
			</tbody>
		</table>
		</div>
		<div id="ps-foot"></div>
	</div>
</form>
</div>
<script src="/js/jquery-1.9.1.js"></script>
<script>

	function getSelected(a) {
    console.log("Parent:", a);
    var itemcode = $('#itemcode' + a).val();
    var stockitem = $('#stockitem' + a).val();
    var StokBarang = $('#StokBarang' + a).val();
    var StockName = $('#StockName' + a).val();
    window.opener.parent.document.getElementById("itemcode").value = itemcode;
    window.opener.parent.document.getElementById("stockitem").value = stockitem;
    window.opener.parent.document.getElementById("StokBarang").innerHTML = StokBarang;
    window.opener.parent.document.getElementById("StockName").innerHTML = StockName;
    window.close();
      
	}

	// Tabel ala DataTables (cari, jumlah baris, urut kolom, pagination); hanya tampilan
	window.addEventListener('load', function () {
		var table = document.getElementById('product-table');
		if (!window.FbpTable || !table) return;
		FbpTable.init(table, { pageSize: 20 });
		var tools = table.querySelector('.fbt-toolbar-row .fbt-tools'), foot = table.parentNode.querySelector('.fbt-foot');
		if (tools) document.getElementById('ps-tools').appendChild(tools);
		if (foot) document.getElementById('ps-foot').appendChild(foot);
		var cari = document.getElementById('txtcari');
		if (cari && !cari.value) cari.focus();
	});
</script>
</body>
</html>
