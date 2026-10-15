<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
	Dim rsStockOnHand, strStockOnHand, rsCompany, strsql, rsFcba, strfcba
	Dim selectFields,strFieldCons,selectCons,whereCons,GroupCons,orderCons 
	Dim openingqty,receivenote,issue,transportask,receivetask,qtyonhand
	Dim NO,nTotal,v_txtcompany,v_txtfcba, v_txthi, v_txtitemcode, v_txtitemname,v_txtusercode
	Dim tglprevsplit,bulanprevsplit,thnprevsplit
	Dim xbln1,xthn1,xbln2,xthn2
	

	v_txtcompany          = Request("txtCompany")
	v_txtfcba             = Request("txtfcba")
	v_txthi               = Request("tgl")
	v_txtusercode         = Session("CodeName")
	v_txtStockGroup		=	trim(request.Form("txtStockGroup"))
	v_txtStock			=	trim(request.Form("txtStock"))

    If (v_txthi = "" Or Len(Trim(v_txthi)) = 0) Then v_txthi = textdate(date) End if
    tglblnawal  = bulansplit & "/" & "01" & "/" & thnsplit
    'tglbulanprev =  xbln2 & "/" & tglsplit & "/" & 

	
	
	tglbulanprev = v_txthi


	bulanprevsplit  = left(tglbulanprev,2)
    tglprevsplit    = mid(tglbulanprev,4,2)
    thnprevsplit    = right(tglbulanprev,4)

	'Response.Write tglbulanprev
	'Response.Write tglbulanprev

	
	strStockOnHand		=	"select t.storecode,t.storename,t.itemcode,t.stockname,t.uom,sum(t.openingbalanceqty) as openingbalanceqty,sum(t.receivenoteqty) as receivenoteqty, " & _
	" sum(t.issueqty) as issueqty,sum(t.quantitytransfered) as quantitytransfered,sum(t.quantityreceived) as quantityreceived,sum(t.qtywriteoff) as qtywriteoff,GroupCode, stockgroup_api.Get_StockGrpName(GroupCode,'" & v_txtfcba & "') GroupName" & _
	" from( " & _
    " select l.STORECODE,store_api.get_fcname(storecode, fcba) storename,l.itemcode, " & _
    " stock_api.get_stockname(itemcode, fcba) stockname,stock_api.get_uom(itemcode, fcba) uom,nvl(l.openingbalanceqty,0) openingbalanceqty, " & _
    " 0 as receivenoteqty,0 as issueqty,0 as quantitytransfered,0 as quantityreceived," & _ 
	" 0 as qtywriteoff,stock_api.get_StockGroup(itemcode,'" & v_txtfcba & "') GroupCode " & _
    " from stockledger l " & _
    " where l.YEAR='" & thnprevsplit & "' and l.MONTH='" & bulanprevsplit & "' and l.fcba='" & v_txtfcba & "' " & _
    "      union all " & _
    " select r.storecode,store_api.get_fcname(r.storecode, r.fcba) storename,rd.purchaseitemcode, " & _
    " stock_api.get_stockname(rd.purchaseitemcode, r.fcba) stockname, stock_api.get_uom(rd.purchaseitemcode, r.fcba) uom, " & _
    " 0 as openingbalanceqty,rd.quantity,0 as issueqty,0 as quantitytransfered,0 as quantityreceived," & _ 
	" 0 as qtywriteoff,stock_api.get_StockGroup(rd.purchaseitemcode,'" & v_txtfcba & "') GroupCode " & _
    " from receivenote r,receivenotedetail rd " & _
    " where r.receivenotecode=rd.receivenotecode and r.fcba=rd.fcba and r.rndate <=to_date('" & v_txthi & "','mm/dd/yyyy') and to_char(r.rndate,'yyyymm')=to_char(to_date('" & v_txthi & "','mm/dd/yyyy'),'yyyymm')" & _
    " and r.fcba='" & v_txtfcba & "' " & _
    "      union all " & _
    " select h.storecode,store_api.get_fcname(h.storecode, h.fcba) storename,d.STOCKCODE,stock_api.get_stockname(d.STOCKCODE, d.fcba) stockname, " & _
    " stock_api.get_uom(d.STOCKCODE, d.fcba) uom,0 as openingbalanceqty,0 as receivenoteqty,d.quantity,0 as quantitytransfered, " & _
    " 0 as quantityreceived,0 as qtywriteoff,stock_api.get_StockGroup(d.STOCKCODE,'" & v_txtfcba & "') GroupCode " & _
    " from issue h,issuedetails d " & _
    " where h.sivcode=d.sivcode and h.fcba=d.fcba and h.fcba='" & v_txtfcba & "' and h.sivdate<=to_date('" & v_txthi & "','mm/dd/yyyy') and to_char(h.sivdate,'yyyymm')=to_char(to_date('" & v_txthi & "','mm/dd/yyyy'),'yyyymm')" & _
    "      union all " & _
    " select d.storecode,store_api.get_fcname(d.storecode, d.fcba) storename,d.STOCKCODE,stock_api.get_stockname(d.STOCKCODE, d.fcba) stockname, " & _
    " stock_api.get_uom(d.STOCKCODE, d.fcba) uom,0 as openingbalanceqty,0 as receivenoteqty,0 as issueqty,d.quantitytransfered,0 as quantityreceived," & _ 
	" 0 as qtywriteoff,stock_api.get_StockGroup(d.STOCKCODE,'" & v_txtfcba & "') GroupCode " & _
    " from transporttaskdtl_tab d,transporttask_tab h " & _
    " where h.orderno=d.orderno and h.rowstateid not in('Rejected') and h.fcba=d.fcba and h.fcba='" & v_txtfcba & "' AND h.orderdate<=to_date('" & v_txthi & "','mm/dd/yyyy') and to_char(h.orderdate,'yyyymm')=to_char(to_date('" & v_txthi & "','mm/dd/yyyy'),'yyyymm')" & _
    "     union all " & _
    " select d.locationcode,store_api.get_fcname(d.locationcode, d.fcba) storename,d.STOCKCODE,stock_api.get_stockname(d.STOCKCODE, d.fcba) stockname, " & _
    " stock_api.get_uom(d.STOCKCODE, d.fcba) uom,0 as openingbalanceqty,0 as receivenoteqty,0 as issueqty,0 as quantitytransfered,d.quantityreceived," & _ 
	" 0 as qtywriteoff,stock_api.get_StockGroup(d.STOCKCODE,'" & v_txtfcba & "') GroupCode " & _
    " from receivedtask_tab h,receivedtaskdtl_tab d " & _
    " where h.orderno=d.orderno and h.fcba=d.fcba and h.orderdate<=to_date('" & v_txthi & "','mm/dd/yyyy')  and to_char(h.orderdate,'yyyymm')=to_char(to_date('" & v_txthi & "','mm/dd/yyyy'),'yyyymm') and h.fcba='" & v_txtfcba & "' " & _
	"  UNION ALL " & _
	" select h.storecode,store_api.get_fcname(h.storecode, d.fcba) storename,d.stockcode,stock_api.get_stockname(d.STOCKCODE, d.fcba) stockname, " & _
	" stock_api.get_uom(d.STOCKCODE, d.fcba) uom,0 as openingbalanceqty,0 as receivenoteqty,0 as issueqty,0 as quantitytransfered, " & _
	" 0 as quantityreceived,d.quantity as qtywriteoff,stock_api.get_StockGroup(d.STOCKCODE,'" & v_txtfcba & "') GroupCode  " & _
	" from writeoff h,writeoffdetails d " & _
	" where h.writeoffcode=d.writeoffcode and h.fcba=d.fcba and h.wodate<=to_date('" & v_txthi & "','mm/dd/yyyy') and to_char(h.wodate,'yyyymm')=to_char(to_date('" & v_txthi & "','mm/dd/yyyy'),'yyyymm') and h.fcba='" & v_txtfcba & "' " & _ 
	"  UNION ALL " & _
	" select h.store storecode,store_api.get_fcname(h.store, d.fcba) storename,d.itemcode stockcode,stock_api.get_stockname(d.itemcode, d.fcba) stockname, " & _
    " stock_api.get_uom(d.itemcode, d.fcba) uom,0 as openingbalanceqty,0 as receivenoteqty,0 as issueqty,0 as quantitytransfered, " & _
	" qtyreturn as quantityreceived,0 as qtywriteoff,stock_api.get_StockGroup(d.itemcode,'" & fcba & "') GroupCode  " & _
	" from returnissuehdr h,returnissuedtl d " & _
	" where h.returnno=d.returnno and h.fcba=d.fcba and h.returndate<=to_date('" & v_txthi & "','mm/dd/yyyy') and to_char(h.returndate,'yyyymm')=to_char(to_date('" & v_txthi & "','mm/dd/yyyy'),'yyyymm') and h.fcba='" & v_txtfcba & "' " & _ 				
	")t WHERE 1=1 " 


	if len(trim(v_txtStockGroup)) > 0 then
		strStockOnHand = strStockOnHand & " AND GroupCode = '" & v_txtStockGroup & "'"
	end if
	if len(trim(v_txtStock)) > 0 then
		strStockOnHand = strStockOnHand & " AND itemcode = '" & v_txtStock & "'"
	end if
						
	strGroup	=	" group by t.storecode,t.storename,t.itemcode,t.stockname,t.uom,GroupCode " 
	strOrder	=	" ORDER BY  t.stockname ASC"

	strStockOnHand	=	strStockOnHand + strGroup + strOrder

	'response.write strStockOnHand
	Set rsStockOnHand = Server.CreateObject("ADODB.Recordset")
	rsStockOnHand.Open strStockOnHand, connstr, adOpenStatic 
	
%>


<html>

<head>
	<title>ePlaS Reporting - Stock On Hand</title>
	<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
	<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
	<link rel="stylesheet" href="/css/jquery-ui.css" />
	<script src="/js/jquery-1.9.1.js"></script>
	<script src="/js/jquery-ui.js"></script>
	<link href="../plugins/tabcontent/tabcontent.css" rel="stylesheet" type="text/css" />
	<link href="../plugins/chosen/chosen.min.css" rel="stylesheet" />

	<link rel="stylesheet" href="../plugins/responsivegridsystem/col.css" media="all">
	<link rel="stylesheet" href="../plugins/responsivegridsystem/12cols.css" media="all">
	<script language="JavaScript" src='/js/x.js'></script>
	<script language="JavaScript" src='/js/xtableheaderfixed.js'></script>
	<script language="JavaScript">
	xAddEventListener(window, 'load',
	  function() {
		new xTableHeaderFixed('xthf-site', window);
	  }, false
	);
	</script>
	 <script>
	$(function() {
		$( "#tgl" ).datepicker();
	});
	</script>
	<script language="JavaScript" src="/js/date-picker.js"></script> 
	<script language="JavaScript">
	var newwin;
	function launchwin(winurl,winname,winfeatures)
	{
		//This launches a new window and then
		//focuses it if window.focus() is supported.
		newwin = window.open(winurl,winname,winfeatures);

	}
	</script>
	
	<style type="text/css">
	@media print
	{
	.printhide{display:none;}
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
			<form name="frmEntry" method="post">
			<div class="printhide">
				<table id="form-table" class="centered-margin">
				<tr>
				    <th colspan="6" style="background-color:#FFFFFF; border:none;">  
						<div id="header-wrapper"> 
							<div id="subheader" class="centered" tyle="width:50%; margin:auto;">
								<h2>STOCK ON HAND</h2>
							</div>
						</div>
					</th>
				</tr>
				<tr>
						<th>PT</th>
						<td>
					
						<select name="txtCompany" >
								<option value="">--PILIH COMPANY--</option>
								<%
								
								
								Set rsCompany = Server.CreateObject("ADODB.Recordset")
								rsCompany.CursorLocation = adUseClient
								strsql = " SELECT fccompanycode from user_company_tab  " & _
								         " where fcusercode='" & Session("CodeName") & "' group by fccompanycode" 								
								rsCompany.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
								if not rsCompany.eof then
									DO while not rsCompany.eof
										if v_txtCompany = rsCompany("fccompanycode") then
										%>			
											<option value="<%=rsCompany("fccompanycode")%>" Selected><%=rsCompany("fccompanycode")%></option>
										<% Else %>	
											<option value="<%=rsCompany("fccompanycode")%>"><%=rsCompany("fccompanycode")%></option>
										<% End If
										
										rsCompany.movenext
									loop
								End if
								rsCompany.Close
								set rsCompany = Nothing
								
								%>
								
						</select>
						</th>
						<th>Kebun/Pabrik</th>
						<td>
						<select name="txtfcba" >
								<option value="">--PILIH UNIT--</option>
								<%
								
								
								Set rsFcba = Server.CreateObject("ADODB.Recordset")
								rsFcba.CursorLocation = adUseClient
								strfcba = " SELECT fcba from user_fcba_tab  " & _
								         " where fcusercode='" & Session("CodeName") & "'" 								
										 rsFcba.Open strfcba, connStr, adOpenForwardOnly, adLockReadOnly	
								if not rsFcba.eof then
									DO while not rsFcba.eof
										if v_txtfcba = rsFcba("fcba") then
										%>			
											<option value="<%=rsFcba("fcba")%>" Selected><%=rsFcba("fcba")%></option>
										<% Else %>	
											<option value="<%=rsFcba("fcba")%>"><%=rsFcba("fcba")%></option>
										<% End If
										
										rsFcba.movenext
									loop
								End if
								rsFcba.Close
								set rsFcba = Nothing
								
								%>
								
						</select>
						</td>
						<th> Per Tanggal</th>
						<td>
						    
						  <input type="text" name="tgl" id="tgl" value="<%=v_txthi%>" size="20">
						</td>
				</tr>
				<tr>
						<th> Stock Group</th>
						<td>
						    <select name="txtStockGroup" class="chosen-select">	
									<option value="">--ALL--</option>
							<%
								Dim rsStockGroup
								Set rsStockGroup = Server.CreateObject("ADODB.Recordset")
								rsStockGroup.CursorLocation = adUseClient
								strsql =  " SELECT FCCODE, FCNAME FROM StockGroup WHERE FCBA = '" & Session("fcBA") & "' ORDER BY FCNAME ASC"
								rsStockGroup.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
								if not rsStockGroup.eof then
									DO while not rsStockGroup.eof
									if v_txtStockGroup = rsStockGroup("FCCODE") then
							%>			<option value="<%=rsStockGroup("FCCODE")%>" Selected><%=rsStockGroup("FCNAME")%></option>
									<%else%>
										<option value="<%=rsStockGroup("FCCODE")%>"><%=rsStockGroup("FCNAME")%></option>
									<%end if%>
							<%			rsStockGroup.movenext
									loop
								End if
								rsStockGroup.Close
								set rsStockGroup = nothing
							%>
							</select>	
						</td>   			 
						<th> Stock Code</th>
						<td>
						    <select name="txtStock" class="chosen-select" style="width: 50%">	
								<option value="">--ALL--</option>
							<%
								Dim rsStock
								Set rsStock = Server.CreateObject("ADODB.Recordset")
								rsStock.CursorLocation = adUseClient
								strsql =  " SELECT FCCODE, FCNAME FROM Stock WHERE FCBA = '" & Session("fcBA") & "' ORDER BY FCNAME ASC"
								rsStock.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
								if not rsStock.eof then
									DO while not rsStock.eof
									if v_txtStock = rsStock("FCCODE") then
							%>			<option value="<%=rsStock("FCCODE")%>" Selected><%=rsStock("FCCODE")%> : <%=rsStock("FCNAME")%></option>
									<%else%>
										<option value="<%=rsStock("FCCODE")%>"><%=rsStock("FCCODE")%> : <%=rsStock("FCNAME")%></option>
									<%end if%>
							<%			rsStock.movenext
									loop
								End if
								rsStock.Close
								set rsStock = nothing
							%>
							</select>		
						</td>   			 
						<td align="center"><input type="submit" value="View" /></td>
				</tr>
					<!--		<tr>
						<th> HARI INI</th>
						<th> ITEM CODE</th>
            			<th> ITEM NAME</th>
					</tr>
					<tr>
						<td>
							<input type="text" name="tgl" id="tgl" value="<%=TextDate(date())%>" size="20">
						</td>
				<td>
						    <input type="text" name="txtitemcode" value="<%=v_txtitemcode%>" size="30">
						</td>   			 
            <td>
                <input type="text" name="txtitemname" value="<%=v_txtitemname%>" size="30">
            </td>            
						<td align="center"><input type="submit" value="View" /></td>
                    </tr>-->    
				</table>
			</div>
			</form>
		</div>
		
		<table border="0" width="100%" id="product-table" class="xthf-site">
			<thead>
					<th>COMPANY</th>
					<th>UNIT</th>
					<th>STORE CODE</th>
					<th>STORE NAME</th>
					<th>STOCK GROUP </th>
					<th>STOCK CODE</th>
					<th>STOCK NAME</th>
					<th>UOM</th>
					<th>QUANTITY</th>
				</tr>
			</thead>
		<tbody>
<%if rsStockOnHand.eof and rsStockOnHand.eof then%>
	<tr>
		<td colspan="<%=6 + var_colspan%>" align="left">NO DATA AVAILABLE</td>
	</tr>
<%end if
		NO = 1
		nTotal = 0
		
		while not rsStockOnHand.eof
		openingqty  = rsStockOnHand("openingbalanceqty")
		receivenote = rsStockOnHand("receivenoteqty")
		issue		= rsStockOnHand("issueqty")
		transportask= rsStockOnHand("quantitytransfered")
		receivetask	= rsStockOnHand("quantityreceived")
		qtywriteoff	= rsStockOnHand("qtywriteoff")
		if openingqty = "" then
			openingqty = "0"
		end if
		if receivenote = "" then
			receivenote = "0"
		end if
		if issue = "" then
			issue = "0"
		end if
		if transportask = "" then
			transportask = "0"
		end if
		if receivetask = "" then
			receivetask = "0"
		end if
		if qtywriteoff = "" then
			qtywriteoff = "0"
		end if
		if qtyonhand = "" then
			qtyonhand = "0"
		end if
		qtyonhand	= cdbl(openingqty)+cdbl(receivenote)-cdbl(issue)-cdbl(transportask)+cdbl(receivetask)+cdbl(qtywriteoff)
		
%>
			<tr class="<%=cClass%>">
				 <td align="left"><%=v_txtcompany%></td>
				<td align="left"><%=v_txtfcba%></td>
				<td align="left"><%=rsStockOnHand("STORECODE")%></td>
				<td align="left"><%=rsStockOnHand("STORENAME")%></td>
				<td align="left"><%=rsStockOnHand("GroupName")%></td>
				<td align="left"><%=rsStockOnHand("ITEMCODE")%></td>
				<td align="left"><%=rsStockOnHand("STOCKNAME")%></td>
				<td align="left"><%=rsStockOnHand("UOM")%></td>
				<td align="right"><%=formatnumber(cdbl(to_nol(qtyonhand)),2)%></td>				
			</tr>
<%
			NO = NO + 1
			nTotal = nTotal + Cdbl(qtyonhand)
			rsStockOnHand.MoveNext
		Wend
	rsStockOnHand.Close
	set rsStockOnHand = Nothing
%>
			<tr style="background-color:#F4F4A0;">
				<td align="center" colspan="<%=8 + var_colspan%>"><strong>TOTAL</strong></td>
				<td align="right"><strong><%=FbpNum2(nTotal)%></strong></td>
			</tr>
 </tbody>
</table>
   <table width="90%" class="data" style="margin:auto;">
   	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr>
 		<td width="50%" align="center"><font face="Verdana" size="1">Approved by,</font></td>
		<td width="50%" align="center"><font face="Verdana" size="1">Prepared by,</font></td>
	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr> 
		<td align="center"><font face="Verdana" size="2"><strong><u>&nbsp;&nbsp;<%=Session("FCMANAGER")%>&nbsp;&nbsp;</u></strong></td>
		<td align="center"><font face="Verdana" size="2"><strong><u>&nbsp;&nbsp;<%=Session("KTU")%>&nbsp;&nbsp;</u></strong></td>
	   </tr>
	   	  <tr> 
		<td align="center"><font face="Verdana" size="1"><strong><i>MANAGER</i></strong></td>
		<td align="center"><font face="Verdana" size="1"><strong><i>KTU</i></strong></td>
	   </tr>
   </table>
  <script src="../js/jquery.inputmask.bundle.min.js"></script>
  <script src="../js/inputmask.binding.js"></script>
  <script src="../plugins/chosen/chosen.jquery.min.js"></script>
  <script type="text/javascript"> 
    // $(".chosen-select").empty();
    
    var config = {
      '.chosen-select'           : {},
      '.chosen-select-deselect'  : {allow_single_deselect:true},
      '.chosen-select-no-single' : {disable_search_threshold:10},
      '.chosen-select-no-results': {no_results_text:'Oops, nothing found!'},
      '.chosen-select-width'     : {width:"45%"}
    }
    for (var selector in config) {
      $(selector).chosen(config[selector]);
    }
    // $(".chosen-select").trigger("chosen:updated");

  </script>
</div>