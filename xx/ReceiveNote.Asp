<%
	response.buffer = True
	validatesession
%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="mm.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
	<title>ePlaS - Receive Note</title>
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
<script language="JavaScript">

function getSto(sel){
	removeOptions(document.getElementById("StockTransNo2"));
	removeOptions(document.getElementById("DONo2"));
	//alert(sel.options[sel.selectedIndex].text);
	var url = 'ajax/getstonumber.asp';
	var txttrackno = sel.options[sel.selectedIndex].text;	
	var fcba = '<%=Session("fcBA") %>';	
	$.ajax({
		url:url,
		type:'GET',
		dataType: 'json',
		data: jQuery.param({fcbasource: fcba, txttrackno: txttrackno}),
		contentType: 'application/json; charset=UTF-8',
		success: function( json ) {
		    $('#StockTransNo2').append($('<option>').text('').attr('value', null));
			$.each(json, function(i, value) {
				$('#StockTransNo2').append($('<option>').text(value.STONO).attr('value', value.STONO));
			});
		}
	});
}

function getdono(sel){
	removeOptions(document.getElementById("DONo2"));
	//alert(sel.options[sel.selectedIndex].text);
	var url = 'ajax/getdononumber.asp';
	var txttrackno = document.getElementById("xCode").value;	
	var txtstono = document.getElementById("StockTransNo2").value;
	var fcba = '<%=Session("fcBA") %>';
    //alert(fcba + ' , ' + txttrackno + ' , ' + txtstono) ;
	$.ajax({
		url:url,
		type:'GET',
		dataType: 'json',
		data: jQuery.param({fcbasource: fcba, txttrackno: txttrackno , txtstono: txtstono}),
		contentType: 'application/json; charset=UTF-8',
		success: function( json ) {
		    $('#DONo2').append($('<option>').text('').attr('value', null));
			$.each(json, function(i, value) {
				$('#DONo2').append($('<option>').text(value.DONO).attr('value', value.DONO));
			});
		}
	});
	
}

function removeOptions(selectbox)
{
    var i;
    for(i = selectbox.options.length - 1 ; i >= 0 ; i--)
    {
        selectbox.remove(i);
    }
}

function getstonumber(str){

	var xmlhttp;
	if (str=="")
	  {
	  document.getElementById("StockTransNo").innerHTML="";
	  document.getElementById("receivenote").innerHTML="";
	  return;
	  }
	if (window.XMLHttpRequest)
	  {// code for IE7+, Firefox, Chrome, Opera, Safari
	  xmlhttp=new XMLHttpRequest();
	  StockName=new XMLHttpRequest();
	  }
	else
	  {xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");}
	xmlhttp.onreadystatechange=function()
	  {
	  if (xmlhttp.readyState==4 && xmlhttp.status==200)
		{
		//alert(xmlhttp.responseText);
		document.getElementById("StockTransNo").innerHTML=xmlhttp.responseText;
		document.getElementById("receivenote").innerHTML="";
		}
	  }
	xmlhttp.open("GET","ajax/getstonumber.asp?fcbasource="+str+"&fcbaDistination=<%=Session("fcBA")%>",true);
	xmlhttp.send();
}


xAddEventListener(window, 'load',
  function() {
    new xTableHeaderFixed('xthf-site', window);
    new xTableHeaderFixed('xthf-green', window);
    new xTableHeaderFixed('xthf-blue', window);
  }, false
);
</script>
 <script>
$(function() {
$( ".datepicker" ).datepicker();
});
</script>
<!--<script language="JavaScript" src="/js/date-picker.js"></script> -->
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}

function harusangka(evt) {
    var e = evt || window.event;
    var key = e.keyCode || e.which;

    if (!e.shiftKey && !e.altKey && !e.ctrlKey &&
    // numbers   
    key >= 48 && key <= 57 ||
	e.ctrlKey && key == 97 ||
    // Backspace and Tab and Enter
    key == 8 || key == 9 || key == 13 ||
    // left and right arrows
    key == 37 || key == 39 ||
    // Del and Ins
	key == 46 ) {
        // input is VALID
		}
	else {
			// input is INVALID
			e.returnValue = false;
			if (e.preventDefault) e.preventDefault();
		}
}

</script>
</head>
<body>
	<!--#include file="top.jsp"-->
	<div style=" width:21.7%; float:left;margin-left:0.3%; margin-top:0.3%; margin-bottom:0.3%;">
		<% if Session("fcBa") <> "" then %>
         <!--#include file="menu.asp"-->
        <% end if %>
	</div>
    <div class="judul_head" valign="top" bgcolor="#fff" style=" width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;"> 
    <div class="judul">
	   Module : <%=request("txtModule")%> / <%=request("txtTable")%>
	</div>
    <div class="box_both" style="overflow:scroll;">
	
<%
Function cmdCheckHistNursery(initfcba, initNURSERYCODE, initLineNoHist)
	dim qCheckHistNursery,rsCheckHistNursery,qTotalDataNursery
	
	qCheckHistNursery = " select count(1) totalData" & _
						" from stocktranshist" & _
						" where fcba = '" & initfcba & "'" & _
						" and targetcode = '" & initNURSERYCODE & "'" & _
						" and lineno > " & initLineNoHist & ""
						
	'Response.Write qCheckHistNursery

	set rsCheckHistNursery = Server.CreateObject("ADODB.Recordset")
	rsCheckHistNursery.Open qCheckHistNursery, connstr
	if not rsCheckHistNursery.eof then
		qTotalDataNursery = cdbl(rsCheckHistNursery("totalData"))
	else
		qTotalDataNursery = 0
	end if 	

	cmdCheckHistNursery = qTotalDataNursery
End Function
%>
		<%
			'--Awal--
			Dim xCode,xstoreCode,fcCode,cmdView,DATECREATED,VoucherCode,QUANTITY,AMOUNT,OTHERITEMDESC,v_CrossCekPO
			Dim Remarks
			Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis,TOTALAMOUNT,HeaderRemarks
			Dim Jobcode, LocationType, LocationCode, Reference, Remark
			Dim qLineMax, rsqLineMax, tempMaxLine
			Dim qTotalRecordOnHand, rsTotalRecordOnHand, tempTotalRecordOnHand
			Dim qQuantityOnHand, rsQuantityOnHand, tempQuantityOnHand
			Dim sqlSQLCheck, rsSQL, qtyRemained, qtyDoRemaind
			Dim qCheckStore, rsqCheckStore
			Dim txtLocationType	

			Dim txtGRNCode : txtGRNCode = 0
			Dim oRs : oRs = 0
			Dim oSql : oSql = 0
			Dim txtLocationCode : txtLocationCode = ""
			Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""
			dim StockTransNo2,DONo2,IsUseExtSys

			

			txtTable =Request("txtTable")
			txtModule=Request("txtModule")
			txtSearch=Request("txtSearch")
			txtField =Request("txtField")
			
			if request("fccOde")="" then
				fcCode="" '----- Nantinya Ambil dari Format Master
			else
				fcCode = request("fccOde")
			end if
			
			If Len(Trim(fcCode)) = 0 Then fcCode = "(NEW)"
			
			txtJenis = request("txtJenis") '-- Untuk Membedakan Jenis Data Yang akan diolah [insert,edit,delete]
			
			'---- Data Header
			if Request("DATECREATED") <> "" then
				DATECREATED	=	Request("DATECREATED")
			else
				DATECREATED	=	textDate(date)
			end if
			DATECREATED = FormatDateTime(DATECREATED)
			
			
			xCode = Request("xCode")
			xstoreCode = Request("xstoreCode")
			StockTransNo2 = Request("StockTransNo2")
			DONo2 = Request("DONo2")
			
			HeaderRemarks = Request("HeaderRemarks")
			
			'---- Data Details
			JobCode=Request("JobCode")
			LocationType=Request("LocationType")
			LocationCode=Request("LocationCode")
			Reference=Request("Reference")
			Remarks=Request("Remarks")
			Amount=Request("Amount")
			PurchaseLineNoKey = Request("PurchaseLineNoKey")
			
			
			
			
			
			'--- Action ke Database di Oracle [ insert,edit,delete ]
			if Request("txtJenis")="New" or Request("txtJenis")= "Header" or Request("txtJenis")= "Details" or Request("txtJenis")="Edit" Then
				REMARKS	=	request("REMARKS")
				
				    Set rsPO = Server.CreateObject("ADODB.Recordset")
					rsPO.Open " select company_api.IsUseExtSys(businessunit_api.get_CompanyCode('" & Session("fcBA") & "')) IsUseExtSys from dual" , connstr
					IF Not rsPO.eof then
						IsUseExtSys =  rsPO("IsUseExtSys")
					End if
					rsPO.Close
					set rsPO = Nothing
				
				if Request("txtJenis")= "Header" then
					'Simpan NO PO nya terlebih dahulu sebelum di delete
					connstr.BeginTrans
					
					Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), DATECREATED)
					If Cek_IsDailyClosing = "FALSE" then

						Set rsPO = Server.CreateObject("ADODB.Recordset")
						rsPO.Open " select pocode" & _
								  " from receivenote" & _
								  " where receivenotecode = '" & fcCode & "'" & _
								  "       and fcBA = '" & Session("fcBA") & "'", connstr
						IF Not rsPO.eof then
							txtNoPO =  rsPO("POCODE")
						End if
						rsPO.Close
						set rsPO = Nothing

						q = " delete from receivenote" & _
							" where receivenotecode = '" & fcCode & "'" & _
							"       and fcba = '" & Session("fcBA") & "'"
						
						connStr.execute(q)
						fcCode = "(NEW)"
						
						if err.number = 0 then
							connStr.CommitTrans
						end if
						
					else
						err.number = 1001
						err.description = "Tanggal """ & DATECREATED & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
					End If
					
					cmdView	= "True"
				elseif Request("txtJenis")= "Details" then
					'Simpan NO PO nya terlebih dahulu sebelum di delete
					connstr.BeginTrans
					
					Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), DATECREATED)
					If Cek_IsDailyClosing = "FALSE" then
						
						Set rsPO = Server.CreateObject("ADODB.Recordset")
						rsPO.Open "Select POCODE,DELIVERYORDERCODE,StockOrderNo, STORECODE from ReceiveNote where RECEIVENOTECODE='" & fcCode & "'" & _
						" and fcBA='" & Session("fcBA") & "'", connstr
						IF Not rsPO.eof then
							txtNoPO =  rsPO("POCODE")
							StockTransNo2 =  rsPO("StockOrderNo")
							DONo2 =  rsPO("DELIVERYORDERCODE")
							txtStoreCode = rsPO("STORECODE")
						End if
						rsPO.Close
						set rsPO = Nothing
						
						q = " Select LineNoHist" & _
							"       , quantity" & _
							"       , PURCHASEITEMCODE" & _
							"       , PURCHASELINENOKEY" & _
							"       , LOCATIONTYPE" & _
							"       , LOCATIONCODE" & _
							" from ReceiveNoteDetail" & _
							" where RECEIVENOTECODE = '" & fcCode & "'" & _
							"       and PURCHASELINENOKEY = '" & Request("PURCHASELINENOKEY") & "'" & _
							"       and fcBA = '" & Session("fcBA") & "'"
						
						Set rsReceiveDetail = Server.CreateObject("ADODB.Recordset")
						rsReceiveDetail.Open q, connstr
								
						IF Not rsReceiveDetail.eof then
							txtLineNoHist = rsReceiveDetail("LINENOHIST")
							originalQuantity = CDbl(rsReceiveDetail("QUANTITY"))
							txtItemCodeDelete = rsReceiveDetail("PURCHASEITEMCODE")
							txtItemLineKeyDelete = rsReceiveDetail("PURCHASELINENOKEY")
							txtLocationType = rsReceiveDetail("LOCATIONTYPE")
							txtLocationCode = rsReceiveDetail("LOCATIONCODE")
						End if
						rsReceiveDetail.Close
						set rsReceiveDetail = Nothing					
						
						
						
						' Add by NAF, 20 Jan 2014
						' To handle Check Store for CAPEX
						qCheckStore = " select LOCATIONTYPE_API.get_item('" & txtLocationType & "') item from dual "

						set rsqCheckStore = Server.CreateObject("ADODB.Recordset")
						rsqCheckStore.Open qCheckStore, connstr
						if not rsqCheckStore.eof then
							tempItem = rsqCheckStore("item")
						else
							tempItem = ""
						end if 
						rsqCheckStore.Close
						Set rsqCheckStore = Nothing
						' End of modification
									
						
						if tempItem = "MATERIAL" then
						
							If txtLocationType = "MN" Then
								v_CrossCekDetail = cmdCheckHistNursery(Session("fcBA"), txtLocationCode, txtLineNoHist)
					
									qLessZero = " select islesszero" & _
												" from stock" & _
												" where fccode = '" & Request("ItemCode") & "'" & _
												"       and fcba = '" & session("fcBA") & "'" 
									
									set rsLessZero = Server.CreateObject("ADODB.Recordset")
									rsLessZero.Open qLessZero, connstr
									if not rsLessZero.eof then
										tempislesszero = rsLessZero("islesszero")
									else
										tempislesszero = "FALSE"
									end if
									
									qQuantityOnHand = " select quantity" & _
													   " from stockonhandnursery" & _
													   " where storecode = '" & txtStoreCode & "'" & _
													   "       and stockcode  = '" & Request("ItemCode") & "'" & _
													   "       and nurserycode  = '" & txtLocationCode & "'" & _
													   "       and fcba = '" & session("fcBA") &"' "
													  
									set rsQuantityOnHand = Server.CreateObject("ADODB.Recordset")
									rsQuantityOnHand.Open qQuantityOnHand, connstr
									if not rsQuantityOnHand.eof then
										QuantityOnHandNursery = cdbl(rsQuantityOnHand("quantity"))
									else
										QuantityOnHandNursery = 0
									end if
									
									If tempislesszero = "TRUE" Or (tempislesszero = "FALSE" And ((QuantityOnHandNursery - originalQuantity) >= 0)) Then
									
										q =	" Delete from stocktranshist" & _
											" where lineno = " & txtLineNoHist & "" & _
											"       and fcba = '" & Session("fcBA") & "'"

										connStr.execute(q)

										q = " update stockonhand" & _
											" set quantity = quantity - " & originalQuantity & "" & _
											"     , lastupdate = sysdate " & _
											"     , lasttime = to_char(sysdate,'hh24:mi') " & _
											" where storecode = '" & txtStoreCode & "'" & _
											"       and stockcode = '" & Request("ItemCode") & "'" & _
											"       and fcba = '" & session("fcBA") &"'"

										connStr.execute(q)

										q = " update stockonhandnursery " & _
											" set quantity = quantity - " & originalQuantity & "" & _
											"     , lastupdate = sysdate " & _
											"     , lasttime = to_char(sysdate,'hh24:mi') " & _
											" where storecode = '" & txtStoreCode & "'" & _
											"       and stockcode  = '" & Request("ItemCode") & "'" & _ 
											"       and nurserycode  = '" & txtLocationCode & "'" & _ 
											"       and fcba = '" & session("fcBA") &"'"

										connStr.execute(q)
										
										If IsUseExtSys = "Y"  Then
											q = "BEGIN BridgingSAP_PRPOGRN_API.DeleteGRNLineNoToBridge('" & Session("fcBA") & "','" & fcCode & "','" & PurchaseLineNoKey & "'); END; "
											connStr.execute(q)
										End if
										
										q = " delete from receivenotedetail" & _
											" where receivenotecode = '" & fcCode & "'" & _
											"       and purchaselinenokey = '" & PurchaseLineNoKey & "'" & _
											"       and fcba = '" & Session("fcBA") & "'"
										
										connStr.execute(q)
										
										q = " update purchaseorderdetails" & _
											" set rowstateid = 'Opened'" & _
											" where pocode = '" & txtNoPo & "'" & _
											"       and linenokey = '" & txtItemLineKeyDelete & "'" & _
											"       and fcBA = '" & Session("fcBA") & "'"

										connStr.execute(q)
										
										v_CrossCekPO =  cmdCrossCekPOUpdate (txtNoPo, txtItemLineKeyDelete,  originalQuantity )
										
										if err.number=0 then
											connStr.CommitTrans
										End if
										
									Else
										err.number = 1001
										Err.Description = "Quantity tidak boleh Negatif atau Saldo belum didefinisikan"
									End If
								
							Else
							
								v_CrossCekDetail = cmdCheckHist(txtLineNoHist, Request("ItemCode"))
								
								' Add new function to update qty received of PO Detail, by NAF
								v_CrossCekPO =  cmdCrossCekPOUpdate (txtNoPo, txtItemLineKeyDelete,  originalQuantity ) 
								' End of Modification
								
								qLessZero = " select islesszero" & _
											" from stock" & _
											" where fccode = '" & Request("ItemCode") & "'" & _
											"       and fcba = '" & session("fcBA") & "'" 
								
								set rsLessZero = Server.CreateObject("ADODB.Recordset")
								rsLessZero.Open qLessZero, connstr
								if not rsLessZero.eof then
									tempislesszero = rsLessZero("islesszero")
								else
									tempislesszero = "FALSE"
								end if
								
								qQuantityOnHand = " select quantity" & _
												   " from stockonhand" & _
												   " where storecode = '" & txtStoreCode & "'" & _
												   "       and stockcode  = '" & Request("ItemCode") & "'" & _
												   "       and fcba = '" & session("fcBA") &"' "
												  
								set rsQuantityOnHand = Server.CreateObject("ADODB.Recordset")
								rsQuantityOnHand.Open qQuantityOnHand, connstr
								if not rsQuantityOnHand.eof then
									QuantityOnHandStock = cdbl(rsQuantityOnHand("quantity"))
								else
									QuantityOnHandStock = 0
								end if
								
								If tempislesszero = "TRUE" Or (tempislesszero = "FALSE" And ((QuantityOnHandStock - originalQuantity) >= 0)) Then
									
									q =  "Delete from stocktranshist " & _
										 "where " & _
										 "lineno = " & txtLineNoHist & " and " & _
										 "fcba = '" & Session("fcBA") & "'"
								
									connStr.execute(q)
									
									q = " update stockonhand " & _
										" set " & _
										" quantity = quantity - " & originalQuantity & " " & _
										", LASTUPDATE = sysdate " & _
										", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
										"where " & _
										"storecode = '" & txtStoreCode & "' and " & _
										"stockcode  = '" & txtItemCodeDelete & "' and " & _ 
										"fcba = '" & session("fcBA") &"' "
										
									connStr.execute(q)
									
									'add By Akbar 11 May 2018
									
									If IsUseExtSys = "Y"  Then
									
										q = "Update STOCKTRANSFERORDER sto  set sto.dorcvqty = sto.dorcvqty  - " & originalQuantity & _
											 ", rowstateid = 'Released', LASTUPDATE = sysdate " & _
											 ", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
											 "	 where sto.trackno = '" & txtNoPo & "' " & _ 
											 "	 and sto.tracklineno = '" & txtItemLineKeyDelete & "'" & _ 
											 "	 and sto.stono =   '" & StockTransNo2 & "' " & _  
											 "	 and sto.dono =  '" & DONo2 & "' " & _  
											 "	 and sto.fcba =  '" & session("fcBA") & "'" 
										'response.write	 q
										connStr.execute(q)	

										strQueryStatus = "BEGIN BridgingSAP_PRPOGRN_API.DeleteGRNLineNoToBridge('" & Session("fcBA") & "','" & fcCode & "','" & PurchaseLineNoKey & "'); END; "
										connStr.execute(strQueryStatus)
									
										q="Delete from RECEIVENOTEDETAIL where  RECEIVENOTECODE='" & fcCode & "' and PURCHASELINENOKEY='" & PurchaseLineNoKey & _
										"' and fcBA='" & Session("fcBA") & "'"
										
										connStr.execute(q)
									end if	
									
									q = " update purchaseorderdetails" & _
										" set rowstateid = 'Opened'" & _
										" where pocode = '" & txtNoPo & "'" & _
										"       and linenokey = '" & txtItemLineKeyDelete & "'" & _
										"       and fcBA = '" & Session("fcBA") & "'"

									connStr.execute(q)
								
									if err.number=0 then
										connStr.CommitTrans
									End if
									
								Else
									err.number = 1001
									Err.Description = "Quantity tidak boleh Negatif atau Saldo belum didefinisikan"
								End If
								
							End If
							
						else
							' Add new function to update qty received of PO Detail, by NAF
							v_CrossCekPO =  cmdCrossCekPOUpdate (txtNoPo, txtItemLineKeyDelete,  originalQuantity ) 
							' End of Modification
							
							'add By Akbar 11 May 2018
							If IsUseExtSys = "Y"  Then
								 q = "Update STOCKTRANSFERORDER sto  set sto.dorcvqty = sto.dorcvqty - " & originalQuantity & _
									 ", rowstateid = 'Released', LASTUPDATE = sysdate " & _
									 ", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
									 "	 where sto.trackno = '" & txtNoPo & "' " & _ 
									 "	 and sto.tracklineno = '" & txtItemLineKeyDelete & "'" & _ 
									 "	 and sto.stono =   '" & StockTransNo2 & "' " & _  
									 "	 and sto.dono =  '" & DONo2 & "' " & _  
									 "	 and sto.fcba =  '" & session("fcBA") & "'" 
								connStr.execute(q)	
								
								strQueryStatus = "BEGIN BridgingSAP_PRPOGRN_API.DeleteGRNLineNoToBridge('" & Session("fcBA") & "','" & fcCode & "','" & PurchaseLineNoKey & "'); END; "
								connStr.execute(strQueryStatus)
							End If
							
							
							q="Delete from RECEIVENOTEDETAIL where  RECEIVENOTECODE='" & fcCode & "' and PURCHASELINENOKEY='" & PurchaseLineNoKey & _
									"' and fcBA='" & Session("fcBA") & "'"
									
							connStr.execute(q)
						
							if err.number=0 then
								connStr.CommitTrans
							End if
								
						end if
						
						
					else
						err.number = 1001
						err.description = "Tanggal """ & DATECREATED & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
					End If
					
					cmdView	= "True"
					
				elseif Request("txtJenis")="New" then
					connstr.BeginTrans
					
					Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), DATECREATED)
					If Cek_IsDailyClosing = "FALSE" then
					
					'--Add by Ukie to handle Auto Number of PR
					If Ucase(fcCode) = Ucase("(NEW)") Then
						oSql = " select nvl(controlnumber_api.get_nextnumber(upper('receivenote')" & _
							   "                                         , " & year(textDate(DATECREATED)) & "" & _
							   "                                         , " & month(textDate(DATECREATED)) & "" & _
							   "                                         , '" & Session("fcBA") & "'" & _
							   " ), 'FALSE') grncode from dual"
						
						'Response.Write oSql
							   
						Set oRs = Server.CreateObject("ADODB.Recordset")
						oRs.Open oSql, connstr, 3, 3
						If Not oRs.EOF Then
							txtGRNCode = oRs("grncode")
						End If
						oRs.Close
						Set oRs = Nothing
						fcCode = txtGRNCode
					End If
					
					If txtGRNCode = "FALSE" Then
						err.Number = 1001
						err.description = "<br/>Automatic Number Not Have Been Set, Call your Administrator!"
					End If
					
					'Response.Write fcCode & txtGRNCode
					'--End Added
					
					
					Dim rsPO
					  Set rsPO = Server.CreateObject("ADODB.Recordset")
					  rsPO.Open "Select RECEIVENOTECODE from ReceiveNote where RECEIVENOTECODE='" & fcCode & "'" & _
								" and fcBA='" & Session("fcBA") & "'", connstr
					  IF rsPO.eof then	    
						 Dim txtCodeSupplier
						 q="Select SUPPLIERCODE From PurchaseOrder where Pocode ='" & xCode & "'"  & _
						   " and fcBA='" & Session("fcBA") & "'"
						 Dim rsLPO  
						 set rsLPO = Server.CreateObject("ADODB.Recordset")
						 rsLPO.Open q,connstr,1,1
						 if not rsLPO.eof then
							txtCodeSupplier=rsLPO("SupplierCode")
						 End if 
						 rsLPO.Close
						 set rsLPO = Nothing
						 q="Insert into ReceiveNote (RECEIVENOTECODE, RNDATE, POCODE,STOCKORDERNO,DELIVERYORDERCODE,SupplierCode, Storecode, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
						 "values ('" & fcCode & "'," & " to_date('" & DATECREATED & "','mm/dd/yyyy'),'" & xCode & "','" & StockTransNo2 & "','" & DONo2 & "','" & txtCodeSupplier & "','" & _
						 xstoreCode & "','" & _
						 Session("CodeName") & "','" & Session("CodeName") & "','" & _
						 Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI'))"
						 connStr.execute(q)
						 'response.Write(q)
					  End if
					  rsPO.Close	
					  Dim n
	  
						For n=1 to (cdbl(request("no"))-1) 
							if request("LT_"&TRIM(n))<>"" and request("LC_"&TRIM(n))<>"" and request("QtyReceived_"&TRIM(n))<>"" then
								 'Cari Control Job Stock Code 
								 if  request("JobCode_"&TRIM(n))="" Then
									 Dim txtControlJOb
									 q="select sg.CONTROLJOB as CONTROLJOB  from stockgroup sg,stock s" & _
									   " where (sg.FCCODE = s.STOCKGROUPCODE and sg.fcba = s.fcba) " & _
									   " and s.FCCODE='" & request("Stock_"&TRIM(n)) & "'" & _
									   " and s.fcBA='" & Session("fcBA") & "'"
									 set rsStock = Server.CreateObject("ADODB.Recordset")
									 rsStock.Open q, connstr
									 if not rsStock.eof then
										txtControlJob=rsStock("CONTROLJOB")
									 End if 
									 'rsStock.Close
									 'set rsStock = Nothing
								 else
									txtControlJob=request("JobCode_"&TRIM(n))
								 end if
								 
								if  request("AMOUNTReceived_"&TRIM(n))="" Then
									Dim txtamountreceived,txtQtyPO
									
									q = " select p.linenokey" & _
										"        , p.itemcode" & _
										"        , s.fcname" & _
										"        , sum(p.quantity) as quantity" & _
										"        , sum(amount) as amount" & _
										" from purchaseorderdetails p" & _
										"      , stock s" & _
										" where (p.itemcode = s.fccode and p.fcba = s.fcba)" & _
										"       and pocode = '" & xCode & "'" & _
										"       and s.fccode = '" & request("Stock_"&TRIM(n)) & "'" & _
										"       and p.linenokey = '" & request("LineNoKey_"&TRIM(n)) & "'" & _
										"       and p.fcba = '" & session("fcBA") & "'" & _
										" group by p.linenokey" & _
										"          , p.itemcode" & _
										"          , s.fcname"

									set rsamr = Server.CreateObject("ADODB.Recordset")
									rsamr.Open q, connstr
									
									if not rsamr.eof then
										txtamountreceived = Round((cdbl(rsamr("Amount"))/cdbl(rsamr("Quantity")))*request("QtyReceived_"&TRIM(n)), 0)
										PurchaseLineNoKey = rsamr("lineNoKey")
										txtQtyPO = cdbl(rsamr("Quantity"))
										'response.Write  PurchaseLineNoKey & ". rsamr(Amount) -> " & rsamr("Amount") & " , rsamr(Quantity) -> " & rsamr("Quantity") 
									End if 
									rsamr.close
									set rsamr = nothing	
									
									
									q = " select s.stono, " & _
										"		 s.dono,  " & _
										"		 s.trackno," & _
										"		 s.tracklineno," & _
										"		 s.material, " & _
										"		 s.doqty," & _
										"		 s.netprice " & _
										"   from stocktransferorder s " & _
										"  where s.trackno = '" & xCode & "'" & _
										"	and s.tracklineno = '" & request("LineNoKey_"&TRIM(n)) & "'" & _
										"	and s.material = '" & request("Stock_"&TRIM(n)) & "'" & _
										"	 and s.stono =   '" & StockTransNo2 & "' " & _  
										"	 and s.dono =  '" & DONo2 & "' " & _  
										"	 and s.fcba =  '" & session("fcBA") & "'" 

									set rsamr = Server.CreateObject("ADODB.Recordset")
									rsamr.Open q, connstr
									if not rsamr.eof then
										txtamountreceived = cdbl(txtamountreceived) + Round((cdbl(rsamr("netprice"))/cdbl(txtQtyPO))*cdbl(rsamr("doqty")), 0)
									End if 
									rsamr.close
									set rsamr = nothing
									
									
								end if
								 
								 
								' Perlu diperbaiki, 2012-12-28
								sqlSQLCheck = " select (quantity - received) as qtyRemained from purchaseorderdetails " & _ 
												 " where POCode='" & xCode & "' and LineNoKey = '" & request("LineNoKey_"&TRIM(n)) & "'" & _
												 " and fcBA='" & session("fcBA") & "'"
												
								Set rsSQL = Server.CreateObject("ADODB.Recordset")

								rsSQL.Open sqlSQLCheck, connstr, adOpenStatic 

								if not rsSQL.eof then
									qtyRemained = cdbl(rsSQL("qtyRemained"))
								End if
								rsSQL.Close
								set rsSQL = nothing	
                                
								' Add By Akbar  11 Mei 2018
								'response.write StockTransNo2
								If IsUseExtSys = "Y"  Then
									If StockTransNo2 <> ""  Then	
									 
										sqlSQLCheck = "  Select  nvl(sum(nvl(sto.doqty,0)),0) - nvl(sum(nvl(sto.dorcvqty,0)),0) qtydoremaind " & _ 
													  "	 from STOCKTRANSFERORDER sto " & _ 
													  "	 where sto.trackno = '" & xCode & "' " & _ 
													  "	 and sto.tracklineno = '" & request("LineNoKey_"&TRIM(n)) & "'" & _ 
													  "	 and sto.stono =   '" & StockTransNo2 & "' " & _  
													  "	 and sto.dono =  '" & DONo2 & "' " & _  
													  "	 and sto.fcba =  '" & session("fcBA") & "'" 
													
														
										Set rsSQL = Server.CreateObject("ADODB.Recordset")

										rsSQL.Open sqlSQLCheck, connstr, adOpenStatic 
										
										'response.write sqlSQLCheck
										if not rsSQL.eof then
											qtyDoRemaind = cdbl(rsSQL("qtydoremaind"))
										End if
										rsSQL.Close
										set rsSQL = nothing	
										
										If cdbl(qtyDoRemaind) > 0 and  cdbl(qtyDoRemaind) <= cdbl(qtyRemained) then
										   qtyRemained = qtyDoRemaind
										End if									
					
									End If  
                                End If 									
								
								if ( qtyRemained - cdbl(request("QtyReceived_"&TRIM(n))) ) < 0 then
										err.Number = 1001
										err.description = "Total QTY yang diterima tidak boleh lebih besar dari total QTY dalam item PO atau DO"
								Else
									q="Update PurchaseOrderDetails  set RECEIVED = received + " & request("QtyReceived_"&TRIM(n))  & _
									  ", LASTUPDATE = sysdate " & _
									  ", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
									" where POCode='" & xCode & "' and LineNoKey= '" & request("LineNoKey_"&TRIM(n)) & "'" & _
									" and fcBA='" & session("fcBA") & "'"
									connStr.execute(q)		

                                    q="Update STOCKTRANSFERORDER sto  set sto.dorcvqty = sto.dorcvqty  + " & request("QtyReceived_"&TRIM(n))  & _
									  ", LASTUPDATE = sysdate " & _
									  ", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
									  "	 where sto.trackno = '" & xCode & "' " & _ 
									  "	 and sto.tracklineno = '" & request("LineNoKey_"&TRIM(n)) & "'" & _ 
									  "	 and sto.stono =   '" & StockTransNo2 & "' " & _  
									  "	 and sto.dono =  '" & DONo2 & "' " & _  
									  "	 and sto.fcba =  '" & session("fcBA") & "'" 
									connStr.execute(q)											
												
									' Add by NAF, 20 Jan 2014
									' To handle Check Store for CAPEX
									qCheckStore = " select LOCATIONTYPE_API.get_item('" & request("LT_"&TRIM(n)) & "') item from dual "
									
									set rsqCheckStore = Server.CreateObject("ADODB.Recordset") 
									rsqCheckStore.Open qCheckStore, connstr
									if not rsqCheckStore.eof then
										tempItem = rsqCheckStore("item")
									else
										tempItem = ""
									end if 
									rsqCheckStore.Close
									Set rsqCheckStore = Nothing
									' End of modification
												
									' Add by NAF to handle if condition, except MATERIAL
									if tempItem = "MATERIAL" then 
									
										' Add by NAF to handle Inventory Stock On Hand
										' Insert the History of Received Inventory Transaction
										
										qLineMax = "select nvl(max(lineno),0) tempMaxLine from stocktranshist " & _
												   "where " & _
												   "fcba = '" & session("fcBA") & "'" 
											
										set rsqLineMax = Server.CreateObject("ADODB.Recordset")
										rsqLineMax.Open qLineMax, connstr
										if not rsqLineMax.eof then
											tempMaxLine = cdbl(rsqLineMax("tempMaxLine"))
										else
											tempMaxLine = 0
										end if 

										tempMaxLine = tempMaxLine + 1
											
										
										q = "Insert into StockTransHist ( " & _
											" LINENO, FCBA, STOCKCODE, TRANSDATE, TRANSTYPE, STORECODE, SOURCETYPE, SOURCECODE, TARGETTYPE, TARGETCODE, QUANTITY, LASTUPDATE, LASTTIME) values ( "& _
											  tempMaxLine & ",'" & _ 
											  session("fcBA") & "','" & _
											  request("Stock_"&TRIM(n)) &"'," & _
											  "to_date('" & DATECREATED & "','mm/dd/yyyy') ,'" & _
											  "RECEIVED" & "','" & _
											  xstoreCode & "'," & _
											  "'PR','" & _
											  xCode & "','" & _
											  request("LT_"&TRIM(n)) & "','" & _
											  request("LC_"&TRIM(n)) & "'," & _
											  request("QtyReceived_"&TRIM(n)) & ", sysdate, To_Char(sysdate,'HH24:MI') ) "
											  
										'response.Write(q)  
											  
										connStr.execute(q)
										
										'****** Add By Ukie at 20.03.2013 to handle rounding
										' Update by Akbar 06.06.2018
										  
										 Dim  QtySisa  
										 oQuery = " select (quantity - received) qtysisa" & _
													 " from purchaseorderdetails" & _
													 " where pocode = '" & xCode & "'" & _
													 "       and linenokey = '" & request("LineNoKey_"&TRIM(n)) & "'" & _
													 "       and fcba = '" & session("fcBA") & "'"

										 set oRecordSet = Server.CreateObject("ADODB.Recordset")
										 	oRecordSet.Open oQuery, connstr
										 If Not oRecordSet.eof Then
												QtySisa = cdbl(oRecordSet("qtysisa"))
										Else
										    	QtySisa = 0
										End If
										oRecordSet.Close()
										set oRecordSet = Nothing
										
										'If ( qtyRemained - cdbl(request("QtyReceived_"&TRIM(n))) ) = 0 Then
										 If QtySisa = 0 Then
    										oQuery = " select nvl(sum(nvl(amount, 0)), 0) amount" & _
													 " from purchaseorderdetails" & _
													 " where pocode = '" & xCode & "'" & _
													 "       and linenokey = '" & request("LineNoKey_"&TRIM(n)) & "'" & _
													 "       and fcba = '" & session("fcBA") & "'"

											set oRecordSet = Server.CreateObject("ADODB.Recordset")
											oRecordSet.Open oQuery, connstr
											If Not oRecordSet.eof Then
												nFixedPOAmount = cdbl(oRecordSet("amount"))
											Else
												nFixedPOAmount = 0
											End If
											oRecordSet.Close()
											set oRecordSet = Nothing
											
											oQuery = " select nvl(sum(nvl(amountreceived, 0)), 0) amountreceived" & _
													 " from receivenotedetail rnd" & _
													 "      , receivenote rn" & _
													 " where rnd.receivenotecode = rn.receivenotecode" & _
													 "       and rnd.fcba = rn.fcba" & _
													 "       and rn.pocode = '" & xCode & "'" & _
													 "       and rnd.purchaselinenokey = '" & request("LineNoKey_"&TRIM(n)) & "'" & _
													 "       and rnd.fcba = '" & session("fcBA") & "'"

											set oRecordSet = Server.CreateObject("ADODB.Recordset")
											oRecordSet.Open oQuery, connstr
											If Not oRecordSet.eof Then
												nCurrentRNAmount = cdbl(oRecordSet("amountreceived"))
											Else
												nCurrentRNAmount = 0
											End If
											oRecordSet.Close()
											set oRecordSet = Nothing
											
											oQuery = "  select nvl(sum(s.netprice), 0) StoPrice " & _
													 "	   from stocktransferorder s " & _
													 "	  where s.trackno = '" & xCode & "'" & _
													 "		and s.tracklineno = '" & request("LineNoKey_"&TRIM(n)) & "'" & _
													 "      and s.fcba = '" & session("fcBA") & "'"

											set oRecordSet = Server.CreateObject("ADODB.Recordset")
											oRecordSet.Open oQuery, connstr
											If Not oRecordSet.eof Then
												nFixedPOAmount = cdbl(nFixedPOAmount) + cdbl(oRecordSet("StoPrice"))
											End If
											oRecordSet.Close()
											set oRecordSet = Nothing
											
											Selisih = (nFixedPOAmount - nCurrentRNAmount)
											txtamountreceived = Selisih
											
											
											q = " update purchaseorderdetails" & _
												" set rowstateid = 'Completed'" & _
												" where pocode = '" & xCode & "'" & _
												"       And (quantity - received) = 0 and linenokey = '" & request("LineNoKey_"&TRIM(n)) & "'" & _
												"       and fcba = '" & session("fcBA") & "'"
												
											'response.Write q
											connStr.execute(q)
											
											If IsUseExtSys = "Y"  Then
												q = " Update STOCKTRANSFERORDER sto " & _
													" set rowstateid = 'Completed'" & _
													"	 where sto.trackno = '" & xCode & "' " & _ 
													"	 and sto.tracklineno = '" & request("LineNoKey_"&TRIM(n)) & "'" & _ 
													"	 and sto.stono =   '" & StockTransNo2 & "' " & _  
													"	 and sto.dono =  '" & DONo2 & "' " & _  
													"	 and sto.fcba =  '" & session("fcBA") & "'  and (doqty - dorcvqty) = 0 " 
													
												'response.Write q
												connStr.execute(q)
										    End if		
											
										End If
										
										q="Insert into  RECEIVENOTEDETAIL (RECEIVENOTECODE, LOCATIONTYPE, LOCATIONCODE, PURCHASEITEMCODE, " & _
										"JOBCODE, QUANTITY, QTYRETURN, AMOUNTRECEIVED, AMOUNTRETURN, FCENTRY, FCEDIT, FCIP, FCBA, linenohist, purchaselinenokey, LASTUPDATE, LASTTIME) " & _
										"values ('" & fcCode & "','" & request("LT_"&TRIM(n)) & "','" & request("LC_"&TRIM(n)) &  _
										"','" & request("Stock_"&TRIM(n)) & "','" & txtControlJOb & "','" & request("QtyReceived_"&TRIM(n)) & _
										"',0,'"& txtamountreceived & "',0," & _
										"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
										Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "',"& tempMaxLine & ",'" & request("LineNoKey_"&TRIM(n)) & "', sysdate, To_Char(sysdate,'HH24:MI'))"
										'response.Write(q)
									
										connStr.execute(q)
									else
									
										q="Insert into  RECEIVENOTEDETAIL (RECEIVENOTECODE, LOCATIONTYPE, LOCATIONCODE, PURCHASEITEMCODE, " & _
											"JOBCODE, QUANTITY, QTYRETURN, AMOUNTRECEIVED, AMOUNTRETURN, FCENTRY, FCEDIT, FCIP, FCBA, purchaselinenokey, LASTUPDATE, LASTTIME) " & _
											"values ('" & fcCode & "','" & request("LT_"&TRIM(n)) & "','" & request("LC_"&TRIM(n)) &  _
											"','" & request("Stock_"&TRIM(n)) & "','" & txtControlJOb & "','" & request("QtyReceived_"&TRIM(n)) & _
											"',0,'"& txtamountreceived & "',0," & _
											"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
											Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "','" & request("LineNoKey_"&TRIM(n)) & "', sysdate, To_Char(sysdate,'HH24:MI'))"
											'response.Write(q)
										
										connStr.execute(q)
										
										
									end if

									' Add by NAF to handle if condition, except MATERIAL
									if tempItem = "MATERIAL" then 										
										If request("LT_"&TRIM(n)) = "MN" Then
										
											oQuery = " select count(*) total" & _
													 " from stockonhandnursery" & _
													 " where storecode = '" & xstoreCode & "'" & _
													 "       and stockcode = '" & request("Stock_"&TRIM(n)) & "'" & _
													 "       and nurserycode = '" & request("LC_"&TRIM(n)) & "'" & _
													 "       and fcba = '" &  session("fcBA") & "'"

											set oRecordSet = Server.CreateObject("ADODB.Recordset")
											oRecordSet.Open oQuery, connstr
											if not oRecordSet.eof then
												nTotalRecordOnStockOnHandNursery = cdbl(oRecordSet("total"))
											else
												nTotalRecordOnStockOnHandNursery = 0
											end if
											oRecordSet.Close()
											set oRecordSet = Nothing
											
											If nTotalRecordOnStockOnHandNursery = 0 then
											
												' Just Insert the New Record with the following/received Quantity
												q = " Insert into stockonhandnursery" & _
													" ( storecode, stockcode, nurserycode, fcba, quantity, lastupdate, lasttime ) values (" & _
													" '" & xstoreCode & "'" & _
													" , '" & request("Stock_"&TRIM(n)) & "'" & _
													" , '" & request("LC_"&TRIM(n)) & "'" & _
													" , '" & session("fcBA") & "'" & _
													" , '" & cdbl(request("QtyReceived_"&TRIM(n))) & "'" & _
													" , sysdate" & _
													" , to_char(sysdate,'hh24:mi')" & _
													" )"

												connStr.execute(q)

											Else
											
												q = " update stockonhandnursery " & _
													" set quantity = quantity + " & cdbl(request("QtyReceived_"&TRIM(n))) & " " & _
													"     , LASTUPDATE = sysdate " & _
													"     , LASTTIME = To_Char(sysdate,'HH24:MI') "& _
													" where storecode = '" & xstoreCode & "'" & _
													"       and stockcode  = '" & request("Stock_"&TRIM(n)) & "'" & _
													"       and nurserycode  = '" & request("LC_"&TRIM(n)) & "'" & _
													"       and fcba = '" & session("fcBA") &"'"

												connStr.execute(q)
											End If
										End If

										qTotalRecordOnHand = "select count(*) totalRecordOnHand from stockonhand " & _
															 "where " & _
															 "storecode = '" & xstoreCode & "' and " & _												 
															 "stockcode = '" & request("Stock_"&TRIM(n)) & "' and " & _
															 "fcba = '" &  session("fcBA") & "'"

										set rsTotalRecordOnHand = Server.CreateObject("ADODB.Recordset")
										rsTotalRecordOnHand.Open qTotalRecordOnHand, connstr
										if not rsTotalRecordOnHand.eof then
											tempTotalRecordOnHand = cdbl(rsTotalRecordOnHand("totalRecordOnHand"))
										else
											tempTotalRecordOnHand = 0
										end if 
										
										if tempTotalRecordOnHand = 0 then
										
											' Just Insert the New Record with the following/received Quantity
											q = "Insert into stockonhand ( " & _
												" STORECODE, STOCKCODE, FCBA, QUANTITY, LASTUPDATE, LASTTIME ) values ( '" & _
												xstoreCode & "','" & _
												request("Stock_"&TRIM(n)) & "','" & _
												session("fcBA") & "', '" & cdbl(request("QtyReceived_"&TRIM(n))) & "', sysdate, To_Char(sysdate,'HH24:MI') ) "
												
												
											connStr.execute(q)
											
										Else	  

											q = "update stockonhand " & _
												"set quantity = quantity + " & cdbl(request("QtyReceived_"&TRIM(n))) & " " & _ 
												", LASTUPDATE = sysdate " & _
												", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
												"where " & _
												"storecode = '" & xstoreCode & "' and " & _
												"stockcode  = '" & request("Stock_"&TRIM(n)) & "' and " & _ 
												"fcba = '" & session("fcBA") &"' "
												 
											connStr.execute(q)
										End If
									End if
								End If
							End If
						Next
						
						if err.number=0 then
						    connStr.CommitTrans	
							If IsUseExtSys = "Y"  Then
								strQueryStatus = "BEGIN BridgingSAP_PRPOGRN_API.ReleaseGRNToBridge('" & Session("fcBA") & "','" & fcCode & "'); END; "
								connStr.execute(strQueryStatus)	
                            End If 								
						End if
						
					else
						err.number = 1001
						err.description = "Tanggal """ & DATECREATED & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
					End If
		
					cmdView	= "True"
    else
        cmdView	= "True"
	end if								 
	'Response.Write "sql " & q
 else
    cmdView	= request("cmdView")		    
 End if

 Dim rsVoucher,SudahClose
 set rsVoucher = server.CreateObject("Adodb.recordset")
 q="select * from RECEIVENOTE  " & _
  " where RECEIVENOTECODE='" & fcCode & "' and " & _
  " fcBA='" & session("fcBA") & "'"
 SudahClose="N"
 'response.Write	q
 rsVoucher.Open q, connstr
 if Not rsVoucher.Eof then
    DATECREATED=FormatDateTime(rsVoucher("RNDATE"))
	DATECREATED = FormatDateTime(DATECREATED)
	'response.write FormatDateTime(rsVoucher("RNDATE"))
	xCode=rsVoucher("POCODE")
	StockTransNo2 =  rsVoucher("StockOrderNo")
	DONo2 =  rsVoucher("DELIVERYORDERCODE")
	xstoreCode = rsVoucher("STORECODE")
	SudahClose="Y"
 End if
 rsVoucher.Close
 Set rsVoucher = Nothing
 'response.Write	StockTransNo2 & ", " & DONo2


Dim IsAvlSto : IsAvlSto = "N"
'Dim rs
oSQLQuery = "  Select nvl(sum(s.doqty - s.dorcvqty),0) qtyOSTD  from STOCKTRANSFERORDER s " & _
			"	Where s.doqty - s.dorcvqty > 0 " & _
			"	  and s.trackno = '" & xCode& "'  " 
			
'response.Write	oSQLQuery   
set rs = server.CreateObject("Adodb.recordset")
rs.Open oSQLQuery, connstr
If Not rs.Eof then
    If cdbl(rs("qtyOSTD")) > 0 then
	    IsAvlSto = "Y"
	end if	
End if
'response.Write "IsAvlSto " & IsAvlSto
rs.Close
Set rs = Nothing


Dim IsAvlLCO : IsAvlLCO = ""
'Dim rs
oSQLQuery = "  Select incoterm1  from PURCHASEORDER s " & _
			"	Where  s.pocode = '" & xCode& "'  " 
			
'response.Write	oSQLQuery   
set rs = server.CreateObject("Adodb.recordset")
rs.Open oSQLQuery, connstr
If Not rs.Eof then
	IsAvlLCO = rs("incoterm1")
End if
'response.Write "IsAvlSto " & IsAvlSto
rs.Close
Set rs = Nothing



Dim val_IsTrue : val_IsTrue = "N"
Dim oRecordSet, oSQLQuery
oSQLQuery = " select 'Y' istrue" & _
            " from periodcontrol" & _
	        " where fcba = '" & Session("fcBA") & "'" & _
	        "       and to_date ('" & DATECREATED & "', 'mm/dd/yyyy') >= currentdate" & _
	        "       and to_date ('" & DATECREATED & "', 'mm/dd/yyyy') <= last_day(add_months(currentdate, 1))" & _
	        "       and upper(system) = upper('Purchase')"
	   
set oRecordSet = server.CreateObject("Adodb.recordset")
oRecordSet.Open oSQLQuery, connstr
If Not oRecordSet.Eof then
	val_IsTrue = oRecordSet("istrue")
Else
	val_IsTrue = "N"
End if
'response.Write strNurseryCulling
oRecordSet.Close
Set oRecordSet = Nothing


If Err.Number <> 0 then
	connStr.RollbackTrans
	Response.Write "<div class = ""error"">" & Err.Description & "</div>"
End if

 %>
 <form name="frmEntry" method="post" action="ReceiveNote.Asp">
	<table class="data" width="100%">
		<tr>
		<td align="left">GRN No</td>
		<td align="left"><input type="text" maxlength="50" name="fcCode" size="30" value="<%=fcCode%>"><font color="#FF0000">&nbsp;*</font></td></tr><tr>
		<td align="left">Date</td>
		<td align="left"><input class="datepicker" type="text" maxlength="10" name="DATECREATED" size="10" value="<%=DATECREATED%>">
		<font color="#FF0000">&nbsp;*</font>
		</td></tr>
		<tr>
			<td align="left">PO Number</td>
			<td align="left">      
				<%if SudahClose="Y" then%>
					<select name="xCode" id="xCode"><option value="<%=xCode%>" Selected><%=xCode%></option></select>
				<%else%>
					<select name="xCode" id="xCode" onchange="getSto(this)" >
					<option value="" ></option> 
					<%
					dim rsCallCombo
					Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
					rsCallCombo.CursorType = 3
					rsCallCombo.CursorLocation = 3
					rsCallCombo.Open " Select distinct(a.poCode)" & _
									 " from PurchaseOrderDetails a" & _ 
									 "      , purchaseorder b " & _
									 " where a.pocode = b.pocode" & _
									 "       and a.fcba = b.fcba" & _
									 "       and nvl(b.rowstateid, 'Ordered') = 'Ordered'" & _
									 "       and a.fcba = '" & session("fcba") & "'" & _
									 "       and a.received < a.quantity", connstr
					if not rsCallCombo.eof then
					 Do While Not rsCallCombo.eof
					   
						if rsCallCombo("poCode") = xCode then%>
						   <option value="<%=rsCallCombo("poCode")%>" Selected><%=rsCallCombo("poCode")%></option> 
						<%Else%>
						   <option value="<%=rsCallCombo("poCode")%>"><%=rsCallCombo("poCode")%></option>  
						<%End if%>
						<%rsCallCombo.movenext
					 loop
					End if
					rsCallCombo.Close
					set rsCallCombo = Nothing%>
					</select>
				<%End if%>
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="left">STO No</td>
			<td align="left" > 
			<%if SudahClose="Y" then  
					%>
				<select name="StockTransNo2" id="StockTransNo2" ><option value="<%=StockTransNo2%>" Selected><%=StockTransNo2%></option></select>
				
			<%else%>		
			    <select name="StockTransNo2" id="StockTransNo2" onchange="getdono(this)">
				<%
			              				 
						   rst = " Select distinct a.stono from stocktransferorder  a where (a.doqty - nvl(a.dorcvqty,0 )) > 0 And  a.fcba = '" & session("fcba") & "' And a.trackno='" & xCode & "'  "
						   'response.write rst
						   Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
						   rsCallCombo.Open rst, connstr			  
							 if not rsCallCombo.eof then
							 Do While Not rsCallCombo.eof
							   
								if rsCallCombo("stono") = StockTransNo2 then%>
								   <option value="<%=rsCallCombo("stono")%>" Selected><%=rsCallCombo("stono")%></option> 
								<%Else%>
								   <option value="<%=rsCallCombo("stono")%>"><%=rsCallCombo("stono")%></option>  
								<%End if%>
								<%rsCallCombo.movenext
							 loop
							End if
							rsCallCombo.Close
				            set rsCallCombo = Nothing%>
				</select>
				  
			 <% End if %>	  
			</td>
		</tr>
		<tr>
			<td align="left">DO No</td>
			<td align="left" > 
			<%if SudahClose="Y" then  
					%>
				<select name="DONo2" id="DONo2"  ><option value="<%=DONo2%>" Selected><%=DONo2%></option></select>
			<%else	%>	
					 <select name="DONo2" id="DONo2"  >  
					<%
			              dim rst					 
						  rst = " select distinct a.dono from stocktransferorder  a where (a.doqty - nvl(a.dorcvqty,0 )) > 0 And a.fcba = '" & session("fcba") & "' And a.trackno='" & xCode & "' and stono = '" & StockTransNo2 & "'"
						   'response.write rst
						   Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
						   rsCallCombo.Open rst, connstr			  
							 if not rsCallCombo.eof then
							 Do While Not rsCallCombo.eof
							   
								if rsCallCombo("dono") = DONo2 then%>
								   <option value="<%=rsCallCombo("dono")%>" Selected><%=rsCallCombo("dono")%></option> 
								<%Else%>
								   <option value="<%=rsCallCombo("dono")%>"><%=rsCallCombo("dono")%></option>  
								<%End if%>
								<%rsCallCombo.movenext
							 loop
							End if
							rsCallCombo.Close
				            set rsCallCombo = Nothing%>
				</select>
				  
			 <% End if %>	  
			</td>
		</tr>
		
		<tr>
		<td align="left">Store Code</td>
		<td align="left">        
				<select name="xstoreCode" value="<%=xstoreCode%>"> 
					<%cmdCallCombo "store",xstoreCode,"",""%>
				</select>
				<font color="#FF0000">&nbsp;*</font>
        </td>
		</tr>
		
		
		<tr>
			<td colspan="2"  align="left"> 
				<input Type="Hidden" Value="True" Name="cmdView">  
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">  
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">   	 
				<input Type="Submit" name="Add" value="Select">&nbsp;&nbsp;  
				<a href="goods_received_note.asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&grnno=<%=fcCode%>" target="_blank"><img src="IMAGES/Print.gif" border="0" title="Print Preview"><font face="Verdana" size="1">Print</font></a>&nbsp;&nbsp;
				<%If val_IsTrue = "Y" Then%>
					<a href="ReceiveNote.Asp?fcCode=<%=fcCode%>&DATECREATED=<%=DATECREATED%>&txtJenis=Header"><img src="IMAGES/delete.gif" alt="Delete Data : <%=fccode%>" border="0" onClick="return confirm('Are you sure to delete');"></a>
				<%Else%>
					<img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18">
				<%End If%>
				<%
		dim dataModule,dataSubmodule,dataDetail
		dataModule 		= "Purchase"
		dataSubmodule 	= "Receive Note"
		dataDetail 		= "NULL"
		%>
		<a class="jqplot-image-button" style="" href="ajax/hadirAjax.asp?dataModule=<%=dataModule%>&dataSubmodule=<%=dataSubmodule%>&dataDetail=<%=dataDetail%>"><i class="fa fa-check-square-o"></i>&nbsp;Hadir</a>
		<!-- END:HADIR -->
		</td> 
		</tr> 
    </table>    
 </form>
<%
'response.Write "IsAvlSto " & IsAvlSto
 'response.Write "IsAvlLCO " &  IsAvlLCO
 
 
If cmdView = "True" and LEN(TRIM(fcCode)) > 4 and len(trim(xstoreCode)) > 0 Then
	Dim rsGRN,intRecordsPerPage,cColor
	set rsGRN = server.CreateObject("Adodb.recordset")
	If SudahClose = "Y" Then
		q	= " select purchaseitemcode  " & _
		      "      , purchaselinenokey " & _
			  "      , linenohist " & _
			  "      , jobcode " & _
			  "      , locationtype " & _
			  "      , locationcode " & _
			  "      , nvl(sum(nvl(quantity, 0)), 0) as quantity " & _
			  "      , nvl(sum(nvl(amountreceived, 0)), 0) as amountreceived " & _			  
			  " from receivenotedetail p " & _
			  "      , stock s " & _
			  " where (p.purchaseitemcode = s.fccode and p.fcba = s.fcba) " & _
			  "       and receivenotecode = '" & fcCode & "'" & _
			  "       and p.fcBA = '" & session("fcBA") & "'" & _
			  " group by purchaseitemcode" & _
			  "        , purchaselinenokey" & _
			  "        , linenohist" & _
			  "        , jobcode" & _
			  "        , locationtype" & _
			  "        , locationcode" & _
			  " order by purchaseitemcode" & _
			  "        , purchaselinenokey"
	end if	

    If 	IsAvlLCO = "LCO" and StockTransNo2 = ""  then
	    Response.Write " This PO Must Be Stock Transfer Order ( STO ) , Please choose or download"
		q	= " select * from purchaseorderdetails p" & _	
		      " where pocode='TidakAda' "
	Else
		If SudahClose = "N" AND IsAvlSto ="N" then	 
				q	= " select p.itemcode" & _
					  "        , s.fcname" & _
					  "        , p.linenokey purchaselinenokey" & _
					  "        , nvl(sum(nvl(p.quantity, 0)), 0) quantityorder" & _
					  "        , nvl(sum(nvl(p.quantity, 0)), 0) - nvl(sum(nvl(p.received, 0)), 0) as quantityreceived" & _
					  "        , nvl(sum(nvl(amount, 0)), 0) as amount" & _		      
					  " from purchaseorderdetails p" & _
					  "      , stock s" & _
					  " where (p.itemcode = s.fccode and p.fcba = s.fcba)" & _
					  "       and pocode = '" & xCode & "'" & _
					  "       and p.received < quantity" & _
					  "       and p.fcba = '" & session("fcBA") & "'" & _
					  " group by p.itemcode" & _
					  "          , s.fcname" & _
					  "          , p.linenokey"
		Elseif SudahClose = "N" AND IsAvlSto ="Y" then	 
				q	= " select p.itemcode " & _
					  "      , stock_api.get_stockName(p.itemcode,p.fcba) fcname " & _
					  "      , p.linenokey purchaselinenokey " & _
					  "      , nvl(sum(nvl(p.quantity, 0)), 0) quantityorder " & _
					  "      , nvl(sum(nvl(p.quantity, 0)), 0) - nvl(sum(nvl(p.received, 0)), 0) as quantityreceived " & _
					  "      , sum(s.doqty) doqty " & _
					  "      , sum(s.doqty)- sum(s.dorcvqty) doreceiveqty " & _
					  "      , round(((nvl(sum(nvl(amount, 0)), 0)/sum(p.quantity)) * sum(s.doqty))  + ((sum(nvl(s.netprice,0))/sum(p.quantity))* sum(s.doqty))) as amount " & _         
					  " from purchaseorderdetails p right join STOCKTRANSFERORDER s " & _
					  " on ( p.pocode = s.trackno  and p.linenokey = s.tracklineno and p.fcba=s.fcba ) " & _
					  " where p.pocode = '" & xCode& "'  " & _
					  " And s.stono = '" & StockTransNo2 & "'  " & _
					  " And s.dono= '" & DONo2 & "'  " & _
					  " And s.ROWSTATEID = 'Released' " & _
					  "      and s.doqty - s.dorcvqty > 0" & _
					  "      and p.fcba = '" & session("fcBA") & "'  " & _
					  " group by p.itemcode " & _
					  "          , p.fcba  " & _
					  "          , p.linenokey"		
			
		End if
	End If		
   'response.Write q	
   Set rsGRN = Server.CreateObject("ADODB.Recordset")
   rsGRN.CursorType = 3
   rsGRN.CursorLocation = 3
   rsGRN.Open q, connstr%>
   <form   action="ReceiveNote.Asp" method="post" name="eForm" >
   <div style="overflow-x: auto">
   <table class="data" border="1" width="90%" >
   <%
     'response.Write "IsAvlSto " & IsAvlSto
	 if SudahClose<>"Y" AND IsAvlSto ="Y" Then%>
	   <tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;">
	   <td align="center">No</td>
	   <td align="center">LineNoKey</td>
	   <td align="center">Code</td>
	   <td align="center">Desc.</td>
	   <td align="center">Qty PO</td>
	   <td align="center">Sisa Qty PO</td>
	   <td align="center">Qty DO</td>
	   <td align="center">Sisa Qty DO</td>
	   <td align="center">Amount</td>
	   <td align="center">Job Code</td>
	   <td align="center">Target Type</td>
	   <td align="center">Target Code</td>
	   <td align="center">Quantity</td>
	   <!--<td align="center">Amount</td>-->
   	   </tr>
   <%ElseIf  SudahClose<>"Y" AND IsAvlSto ="N" Then	%>
	   <tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;">
	   <td align="center">No</td>
	   <td align="center">LineNoKey</td>
	   <td align="center">Code</td>
	   <td align="center">Desc.</td>
	   <td align="center">Order (Qty)</td>
	   <td align="center">Sisa order</td>
	   <td align="center">Amount</td>
	   <td align="center">Job Code</td>
	   <td align="center">Target Type</td>
	   <td align="center">Target Code</td>
	   <td align="center">Quantity</td>
	   <!--<td align="center">Amount</td>-->
   	   </tr>
     <%Else	%>
	   <tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;">
	   <td align="center">No</td>
	   <td align="center">&nbsp;</td>
	   <td align="center">Stock Code</td>
	   <td align="center">LineNoKey</td>
	   <td align="center">Line No Hist</td>
	   <td align="center">Control Job</td>
	   <td align="center">Target Type</td>
	   <td align="center">Target Code</td>
	   <td align="center">Quantity</td>
	   <!--<td align="center">Amount</td>-->
   	   </tr>
   <%End if
   
   
    if Not rsGRN.eof then
      nCountofField=rsGRN.fields.count-1
	  no=1
	  While Not rsGRN.EOF

		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
			<tr bgcolor="<%=cColor%>">
				<td align="left" valign="top"><%=NO%></td>
				<%IF SudahClose = "Y" then%>
					<%If val_IsTrue = "Y" Then%>
					<td valign="middle">
						<a href="ReceiveNote.Asp?fcCode=<%=fcCode%>&ItemCode=<%=rsGRN("PURCHASEITEMCODE")%>&purchaseLineNoKey=<%=rsGRN("PurchaseLineNoKey")%>&linenomorhist=<%=rsGRN("linenohist")%>&DATECREATED=<%=DATECREATED%>&txtJenis=Details">
						<img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGRN(1)%>" border="0" onClick="return confirm('Are you sure to delete');"></a>
					</td>
					<%Else%>
					<td>
						<img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18">
					</td>
					<%End If%>
					<%for i = 0 to nCountofField-1%>
						<td align="left" title="<%=rsGRN(i).Name%>" valign="top"><%=rsGRN(i)%></td>
					<%next%>
				<%End if%>
				
				<%if SudahClose<>"Y"  AND IsAvlSto ="N"  Then%>
					<td align="Right" title="LineNoKey" valign="top"><%=rsGRN("purchaselinenokey")%></td>
					<td align="left" title="Item Code" valign="top"><%=rsGRN("itemcode")%></td>
					<td align="left" title="Item Name" valign="top"><%=rsGRN("fcname")%></td>
					<td align="Right" title="QuantityOrder" valign="top"><%=rsGRN("quantityorder")%></td>
					<td align="Right" title="QuantityReceived" valign="top"><%=rsGRN("quantityreceived")%></td>
					<td align="Right" title="Amount" valign="top"><%=rsGRN("Amount")%></td>

					 <td align="left" width="2"><input type="text" name="Jobcode_<%=no%>" width="2" value="<%=request("Jobcode_"&TRIM(no))%>"><a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=Job&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana"   size="1">Activity</font></a></td>
					 <td align="left" width="2"><input type="text" name="LT_<%=no%>" width="2" value="<%=request("LT_"&TRIM(no))%>"><a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=LocationType&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana"   size="1">Target Type</font></a><font color="#FF0000">&nbsp;*</font></td>
					 <td align="left" width="10"><input type="text" name="LC_<%=no%>" width="10" value="<%=request("LC_"&TRIM(no))%>"></td>
					 <td align="right" width="10"><input type="text" name="QtyReceived_<%=no%>" width="10" value="<%=request("QtyReceived_"&TRIM(no))%>"  onkeypress="return harusangka(event);" >
					 <input type="hidden" name="AmountReceived_<%=no%>" width="10" value="<%=txtamoutreceived%>">
					 <input type="hidden" name="LineNoKey_<%=no%>" width="10" value="<%=rsGRN("PurchaseLineNoKey")%>">
					 <input type="hidden" name="Stock_<%=no%>" value="<%=rsGRN("ITEMCODE")%>"></td>
				<%elseif SudahClose<>"Y"  AND IsAvlSto ="Y"  Then%>
					<td align="Right" title="LineNoKey" valign="top"><%=rsGRN("purchaselinenokey")%></td>
					<td align="left" title="Item Code" valign="top"><%=rsGRN("itemcode")%></td>
					<td align="left" title="Item Name" valign="top"><%=rsGRN("fcname")%></td>
					<td align="Right" title="QuantityOrder" valign="top"><%=rsGRN("quantityorder")%></td>
					<td align="Right" title="QuantityReceived" valign="top"><%=rsGRN("quantityreceived")%></td>
					<td align="Right" title="QuantityReceived" valign="top"><%=rsGRN("doqty")%></td>
					<td align="Right" title="QuantityReceived" valign="top"><%=rsGRN("doreceiveqty")%></td>
					<td align="Right" title="Amount" valign="top"><%=rsGRN("Amount")%></td>

					 <td align="left" width="2"><input type="text" name="Jobcode_<%=no%>" width="2" value="<%=request("Jobcode_"&TRIM(no))%>"><a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=Job&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana"   size="1">Activity</font></a></td>
					 <td align="left" width="2"><input type="text" name="LT_<%=no%>" width="2" value="<%=request("LT_"&TRIM(no))%>"><a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=LocationType&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana"   size="1">Target Type</font></a><font color="#FF0000">&nbsp;*</font></td>
					 <td align="left" width="10"><input type="text" name="LC_<%=no%>" width="10" value="<%=request("LC_"&TRIM(no))%>"></td>
					 <td align="right" width="10"><input type="text" name="QtyReceived_<%=no%>" width="10" value="<%=request("QtyReceived_"&TRIM(no))%>" onkeypress="return harusangka(event);" >
					 <input type="hidden" name="AmountReceived_<%=no%>" width="10" value="<%=txtamoutreceived%>">
					 <input type="hidden" name="LineNoKey_<%=no%>" width="10" value="<%=rsGRN("PurchaseLineNoKey")%>">
					 <input type="hidden" name="Stock_<%=no%>" value="<%=rsGRN("ITEMCODE")%>"></td>
			    <%End if%>
		  </tr>
		  <%
		  rsGRN.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   else
      response.Write("Data is empty")	  
   end if 
   rsGRN.Close
   set rsGRN = Nothing%>
	  <tr><td colspan="13" align="left">
		  <!-- diambil dari form yang diatas -->
	      <input type="hidden" name="fcCode" value="<%=fcCode%>">
		  <input type="hidden" name="DATECREATED" value="<%=DATECREATED%>">
		  <input type="hidden" name="xCode" value="<%=xCode%>">
		  <input type="hidden" name="StockTransNo2" value="<%=StockTransNo2%>">
		  <input type="hidden" name="DONo2" value="<%=DONo2%>">
		  <input type="hidden" name="xstoreCode" value="<%=xstoreCode%>">
		  <input type="hidden" name="No" value="<%=No%>">
		  
		  <input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		  <input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		  <%If val_IsTrue = "Y" Then%>
		  <%IF Request("txtJenis")="Edit" then%>
		       <input type="hidden" name="txtJenis" value="Edit">
		       <input Type="Submit"  name="Add" value="Update">
		  <%Else%>
		       <input type="hidden" name="txtJenis" value="New">
		       <input Type="Submit"  name="Add" value="Submit" onClick="return confirm('Apakah Anda sudah Cek jumlah dan qualitas dari tiap Barang yg diterima ?');">
		  <%End if%>
		  <%End if%>
		  <input Type="Reset"  name="Reset" value="Reset">
		  <input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
	  </tr>
     
   </TABLE>
   </div>
    </div>
   </form>
<%
End if  
 
'-------------akhir----------------------
  %>
 </div> </div>
  <!--#include file="bottom.jsp"-->

</body>
</html>
