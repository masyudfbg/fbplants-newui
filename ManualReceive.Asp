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
	<title>ePlaS - RECEIVED TASK</title>
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
function getordernumber(str){

	var xmlhttp;
	if (str=="")
	  {
	  document.getElementById("TransportTaskNo").innerHTML="";
	  document.getElementById("manualreceive").innerHTML="";
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
		document.getElementById("TransportTaskNo").innerHTML=xmlhttp.responseText;
		document.getElementById("manualreceive").innerHTML="";
		}
	  }
	xmlhttp.open("GET","ajax/getordernumber.asp?fcbasource="+str+"&fcbaDistination=<%=Session("fcBA")%>",true);
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
Dim xCode,xstoreCode,TempOrderNo,cmdView,DATECREATED,VoucherCode,QUANTITY,AMOUNT,OTHERITEMDESC,v_CrossCekPO
Dim Remarks
Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis,TOTALAMOUNT,HeaderRemarks
Dim Jobcode, LocationType, LocationCode, Reference, Remark
Dim qLineMax, rsqLineMax, tempMaxLine
Dim qTotalRecordOnHand, rsTotalRecordOnHand, tempTotalRecordOnHand
Dim qQuantityOnHand, rsQuantityOnHand, tempQuantityOnHand
Dim sqlSQLCheck, rsSQL, qtyRemained
Dim qCheckStore, rsqCheckStore
Dim txtLocationType	
Dim TransportTaskNo

Dim txtOrderNo : txtOrderNo = 0
Dim oRs : oRs = 0
Dim oSql : oSql = 0
Dim txtLocationCode : txtLocationCode = ""
Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""

			

txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")
			
			
if request("TempOrderNo")="" then
	TempOrderNo="" '----- Nantinya Ambil dari Format Master
else
	TempOrderNo = request("TempOrderNo")
end if

'response.write TempOrderNo
			
If  Len(Trim(TempOrderNo)) = 0 Then TempOrderNo = "(NEW)"			
	txtJenis = request("txtJenis") '-- Untuk Membedakan Jenis Data Yang akan diolah [insert,edit,delete]
	'---- Data Header
	if Request("DATECREATED") <> "" then
		DATECREATED	=	Request("DATECREATED")
	else
		DATECREATED	=	textDate(date)
	end if
			
	xCode = Request("xCode")  '-- fcba source

			
	HeaderRemarks = Request("HeaderRemarks")
			
	'---- Data Details

	TransportTaskNo = Request("TransportTaskNo")
	
	'response.write Request("txtJenis")
			
	'--- Action ke Database di Oracle [ insert,edit,delete ]
	If  Request("txtJenis")="New" or Request("txtJenis")= "Header" or Request("txtJenis")= "Details" or Request("txtJenis")="Edit" Then
		REMARKS	=	request("REMARKS")
		if Request("txtJenis")= "Header" then
					'Simpan Order No nya terlebih dahulu sebelum di delete
					connstr.BeginTrans
					
					Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), DATECREATED)
					If Cek_IsDailyClosing = "FALSE" then

						Set rsMNLRCV = Server.CreateObject("ADODB.Recordset")
						rsMNLRCV.Open " select OrderNo" & _
								  " from MANUALRECEIVE_TAB" & _
								  " where OrderNo = '" & TempOrderNo & "'" & _
								  "       and fcBA = '" & Session("fcBA") & "'", connstr
						IF Not rsMNLRCV.eof then
							txtOrderNo =  rsMNLRCV("OrderNo")
						End if
						rsMNLRCV.Close
						set rsMNLRCV = Nothing

						q = " delete from MANUALRECEIVE_TAB" & _
							" where OrderNo = '" & TempOrderNo & "'" & _
							"       and fcba = '" & Session("fcBA") & "'"
						
						connStr.execute(q)
						TempOrderNo = "(NEW)"
						
						if err.number = 0 then
							connStr.CommitTrans
						end if
						
					else
						err.number = 1001
						err.description = "Tanggal """ & DATECREATED & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
					End If
					
					cmdView	= "True"
		elseif Request("txtJenis")= "Details" then
					'Simpan Order No nya terlebih dahulu sebelum di delete
					connstr.BeginTrans
					
					Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), DATECREATED)
					
					
					If Cek_IsDailyClosing = "FALSE" then
						
						Set rsMNLRCV = Server.CreateObject("ADODB.Recordset")
						rsMNLRCV.Open "Select OrderNo from MANUALRECEIVE_TAB where  OrderNo='" & TempOrderNo & "'" & _
						" and fcBA='" & Session("fcBA") & "'", connstr
						IF Not rsMNLRCV.eof then
							txtOrderNo =  rsMNLRCV("OrderNo")
						End if
						rsMNLRCV.Close
						set rsMNLRCV = Nothing
						
						q = " Select LineNoHist" & _
							"       , QUANTITYRECEIVED" & _
							"       , STOCKCODE" & _
							"       , FCBASOURCE" & _
							"       , ORDERSOURCE" & _ 
							"       , LINENOSOURCE" & _
							"       , LOCATIONTYPE" & _
							"       , LOCATIONCODE" & _
							" from MANUALRECEIVEDTL_TAB" & _
							" where OrderNo = '" & TempOrderNo & "'" & _
							"       and LINENOSOURCE = '" & Request("linenosource") & "'" & _
							"       and fcBA = '" & Session("fcBA") & "'"
						
						Set rsMNLRCVDetail = Server.CreateObject("ADODB.Recordset")
						rsMNLRCVDetail.Open q, connstr
								
						IF Not rsMNLRCVDetail.eof then
							txtLineNoHist = rsMNLRCVDetail("LINENOHIST")
							originalQuantity = CDbl(rsMNLRCVDetail("QUANTITYRECEIVED"))
							txtItemCodeDelete = rsMNLRCVDetail("STOCKCODE")
							txtItemLineKeyDelete = rsMNLRCVDetail("LINENOSOURCE")
							txtOrderSource = rsMNLRCVDetail("ORDERSOURCE")
							txtFcbaSource = rsMNLRCVDetail("FCBASOURCE")
							txtLocationType = rsMNLRCVDetail("LOCATIONTYPE")
							txtLocationCode = rsMNLRCVDetail("LOCATIONCODE")
						End if
						rsMNLRCVDetail.Close
						set rsMNLRCVDetail = Nothing					
						
						qLessZero = " select islesszero" & _
									" from stock" & _
									" where fccode = '" & txtItemCodeDelete & "'" & _
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
										  " where storecode = '" & txtLocationCode & "'" & _
										  "       and stockcode  = '" & txtItemCodeDelete & "'" & _
								   	      "       and fcba = '" & session("fcBA") &"' "
												  
						set rsQuantityOnHand = Server.CreateObject("ADODB.Recordset")
						rsQuantityOnHand.Open qQuantityOnHand, connstr
						if not rsQuantityOnHand.eof then
							QuantityOnHandStock = cdbl(rsQuantityOnHand("quantity"))
						else
							QuantityOnHandStock = 0
						end if
						'response.write qQuantityOnHand
						'response.write tempislesszero & " - " & QuantityOnHandStock & " , "  & originalQuantity
						
						If tempislesszero = "TRUE" Or (tempislesszero = "FALSE" And ((QuantityOnHandStock - originalQuantity) >= 0)) Then
							
                            q = "Delete from stocktranshist " & _
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
								"storecode = '" &  txtLocationCode & "' and " & _
								"stockcode = '" & txtItemCodeDelete & "' and " & _ 
								"fcba = '" & session("fcBA") &"' "
										
							connStr.execute(q)								
								
							q = "Delete from MANUALRECEIVEDTL_TAB where  OrderNo='" & TempOrderNo & "' and LINENOHIST='" & txtLineNoHist & _
								"' and fcBA='" & Session("fcBA") & "'"
									
							connStr.execute(q)						    
							
							q = " Update TRANSPORTTASKDTL_TAB  a set a.QUANTITYRECEIVED = a.QUANTITYRECEIVED - " &  originalQuantity & " " & _
								" where a.fcba = '" & txtFcbaSource & "' and a.orderno = '" & txtOrderSource & "'  and a.lineno = " & txtItemLineKeyDelete 
							
                           ' response.write q							
						    connStr.execute(q)	
						    if err.number=0 then
								connStr.CommitTrans
							End if
									
						Else
							err.number = 1001
							Err.Description = "Quantity tidak boleh Negatif atau Saldo belum didefinisikan"
						End If    	
         				
                    else
						err.number = 1001
						err.description = "Tanggal """ & DATECREATED & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
					End If
					
					cmdView	= "True"   							
		   	
        elseif Request("txtJenis")="New" then
			connstr.BeginTrans
					
			Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), DATECREATED)
			If  Cek_IsDailyClosing = "FALSE" then
			
			    Dim AdaHeader
				
				q= "select * from MANUALRECEIVE_TAB  " & _
				   " where OrderNo='" & TempOrderNo & "' and " & _
				   " fcBA='" & session("fcBA") & "'"									
				Set rsCek = Server.CreateObject("ADODB.Recordset")
				rsCek.open q , connstr
				If not rsCek.eof then
					AdaHeader ="Y"
				Else
					AdaHeader ="N"
				end if 
				rsCek.Close
				Set rsCek = Nothing
				
				
				If AdaHeader ="N"	Then	
				    '--handle Auto Number of ManualReceive
					If  Ucase(TempOrderNo) = Ucase("(NEW)") Then
						oSql =  " select nvl(controlnumber_api.get_nextnumber(upper('MANUALRECEIVED')" & _
								"                                         , " & year(textDate(DATECREATED)) & "" & _
								"                                         , " & month(textDate(DATECREATED)) & "" & _
								"                                         , '" & Session("fcBA") & "'" & _
								" ), 'FALSE') OrderNo from dual"
									
						'Response.Write oSql
										   
						Set oRs = Server.CreateObject("ADODB.Recordset")
						oRs.Open oSql, connstr, 3, 3
						If  Not oRs.EOF Then
							txtOrderNo = oRs("OrderNo")
						End If
						oRs.Close
						Set oRs = Nothing
						TempOrderNo = txtOrderNo
					End If
					
					'response.write OrderNo
							
					If  txtOrderNo = "FALSE" Then
						err.Number = 1001
						err.description = "<br/>Automatic Number Not Have Been Set, Call your Administrator!"
					End If
					
					
					q = "Insert into MANUALRECEIVE_TAB (OrderNo, OrderDate,FcbaSource,OrderSource, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
						"values ('" & TempOrderNo & "'," & " to_date('" & DATECREATED & "','mm/dd/yyyy'),'" & xCode & "','" & TransportTaskNo & "','" & _
						 Session("CodeName") & "','" & Session("CodeName") & "','" & _
						 Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI'))"
					
					'response.Write(q)
					
					connStr.execute(q)
				
				End if
				
				
				Dim TranportTaskDate_
				
				
							
				Dim n, ItemPrice, ItemAmount 
		  
				For n=1 to (cdbl(request("no"))-1) 
								
						' Insert the History of Received Inventory Transaction
												
						qLineMax =  " Select nvl(max(lineno),0) tempMaxLine from stocktranshist " & _
									" Where " & _
									" Fcba = '" & session("fcBA") & "'" 
													
						set rsqLineMax = Server.CreateObject("ADODB.Recordset")
						rsqLineMax.Open qLineMax, connstr
						if not rsqLineMax.eof then
							tempMaxLine = cdbl(rsqLineMax("tempMaxLine"))
						else
							tempMaxLine = 0
						end if 

						tempMaxLine = tempMaxLine + 1
						
						q= "select last_day(TRANSPORTTASK_APi.GET_RCVDATE('" & xCode & "', '" & request("Ordersource_"&TRIM(n)) & "' )) TranportTaskDate from dual "
						'response.Write(q)				
						Set rsTranportTask = Server.CreateObject("ADODB.Recordset")
						rsTranportTask.open q , connstr
						If not rsTranportTask.eof then
							TranportTaskDate_ =rsTranportTask("TranportTaskDate")
						end if 
						rsTranportTask.Close
						Set rsTranportTask = Nothing
						
													
									
						q = "select STOCKLEDGER_API.GetUnitPriceThisDate(to_date('" & TranportTaskDate_ & "','mm/dd/yyyy')  ,'" & request("STOCKCODE_"&TRIM(n))   & "','" & xCode & "')  itemprice from dual" 
 
                        'response.write q 						
						set rsitemprice = Server.CreateObject("ADODB.Recordset")
						rsitemprice.open q , connstr
						If not rsitemprice.eof then
							ItemPrice   = cdbl(rsitemprice("itemprice"))
							ItemAmount  = cdbl(rsitemprice("itemprice")) * cdbl(request("quantityreceived_"&TRIM(n)))
						Else
							ItemPrice   = 0
							ItemAmount  = 0
						end if 
						rsitemprice.Close
						Set rsitemprice = Nothing
									
						q = " Update TRANSPORTTASKDTL_TAB  a set a.price = " &  ItemPrice & ", a.amount = " & ItemAmount & " , QUANTITYRECEIVED = " &  request("quantityreceived_"&TRIM(n)) & " " & _
							" where a.fcba = '" & xCode & "' and a.orderno = '" & request("Ordersource_"&TRIM(n)) & "'  and a.lineno = " & request("LineNoKey_"&TRIM(n))  
									
						connStr.execute(q)
									
						q = "Insert into StockTransHist ( " & _
							" LINENO, FCBA, STOCKCODE, TRANSDATE, TRANSTYPE, STORECODE, SOURCETYPE, SOURCECODE, TARGETTYPE, TARGETCODE, QUANTITY, LASTUPDATE, LASTTIME) values ( "& _
							tempMaxLine & ",'" & _ 
							session("fcBA") & "','" & _
							request("STOCKCODE_"&TRIM(n)) &"'," & _
							"to_date('" & DATECREATED & "','mm/dd/yyyy') ,'" & _
							"MANUAL RECEIVED" & "','" & _
							request("xstoreCode_"&TRIM(n)) & "'," & _
							"'CB','BS','" & _
							request("locationtype_"&TRIM(n)) & "','" & _
							request("xstoreCode_"&TRIM(n)) & "'," & _
							request("quantityreceived_"&TRIM(n)) & ", sysdate, To_Char(sysdate,'HH24:MI') ) "
													  
						'response.Write(q)  
													  
						connStr.execute(q)
									
										
														
						q = "Insert into  MANUALRECEIVEDTL_TAB (OrderNo, LOCATIONTYPE, LOCATIONCODE, STOCKCODE, JOBCODE, QUANTITYRECEIVED, PRICERECEIVED, AMOUNTRECEIVED, " & _
							"  FCBASOURCE, ORDERSOURCE, LINENOSOURCE, QUANTITYSOURCE, LINENOHIST,FCENTRY, FCEDIT, FCIP, FCBA,LASTUPDATE, LASTTIME) " & _
							" values ('" & TempOrderNo & "','" & request("locationtype_"&TRIM(n)) & "','" & request("xstoreCode_"&TRIM(n)) &  _
							"','" & request("STOCKCODE_"&TRIM(n)) & "','" & request("Jobcode_"&TRIM(n)) & "'," & request("quantityreceived_"&TRIM(n)) & _
							"," & ItemPrice & "," & ItemAmount & ",'" & xCode & "','" & request("Ordersource_"&TRIM(n)) & "'," & request("LineNoKey_"&TRIM(n)) & "," & request("QuantitySource_"&TRIM(n)) & ","& tempMaxLine & "," & _
							"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
							Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI'))"
						
						'response.Write(q)
											
						connStr.execute(q)
									
						qTotalRecordOnHand = "select count(*) totalRecordOnHand from stockonhand " & _
											 "where " & _
											 "storecode = '" & request("xstoreCode_"&TRIM(n))  & "' and " & _												 
											 "stockcode = '" & request("STOCKCODE_"&TRIM(n)) & "' and " & _
											 "fcba = '" &  session("fcBA") & "'"

						set rsTotalRecordOnHand = Server.CreateObject("ADODB.Recordset")
						rsTotalRecordOnHand.Open qTotalRecordOnHand, connstr
						if not rsTotalRecordOnHand.eof then
							tempTotalRecordOnHand = cdbl(rsTotalRecordOnHand("totalRecordOnHand"))
						else
							tempTotalRecordOnHand = 0
						end if 
												
						if  tempTotalRecordOnHand = 0 then			
							' Just Insert the New Record with the following/received Quantity
							q = "Insert into stockonhand ( " & _
								" STORECODE, STOCKCODE, FCBA, QUANTITY, LASTUPDATE, LASTTIME ) values ( '" & _
								request("xstoreCode_"&TRIM(n))  & "','" & _
								request("STOCKCODE_"&TRIM(n)) & "','" & _
								session("fcBA") & "', '" & cdbl(request("quantityreceived_"&TRIM(n))) & "', sysdate, To_Char(sysdate,'HH24:MI') ) "
								connStr.execute(q)
													
						Else	  

						   q =  "update stockonhand " & _
								"set quantity = quantity + " & cdbl(request("quantityreceived_"&TRIM(n))) & " " & _ 
								", LASTUPDATE = sysdate " & _
								", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
								"where " & _
								"storecode = '" & request("xstoreCode_"&TRIM(n)) & "' and " & _
								"stockcode = '" & request("STOCKCODE_"&TRIM(n)) & "' and " & _ 
								"fcba = '" & session("fcBA") &"' "
														 
							connStr.execute(q)
						End If
									
				Next
							
				if  err.number=0 then
					connStr.CommitTrans
				End if
							
			Else
				err.number = 1001
				err.description = "Tanggal """ & DATECREATED & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
			End If
		   cmdView	= "True"
        End if
		cmdView	= "True"
 else
    cmdView	= request("cmdView")
 End if								 

 Dim rsVoucher,SudahClose
 set rsVoucher = server.CreateObject("Adodb.recordset")

 q="select * from MANUALRECEIVE_TAB  " & _
  " where OrderNo='" & TempOrderNo & "' and " & _
  " fcBA='" & session("fcBA") & "'"
SudahClose="N"
'response.write q
 rsVoucher.Open q, connstr
 if Not rsVoucher.Eof then
    DATECREATED=rsVoucher("OrderDate")
	xCode=rsVoucher("fcbasource")
	TransportTaskNo = rsVoucher("Ordersource")
	SudahClose="Y"
 End if
 rsVoucher.Close
 Set rsVoucher = Nothing
 
   

 
 'response.write xCode & TransportTaskNo
 
If Err.Number <> 0 then
	connStr.RollbackTrans
	Response.Write "<div class = ""error"">" & Err.Description & "</div>"
End if


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
'response.Write oSQLQuery
oRecordSet.Close
Set oRecordSet = Nothing
 %>
 <form name="frmEntry" method="post" action="ReceivedTask.Asp">
	<table class="data" width="100%">
		<tr>
		<td align="left">Order No</td>
		<td align="left"><input type="text" maxlength="50" name="TempOrderNo" size="30" value="<%=TempOrderNo%>"><font color="#FF0000">&nbsp;*</font></td></tr><tr>
		<td align="left">Order Date</td>
		<td align="left"><input class="datepicker" type="text" maxlength="10" name="DATECREATED" size="10" value="<%=DATECREATED%>">
		<font color="#FF0000">&nbsp;*</font>
		</td></tr>
		<tr>
		<td align="left">Fcba Source</td>
		<td align="left" readonly="RUE" > 
        <%if SudahClose="Y" then%>
		    <select name="xCode" ><option disabled="disabled" value="<%=xCode%>" Selected><%=xCode%></option></select>
		<%else%>   		
		    <select name="xCode" onchange="getordernumber(this.value);">
		    <% 
			dim rsCallCombo
			q = " select fccode,fcname " & _
							 "	from businessunit " & _
							 "	where fccompanycode = businessunit_api.get_CompanyCode('" & session("fcba") & "') " & _
							 "	and businessunit_api.Isactive(fccode)='TRUE'  and fccode <> '" & session("fcba") & "' order by orderdata "
							
			'response.write q
			Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
			rsCallCombo.CursorType = 3
			rsCallCombo.CursorLocation = 3
			rsCallCombo.Open q	,connstr			 
			if not rsCallCombo.eof then
			 Do While Not rsCallCombo.eof
			    %> 
				   <option value="<%=rsCallCombo("Fccode")%>" <% If Request("xCode") = rsCallCombo("Fccode") Then %> Selected <%End if%>><%=rsCallCombo("Fccode")%> : <%=rsCallCombo("fcname")%></option> 
				<%
				rsCallCombo.movenext
			 loop
			End if
			rsCallCombo.Close
			set rsCallCombo = Nothing
			%>
            </select>
	       <%End if%>
		<font color="#FF0000">&nbsp;*</font>
        </td>
		</tr>
		<tr>
		<td align="left">Transport Task No</td>
		<td align="left" id="TransportTaskNo"> 
		<%if SudahClose="Y" then  
		        %>
		    <select name="TransportTaskNo"><option disabled="disabled" value="<%=TransportTaskNo%>" Selected><%=TransportTaskNo%></option></select>
		<%else%>   
		    
			  
			  <%		
                     dim rst					 
					 rst = " select distinct  a.orderno from TRANSPORTTASKDTL_TAB  a where fcba='" & xCode & "'  and (quantitytransfered -quantityreceived ) > 0"
                      'response.write rst
					  Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
					  rsCallCombo.Open rst, connstr
					  if not rsCallCombo.eof then
							html = "<select name='TransportTaskNo'>"
							Do While Not rsCallCombo.eof
							html = html & "<option value='" & rsCallCombo("orderno") & "'>" & rsCallCombo("orderno") & "</option>" & _
								rsCallCombo.movenext
							loop
							html = html & "</select>" 
							else
							
							html = "<select name='TransportTaskNo'></select>"
					  End if
					  rsCallCombo.Close
					  set rsCallCombo = Nothing  
					  response.write(html)
			  
			  
		  End if %>	  
		</td>
		</tr>
		
		
		<tr>
			<td colspan="2"  align="left"> 
				<input Type="Hidden" Value="True" Name="cmdView">  
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">  
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">   	 
				<input Type="Submit" name="Add" value="Select">&nbsp;&nbsp;  
				<a href="ReceivedTask.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&TempOrderNo=<%=TempOrderNo%>&TransportTaskNo=<%=TransportTaskNo%>" target="_blank"><img src="IMAGES/Print.gif" border="0" title="Print Preview"><font face="Verdana" size="1">Print</font></a>&nbsp;&nbsp;
				<%If val_IsTrue = "Y" Then%>
					<a href="ReceivedTask.Asp?TempOrderNo=<%=TempOrderNo%>&DATECREATED=<%=DATECREATED%>&TransportTaskNo=<%=TransportTaskNo%>&txtJenis=Header"><img src="IMAGES/delete.gif" alt="Delete Data : <%=TempOrderNo%>" border="0" onClick="return confirm('Are you sure to delete');"></a>
				<%Else%>
					<img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18">
				<%End If%>
				<%
		dim dataModule,dataSubmodule,dataDetail
		dataModule 		= "Purchase"
		dataSubmodule 	= "Manual Receive "
		dataDetail 		= "NULL"
		%>
		<a class="jqplot-image-button" style="" href="ajax/hadirAjax.asp?dataModule=<%=dataModule%>&dataSubmodule=<%=dataSubmodule%>&dataDetail=<%=dataDetail%>"><i class="fa fa-check-square-o"></i>&nbsp;Hadir</a>
		<!-- END:HADIR -->
		</td> 
		</tr> 
    </table>    
 </form>
<%
If cmdView = "True" and LEN(TRIM(TempOrderNo)) > 4 and len(trim(TempOrderNo)) > 0 Then
	Dim rsMaRcv,intRecordsPerPage,cColor
	set rsMaRcv = server.CreateObject("Adodb.recordset")
	If SudahClose = "Y" Then
		q	= " select a.Linenosource, " & _
			  "		   a.ordersource, " & _
			  "		   a.stockcode, " & _
			  "		   stock_api.get_stockName(a.stockcode,a.fcba) Stockname, " & _
			  "		   a.quantityreceived, " & _
			  "		   a.PriceReceived PRICE, " & _
			  "		   a.amountreceived amount, " & _
			  "		   a.jobcode, " & _
			  "        job_api.get_jobcode_name(a.jobcode,a.fcba) Jobname,"& _
			  " 	   a.locationtype, " & _
			  "		   a.locationcode  " & _
			  "	from manualreceivedtl_tab a" & _
			  "	where fcba='" & session("fcba") & "' and a.orderno ='" & TempOrderNo & "' " & _
			  "	order by a.Linenosource"
	Else
		q	= " select a.lineno     Linenosource, " & _
			  "		   a.orderno    Ordersource, " & _
			  "		   a.stockcode, " & _
			  "		   stock_api.get_stockName(a.stockcode,a.fcbadistination) Stockname, " & _
			  "		   a.quantitytransfered QuantitySource,  " & _
			  "		   a.quantitytransfered quantityreceived," & _
			  "		   round(a.PRICE)  PRICE, " & _
			  "		   round(a.AMOUNT) AMOUNT, " & _
			  "		   stockgroup_api.Get_ControlJob(stock_api.get_StockGroup(a.Stockcode,a.fcbadistination),a.fcbadistination) jobcode," & _
			  "        job_api.get_jobcode_name(stockgroup_api.Get_ControlJob(stock_api.get_StockGroup(a.Stockcode,a.fcbadistination),a.fcbadistination),fcba) Jobname, " & _
			  "		   'ST' locationtype, " & _
			  "		   'ST1' locationcode" & _
			  "	from TRANSPORTTASKDTL_TAB  a" & _
			  "	where fcba='" & xCode & "'  and a.orderno ='" & TransportTaskNo & "'  and (quantitytransfered -quantityreceived ) > 0" & _
			  "	order by a.lineno"
    End If	
	
   'response.Write q	& dataSubmodule
   Set rsMaRcv = Server.CreateObject("ADODB.Recordset")
   rsMaRcv.CursorType = 3
   rsMaRcv.CursorLocation = 3
   rsMaRcv.Open q, connstr%>
   <form  id="manualreceive"  action="ReceivedTask.Asp" method="post" name="eForm">
   <table class="data" border="1" width="100%" >
    <%if SudahClose<>"Y" Then%>
	   <tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;">
	   <td align="center">No</td>
	   <td align="center">LineNoKey</td>
	   <td align="center">Order No Source</td>
	   <td align="center">Item Code</td>
	   <td align="center">Item Name</td>
	   <td align="center">Qty Source</td>
	   <td align="center">Price Source</td>
	   <td align="center">Amount Source</td>
	   <td align="center">Job Code</td>
	   <td align="center">Job Name</td>
	   <td align="center">Target Type</td>
	   <td align="center">Target Code</td>
	   <td align="center">Quantity Received</td>
	   <!--<td align="center">Amount</td>-->
   	   </tr>
	 <%Else%>
       <tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;">
	   <td align="center">No</td>
	   <td align="center">&nbsp;</td>
	   <td align="center">LineNoKey</td>
	   <td align="center">Order No Source</td>
	   <td align="center">Item Code</td>
	   <td align="center">Item Name</td>
	   <td align="center">Quantity Received</td>
	   <td align="center">Price Received</td>
	   <td align="center">Amount Received</td>
	   <td align="center">Job Code</td>
	   <td align="center">Job Name</td>
	   <td align="center">Target Type</td>
	   <td align="center">Target Code</td>
	   <!--<td align="center">Amount</td>-->
   	   </tr>
     	 
   <% End if
   
    Dim qkt : qkt = "quantityreceived_"
	
   if Not rsMaRcv.eof then
      nCountofField=rsMaRcv.fields.count-1
	  no=1
	  While Not rsMaRcv.EOF

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
						<a href="ReceivedTask.Asp?TempOrderNo=<%=TempOrderNo%>&ItemCode=<%=rsMaRcv("STOCKCODE")%>&linenosource=<%=rsMaRcv("linenosource")%>&DATECREATED=<%=DATECREATED%>&txtJenis=Details">
						<img src="IMAGES/delete.gif" alt="Delete Data : <%=rsMaRcv(1)%>" border="0" onClick="return confirm('Are you sure to delete');"></a>
					</td>
					<%Else%>
					<td>
						<img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18">
					</td>
					<%End If%>
					<%for i = 0 to nCountofField%>
						<td align="left" title="<%=rsMaRcv(i).Name%>" valign="top"><%=rsMaRcv(i)%></td>
					<%next%>
				<%End if%>
				
				<%if SudahClose<>"Y" Then%>
					<td align="Right" title="LineNoKey" valign="top"><%=rsMaRcv("Linenosource")%></td>
					<td align="left" title="Order Source" valign="top"><%=rsMaRcv("Ordersource")%></td>
					<td align="left" title="Item Code" valign="top"><%=rsMaRcv("STOCKCODE")%></td>
					<td align="left" title="Item Name" valign="top"><%=rsMaRcv("Stockname")%></td>
					<td align="left" title="QuantitySource" valign="top"><%=rsMaRcv("QuantitySource")%></td>
					<td align="Right" title="Price" valign="top"><%=FormatNumber(rsMaRcv("Price"),2)%></td>
					<td align="Right" title="Amount" valign="top"><%=FormatNumber(rsMaRcv("Amount"),2)%></td>
                    <td align="Right" title="Jobcode" valign="top"><%=rsMaRcv("Jobcode")%></td> 
					<td align="Right" title="Jobname" valign="top"><%=rsMaRcv("Jobname")%></td>
					<td align="Right" title="locationtype" valign="top"><%=rsMaRcv("locationtype")%></td>
                    <td align="left">        
							<select name="xstoreCode_<%=no%>" value="<%=request("xstoreCode_"&TRIM(no))%>"> 
								<%cmdCallCombo "Store",xstoreCode_&TRIM(no),"ACTIVATION", "Y"%>
							</select>
							<font color="#FF0000"></font>
					</td>
					<td align="Right" title="quantityreceived" valign="top"><%=rsMaRcv("quantityreceived")%></td>
					
					<input type="hidden" name="LineNoKey_<%=no%>" value="<%=rsMaRcv("Linenosource")%>"></td>
					<input type="hidden" name="Ordersource_<%=no%>" value="<%=rsMaRcv("Ordersource")%>"></td>
					<input type="hidden" name="STOCKCODE_<%=no%>" value="<%=rsMaRcv("STOCKCODE")%>"></td>
					<input type="hidden" name="QuantitySource_<%=no%>" value="<%=rsMaRcv("QuantitySource")%>"></td>
					<input type="hidden" name="Jobcode_<%=no%>" value="<%=rsMaRcv("Jobcode")%>"></td>
					<input type="hidden" name="locationtype_<%=no%>" value="<%=rsMaRcv("locationtype")%>"></td>
					<input type="hidden" name="quantityreceived_<%=no%>" value="<%=rsMaRcv("quantityreceived")%>"></td>
					
			<%End if%>
		  </tr>
		  <%
		  rsMaRcv.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   else
      response.Write("Data is empty")	  
   end if 
   rsMaRcv.Close
   set rsMaRcv = Nothing
   %>
	  <tr><td colspan="14" align="left">
		  <!-- diambil dari form yang diatas -->
	      <input type="hidden" name="TempOrderNo" value="<%=TempOrderNo%>">
		  <input type="hidden" name="DATECREATED" value="<%=DATECREATED%>">
		  <input type="hidden" name="xCode" value="<%=xCode%>">
		  <input type="hidden" name="TransportTaskNo" value="<%=TransportTaskNo%>">
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
   </form>
<%
End if  
 
'-------------akhir----------------------
  %>
 </div> </div>
  <!--#include file="bottom.jsp"-->

</body>
</html>
