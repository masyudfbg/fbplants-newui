<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="mm.inc"-->
<!--#include file="DailyValidation.inc"-->
<!--#include file="qtyOnHandFunction.inc"-->
<html>
<head>
<title>ePlaS - Stock Issue</title>
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

</script>
<script language="javascript" type="text/javascript">



                function showHint(str)
					{
				var xmlhttp;
				var StockName;
			
					if (str=="")
					  {
					  document.getElementById("announ").innerHTML="";
					  document.getElementById("StockName").innerHTML="";
					  return;
					  }
					if (window.XMLHttpRequest)
					  {// code for IE7+, Firefox, Chrome, Opera, Safari
					  xmlhttp=new XMLHttpRequest();
					  StockName=new XMLHttpRequest();
					  }
					else
					  {// code for IE6, IE5
					  xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");
					  StockName=new ActiveXObject("Microsoft.XMLHTTP");
					  
					  }
					xmlhttp.onreadystatechange=function()
					  {
					  if (xmlhttp.readyState==4 && xmlhttp.status==200)
						{
						document.getElementById("announ").innerHTML=xmlhttp.responseText;
						}
					  }
					xmlhttp.open("GET","ajax/qualityAjax.asp?ajaxUoM=quality&q="+str,true);
					xmlhttp.send();
					
					StockName.onreadystatechange=function()
						  {
						  if (StockName.readyState==4 && StockName.status==200)
							{
							document.getElementById("StockName").innerHTML=StockName.responseText;
							}
						  }
						StockName.open("GET","ajax/qualityAjax.asp?ajaxUoM=StockName&q="+str,true);
						StockName.send();
						
					}
					
					
				function JobCodeHint(str)
					{
					var jobname;
					
					if (str=="")
					  {
					  document.getElementById("jobname").innerHTML="";
					  return;
					  }
					if (window.XMLHttpRequest)
					  {// code for IE7+, Firefox, Chrome, Opera, Safari
					  jobname=new XMLHttpRequest();
					  }
					else
					  {// code for IE6, IE5
					  jobname=new ActiveXObject("Microsoft.XMLHTTP");
					  }
				    jobname.onreadystatechange=function()
						  {
						  if (jobname.readyState==4 && jobname.status==200)
							{
							document.getElementById("jobname").innerHTML=jobname.responseText;
							}
						  }
						jobname.open("GET","ajax/MANDAYS.asp?ajaxUoM=jobname&q="+str,true);
						jobname.send();
					}
					

</script>

 <script>
$(function() {
$( ".datepicker" ).datepicker();
});

function validateFloatKeyPress(el, evt) {
    var charCode = (evt.which) ? evt.which : event.keyCode;
    var number = el.value.split('.');
    if (charCode != 46 && charCode > 31 && (charCode < 48 || charCode > 57)) {
        return false;
    }
    //just one dot
    if(number.length>1 && charCode == 46){
         return false;
    }
    //get the carat position
    var caratPos = getSelectionStart(el);
    var dotPos = el.value.indexOf(".");
    if( caratPos > dotPos && dotPos>-1 && (number[1].length > 1)){
        return false;
    }
    return true;
}

//thanks: http://javascript.nwbox.com/cursor_position/
function getSelectionStart(o) {
	if (o.createTextRange) {
		var r = document.selection.createRange().duplicate()
		r.moveEnd('character', o.value.length)
		if (r.text == '') return o.value.length
		return o.value.lastIndexOf(r.text)
	} else return o.selectionStart
}
</script>
<script language="JavaScript"> 
var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}

var reloadSivCode = function () {
	var sivdate =$( ".datepicker" ).val();
	window.location.href = window.location.href.replace( /[\?#].*|$/, "?txtModule=Procurement&txtTable=IssueDetails&sivdateurl="+sivdate );
}

</script> 

 <script language="JavaScript">
        function PopupTType() {
			var txtttype = document.getElementById("locationtypecode").value;
			var txtdate  = document.getElementById("IssueDate").value;
			var newwin   = window.open("popup_loccode.Asp?txtSearch="+txtttype+"&txtdate="+txtdate+"&Mode=Print","newwindow","width = 700, height = 505, screenX = 0, screenY = 0, directories = 0, fullscreen = 0, location = 0, menubar = 0, scrollbars = 1, status = 0, toolbar = 1");
        }

</script>
</head>
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
 <%
 '-------------awal-----------------------
 Dim xCode,fcCode,cmdView,PoCode,IssueDate,ITEMCODE,QUANTITY,AMOUNT, LineNoHist
 Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis
 Dim Jobcode, LocationType, LocationCode, Remarks, Reference
 Dim qLineMax, rsqLineMax, tempMaxLine
 Dim qTotalRecordOnHand, rsTotalRecordOnHand, tempTotalRecordOnHand
 Dim qQuantityOnHand, rsQuantityOnHand, tempQuantityOnHand, QuantityOriginal
 Dim done, qq  
 Dim qLessZero, rsLessZero, tempislesszero
 Dim qLineUpper, rsqLineUpper, totalUpper
 Dim qQuantity, rsqQuantity, totalQuantity,LineNoKey

 Dim qStockMin, rsqStockMin, stockmin, tempStringError

 Dim monthprev,bulanprevsplit,tglprevsplit,thnprevsplit,tglnow,tglawal
 Dim tglbulanprev,bulanprev_split,tglprev_split,thnprev_split
 

 Dim rsReceiveDetail
 Dim v_CrossCekDetail, txtLineNoHist, originalQuantity
 Dim txtItemCodeDelete
 Dim qDeleteHist
 Dim strLineMax, rsstrLineMax, tempMaxLineNo , txtLineNoKey
 Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""
 Dim Cek_IsAccClosing : Cek_IsAccClosing = ""
 Dim ACCOUNTINGYEAR, ACCOUNTINGPERIOD,ACCOUNTINGYEARPERIOD,postingflg,postingflgtxt,closeflg_
 Dim modulFile,executeQry
 dim arr_dtlsiv_eof
 dim txtstorecode,txtstockcode,txtqty,txtrowstate


 txtTable =Request("txtTable")
 txtModule=Request("txtModule")
 txtSearch=Request("txtSearch")
 txtField =Request("txtField")
 
 if Request("fcba") <> "" then
	fcba = Request("fcba")
	else	
	fcba = session("fcba") 
 end if

 if request("fccOde")="" then
    fcCode="" '----- Nantinya Ambil dari Format Master
 else
    fcCode = request("fccOde")
 end if
 
 
 txtJenis = request("txtJenis") '---- Untuk Membedakan Jenis Data Yang akan diolah [insert,edit,delete]
 
 '---- Data Header
 if request("sivdateurl") <>"" then
   IssueDate   = request("sivdateurl")
 elseif Request("IssueDate")<>"" then
   IssueDate   = Request("IssueDate")
 else
   IssueDate   = textDate(date)
 end if

 if KonversiTanggalKeDbl(IssueDate) > KonversiTanggalKeDbl(textDate(date)) then
		IssueDate = textDate(date)
 end if
 
 xCode 		= Request("xCode")

 'monthprev 		 = DateAdd("m", -1, IssueDate)
 tglnow		 	 = IssueDate
 bulanprevsplit  = left(tglnow,2)
 tglprevsplit    = mid(tglnow,4,2)
 thnprevsplit    = right(tglnow,4)

 tglawal 		 = bulanprevsplit & "/" & "01" & "/" & thnprevsplit
 
 'tglbulanprev = DateAdd("m", -1, IssueDate)
 tglbulanprev = IssueDate
 bulanprev_split  = month(tglbulanprev)
 tglprev_split    = day(tglbulanprev)
 thnprev_split    = year(tglbulanprev)


 ACCOUNTINGPERIOD  = left(tglnow,2)
 ACCOUNTINGYEAR    = right(tglnow,4)
 
 
 '---- Data Details
 JobCode=Request("JobCode")
 doc_ref=Request("doc_ref")
 LocationType=Request("LocationType")
 LocationCode=Request("LocationCode")
 Remarks = Request("Remarks")
 Reference = Request("Reference")
 ITEMCODE=Request("ITEMCODE")
 LineNoKey=Request("LineNoKey")
 QUANTITY=Request("QUANTITY")
 AMOUNT =Request("AMOUNT")
 Linenohist = Request ("linenohist")
 
 

Cek_IsAccClosing = StatusAccClosingTgl(IssueDate,fcba)


if 	Cek_IsAccClosing = "FALSE" then  
 
 '--- Action ke Database di Oracle [ insert,edit,delete ]
 if Request("txtJenis")="New" or Request("txtJenis")= "Header" or Request("txtJenis")= "Details" or Request("txtJenis")="Edit" or Request("txtJenis")="Approved" Then
	
	if Request("txtJenis")= "Header" then
	   connstr.BeginTrans
	   Cek_IsDailyClosing = IsDailyClosing(fcba, IssueDate)
		If Cek_IsDailyClosing = "FALSE" then
			   If err.number=0 then
				  q="Delete from Issue where  SivCODE='" & fcCode & "' and fcBA='" & fcba & "'"
				  connStr.execute(q)
				  if err.number  = 0 then
					 connStr.CommitTrans
				  end if
			   end if
		else
			err.number = 1001
			err.description = "Tanggal """ & IssueDate & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
       cmdView	= "True"
	elseif Request("txtJenis")= "Posting" then
					
					strQueryStatus = "UPDATE issue SET POSTINGFLG='1',lastupdate=sysdate,fcedit='" & session("CodeName") & "' WHERE fcba = '" & fcba & "' AND sivcode = '" & fccode & "'"
					'response.write strQueryStatus & "<br>"
					connStr.execute(strQueryStatus)
					
					strQueryStatus = "BEGIN fac_journal_mam_api.MAM_SIS_DOCNO('" & fccode & "', TO_NUMBER(TO_CHAR(TO_DATE('" & IssueDate  & "','MM/DD/YYYY'),'YYYY')), " & _ 
			                         "TO_NUMBER(TO_CHAR(TO_DATE('" & IssueDate  & "','MM/DD/YYYY'),'MM')) , '" & fcba & "','System'); " & _ 
									 "END;"			
					'response.write strQueryStatus & "<br>"
					connStr.execute(strQueryStatus)
					
					
					strQueryStatus = "update BKMJOURNALHDR set postingflg=1,lastupdate=sysdate,fcedit='" & session("CodeName") & "' where fcba='" & fcba & "' AND docno='" & fccode & "'"
					'response.write strQueryStatus & "<br>"
					connStr.execute(strQueryStatus)
	
					strQueryStatus = "BEGIN fac_closing_accounting_api.ProsesClosing(TO_NUMBER(TO_CHAR(TO_DATE('" & IssueDate  & "','MM/DD/YYYY'),'YYYY')), " & _ 
			                         "TO_NUMBER(TO_CHAR(TO_DATE('" & IssueDate  & "','MM/DD/YYYY'),'MM')) , '" & fcba & "','System'); " & _ 
									 "END;"			
					'response.write strQueryStatus	& "<br>"			 
					connStr.execute(strQueryStatus)
		      
		cmdView	= "True"	
 	elseif Request("txtJenis")= "Details" then
	
	    connstr.BeginTrans
		
		Cek_IsDailyClosing = IsDailyClosing(fcba, IssueDate)
		If Cek_IsDailyClosing = "FALSE" then
	   
			q = "Select storecode,sivdate,rowstate from Issue where SivCODE='" & fcCode & _
				"' and fcBA='" & fcba & "'"
				
				
			Set rsReceiveDetail = Server.CreateObject("ADODB.Recordset")
			rsReceiveDetail.Open q, connstr
					
			IF Not rsReceiveDetail.eof then
				txtStoreCode = rsReceiveDetail("STORECODE")
				txtSivdate 		= rsReceiveDetail("sivdate")
				txtrowstate 		= rsReceiveDetail("rowstate")
			End if
			rsReceiveDetail.Close
			set rsReceiveDetail = Nothing		   
		   
		   
			q = "Select stockcode, linenohist, quantity, lineNoKey from IssueDetails where SivCODE='" & fcCode & _
				"' and LineNoKey='" & request("LineNoKey") & _
				"' and STOCKCODE='" & request("ItemCode") & _
				"' and fcBA='" & fcba & "' " & _
				"  and linenohist = " & Linenohist
			

			Set rsReceiveDetail = Server.CreateObject("ADODB.Recordset")
			rsReceiveDetail.Open q, connstr
			
		
					
			IF Not rsReceiveDetail.eof then
				txtLineNoHist = rsReceiveDetail("LINENOHIST")
				originalQuantity = rsReceiveDetail("QUANTITY")
				txtItemCodeDelete = rsReceiveDetail("stockcode")
				txtLineNoKey = rsReceiveDetail("LineNoKey")
			End if
			rsReceiveDetail.Close
			set rsReceiveDetail = Nothing			

			v_CrossCekDetail = cmdCheckHist(txtLineNoHist, Request("ItemCode"))

			
			'function untuk mengembalikan qty ketika ada proses delete
			'qrylessstockonhand	= cmdUpdateStockOnHand(txtStoreCode,txtItemCodeDelete,originalQuantity,"SIV","Delete")
			'qrylessstockledger 	= cmdUpdateStockLedger(txtStoreCode,txtItemCodeDelete,IssueDate,originalQuantity,"SIV","Delete")

			

				q =  "Delete from stocktranshist " & _
					 "where " & _
					 "lineno = " & txtLineNoHist & " and " & _
					 "fcba = '" & fcba & "'"
			
				'connStr.execute(q)
				ExecuteAndLog "IssueDetail.asp", q
							
									
				q="Delete from IssueDetails where  SivCODE='" & fcCode & "' and LineNoKey=" & LineNoKey & " and fcBA='" & fcba & "'" & _
				   "  and linenohist = " & Linenohist
				'connStr.execute(q)
				ExecuteAndLog "IssueDetail.asp", q



				if err.number  = 0 then
				   connStr.CommitTrans
				end if
		else
			err.number = 1001
			err.description = "Tanggal """ & IssueDate & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If

		cmdView	= "True"
	elseif Request("txtJenis")="New" then
	    connstr.BeginTrans	  
		Cek_IsDailyClosing = IsDailyClosing(fcba, IssueDate)
		If Cek_IsDailyClosing = "FALSE" then
		if Request("txtJenis") <> "Approved" then
		dim sisaqty
		sisaqty 	= getSisaStock(xCode,ITEMCODE,IssueDate)
		' response.write(sisaqty)
		' response.End
			 if cdbl(Quantity) <= cdbl(sisaqty) then
			  Dim rsPO
			  Set rsPO = Server.CreateObject("ADODB.Recordset")
			  rsPO.Open "Select sivcode from Issue where SIVCODE='" & fcCode & "'AND FCBA = '" & fcba & "'", connstr
			  IF rsPO.eof then	    
				 q="Insert into Issue (SIVCODE, DOC_REF, STORECODE, SIVDATE, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME, ROWSTATE, DOCUMENT_CLASSIFICATION) " & _
				 "values ('" & fcCode & "','" &  DOC_REF & "','" &  xCode & "', to_date('" & IssueDate & "','mm/dd/yyyy') " & _
				 ",'" &  Session("CodeName") & "','" & Session("CodeName") & "','" & _
				 Request.ServerVariables("REMOTE_ADDR") &"','" & fcba & "', sysdate, To_Char(sysdate,'HH24:MI'),'Planned','506')"
				connStr.execute(q)
				modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
				executeQry	= "Query executed:" & q
				ExecuteAndLog modulFile,executeQry
			  End if
			  rsPO.Close
			    
				
			
				qLessZero = "select islesszero from stock " & _	   
					"where " & _
					"fccode = '" & ITEMCODE & "' and " & _
					"fcba = '" & fcba & "'" 
				
				set rsLessZero = Server.CreateObject("ADODB.Recordset")
				rsLessZero.Open qLessZero, connstr
				if not rsLessZero.eof then
					tempislesszero = rsLessZero("islesszero")
				else
					tempislesszero = "FALSE"
				end if 
					


				qStockMin = "select nvl(min_value,0) stockmin from stock " & _
							"where fccode = '"& ITEMCODE &"'" & _
							"and fcba = '" & fcba & "'" 
					
					set rsqStockMin = Server.CreateObject("ADODB.Recordset")
					rsqStockMin.Open qStockMin, connstr
					if not rsqStockMin.eof then
						stockmin = cdbl(to_nol(rsqStockMin("stockmin")))
					else
						stockmin = 0
					end if 	
			

					if (Quantity < stockmin ) then 
						tempString = "Anda sudah melewati batas stock minimum, silahkan order barang lagi."				 
						response.Write  "<font color=#000000 pointsize=9 face=Arial, Helvetica, sans-serif>" & tempString & "</font>"
					' response.write quantity
					' response.write stockmin
					End if 
				
					qLineMax = "select nvl(max(lineno),0) tempMaxLine from stocktranshist " & _
					"where " & _
					"fcba = '" & fcba & "'" 
					
					set rsqLineMax = Server.CreateObject("ADODB.Recordset")
					rsqLineMax.Open qLineMax, connstr
					if not rsqLineMax.eof then
						tempMaxLine = cdbl(rsqLineMax("tempMaxLine"))
					else
						tempMaxLine = 0
					end if 

					tempMaxLine = tempMaxLine + 1
						
					
					q = "Insert into StockTransHist ( " & _
						" LINENO, FCBA, DOCUMENTNO, STOCKCODE, TRANSDATE, TRANSTYPE, STORECODE, SOURCETYPE, SOURCECODE, TARGETTYPE, TARGETCODE, QUANTITY, LASTUPDATE, LASTTIME) values ( "& _
						  tempMaxLine & ",'" & _ 
						  fcba & "','" & _
						  fccode & "','" & _
						  ITEMCODE &"'," & _
						  "to_date('" & IssueDate & "','mm/dd/yyyy') ,'" & _
						  "ST_ISSUED" & "','ST1'," & _
						  "'ST','" & _
						  xCode & "','" & _
						  LOCATIONTYPE & "','" & _
						  LOCATIONCODE & "'," & _
						  -1 * Quantity & ", sysdate, To_Char(sysdate,'HH24:MI') ) "
													
					connStr.execute(q)
					modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
					executeQry	= "Query executed:" & q
					ExecuteAndLog modulFile,executeQry

					
					
					' Added by NAF, 2012-12-01
					 strLineMax = " select nvl(max(linenokey),0) maxLineNoKey from IssueDetails " & _
							   " where " & _
							   " fcba = '" & fcba & "'"
							   
				  
				
					Set rsstrLineMax = Server.CreateObject("ADODB.Recordset")
					rsstrLineMax.CursorType = 3
					rsstrLineMax.CursorLocation = 3
					rsstrLineMax.Open strLineMax, connstr
					
					if not rsstrLineMax.eof then
						tempMaxLineNo = cdbl(rsstrLineMax("maxLineNoKey")) + 1
					end if
								
					rsstrLineMax.Close
					set rsstrLineMax = Nothing 
					' End of Added by NAF, 2012-12-01

					q=  "Insert into  IssueDetails (SIVCODE, STOCKCODE, LOCATIONCODE, LOCATIONTYPE, QUANTITY, PRICE, AMOUNT, " & _
						"JOBCODE, REFERENCE, REMARKS, FCENTRY, FCEDIT, FCIP, FCBA, linenohist, LineNoKey, LASTUPDATE, LASTTIME ) " & _
						"values ('" & fcCode & "','" & ITEMCODE & "','"  &  LOCATIONCODE & "','" & LOCATIONTYPE & "'" & _
						",'" & Quantity & "',0,0,'" & Jobcode & "','" &  Reference & "','" & Remarks & "','" & Session("CodeName")  & "','" & Session("CodeName") & "','" & _
						Request.ServerVariables("REMOTE_ADDR") &"','" & fcba & "',"& tempMaxLine & "," &tempMaxLineNo & ", sysdate, To_Char(sysdate,'HH24:MI') )"
					connStr.execute(q)
					modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
					executeQry	= "Query executed:" & q
					ExecuteAndLog modulFile,executeQry

						if err.number  = 0 then
							connStr.CommitTrans
							else
							connStr.RollbackTrans   
							response.Write  "<div class = ""error"">" & err.description & "</div>"
						end if

					done = 1		
			  else 
			 	 	err.number = 1001
			 	 	err.description = "Stock Kode Barang ini """ & ITEMCODE & """, Kurang dari 0 , Silahkan tambah stock dulu, atau input dengan kodebarang lain"
			 	 END IF
				End If
		else
			err.number = 1001
			err.description = "Tanggal """ & IssueDate & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
	
        cmdView	= "True"
	
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
		'Response.Write QUANTITY
	     connstr.BeginTrans
		
		Cek_IsDailyClosing = IsDailyClosing(fcba, IssueDate)
		If Cek_IsDailyClosing = "FALSE" then
		
				
					qDeleteHist = "delete stocktranshist " & _
					" where " & _
					" lineno = " & linenohist & " and " & _ 
					" fcba = '" & fcba & "'"
					connStr.execute(qDeleteHist)
					modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
					executeQry	= "Query executed:" & qDeleteHist
					ExecuteAndLog modulFile,executeQry

					qq = "Insert into StockTransHist ( " & _
						" LINENO, FCBA, DOCUMENTNO, STOCKCODE, TRANSDATE, TRANSTYPE, STORECODE, SOURCETYPE, SOURCECODE, TARGETTYPE, TARGETCODE, QUANTITY, LASTUPDATE, LASTTIME) values ( "& _
						  linenohist & ",'" & _ 
						  fcba & "','" & _
						  fccode & "','" & _
						  ITEMCODE &"'," & _
						  "to_date('" & IssueDate & "','mm/dd/yyyy') ,'" & _
						  "ST_ISSUED" & "','" & _
						  xCode & "'," & _
						  "'ST','" & _
						  xCode & "','" & _
						  LOCATIONTYPE & "','" & _
						  LOCATIONCODE & "'," & _
						  -1 * Quantity & ", sysdate, To_Char(sysdate,'HH24:MI') ) "
													
					connStr.execute(qq)
					modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
					executeQry	= "Query executed:" & qq
					ExecuteAndLog modulFile,executeQry
			   
					q = " Update IssueDetails set" & _
						"     Quantity = " & Quantity & _
						"     , REFERENCE = '" & Reference & "'" & _
						"     , JobCode = '" & JobCode & "'" & _
						"     , Locationtype = '" & Locationtype & "'" & _
						"     , Locationcode = '" & Locationcode & "'" & _
						"     , REMARKS = '" & Remarks & "'" & _
						"     , LASTUPDATE = sysdate " & _
						"     , LASTTIME = To_Char(sysdate,'HH24:MI') " & _
						" where SIVCODE = '" & fcCode & "' and StockCode = '" & ItemCode & "' and fcBA = '" & fcba & "' and linenokey ='"& linenokey &"'"
					connStr.execute(q)	
					modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
					executeQry	= "Query executed:" & q
					ExecuteAndLog modulFile,executeQry

				
					
					'function untuk mengembalikan qty ketika ada proses delete atau update
					'	qrylessstockonhand	= cmdUpdateStockOnHand(xCode,ITEMCODE,originalQuantity,"SIV","Update")
					'	qrylessstockonhand	= UpdateNewStockOnHand(xCode,ITEMCODE,Quantity,"SIV","Update")
					'	qrylessstockledger 	= cmdUpdateStockLedger(xCode,ITEMCODE,IssueDate,originalQuantity,"SIV","Update")
					'	qrylessstockledger 	= UpdateNewStockLedger(xCode,ITEMCODE,IssueDate,Quantity,"SIV","Update")

					if err.number  = 0 then
						connStr.CommitTrans
					end if
			
			cmdView	= "True"
		else
			err.number = 1001
			err.description = "Tanggal """ & IssueDate & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
			
        	cmdView	= "True"
    elseif Request("txtJenis")="Approved" then
		'query 
		
		arr_dtlsiv_eof =  true
		q = "SELECT H.STORECODE, D.STOCKCODE,D.QUANTITY " & _ 
		" FROM ISSUE H " & _ 
		" INNER JOIN  ISSUEDETAILS D ON H.SIVCODE=D.SIVCODE AND H.FCBA = D.FCBA " & _ 
		" WHERE H.SIVCODE ='" & fcCode & "'"
		set rsSivetail = Server.CreateObject("ADODB.Recordset")
		rsSivetail.Open q, connstr
		if not rsSivetail.eof then 
			arr_dtlsiv = rsSivetail.getrows()
			arr_dtlsiv_eof = false
		End IF

		rsSivetail.close : set rsSivetail = nothing
		if not arr_dtlsiv_eof then 
			for x=0 to ubound(arr_dtlsiv,2)
				txtstorecode 	= arr_dtlsiv(0,x)
				txtstockcode	= arr_dtlsiv(1,x)
				txtqty 			= arr_dtlsiv(2,x)

				'Response.Write txtstorecode & " - " & txtstockcode & " - " & txtqty
				'Response.end

				' modify jun (19 juni 2024)
				qTotalRecordOnHand  =   "select count(*) totalRecordOnHand from stockonhand " & _
										"where " & _
										"storecode = '" & txtstorecode & "' and " & _												 
										"stockcode = '" & txtstockcode & "' and " & _
										"fcba = '" &  fcba & "'"
				'Response.Write qTotalRecordOnHand
				set rsTotalRecordOnHand = Server.CreateObject("ADODB.Recordset")
				rsTotalRecordOnHand.Open qTotalRecordOnHand, connstr
				if not rsTotalRecordOnHand.eof then
						tempTotalRecordOnHand = cdbl(rsTotalRecordOnHand("totalRecordOnHand"))
				else
						tempTotalRecordOnHand = 0
				end if 
					
				If tempTotalRecordOnHand = 0 then
					
						' Tambahkan item stock dan qty di tabel stock onhand
						qryinsertstockOnhand 	= cmdInsertStockOnHand(txtstorecode,txtstockcode,txtqty,"SIV")
						
				Else	 
						' Update qty di tabel stock onhand
						qryUpdatestockOnhand 	= cmdUpdateStockOnHand(txtstorecode,txtstockcode,txtqty,"SIV","Insert")
				End If	
				
				'CEK ITEM DAHULU DI TABEL STOCKLEDGER
				'JIKA TIDAK ADA INSERT BARU, SELAIN ITU UPDATE
				dim qryStockLedger, rsStockLedger,temp,totalRecordStockLedger
				qryStockLedger = "SELECT COUNT(*) totalRecordStockLedger FROM STOCKLEDGER" & _ 
				" WHERE " & _ 
				" STORECODE = '" & txtstorecode & "' and " & _
				" ITEMCODE  = '" & txtstockcode & "' and " & _ 
				" year  	= to_char(to_date('" & IssueDate & "','mm/dd/yyyy'),'yyyy') and " & _ 
				" month 	= to_char(to_date('" & IssueDate & "','mm/dd/yyyy'),'mm') and " & _ 
				" FCBA 		= '" & fcba & "' "
				
				set rsStockLedger = Server.CreateObject("ADODB.Recordset")
				rsStockLedger.Open qryStockLedger, connstr
				if not rsStockLedger.eof then
						totalRecordStockLedger = cdbl(rsStockLedger("totalRecordStockLedger"))
				else
						totalRecordStockLedger = 0
				end if 
			

				'function ini digunakan untuk insert/update data di tabel stockledger
				' modify jun (19 juni 2024)
				If totalRecordStockLedger = 0 then
				
					'insert stockledger
					qryinsertstockledger 	= cmdInsertStockLedger(txtstorecode,txtstockcode,IssueDate,txtqty, "SIV")
				ELSE
				
					'update stockledger
					qryupdateStockLedger 	= cmdUpdateStockLedger(txtstorecode,txtstockcode,IssueDate,txtqty, "SIV", "Insert")
				END IF
			Next
		end if

		strQueryStatus = "UPDATE issue SET POSTINGFLG='1',rowstate='Approved',lastupdate=sysdate,fcedit='" & session("CodeName") & "' WHERE fcba = '" & fcba & "' AND sivcode = '" & fccode & "'"
		connStr.execute(strQueryStatus)
		modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
		executeQry	= "Query executed:" & strQueryStatus
		ExecuteAndLog modulFile,executeQry
		cmdView	= "True" 
	else
        cmdView	= "True"
	end if								 
	'Response.Write "sql " & q
 else
    cmdView	= request("cmdView")
 End if
else
    err.number = 1001
	err.description = "Accounting Period  " & ACCOUNTINGPERIOD & "/" & ACCOUNTINGYEAR & "  LOCKED !!!"
	cmdView	= request("cmdView")
End if  
 

 Dim rsPurchase
 set rsPurchase = server.CreateObject("Adodb.recordset")
 q="select to_char(SivDate,'mm/dd/yyyy') SivDate,StoreCode,postingflg,rowstate from Issue  " & _
  " where SivCode='" & fcCode & "' and " & _
  " fcBA='" & fcba & "'"

 rsPurchase.Open q, connstr
 postingflg = "0"
 if Not rsPurchase.Eof then
    IssueDate	= rsPurchase("SivDate")
	xCode		= rsPurchase("StoreCode")
	postingflg  = rsPurchase("postingflg")
	txtrowstate = rsPurchase("rowstate")

 End if

Cek_IsAccClosing = StatusAccClosingTgl(IssueDate,fcba)


IF postingflg = "1" THEN 
   postingflgtxt = "YA"
ELSE
   postingflgtxt	= "TIDAK"
END IF

 '-------------ISI Payment Form
 
Dim val_IsTrue : val_IsTrue = "N"
Dim oRecordSet, oSQLQuery
oSQLQuery = " select 'Y' istrue" & _
            " from periodcontrol" & _
	        " where fcba = '" & fcba & "'" & _
	        "       and to_date ('" & IssueDate & "', 'mm/dd/yyyy') >= currentdate" & _
	        "       and to_date ('" & IssueDate & "', 'mm/dd/yyyy') <= last_day(add_months(currentdate, 1))" & _
	        "       and upper(system) = upper('Stores')"
			
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

	' Added By Sendy untuk validasi menu 11/17/2018
	Dim Menu : Menu = "N"
	Dim oRecordSetMenu, oSQLQueryMenu
	oSQLQueryMenu = " SELECT 'Y' istrue FROM users_tab " & _
				" WHERE fccode = '" & Session("CodeName") & "' AND MENU_ISSUE = 'TRUE'" 
	' response.write oSQLQueryMenu
	set oRecordSetMenu = server.CreateObject("Adodb.recordset")
	oRecordSetMenu.Open oSQLQueryMenu, connstr
	If Not oRecordSetMenu.Eof then
		Menu = oRecordSetMenu("istrue")
	Else
		Menu = "N"
	End if


'QUERY GET AUTHORIZE
dim isApproved,rsApproved
set rsApproved = server.CreateObject("Adodb.recordset")
q = " SELECT COUNT(1) total" & _
	" FROM authorizelevel al" & _
	" INNER JOIN authorizestructure aus ON aus.authorizelevelcode = al.fccode" & _
	" INNER JOIN authorizetemplate at ON aus.authorizetemplatecode = at.fccode" & _
	" INNER JOIN authorizelevel_users alu ON (alu.authorizelevelcode = al.fccode)" & _
	" WHERE at.document_type = 'SI' and at.fccode='506' " & _
	"       AND al.isactive = 'TRUE'" & _
	"       AND al.finallevel = 'TRUE'" & _
	"		AND alu.isactive = 'TRUE'" & _
	"       AND '"& session("CodeName") &"' = '" & session("CodeName") & "'" & _
	"       AND aus.fcusercode = '" & session("CodeName") & "'" & _
	"       AND aus.fcba = '" & session("fcba") & "'"
' response.write q
rsApproved.Open q, connstr
if Not rsApproved.Eof then
	isApproved =  cdbl(to_nol(rsApproved("total")))
else
	isApproved = 0
End if
rsApproved.Close : set rsApproved = Nothing 


 %>
 <form id="frmEntry" name="frmEntry" method="post" action="IssueDetails.Asp">
	<table class="data" width="100%">
		<tr>
			<td align="left">Tahun</td>
			<%
			'response.write fcCode
			if fcCode = "" then
				q = " select " & _
				    " Decode(Nvl(Max(cn.PREFIX),'-'), '-', '" & fcba & "'|| '/SIV/' || TO_CHAR(TO_DATE('"& IssueDate &"','mm/dd/yyyy'),'YY') ||'/'|| LPAD(NVL(MAX(regexp_substr(iss.SIVCODE, '[^/]+', 1, 4)),0000)+0001,4,0)," & _
				    " MAX(cn.PREFIX) || LPAD(NVL(MAX(regexp_substr(iss.SIVCODE, '[^/]+', 1, 4)),0000)+0001,4,0)) DOCNUMBER  " & _
					" from  " & _
					" controlnumber cn  " & _
					" left join Issue iss on cn.fcba = iss.fcba and cn.accountingyear = to_number(to_char(iss.SIVDATE,'yyyy'))  " & _
					" where cn.fcba ='" & fcba & "' " & _
					" and cn.LU_NAME ='INVENTORYISSUE' "&_
					" and cn.accountingyear =  TO_NUMBER(TO_CHAR(TO_DATE('"& IssueDate &"','mm/dd/yyyy'),'yyyy')) " & _
					" and cn.accountingperiod = TO_NUMBER(TO_CHAR(TO_DATE('"& IssueDate &"','mm/dd/yyyy'),'mm')) and NOT REGEXP_LIKE (iss.SIVCODE,'^BKM|^MBKM') "
				'Response.write q
				set rsPONo = server.CreateObject("Adodb.recordset")
				rsPONo.Open q, connstr
				if Not rsPONo.Eof then
					fcCode = rsPONo("DOCNUMBER")
				End if
			End if
			%>
			<td align="left"><input type="text" maxlength="30" name="fcCode" size="30" value="<%=fcCode%>"><font color="#FF0000">&nbsp;*</font></td></tr>
		<tr>
			<td align="left">Jabatan</td>
			<td align="left"><input type="text" maxlength="50" name="doc_ref" size="30" value="<%=doc_ref%>"></td>
		<tr>
		<td align="left">Status</td>
		<td align="left">      
		<select name="xCode" id="xcode" onchange="get_Allowentry()" value="<%=xCode%>
	    <%cmdCallCombo "Store",xCode,"",""%>
        </select><font color="#FF0000">&nbsp;*</font>
        </td>
		</td>
		</tr>
		<tr>

		<tr><td colspan="2"  align="left">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Hidden" Value="<%=postingflg%>" id="postingflg_">
		<input Type="Hidden" Value="<%=closeflg_%>" id="closeflg_">
		
		
		<input Type="Submit"  name="Add" id="Add" value="Add">&nbsp;&nbsp;
        <input Type="Submit"  name="Add" id="Add" value="Select">&nbsp;&nbsp;
	    <% IF LEFT(fcCode,3) <> "BKM" THEN %>
		<% if isApproved <> 0 then %>
		      <a id="poshdr" href="?fcCode=<%=fcCode%>&IssueDate=<%=IssueDate%>&fcba=<%=fcba%>&txtJenis=Approved" onClick="return confirm('Are you sure to Approved Stock Issue this data :  <%=fcCode%> ?');">Approved</a>
		<%Else%>
			<%IF txtrowstate = "Approved" then %>
				<img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18">
			<%End If%>
		<%End If%>
	
		<% END IF%>				
		<%
		dim dataModule,dataSubmodule,dataDetail
		dataModule 		= "Stores"
		dataSubmodule 	= "Stock Issue"
		dataDetail 		= "NULL"
		%>
		
		</td>
		</tr>
    </table>  
 </form>
<%
if cmdView = "True" AND LEN(TRIM(fcCode)) > 4 and len(trim(xCode)) > 0 Then
   Dim rsPurchaseOrder,intRecordsPerPage,cColor
   set rsPurchaseOrder = server.CreateObject("Adodb.recordset")
   q="select SIVCODE, LINENOKEY, STOCKCODE, STOCK_API.get_stockname(STOCKCODE,FCBA) STOCKNAME, JOBCODE, JOB_API.get_jobcode_name(JOBCODE,FCBA) JOBNAME, LOCATIONTYPE, LOCATIONCODE, QUANTITY, PRICE, AMOUNT, LINENOHIST, REFERENCE, REMARKS, FCENTRY, FCEDIT, FCIP, FCBA  from IssueDetails  " & _
     " where SivCode='" & fcCode & "' and " & _
     " fcBA='" & fcba & "'"
   'response.Write q	
   Set rsPurchaseOrder = Server.CreateObject("ADODB.Recordset")
   rsPurchaseOrder.CursorType = 3
   rsPurchaseOrder.CursorLocation = 3
   rsPurchaseOrder.Open q, connstr%>
   <form action="IssueDetails.Asp" method="post" name="eForm" id="eForm">
   <table class="data" cellpadding="0" cellspacing="0" width="100%">
		<tr><td align="left">Item Code</td>
		<td align="left">
		<%IF txtJenis="Edit" Then%>
		     <input id="itemcode" type="text" onkeyup="showHint(this.value)" size="15" name="ITEMCODE" style="background:#dddddd" readonly value=<%=ITEMCODE%>  >&nbsp;<font color="#FF0000">* <a href="javascript:launchwin('popup_stock_new.asp?fcba=<%=fcba%>&IssueDate=<%=IssueDate%>&Mode=Print','newwindow','width='+(1000)+',height='+(705)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')">Stock<span class="satuan_item" id="StockName">&nbsp;...</span></a>
			<tr><td align="left">Job Code</td><td align="left"><input type="text"  id="jobcode" size="15" name="Jobcode" onkeyup="getTargettype()"  style="background:#dddddd" readonly value=<%=Jobcode%>  >&nbsp;<font color="#FF0000">*
			
				<a href="javascript:launchwin('popup_activity.asp','newwindow','width='+(1000)+',height='+(705)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')">Job <span class="satuan_item" id="jobname"></span></a>
			
			</tr>
			<tr>
				<td align="left">Target Type</td><td align="left"><input type="text" size="15" id="locationtypecode"  style="background:#dddddd" readonly  name="LocationType" value=<%=LocationType%>  >&nbsp;</td></tr>
			<tr>
				<td align="left">Target Code</td><td align="left"><input type="text" size="15" id="locationcode" name="LocationCode" style="background:#dddddd"  readonly value=<%=LocationCode%>  >&nbsp;<font color="#FF0000">*<a href="javascript:PopupTType('<%=fcba %>')">
                <font face="Verdana" size="1">Location Code</font> <span class="satuan_item" id="locationname">&nbsp;</span></a></td></tr>


			
			<input type="hidden" size="15" name="LineNoKey" value="<%=LineNoKey%>" readonly="True" style="background:#999999">
			<input type="hidden" size="15" name="Linenohist" value="<%=Linenohist%>" readonly="True" style="background:#999999">
			
		<%else%>
		    <input id="itemcode" type="text"  onkeyup="showHint(this.value)" size="15" name="ITEMCODE" style="background:#dddddd" readonly value=<%=ITEMCODE%> >&nbsp;<font color="#FF0000">* <a href="javascript:launchwin('popup_stock_new.asp?fcba=<%=fcba%>&IssueDate=<%=IssueDate%>','newwindow','width='+(1000)+',height='+(705)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')">Stock<span class="satuan_item" id="StockName">&nbsp;...</span></a>
			<tr>
				<td align="left">Job Code</td>
				<td align="left">
					<input type="text" id="jobcode" size="15"  name="Jobcode" readonly  style="background:#dddddd"  value="<%=Jobcode%>"  >&nbsp;<font color="#FF0000">*
			
			<a href="javascript:launchwin('popup_activity.asp','newwindow','width='+(1000)+',height='+(705)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')">Job <span class="satuan_item" id="jobname"></span></a>
			
			</tr>
			<tr><td align="left">Target Type</td><td align="left"><input type="text" size="15" id="locationtypecode"  style="background:#dddddd" readonly name="LocationType" value=<%=LocationType%>>&nbsp;</td></tr>
			
			<tr><td align="left">Target Code</td><td align="left"><input type="text" size="15" id="locationcode" name="LocationCode" readonly style="background:#dddddd"  value=<%=LocationCode%>>&nbsp;<font color="#FF0000">*<a href="javascript:PopupTType('<%=fcba %>')">
                <font face="Verdana" size="1">Location Code</font> <span class="satuan_item" id="locationname">&nbsp;</span></a></td></tr>
			
			<input type="hidden" size="15" name="LineNoKey" value="<%=LineNoKey%>" readonly="True" style="background:#999999">
			<input type="hidden" size="15" name="Linenohist" value="<%=Linenohist%>" readonly="True" style="background:#999999">
		<%End if%>
		</tr>
		<tr><td align="left">Quantity</td><td valign="center" align="left"><input class="quantity" type="text" size="15" name="QUANTITY" id="qtyissue" onkeyup="validSisaStock()" onkeypress="return validateFloatKeyPress(this,event);" value="<%=QUANTITY%>">&nbsp;<font color="#FF0000">*</font><span id="announ" class="satuan_item" >...</span></td></tr>
		<tr><td align="left">REFERENCE</td><td align="left"><input type="text" size="30" maxlength = "30" name="REFERENCE" value="<%=REFERENCE%>"></td></tr>
		<tr><td align="left">REMARKS</td><td align="left"><textarea  name="REMARKS" maxlength = "100" cols="50" rows="3"><%=REMARKS%></textarea></td></tr>
	    <TR><td colspan="2" align="left">
		  <!-- diambil dari form yang diatas -->
	      <input type="hidden" name="fcCode" value="<%=fcCode%>">
		  <input type="hidden" name="IssueDate" value="<%=IssueDate%>">
		  <input type="hidden" name="xCode" value="<%=xCode%>">
		  <input type="hidden" name="doc_ref" value="<%=doc_ref%>">
		  
		  <input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		  <input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		  <% if postingflg ="0" and Cek_IsAccClosing="FALSE" then%>
		  <%IF Request("txtJenis")="Edit" then%>
		       <input type="hidden" name="txtJenis" value="Edit">
		       <input Type="Submit" name="Add" value="Update">
		  <%Else%>
		       <input type="hidden" name="txtJenis" value="New">
		       <input Type="Submit" name="Add" value="Add">
		  <%End if%>
		  <%End if%>
		  
		  <input Type="Reset" name="Reset" value="Reset">
		  <input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
	  </tr>
   </table>

   <table class="data">
	<BR>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>	
	<TD align="left">ISSUE CODE</TD>
	<TD align="left">LineNoKey</TD>	
	<TD align="left">STOCK CODE</TD>
	<TD align="left">STOCK NAME</TD>
	<TD align="left">JOB CODE</TD>
	<TD align="left">JOB NAME</TD>
    <TD align="left">LOC. TYPE</TD>	
	<TD align="left">LOC. CODE</TD>	
	<TD align="left">QUANTITY</TD>	
	<%If Menu = "Y" Then%>
	<TD align="left">PRICE</TD>
	<TD align="left">AMOUNT</TD>
	<%End If%>	
	<TD align="left">Line No (history)</TD>		
	<TD align="left">Reference</TD>	
	<TD align="left">Remarks</TD>	
	</tr>   
   <%
   if Not rsPurchaseOrder.eof then
      nCountofField=rsPurchaseOrder.fields.count-1
	  no=1
	  While Not rsPurchaseOrder.EOF

		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
		  <td align="left" valign="top"><%=NO%></td>
			<!--  Validation Period Control In Here -->
			<% 	if postingflg ="0" and Cek_IsAccClosing="FALSE" then%>
				<td valign="middle" width="30">
					<a  href="IssueDetails.Asp?fcCode=<%=rsPurchaseOrder("SivCode")%>&LineNoKey=<%=rsPurchaseOrder("LineNoKey")%>&ItemCode=<%=rsPurchaseOrder("StockCode")%>&Quantity=<%=rsPurchaseOrder("Quantity")%>&JobCode=<%=rsPurchaseOrder("JobCode")%>&LocationType=<%=rsPurchaseOrder("LocationType")%>&LocationCode=<%=rsPurchaseOrder("LocationCode")%>&LineNoHist=<%=rsPurchaseOrder("LineNoHist")%>&Reference=<%=rsPurchaseOrder("Reference")%>&Remarks=<%=rsPurchaseOrder("Remarks")%>&txtJenis=Edit">
					<img src="IMAGES/edit.gif" alt="Edit Data : <%=rsPurchaseOrder(1)%>" border="0"></a>
				</td>
				<td valign="middle" width="10">
					<a  href="IssueDetails.Asp?fcCode=<%=rsPurchaseOrder("SivCode")%>&LineNoKey=<%=rsPurchaseOrder("LineNoKey")%>&ItemCode=<%=rsPurchaseOrder("StockCode")%>&LineNoHist=<%=rsPurchaseOrder("LineNoHist")%>&IssueDate=<%=IssueDate%>&txtJenis=Details">
					<img src="IMAGES/delete.gif" alt="Delete Data : <%=rsPurchaseOrder(1)%>" border="0" onClick="return confirm('Are you sure to delete');"></a>
				</td>
			<%else%>
					<td valign="middle" width="30">
						<img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18">
					</td>
					<td valign="middle" width="10">
						<img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18">
					</td>
			<% end if%>	
			<!-- END Validation Period Control In Here -->
		  			<td align="center" title="SivCode" valign="top"><%=rsPurchaseOrder("SivCode")%></td>
					<td align="center" title="LineNoKey" valign="top"><%=rsPurchaseOrder("LineNoKey")%></td>
					<td align="center" title="StockCode" valign="top"><%=rsPurchaseOrder("StockCode")%></td>
					<td align="center" title="StockCode" valign="top"><%=rsPurchaseOrder("StockName")%></td>
					<td align="center" title="JobCode" valign="top"><%=rsPurchaseOrder("JobCode")%></td>
					<td align="center" title="JobCode" valign="top"><%=rsPurchaseOrder("JobName")%></td>
					<td align="center" title="LocationType" valign="top"><%=rsPurchaseOrder("LocationType")%></td>
					<td align="center" title="LocationCode" valign="top"><%=rsPurchaseOrder("LocationCode")%></td>
					<td align="Right" title="Quantity" valign="top"><%=FbpNum2(rsPurchaseOrder("Quantity"))%></td>
					<%If Menu = "Y" Then%>
					<td align="center" title="Price" valign="top"><%=rsPurchaseOrder("Price")%></td>
					<td align="center" title="Amount" valign="top"><%=rsPurchaseOrder("Amount")%></td>
					<%End If%>
					<td align="center" title="LineNoHist" valign="top"><%=rsPurchaseOrder("LineNoHist")%></td>	
					<td align="left" title="Reference" valign="top"><%=rsPurchaseOrder("Reference")%></td>	
					<td align="left" title="Remarks" valign="top"><%=rsPurchaseOrder("Remarks")%></td>	    
		  </tr>
		  <%
		  rsPurchaseOrder.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   else
      response.Write("Data is empty")	  
   end if 
   rsPurchaseOrder.Close
   set rsPurchaseOrder = Nothing%>
   </div>
   </TABLE>
   </form>
<%End if%>  
<%'-------------akhir----------------------%>
</div></div>
  <!--#include file="bottom.jsp"-->

</body>
</html>

<script>
$(function() {	

	//get_AllowentryIssue();
	get_Allowentry();
	
});

function validSisaStock()
{
	var itemcode 	=  $('#itemcode').val();
	var storecode 	=  $('#xcode').val();
	var qtyissue 	=  $('#qtyissue').val();
	if (storecode==""){
		alert("Gudang Belum Dipilih..!!");
	}else{
		if (itemcode != "") {
			$.post("validasi_stock.asp?itemcode=" + itemcode + "&storecode=" + storecode, function(data) {
				$('#msg').empty();
				var json  	= JSON.parse(data);
				var sisastock = parseFloat(json.results[0].quantity);
				
				if(qtyissue>sisastock){
					alert("Stock Tidak Mencukupi...!!");
					$('#qtyissue').val(0);
					return 0;
				}
							
			});
			
		} 
	}
}


//function lock inputan dari periode akunting
function get_AllowentryIssue()
{
	var tglbkm = $('#IssueDate').val();
	var postingflg =  $('#postingflg_').val();
	if (tglbkm != "") {
		$.post("acc_allow_entry.asp?tglbkm=" + tglbkm, function(data) {
			$('#bodytable').empty();
			var json  = JSON.parse(data);
			var closingStatus = json.results[0].msg;
			if(closingStatus=="FALSE"){
				if (postingflg == 1) 
				{
					$('#delhdr').hide();
				//	$('#poshdr').hide();
					$(':input[type="text"]').prop('disabled', true); 
					$(':input[type="Submit"]').prop('disabled', true);
				}
				else
				{
					$('#eForm').show();
					$('#delhdr').show();
					//$('#poshdr').show();
					$('#Add').show();
				}
				
				
			}else{
				alert("PERIODE SUDAH DI CLOSE !!!");
				$('#eForm').hide();
				$('#delhdr').hide();
				//$('#poshdr').hide();
				$('#Add').hide();
				
			}
			
		});
		
	} 
}

//validasi lock inputan dri H+3
function get_Allowentry()
{
	var tglbkm = $('#IssueDate').val();
	var postingflg_ = $('#postingflg_').val();
	if (tglbkm != "") {
		$.post("bkm_allow_entry.asp?tglbkm=" + tglbkm, function(data) {
			$('#bodytable').empty();
			var json  = JSON.parse(data);
			var closingStatus = json.results[0].msg;
			if(closingStatus=="FALSE"){
				if(postingflg_==1){
					$('#deletehdr').hide();
					$('#poshdr').hide();
				}else{
					$('#add').show();
					$('#eForm').show();
					$('#deletehdr').show();
					$('#poshdr').show();
				}
			}else{
				alert("PERIODE SUDAH DI CLOSE !!!");
				$('#add').hide();
				$('#eForm').hide();
				$('#deletehdr').hide();
				$('#Add').hide();
				$('#poshdr').hide();
			}
			
		});
		
	} 
}

function getTargettype()
{
	$('#locationtypecode').empty();
	//$('#locationcode').empty();
	var jobcode = $('#jobcode').val();
	//alert(jobcode);
	if (jobcode != "") {
		$.post("bkm_viewdata.asp?jobcode=" + jobcode + "&idview=" + 33, function(data) {
			var json  	= JSON.parse(data);
			var job_category = json.results[0].job_category;
			$('#locationtypecode').val(job_category);
		
		});
	} 

}

</script>