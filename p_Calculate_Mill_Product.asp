<%Option Explicit%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="mm.inc"-->
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
%>
<html>
<head>
	<title>ePlaS - Calculate Mill Production</title>
	<script type="text/javascript" language="JavaScript1.2" src="stlib.js"></script>
	<script language="JavaScript" src="date-picker.js"></script>
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
<%
Dim strNormalisasiProduksi
Dim strSQLPreParameter, rsSQLPreParameter,v_txtFromDayDate, v_txtToDayDate, v_txtfcba, v_txtyear, v_txtmonth,BRG
Dim  strSQL,strSQL2,strSQL3,Rsprd,Rsprd2,Rsprd3,q
Dim  sql1,v_txtprodcode 
Dim  snd_today,snd_yesterday,sales_today,prod_today
Dim TimeAwal,TimeAkhir,Hitung,Lama,CHITNO
TimeAwal=Time()
Dim Startingtime, Endingtime
Startingtime = FormatDateTime(CDate(Now))

Response.Write "Current Period for Production is " & FormatDateTime(Session("v_Delivery"),1) & "<br/>"
Response.Write "Please Wait.." & "<br/>"
Response.Write "Start ..." & Time() & "<br/>" 

' Normalisasi Produksi					    
strNormalisasiProduksi	=	" DELETE FROM MILL_PRODUCTION MP" & _
							" WHERE" & _
							" TO_NUMBER(TO_CHAR(TDATE,'YYYY')) = " & Year(Session("v_Delivery")) & " AND" & _
							" TO_NUMBER(TO_CHAR(TDATE,'MM')) = " & Month(Session("v_Delivery")) & " AND" & _
							" MP.FCBA = '" & Session("fcBA") & "'"
							
' strNormalisasiProduksi	=	" DELETE FROM MILL_PRODUCTION MP" & _
							' " WHERE" & _
							' " TO_NUMBER(TO_CHAR(TDATE,'YYYY')) = " & Year(Session("v_Delivery")) & " AND" & _
							' " TO_NUMBER(TO_CHAR(TDATE,'MM')) = " & Month(Session("v_Delivery")) & " AND" & _
							' " MP.FCBA = '" & Session("fcBA") & "'"
							
connStr.execute(strNormalisasiProduksi)
'Response.Write strNormalisasiProduksi


strSQL	=	" SELECT PS.READINGDATE" & _
			"	, TO_CHAR(PS.READINGDATE,'MM/DD/YYYY') DNOTE" & _
			"	, TO_CHAR(PS.READINGDATE-1,'MM/DD/YYYY') H_MINUS" & _
			"	, PS.PSTORAGE_STORAGECODE" & _
			"	, NVL(PS.QUANTITY,0) QUANTITY" & _
			"	, PS.FCBA" & _
			"	, MILL_PRODUCTION_API.GET_PRODCODE(PS.PSTORAGE_STORAGECODE,PS.FCBA) PRODCODE , sysdate lastupdate, to_char(sysdate,'HH24:MI') lasttime  " & _
			" FROM PRODUCTSTORAGEREADING PS" & _
			" WHERE " & _
			" PS.PSTORAGE_STORAGECODE||FCBA IN (SELECT STORAGECODE||FCBA  FROM PRODUCTSTORAGE) AND" & _
			" TO_NUMBER(TO_CHAR(PS.READINGDATE,'YYYY')) = " & Year(Session("v_Delivery")) & " AND" & _
			" TO_NUMBER(TO_CHAR(PS.READINGDATE,'MM')) = " & Month(Session("v_Delivery")) & "  AND" & _
			" PS.FCBA = '" & Session("fcBA") & "'"
			

'Response.Write strSQL
Set Rsprd = Server.CreateObject("ADODB.Recordset")
Rsprd.CursorType = 3
Rsprd.CursorLocation = 3
Rsprd.Open strSQL, connstr


if not Rsprd.EOF then
  
   while not Rsprd.EOF
	  snd_today		=	Rsprd("quantity") 
	  snd_yesterday	=	0
	  sales_today	=	0
	  'Response.Write "Production Date : " & Rsprd("readingdate")  & " , prod code : " & Rsprd("prodcode")  & " , storage code : " & Rsprd("pstorage_storagecode")  & "<br><br>" 
	 ' Response.Write "sounding today : " & snd_today  & "<br><br>"
		 strSQL2 = " select nvl(ps.quantity,0) quantity from productstoragereading ps " &_
					" where ps.fcba= '" & Session("fcBA") &"' " &_
					" and to_char(ps.readingdate,'MM/DD/YYYY') = '" & Rsprd("h_minus") & "'  and ps.pstorage_storagecode= '" & Rsprd("pstorage_storagecode") & "'"
		  set Rsprd2 = Server.CreateObject("ADODB.Recordset")
		  Rsprd2.CursorType = 3
		  Rsprd2.CursorLocation = 3
		  Rsprd2.Open strSQL2, connstr
		  'Response.Write strSQL2 & "<br>" 
		  if not Rsprd2.eof then
			 snd_yesterday = Rsprd2("quantity")
			' Response.Write "sounding Yesterday : " & snd_yesterday  & "<br><br>"
		  end if  
		 
		  strSQL3 = "  select nvl(sum(nvl(dn.quantity,0)),0) quantity from deliverynote dn where dn.fcba='" & Session("fcBA") &"' and to_char(dn.dnote_date,'MM/DD/YYYY') = '" & Rsprd("dnote") & "' and dn.storagecode = '" & Rsprd("pstorage_storagecode") & "'"
		  set Rsprd3 = Server.CreateObject("ADODB.Recordset")
		  Rsprd3.CursorType = 3
		  Rsprd3.CursorLocation = 3
		  Rsprd3.Open strSQL3, connstr
		 'Response.Write strSQL3 & "<br>" 
		  if not Rsprd2.eof then
			 sales_today = Rsprd3("quantity")
			'Response.Write "sales today : " & sales_today  & "<br><br>"
		  end if	
		  prod_today = cdbl(snd_today) - cdbl(snd_yesterday) + cdbl(sales_today)
		  'insert ke WEIGHBRIDGETICKET
		   q=  " INSERT INTO MILL_PRODUCTION MP VALUES (to_Date('" & Rsprd("readingdate") & "','MM/DD/YYYY'),'" & Rsprd("PSTORAGE_STORAGECODE") & "','" & Rsprd("prodcode") & "'," &  cdbl(snd_today)& "," & cdbl(snd_yesterday) & "," & cdbl(sales_today) & "," & cdbl(prod_today) & ", "  &_
			   "  '','','" & Session("CodeName") & "','" & Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "',to_date(substr('" & Rsprd("LASTUPDATE")& "',1,9),'MM/DD/YYYY') ,'"  & Rsprd("LASTTIME") & "') " 
'response.write  q & "<br/>"	
		  connStr.execute(q)	
		  						  
			  
		 
	  Rsprd.MoveNext
	 Wend
 end if
 Rsprd.Close
 set Rsprd = Nothing
						
					    TimeAkhir=Time()
						Lama =TimeAkhir-TimeAwal
						Response.Write "<br>Finish ... " & time() & "<br><br>"
						Response.Write " Time Elapsed " & Hour(Lama) & " Hours " & Minute(Lama) & " Minutes " &  Second(Lama) & " Seconds."
					%>
</body>
</html>