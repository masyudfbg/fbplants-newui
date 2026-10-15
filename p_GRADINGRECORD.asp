<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<script language="JavaScript" src='/js/x.js'></script>
<script language="JavaScript" src='/js/xtableheaderfixed.js'></script>
<script language="JavaScript">
xAddEventListener(window, 'load',
  function() {
    new xTableHeaderFixed('xthf-site', window);
    new xTableHeaderFixed('xthf-green', window);
    new xTableHeaderFixed('xthf-blue', window);
  }, false
);
</script>
<title>ePlaS Reporting - Grading Record</title>
</head>

<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

v_txtDATETIME_From= trim(request("DATETIME_From"))
v_txtDATETIME_To= trim(request("DATETIME_To"))
v_txtGRADINGNO= trim(request("GRADINGNO"))
v_txtWBT_CHITNO= trim(request("WBT_CHITNO"))
v_txtVEHICLENO=trim(request("VEHICLENO"))

txtSQL = 	"SELECT " & _
			" DATETIME," & _ 
			" VEHICLENO," & _ 
			" nvl(WEIGHT,0) as WEIGHT," & _ 
			" nvl(BUNCHES,0) as BUNCHES," & _
			" nvl(TKOSONG,0) TKOSONG," & _
			" nvl(TPANJANG,0) TPANJANG," & _
			" nvl(DURA,0) DURA," & _
			" nvl(BKECIL,0) BKECIL," & _
			" nvl(AIR,0) AIR," & _
			" nvl(SAMPAH,0) SAMPAH," & _
			" nvl(PASIR,0) PASIR," & _
			" nvl(LAIN,0) LAIN," & _
			" nvl(eatenbyrat,0) eatenbyrat," & _
			" nvl(brondolan,0) brondolan," & _
			" nvl(mentah,0) mentah," & _
			" nvl(kmt,0) kmt," & _
			" nvl(mt,0) mt," & _
			" nvl(tmt,0) tmt," & _
			" nvl(bsk,0) bsk," & _
			" nvl(an,0) an," & _
			" REMARKS," & _
			" WBT_CHITNO FROM GRADINGRECORD" & _
			" where fcba='" & session("fcBA") & "'"
'response.write(txtSQL)
if  len(trim(v_txtDATETIME_From)) > 0 then
    txtSQL = txtSQL & " and DATETIME >=to_Date('" & v_txtDATETIME_From & "','MM/DD/YYYY')"
	txtSelect = " From " & FormatDateTime(v_txtDATETIME_From,1)
end if
if  len(trim(v_txtDATETIME_To)) > 0 then
    txtSQL =txtSQL &  " and DATETIME <=to_Date('" & v_txtDATETIME_To & "','MM/DD/YYYY')"
	txtSelect = txtSelect & "  to " & FormatDateTime(v_txtDATETIME_To,1)
end if
if  len(trim(v_txtGRADINGNO)) > 0 then
    txtSQL = txtSQL & " and GRADINGNO='" & v_txtGRADINGNO & "'"
	txtSelect = txtSelect & " Grading No " & v_txtGRADINGNO
end if
if  len(trim(v_txtWBT_CHITNO)) > 0 then
    txtSQL =txtSQL &  " and WBT_CHITNO='" & v_txtWBT_CHITNO & "'"
	txtSelect = txtSelect & " Chit No " & v_txtWBT_CHITNO
end if 
if  len(trim(v_txtVEHICLENO)) > 0 then
    txtSQL =txtSQL &  " and VEHICLENO='" & v_txtVEHICLENO & "'"
	txtSelect = txtSelect & " VEHICLENO " & v_txtVEHICLENO
end if 

'txtsql=(txtsql) & " group by GRADINGNO, DATETIME, VEHICLENO, REMARKS, CHECKEDBY, WBT_CHITNO " 
'                     "VEHICLE, LOADERCODE, RAMPATTENDANTCODE,WEIGHT"
if len(trim(txtSelect)) > 0 then
   txtSelect = "Selected  by : " & txtSelect
else
   txtSelect = "Selected by All Data"
end if
'response.Write(txtsql)
'response.Flush()
%>
<body>
		<div id="wrapper">
        <!--#include file="print_excel.asp"-->
		 <table border="0" width="100%" id="product-table" class="xthf-site">
         <thead>
			 <tr>
				<th colspan="16" style=" background-color:#FFFFFF; color:#000000; border:none">
					<div id="header-wrapper">
						<div id="head_judul" style="float:left; text-align:left;">
							<h4>PT. <%=Session("COMPANY")%></h4>
							<h5><%=Session("BA")%></h5>
						</div>
						<div id="subheader" class="centered" style="width:50%; margin:auto;">
							<h2>GRADING RECORD</h2>
						</div>
					</div>
				</th>
			</tr>
							<!--<div class="printhide">-->
			<tr>
				<td colspan="20" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></td>
			</tr>
				<!--</div>-->
				<%	
				Set rsCat = connStr.Execute(txtSQL)
				if Not rsCat.eof then
					nCountofField=rsCat.fields.count-1
				end if
				NO=1%>
				<tr>
					<th style="width:70px;">Date</th>
					<th>Chit No</th>
					<th style="width:70px;">VEHICLE No</th>
					
					<th style="width:10px;">DURA</th>
					<th style="width:10px;">B.KECIL</th>
					<th style="width:10px;">AIR</th>
					<th style="width:10px;">SAMPAH</th>
					<th style="width:10px;">PASIR</th>
					<th style="width:10px;">LAIN</th>
					<th style="width:10px;">B.DIMAKAN TIKUS</th>
					<th style="width:10px;">P.BRONDOLAN</th>
					<th style="width:10px;">MENTAH</th>
					<th style="width:10px;">KMT</th>
					<th style="width:10px;">MT</th>
					<th style="width:10px;">TMT</th>
					<th style="width:10px;">BSK</th>
					<th style="width:10px;">AN</th>
					<th style="width:10px;">JJK</th>
					<th style="width:10px;">TP</th>
					<th style="width:10px;">BUNCHES</th>
					<th style="width:10px;">WEIGHT</th>
					<th>REMARKS</th>
				</tr>
                </thead>
               <tbody>
	<%
	Dim v_WEIGHT,v_BUNCHES
	v_WEIGHT=0
	v_BUNCHES=0
	While Not rsCat.EOF%>
		<tr>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("DATETIME")%>&nbsp;</td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("WBT_CHITNO")%>&nbsp;</td>
		<td align="CENTER" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("VEHICLENO")%>&nbsp;</td>
		
		<td align="CENTER" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("DURA")%>&nbsp;</td>
		<td align="CENTER" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("BKECIL")%>&nbsp;</td>
		<td align="CENTER" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("AIR")%>&nbsp;</td>
		<td align="CENTER" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("SAMPAH")%>&nbsp;</td>
		<td align="CENTER" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("PASIR")%>&nbsp;</td>
		<td align="CENTER" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("LAIN")%>&nbsp;</td>
		<td align="CENTER" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("EATENBYRAT")%>&nbsp;</td>
		<td align="CENTER" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("BRONDOLAN")%>&nbsp;</td>
		<td align="CENTER" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("MENTAH")%>&nbsp;</td>
		<td align="CENTER" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("KMT")%>&nbsp;</td>
		<td align="CENTER" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("MT")%>&nbsp;</td>
		<td align="CENTER" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("TMT")%>&nbsp;</td>
		<td align="CENTER" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("BSK")%>&nbsp;</td>
		<td align="CENTER" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("AN")%>&nbsp;</td>
		<td align="CENTER" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("TKOSONG")%>&nbsp;</td>
		<td align="CENTER" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("TPANJANG")%>&nbsp;</td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("BUNCHES"),2)%></td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("WEIGHT"),2)%></td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("REMARKS")%>&nbsp;</td>
		</tr>
	  <%
	  v_WEIGHT=v_WEIGHT+trim(cdbl(rsCat("WEIGHT")))
	  v_BUNCHES=v_BUNCHES+trim(cdbl(rsCat("BUNCHES")))
	  'response.write cdbl(trim(rsCat("MTHOPENBALANCE")))+1
	  rsCat.MoveNext
	  'response.Flush()
	  no=no+1
	Wend
	rsCat.close()
	rsCat=null%>
	<input  type="hidden" value="<%=txtTable%>" name="txtTable">
	<input  type="hidden" value="<%=txtModule%>" name="txtModule">
	</form>
	<tr bgcolor="#F4F4A0">
	<td align="Center" colspan="19"> <strong><i> <font face="Verdana" size="2">T o t a l</i></strong></td>
	<td align="Right"> <strong><i> <font face="Verdana" size="1"><%=FormatNumber(v_WEIGHT,2)%></i></strong></td>
	<td align="Right"> <strong><i> <font face="Verdana" size="1"><%=FormatNumber(v_BUNCHES,2)%></i></strong></td>
	<td align="Right"> <strong><i> <font face="Verdana" size="1">&nbsp;</i></strong></td>
	</tr>
	</tbody>
</table>
<!--#include file="Bottom_Report.asp"-->