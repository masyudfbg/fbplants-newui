<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"
%>
<html>
<head>
<title>LHK Non Produksi</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js"></script>
<script language="JavaScript" src="date-picker.js"></script> 
</head>
<body>
 <%
    Dim txtNoWS,txtDateWS,txtRemarks,q,nCountofField,no,i
    Dim Remarks
    Dim rsWS,intRecordsPerPage,cColor,txtVehicle
	txtVehicle =cmdFindTitle("Vehicle","fcCode",request("VehicleCode"))
%>
	 <table width="700">
	  <tr> 
		<td rowspan="6" width="100">
			<% if len(trim(PROFIL("REPORT_IMAGE"))) > 0 then %>
				<img src="IMAGES/<%=PROFIL("REPORT_IMAGE")%>" width="210" height="80">
			<% end if %>
		</td>
	  </tr>
		<tr><td width="100">Company</td><td  width="300"><b><FONT color="#000099"><%=Session("COMPANY")%></FONT></td></tr>
		<tr><td width="100">Business Unit</td><td  width="300"><b><FONT color="#000099"><%=Session("BA")%> / <%=Session("FCBA")%></FONT></td></tr>
		<tr><td width="100">Manager</td><td  width="300"><b><FONT color="#000099"><%=Session("FCMANAGER")%></FONT></td></tr>
		<tr><td width="100">Report</td><td  width="300"><b><FONT color="#000099">HARVESTING</FONT></td></tr>
		<tr><td width="100">For Date</td><td  width="300"><b><FONT color="#000099"><%=FormatDateTime(REQUEST("TRANS_DATE"),1)%></FONT></td></tr>
	 </table>
	<BR><BR>


<%
	   Dim rsVehicle
	   q	=	" SELECT " & _
				" STACKING_HA, TANAMBARU_PKK, TANAMBARU_HA, KOMPLETING_PKK, KOMPLETING_HA, TANAMSISIP_TBM_PKK, " & _
				" TANAMSISIP_TM_PKK, KASTRASI_HA, SANITASI_HA, " & _
				" PEMEL_JALAN_TIMBUN_MR_MTR, PEMEL_JALAN_TIMBUN_CR_MTR, PEMEL_JALAN_PENGERASAN_MR_MTR, PEMEL_JALAN_PENGERASAN_CR_MTR," & _
				" TO_CHAR (TRANS_DATE, 'mm/yyyy') as TRANS_DATE, ID, FCBA" & _
				" FROM LHK_NON_PROD_BUDGET " & _
				" WHERE FCBA='" & session("FCBA") & "'"
		
			if  len(trim(request("TRANS_DATE"))) > 0 then
				q = q & " AND to_char(TRANS_DATE,'mm/yyyy')=to_char(to_date('" & request("TRANS_DATE") & "','" & "MM/DD/YYYY" & "'),'mm/yyyy')"
			end if
			
			' if  len(trim(request("Txt_Fieldcode"))) > 0 then
				' q = q & " and FIELDCODE='" & request("Txt_Fieldcode") & "'"
			' end if
			
			' q = q & " group by CHITNO, FIELDCODE, CROPCODE, PRODUCTCODE, to_char(RECEPTIONDATE,'mm/dd/yyyy'), VEHICLE, OWN"
  
	   'response.Write (q)
	   'response.Flush()	
	   Set rsVehicle = Server.CreateObject("ADODB.Recordset")
	   rsVehicle.CursorType = 3
	   rsVehicle.CursorLocation = 3
	   rsVehicle.Open q, connstr
   
%>
   <br><br>
	<table Cellpadding="2" bgcolor="#F7F7F7" cellspacing="0">
   <tr>   
	   <td style="<%=cmdstyleLeft%>;border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"><font face="Verdana" size="1">NO</td>	   
	   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">Stacking (Ha)</td>
	   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">Tanam Baru (Pkk)</td>
	   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">Tanam Baru (Ha)</td>
	   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">Kompleting (Pkk)</td>
	   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">Kompleting (Ha)</td>
	   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">Tanam Sisip TBM (Pkk)</td>
	   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">Tanam Sisip TM (Pkk)</td>
	   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">Kastrasi (Ha)</td>
	   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">Sanitasi (Ha)</td>
	   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">Timbun MR (Mtr)</td>
	   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">Timbun CR (Mtr)</td>
	   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">Pengerasan MR (Mtr)</td>
	   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">Pengerasan CR (Mtr)</td>
	   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">Tanggal Transaksi</td>

   </tr>
   <%
		'Dim nMillweight,nEstate,nQuantity
		if Not rsVehicle.eof then
			nCountofField=rsVehicle.fields.count-1
			no=1
			' nMillweight=0
			' nEstate=0
			' nQuantity=0
		  While Not rsVehicle.EOF
			  ' nMillweight = nMillweight+cdbl(rsVehicle("MILL_WEIGHT"))
			  ' nEstate = nEstate+cdbl(rsVehicle("ESTATE_WEIGHT"))
			  ' nQuantity = nQuantity+cdbl(rsVehicle("QUANTITY"))
			  
			  if no  mod 2 = 1 then
				 cColor = "#EEEEEE"
			  else
				 cColor = "#E5E5E5"
			  end if
	%>
			<tr bgcolor="<%=cColor%>">
				<td align="left" style="<%=cmdstyleLeft%>"><font face="Verdana" size="1"><%=NO%>&nbsp;</td>
				<td align="left"  title="STACKING_HA" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsVehicle("STACKING_HA")%>&nbsp;</td>
				<td align="left"  title="TANAMBARU_PKK" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsVehicle("TANAMBARU_PKK")%>&nbsp;</td>
				<td align="left"  title="TANAMBARU_HA" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsVehicle("TANAMBARU_HA")%>&nbsp;</td>
				<td align="left"  title="KOMPLETING_PKK" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsVehicle("KOMPLETING_PKK")%>&nbsp;</td>
				<td align="left"  title="KOMPLETING_HA" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsVehicle("KOMPLETING_HA")%>&nbsp;</td>
				<td align="left"  title="TANAMSISIP_TBM_PKK" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsVehicle("TANAMSISIP_TBM_PKK")%>&nbsp;</td>
				<td align="left"  title="TANAMSISIP_TM_PKK" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsVehicle("TANAMSISIP_TM_PKK")%>&nbsp;</td>		  
				<td align="right" title="KASTRASI_HA" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsVehicle("KASTRASI_HA")%>&nbsp;</td>
				<td align="right" title="SANITASI_HA" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsVehicle("SANITASI_HA")%>&nbsp;</td>
				<td align="left"  title="PEMEL_JALAN_TIMBUN_MR_MTR" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsVehicle("PEMEL_JALAN_TIMBUN_MR_MTR")%>&nbsp;</td>
				<td align="left"  title="PEMEL_JALAN_TIMBUN_CR_MTR" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsVehicle("PEMEL_JALAN_TIMBUN_CR_MTR")%>&nbsp;</td>
				<td align="left"  title="PEMEL_JALAN_PENGERASAN_MR_MTR" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsVehicle("PEMEL_JALAN_PENGERASAN_MR_MTR")%>&nbsp;</td>
				<td align="left"  title="PEMEL_JALAN_PENGERASAN_CR_MTR" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsVehicle("PEMEL_JALAN_PENGERASAN_CR_MTR")%>&nbsp;</td>
				<td align="left"  title="TRANS_DATE" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsVehicle("TRANS_DATE")%>&nbsp;</td>
			</tr>
	<%
			  rsVehicle.MoveNext
			  intRecordsPerPage = intRecordsPerPage - 1
			  no=no+1
		  Wend
	%>
		 <!-- 
		  <tr>
		  <td colspan="8" align="center" style="<%=cmdstyleLeft%>"><font face="Verdana" size="1">T o t a l</td>
		  <td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(nQuantity,2)%>&nbsp;</td>
		  <td align="right"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(nEstate,2)%>&nbsp;</td>
		  <td align="right"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(nMillweight,2)%>&nbsp;</td>
		  </tr>
		  -->
	<%
	   else
		  response.Write("Data is empty")	  
	   end if 
	   rsVehicle.Close
	   set rsVehicle = Nothing
   %>
   </TABLE>
<!--#include file="Bottom_Report.asp"-->