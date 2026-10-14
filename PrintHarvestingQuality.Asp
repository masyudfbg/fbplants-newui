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

v_txtPeriodDate=REQUEST("PeriodDate")

txtSQL = "select H.EMPCODE,H.FDDATE,H.FIELDCODE,SUM(H.UNDER_RIPE) AS UNDER_RIPE ,nvl(SUM(H.OVER_RIPE),0) AS OVER_RIPE,nvl(SUM(H.ABNORMAL),0) AS ABNORMAL," & _
         " nvl(SUM(H.LONG_STALK),0) AS LONG_STALK,nvl(SUM(H.EATEN_BY_RAT),0) AS EATEN_BY_RAT,nvl(SUM(H.UNHARVEST_FFB),0) as  UNHARVEST_FFB, " & _
		 " nvl(SUM(H.UNCOLLECT_LF_CIRCLE),0) AS UNCOLLECT_LF_CIRCLE,nvl(SUM(H.UNCOLLECT_LF_PIECE),0) AS UNCOLLECT_LF_PIECE," & _
	     " nvl(SUM(H.UNARRANGE_FFB),0) AS UNARRANGE_FFB,nvl(SUM(H.UNPRUNE_FROND),0) AS UNPRUNE_FROND,nvl(SUM(H.QE_1),0) AS QE_1," & _
	     " nvl(SUM(H.QE_2),0) AS QE_2,nvl(SUM(H.QE_3),0) AS QE_3,nvl(SUM(H.QE_4),0) AS QE_4,nvl(SUM(H.QE_5),0) AS QE_5," & _
		 " nvl(SUM(H.QE_6),0) AS QE_6,nvl(SUM(H.QE_7),0) AS QE_7,nvl(SUM(H.QE_8),0) AS QE_8,nvl(SUM(H.QE_9),0) AS QE_9," & _
		 " nvl(SUM(H.QE_10),0) AS QE_10" & _
	     " from Harvestingquality H where fddate =to_Date('" & v_txtPeriodDate & "','MM/DD/YYYY') and fcBA='" & Session("fcBA") & "'"

txtsql=(txtsql) & " group by H.EMPCODE,H.FDDATE,H.FIELDCODE"

'response.Write(txtsql)
'response.Flush()
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
	<tr><td width="100">Report</td><td  width="300"><b><FONT color="#000099">HARVESTING QUALITY</FONT></td></tr>
 </table>

<BR><BR>
<table Cellpadding="2" bgcolor="#F7F7F7" cellspacing="0">
		<tr>
		<%	
		Set rsCat = connStr.Execute(txtSQL)
		if Not rsCat.eof then
		   nCountofField=rsCat.fields.count-1
		end if
		NO=1%>
		</tr>
		<tr>
	    <td style="<%=cmdstyleLeft%>;border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"><font face="Verdana" size="1">NO</td>
			
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">DATE</td>
		<td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">EMP. CODE</td>
		<td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">FIELD CODE</td>	
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">UNDER RIPE</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">OVER RIPE</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">ABNORMAL</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">LONG STALK</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">EATEN BY RAT</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">UNHARVEST FFB</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">UNCOLLECT LC</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">UNCOLLECT LP</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">UNARRANGE</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">UNPRUNE</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">QE 1</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">QE 2</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">QE 3</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">QE 4</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">QE 5</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">QE 6</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">QE 7</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">QE 8</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">QE 9</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">QE 10</td>
	</tr>
    <%
    Set rsCat = Server.CreateObject("ADODB.Recordset")
   rsCat.CursorType = 3
   rsCat.CursorLocation = 3
   rsCat.Open txtSql, connstr
   if Not rsCat.eof then
	nCountofField=rsCat.fields.count-1
	no=1
	nQty=0
	nAmount=0
	Dim v_UNDER_RIPE,v_OVER_RIPE,v_ABNORMAL,v_LONG_STALK,v_EATEN_BY_RAT,v_UNHARVEST_FFB,v_UNCOLLECT_LF_CIRCLE,v_UNCOLLECT_LF_PIECE
	Dim v_UNARRANGE_FFB,v_UNPRUNE_FROND,v_QE_1,v_QE_2,v_QE_3,v_QE_4,v_QE_5,v_QE_6,v_QE_7,v_QE_8,v_QE_9,v_QE_10
	
	v_UNDER_RIPE=0
	v_OVER_RIPE=0
	v_ABNORMAL=0
	v_LONG_STALK=0
	v_EATEN_BY_RAT=0
	v_UNHARVEST_FFB=0
	v_UNCOLLECT_LF_CIRCLE=0
	v_UNCOLLECT_LF_PIECE=0
	v_UNARRANGE_FFB=0
	v_UNPRUNE_FROND=0
	v_QE_1=0
	v_QE_2=0
	v_QE_3=0
	v_QE_4=0
	v_QE_5=0
	v_QE_6=0
	v_QE_7=0
	v_QE_8=0
	v_QE_9=0
	v_QE_10=0
	While Not rsCat.EOF%>
	  		<tr>
			<td align="left" style="<%=cmdstyleLeft%>"><font face="Verdana" size="1"><%=NO%></td>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("FDDATE")%></td>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("EMPCODE")%>&nbsp;</td>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("FIELDCODE")%>&nbsp;</td>
			<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("UNDER_RIPE"),2)%></td>
			<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(rscat("OVER_RIPE")),2)%></td>
			<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("ABNORMAL"),2)%></td>
			<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("LONG_STALK"),2)%></td>
			<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("EATEN_BY_RAT"),2)%></td>
			<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("UNHARVEST_FFB"),2)%></td>
			<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("UNCOLLECT_LF_CIRCLE"),2)%></td>
			<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("UNCOLLECT_LF_PIECE"),2)%></td>
			<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("UNARRANGE_FFB"),2)%></td>
			<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("UNPRUNE_FROND"),2)%></td>
			<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("QE_1"),2)%></td>
			<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("QE_2"),2)%></td>
			<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("QE_3"),2)%></td>
			<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("QE_4"),2)%></td>
			<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("QE_5"),2)%></td>
			<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("QE_6"),2)%></td>
			<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("QE_7"),2)%></td>
			<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("QE_8"),2)%></td>
			<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("QE_9"),2)%></td>
			<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("QE_10"),2)%></td>
			</tr>
		  <%
		  rsCat.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend%>
   <%
   else
      response.Write("Data is empty")	  
   end if 
   rsCat.Close
   set rsCat = Nothing%>
   </TABLE>
<!--#include file="Bottom_Report.asp"-->