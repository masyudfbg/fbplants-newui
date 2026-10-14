<!--#include file="../connDb1.Inc"-->
<!--#include file="../txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>iPS Report - Employee Absence</title>
<style type="text/css">
#hor-zebra
{
	font-family: "Lucida Sans Unicode", "Lucida Grande", Sans-Serif;
	font-size: 12px;
	margin: 45px;
	width: 480px;
	text-align: left;
	border-collapse: collapse;
}
#hor-zebra th
{
	font-size: 14px;
	font-weight: normal;
	padding: 10px 8px;
	color: #039;
}
#hor-zebra td
{
	padding: 8px;
	color: #669;
}
#hor-zebra .odd
{
	background: #e8edff; 
}


</style>
<%
Dim rsAbsence, strAbsence

	v_txtyear			=	trim(request("txtyear"))
	v_txtmonth			=	trim(request("txtmonth"))
	v_txtEmployeeCode	=	trim(request("txtEmployeeCode"))
	
strAbsence = " select " & _
"     employeecode, " & _
"     fcname," & _
"     (select fccode from parameter where parcode = 'E0102' and fcba = '" & Session("FCBA") & "' and rate_1 = day_1) day_1," & _
"     (select fccode from parameter where parcode = 'E0102' and fcba = '" & Session("FCBA") & "' and rate_1 = day_2) day_2," & _
"     (select fccode from parameter where parcode = 'E0102' and fcba = '" & Session("FCBA") & "' and rate_1 = day_3) day_3," & _
"     (select fccode from parameter where parcode = 'E0102' and fcba = '" & Session("FCBA") & "' and rate_1 = day_4) day_4," & _
"     (select fccode from parameter where parcode = 'E0102' and fcba = '" & Session("FCBA") & "' and rate_1 = day_5) day_5," & _
"     (select fccode from parameter where parcode = 'E0102' and fcba = '" & Session("FCBA") & "' and rate_1 = day_6) day_6," & _
"     (select fccode from parameter where parcode = 'E0102' and fcba = '" & Session("FCBA") & "' and rate_1 = day_7) day_7," & _
"     (select fccode from parameter where parcode = 'E0102' and fcba = '" & Session("FCBA") & "' and rate_1 = day_8) day_8," & _
"     (select fccode from parameter where parcode = 'E0102' and fcba = '" & Session("FCBA") & "' and rate_1 = day_9) day_9," & _
"     (select fccode from parameter where parcode = 'E0102' and fcba = '" & Session("FCBA") & "' and rate_1 = day_10) day_10," & _
"     (select fccode from parameter where parcode = 'E0102' and fcba = '" & Session("FCBA") & "' and rate_1 = day_11) day_11," & _
"     (select fccode from parameter where parcode = 'E0102' and fcba = '" & Session("FCBA") & "' and rate_1 = day_12) day_12," & _
"     (select fccode from parameter where parcode = 'E0102' and fcba = '" & Session("FCBA") & "' and rate_1 = day_13) day_13," & _
"     (select fccode from parameter where parcode = 'E0102' and fcba = '" & Session("FCBA") & "' and rate_1 = day_14) day_14," & _
"     (select fccode from parameter where parcode = 'E0102' and fcba = '" & Session("FCBA") & "' and rate_1 = day_15) day_15," & _
"     (select fccode from parameter where parcode = 'E0102' and fcba = '" & Session("FCBA") & "' and rate_1 = day_16) day_16," & _
"     (select fccode from parameter where parcode = 'E0102' and fcba = '" & Session("FCBA") & "' and rate_1 = day_17) day_17," & _
"     (select fccode from parameter where parcode = 'E0102' and fcba = '" & Session("FCBA") & "' and rate_1 = day_18) day_18," & _
"     (select fccode from parameter where parcode = 'E0102' and fcba = '" & Session("FCBA") & "' and rate_1 = day_19) day_19," & _
"     (select fccode from parameter where parcode = 'E0102' and fcba = '" & Session("FCBA") & "' and rate_1 = day_20) day_20," & _
"     (select fccode from parameter where parcode = 'E0102' and fcba = '" & Session("FCBA") & "' and rate_1 = day_21) day_21," & _
"     (select fccode from parameter where parcode = 'E0102' and fcba = '" & Session("FCBA") & "' and rate_1 = day_22) day_22," & _
"     (select fccode from parameter where parcode = 'E0102' and fcba = '" & Session("FCBA") & "' and rate_1 = day_23) day_23," & _
"     (select fccode from parameter where parcode = 'E0102' and fcba = '" & Session("FCBA") & "' and rate_1 = day_24) day_24," & _
"     (select fccode from parameter where parcode = 'E0102' and fcba = '" & Session("FCBA") & "' and rate_1 = day_25) day_25," & _
"     (select fccode from parameter where parcode = 'E0102' and fcba = '" & Session("FCBA") & "' and rate_1 = day_26) day_26," & _
"     (select fccode from parameter where parcode = 'E0102' and fcba = '" & Session("FCBA") & "' and rate_1 = day_27) day_27," & _
"     (select fccode from parameter where parcode = 'E0102' and fcba = '" & Session("FCBA") & "' and rate_1 = day_28) day_28," & _
"     (select fccode from parameter where parcode = 'E0102' and fcba = '" & Session("FCBA") & "' and rate_1 = day_29) day_29," & _
"     (select fccode from parameter where parcode = 'E0102' and fcba = '" & Session("FCBA") & "' and rate_1 = day_30) day_30," & _
"     (select fccode from parameter where parcode = 'E0102' and fcba = '" & Session("FCBA") & "' and rate_1 = day_31) day_31 " & _
" from (" & _
" select " & _
"     employeecode," & _
"     fcname," & _
"     sum(day_1) day_1," & _
"     sum(day_2) day_2," & _
"     sum(day_3) day_3," & _
"     sum(day_4) day_4," & _
"     sum(day_5) day_5," & _
"     sum(day_6) day_6," & _
"     sum(day_7) day_7," & _
"     sum(day_8) day_8," & _
"     sum(day_9) day_9," & _
"     sum(day_10) day_10," & _
"     sum(day_11) day_11," & _
"     sum(day_12) day_12," & _
"     sum(day_13) day_13," & _
"     sum(day_14) day_14," & _
"     sum(day_15) day_15," & _
"     sum(day_16) day_16," & _
"     sum(day_17) day_17," & _
"     sum(day_18) day_18," & _
"     sum(day_19) day_19," & _
"     sum(day_20) day_20," & _
"     sum(day_21) day_21," & _
"     sum(day_22) day_22," & _
"     sum(day_23) day_23," & _
"     sum(day_24) day_24," & _
"     sum(day_25) day_25," & _
"     sum(day_26) day_26," & _
"     sum(day_27) day_27," & _
"     sum(day_28) day_28," & _
"     sum(day_29) day_29," & _
"     sum(day_30) day_30," & _
"     sum(day_31) day_31" & _
" from (" & _
" SELECT " & _
"     employeecode," & _
"     fcname," & _
"     DECODE (TO_NUMBER (TO_CHAR (fddate, 'DD')),1, (SELECT rate_1 FROM parameter WHERE parcode = 'E0102' AND fcba = '" & Session("FCBA") & "' AND fccode = attendance), 0) day_1," & _
"     DECODE (TO_NUMBER (TO_CHAR (fddate, 'DD')),2, (SELECT rate_1 FROM parameter WHERE parcode = 'E0102' AND fcba = '" & Session("FCBA") & "' AND fccode = attendance), 0) day_2," & _
"     DECODE (TO_NUMBER (TO_CHAR (fddate, 'DD')),3, (SELECT rate_1 FROM parameter WHERE parcode = 'E0102' AND fcba = '" & Session("FCBA") & "' AND fccode = attendance), 0) day_3," & _
"     DECODE (TO_NUMBER (TO_CHAR (fddate, 'DD')),4, (SELECT rate_1 FROM parameter WHERE parcode = 'E0102' AND fcba = '" & Session("FCBA") & "' AND fccode = attendance), 0) day_4," & _
"     DECODE (TO_NUMBER (TO_CHAR (fddate, 'DD')),5, (SELECT rate_1 FROM parameter WHERE parcode = 'E0102' AND fcba = '" & Session("FCBA") & "' AND fccode = attendance), 0) day_5," & _
"     DECODE (TO_NUMBER (TO_CHAR (fddate, 'DD')),6, (SELECT rate_1 FROM parameter WHERE parcode = 'E0102' AND fcba = '" & Session("FCBA") & "' AND fccode = attendance), 0) day_6," & _
"     DECODE (TO_NUMBER (TO_CHAR (fddate, 'DD')),7, (SELECT rate_1 FROM parameter WHERE parcode = 'E0102' AND fcba = '" & Session("FCBA") & "' AND fccode = attendance), 0) day_7," & _
"     DECODE (TO_NUMBER (TO_CHAR (fddate, 'DD')),8, (SELECT rate_1 FROM parameter WHERE parcode = 'E0102' AND fcba = '" & Session("FCBA") & "' AND fccode = attendance), 0) day_8," & _
"     DECODE (TO_NUMBER (TO_CHAR (fddate, 'DD')),9, (SELECT rate_1 FROM parameter WHERE parcode = 'E0102' AND fcba = '" & Session("FCBA") & "' AND fccode = attendance), 0) day_9," & _
"     DECODE (TO_NUMBER (TO_CHAR (fddate, 'DD')),10, (SELECT rate_1 FROM parameter WHERE parcode = 'E0102' AND fcba = '" & Session("FCBA") & "' AND fccode = attendance), 0) day_10," & _
"     DECODE (TO_NUMBER (TO_CHAR (fddate, 'DD')),11, (SELECT rate_1 FROM parameter WHERE parcode = 'E0102' AND fcba = '" & Session("FCBA") & "' AND fccode = attendance), 0) day_11," & _
"     DECODE (TO_NUMBER (TO_CHAR (fddate, 'DD')),12, (SELECT rate_1 FROM parameter WHERE parcode = 'E0102' AND fcba = '" & Session("FCBA") & "' AND fccode = attendance), 0) day_12," & _
"     DECODE (TO_NUMBER (TO_CHAR (fddate, 'DD')),13, (SELECT rate_1 FROM parameter WHERE parcode = 'E0102' AND fcba = '" & Session("FCBA") & "' AND fccode = attendance), 0) day_13," & _
"     DECODE (TO_NUMBER (TO_CHAR (fddate, 'DD')),14, (SELECT rate_1 FROM parameter WHERE parcode = 'E0102' AND fcba = '" & Session("FCBA") & "' AND fccode = attendance), 0) day_14," & _
"     DECODE (TO_NUMBER (TO_CHAR (fddate, 'DD')),15, (SELECT rate_1 FROM parameter WHERE parcode = 'E0102' AND fcba = '" & Session("FCBA") & "' AND fccode = attendance), 0) day_15," & _
"     DECODE (TO_NUMBER (TO_CHAR (fddate, 'DD')),16, (SELECT rate_1 FROM parameter WHERE parcode = 'E0102' AND fcba = '" & Session("FCBA") & "' AND fccode = attendance), 0) day_16," & _
"     DECODE (TO_NUMBER (TO_CHAR (fddate, 'DD')),17, (SELECT rate_1 FROM parameter WHERE parcode = 'E0102' AND fcba = '" & Session("FCBA") & "' AND fccode = attendance), 0) day_17," & _
"     DECODE (TO_NUMBER (TO_CHAR (fddate, 'DD')),18, (SELECT rate_1 FROM parameter WHERE parcode = 'E0102' AND fcba = '" & Session("FCBA") & "' AND fccode = attendance), 0) day_18," & _
"     DECODE (TO_NUMBER (TO_CHAR (fddate, 'DD')),19, (SELECT rate_1 FROM parameter WHERE parcode = 'E0102' AND fcba = '" & Session("FCBA") & "' AND fccode = attendance), 0) day_19," & _
"     DECODE (TO_NUMBER (TO_CHAR (fddate, 'DD')),20, (SELECT rate_1 FROM parameter WHERE parcode = 'E0102' AND fcba = '" & Session("FCBA") & "' AND fccode = attendance), 0) day_20," & _
"     DECODE (TO_NUMBER (TO_CHAR (fddate, 'DD')),21, (SELECT rate_1 FROM parameter WHERE parcode = 'E0102' AND fcba = '" & Session("FCBA") & "' AND fccode = attendance), 0) day_21," & _
"     DECODE (TO_NUMBER (TO_CHAR (fddate, 'DD')),22, (SELECT rate_1 FROM parameter WHERE parcode = 'E0102' AND fcba = '" & Session("FCBA") & "' AND fccode = attendance), 0) day_22," & _
"     DECODE (TO_NUMBER (TO_CHAR (fddate, 'DD')),23, (SELECT rate_1 FROM parameter WHERE parcode = 'E0102' AND fcba = '" & Session("FCBA") & "' AND fccode = attendance), 0) day_23," & _
"     DECODE (TO_NUMBER (TO_CHAR (fddate, 'DD')),24, (SELECT rate_1 FROM parameter WHERE parcode = 'E0102' AND fcba = '" & Session("FCBA") & "' AND fccode = attendance), 0) day_24," & _
"     DECODE (TO_NUMBER (TO_CHAR (fddate, 'DD')),25, (SELECT rate_1 FROM parameter WHERE parcode = 'E0102' AND fcba = '" & Session("FCBA") & "' AND fccode = attendance), 0) day_25," & _
"     DECODE (TO_NUMBER (TO_CHAR (fddate, 'DD')),26, (SELECT rate_1 FROM parameter WHERE parcode = 'E0102' AND fcba = '" & Session("FCBA") & "' AND fccode = attendance), 0) day_26," & _
"     DECODE (TO_NUMBER (TO_CHAR (fddate, 'DD')),27, (SELECT rate_1 FROM parameter WHERE parcode = 'E0102' AND fcba = '" & Session("FCBA") & "' AND fccode = attendance), 0) day_27," & _
"     DECODE (TO_NUMBER (TO_CHAR (fddate, 'DD')),28, (SELECT rate_1 FROM parameter WHERE parcode = 'E0102' AND fcba = '" & Session("FCBA") & "' AND fccode = attendance), 0) day_28," & _
"     DECODE (TO_NUMBER (TO_CHAR (fddate, 'DD')),29, (SELECT rate_1 FROM parameter WHERE parcode = 'E0102' AND fcba = '" & Session("FCBA") & "' AND fccode = attendance), 0) day_29," & _
"     DECODE (TO_NUMBER (TO_CHAR (fddate, 'DD')),30, (SELECT rate_1 FROM parameter WHERE parcode = 'E0102' AND fcba = '" & Session("FCBA") & "' AND fccode = attendance), 0) day_30," & _
"     DECODE (TO_NUMBER (TO_CHAR (fddate, 'DD')),31, (SELECT rate_1 FROM parameter WHERE parcode = 'E0102' AND fcba = '" & Session("FCBA") & "' AND fccode = attendance), 0) day_31 " & _
" FROM" & _
"     attendance_gad ag, " & _
"     employee em" & _
" where" & _
"     (ag.employeecode = em.fccode and ag.fcba = em.fcba)" & _
" and    " & _
"     ag.fcba = '" & Session("FCBA") & "'" & _
" )" & _
" group by " & _
"     employeecode," & _
"     fcname" & _
" )"

		Set rsAbsence = Server.CreateObject("ADODB.Recordset")
		rsAbsence.CursorType = 3
		rsAbsence.CursorLocation = 3
		rsAbsence.Open strAbsence, connstr
		
%>
<table id="hor-zebra">
	<tr>
		<th>No</th>
		<th>Employee</th>
		<th>1</th>
		<th>2</th>
		<th>3</th>
		<th>4</th>
		<th>5</th>
		<th>6</th>
		<th>7</th>
		<th>8</th>
		<th>9</th>
		<th>10</th>
		<th>11</th>
		<th>12</th>
		<th>13</th>
		<th>14</th>
		<th>15</th>
		<th>16</th>
		<th>17</th>
		<th>18</th>
		<th>19</th>
		<th>20</th>
		<th>21</th>
		<th>22</th>
		<th>23</th>
		<th>24</th>
		<th>25</th>
		<th>26</th>
		<th>27</th>
		<th>28</th>
		<th>29</th>
		<th>30</th>
		<th>31</th>
	</tr>
	<%
		NO = 1
		While Not rsAbsence.EOF
		if NO mod 2 = 1 then
             cClass = "odd"
		  else
		     cClass = "default"
		  end if
	%>
	<tr class = "<%=cClass%>">
		<td><%=NO%></td>
		<td><%=rsAbsence("FCNAME")%></td>
		<td><%=rsAbsence("day_1")%></td>
		<td><%=rsAbsence("day_2")%></td>
		<td><%=rsAbsence("day_3")%></td>
		<td><%=rsAbsence("day_4")%></td>
		<td><%=rsAbsence("day_5")%></td>
		<td><%=rsAbsence("day_6")%></td>
		<td><%=rsAbsence("day_7")%></td>
		<td><%=rsAbsence("day_8")%></td>
		<td><%=rsAbsence("day_9")%></td>
		<td><%=rsAbsence("day_10")%></td>
		<td><%=rsAbsence("day_11")%></td>
		<td><%=rsAbsence("day_12")%></td>
		<td><%=rsAbsence("day_13")%></td>
		<td><%=rsAbsence("day_14")%></td>
		<td><%=rsAbsence("day_15")%></td>
		<td><%=rsAbsence("day_16")%></td>
		<td><%=rsAbsence("day_17")%></td>
		<td><%=rsAbsence("day_18")%></td>
		<td><%=rsAbsence("day_19")%></td>
		<td><%=rsAbsence("day_20")%></td>
		<td><%=rsAbsence("day_21")%></td>
		<td><%=rsAbsence("day_22")%></td>
		<td><%=rsAbsence("day_23")%></td>
		<td><%=rsAbsence("day_24")%></td>
		<td><%=rsAbsence("day_25")%></td>
		<td><%=rsAbsence("day_26")%></td>
		<td><%=rsAbsence("day_27")%></td>
		<td><%=rsAbsence("day_28")%></td>
		<td><%=rsAbsence("day_29")%></td>
		<td><%=rsAbsence("day_30")%></td>
		<td><%=rsAbsence("day_31")%></td>
	</tr>
	<%rsAbsence.MoveNext%>
	<%NO = NO + 1%>
	<%Wend%>
	<%rsAbsence.close()%>
	<%rsAbsence = null%>
</table>