<!--#include file="../connDb1.Inc"-->
<!--#include file="../txtFunction.inc"-->
<%
	'---- CursorTypeEnum Values ----
	Const adOpenForwardOnly = 0
	Const adOpenKeyset = 1
	Const adOpenDynamic = 2
	Const adOpenStatic = 3
	
	v_txtyear			=	Cint(trim(request("txtyear")))
	v_txtmonth			=	Cint(trim(request("txtmonth")))
	v_txtEmployeeCode	=	trim(request("txtEmployeeCode"))

	if v_txtyear <> "" and v_txtmonth <> "" and v_txtEmployeeCode <> "" then
		strDataEmployee = 	" SELECT" & _
							" *" & _
							" FROM VIEW_TOTALATTENDANCE" & _
							" WHERE" & _
							" NYEAR = " & v_txtyear & "" & _
							" AND NMONTH = " & v_txtmonth & "" & _
							" AND EMPLOYEECODE = '" & v_txtEmployeeCode & "'" & _
							" AND FCBA ='" & Session("fcBA") & "'"

		Set rsDataEmployee = Server.CreateObject("ADODB.Recordset")
		rsDataEmployee.CursorType = 3
		rsDataEmployee.CursorLocation = 3
		rsDataEmployee.Open strDataEmployee, connstr
			
		strGaji		=	" select nvl(sum(amount),0) amount from empallded where allowdedcode = 'A01'"
		strLembur	=	" select nvl(sum(amount),0) amount from empallded where allowdedcode = 'A18'"
		strPSB		=	" select nvl(sum(amount),0) amount from empallded where allowdedcode = 'A27'"
		strPLB		=	" select nvl(sum(amount),0) amount from empallded where allowdedcode = 'A24'"
		strPB		=	" select nvl(sum(amount),0) amount from empallded where allowdedcode = 'A45'"
		strPK		=	" select nvl(sum(amount),0) amount from empallded where allowdedcode = 'A23'"
		strPT		=	" select nvl(sum(amount),0) amount from empallded where allowdedcode = 'A32'"
		strPL		=	" select nvl(sum(amount),0) amount from empallded where allowdedcode not in" & _
						" ('A01','A18','A27','A24','A45','A23','A32','A17')" & _
						" and allowdedcode like 'A%'"
		strTJ		=	" select nvl(sum(amount),0) amount from empallded where allowdedcode = 'A17'"
		strsubPR	=	" select nvl(sum(amount),0) amount from empallded where allowdedcode not in ('A01') and allowdedcode like 'A%'"
		
		strGajiBruto	=	" select nvl(sum(amount),0) amount from empallded where allowdedcode like 'A%'"
		
		strPinjam		=	" select nvl(sum(amount),0) amount from empallded where allowdedcode in ('D04','D06','D07')"
		strPAK			=	" select nvl(sum(amount),0) amount from empallded where allowdedcode = 'D05'"
		strMP1			=	" select nvl(sum(amount),0) amount from empallded where allowdedcode in ('D01','D18')"
		strTDB			=	" select nvl(sum(amount),0) amount from empallded where allowdedcode = 'D19'"
		strDPN			=	" select nvl(sum(amount),0) amount from empallded where allowdedcode in ('D09','D10','D12','D13','D14','D15','D16')"
		strPLL			=	" select nvl(sum(amount),0) amount from empallded where allowdedcode not in" & _
							" ('D04','D06','D07','D05','D02','D18','D19','D09','D10','D12','D13','D14','D15','D16','D20')" & _
							" and allowdedcode like 'D%'"
		strPPH21		=	" select nvl(sum(amount),0) amount from empallded where allowdedcode = 'D20'"
		strPJ			=	" select nvl(sum(amount),0) amount from empallded where allowdedcode = 'D02'"
		strsubPotong	=	" select nvl(sum(amount),0) amount from empallded where allowdedcode like 'D%'"

		strGajiNetto	=	" select nvl(sum(amount),0) amount from empallded where"
		
		strWhere0	=	" fcba = '" & Session("fcBA") & "'"
		
		if v_txtyear <> "" then
			strWhere1	=	" and to_number(to_char(adhocdate,'yyyy')) = " & v_txtyear
		end if
		if v_txtmonth <> "" then
			strWhere2	=	" and to_number(to_char(adhocdate,'mm')) = " & v_txtmonth
		end if
		if v_txtEmployeeCode <> "" then
			strWhere3	=	" and empcode = '" & v_txtEmployeeCode & "'"
		end if
		
		strWhere4	=	" and fcba = '" & Session("fcBA") & "'"
		
		strGaji		=	strGaji + strWhere1 + strWhere2 + strWhere3 + strWhere4
		strLembur	=	strLembur + strWhere1 + strWhere2 + strWhere3 + strWhere4
		strPSB		=	strPSB + strWhere1 + strWhere2 + strWhere3 + strWhere4
		strPLB		=	strPLB + strWhere1 + strWhere2 + strWhere3 + strWhere4
		strPB		=	strPB + strWhere1 + strWhere2 + strWhere3 + strWhere4
		strPK		=	strPK + strWhere1 + strWhere2 + strWhere3 + strWhere4
		strPT		=	strPT + strWhere1 + strWhere2 + strWhere3 + strWhere4
		strPL		=	strPL + strWhere1 + strWhere2 + strWhere3 + strWhere4
		strTJ		=	strTJ + strWhere1 + strWhere2 + strWhere3 + strWhere4
		strsubPR	=	strsubPR + strWhere1 + strWhere2 + strWhere3 + strWhere4
		
		strGajiBruto = strGajiBruto + strWhere1 + strWhere2 + strWhere3 + strWhere4
		
		strPinjam 		= strPinjam + strWhere1 + strWhere2 + strWhere3 + strWhere4
		strPAK 			= strPAK + strWhere1 + strWhere2 + strWhere3 + strWhere4
		strMP1 			= strMP1 + strWhere1 + strWhere2 + strWhere3 + strWhere4
		strTDB 			= strTDB + strWhere1 + strWhere2 + strWhere3 + strWhere4
		strDPN 			= strDPN + strWhere1 + strWhere2 + strWhere3 + strWhere4
		strPLL 			= strPLL + strWhere1 + strWhere2 + strWhere3 + strWhere4
		strPPH21 		= strPPH21 + strWhere1 + strWhere2 + strWhere3 + strWhere4
		strPJ 			= strPJ + strWhere1 + strWhere2 + strWhere3 + strWhere4
		strsubPotong 	= strsubPotong + strWhere1 + strWhere2 + strWhere3 + strWhere4
		
		strGajiNetto = strGajiNetto + strWhere0 + strWhere1 + strWhere2 + strWhere3

		
		Set rsGaji		= Server.CreateObject("ADODB.Recordset")		
		Set rsLembur	= Server.CreateObject("ADODB.Recordset")
		Set rsPSB		= Server.CreateObject("ADODB.Recordset")		
		Set rsPLB		= Server.CreateObject("ADODB.Recordset")
		Set rsPB		= Server.CreateObject("ADODB.Recordset")		
		Set rsPK		= Server.CreateObject("ADODB.Recordset")
		Set rsPT		= Server.CreateObject("ADODB.Recordset")		
		Set rsPL		= Server.CreateObject("ADODB.Recordset")
		Set rsTJ		= Server.CreateObject("ADODB.Recordset")
		Set rssubPR		= Server.CreateObject("ADODB.Recordset")
		
		Set rsGajiBruto = Server.CreateObject("ADODB.Recordset")
		
		Set rsPinjam 	= Server.CreateObject("ADODB.Recordset")
		Set rsPAK 		= Server.CreateObject("ADODB.Recordset")
		Set rsMP1 		= Server.CreateObject("ADODB.Recordset")
		Set rsTDB 		= Server.CreateObject("ADODB.Recordset")
		Set rsDPN 		= Server.CreateObject("ADODB.Recordset")
		Set rsPLL 		= Server.CreateObject("ADODB.Recordset")
		Set rsPPH21 	= Server.CreateObject("ADODB.Recordset")
		Set rsPJ 		= Server.CreateObject("ADODB.Recordset")
		Set rssubPotong	= Server.CreateObject("ADODB.Recordset")
		
		Set rsGajiNetto = Server.CreateObject("ADODB.Recordset")

		
		rsGaji.Open strGaji, connstr, adOpenStatic
		rsLembur.Open strLembur, connstr, adOpenStatic
		rsPSB.Open strPSB, connstr, adOpenStatic
		rsPLB.Open strPLB, connstr, adOpenStatic
		rsPB.Open strPB, connstr, adOpenStatic
		rsPK.Open strPK, connstr, adOpenStatic
		rsPT.Open strPT, connstr, adOpenStatic
		rsPL.Open strPL, connstr, adOpenStatic
		rsTJ.Open strTJ, connstr, adOpenStatic
		rssubPR.Open strsubPR, connstr, adOpenStatic
		
		rsGajiBruto.Open strGajiBruto, connstr, adOpenStatic
		
		rsPinjam.Open strPinjam, connstr, adOpenStatic
		rsPAK.Open strPAK, connstr, adOpenStatic
		rsMP1.Open strMP1, connstr, adOpenStatic
		rsTDB.Open strTDB, connstr, adOpenStatic
		rsDPN.Open strDPN, connstr, adOpenStatic
		rsPLL.Open strPLL, connstr, adOpenStatic
		rsPPH21.Open strPPH21, connstr, adOpenStatic
		rsPJ.Open strPJ, connstr, adOpenStatic
		rssubPotong.Open strsubPotong, connstr, adOpenStatic
		
		rsGajiNetto.Open strGajiNetto, connstr, adOpenStatic

	end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>iPS Reporting - Employee Payslip</title>
<style type="text/css">
* {
margin:0;
padding:0;
}

.clear {
  clear: both;
  display: block;
  overflow: hidden;
  visibility: hidden;
  width: 0;
  height: 0;
}

.clearfix:before,
.clearfix:after {
  content: '\0020';
  display: block;
  overflow: hidden;
  visibility: hidden;
  width: 0;
  height: 0;
}

.clearfix:after {
  clear: both;
}

.clearfix {
  zoom: 1;
}

.left {
text-align:left;
}

.right {
text-align:right;
}

.center{
text-align:center;
}

h1,h2,h3,h4,h5,h6{
text-transform:capitalize;
}
body {
line-height:1.4em;
font-size:14px;
background:#000;
font-family:Verdana, Arial, Helvetica, sans-serif;
color:#fff;
}
#wrapper{
width:940px;
margin: 0 auto;
padding:20px;
background:#fff;
color:#000;
}
.fixed-1{
width:940px;
margin-bottom:20px;
}
</style>
</head>
<body>
	<div id="wrapper">
	<div id="header" class="fixed-1 clearfix left">
		<h2><%=Session("COMPANY")%></h2>
		<h3><%=Session("BA")%></h3>
	</div>
	<div id="subheader"class="fixed-1 clearfix center">
		<h4><u>Slip Gaji <%=rsDataEmployee("POSITIONNAME")%></u></h4>
		<h5><%=textMonth(v_txtmonth)%>, <%=v_txtyear%></h5>
	</div>
	<table>
		<tr>
			<td>Nama</td>
			<td>:</td>
			<td><%=rsDataEmployee("EMPLOYEENAME")%></td>
		</tr>
		<tr>
			<td>Jabatan</td>
			<td>:</td>
			<td><%=rsDataEmployee("POSITIONNAME")%>&nbsp;<%=rsDataEmployee("GRADECODE")%></td>
		</tr>
		<tr>
			<td>Status</td>
			<td>:</td>
			<td><%=rsDataEmployee("MARITIAL")%></td>
		</tr>
	</table>
	<table>
		<tr>
			<td>HK BULAN INI</td>
			<td>:</td>
			<td><%=rsDataEmployee("TOTALKJ")%></td>
			<TD>HARI</TD>
		</tr>
		<tr>
			<td>MANGKIR</td>
			<td>:</td>
			<td><%=rsDataEmployee("TOTALMK")%></td>
			<TD>HARI</TD>
		</tr>
		<tr>
			<td>P1</td>
			<td>:</td>
			<td><%=rsDataEmployee("TOTALP1")%></td>
			<TD>HARI</TD>
		</tr>
		<tr>
			<td>P2</td>
			<td>:</td>
			<td><%=rsDataEmployee("TOTALP2")%></td>
			<TD>HARI</TD>
		</tr>
		<tr>
			<td>SAKIT</td>
			<td>:</td>
			<td><%=rsDataEmployee("TOTALSK")%></td>
			<TD>HARI</TD>
		</tr>
		<tr>
			<td>CUTI</td>
			<td>:</td>
			<td><%=rsDataEmployee("TOTALCT")%></td>
			<TD>HARI</TD>
		</tr>
	</table>	
	<table>
		<tr>
			<td>A</td>
			<td>Gaji</td>
			<td></td>
			<td></td>
		</tr>
		<tr>
			<td></td>
			<td>Gaji Pokok</td>
			<td>Rp</td>
			<td><%=rsGaji("amount")%></td>
		</tr>
		<tr>
			<td>B</td>
			<td>Premi/ Lembur</td>
			<td></td>
			<td></td>
		</tr>
		<tr>
			<td></td>
			<td>Lembur</td>
			<td>Rp</td>
			<td><%=rsLembur("amount")%></td>
		</tr>
		<tr>
			<td></td>
			<td>Premi Siap Borong</td>
			<td>Rp</td>
			<td><%=rsPSB("amount")%></td>
		</tr>
		<tr>
			<td></td>
			<td>Premi Lebih Borong</td>
			<td>Rp</td>
			<td><%=rsPLB("amount")%></td>
		</tr>
		<tr>
			<td></td>
			<td>Premi Borongan</td>
			<td>Rp</td>
			<td><%=rsPB("amount")%></td>
		</tr>
		<tr>
			<td></td>
			<td>Premi Kerajinan</td>
			<td>Rp</td>
			<td><%=rsPK("amount")%></td>
		</tr>
		<tr>
			<td></td>
			<td>Premi Transport</td>
			<td>Rp</td>
			<td><%=rsPT("amount")%></td>
		</tr>
		<tr>
			<td></td>
			<td>Premi Lain-lain</td>
			<td>Rp</td>
			<td><%=rsPL("amount")%></td>
		</tr>
		<tr>
			<td></td>
			<td>Tunjangan Astek 4.54%</td>
			<td>Rp</td>
			<td><%=rsTJ("amount")%></td>
		</tr>
		<tr>
			<td></td>
			<td>Sub Total</td>
			<td>Rp</td>
			<td><%=rssubPR("amount")%></td>
		</tr>
		<tr>
			<td>C</td>
			<td>Gaji Bruto (A + B)</td>
			<td>Rp</td>
			<td><%=rsGajiBruto("amount")%></td>
		</tr>		
		<tr>
			<td>D</td>
			<td>Pengurangan</td>
			<td></td>
			<td></td>
		</tr>	
		<tr>
			<td></td>
			<td>Pinjaman ( Gaji, Uang, material )</td>
			<td>Rp</td>
			<td><%=rsPinjam("amount")%></td>
		</tr>
		<tr>
			<td></td>
			<td>Potongan Alat Kerja</td>
			<td>Rp</td>
			<td><%=rsPAK("amount")%></td>
		</tr>
		<tr>
			<td></td>
			<td>Mangkir + P1</td>
			<td>Rp</td>
			<td><%=rsMP1("amount")%></td>
		</tr>
		<tr>
			<td></td>
			<td>Tidak Dapat Basis</td>
			<td>Rp</td>
			<td><%=rsTDB("amount")%></td>
		</tr>
		<tr>
			<td></td>
			<td>Denda panen</td>
			<td>Rp</td>
			<td><%=rsDPN("amount")%></td>
		</tr>
		<tr>
			<td></td>
			<td>Potongan Lain - lain</td>
			<td>Rp</td>
			<td><%=rsPLL("amount")%></td>
		</tr>
		<tr>
			<td></td>
			<td>Potongan PPH 21</td>
			<td>Rp</td>
			<td><%=rsPPH21("amount")%></td>
		</tr>
		<tr>
			<td></td>
			<td>Iuran Astek 6,54%</td>
			<td>Rp</td>
			<td><%=rsPJ("amount")%></td>
		</tr>
		<tr>
			<td></td>
			<td>Sub Total</td>
			<td>Rp</td>
			<td><%=rssubPotong("amount")%></td>
		</tr>
		<tr>
			<td>E</td>
			<td>Gaji Netto (C - E)</td>
			<td>Rp</td>
			<td><%=rsGajiNetto("amount")%></td>
		</tr>	
	</table>
	</div> <!--end wrapper-->
	<div class = "clear" />
	
</body>
</html>	