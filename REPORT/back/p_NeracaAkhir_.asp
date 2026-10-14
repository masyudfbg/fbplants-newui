<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
Dim v_txtYear, v_txtMonth
Dim ScriptName, NO, cClass

v_txtYear		=	Request("txtYear")
v_txtMonth		=	Request("txtMonth")

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

if request("txtYear")="" then
   v_txtYear	=	Year(Session("v_Costbook"))
end if
if request("txtMonth")="" then
   v_txtMonth	=	Month(Session("v_Costbook"))
end if

Dim rsTipeBa, strTipeBa, validasiTipeBa

strTipeBa		=	" SELECT FCTYPE FROM BUSINESSUNIT WHERE FCCODE = '" & Session("FCBA") & "'"

Set rsTipeBa	=	Server.CreateObject("ADODB.Recordset")
rsTipeBa.Open strTipeBa, connstr, adLockOptimistic, adOpenStatic

if not rsTipeBa.eof then
validasiTipeBa	=	rsTipeBa("FCTYPE")
else
Response.Write  "<font face=""Verdana"" size=""1"" color=""red"">" & "iPla'S Error: Business Area ( Type ) is Empty" & "</font>"
err.number		=	1101
validasiTipeBa	=	""
end if
rsTipeBa.close
set rsTipeBa = Nothing
set strTipeBa = Nothing

Dim rsAktiva, strAktiva

if validasiTipeBa	=	"E" then
	strAktiva	=	" SELECT DECODE (R," & _
					" 1, 'KAS'," & _
					" 2, 'BANK KEBUN'," & _
					" 3, 'BANK PABRIK'," & _
					" 4, 'PINJAMAN KARYAWAN'," & _
					" 5, 'PINJAMAN KONTRAKTOR'," & _
					" 6, 'PINJAMAN PETANI PLASMA'," & _
					" 7, 'BIAYA DIBAYAR DIMUKA'," & _
					" 8, 'AKTIVA DALAM PROSES'," & _
					" 9, 'MATERIAL'," & _
					" 10, 'Bahan Bakar dan Pelumas'," & _
					" 11, 'Bahan Kimia'," & _
					" 12, 'Pupuk'," & _
					" 13, 'Suku Cadang'," & _
					" 14, 'Barang Umum'," & _
					" 15, 'Subtotal'," & _
					" 16, 'BIAYA LANGSUNG'," & _
					" 17, 'BIAYA TANAMAN MENGHASILKAN'," & _
					" 18, 'BIAYA TANAMAN BELUM MENGHASILKAN'," & _
					" 19, 'BIAYA BIBITAN'," & _
					" 20, 'INVESTASI PLASMA'," & _
					" 21, NULL," & _
					" 22, 'Rugi/ Laba Ditahan'," & _
					" 23, 'TOTAL AKTIVA')" & _
					" perkiraan," & _
					" DECODE (R," & _
					" 1, KAS," & _
					" 2, BANKKEBUN," & _
					" 3, BANKPABRIK," & _
					" 4, PINJAMANKARYAWAN," & _
					" 5, PINJAMANKONTRAKTOR," & _
					" 6, PINJAMANPETANIPLASMA," & _
					" 7, BIAYADIBAYARDIMUKA," & _
					" 8, AKTIVADALAMPROSES," & _
					" 9, NULL," & _
					" 10, BBM," & _
					" 11, KIMIA," & _
					" 12, PUPUK," & _
					" 13, SUKUCADANG," & _
					" 14, BARANGUMUM," & _
					" 15, SUBMATERIAL," & _
					" 16, NULL," & _
					" 17, BIAYATM," & _
					" 18, BIAYATBM," & _
					" 19, BIAYABIBIT," & _
					" 20, BIAYAINVESTASIPLASMA," & _
					" 21, NULL," & _
					" 22, RUGILABA, 23, AKTIVA)" & _
					" SALDODEBET," & _
					" DECODE (R," & _
					" 1, KASP," & _
					" 2, BANKKEBUNP," & _
					" 3, BANKPABRIKP," & _
					" 4, PINJAMANKARYAWANP," & _
					" 5, PINJAMANKONTRAKTORP," & _
					" 6, PINJAMANPETANIPLASMAP," & _
					" 7, BIAYADIBAYARDIMUKAP," & _
					" 8, AKTIVADALAMPROSESP," & _
					" 9, NULL," & _
					" 10, NULL," & _
					" 11, NULL," & _
					" 12, NULL," & _
					" 13, NULL," & _
					" 14, NULL," & _
					" 15, SUBMATERIALP," & _
					" 16, NULL," & _
					" 17, BIAYATMP," & _
					" 18, BIAYATBMP," & _
					" 19, BIAYABIBITP," & _
					" 20, BIAYAINVESTASIPLASMAP," & _
					" 21, NULL," & _
					" 22, RUGILABAP, 23, AKTIVAP)" & _
					" PERCENTAGE" & _
					" FROM ( SELECT FCBA," & _
					" NYEAR," & _
					" NMONTH," & _
					" SUM (KAS) KAS," & _
					" SUM (KAS / DECODE (AKTIVA, 0, 1, AKTIVA) * 100) KASP," & _
					" SUM (BANKKEBUN) BANKKEBUN," & _
					" SUM (BANKKEBUN / DECODE (AKTIVA, 0, 1, AKTIVA) * 100)" & _
					" BANKKEBUNP," & _
					" SUM (BANKPABRIK) BANKPABRIK," & _
					" SUM (BANKPABRIK / DECODE (AKTIVA, 0, 1, AKTIVA) * 100)" & _
					" BANKPABRIKP," & _
					" SUM (PINJAMANKARYAWAN) PINJAMANKARYAWAN," & _
					" SUM (PINJAMANKARYAWAN / DECODE (AKTIVA, 0, 1, AKTIVA) * 100)" & _
					" PINJAMANKARYAWANP," & _
					" SUM (PINJAMANKONTRAKTOR) PINJAMANKONTRAKTOR," & _
					" SUM (PINJAMANKONTRAKTOR / DECODE (AKTIVA, 0, 1, AKTIVA) * 100)" & _
					" PINJAMANKONTRAKTORP," & _
					" SUM (PINJAMANPETANIPLASMA) PINJAMANPETANIPLASMA," & _
					" SUM (" & _
					" PINJAMANPETANIPLASMA / DECODE (AKTIVA, 0, 1, AKTIVA) * 100)" & _
					" PINJAMANPETANIPLASMAP," & _
					" SUM (BIAYADIBAYARDIMUKA) BIAYADIBAYARDIMUKA," & _
					" SUM (" & _
					" BIAYADIBAYARDIMUKA / DECODE (AKTIVA, 0, 1, AKTIVA) * 100)" & _
					" BIAYADIBAYARDIMUKAP," & _
					" SUM (AKTIVADALAMPROSES) AKTIVADALAMPROSES," & _
					" SUM (AKTIVADALAMPROSES / DECODE (AKTIVA, 0, 1, AKTIVA) * 100)" & _
					" AKTIVADALAMPROSESP," & _
					" SUM (BBM) BBM," & _
					" SUM (KIMIA) KIMIA," & _
					" SUM (PUPUK) PUPUK," & _
					" SUM (SUKUCADANG) SUKUCADANG," & _
					" SUM (BARANGUMUM) BARANGUMUM," & _
					" SUM (SUBMATERIAL) SUBMATERIAL," & _
					" SUM (SUBMATERIAL / DECODE (AKTIVA, 0, 1, AKTIVA) * 100)" & _
					" SUBMATERIALP," & _
					" SUM (BIAYATBM) BIAYATBM," & _
					" SUM (BIAYATBM / DECODE (AKTIVA, 0, 1, AKTIVA) * 100) BIAYATBMP," & _
					" SUM (BIAYABIBIT) BIAYABIBIT," & _
					" SUM (BIAYABIBIT / DECODE (AKTIVA, 0, 1, AKTIVA) * 100)" & _
					" BIAYABIBITP," & _
					" SUM (BIAYAINVESTASIPLASMA) BIAYAINVESTASIPLASMA," & _
					" SUM (" & _
					" BIAYAINVESTASIPLASMA / DECODE (AKTIVA, 0, 1, AKTIVA) * 100)" & _
					" BIAYAINVESTASIPLASMAP," & _
					" SUM (RUGILABA) RUGILABA," & _
					" SUM (" & _
					" RUGILABA / DECODE (AKTIVA, 0, 1, AKTIVA) * 100)" & _
					" RUGILABAP," & _
					" SUM (BIAYATM) BIAYATM," & _
					" SUM (BIAYATM / DECODE (AKTIVA, 0, 1, AKTIVA) * 100) BIAYATMP," & _
					" SUM (AKTIVA) AKTIVA," & _
					" SUM (AKTIVA / DECODE (AKTIVA, 0, 1, AKTIVA) * 100) AKTIVAP" & _
					" FROM ( SELECT FCBA," & _
					" NYEAR," & _
					" NMONTH," & _
					" SUM (KAS) KAS," & _
					" SUM (BANKKEBUN) BANKKEBUN," & _
					" SUM (BANKPABRIK) BANKPABRIK," & _
					" SUM (PINJAMANKARYAWAN) PINJAMANKARYAWAN," & _
					" SUM (PINJAMANKONTRAKTOR) PINJAMANKONTRAKTOR," & _
					" SUM (PINJAMANPETANIPLASMA) PINJAMANPETANIPLASMA," & _
					" SUM (BIAYADIBAYARDIMUKA) BIAYADIBAYARDIMUKA," & _
					" SUM (AKTIVADALAMPROSES) AKTIVADALAMPROSES," & _
					" SUM (BBM) BBM," & _
					" SUM (KIMIA) KIMIA," & _
					" SUM (PUPUK) PUPUK," & _
					" SUM (SUKUCADANG) SUKUCADANG," & _
					" SUM (BARANGUMUM) BARANGUMUM," & _
					" SUM (SUBMATERIAL) SUBMATERIAL," & _
					" SUM (BIAYATBM) BIAYATBM," & _
					" SUM (BIAYABIBIT) BIAYABIBIT," & _
					" SUM (BIAYAINVESTASIPLASMA) BIAYAINVESTASIPLASMA," & _
					" SUM (RUGILABA) RUGILABA," & _
					" SUM (BIAYATM) BIAYATM," & _
					" SUM (KAS)" & _
					" + SUM (BANKKEBUN)" & _
					" + SUM (BANKPABRIK)" & _
					" + SUM (PINJAMANKARYAWAN)" & _
					" + SUM (PINJAMANKONTRAKTOR)" & _
					" + SUM (PINJAMANPETANIPLASMA)" & _
					" + SUM (BIAYADIBAYARDIMUKA)" & _
					" + SUM (AKTIVADALAMPROSES)" & _
					" + SUM (SUBMATERIAL)" & _
					" + SUM (BIAYATBM)" & _
					" + SUM (BIAYABIBIT)" & _
					" + SUM (BIAYAINVESTASIPLASMA)" & _
					" + SUM (RUGILABA)" & _
					" + SUM (BIAYATM)" & _
					" AKTIVA" & _
					" FROM ( SELECT TBGROUP.FCBA," & _
					" NYEAR," & _
					" NMONTH," & _
					" SUM (KAS) KAS," & _
					" SUM (BANKKEBUN) BANKKEBUN," & _
					" SUM (BANKPABRIK) BANKPABRIK," & _
					" SUM (PINJAMANKARYAWAN) PINJAMANKARYAWAN," & _
					" SUM (PINJAMANKONTRAKTOR) PINJAMANKONTRAKTOR," & _
					" SUM (PINJAMANPETANIPLASMA)" & _
					" PINJAMANPETANIPLASMA," & _
					" SUM (BIAYADIBAYARDIMUKA)+SUM (CLEARINGAKUN)" & _
					" BIAYADIBAYARDIMUKA," & _
					" SUM (AKTIVADALAMPROSES) AKTIVADALAMPROSES," & _
					" SUM (BBM) BBM," & _
					" SUM (KIMIA) KIMIA," & _
					" SUM (PUPUK) PUPUK," & _
					" SUM (SUKUCADANG) SUKUCADANG," & _
					" SUM (BARANGUMUM) BARANGUMUM," & _
					" SUM (SUBMATERIAL) SUBMATERIAL," & _
					" SUM (BIAYATBM) BIAYATBM," & _
					" SUM (BIAYABIBIT) BIAYABIBIT," & _
					" SUM (BIAYAINVESTASIPLASMA)" & _
					" BIAYAINVESTASIPLASMA," & _
					" SUM (RUGILABA) RUGILABA," & _
					" SUM (BIAYATM) BIAYATM" & _
					" FROM ( SELECT FCBA," & _
					" NYEAR," & _
					" NMONTH," & _
					" SUM (" & _
					" DECODE (JOBCODE," & _
					" '1110103', SALDO," & _
					" 0))" & _
					" KAS," & _
					" SUM (" & _
					" DECODE (JOBCODE," & _
					" '1110201', SALDO, '1110206', SALDO," & _
					" 0))" & _
					" BANKKEBUN," & _
					" SUM (" & _
					" DECODE (JOBCODE," & _
					" '1110202', SALDO," & _
					" 0))" & _
					" BANKPABRIK," & _
					" SUM (" & _
					" DECODE (JOBCODE," & _
					" '1130100', SALDO," & _
					" 0))" & _
					" PINJAMANKARYAWAN," & _
					" SUM (" & _
					" DECODE (JOBCODE," & _
					" '1130200', SALDO," & _
					" 0))" & _
					" PINJAMANKONTRAKTOR," & _
					" SUM (" & _
					" DECODE (JOBCODE," & _
					" '1130300', SALDO," & _
					" 0))" & _
					" PINJAMANPETANIPLASMA," & _
					" SUM (" & _
					" DECODE (JOBCODE," & _
					" '1140100', SALDO,'1140200', SALDO,'1140301', SALDO,'1140302', SALDO, '1140400', SALDO, '1140700', SALDO, '1140800', SALDO," & _
					" 0))" & _
					" BIAYADIBAYARDIMUKA," & _
					" SUM (" & _
					" DECODE (JOBCODE," & _
					" '1140600', SALDO," & _
					" 0))" & _
					" AKTIVADALAMPROSES," & _
					" SUM (" & _
					" DECODE (JOBCODE," & _
					" '1120400', SALDO," & _
					" 0))" & _
					" CLEARINGAKUN," & _
					" SUM (" & _
					" DECODE (JOBCODE," & _
					" '1210100', SALDO," & _
					" 0))" & _
					" BBM," & _
					" SUM (" & _
					" DECODE (JOBCODE," & _
					" '1210200', SALDO," & _
					" 0))" & _
					" KIMIA," & _
					" SUM (" & _
					" DECODE (JOBCODE," & _
					" '1210300', SALDO," & _
					" 0))" & _
					" PUPUK," & _
					" SUM (" & _
					" DECODE (JOBCODE," & _
					" '1210400', SALDO," & _
					" 0))" & _
					" SUKUCADANG," & _
					" SUM (" & _
					" DECODE (JOBCODE," & _
					" '1210500', SALDO," & _
					" 0))" & _
					" BARANGUMUM," & _
					" SUM (" & _
					" DECODE (JOBCODE," & _
					" '1210100', SALDO," & _
					" 0))" & _
					" + SUM (" & _
					" DECODE (JOBCODE," & _
					" '1210200', SALDO," & _
					" 0))" & _
					" + SUM (" & _
					" DECODE (JOBCODE," & _
					" '1210300', SALDO," & _
					" 0))" & _
					" + SUM (" & _
					" DECODE (JOBCODE," & _
					" '1210400', SALDO," & _
					" 0))" & _
					" + SUM (" & _
					" DECODE (JOBCODE," & _
					" '1210500', SALDO," & _
					" 0))" & _
					" SUBMATERIAL," & _
					" SUM (" & _
					" DECODE (JOBCODE," & _
					" '1330200', SALDO," & _
					" 0))" & _
					" BIAYATBM," & _
					" SUM (" & _
					" DECODE (JOBCODE," & _
					" '1330100', SALDO," & _
					" 0))" & _
					" BIAYABIBIT," & _
					" SUM (" & _
					" DECODE (JOBFORSTATUS," & _
					" '1340200', SALDO," & _
					" 0))" & _
					" BIAYAINVESTASIPLASMA," & _
					" SUM (" & _
					" DECODE (JOBCODE," & _
					" '3110300', SALDO," & _
					" 0))" & _
					" RUGILABA," & _
					" SUM (" & _
					" DECODE (JOBFORSTATUS," & _
					" 'TM', SALDO," & _
					" 0))" & _
					" BIAYATM" & _
					" FROM (SELECT TB.FCBA," & _
					" NYEAR," & _
					" NMONTH," & _
					" JOBCODE," & _
					" SALDO," & _
					" J.JOB_FORFIELDSTATUS" & _
					" JOBFORSTATUS" & _
					" FROM ( SELECT FCBA," & _
					" PERIOD_YEAR" & _
					" NYEAR," & _
					" PERIOD_MONTH" & _
					" NMONTH," & _
					" FCCODE JOBCODE," & _
					" SUM (" & _
					" ENDING_BALANCE)" & _
					" SALDO" & _
					" FROM TRIALBALANCE" & _
					" GROUP BY FCBA," & _
					" PERIOD_YEAR," & _
					" PERIOD_MONTH," & _
					" FCCODE) TB" & _
					" INNER JOIN" & _
					" JOB J" & _
					" ON (TB.FCBA = J.FCBA" & _
					" AND TB.JOBCODE =" & _
					" J.FCCODE))" & _
					" GROUP BY FCBA, NYEAR, NMONTH) TBGROUP" & _
					" INNER JOIN" & _
					" BUSINESSUNIT BU" & _
					" ON (TBGROUP.FCBA = BU.FCCODE)" & _
					" GROUP BY TBGROUP.FCBA, NYEAR, NMONTH)" & _
					" GROUP BY FCBA, NYEAR, NMONTH)" & _
					" GROUP BY FCBA, NYEAR, NMONTH)," & _
					" (SELECT ROWNUM R" & _
					" FROM ALL_OBJECTS" & _
					" WHERE ROWNUM <= 23)" & _
					" WHERE FCBA = '" & Session("FCBA") & "' AND NYEAR = " & v_txtYear & " AND NMONTH = " & v_txtmonth
elseif validasiTipeBa	=	"M" then
strAktiva	=	" SELECT DECODE (R," & _
				" 1, 'KAS'," & _
				" 2, 'BANK KEBUN'," & _
				" 3, 'BANK PABRIK'," & _
				" 4, 'PINJAMAN KARYAWAN'," & _
				" 5, 'PINJAMAN KONTRAKTOR'," & _
				" 6, 'PINJAMAN PETANI PLASMA'," & _
				" 7, 'BIAYA DIBAYAR DIMUKA'," & _
				" 8, 'AKTIVA DALAM PROSES'," & _
				" 9, 'MATERIAL'," & _
				" 10, 'Bahan Bakar dan Pelumas'," & _
				" 11, 'Bahan Kimia'," & _
				" 12, 'Pupuk'," & _
				" 13, 'Suku Cadang'," & _
				" 14, 'Barang Umum'," & _
				" 15, 'Subtotal'," & _
				" 16, 'BIAYA LANGSUNG'," & _
				" 17, 'BIAYA POM'," & _
				" 18, NULL," & _
				" 19, 'Rugi/ Laba Ditahan'," & _
				" 20, NULL," & _
				" 21, NULL," & _
				" 22, NULL," & _
				" 23, 'TOTAL AKTIVA')" & _
				" PERKIRAAN," & _
				" DECODE (R," & _
				" 1, KAS," & _
				" 2, BANKKEBUN," & _
				" 3, BANKPABRIK," & _
				" 4, PINJAMANKARYAWAN," & _
				" 5, PINJAMANKONTRAKTOR," & _
				" 6, PINJAMANPETANIPLASMA," & _
				" 7, BIAYADIBAYARDIMUKA," & _
				" 8, AKTIVADALAMPROSES," & _
				" 9, NULL," & _
				" 10, BBM," & _
				" 11, KIMIA," & _
				" 12, PUPUK," & _
				" 13, SUKUCADANG," & _
				" 14, BARANGUMUM," & _
				" 15, SUBMATERIAL," & _
				" 16, NULL," & _
				" 17, BIAYAPOM," & _
				" 18, NULL," & _
				" 19, RUGILABA," & _
				" 20, NULL," & _
				" 21, NULL," & _
				" 22, NULL," & _
				" 23, AKTIVA)" & _
				" SALDODEBET," & _
				" DECODE (R," & _
				" 1, KASP," & _
				" 2, BANKKEBUNP," & _
				" 3, BANKPABRIKP," & _
				" 4, PINJAMANKARYAWANP," & _
				" 5, PINJAMANKONTRAKTORP," & _
				" 6, PINJAMANPETANIPLASMAP," & _
				" 7, BIAYADIBAYARDIMUKAP," & _
				" 8, AKTIVADALAMPROSESP," & _
				" 9, NULL," & _
				" 10, NULL," & _
				" 11, NULL," & _
				" 12, NULL," & _
				" 13, NULL," & _
				" 14, NULL," & _
				" 15, SUBMATERIALP," & _
				" 16, NULL," & _
				" 17, BIAYAPOMP," & _
				" 18, RUGILABAP," & _
				" 19, NULL," & _
				" 20, NULL," & _
				" 21, NULL," & _
				" 22, NULL," & _
				" 23, AKTIVAP)" & _
				" PERCENTAGE" & _
				" FROM ( SELECT FCBA," & _
				" NYEAR," & _
				" NMONTH," & _
				" SUM (KAS) KAS," & _
				" SUM (KAS / DECODE (AKTIVA, 0, 1, AKTIVA) * 100) KASP," & _
				" SUM (BANKKEBUN) BANKKEBUN," & _
				" SUM (BANKKEBUN / DECODE (AKTIVA, 0, 1, AKTIVA) * 100)" & _
				" BANKKEBUNP," & _
				" SUM (BANKPABRIK) BANKPABRIK," & _
				" SUM (BANKPABRIK / DECODE (AKTIVA, 0, 1, AKTIVA) * 100)" & _
				" BANKPABRIKP," & _
				" SUM (PINJAMANKARYAWAN) PINJAMANKARYAWAN," & _
				" SUM (PINJAMANKARYAWAN / DECODE (AKTIVA, 0, 1, AKTIVA) * 100)" & _
				" PINJAMANKARYAWANP," & _
				" SUM (PINJAMANKONTRAKTOR) PINJAMANKONTRAKTOR," & _
				" SUM (PINJAMANKONTRAKTOR / DECODE (AKTIVA, 0, 1, AKTIVA) * 100)" & _
				" PINJAMANKONTRAKTORP," & _
				" SUM (PINJAMANPETANIPLASMA) PINJAMANPETANIPLASMA," & _
				" SUM (" & _
				" PINJAMANPETANIPLASMA / DECODE (AKTIVA, 0, 1, AKTIVA) * 100)" & _
				" PINJAMANPETANIPLASMAP," & _
				" SUM (BIAYADIBAYARDIMUKA) BIAYADIBAYARDIMUKA," & _
				" SUM (" & _
				" BIAYADIBAYARDIMUKA / DECODE (AKTIVA, 0, 1, AKTIVA) * 100)" & _
				" BIAYADIBAYARDIMUKAP," & _
				" SUM (AKTIVADALAMPROSES) AKTIVADALAMPROSES," & _
				" SUM (AKTIVADALAMPROSES / DECODE (AKTIVA, 0, 1, AKTIVA) * 100)" & _
				" AKTIVADALAMPROSESP," & _
				" SUM (BBM) BBM," & _
				" SUM (KIMIA) KIMIA," & _
				" SUM (PUPUK) PUPUK," & _
				" SUM (SUKUCADANG) SUKUCADANG," & _
				" SUM (BARANGUMUM) BARANGUMUM," & _
				" SUM (SUBMATERIAL) SUBMATERIAL," & _
				" SUM (SUBMATERIAL / DECODE (AKTIVA, 0, 1, AKTIVA) * 100)" & _
				" SUBMATERIALP," & _
				" SUM (RUGILABA) RUGILABA," & _
				" SUM (RUGILABA / DECODE (AKTIVA, 0, 1, AKTIVA) * 100)" & _
				" RUGILABAP," & _
				" SUM (BIAYAPOM) BIAYAPOM," & _
				" SUM (BIAYAPOM / DECODE (AKTIVA, 0, 1, AKTIVA) * 100) BIAYAPOMP," & _
				" SUM (AKTIVA) AKTIVA," & _
				" SUM (AKTIVA / DECODE (AKTIVA, 0, 1, AKTIVA) * 100) AKTIVAP" & _
				" FROM ( SELECT FCBA," & _
				" NYEAR," & _
				" NMONTH," & _
				" SUM (KAS) KAS," & _
				" SUM (BANKKEBUN) BANKKEBUN," & _
				" SUM (BANKPABRIK) BANKPABRIK," & _
				" SUM (PINJAMANKARYAWAN) PINJAMANKARYAWAN," & _
				" SUM (PINJAMANKONTRAKTOR) PINJAMANKONTRAKTOR," & _
				" SUM (PINJAMANPETANIPLASMA) PINJAMANPETANIPLASMA," & _
				" SUM (BIAYADIBAYARDIMUKA) BIAYADIBAYARDIMUKA," & _
				" SUM (AKTIVADALAMPROSES) AKTIVADALAMPROSES," & _
				" SUM (BBM) BBM," & _
				" SUM (KIMIA) KIMIA," & _
				" SUM (PUPUK) PUPUK," & _
				" SUM (SUKUCADANG) SUKUCADANG," & _
				" SUM (BARANGUMUM) BARANGUMUM," & _
				" SUM (SUBMATERIAL) SUBMATERIAL," & _
				" SUM (RUGILABA) RUGILABA," & _
				" SUM (BIAYAPOM) BIAYAPOM," & _
				" SUM (KAS)" & _
				" + SUM (BANKKEBUN)" & _
				" + SUM (BANKPABRIK)" & _
				" + SUM (PINJAMANKARYAWAN)" & _
				" + SUM (PINJAMANKONTRAKTOR)" & _
				" + SUM (PINJAMANPETANIPLASMA)" & _
				" + SUM (BIAYADIBAYARDIMUKA)" & _
				" + SUM (AKTIVADALAMPROSES)" & _
				" + SUM (SUBMATERIAL)" & _
				" + SUM (RUGILABA)" & _
				" + SUM (BIAYAPOM)" & _
				" AKTIVA" & _
				" FROM ( SELECT TBGROUP.FCBA," & _
				" NYEAR," & _
				" NMONTH," & _
				" SUM (KAS) KAS," & _
				" SUM (BANKKEBUN) BANKKEBUN," & _
				" SUM (BANKPABRIK) BANKPABRIK," & _
				" SUM (PINJAMANKARYAWAN) PINJAMANKARYAWAN," & _
				" SUM (PINJAMANKONTRAKTOR) PINJAMANKONTRAKTOR," & _
				" SUM (PINJAMANPETANIPLASMA)" & _
				" PINJAMANPETANIPLASMA," & _
				" SUM (BIAYADIBAYARDIMUKA)" & _
				" BIAYADIBAYARDIMUKA," & _
				" SUM (AKTIVADALAMPROSES) AKTIVADALAMPROSES," & _
				" SUM (BBM) BBM," & _
				" SUM (KIMIA) KIMIA," & _
				" SUM (PUPUK) PUPUK," & _
				" SUM (SUKUCADANG) SUKUCADANG," & _
				" SUM (BARANGUMUM) BARANGUMUM," & _
				" SUM (SUBMATERIAL) SUBMATERIAL," & _
				" SUM (RUGILABA) RUGILABA," & _
				" SUM (BIAYAPOM) BIAYAPOM" & _
				" FROM ( SELECT FCBA," & _
				" NYEAR," & _
				" NMONTH," & _
				" SUM (" & _
				" DECODE (JOBCODE," & _
				" '1110103', SALDO," & _
				" 0))" & _
				" KAS," & _
				" SUM (" & _
				" DECODE (JOBCODE," & _
				" '1110201', SALDO," & _
				" 0))" & _
				" BANKKEBUN," & _
				" SUM (" & _
				" DECODE (JOBCODE," & _
				" '1110202', SALDO," & _
				" 0))" & _
				" BANKPABRIK," & _
				" SUM (" & _
				" DECODE (JOBCODE," & _
				" '1130100', SALDO," & _
				" 0))" & _
				" PINJAMANKARYAWAN," & _
				" SUM (" & _
				" DECODE (JOBCODE," & _
				" '1130200', SALDO," & _
				" 0))" & _
				" PINJAMANKONTRAKTOR," & _
				" SUM (" & _
				" DECODE (JOBCODE," & _
				" '1130300', SALDO," & _
				" 0))" & _
				" PINJAMANPETANIPLASMA," & _
				" SUM (" & _
				" DECODE (JOBCODE," & _
				" '1140100', SALDO,'1140200', SALDO,'1140301', SALDO,'1140302', SALDO, '1140400', SALDO, '1140700', SALDO, '1140800', SALDO," & _
				" 0))" & _
				" BIAYADIBAYARDIMUKA," & _
				" SUM (" & _
				" DECODE (JOBCODE," & _
				" '1140600', SALDO," & _
				" 0))" & _
				" AKTIVADALAMPROSES," & _
				" SUM (" & _
				" DECODE (JOBCODE," & _
				" '1210100', SALDO," & _
				" 0))" & _
				" BBM," & _
				" SUM (" & _
				" DECODE (JOBCODE," & _
				" '1210200', SALDO," & _
				" 0))" & _
				" KIMIA," & _
				" SUM (" & _
				" DECODE (JOBCODE," & _
				" '1210300', SALDO," & _
				" 0))" & _
				" PUPUK," & _
				" SUM (" & _
				" DECODE (JOBCODE," & _
				" '1210400', SALDO," & _
				" 0))" & _
				" SUKUCADANG," & _
				" SUM (" & _
				" DECODE (JOBCODE," & _
				" '1210500', SALDO," & _
				" 0))" & _
				" BARANGUMUM," & _
				" SUM (" & _
				" DECODE (JOBCODE," & _
				" '1210100', SALDO," & _
				" 0))" & _
				" + SUM (" & _
				" DECODE (JOBCODE," & _
				" '1210200', SALDO," & _
				" 0))" & _
				" + SUM (" & _
				" DECODE (JOBCODE," & _
				" '1210300', SALDO," & _
				" 0))" & _
				" + SUM (" & _
				" DECODE (JOBCODE," & _
				" '1210400', SALDO," & _
				" 0))" & _
				" + SUM (" & _
				" DECODE (JOBCODE," & _
				" '1210500', SALDO," & _
				" 0))" & _
				" SUBMATERIAL," & _
				" SUM (" & _
				" DECODE (JOBCODE," & _
				" '3110300', SALDO," & _
				" 0))" & _
				" RUGILABA," & _
				" SUM (" & _
				" DECODE (JOBFORSTATUS," & _
				" 'POM', SALDO," & _
				" 0))" & _
				" BIAYAPOM" & _
				" FROM (SELECT TB.FCBA," & _
				" NYEAR," & _
				" NMONTH," & _
				" JOBCODE," & _
				" SALDO," & _
				" J.JOB_FORFIELDSTATUS" & _
				" JOBFORSTATUS" & _
				" FROM ( SELECT FCBA," & _
				" PERIOD_YEAR" & _
				" NYEAR," & _
				" PERIOD_MONTH" & _
				" NMONTH," & _
				" FCCODE JOBCODE," & _
				" SUM (" & _
				" ENDING_BALANCE)" & _
				" SALDO" & _
				" FROM TRIALBALANCE" & _
				" GROUP BY FCBA," & _
				" PERIOD_YEAR," & _
				" PERIOD_MONTH," & _
				" FCCODE) TB" & _
				" INNER JOIN" & _
				" JOB J" & _
				" ON (TB.FCBA = J.FCBA" & _
				" AND TB.JOBCODE =" & _
				" J.FCCODE))" & _
				" GROUP BY FCBA, NYEAR, NMONTH) TBGROUP" & _
				" INNER JOIN" & _
				" BUSINESSUNIT BU" & _
				" ON (TBGROUP.FCBA = BU.FCCODE)" & _
				" GROUP BY TBGROUP.FCBA, NYEAR, NMONTH)" & _
				" GROUP BY FCBA, NYEAR, NMONTH)" & _
				" GROUP BY FCBA, NYEAR, NMONTH)," & _
				" (SELECT ROWNUM R" & _
				" FROM ALL_OBJECTS" & _
				" WHERE ROWNUM <= 23)" & _
				" WHERE FCBA = '" & Session("FCBA") & "' AND NYEAR = " & v_txtYear & " AND NMONTH = " & v_txtmonth
else
	strAktiva	=	"NOQUERY"
end if

Dim rsPasiva, strPasiva

strPasiva	=	" SELECT DECODE (R," & _
				" 1, 'HUTANG'," & _
				" 2, 'Hutang Suplier Lokal'," & _
				" 3, 'Hutang Pokok Kontraktor'," & _
				" 4, 'Hutang Retensi'," & _
				" 5, 'Subtotal'," & _
				" 6, NULL," & _
				" 7, 'HUTANG PPH 21'," & _
				" 8, 'HUTANG GAJI'," & _
				" 9, 'HUTANG JAMSOSTEK'," & _
				" 10, 'Intraco HO'," & _
				" 11, 'Intraco RO'," & _
				" 12, 'Intraco KSB ( Kebun Sei Bangko )'," & _
				" 13, 'Intraco KSR ( Kebun Sei Rokan )'," & _
				" 14, 'Intraco KSD ( Kebun Simpang Damar )'," & _
				" 15, 'Intraco PSD ( Pabrik Simpang Damar )'," & _
				" 16, 'ACCRUAL BONUS'," & _
				" 17, 'ACCRUAL THR'," & _
				" 18, 'ACCRUED OTHERS'," & _
				" 19, NULL," & _
				" 20, NULL," & _
				" 21, NULL, 22, NULL, 23, 'TOTAL PASIVA')" & _
				" perkiraan," & _
				" DECODE (R," & _
				" 1, NULL," & _
				" 2, HUTANGSUPPLIERLOKAL*-1," & _
				" 3, HUTANGPOKOKKONTRAKTOR*-1," & _
				" 4, HUTANGRETENSI*-1," & _
				" 5, SUBHUTANG*-1," & _
				" 6, NULL," & _
				" 7, HUTANGPPH21*-1," & _
				" 8, HUTANGGAJI*-1," & _
				" 9, HUTANGJAMSOSTEK*-1," & _
				" 10, INTRACOHO*-1," & _
				" 11, INTRACORO*-1," & _
				" 12, INTRACOKSB*-1," & _
				" 13, INTRACOKSR*-1," & _
				" 14, INTRACOKSD*-1," & _
				" 15, INTRACOPSD*-1," & _
				" 16, ACCRUALBONUS*-1," & _
				" 17, ACCRUALTHR*-1," & _
				" 18, ACCRUEDOTHERS*-1," & _
				" 19, NULL," & _
				" 20, NULL," & _
				" 21, NULL, 22, NULL, 23, PASIVA*-1)" & _
				" SALDODEBET," & _
				" DECODE (R," & _
				" 1, NULL," & _
				" 2, NULL," & _
				" 3, NULL," & _
				" 4, NULL," & _
				" 5, SUBHUTANGP," & _
				" 6, NULL," & _
				" 7, HUTANGPPH21P," & _
				" 8, HUTANGGAJIP," & _
				" 9, HUTANGJAMSOSTEKP," & _
				" 10, INTRACOHOP," & _
				" 11, INTRACOROP," & _
				" 12, INTRACOKSBP," & _
				" 13, INTRACOKSRP," & _
				" 14, INTRACOKSDP," & _
				" 15, INTRACOPSDP," & _
				" 16, ACCRUALBONUSP," & _
				" 17, ACCRUALTHRP," & _
				" 18, ACCRUEDOTHERSP," & _
				" 19, NULL," & _
				" 20, NULL," & _
				" 21, NULL, 22, NULL , 23, PASIVAP)" & _
				" PERCENTAGE" & _
				" FROM ( SELECT FCBA," & _
				" NYEAR," & _
				" NMONTH," & _
				" SUM (HUTANGSUPPLIERLOKAL) HUTANGSUPPLIERLOKAL," & _
				" SUM (HUTANGPOKOKKONTRAKTOR) HUTANGPOKOKKONTRAKTOR," & _
				" SUM (HUTANGRETENSI) HUTANGRETENSI," & _
				" SUM (SUBHUTANG) SUBHUTANG," & _
				" SUM (SUBHUTANG / DECODE (PASIVA, 0, 1, PASIVA) * 100)" & _
				" SUBHUTANGP," & _
				" SUM (HUTANGPPH21) HUTANGPPH21," & _
				" SUM (HUTANGPPH21 / DECODE (PASIVA, 0, 1, PASIVA) * 100)" & _
				" HUTANGPPH21P," & _
				" SUM (HUTANGGAJI) HUTANGGAJI," & _
				" SUM (HUTANGGAJI / DECODE (PASIVA, 0, 1, PASIVA) * 100)" & _
				" HUTANGGAJIP," & _
				" SUM (HUTANGJAMSOSTEK) HUTANGJAMSOSTEK," & _
				" SUM (HUTANGJAMSOSTEK / DECODE (PASIVA, 0, 1, PASIVA) * 100)" & _
				" HUTANGJAMSOSTEKP," & _
				" SUM (INTRACOHO) INTRACOHO," & _
				" SUM (INTRACOHO / DECODE (PASIVA, 0, 1, PASIVA) * 100)" & _
				" INTRACOHOP," & _
				" SUM (INTRACORO) INTRACORO," & _
				" SUM (INTRACORO / DECODE (PASIVA, 0, 1, PASIVA) * 100)" & _
				" INTRACOROP," & _
				" SUM (INTRACOKSB) INTRACOKSB," & _
				" SUM (INTRACOKSB / DECODE (PASIVA, 0, 1, PASIVA) * 100)" & _
				" INTRACOKSBP," & _
				" SUM (INTRACOKSR) INTRACOKSR," & _
				" SUM (INTRACOKSR / DECODE (PASIVA, 0, 1, PASIVA) * 100)" & _
				" INTRACOKSRP," & _
				" SUM (INTRACOKSD) INTRACOKSD," & _
				" SUM (INTRACOKSD / DECODE (PASIVA, 0, 1, PASIVA) * 100)" & _
				" INTRACOKSDP," & _
				" SUM (INTRACOPSD) INTRACOPSD," & _
				" SUM (INTRACOPSD / DECODE (PASIVA, 0, 1, PASIVA) * 100)" & _
				" INTRACOPSDP," & _
				" SUM (ACCRUALBONUS) ACCRUALBONUS," & _
				" SUM (ACCRUALBONUS / DECODE (PASIVA, 0, 1, PASIVA) * 100)" & _
				" ACCRUALBONUSP," & _
				" SUM (ACCRUALTHR) ACCRUALTHR," & _
				" SUM (ACCRUALTHR / DECODE (PASIVA, 0, 1, PASIVA) * 100)" & _
				" ACCRUALTHRP," & _
				" SUM (ACCRUEDOTHERS) ACCRUEDOTHERS," & _
				" SUM (ACCRUEDOTHERS / DECODE (PASIVA, 0, 1, PASIVA) * 100)" & _
				" ACCRUEDOTHERSP," & _
				" SUM (PASIVA) PASIVA," & _
				" SUM (PASIVA / DECODE (PASIVA, 0, 1, PASIVA) * 100) PASIVAP" & _
				" FROM ( SELECT FCBA," & _
				" NYEAR," & _
				" NMONTH," & _
				" SUM (HUTANGSUPPLIERLOKAL) HUTANGSUPPLIERLOKAL," & _
				" SUM (HUTANGPOKOKKONTRAKTOR) HUTANGPOKOKKONTRAKTOR," & _
				" SUM (HUTANGRETENSI) HUTANGRETENSI," & _
				" SUM (SUBHUTANG) SUBHUTANG," & _
				" SUM (HUTANGPPH21) HUTANGPPH21," & _
				" SUM (HUTANGGAJI) HUTANGGAJI," & _
				" SUM (HUTANGJAMSOSTEK) HUTANGJAMSOSTEK," & _
				" SUM (INTRACOHO) INTRACOHO," & _
				" SUM (INTRACORO) INTRACORO," & _
				" SUM (INTRACOKSB) INTRACOKSB," & _
				" SUM (INTRACOKSR) INTRACOKSR," & _
				" SUM (INTRACOKSD) INTRACOKSD," & _
				" SUM (INTRACOPSD) INTRACOPSD," & _
				" SUM (ACCRUALBONUS) ACCRUALBONUS," & _
				" SUM (ACCRUALTHR) ACCRUALTHR," & _
				" SUM (ACCRUEDOTHERS) ACCRUEDOTHERS," & _
				" SUM (SUBHUTANG)" & _
				" + SUM (HUTANGPPH21)" & _
				" + SUM (HUTANGGAJI)" & _
				" + SUM (HUTANGJAMSOSTEK)" & _
				" + SUM (INTRACOHO)" & _
				" + SUM (INTRACORO)" & _
				" + SUM (INTRACOKSB)" & _
				" + SUM (INTRACOKSR)" & _
				" + SUM (INTRACOKSD)" & _
				" + SUM (INTRACOPSD)" & _
				" + SUM (ACCRUALBONUS)" & _
				" + SUM (ACCRUALTHR)" & _
				" + SUM (ACCRUEDOTHERS)" & _
				" PASIVA" & _
				" FROM ( SELECT FCBA," & _
				" NYEAR," & _
				" NMONTH," & _
				" SUM (DECODE (JOBCODE, '2110100', SALDO, '2110300', SALDO, 0))" & _
				" HUTANGSUPPLIERLOKAL," & _
				" SUM (DECODE (JOBCODE, '2120100', SALDO, 0))" & _
				" HUTANGPOKOKKONTRAKTOR," & _
				" SUM (DECODE (JOBCODE, '2120200', SALDO, 0))" & _
				" HUTANGRETENSI," & _
				" SUM (DECODE (JOBCODE, '2110100', SALDO, '2110300', SALDO, 0))" & _
				" + SUM (DECODE (JOBCODE, '2120100', SALDO, 0))" & _
				" + SUM (DECODE (JOBCODE, '2120200', SALDO, 0))" & _
				" SUBHUTANG," & _
				" SUM (DECODE (JOBCODE, '2140100', SALDO, 0))" & _
				" HUTANGPPH21," & _
				" SUM (DECODE (JOBCODE, '2160100', SALDO, 0))" & _
				" HUTANGGAJI," & _
				" SUM (DECODE (JOBCODE, '2160300', SALDO, '2160200', SALDO, 0))" & _
				" HUTANGJAMSOSTEK," & _
				" SUM (DECODE (JOBCODE, '1120101', SALDO, 0))" & _
				" INTRACOHO," & _
				" SUM (DECODE (JOBCODE, '1120102', SALDO, 0))" & _
				" INTRACORO," & _
				" SUM (DECODE (JOBCODE, '1120201', SALDO, 0))" & _
				" INTRACOKSB," & _
				" SUM (DECODE (JOBCODE, '1120202', SALDO, 0))" & _
				" INTRACOKSR," & _
				" SUM (DECODE (JOBCODE, '1120203', SALDO, 0))" & _
				" INTRACOKSD," & _
				" SUM (DECODE (JOBCODE, '1120204', SALDO, 0))" & _
				" INTRACOPSD," & _
				" SUM (DECODE (JOBCODE, '2160400', SALDO, 0))" & _
				" ACCRUALBONUS," & _
				" SUM (DECODE (JOBCODE, '2210101', SALDO, 0))" & _
				" ACCRUALTHR," & _
				" SUM (DECODE (JOBCODE, '2210102', SALDO, 0))" & _
				" ACCRUEDOTHERS" & _
				" FROM (SELECT FCBA," & _
				" NYEAR," & _
				" NMONTH," & _
				" JOBCODE," & _
				" SALDO" & _
				" FROM ( SELECT FCBA," & _
				" PERIOD_YEAR NYEAR," & _
				" PERIOD_MONTH NMONTH," & _
				" FCCODE JOBCODE," & _
				" SUM (ENDING_BALANCE) SALDO" & _
				" FROM TRIALBALANCE" & _
				" GROUP BY FCBA," & _
				" PERIOD_YEAR," & _
				" PERIOD_MONTH," & _
				" FCCODE))" & _
				" GROUP BY FCBA, NYEAR, NMONTH)" & _
				" GROUP BY FCBA, NYEAR, NMONTH)" & _
				" GROUP BY FCBA, NYEAR, NMONTH)," & _
				" (SELECT ROWNUM R" & _
				" FROM ALL_OBJECTS" & _
				" WHERE ROWNUM <= 23)" & _
				" WHERE FCBA = '" & Session("FCBA") & "' AND NYEAR = " & v_txtYear & " AND NMONTH = " & v_txtmonth


Set rsAktiva = Server.CreateObject("ADODB.Recordset")
rsAktiva.Open strAktiva, connstr, adOpenStatic

Set rsPasiva = Server.CreateObject("ADODB.Recordset")
rsPasiva.Open strPasiva, connstr, adOpenStatic

%>
<html>

<head>
<link rel="stylesheet" href="/css/Neraca.css" type="text/css" />
<title>iPla'S Reporting - Neraca Akhir</title>
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
body{font-family:'Arial Narrow';font-size:12px;}
.printhide{display:none;}
}
</style>
</head>

<body>
	<div id="wrapper">
	<div id="header-wrapper">
	<div class="printhide">
	<!--#include file="menutop.asp"-->
	</div>
	<div id="header">
		<h4>PT. <%=Session("COMPANY")%></h4>
		<h5><%=Session("BA")%></h5>
	</div>
	<div id="subheader" class="centered">
		<h2>NERACA AKHIR</h3>
		<br/>
		<h4>PERIODE&nbsp;:&nbsp;<%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
	</div>
	</div>
	<div class="printhide">
<table id="form-table" class="centered-margin">
	<form action="<%=ScriptName%>" method="post" name="frmEntry">
	<tr>
		<td><font face="verdana" size="1">YEAR</font></td>
<% 
	Dim mon, Yr, MinYr, MaxYr, i, MM
    mon = Month(Session("v_Costbook"))
    If len(v_txtMonth)=0 then v_txtMonth=mon

    Yr = Year(Session("v_Costbook"))
    MinYr = Yr - 2
    MaxYr = Yr  + 2
    If len(v_txtYear)= 0 then v_txtYear = Yr
%>
		<td width="4">
			<select name="txtYear">
<%
    For i = MinYr to MaxYr
%>
        <Option Value="<%=i%>" <%If cint(i) = cint(v_txtYear) then%> Selected <%End If%>><%=i%></option>
<%
    Next
%>
			</select>
		</td>
		<td><font face="verdana" size="1">MONTH</font></td>
		<td width="2">
			<select name="txtMonth">
<%
    For i = 1 to 12
%>
        <Option Value="<%=i%>" <%If cint(i) = cint(v_txtMonth) then%> Selected <%End If%>><%=Bulan(i)%></option>
<%
    Next
%>
			</select>
		</td>
		<td><input type="submit" value="View" /></td>
	</tr>
	</form>
</table>
</div>

<table border="0" width="100%" id="product-table">
	<tr>
		<td align="top">
			<table border="0" width="100%" id="#">
				<tr>
					<th align="center" colspan="2">PERKIRAAN</th>
					<th align="center">SALDO DEBET</th>
					<th align="center">%</th>
					<th align="center">&nbsp;</th>
				</tr>
				<tr><td colspan="4"><strong><i>AKTIVA</td></i></strong></tr>
<%
if rsAktiva.bof and rsAktiva.eof then
%>
	<tr><td colspan="4">NO DATA AVAILABLE</td></tr>
<%end if
		while not rsAktiva.eof
			if NO  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
<%if validasiTipeBa = "E" then%>
			<tr>
				<td width="1%"></td>
				<%if Not rsAktiva("PERKIRAAN") = "" then %>
					<%if rsAktiva("PERKIRAAN") = "MATERIAL" or rsAktiva("PERKIRAAN") = "BIAYA LANGSUNG" then%>				
						<td colspan="4"><strong><i><%=rsAktiva("PERKIRAAN")%></i></strong></td>
					<%elseif rsAktiva("PERKIRAAN") = "Subtotal" then%>
						<td align="left"><strong><i><%=rsAktiva("PERKIRAAN")%></i></strong></td>
						<td align="right"><strong><i><%=FormatNumber(rsAktiva("SALDODEBET"),0)%></i></strong></td>
						<td align="right"><strong><i><%=FormatNumber(rsAktiva("PERCENTAGE"),2)%></i></strong></td>
						<td align="left"><strong><i>%</i></strong></td>
					<%elseif rsAktiva("PERKIRAAN") = "TOTAL AKTIVA" then%>
						<th align="left"><strong><%=rsAktiva("PERKIRAAN")%></strong></th>
						<th align="right"><strong><i><%=FormatNumber(rsAktiva("SALDODEBET"),0)%></i></strong></th>
						<th align="right"><strong><i><%=FormatNumber(rsAktiva("PERCENTAGE"),2)%></i></strong></th>
						<th align="left"><strong><i>%</i></strong></th>
					<%else%>
						<td align="left"><%=rsAktiva("PERKIRAAN")%></td>
						<%if Not rsAktiva("SALDODEBET") = "" then%>
						<td align="right"><%=FormatNumber(rsAktiva("SALDODEBET"),0)%></td>
						<%end if%>
						<%if Not rsAktiva("PERCENTAGE") = "" then%>
						<td align="right"><%=FormatNumber(rsAktiva("PERCENTAGE"),2)%></td>
						<td align="left">%</i></td>
						<%end if%>
					<%end if%>
				<%else%>
						<td align="left">&nbsp;</td>
						<td align="right">&nbsp;</td>
						<td align="right">&nbsp;</td>
				<%end if%>
			</tr>
<%elseif validasiTipeBa = "M" then%>
			<tr>
				<td width="1%"></td>
				<%if Not rsAktiva("PERKIRAAN") = "" then %>
					<%if rsAktiva("PERKIRAAN") = "MATERIAL" or rsAktiva("PERKIRAAN") = "BIAYA LANGSUNG" then%>				
						<td colspan="4"><strong><i><%=rsAktiva("PERKIRAAN")%></i></strong></td>
					<%elseif rsAktiva("PERKIRAAN") = "Subtotal" then%>
						<td align="left"><strong><i><%=rsAktiva("PERKIRAAN")%></i></strong></td>
						<td align="right"><strong><i><%=FormatNumber(rsAktiva("SALDODEBET"),0)%></i></strong></td>
						<td align="right"><strong><i><%=FormatNumber(rsAktiva("PERCENTAGE"),2)%></i></strong></td>
						<td align="left"><strong><i>%</i></strong></td>
					<%elseif rsAktiva("PERKIRAAN") = "TOTAL AKTIVA" then%>
						<th align="left"><strong><%=rsAktiva("PERKIRAAN")%></strong></th>
						<th align="right"><strong><i><%=FormatNumber(rsAktiva("SALDODEBET"),0)%></i></strong></th>
						<th align="right"><strong><i><%=FormatNumber(rsAktiva("PERCENTAGE"),2)%></i></strong></th>
						<th align="left"><strong><i>%</i></strong></th>
					<%else%>
						<td align="left"><%=rsAktiva("PERKIRAAN")%></td>
						<%if Not rsAktiva("SALDODEBET") = "" then%>
						<td align="right"><%=FormatNumber(rsAktiva("SALDODEBET"),0)%></td>
						<%end if%>
						<%if Not rsAktiva("PERCENTAGE") = "" then%>
						<td align="right"><%=FormatNumber(rsAktiva("PERCENTAGE"),2)%></td>
						<td align="left">%</i></td>
						<%end if%>
					<%end if%>
				<%else%>
						<td align="left">&nbsp;</td>
						<td align="right">&nbsp;</td>
						<td align="right">&nbsp;</td>
				<%end if%>
			</tr>
<%end if%>
<%
			rsAktiva.MoveNext
		Wend
rsAktiva.Close
set rsAktiva = Nothing
%>
			</table>
		</td>
		<td>
			<table border="0" width="100%" id="#">
				<tr valign="top">
					<th align="center" colspan="2">PERKIRAAN</th>
					<th align="center">SALDO DEBET</th>
					<th align="center">%</th>
					<th align="center">&nbsp;</th>
				</tr>
				<tr><td colspan="4"><strong><i>PASIVA</td></i></strong></tr>
<%
if rsPasiva.bof and rsPasiva.eof then
%>
	<tr><td colspan="4">NO DATA AVAILABLE</td></tr>
<%end if
		while not rsPasiva.eof
			if NO  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
			<tr>
				<td width="1%"></td>
				<%if Not rsPasiva("PERKIRAAN") = "" then %>
					<%if rsPasiva("PERKIRAAN") = "HUTANG" then%>				
						<td colspan="4"><strong><i><%=rsPasiva("PERKIRAAN")%></i></strong></td>
					<%elseif rsPasiva("PERKIRAAN") = "Subtotal" then%>
						<td align="left"><strong><i><%=rsPasiva("PERKIRAAN")%></i></strong></td>
						<td align="right"><strong><i><%=FormatNumber(rsPasiva("SALDODEBET"),0)%></i></strong></td>
						<td align="right"><strong><i><%=FormatNumber(rsPasiva("PERCENTAGE"),2)%></i></strong></td>
						<td align="left"><strong><i>%</i></strong></td>
					<%elseif rsPasiva("PERKIRAAN") = "TOTAL PASIVA" then%>
						<th align="left"><strong><%=rsPasiva("PERKIRAAN")%></strong></th>
						<th align="right"><strong><i><%=FormatNumber(rsPasiva("SALDODEBET"),0)%></i></strong></th>
						<th align="right"><strong><i><%=FormatNumber(rsPasiva("PERCENTAGE"),2)%></i></strong></th>
						<th align="left"><strong><i>%</i></strong></th>
					<%else%>
						<td align="left"><%=rsPasiva("PERKIRAAN")%></td>
						<%if Not rsPasiva("SALDODEBET") = "" then%>
						<td align="right"><%=FormatNumber(rsPasiva("SALDODEBET"),0)%></td>
						<%end if%>
						<%if Not rsPasiva("PERCENTAGE") = "" then%>
						<td align="right"><%=FormatNumber(rsPasiva("PERCENTAGE"),2)%></td>
						<td align="left">%</i></td>
						<%end if%>
					<%end if%>
				<%else%>
						<td align="left">&nbsp;</td>
						<td align="right">&nbsp;</td>
						<td align="right">&nbsp;</td>
				<%end if%>
<%
			rsPasiva.MoveNext
		Wend
rsPasiva.Close
set rsPasiva = Nothing
%>
			</table>
		</td>
	</tr>
</table>

   <table width="90%" class="data">
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
		<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=Session("FCMANAGER")%></u></strong></td>
		<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=Session("KTU")%></u></strong></td>
	   </tr>
	   	  <tr> 
		<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>MANAGER</i></strong></td>
		<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>KTU</i></strong></td>
	   </tr>
   </table>
   </body>
   </html>
