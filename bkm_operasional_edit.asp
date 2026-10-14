<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->

<%
dim businessunit,fcbaname,CompanyName,q,sqltotdata,bkmno,fcbahdr,rkbno,postingflgtxt
bkmno		= trim(request("bkmno"))
'postingflg		= trim(request("postingflg"))
businessunit= trim(session("fcba"))

'response.write bkmno

q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & businessunit & "' "
	' response.write q
set rsCat = connStr.Execute(q)
if not rsCat.eof then
   fcbaname = rsCat("fcName")
   CompanyName=rsCat("FCCOMPANYNAME")
 
end if


' get data header
dim sqlheader,rsHeader,bkmnohdr,DATECREATED,SECTIONNAME,GANGCODE,ROWSTATEID,SUPERVISI1,SUPERVISI2,SUPERVISI3,SUPERVISI4,SUPERVISINAME1,SUPERVISINAME2,SUPERVISINAME3,SUPERVISINAME4,Clasification,Clasificationname,ROWSTATE
sqlheader = "select h.BKMNO,to_char(h.BKMDATE,'mm/dd/yyyy')BKMDATE,h.GANGCODE,h.SUPERVISI1,(SELECT s1.FCNAME FROM EMPLOYEE  s1 WHERE h.SUPERVISI1=s1.FCCODE and s1.fcba=h.fcba ) AS SUPERVISINAME1, " & _
" h.SUPERVISI2,(SELECT s2.FCNAME FROM EMPLOYEE s2 WHERE h.SUPERVISI2=s2.FCCODE and s2.fcba=h.fcba) AS SUPERVISINAME2,h.SUPERVISI3, " & _
" (SELECT s3.FCNAME FROM EMPLOYEE s3 WHERE h.SUPERVISI3=s3.FCCODE and s3.fcba=h.fcba) AS SUPERVISINAME3,h.SUPERVISI4, " & _
" (SELECT s4.FCNAME FROM EMPLOYEE s4 WHERE h.SUPERVISI4=s4.FCCODE and s4.fcba=h.fcba) AS SUPERVISINAME4,h.SECTIONCODE,h.DOCUMENT_CLASSIFICATION,c.fcname,h.fcba, h.ROWSTATE,h.POSTINGFLG,h.rkbno " & _
" from bkmhdr h,DOCUMENT_CLASSIFICATION c where h.DOCUMENT_CLASSIFICATION=c.FCCODE and h.BKMNO='" & bkmno & "'"
'response.write sqlheader
set rsHeader = connStr.Execute(sqlheader)

if not rsHeader.eof then 
	bkmnohdr = rsHeader("BKMNO")
	DATECREATED = rsHeader("BKMDATE")
	SECTIONNAME = rsHeader("SECTIONCODE")
	GANGCODE = rsHeader("GANGCODE")
	SUPERVISI1 = rsHeader("SUPERVISI1")
	SUPERVISI2 = rsHeader("SUPERVISI2")
	SUPERVISI3 = rsHeader("SUPERVISI3")
	SUPERVISI4 = rsHeader("SUPERVISI4")
	SUPERVISINAME1 = rsHeader("SUPERVISINAME1")
	SUPERVISINAME2 = rsHeader("SUPERVISINAME2")
	SUPERVISINAME3 = rsHeader("SUPERVISINAME3")
	SUPERVISINAME4 = rsHeader("SUPERVISINAME4")
	Clasification = rsHeader("DOCUMENT_CLASSIFICATION")
	Clasificationname 	= rsHeader("fcname")
	postingflgtxt = rsHeader("POSTINGFLG")
	ROWSTATE = rsHeader("ROWSTATE")
	fcbahdr = rsHeader("fcba")
	rkbno = rsHeader("rkbno")
	
	
end if

IF postingflgtxt = "0" THEN 
   postingflgtxt = "TIDAK"
ELSE
   postingflgtxt	= "YA"
END IF

'count detailupah
Dim SQL_Rows, NoOfRows,RSCount

SQL_Rows = "SELECT Count (*) from BKMUPAHDTL " & _ 
" where BKMNO='" & bkmno & "'"
set RSCount = connStr.Execute(SQL_Rows)
if not RSCount.eof then 
	NoOfRows = RSCount(0).value
end if
RSCount.close

'count detailbahan
Dim SQL_Rowsitem, NoOfRowsitem,RSCountitem

SQL_Rowsitem = "SELECT Count (*) from BKMBAHANDTL where BKMNO='" & bkmno & "'"
set RSCountitem = connStr.Execute(SQL_Rowsitem)
if not RSCountitem.eof then 
	NoOfRowsitem = RSCountitem(0).value
end if
RSCountitem.close


%>

<html>
	<head>
		<title>FBPlants - BKM Operasional</title>
		<link href="/js/tabbedpanels/SpryTabbedPanels.css" rel="stylesheet" type="text/css"/>
		<script src="/js/tabbedpanels/SpryTabbedPanels.js" type="text/javascript"></script>
		<script src="/js/jquery-1.9.1.js"></script>
		<script src="/js/jquery-ui.js"></script>
		<link rel="stylesheet" href="css/jquery-ui.css" />
		<link rel="StyleSheet" href="style-content.css" type="text/css" />
		<link href="/js/spry/widgets/textfieldvalidation/SpryValidationTextField.css" rel="stylesheet" type="text/css"/>
		<script src="/js/spry/widgets/textfieldvalidation/SpryValidationTextField.js" type="text/javascript"></script>

		<link href="/js/spry/widgets/selectvalidation/SpryValidationSelect.css" rel="stylesheet" type="text/css"/>
		<script src="/js/spry/widgets/selectvalidation/SpryValidationSelect.js" type="text/javascript"></script>
	<% if Request("oldui") <> "1" then %>
<!-- Tampilan baru (css/fbp-shell.css + css/fbp-forms.css). ?oldui=1 = tampilan lama -->
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
<link rel="StyleSheet" href="/css/login/font-awesome.min.css" type="text/css" />
<link rel="StyleSheet" href="/css/fbp-shell.css?v=20261009d" type="text/css" />
<link rel="StyleSheet" href="/css/fbp-forms.css?v=20261009e" type="text/css" />
<script type="text/javascript" src="/js/fbp-shell.js?v=20261014a"></script>
<% end if %>
</head>

<style>
	.posisikirijudul{
		position: relative;
		left: 0px;
	}
	.posisikiri{
		position: relative;
		left: -15px;
	}
	.lebarsection{
		width: 160px;
	}
	.lebargangcode{
		width: 200px;
	}
	.lebarhk{
		width: 45px;
	}
	.lebarabsen{
		width: 48px;
	}
	.lebarlokasi{
		width: 60px;
	}
	.lebartc{
		width: 70px;
	}
	.lebarrate{
		width: 60px;
	}
	.lebaractivity{
		width: 110px;
	}
	.lebarallow{
		width: 90px;
	}
	.lebarupah{
		width: 80px;
	}
	.warnagaris{
		border-color: #FFA800;
	}
	.tebal{
		font-weight: bold;
	}
	.ratakiri{
		text-align: left;
	}
	.ratakanan{
		text-align: right;
	}
	.ratatengah{
		text-align: center;
	}
	.lebaritembarang{
		width: 400px;
	}
	.lebaractbarang{
		width: 300px;
	}
	.lebarvehicle{
		width: 200px;
	}
	.lebartime{
		width: 90px;
	}

	.lebarspkno{
		width: 250px;
	}
	.lebarovtime{
		width: 70px;
	}
	.warnaamount{
		background-color: #eee;
	}
	
</style>



<body<% if Request("oldui") <> "1" then %> class="dash page-show page-legacy page-bkm"<% end if %>>
<!--#include file="top.jsp"-->

<br>
<br>
<br>
<br>
<br>


<div id="head_judul" style="float:left;font-size:10px; text-align:left;">
		<strong class="posisikiri">&nbsp;&nbsp;PT. <%=CompanyName%> - &nbsp;<%=FcbaName%>   -           MODULE : BKM OPERASIONAL EDIT</strong><br>					
</div>
<br>
<br>
<div class="box_bothall" style="overflow:scroll;">			
<form name="frmEntry" method="post" action="bkm_operasional_prosesedit.asp">
	<input type="hidden" name="jumldata" id="jumldata" value="<%=NoOfRows%>" />
	<input type="hidden" name="jumldataitem" id="jumldataitem" value="<%=NoOfRowsitem%>" />
	<input type="hidden" name="fcbahdr" id="fcbahdr" value="<%=fcbahdr%>" />
	<input type="hidden" name="rowstatedata" id="rowstatedata" value="<%=ROWSTATE%>" />
	<input type="hidden" name="proses" id="proses" value="2" />
	<table class="data posisikiri" border="0" width="90%">			
		<tr>
			<td align="left" >BKM No</td>
			<td align="left">
				<input type="text" name="bkmno" readonly id="bkmno" size="30" value="<%=bkmno%>" title="BKM No">
				
			</td>

			<td align="left" >RKB No</td>
			<td align="left">
				<input type="text" name="rkbno" readonly id="rkbno" size="30" value="<%=rkbno%>" title="RKB No">
				
			</td>
		</tr>
		<tr>
			
			<td align="left" >Tanggal</td>
			<td align="left">
				<input type="text" class="" readonly maxlength="15" name="tglbkm" readonly id="tglbkm" size="15" value="<%=DATECREATED%>" title="Tanggal BKM" onchange="get_Allowentry();">
			</td>
		

			<td align="left" >Section Name</td>
			<td align="left">
				<input type="text" name="sectionname" readonly id="sectionname" size="20" value="<%=SECTIONNAME%>">				
			</td>
			
			<td align="left" >Gang</td>
			<td align="left">
				<input type="text" name="gangcode"  readonly id="gangcode" size="15" value="<%=GANGCODE%>">
			</td>
			
			
		</tr>
		<tr>
			<td align="left" >Mandor I</td>
			<td align="left">
				<select name="supervisi1" id="supervisi1" class="lebarsection" >
					<option value="<%=SUPERVISI1%>"><%=SUPERVISINAME1%></option>
				</select>
			</td>
			<td align="left" >Mandor</td>
			<td align="left">
				<select name="supervisi2" id="supervisi2" class="lebarsection" >
					<option value="<%=SUPERVISI2%>"><%=SUPERVISI2%></option>
				</select>
			</td>
			<td align="left" >Kerani Panen</td>
			<td align="left">
				<select name="supervisi3" id="supervisi3" class="lebarsection" >
					<option value="<%=SUPERVISI3%>"><%=SUPERVISINAME3%></option>
				</select>
			</td>
			<td align="left" >Kerani Transport</td>
			<td align="left">
				<select name="supervisi4" id="supervisi4" class="lebarsection" >
					<option value="<%=SUPERVISI4%>"><%=SUPERVISINAME4%></option>
				</select>
			</td>
		</tr>
		<tr>
			<td align="left" >Posting</td>
			<td align="left"><%=postingflgtxt%>-<%=ROWSTATE%> <input type="hidden" value="<%=ROWSTATE%>" id="statusbkm"/></td>
			<td align="left">&nbsp;</td>
			<td align="left"><span id="msg" >-</span></td>
		</tr>
	</table>
	
	<br>
		
	<table align=center border="0" class='data posisikiri' width='900' id='tableall'>
		<tr>
			<td>
				<div class="TabbedPanels" id="tp1" align='center'>
				
					<ul class="TabbedPanelsTabGroup">
						<li class="TabbedPanelsTab">BKM UPAH</li>
						<li class="TabbedPanelsTab">BKM BAHAN</li>
						
					</ul>

					<div class="TabbedPanelsContentGroup">

						<!-- view bkm upah-->
						<div class="TabbedPanelsContent">
							<table align=center border="0" width="600" class='data' id='table2'>
								<tr id="trheader2">
									<th>Employee</th>
									<th>TimeIn</th>
									<th>TimeOut</th>
									<th>Lokasi</th>
									<th>No.RKB</th>
									<th>Absen</th>
									<th>Activity</th>
									<th>TC</th>
									<th>HK</th>
									<th>Unit</th>
									<th>Uom</th>
									<th>Output</th>
									<th>Rate</th>
									<th>Allowed Code</th>
									<th>Rate Allowed</th>
									<th>OVTH</th>
									<th>TOVTime</th>
									<th>Remark</th>
									<th colspan="2">&nbsp;</th>
								</tr>

								<% 
								
								
								'get data detail upah
								dim sqldtlupah,rsdtlupah,no,totalamount,ttlhk,ttlunit,ttloutput,ttlrate,ttlrateall,ttlovth
								sqldtlupah = "SELECT u.BKMNO,u.chargeunit fcba, u.EMPLOYEECODE,e.fcname || '-' ||  e.fccode as fcname, u.ACTCODE,a.fcname || '-' ||  u.ACTCODE as activityname, NVL(u.MANDAYS,0)MANDAYS," & _
								" u.ALLOWANCE_CODE, nvl(u.ALLOWANCE_AMT,0)ALLOWANCE_AMT, u.ATTENDANCE,nvl(u.OVERTIME_HOURS,0)OVERTIME_HOURS,u.TYPE_OVERTIME,o.fcname type_ovname," & _ 
								" to_char(u.IN_FIN,'hh24:mi:ss')IN_FINDEPAN,to_char(u.OUT_FIN,'hh24:mi:ss')OUT_FINDEPAN, " & _
								" to_char(u.IN_FIN,'dd/mm/yyyy hh24:mi:ss')IN_FIN,to_char(u.OUT_FIN,'dd/mm/yyyy hh24:mi:ss')OUT_FIN,  " & _
								" u.LOCATIONCODE, nvl(u.OUTPUT,0)OUTPUT,nvl(u.rate,0)rate, nvl(u.UNIT,0)UNIT,al.FCCODE || ' - ' || al.fcname as ALLOWDEDNAME,nvl(u.OTHCONVERTION,0)OTHCONVERTION,u.remarks,a.uom_unit,u.rkbno " & _
								" FROM BKMUPAHDTL u " & _ 
								" INNER JOIN employee e ON (u.EMPLOYEECODE=e.fccode and u.fcba=e.fcba) " & _
								" INNER JOIN JOBMASTER a ON (u.ACTCODE=a.FCCODE) " & _
								" LEFT JOIN ALLOWDED al ON (al.FCCODE = u.ALLOWANCE_CODE and al.fcba=u.fcba) " & _ 
								" LEFT JOIN overtime_convertion o ON (o.type=u.TYPE_OVERTIME and o.HOURS=u.OVERTIME_HOURS and o.fcba=u.fcba)" & _
								" where u.BKMNO='" & bkmno & "' order by e.fcname, a.fcname,u.LOCATIONCODE,u.ALLOWANCE_CODE"
								set rsdtlupah = connStr.Execute(sqldtlupah)
								'response.write sqldtlupah
								

																
								no = 0
								ttlhk = 0
								ttlunit = 0
								ttloutput = 0
								ttlrate = 0
								ttlrateall = 0
								ttlovth = 0
								totalamount = 0
								While Not rsdtlupah.EOF %>
								<tr id="tru<%=no%>">
									
									<td class="ratakiri">
										<select class="lebaractivity " name="fccode<%=no%>" id="fccode<%=no%>" onchange="getFinger(<%=no%>)" onClick="getFinger(<%=no%>)">
											<option value=<%=rsdtlupah("EMPLOYEECODE")%>><%=rsdtlupah("fcname")%></option>
										</select>
									</td>
									<td class="ratatengah">
										<%=rsdtlupah("IN_FINDEPAN")%>
										<span id="timein<%=no%>"></span>
										<input type="hidden" class="lebarabsen" name="timeinx<%=no%>" id="timeinx<%=no%>" value="<%=rsdtlupah("IN_FIN")%>">	
									</td>
									<td class="ratatengah">
										<%=rsdtlupah("OUT_FINDEPAN")%>
										<span id="timeout<%=no%>"></span>
										<input type="hidden" class="lebarabsen" name="timeoutx<%=no%>" id="timeoutx<%=no%>" value="<%=rsdtlupah("OUT_FIN")%>">	
									</td>
									<td class="ratatengah">
										<select class="lebarabsen " name="lokasi<%=no%>" id="lokasi<%=no%>" onchange="getNoRKBFCBA(<%=no%>)">
											<option value=<%=rsdtlupah("fcba")%>><%=rsdtlupah("fcba")%></option>
										</select>
									</td>
									<td class="ratatengah">
										<select class="lebaractivity ratakiri" name="rkbnofcba<%=no%>" id="rkbnofcba<%=no%>" onchange="showActivity(<%=no%>)">
											<option value=<%=rsdtlupah("rkbno")%>><%=rsdtlupah("rkbno")%></option>
										</select>
									</td>
									
									<td class="ratatengah">
										<select class="lebarabsen " name="absen<%=no%>" id="absen<%=no%>" onchange="getDisable(<%=no%>)">
											<option value=<%=rsdtlupah("ATTENDANCE")%>><%=rsdtlupah("ATTENDANCE")%></option>
										</select>
									</td>
									<td class="ratakiri">
										<select class="lebaractivity ratakiri" name="activity<%=no%>" id="activity<%=no%>" onchange="getTC(<%=no%>);">
											<option value=<%=rsdtlupah("ACTCODE")%>><%=rsdtlupah("activityname")%></option>
										</select>
									</td>
									<td class="ratatengah">
										<select class="lebartc ratakiri" name="tcupah<%=no%>" id="tcupah<%=no%>">
											<option value=<%=rsdtlupah("LOCATIONCODE")%>><%=rsdtlupah("LOCATIONCODE")%></option>
										</select>
									</td>
									<td>
										<input class="ratatengah lebarhk" type="text" onkeyup="getHKhi(<%=no%>);getValidasiHK(<%=no%>);hanyaAngka(this)" name="hk<%=no%>" id="hk<%=no%>" value="<%=rsdtlupah("MANDAYS")%>"/>
										<input class="ratatengah" type="hidden" onkeyup="hanyaAngka(this)" name="ttlhkkhl' + no + '" id="ttlhkkhl' + no + '" value="' + ttlmandays + '"/>
										<input class="ratatengah" type="hidden" onkeyup="hanyaAngka(this)" name="hkhidden<%=no%>" id="hkhidden<%=no%>" value="<%=formatnumber(rsdtlupah("MANDAYS"),2)%>"/>
									</td>
									
									<td>
										<input class="ratatengah lebarabsen" type="text" name="unit<%=no%>" id="unit<%=no%>" value="<%=rsdtlupah("UNIT")%>"/>
									</td>
									
									<td>
										<input class="ratatengah lebarabsen" type="text" name="uom_unit<%=no%>" id="uom_unit<%=no%>" value="<%=rsdtlupah("uom_unit")%>"/>
									</td>
									
									<td class="ratakiri">
										<input class="ratatengah lebarabsen" type="text" name="output<%=no%>" id="output<%=no%>" value="<%=rsdtlupah("OUTPUT")%>"/>
									</td>
									
									<td class="ratakiri">
										<input class="ratatengah lebarrate" type="text" name="ratekb<%=no%>" id="ratekb<%=no%>" value="<%=rsdtlupah("rate")%>"/>
									</td>
									<td class="ratakiri">
										<select class="lebarallow ratakiri" name="allowedcode<%=no%>" id="allowedcode<%=no%>">
											<option value=<%=rsdtlupah("ALLOWANCE_CODE")%>><%=rsdtlupah("ALLOWDEDNAME")%></option>
										</select>
									</td>
																		
									<td class="ratakiri">
										<input class="ratakanan lebarrate" type="text" onkeyup="getIsibelakang(<%=no%>)"  name="ratepremidepan<%=no%>" id="ratepremidepan<%=no%>" value="<%=FormatNumber(rsdtlupah("ALLOWANCE_AMT"),0)%>"/>
										<input class="ratakanan lebarrate" type="hidden"   name="ratepremi<%=no%>" id="ratepremi<%=no%>" value="<%=rsdtlupah("ALLOWANCE_AMT")%>"/>
									</td>
									<td class="ratatengah">
										<input class="ratatengah lebarhk" type="text"   name="ovtimehours<%=no%>" id="ovtimehours<%=no%>" value="<%=rsdtlupah("OVERTIME_HOURS")%>"/>
									</td>
									<td class="ratakiri">
										<select class="lebarovtime " name="typeovrtime<%=no%>" id="typeovrtime<%=no%>" onchange="getOTHConvertion(<%=no%>)">
											<option value=<%=rsdtlupah("TYPE_OVERTIME")%>><%=rsdtlupah("type_ovname")%></option>

										</select>
										<input type="hidden"   name="othconvertion<%=no%>" id="othconvertion<%=no%>" value="<%=rsdtlupah("othconvertion")%>"/>
									</td>
									<td class="ratakiri">
										<textarea  name="remark<%=no%>" id="remark<%=no%>" rows="3" cols="10"><%=rsdtlupah("remarks")%></textarea>
									</td>
									
									<td>
										<img src='css/login/images/add.jpg' onclick="addrowUpah('<%=no%>')" />
									</td>
									<td>
										<% if no >= 0 then %>
											<img src='css/login/images/del.png' onclick="deleterowUpah('<%=no%>')" />
										<% else 
											response.write "&nbsp;"
										end if %>
									</td>
									<td></td>
									
								</tr>
								<%
									'dim ttlhk
									no = no +1
									ttlhk = cdbl(ttlhk) + cdbl(rsdtlupah("MANDAYS"))
									ttlunit = cdbl(ttlunit) + cdbl(rsdtlupah("unit"))
									ttloutput = cdbl(ttloutput) + cdbl(rsdtlupah("output"))
									ttlrate = cdbl(ttlrate) + cdbl(rsdtlupah("rate"))
									ttlrateall = cdbl(ttlrateall) + cdbl(rsdtlupah("ALLOWANCE_AMT"))
									ttlovth = cdbl(ttlovth) + cdbl(rsdtlupah("OVERTIME_HOURS"))
									
									
									rsdtlupah.MoveNext
									Wend
								%>
									<tr class="ratakanan">
										<td colspan="8" class="tebal">Total</td>
										<td class="tebal"><input class="ratatengah lebarhk" type="text" readonly  style="background-color:#eee;" name="ttlhk" id="ttlhk<%=no%>" value="<%=FormatNumber(ttlhk,2)%>"/></td>
										<td class="tebal ratakiri"><input class="ratatengah lebarhk" type="text" readonly  style="background-color:#eee;" name="ttlhk" id="ttlhk<%=no%>" value="<%=FormatNumber(ttlunit,2)%>"/></td>
										<td></td>
										<td class="tebal ratakiri"><input class="ratatengah lebarhk" type="text" readonly  style="background-color:#eee;" name="ttlhk" id="ttlhk<%=no%>" value="<%=FormatNumber(ttloutput,2)%>"/></td>
										<td class="tebal ratakiri"><input class="ratatengah lebarrate" type="text" readonly  style="background-color:#eee;" name="ttlhk" id="ttlhk<%=no%>" value="<%=FormatNumber(ttlrate,2)%>"/></td>
										<td></td>
										<td class="tebal ratakiri"><input class="ratakanan lebarrate" type="text" readonly  style="background-color:#eee;" name="amount" id="amount<%=no%>" value="<%=FormatNumber(ttlrateall,0)%>"/></td>
										<td class="tebal ratatengah"><input class="ratatengah lebarhk" type="text" readonly  style="background-color:#eee;" name="amount" id="amount<%=no%>" value="<%=FormatNumber(ttlovth,0)%>"/></td>
										<td></td>
										
									</tr>
									<tr class="ratatengah">
										<%
											if postingflgtxt = "TIDAK" then%>
										<td  colspan="15"><input type=Submit name="btnsave" id="SAVE_UPAH" value="SAVE UPAH">&nbsp;&nbsp;<input type=button onclick="window.close();" name="btncancel" value="CLOSE">&nbsp;&nbsp;<input type=submit name="btnsave" id="DELETE_UPAH" value="DELETE"></td>
											<%End if%>
									</tr>
							</table>
						</div>

						<!-- view bkm bahan-->
						<div class="TabbedPanelsContent">
							<!-- ============================================================================================================================= -->
							<div id="datadetail">
								<table align=center border="0" width="850" class='data' id='table1'>
									<tr id="trheader1">
										
										<th>Activity</th>
										<th>TC</th>
										<th>Item Barang</th>
										<th>Satuan</th>
										<th>Qty</th>
										<th colspan="2">&nbsp;</th>
									</tr>

									<%
									
										dim sqlbarang,rsbarang 
										sqlbarang = "select b.BKMNO,b.ACTCODE,b.ACTCODE || '-' ||  a.fcname as activityname,b.ITEMCODE,s.FCCODE || '-' || s.FCNAME ITEMBARANG, " & _
										" s.UNITOFMEASUREMENT,b.LOCATIONCODE,NVL(b.qty,0)qty " & _
										" from BKMBAHANDTL b " & _
										" inner join STOCK s on (b.ITEMCODE=s.FCCODE AND b.fcba=s.fcba) " & _
										" inner join ACTMASTER a on (b.ACTCODE=a.FCCODE) where b.BKMNO='" & bkmno & "' order by b.ITEMCODE asc"
										 set rsbarang = connStr.Execute(sqlbarang)
										 'response.write sqlbarang
									%>

									<% 
										dim nobr
										nobr = 0
										While Not rsbarang.EOF 
									%>
										<tr id='trb<%=nobr%>'>
											
											
											<td class="ratakiri">
												<select class="lebaractbarang ratakiri" onchange="getTCItem(<%=nobr%>)" name="activityitem<%=nobr%>" id="activityitem<%=nobr%>">
													<option value=<%=rsbarang("ACTCODE")%>><%=rsbarang("activityname")%></option>
												</select>
											</td>
											<td class="ratatengah">
												<select class="lebarabsen ratakiri" name="tcitem<%=nobr%>" id="tcitem<%=nobr%>">
													<option value=<%=rsbarang("LOCATIONCODE")%>><%=rsbarang("LOCATIONCODE")%></option>
												</select>
											</td>
											<td class="ratakiri">
												<select class="lebaritembarang " name="ITEMCODE<%=nobr%>" id="ITEMCODE<%=nobr%>">
													<option value=<%=rsbarang("ITEMCODE")%>><%=rsbarang("ITEMBARANG")%></option>
												</select>
											</td>
											<td class="ratatengah">
												<input class="ratatengah lebarabsen" type="text" name="uom<%=nobr%>" id="uom<%=nobr%>" value="<%=rsbarang("UNITOFMEASUREMENT")%>"/>
											</td>
											<td class="ratatengah">
												<input class="ratatengah lebarabsen" type="text" name="qtyitem<%=nobr%>" id="qtyitem<%=nobr%>" data-value<%=nobr%>="<%=rsbarang("qty")%>" onkeyup="validasi(<%=nobr%>)" value="<%=Replace(rsbarang("qty"), ",", ".")%>"/>
											</td>


											<td>
												<img src='css/login/images/add.jpg' onclick="addrowBahan('<%=nobr%>')" />
											</td>
											<td>
												<% if nobr >= 0 then %>
													<img src='css/login/images/del.png' onclick="deleterowBahan('<%=nobr%>')" />
												<% else 
													response.write "&nbsp;"
													
												end if %>
											</td>
											<td></td>
										</tr>

											
										
									<%
										nobr = nobr + 1
										rsbarang.MoveNext
										Wend
									
									%>

									<tr class="ratatengah">
										<td  colspan="11"><input type=Submit name="btnsave" id="SAVE_BAHAN" value="SAVE BAHAN"></td>
										
									</tr>
								</table>
							</div>
							
						</div>

						
					</div>
				</div>
			</td>
		</tr>
		
	</table>
	<br>
	<table align=center border="0" class='data ratakiri' width='1100' id='tbldtl'>
		<tr>
			<th>Employee</th>
			<th>Lokasi</th>
			<th>No.RKB</th>
			<th>Absen</th>
			<th>Activity</th>
			<th>TC</th>
			<th>HK</th>
			<th>Unit</th>
			<th>Uom</th>
			<th>Output</th>
			<th>Rate</th>
			<th>Allowed Code</th>
			<th>Rate Allowed</th>
			<th>OVTH</th>
			<th>Amount</th>
			<th>Remark</th>
		</tr>
		<%
			'get data detail upah
			dim rsupah,allcode,rateallow,hkview, dtlamt
			sqldtlupah = "SELECT u.BKMNO,u.chargeunit fcba, u.EMPLOYEECODE,e.fcname || '-' ||  e.fccode as fcname, u.ACTCODE,a.fcname || '-' ||  u.ACTCODE as activityname, NVL(u.MANDAYS,0)MANDAYS, " & _
			" p.ALLOWDEDCODE, nvl(u.ALLOWANCE_AMT,0)ALLOWANCE_AMT, u.ATTENDANCE,nvl(u.OVERTIME_HOURS,0)OVERTIME_HOURS,u.TYPE_OVERTIME,o.fcname type_ovname," & _ 
			" to_char(u.IN_FIN,'hh24:mi:ss')IN_FINDEPAN,to_char(u.OUT_FIN,'hh24:mi:ss')OUT_FINDEPAN, " & _
			" to_char(u.IN_FIN,'dd/mm/yyyy hh24:mi:ss')IN_FIN,to_char(u.OUT_FIN,'dd/mm/yyyy hh24:mi:ss')OUT_FIN,  " & _
			" u.LOCATIONCODE, nvl(u.OUTPUT,0)OUTPUT,nvl(u.rate,0)rate, nvl(u.UNIT,0)unit,p.ALLOWDEDCODE,al.FCCODE || ' - ' || al.fcname as ALLOWDEDNAME,nvl(p.AMOUNT,0)AMOUNT,u.remarks,a.uom_unit,u.rkbno " & _
			" FROM BKMUPAHDTL u " & _ 
			" INNER JOIN employee e ON (u.EMPLOYEECODE=e.fccode and u.fcba=e.fcba) " & _
			" INNER JOIN JOBMASTER a ON (u.ACTCODE=a.FCCODE) " & _
			" LEFT JOIN BKMPROSES  p ON (u.EMPLOYEECODE=p.EMPLOYEECODE AND p.BKMNO = u.BKMNO AND p.LINENOKEY = u.LINENOKEY AND u.fcba=p.fcba) " & _ 
			" LEFT JOIN ALLOWDED al ON (al.FCCODE = p.ALLOWDEDCODE and al.fcba=p.fcba) " & _ 
			" LEFT JOIN overtime_convertion o ON (o.type=u.TYPE_OVERTIME and o.HOURS=u.OVERTIME_HOURS and o.fcba=u.fcba)" & _
			" where u.BKMNO='" & bkmno & "' order by e.fcname, a.fcname,u.LOCATIONCODE,p.ALLOWDEDCODE"
			set rsupah = connStr.Execute(sqldtlupah)
			'response.write sqldtlupah
			totalamount = 0
			While Not rsupah.EOF

			allcode = rsupah("ALLOWDEDCODE")

			if allcode = "A01" THEN 
				rateallow = 0
				hkview    = rsupah("MANDAYS")
			ELSE 
				rateallow = rsupah("ALLOWANCE_AMT")
				hkview    = 0
			END IF 

			

		%>
		<tr>
			<td class="ratakiri"><%=rsupah("fcname")%></td>
			<td class="ratatengah"><%=rsupah("fcba")%></td>
			<td class="ratatengah"><%=rsupah("rkbno")%></td>
			<td class="ratatengah"><%=rsupah("ATTENDANCE")%></td>
			<td class="ratakiri"><%=rsupah("activityname")%></td>
			<td class="ratatengah"><%=rsupah("LOCATIONCODE")%></td>
			<td class="ratatengah"><%=hkview%></td>
			<td class="ratatengah"><%=rsupah("UNIT")%></td>
			<td class="ratatengah"><%=rsupah("uom_unit")%></td>
			<td class="ratatengah"><%=rsupah("OUTPUT")%></td>
			<td class="ratatengah"><%=rsupah("RATE")%></td>
			<td class="ratakiri"><%=rsupah("ALLOWDEDNAME")%></td>
			<td class="ratakanan"><%=formatnumber(rateallow,0)%></td>
			<td class="ratatengah"><%=rsupah("OVERTIME_HOURS")%></td>
			<td class="ratakanan"><%=formatnumber(rsupah("AMOUNT"),0)%></td>
			<td class="ratakanan"><%=rsupah("remarks")%></td>
		</tr>
		<%
			totalamount = cdbl(totalamount) + cdbl(rsupah("AMOUNT"))
			rsupah.MoveNext
			Wend
		%>
		<tr>
			<td colspan="12" class="ratakanan tebal">Total</td>
			<td class="ratakanan tebal"><%=formatnumber(totalamount,0)%></td>
		</tr>
	</table>

</form>
</div>
</body>
<!--#include file="bottom.jsp"-->
</html>
<link href="/js/select2/select2.min.css" rel="stylesheet" type="text/css"/>
<script src="/js/select2/select2.min.js" type="text/javascript"></script>
<script>
document.getElementById("SAVE_BAHAN").addEventListener("click", function(event) {
    // Ambil semua input qtyitem
    const qtyInputs = document.querySelectorAll('input[name^="qtyitem"]');
    let valid = true;

    qtyInputs.forEach((input, index) => {
        // Ambil stok maksimum dari atribut data-value
        const dataValueAttr = input.getAttribute(`data-value${index}`);
        if (!dataValueAttr) return; // skip jika tidak ada atributnya

        // Ubah format koma ke titik dan konversi ke float
        const maxQty = parseFloat(dataValueAttr.replace(",", "."));
        const currentQty = parseFloat(input.value.replace(",", "."));

        // Cek apakah qty melebihi stok
        if (currentQty > maxQty) {
            alert(`Qty pada baris ${index + 1} melebihi stok! (Stok: ${maxQty}, Input: ${currentQty})`);
            input.focus();
            valid = false;
        }
    });

    // Jika ada kesalahan, batalkan submit
    if (!valid) {
        event.preventDefault();
    }
});
</script>
<script type="text/javascript">
	new Spry.Widget.TabbedPanels("tp1", {
		defaultTab: 0
	});

	function hanyaAngka(evt) {
		evt.value = evt.value.replace(/[^0-9,.-]/g,"");
	}

$(function() {	
	
	//$("#tglbkm").datepicker();
	showSupervisi1();
	showSupervisi2();
	showSupervisi3();
	showSupervisi4();
	//showSection();
	
	getLoad();
	getLoaditem();

	var jumldataitem = $('#jumldataitem').val();

	if (jumldataitem=="0"){
		addrowBahan(1);
	}
	get_Allowentry();
	//alert(jumldataitem);
});



function numformat(val) {
	var nStr = val.toString();
	nStr = nStr.replace(/\,/g, "");
	var x = nStr.split('.');
	var x1 = x[0];
	var x2 = x.length > 1 ? '.' + x[1] : '';
	var rgx = /(\d+)(\d{3})/;
	while (rgx.test(x1)) {
		x1 = x1.replace(rgx, '$1' + '.' + '$2');
	}
	return x1 + x2;
}

//get lokasi upah
function showFCBA(selid) {
	var fcba = 'fcba';
	$('#lokasi').empty();
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getabsen.asp?viewdata=' + fcba,
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});
}

//get section
function showSection() {

	$('#sectionname').select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getsection.asp',
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
				var query = { keyword: params.term }
				return query;
			}
		}
	});
}

//get GANGCODE 
function showGang() {
	var sectionname = $("#sectionname").val();
	$('#gangcode').empty();
	$('#gangcode').select2({
		ajax: {
			delay: 1000,
			url : 'bkm_gangcode.asp?sectionname=' + sectionname,
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});
}


//get supervisi 1
function showSupervisi1() {
	
	$('#supervisi1').select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getsupervisi.asp',
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});
}

//get supervisi 2
function showSupervisi2() {
	
	$('#supervisi2').select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getsupervisi.asp',
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});
}

//get supervisi 3
function showSupervisi3() {
	
	$('#supervisi3').select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getsupervisi.asp',
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});
}

//get supervisi 1
function showSupervisi4() {
	
	$('#supervisi4').select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getsupervisi.asp',
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});
}

//get absen
function showAbsen(selid) {
	
	$('#absen').empty();
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getabsen.asp',
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});
}
//get RKB Lokasi upah
function getNoRKBFCBA(selid) {
	var tglbkm = $('#tglbkm').val();
	var lokasi = $('#lokasi' + selid).val();
	$('#rkbnofcba' + selid).empty();
	$('#rkbnofcba' + selid).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getrkb.asp?tglbkm=' + tglbkm + '&fcba=' + lokasi + '&id=' + 1+'&doc='+501,
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});
}

//get activity
function showActivity(selid) {
	var rkbno = $('#rkbnofcba' + selid).val();
	$('#activity').empty();
	$('#activity' + selid).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getactivity.asp?panen='+9+'&rkbno='+rkbno,
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});
}

//get tc
function showTC(selid,id) {
	var rkbno    = $('#rkbnofcba' + id).val();
	var activity = $('#activity' + id).val();
	var lokasi   = $('#lokasi' + id).val();
	var tglbkm   = $('#tglbkm').val();
	$('#tcupah').empty();
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			//url : 'bkm_getblok.asp',
			url : 'bkm_getblok.asp?blok='+9+'&activity=' + activity +'&lokasi=' + lokasi +'&rkbno=' + rkbno ,
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});
}

//tcitem
function getTCItem(i)
{
	var rkbno = $('#rkbno').val();
	var tglbkm = $('#tglbkm').val();
	var activityitem = $('#activityitem' + i).val();
	var lokasi = '<%= Session("fcba") %>';

	$('#tcitem' + i).empty();
	$('#tcitem' + i).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getblok.asp?blok='+9+'&activity=' + activityitem +'&lokasi=' + lokasi +'&rkbno=' + rkbno ,
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});

}


//get allowedcode
function showAllowed(selid) {
	
	$('#allowedcode').empty();
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getallowed.asp',
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});
}


//get employee
function showEmployee(selid) {
	var gangcode = $("#gangcode").val();
	$('#fccode').empty();
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			url : "bkm_employee_edit.asp?gangcode=" + gangcode,
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});
}

//getLoadUpah
function getLoad() {
		
	var bkmno = $("#bkmno").val();
	var upahedit = 2;
	$.post("bkm_employee_load.asp?upahedit=" + upahedit +"&bkmno=" + bkmno, function(data) {
		
		var json  = JSON.parse(data);
		var jumlview  = json.results.length;
		for (var i = 0; i < jumlview; i++) {

			showEmployee("fccode" + i);
			showFCBA("lokasi" + i);
			showAbsen("absen" + i);
			showActivity(i);
			showTC("tcupah" + i,i);
			showAllowed("allowedcode" + i);
			showTypeOv("typeovrtime" + i);
			
		}
	});
	
}


var jumlinv = $('#jumldata').val();
var StartIteration =  parseInt(jumlinv) + 1	

function addrowUpah(trid) {
	var fcba = '<%= Session("fcba") %>';
	var rkbno = $('#rkbno').val();
	var lastrow = this.StartIteration++;
	

	if (lastrow < 2) {
		var row = $("#trheader2").after("<tr id='tru" + lastrow + "'></tr>");
		var kurangbutton = "&nbsp;";
	} else {
		var row = $('#tru' + trid).after("<tr id='tru" + lastrow + "'></tr>");
		var kurangbutton = "<img src='css/login/images/del.png' onclick='deleterowUpah(" + lastrow + ")' />";
	}

	var fccode = "<select name='fccode" + lastrow + "' id='fccode" + lastrow + "' class='ratakiri lebaractivity' onchange='getFinger(" + lastrow + ");'></select>";
	$('#tru' + lastrow).append("<td class='ratakiri'>" + fccode + "</td>");

	var timein = "<span name='timein" + lastrow + "' id='timein" + lastrow + "' class='ratatengah lebarrate'></span><input type='hidden' name='timeinx" + lastrow + "' id='timeinx" + lastrow + "'>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + timein + "</td>");

	var timeout = "<span name='timeout" + lastrow + "' id='timeout" + lastrow + "' class='ratatengah lebarrate'></span><input type='hidden' name='timeoutx" + lastrow + "' id='timeoutx" + lastrow + "'>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + timeout + "</td>");

	var lokasi = "<select name='lokasi" + lastrow + "' id='lokasi" + lastrow + "' class='ratakiri lebarabsen' onchange='getNoRKBFCBA(" + lastrow + ");'><option value='" + fcba + "'>" + fcba + "</option></select>";
	$('#tru' + lastrow).append("<td class='ratatengah lebarabsen'>" + lokasi + "</td>");

	var rkbnofcba = "<select name='rkbnofcba" + lastrow + "' id='rkbnofcba" + lastrow + "' class='ratakiri lebaractivity' onchange='showActivity(" + lastrow + ");'><option value='" + rkbno + "'>" + rkbno + "</option></select>";
	$('#tru' + lastrow).append("<td class='ratakiri lebaractivity'>" + rkbnofcba + "</td>");


	var absen = "<select name='absen" + lastrow + "' id='absen" + lastrow + "' class='ratakiri lebarabsen' onchange='getDisable(" + lastrow + ");'></select>";
	$('#tru' + lastrow).append("<td class='ratatengah lebarabsen'>" + absen + "</td>");

	var activity = "<select name='activity" + lastrow + "' id='activity" + lastrow + "' class='ratakiri lebaractivity' onchange='getTC(" + lastrow + ");getuom(" + lastrow + ");getDisable(" + lastrow + ");'></select>";
	$('#tru' + lastrow).append("<td class='ratakiri lebaractivity'>" + activity + "</td>");

	var blok = "<select type='text' name='tcupah" + lastrow + "' id='tcupah" + lastrow + "' value='0' class='ratatengah lebartc'></select>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + blok + "</td>");

	var hk = "<input type='text' name='hk" + lastrow + "' id='hk" + lastrow + "' onkeyup='getHKhinew(" + lastrow + ");getValidasiHK(" + lastrow + ");hanyaAngka(this)' value='0' class='ratatengah lebarhk'>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + hk + "</td>");

	var unit = "<input type='text' name='unit" + lastrow + "' id='unit" + lastrow + "' onkeyup='hanyaAngka(this)' value='0' class='ratatengah lebarabsen'>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + unit + "</td>");

	var uom_unit = "<input type='text' name='uom_unit" + lastrow + "' id='uom_unit" + lastrow + "'  value='HA' readonly class='ratatengah lebarabsen'>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + uom_unit + "</td>");

	var output = "<input type='text' name='output" + lastrow + "' id='output" + lastrow + "' onkeyup='hanyaAngka(this)' value='0' class='ratatengah lebarabsen'>";
	$('#tru' + lastrow).append("<td class='ratakiri'>" + output + "</td>");

	var ratekb = "<input type='text' name='ratekb" + lastrow + "' id='ratekb" + lastrow + "' onkeyup='hanyaAngka(this)' value='0' class='ratatengah lebarrate'>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + ratekb + "</td>");

	var allowedcode = "<select name='allowedcode" + lastrow + "' id='allowedcode" + lastrow + "' class='ratakiri lebarallow'></select>";
	$('#tru' + lastrow).append("<td class='ratakiri'>" + allowedcode + "</td>");

	var ratepremi = "<input type='text' name='ratepremi" + lastrow + "' id='ratepremi" + lastrow + "' onkeyup='hanyaAngka(this)' value='0' class='ratakanan lebarrate'>";
	$('#tru' + lastrow).append("<td class='ratakiri'>" + ratepremi + "</td>");

	var ovtimehours = "<input type='text' name='ovtimehours" + lastrow + "' id='ovtimehours" + lastrow + "' onkeyup='hanyaAngka(this);getOTHConvertion(" + lastrow + ")' value='0' class='ratatengah lebarabsen'>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + ovtimehours + "</td>");

	var typeovrtime = "<select type='text' name='typeovrtime" + lastrow + "' id='typeovrtime" + lastrow + "' onchange='getOTHConvertion(" + lastrow + ");' value='0' class='ratatengah lebarovtime'></select><input type='hidden' name='othconvertion" + lastrow + "' id='othconvertion" + lastrow + "' onkeyup='hanyaAngka(this)' value='0' class='ratatengah'>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + typeovrtime + "</td>");

	var remark = "<textarea type='text' name='remark" + lastrow + "' id='remark" + lastrow + "' rows='3' cols='10' ></textarea>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + remark + "</td>");
	
	var tambahbutton = "<img src='css/login/images/add.jpg' onclick='addrowUpah(" + lastrow + ")' />";
	$('#tru' + lastrow).append("<td>" + tambahbutton + "</td>");

	$('#tru' + lastrow).append("<td>" + kurangbutton + "</td>");

	$('#jumldata').val(lastrow);

	showEmployee("fccode" + lastrow);
	showFCBA("lokasi" + lastrow);
	showAbsen("absen" + lastrow);
	showActivity(lastrow);
	showTC("tc" + lastrow,lastrow);
	showAllowed("allowedcode" + lastrow);
	showTypeOv("typeovrtime" + lastrow);

	new Spry.Widget.ValidationSelect("fccode" + lastrow);
	new Spry.Widget.ValidationSelect("absen" + lastrow);
	new Spry.Widget.ValidationSelect("activity" + lastrow);
	new Spry.Widget.ValidationSelect("tcupah" + lastrow);
	new Spry.Widget.ValidationTextField("hk" + lastrow);

	
}

function getHKhi(i)
{
	var emp 	= $('#fccode' + i).val();
	var tglbkm 	= $('#tglbkm').val();
	var hk 		= $('#hk' + i).val();
	var hkhidden 		= $('#hkhidden' + i).val();

	if (emp != "") {
		$.post("bkm_employee.asp?emp=" + emp + "&employee=" + 300 + "&tglbkm=" + tglbkm, function(data) {
			
			var json  = JSON.parse(data);
			var hkhi  = json.results[0].mandayshi;
			var ttlhkhi = parseFloat(hkhi) - parseFloat(hkhidden);
			var ttlhk	  = parseFloat(ttlhkhi)+parseFloat(hk)

			//alert(ttlhk);
			if((ttlhk)>1){
				alert("TOTAL HK PEKERJA INI SUDAH LEBIH DARI 1 HK, SILAHKAN CEK KEMBALI INPUTAN ANDA..");
				$('#hk' + i).val("0");
				
				 
			}
						
		});
		
	}

	if(hk>1){
		alert("INPUTAN HK TIDAK BOLEH LEBIH DARI 1 HK");
		$('#hk' + i).val(0);
	}
}

function getHKhinew(i)
{
	var emp 	= $('#fccode' + i).val();
	var tglbkm 	= $('#tglbkm').val();
	var hk 		= $('#hk' + i).val();

	if (emp != "") {
		$.post("bkm_employee.asp?emp=" + emp + "&employee=" + 300 + "&tglbkm=" + tglbkm, function(data) {
			$('#hk' + i).empty();
			var json  = JSON.parse(data);
			var hkhi  = json.results[0].mandayshi;
			var ttlhk = parseFloat(hkhi)+parseFloat(hk);
			
			if((ttlhk)>1){
				alert("TOTAL HK PEKERJA INI SUDAH LEBIH DARI 1 HK, SILAHKAN CEK KEMBALI INPUTAN ANDA..");
				 $('#hk' + i).val("0");
				//  $('#fccode' + i).empty();
				//  $('#lokasi' + i).empty();
				//  $('#absen' + i).empty();
				//  $('#activity' + i).empty();
				//  $('#tcupah' + i).empty();
				//  $('#timein' + i).empty();
				//  $('#timeout' + i).empty();
				//  $('#timeinx' + i).empty();
				//  $('#timeoutx' + i).empty();
				 
			}
						
		});
		
	}

	if(hk>1){
		alert("INPUTAN HK TIDAK BOLEH LEBIH DARI 1 HK");
		$('#hk' + i).val(0);
	}
}

function getValidasiHK(i)
{
	var tglbkm = $('#tglbkm').val();
	var thnblntgl  = tglbkm.split('/');
	var blnview 	= thnblntgl[0];
	var tglview 	= thnblntgl[1];
	var thnview 	= thnblntgl[2];
	var thnblntglview = thnview + blnview + tglview;

	var statusemp = $('#statusemp' + i).val();
	var ttlhkkhl = $('#ttlhkkhl' + i).val();
	var hk = $('#hk' + i).val();
	if(statusemp=="KHL" & parseFloat(ttlhkkhl)>=20  & thnblntglview>20230126){
		alert("KARYAWAN DENGAN STATUS KHL TIDAK BOLEH MELEBIHI DARI 20 HK");
		$('#hk' + i).val(0);
	}else if(hk>1){
		alert("INPUTAN HK TIDAK BOLEH LEBIH DARI 1 HK");
		$('#hk' + i).val(0);
	}
	
}

function deleterowUpah(i) {
	Spry.Widget.Utils.destroyWidgets('fccode' + i);
	Spry.Widget.Utils.destroyWidgets('lokasi' + i);
	Spry.Widget.Utils.destroyWidgets('rkbnofcba' + i);	
	Spry.Widget.Utils.destroyWidgets('absen' + i);
	Spry.Widget.Utils.destroyWidgets('activity' + i);
	Spry.Widget.Utils.destroyWidgets('tcupah' + i);
	Spry.Widget.Utils.destroyWidgets('hk' + i);
	Spry.Widget.Utils.destroyWidgets('timein' + i);
	Spry.Widget.Utils.destroyWidgets('timeout' + i);
	Spry.Widget.Utils.destroyWidgets('timeinx' + i);
	Spry.Widget.Utils.destroyWidgets('timeoutx' + i);
	Spry.Widget.Utils.destroyWidgets('unit' + i);
	Spry.Widget.Utils.destroyWidgets('output' + i);
	Spry.Widget.Utils.destroyWidgets('ratekb' + i);
	Spry.Widget.Utils.destroyWidgets('allowedcode' + i);
	Spry.Widget.Utils.destroyWidgets('ratepremi' + i);
	Spry.Widget.Utils.destroyWidgets('ovtimehours' + i);
	Spry.Widget.Utils.destroyWidgets('typeovrtime' + i);
	Spry.Widget.Utils.destroyWidgets('remark' + i);
		
	 $('#tru' + i).remove();
	 //getdeljmlRow();
}

//show time from finger employee
function getFinger(i)
{
	var tglbkm = $('#tglbkm').val();
	var fccode = $('#fccode' + i).val();
	//alert(tglbkm);
	if (fccode != "") {
		$.post("bkm_employee_edit.asp?emp=" + 1 + "&tglbkm=" + tglbkm + "&fccode=" + fccode, function(data) {
			var json  = JSON.parse(data);
			//var jumlview  = json.results.length;
			var timein = json.results[0].timein;
			var timeout = json.results[0].timeout;
				$('#timein' + i).html(timein);
				$('#timeout' + i).html(timeout);
			
			
		});
		
	}

}

//getLoaditem
function getLoaditem() {
	
	var bkmno = $("#bkmno").val();
	var jmlitem = 1;
	$.post("bkm_show_item.asp?jmlitem=" + jmlitem +"&bkmno=" + bkmno, function(data) {
		var json  = JSON.parse(data);
		var jumlview  = json.results.length;
		
		for (var i = 0; i < jumlview; i++) {

			//showTC("tc" + i);
			showItem("ITEMCODE" + i);
			showActivityitem("activityitem" + i);
			
		}
	});
	
	
}


//get item
function showItem(selid) {
	var tglbkm = $('#tglbkm').val();
	var section = $('#sectionname').val();
	$('#ITEMCODE').empty();
	var jmlitem = 0;
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			url : "bkm_show_item.asp?jmlitem=" + jmlitem + "&tglbkm=" + tglbkm + "&section=" + section,
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
				var query = { keyword: params.term }
				return query;
			}
		}
	});
}


//get activity item
function showActivityitem(selid) {
	var rkbno = $('#rkbno').val();
	$('#activityitem').empty();
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getactivity.asp?panen='+9+'&rkbno='+rkbno,
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});
}


//var jumlitem = $('#jumldataitem').val();
//var StartIterationitem =  jumlitem + 1	
var jumlitem = $('#jumldataitem').val();
if (jumlitem=="0"){
		var StartIterationitem = 1
}else{
	var StartIterationitem =  parseInt(jumlitem) + 1
}

function addrowBahan(trid)
{

	var lstitem = this.StartIterationitem++;
	if (lstitem < 2) {
		var row = $("#trheader1").after("<tr id='trb" + lstitem + "'></tr>");
		var kurangbutton = "&nbsp;";
	} else {
		var row = $('#trb' + trid).after("<tr id='trb" + lstitem + "'></tr>");
		var kurangbutton = "<img src='css/login/images/del.png' onclick='deleterowBahan(" + lstitem + ")' />";
	}

	var activityitem = "<select name='activityitem" + lstitem + "' id='activityitem" + lstitem + "' class='lebaractbarang' onchange='getTCItem(" + lstitem + ")'></select>";
	$('#trb' + lstitem).append("<td class='ratakiri'>" + activityitem + "</td>");

	var tcitem = "<select name='tcitem" + lstitem + "' id='tcitem" + lstitem + "' class='lebarabsen ratatengah'></select>";
	$('#trb' + lstitem).append("<td class='ratakiri'>" + tcitem + "</td>");

	var ITEMCODE = "<select name='ITEMCODE" + lstitem + "' id='ITEMCODE" + lstitem + "' class='lebaritembarang'></select>";
	$('#trb' + lstitem).append("<td class='ratakiri'>" + ITEMCODE + "</td>");

	var uom = "<input type='text' name='uom" + lstitem + "' id='uom" + lstitem + "' value='' class='lebarabsen ratatengah'>";
	$('#trb' + lstitem).append("<td class='ratakiri'>" + uom + "</td>");

	var qtyitem = "<input type='text' name='qtyitem" + lstitem + "' id='qtyitem" + lstitem + "' value='' onkeyup='hanyaAngka(this)' class='lebarabsen ratatengah'>";
	$('#trb' + lstitem).append("<td class='ratakiri'>" + qtyitem + "</td>");
	
	var tambahbutton = "<img src='css/login/images/add.jpg' onclick='addrowBahan(" + lstitem + ")' />";
	$('#trb' + lstitem).append("<td>" + tambahbutton + "</td>");

	$('#trb' + lstitem).append("<td>" + kurangbutton + "</td>");

	$('#jumldataitem').val(lstitem);

	showItem("ITEMCODE" + lstitem, $('#select2-sectionname-container').attr('title'));
	$('#qtyitem'+lstitem).on('keyup', function() {
		validasi(lstitem)
	});
	showActivityitem("activityitem" + lstitem);

	//new Spry.Widget.ValidationSelect("activityitem" + lstitem);
	//new Spry.Widget.ValidationSelect("tcitem" + lstitem);
	//new Spry.Widget.ValidationSelect("ITEMCODE" + lstitem);
	//new Spry.Widget.ValidationTextField("qtyitem" + lstitem);	
}

function validasi(i){
	// console.log("YESSSS!", i);
	let elm = document.getElementById("#qtyitem"+i);
	let val = parseFloat($("#qtyitem"+i).val().replace(/,/g, ''));
		
	// ambil stok dari text
		
	// ambil stok dari text
	let text = $("#select2-ITEMCODE" + i + "-container").attr('title');
	let parts = text.split("STOK :");

	let stok = 0;
	if (parts.length > 1) {
		stok = parseFloat(parts[1].trim().replace(/,/g, ''));
	}
	// cek apakah input lebih besar dari stok
	if (val > stok) {
		alert("Jumlah melebihi stok (" + stok + ")");
		$("#qtyitem"+i).val(0); // optional: reset ke nilai stok
	}
}


function deleterowBahan(i) {
	Spry.Widget.Utils.destroyWidgets('ITEMCODE' + i);
	Spry.Widget.Utils.destroyWidgets('uom' + i);
	Spry.Widget.Utils.destroyWidgets('activityitem' + i);
	Spry.Widget.Utils.destroyWidgets('tcitem' + i);
	Spry.Widget.Utils.destroyWidgets('qtyitem' + i);
	 $('#trb' + i).remove();
}


//get tc
function getTC(i)
{
	var rkbno    = $('#rkbnofcba' + i).val();
	var activity = $('#activity' + i).val();
	var lokasi   = $('#lokasi' + i).val();

	$('#tcupah' + i).empty();
	$('#tcupah' + i).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getblok.asp?blok='+9+'&activity=' + activity +'&lokasi=' + lokasi +'&rkbno=' + rkbno ,
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});

}

//ambil data satuan
function getuom(i)
{
	var activity = $('#activity' + i).val();
	if (activity != "") {
		$.post("bkm_getactivity.asp?activity=" + activity +"&panen=" + 10, function(data) {
			$('#uom_unit' + i).empty();
			var json  = JSON.parse(data);
			var uom = json.results[0].uom_unit
			$('#uom_unit' + i).val(uom);
						
		});
		
	}
}

//function disabled input()
function getDisable(i)
{
	var absen = $('#absen' + i).val();
	if (absen=="KB"){
		//disbale
		$( '#hk' + i ).prop( "disabled", true );
		$( '#bunch' + i ).prop( "disabled", true );
		$( '#abw' + i ).prop( "disabled", true );
		$( '#kgbrondolan' + i ).prop( "disabled", true );
		//$( '#allowedcode' + i ).prop( "disabled", true );
		
		//enable
		
		
	}else if (absen == "KH"){
	    $( '#hk' + i ).prop( "disabled", true );
		$( '#hk' + i ).val(0);
	}else{
		$( '#hk' + i ).prop( "disabled", false );
		$( '#bunch' + i ).prop( "disabled", false );
		$( '#abw' + i ).prop( "disabled", false );
		$( '#kgbrondolan' + i ).prop( "disabled", false );
		$( '#unit' + i ).prop( "disabled", false );
		$( '#output' + i ).prop( "disabled", false );
		$( '#ratekb' + i ).prop( "disabled", false );
		$( '#allowedcode' + i ).prop( "disabled", false );
		
	}
	
}

//show tipe overtime
function showTypeOv(selid)
{
	
	//$('#typeovrtime' + i).empty();
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_viewdata.asp?idview=' + 7,
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});

}

//show time from finger employee
function getFinger(i)
{
	var tglbkm = $('#tglbkm').val();
	var fccode = $('#fccode' + i).val();
	//alert(tglbkm);
	if (fccode != "") {
		$.post("bkm_employee_edit.asp?emp=" + 1 + "&tglbkm=" + tglbkm + "&fccode=" + fccode, function(data) {
			var json  = JSON.parse(data);
			//var jumlview  = json.results.length;
			var timein = json.results[0].timein;
			var timeinx = json.results[0].timeinx;
			var timeout = json.results[0].timeout;
			var timeoutx = json.results[0].timeoutx;
				$('#timein' + i).html(timein);
				$('#timeout' + i).html(timeout);
				$('#timeinx' + i).val(timeinx);
				$('#timeoutx' + i).val(timeoutx);
			
			
		});
		
	}

}

//view data oth convertion
function getOTHConvertion(i)
{
	var ovtimehours = $("#ovtimehours" + i).val();
	var typeovrtime = $('#typeovrtime' + i).val();
	if (ovtimehours != "") {
		$.post("bkm_employee.asp?ovtimehours=" + ovtimehours + "&typeovrtime=" +  typeovrtime +  "&satuan=" + 7, function(data) {
			$('#othconvertion' + i).empty();
			var json  = JSON.parse(data);
			var othconvertion = json.results[0].convertion
			$('#othconvertion' + i).val(othconvertion);
						
		});
		
	}

	
}

function getIsibelakang(i)
{
	var ratepremidepan = $('#ratepremidepan' + i).val();
	if(ratepremidepan==""){
		$('#ratepremi' + i).val(0);
	}else{

		$('#ratepremi' + i).val(ratepremidepan);
	}
}

function get_Allowentry() {
    var tglbkm = $('#tglbkm').val();
	var statusbkm = $("#statusbkm").val();

    if (tglbkm != "") {
        $.post("bkm_allow_entry.asp?tglbkm=" + tglbkm, function(data) {
            $('#bodytable').empty();
            var json = JSON.parse(data);
            var closingStatus = json.results[0].msg;

            if (closingStatus == "FALSE" && statusbkm == "Planned") {
                $('#SAVE_UPAH').show();
                $('#DELETE_UPAH').show();
                $('#SAVE_BAHAN').show();
                $('#msg').html('<label color:green>OPEN PERIODE </label>');

            } else if (closingStatus == "TRUE" || closingStatus == "FALSE" && statusbkm == "Review Accounting") {
                // alert('PERIODE SUDAH DI BUKA !!!');
                $('#SAVE_UPAH').show();
                $('#DELETE_UPAH').show();
                $('#SAVE_BAHAN').show();
                $('#msg').html('<label color:green>OPEN PERIODE </label>');

            } else {
                alert('PERIODE SUDAH DI CLOSE !!!');
                $('#msg').html('<label color:red>PERIODE SUDAH DI CLOSE !!!</label>');
                $('#SAVE_UPAH').hide();
                $('#DELETE_UPAH').hide();
                $('#SAVE_BAHAN').hide();
            }
        });
    }
}

</script>
