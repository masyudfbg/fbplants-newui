<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->

<%
dim businessunit,fcbaname,CompanyName,q,sqltotdata,bkmno 
bkmno		= trim(request("bkmno"))
businessunit= trim(session("fcba"))

Dim statusbkm
statusbkm = Request.QueryString("status")

'response.write bkmno

q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & businessunit & "' "
	' response.write q
set rsCat = connStr.Execute(q)
if not rsCat.eof then
   fcbaname = rsCat("fcName")
   CompanyName=rsCat("FCCOMPANYNAME")
 
end if


' get data header
dim sqlheader,rsHeader,bkmnohdr,DATECREATED,SECTIONNAME,GANGCODE,rkbno
dim ROWSTATEID,SUPERVISI1,SUPERVISI2,SUPERVISI3,SUPERVISI4,SUPERVISINAME1
dim SUPERVISINAME2,SUPERVISINAME3,SUPERVISINAME4,Clasification,Clasificationname,ROWSTATE,POSTINGFLG,HUJAN,fcbahdr
sqlheader = "select h.BKMNO,to_char(h.BKMDATE,'mm/dd/yyyy')BKMDATE,h.GANGCODE,h.SUPERVISI1,(SELECT s1.FCNAME FROM EMPLOYEE  s1 WHERE h.SUPERVISI1=s1.FCCODE and s1.fcba=h.fcba ) AS SUPERVISINAME1, " & _
" h.SUPERVISI2,(SELECT s2.FCNAME FROM EMPLOYEE s2 WHERE h.SUPERVISI2=s2.FCCODE and s2.fcba=h.fcba) AS SUPERVISINAME2,h.SUPERVISI3, " & _
" (SELECT s3.FCNAME FROM EMPLOYEE s3 WHERE h.SUPERVISI3=s3.FCCODE and s3.fcba=h.fcba) AS SUPERVISINAME3,h.SUPERVISI4, " & _
" (SELECT s4.FCNAME FROM EMPLOYEE s4 WHERE h.SUPERVISI4=s4.FCCODE and s4.fcba=h.fcba) AS SUPERVISINAME4,h.SECTIONCODE, " & _ 
" h.DOCUMENT_CLASSIFICATION,c.fcname, h.ROWSTATE,h.POSTINGFLG,h.HUJAN,h.FCBA,h.rkbno " & _
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
	POSTINGFLG = rsHeader("POSTINGFLG")
	ROWSTATE = rsHeader("ROWSTATE")
	HUJAN = rsHeader("HUJAN")
	fcbahdr = rsHeader("fcba")
	rkbno = rsHeader("rkbno")
end if

IF POSTINGFLG = "0" THEN 
   POSTINGFLG = "TIDAK"
ELSE
   POSTINGFLG	= "YA"
END IF

'count detailupah
Dim SQL_Rows, NoOfRows,RSCount

SQL_Rows = "SELECT Count (*) from BKMUPAHDTL u " & _ 
" where u.BKMNO='" & bkmno & "'"
set RSCount = connStr.Execute(SQL_Rows)
if not RSCount.eof then 
	NoOfRows = RSCount(0).value
end if
RSCount.close

'count detailbahan
Dim SQL_Rowsitem, NoOfRowsitem,RSCountitem

SQL_Rowsitem = "SELECT Count (*) from BKMMUTU where BKMNO='" & bkmno & "'"
set RSCountitem = connStr.Execute(SQL_Rowsitem)
if not RSCountitem.eof then 
	NoOfRowsitem = RSCountitem(0).value
end if
RSCountitem.close

'count detaildenda
Dim SQL_Rowsdenda, NoOfRowsdenda,RSCountdenda

SQL_Rowsdenda = "SELECT Count (*) from BKMDENDA where BKMNO='" & bkmno & "'"
set RSCountdenda = connStr.Execute(SQL_Rowsdenda)
if not RSCountdenda.eof then 
	NoOfRowsdenda = RSCountdenda(0).value
end if
RSCountdenda.close




%>

<html>
	<head>
		<title>FBPlants - BKM Operasional Panen</title>
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
	.lebardenda{
		width: 60px;
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
		<strong class="posisikiri">&nbsp;&nbsp;PT. <%=CompanyName%> - &nbsp;<%=FcbaName%>   -           MODULE : BKM OPERASIONAL PANEN EDIT</strong><br>					
</div>
<br>
<br>
<div class="box_bothall" style="overflow:scroll;">			
<form name="frmEntry" method="post" action="bkm_panen_prosesedit.asp">
	<input type="hidden" name="jumldata" id="jumldata" value="<%=NoOfRows%>" />
	<input type="hidden" name="jumldataitem" id="jumldataitem" value="<%=NoOfRowsitem%>" />
	<input type="hidden" name="jumldatadenda" id="jumldatadenda" value="<%=NoOfRowsdenda%>" />
	<input type="hidden" name="fcbahdr" id="fcbahdr" value="<%=fcbahdr%>" />
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
				<input type="text" name="gangcode" style= readonly id="gangcode" size="15" value="<%=GANGCODE%>">
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
					<option value="<%=SUPERVISI2%>"><%=SUPERVISINAME2%></option>
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
			<td align="left"><%=POSTINGFLG%>-<%=ROWSTATE%> <input type="hidden" value="<%=ROWSTATE%>" id="statusbkm"/></td>
			<td align="left">&nbsp;</td>
			<td align="left"><span id="msg" >-</span></td>
			<td align="left" >Hujan</td>
			<td align="left">
				<%
				if HUJAN = "1" THEN %>
					<input type="checkbox" name="hujan" id="hujan" checked >
					</input>

				<%else%>
					<input type="checkbox" name="hujan" id="hujan" >
					</input>
				<%end if%>
			</td>
		</tr>
	</table>
	
	<br>
		
	<table align=center border="0" class='data posisikiri' width='900' id='tableall'>
		<tr>
			<td>
				<div class="TabbedPanels" id="tp1" align='center'>
				
					<ul class="TabbedPanelsTabGroup">
						<li class="TabbedPanelsTab">BKM UPAH</li>
						<li class="TabbedPanelsTab">BKM MUTU</li>
						<li class="TabbedPanelsTab">BKM DENDA</li>
						
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
									<th>No. RKB</th>
									<th>Absen</th>
									<th>Activity</th>
									<th>TC</th>
									<th>HK</th>
									<th>JJG</th>
									<th>JJG UN</th>
									<th>JJG BA</th>
									<th>BJR</th>
									<th>Brdln</th>
									<th>KG TBS</th>
									<th>Output</th>
									<th>Rate</th>
									<th>Allowed Code</th>
									<th>Rate Allowed</th>
									<th>Remark</th>
									
									<th colspan="2">&nbsp;</th>
								</tr>

								<% 
								
								
								'get data detail upah
								dim sqldtlupah,rsdtlupah,no,ttlhk,ttlbunch,ttlbunchun,ttlbunchba,ttlbjr,ttlbrndln,ttlkgtbs,ttloutput,ttlrate,ttlrateallowen, totalamount
								sqldtlupah = "SELECT u.BKMNO,u.chargeunit fcba, u.EMPLOYEECODE,e.fcname || '-' ||  e.fccode as fcname, u.ACTCODE,a.fcname || '-' ||  u.ACTCODE as activityname, NVL(u.MANDAYS,0)MANDAYS, " & _
								" u.ALLOWANCE_CODE, nvl(u.ALLOWANCE_AMT,0)ALLOWANCE_AMT, u.ATTENDANCE,u.OVERTIME_HOURS,u.TYPE_OVERTIME,o.fcname type_ovname, " & _ 
								" to_char(u.IN_FIN,'hh24:mi:ss')IN_FINDEPAN,to_char(u.OUT_FIN,'hh24:mi:ss')OUT_FINDEPAN,	" & _
								" to_char(u.IN_FIN,'dd/mm/yyyy hh24:mi:ss')IN_FIN,to_char(u.OUT_FIN,'dd/mm/yyyy hh24:mi:ss')OUT_FIN,  " & _
								" u.LOCATIONCODE, NVL(u.OUTPUT,0)OUTPUT,NVL(u.rate,0)rate, NVL(u.UNIT,0)UNIT ,u.ALLOWANCE_CODE, " & _ 
								" al.FCCODE || ' - ' || al.fcname as ALLOWDEDNAME,NVL(u.bunch_weight,0)bunch_weight, " & _ 
								" NVL(u.bunchun,0)bunchun,NVL(u.bunchba,0)bunchba,NVL(u.abw,0)abw,nvl(u.kg_brondolan,0)kg_brondolan,u.remarks,u.rkbno " & _
								" FROM BKMUPAHDTL u " & _ 
								" INNER JOIN employee e ON (u.EMPLOYEECODE=e.fccode and u.fcba=e.fcba) " & _
								" INNER JOIN ACTMASTER a ON (u.ACTCODE=a.FCCODE) " & _
								" LEFT JOIN ALLOWDED al ON (al.FCCODE = u.ALLOWANCE_CODE and al.fcba=u.fcba) " & _ 
								" LEFT JOIN overtime_convertion o ON (o.type=u.TYPE_OVERTIME and o.HOURS=u.OVERTIME_HOURS and o.fcba=u.fcba)" & _
								" where u.BKMNO='" & bkmno & "' and u.EMPLOYEECODE not in " & _ 
								" (select h1.SUPERVISI1 from BKMHDR h1 where h1.bkmno=u.bkmno AND h1.SUPERVISI1=e.fccode union all  select h2.SUPERVISI2 from BKMHDR h2 where h2.bkmno=u.bkmno AND h2.SUPERVISI2=e.fccode " & _ 
								" union all  select h3.SUPERVISI3 from BKMHDR h3 where h3.bkmno=u.bkmno  AND h3.SUPERVISI3=e.fccode " & _ 
								" union all  select h4.SUPERVISI2 from BKMHDR h4 where h4.bkmno=u.bkmno  AND h4.SUPERVISI4=e.fccode) " & _ 
								" order by e.fcname, a.fcname,u.LOCATIONCODE,u.ALLOWANCE_CODE"
								set rsdtlupah = connStr.Execute(sqldtlupah)
								'response.write sqldtlupah
								

																
								no = 0
								totalamount = 0
								ttlhk = 0
								ttlbunch = 0 
								ttlbunchun = 0
								ttlbunchba = 0 
								ttlbjr = 0 
								ttlbrndln = 0 
								ttlkgtbs = 0 
								ttloutput = 0 
								ttlrate = 0
								ttlrateallowen = 0
								While Not rsdtlupah.EOF %>
								<tr id="tru<%=no%>">
									
									<td class="ratakiri">
										<select class="lebaractivity " name="fccode<%=no%>" id="fccode<%=no%>" onchange="getFinger(<%=no%>)">
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
										<input type="text" class="lebaractivity" name="rkbnofcba<%=no%>" id="rkbnofcba<%=no%>" value="<%=rsdtlupah("rkbno")%>">	
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
									<td class="ratatengah">
										<input class="ratatengah lebarhk" type="text" name="hk<%=no%>" id="hk<%=no%>" value="<%=Replace(rsdtlupah("MANDAYS"), ",", ".")%>"/>
									</td>
									<td class="ratatengah">
										<input class="ratatengah lebarabsen" type="text" name="bunch<%=no%>" id="bunch<%=no%>" onkeyup="getKGTBS(<%=no%>)" value="<%=Replace(rsdtlupah("unit"), ",", ".")%>"/>
									</td>
									<td class="ratatengah">
										<input class="ratatengah lebarabsen" type="text" name="jjgun<%=no%>" id="jjgun<%=no%>" value="<%=Replace(rsdtlupah("bunchun"), ",", ".")%>"/>
									</td>
									<td class="ratatengah">
										<input class="ratatengah lebarabsen" type="text" name="jjgba<%=no%>" id="jjgba<%=no%>" value="<%=Replace(rsdtlupah("bunchba"), ",", ".")%>"/>
									</td>
									<td class="ratatengah">
										<input class="ratatengah lebarabsen" type="text" name="abw<%=no%>" id="abw<%=no%>" onkeyup="getKGTBS(<%=no%>)" value="<%=Replace(rsdtlupah("abw"), ",", ".")%>"/>
									</td>
									<td class="ratatengah">
										<input class="ratatengah lebarabsen" type="text" name="kgbrondolan<%=no%>" id="kgbrondolan<%=no%>" value="<%=Replace(rsdtlupah("kg_brondolan"), ",", ".")%>"/>
									</td>
									<td class="ratatengah">
										<input class="ratatengah lebarabsen" type="text" name="kgtbs<%=no%>" id="kgtbs<%=no%>" value="<%=Replace(rsdtlupah("bunch_weight"), ",", ".")%>"/>
									</td>
									
									<td class="ratatengah">
										<input class="ratatengah lebarabsen" type="text" name="output<%=no%>" id="output<%=no%>" value="<%=Replace(rsdtlupah("OUTPUT"), ",", ".")%>"/>
									</td>
									
									<td class="ratatengah">
										<input class="ratatengah lebarrate" type="text" name="ratekb<%=no%>" id="ratekb<%=no%>" value="<%=Replace(rsdtlupah("rate"), ",", ".")%>"/>
									</td>
									<td class="ratakiri">
										<select class="lebarallow ratakiri" name="allowedcode<%=no%>" id="allowedcode<%=no%>">
											<option value=<%=rsdtlupah("ALLOWANCE_CODE")%>><%=rsdtlupah("ALLOWDEDNAME")%></option>
										</select>
									</td>
																		
									<td class="ratatengah">
										<input class="ratakanan lebarrate" type="text"   name="ratepremi<%=no%>" id="ratepremi<%=no%>" value="<%=Replace(rsdtlupah("ALLOWANCE_AMT"), ",", ".")%>"/>
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
									no = no +1
									ttlhk = cdbl(ttlhk) + cdbl(rsdtlupah("MANDAYS"))
									ttlbunch = cdbl(ttlbunch) + cdbl(rsdtlupah("unit"))
									ttlbunchun = cdbl(ttlbunchun) + cdbl(rsdtlupah("bunchun"))
									ttlbunchba = cdbl(ttlbunchba) + cdbl(rsdtlupah("bunchba"))
									ttlbjr = cdbl(ttlbjr) + cdbl(rsdtlupah("abw"))
									ttlbrndln = cdbl(ttlbrndln) + cdbl(rsdtlupah("kg_brondolan"))
									ttlkgtbs = cdbl(ttlkgtbs) + cdbl(rsdtlupah("bunch_weight"))
									ttloutput = cdbl(ttloutput) + cdbl(rsdtlupah("output"))
									ttlrate = cdbl(ttlrate) + cdbl(rsdtlupah("rate"))
									ttlrateallowen = cdbl(ttlrateallowen) + cdbl(rsdtlupah("ALLOWANCE_AMT"))
									
									rsdtlupah.MoveNext
									Wend
								%>
									<tr class="ratakanan">
										<td colspan="8" class="tebal">Total</td>
										<td class="tebal ratatengah"><input class="ratatengah lebarhk" type="text" readonly  style="background-color:#eee;" name="amount" id="amount<%=no%>" value="<%=FormatNumber(ttlhk,0)%>"/></td>
										<td class="tebal ratakanan"><input class="ratatengah lebarhk" type="text" readonly  style="background-color:#eee;" name="amount" id="amount<%=no%>" value="<%=FormatNumber(ttlbunch,0)%>"/></td>
										<td class="tebal  ratakanan"><input class="ratatengah lebarhk" type="text" readonly  style="background-color:#eee;" name="amount" id="amount<%=no%>" value="<%=FormatNumber(ttlbunchun,0)%>"/></td>
										<td class="tebal ratatengah"><input class="ratatengah lebarhk" type="text" readonly  style="background-color:#eee;" name="amount" id="amount<%=no%>" value="<%=FormatNumber(ttlbunchba,0)%>"/></td>
										<td class="tebal ratatengah"><input class="ratatengah lebarhk" type="text" readonly  style="background-color:#eee;" name="amount" id="amount<%=no%>" value="<%=FormatNumber(ttlbjr,0)%>"/></td>
										<td class="tebal ratatengah"><input class="ratatengah lebarhk" type="text" readonly  style="background-color:#eee;" name="amount" id="amount<%=no%>" value="<%=FormatNumber(ttlbrndln,0)%>"/></td>
										<td class="tebal ratatengah"><input class="ratatengah lebarhk" type="text" readonly  style="background-color:#eee;" name="amount" id="amount<%=no%>" value="<%=FormatNumber(ttlkgtbs,0)%>"/></td>
										<td class="tebal ratakiri"><input class="ratatengah lebarhk" type="text" readonly  style="background-color:#eee;" name="amount" id="amount<%=no%>" value="<%=FormatNumber(ttloutput,2)%>"/></td>
										<td class="tebal ratatengah"><input class="ratatengah lebarrate" type="text" readonly  style="background-color:#eee;" name="amount" id="amount<%=no%>" value="<%=FormatNumber(ttlrate,0)%>"/></td>
										<td></td>
										<td class="tebal ratatengah"><input class="ratakanan lebarrate" type="text" readonly  style="background-color:#eee;" name="amount" id="amount<%=no%>" value="<%=FormatNumber(ttlrateallowen,0)%>"/></td>
										
									</tr>
							</table>
						</div>

						<!-- view bkm MUTU-->
						<div class="TabbedPanelsContent">
							<!-- ============================================================================================================================= -->
							<div id="datadetail">
								<table align=center border="0" width="750" class='data' id='table1'>
									<tr id="trheader1">
										<th>Blok</th>
										<th>B. Mateng</th>
										<th>B. LMateng</th>
										<th>B. KMateng</th>
										<th>B. Mentah</th>
										<th>B. Busuk</th>
										
										<th colspan="2">&nbsp;</th>
									</tr>

									<%
									
										dim sqlbarang,rsbarang 
										sqlbarang = "select * " & _
										" from BKMMUTU " & _
										" where BKMNO='" & bkmno & "' order by blok asc"
										 set rsbarang = connStr.Execute(sqlbarang)
										 'response.write sqlbarang
									%>

									<% 
										dim nobr
										nobr = 0
										While Not rsbarang.EOF 
									%>
										<tr id='trb<%=nobr%>'>
											<td class="ratatengah">
												<select class="lebartc " name="tcmutu" id="tcmutu<%=nobr%>">
													<option value=<%=rsbarang("blok")%>><%=rsbarang("blok")%></option>
												</select>
											</td>
											<td class="ratatengah">
												<input class="ratatengah lebartc" type="text" name="bmateng" id="bmateng<%=nobr%>" value="<%=rsbarang("bmateng")%>"/>
											</td>
											<td class="ratatengah">
												<input class="ratatengah lebartc" type="text" name="blmateng" id="blmateng<%=nobr%>" value="<%=rsbarang("blmateng")%>"/>
											</td>
											<td class="ratatengah">
												<input class="ratatengah lebartc" type="text" name="bkmateng" id="bkmateng<%=nobr%>" value="<%=rsbarang("bkmateng")%>"/>
											</td>
											<td class="ratatengah">
												<input class="ratatengah lebartc" type="text" name="bmentah" id="bmentah<%=nobr%>" value="<%=rsbarang("bmentah")%>"/>
											</td>
											<td class="ratatengah">
												<input class="ratatengah lebartc" type="text" name="bbusuk" id="bbusuk<%=nobr%>" value="<%=rsbarang("bbusuk")%>"/>
											</td>
											

											<td>
												<img src='css/login/images/add.jpg' onclick="addrowMutu('<%=nobr%>')" />
											</td>
											<td>
												<% if nobr >= 0 then %>
													<img src='css/login/images/del.png' onclick="deleterowMutu('<%=nobr%>')" />
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
									
								</table>
							</div>
							
						</div>

						<!-- view bkm denda-->
						<div class="TabbedPanelsContent">
							<!-- ============================================================================================================================= -->
							<div id="detaildenda">
								<table align=center border="0" width="950" class='data' id='table3'>
									<tr >
										
											<th rowspan="2">Employee</th>
											<th colspan="7">Kualitas TBS</th>
											<th colspan="7">Kualitas Ancak</th>
											<th rowspan="2" colspan="2">&nbsp;</th>
									</tr>
									<tr id="trheader3">
										
										<th>(A)</th>
										<th>(BA)</th>
										<th>(TP)</th>
										<th>(BK)</th>
										<th>(O)</th>
										<th>(AB)</th>
										<th>(E)</th>

										<th>(BT)</th>
										<th>(PS)</th>
										<th>(SP)</th>
										<th>(BRD)</th>
										<th>(OP)</th>
										<th>(BM)</th>
										<th>(BDT)</th>
									</tr>

									<%
									
										dim sqldenda,rsdenda,nodenda 
										sqldenda = "Select B.EMPLOYEECODE,E.FCNAME, B.LINENOKEY, B.REMARKS, " & _
										" B.TBSA, B.TBSBA, B.TBSTP, B.TBSBK, B.TBSO, " & _
										" B.TBSAB, B.TBSE, B.ANCAKBT, B.ANCAKPS, B.ANCAKSP, " & _
										" B.ANCAKBRD, B.ANCAKOP, B.ANCAKBM, B.ANCAKBDT " & _
										" From   BKMDENDA B " & _
										" LEFT JOIN EMPLOYEE E ON (B.EMPLOYEECODE=E.FCCODE AND B.FCBA=E.FCBA) " & _
										" WHERE B.BKMNO='" & bkmno & "' order by e.fcname asc"
										 set rsdenda = connStr.Execute(sqldenda)
										 'response.write sqlbarang
									%>

									<% 
										 
										nodenda = 0
										While Not rsdenda.EOF 
									%>
										<tr id='trd<%=nodenda%>'>
											<td class="ratakiri">
												<select class="lebaractivity " name="empcode<%=nodenda%>" id="empcode<%=nodenda%>" onchange="getLineDenda('<%=nodenda%>')">
													<option value=<%=rsdenda("EMPLOYEECODE")%>><%=rsdenda("FCNAME")%></option>
												</select><input type="hidden" name="linedenda<%=nodenda%>" id="linedenda<%=nodenda%>" value="<%=rsdenda("LINENOKEY")%>"></input>
											</td>
											<td class="ratatengah">
												<input class="ratatengah lebardenda" type="text" name="ktbsa<%=nodenda%>" id="ktbsa<%=nodenda%>" value="<%=rsdenda("TBSA")%>"/>
											</td>
											<td class="ratatengah">
												<input class="ratatengah lebardenda" type="text" name="ktbsba<%=nodenda%>" id="ktbsba<%=nodenda%>" value="<%=rsdenda("TBSBA")%>"/>
											</td>
											<td class="ratatengah">
												<input class="ratatengah lebardenda" type="text" name="ktbstp<%=nodenda%>" id="ktbstp<%=nodenda%>" value="<%=rsdenda("TBSTP")%>"/>
											</td>
											<td class="ratatengah">
												<input class="ratatengah lebardenda" type="text" name="ktbsbk<%=nodenda%>" id="ktbsbk<%=nodenda%>" value="<%=rsdenda("TBSBK")%>"/>
											</td>
											<td class="ratatengah">
												<input class="ratatengah lebardenda" type="text" name="ktbso<%=nodenda%>" id="ktbso<%=nodenda%>" value="<%=rsdenda("TBSO")%>"/>
											</td>
											<td class="ratatengah">
												<input class="ratatengah lebardenda" type="text" name="ktbsab<%=nodenda%>" id="ktbsab<%=nodenda%>" value="<%=rsdenda("TBSAB")%>"/>
											</td>
											<td class="ratatengah">
												<input class="ratatengah lebardenda" type="text" name="ktbse<%=nodenda%>" id="ktbse<%=nodenda%>" value="<%=rsdenda("TBSE")%>"/>
											</td>
											<td class="ratatengah">
												<input class="ratatengah lebardenda" type="text" name="kancakbt<%=nodenda%>" id="kancakbt<%=nodenda%>" value="<%=rsdenda("ANCAKBT")%>"/>
											</td>
											<td class="ratatengah">
												<input class="ratatengah lebardenda" type="text" name="kancakps<%=nodenda%>" id="kancakps<%=nodenda%>" value="<%=rsdenda("ANCAKPS")%>"/>
											</td>
											<td class="ratatengah">
												<input class="ratatengah lebardenda" type="text" name="kancaksp<%=nodenda%>" id="kancaksp<%=nodenda%>" value="<%=rsdenda("ANCAKSP")%>"/>
											</td>
											<td class="ratatengah">
												<input class="ratatengah lebardenda" type="text" name="kancakbrd<%=nodenda%>" id="kancakbrd<%=nodenda%>" value="<%=rsdenda("ANCAKBRD")%>"/>
											</td>
											<td class="ratatengah">
												<input class="ratatengah lebardenda" type="text" name="kancakop<%=nodenda%>" id="kancakop<%=nodenda%>" value="<%=rsdenda("ANCAKOP")%>"/>
											</td>
											<td class="ratatengah">
												<input class="ratatengah lebardenda" type="text" name="kancakbm<%=nodenda%>" id="kancakbm<%=nodenda%>" value="<%=rsdenda("ANCAKBM")%>"/>
											</td>
											<td class="ratatengah">
												<input class="ratatengah lebardenda" type="text" name="kancakbdt<%=nodenda%>" id="kancakbdt<%=nodenda%>" value="<%=rsdenda("ANCAKBDT")%>"/>
											</td>
											

											<td>
												<img src='css/login/images/add.jpg' onclick="addrowDenda('<%=nodenda%>')" />
											</td>
											<td>
												<% if nodenda >= 0 then %>
													<img src='css/login/images/del.png' onclick="deleterowDenda('<%=nodenda%>')" />
												<% else 
													response.write "&nbsp;"
												end if %>
											</td>
											<td></td>
										</tr>

											
										
									<%
										nodenda = nodenda + 1
										rsdenda.MoveNext
										Wend
									
									%>
									
								</table>
							</div> 
							<!-- ============================================================================================================================= -->
						</div>

						
					</div>
				</div>
			</td>
		</tr>
		<tr class="ratatengah">
			<% if POSTINGFLG = "TIDAK" then%>
			<td  colspan="11"><input type=Submit name="btnsave" id="SAVE_UPAH" value="SAVE">&nbsp;&nbsp;<input type=button onclick="window.close();" name="btncancel" value="CLOSE">&nbsp;&nbsp;<input type=submit name="btnsave" id="DELETE_UPAH" value="DELETE"></td>
			<%End if%>
		</tr>
	</table>


	<br>
	<table align=center border="0" class='data ratakiri' width='1100' id='tbldtl'>
		<tr>
			<th>Employee</th>
			<th>Lokasi</th>
			<th>No. RKB</th>
			<th>Absen</th>
			<th>Activity</th>
			<th>TC</th>
			<th>HK</th>
			<th>JJG</th>
			<th>Brondln</th>
			<th>KGTbs</th>
			<th>Output</th>
			<th>Rate</th>
			<th>Allowed Code</th>
			<th>Rate Allowed</th>
			<th>Amount</th>
			<th>Remark</th>
		</tr>
		<%
			'get data detail upah
			dim sqlupah, rsupah,allcode,rateallow,hkview, dtlamt
			sqlupah = "SELECT u.BKMNO,u.chargeunit fcba, u.EMPLOYEECODE,e.fcname || '-' ||  e.fccode as fcname, u.ACTCODE,a.fcname || '-' ||  u.ACTCODE as activityname, NVL(u.MANDAYS,0)MANDAYS, " & _
			" p.ALLOWDEDCODE, nvl(u.ALLOWANCE_AMT,0)ALLOWANCE_AMT, u.ATTENDANCE,u.OVERTIME_HOURS,u.TYPE_OVERTIME," & _ 
			" nvl(u.unit,0) unit,nvl(u.kg_brondolan,0)kg_brondolan,NVL(u.bunch_weight,0)bunch_weight, " & _
			" u.LOCATIONCODE, nvl(u.OUTPUT,0)output,u.rate,p.ALLOWDEDCODE,al.FCCODE || ' - ' || al.fcname as ALLOWDEDNAME,nvl(p.AMOUNT,0)AMOUNT,u.remarks,u.rkbno " & _
			" FROM BKMUPAHDTL u " & _ 
			" INNER JOIN employee e ON (u.EMPLOYEECODE=e.fccode and u.fcba=e.fcba) " & _
			" INNER JOIN ACTMASTER a ON (u.ACTCODE=a.FCCODE) " & _
			" LEFT JOIN BKMPROSES  p ON (u.EMPLOYEECODE=p.EMPLOYEECODE AND p.BKMNO = u.BKMNO AND p.LINENOKEY = u.LINENOKEY AND u.fcba=p.fcba) " & _ 
			" LEFT JOIN ALLOWDED al ON (al.FCCODE = p.ALLOWDEDCODE and al.fcba=p.fcba) " & _ 
			" where u.BKMNO='" & bkmno & "' order by e.fcname, a.fcname,u.LOCATIONCODE,p.ALLOWDEDCODE"
			set rsupah = connStr.Execute(sqlupah)
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
			<td class="ratatengah"><%=Replace(hkview, ",", ".")%></td>
			<td class="ratatengah"><%=Replace(rsupah("UNIT"), ",", ".")%></td>
			<td class="ratatengah"><%=Replace(rsupah("kg_brondolan"), ",", ".")%></td>
			<td class="ratatengah"><%=Replace(rsupah("bunch_weight"), ",", ".")%></td>
			<td class="ratatengah"><%=Replace(rsupah("OUTPUT"), ",", ".")%></td>
			<td class="ratatengah"><%=Replace(rsupah("RATE"), ",", ".")%></td>
			<td class="ratakiri"><%=rsupah("ALLOWDEDNAME")%></td>
			<td class="ratakanan"><%=Replace(formatnumber(rateallow,0), ",", ".")%></td>
			<td class="ratakanan"><%=Replace(formatnumber(rsupah("AMOUNT"),0), ",", ".")%></td>
			<td class="ratakanan"><%=rsupah("remarks")%></td>
		</tr>
		<%
			totalamount = cdbl(totalamount) + cdbl(rsupah("AMOUNT"))
			rsupah.MoveNext
			Wend
		%>
		<tr>
			<td colspan="14" class="ratakanan tebal">Total</td>
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

<script type="text/javascript">
	new Spry.Widget.TabbedPanels("tp1", {
		defaultTab: 0
	});

$(function() {	
	//$("#tglbkm").datepicker();
	showSupervisi1();
	showSupervisi2();
	showSupervisi3();
	showSupervisi4();

	getLoad();
	getLoaditem();
	getLoaDenda();
	
	var jumldataitem = $('#jumldataitem').val();

	if (jumldataitem=="0"){
		addrowMutu(1);
		//alert(jumldataitem);
	}

	var jumldatadendax = $('#jumldatadenda').val();

	if (jumldatadendax=="0"){
		addrowDenda(1);
		
	}
	get_Allowentry();
	
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
					var query = {
									keyword: params.term
								}
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

//GET RKB
function getNoRKB()
{
	
	var rkbno = $('#rkbno').val();
	$('#rkbno').empty();
	$('#rkbno').select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getrkb.asp?rkbno=' + rkbno + '&id=' + 3+'&doc='+503,
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
			url : 'bkm_getrkb.asp?tglbkm=' + tglbkm + '&fcba=' + lokasi + '&id=' + 1+'&doc='+503,
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
	var rkbno = $('#rkbno').val();
	$('#activity').empty();
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
		//	url : 'bkm_getactivity.asp?panen=' + 1,
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

function showEmployeedenda(selid) {
	var tglbkm = $('#tglbkm').val();
	var bkmno = $('#bkmno').val();
	//$('#fccode').empty();
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			url : "bkm_employee_edit.asp?emp=" + 4 + "&tglbkm=" + tglbkm + "&bkmno=" + bkmno,
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

function showBlokmutu(selid)
{
	
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getblok.asp?blok=' + 1,
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
			showActivity("activity" + i);
			//getNoRKB("activity" + i);
			showTC("tcupah" + i,i);
			showAllowed("allowedcode" + i);
			showTypeOv("typeovrtime" + i);
			
		}
	});
	
	
}

var jumlinv = $('#jumldata').val();
var StartIteration =  parseInt(jumlinv) + 1	

function addrowUpah(trid,nourut) {	
	var fcba = '<%= Session("fcba") %>';
	var rkbno = $('#rkbno').val();
	 var lastrow	= this.StartIteration++;
	// alert(lastrow);
	if (lastrow < 2) {
		var row = $("#trheader2").after("<tr id='tru" + lastrow + "'></tr>");
		var kurangbutton = "&nbsp;";
	} else {
		var row = $('#tru' + trid).after("<tr id='tru" + lastrow + "'></tr>");
		var kurangbutton = "<img src='css/login/images/del.png' onclick='deleterowUpah(" + lastrow + ");' />";
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

	var activity = "<select name='activity" + lastrow + "' id='activity" + lastrow + "' class='ratakiri lebaractivity' onchange='getTC(" + lastrow + ");'></select>";
	$('#tru' + lastrow).append("<td class='ratakiri'>" + activity + "</td>");

	var blok = "<select type='text' name='tcupah" + lastrow + "' id='tcupah" + lastrow + "' value='0' class='ratatengah lebartc'></select>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + blok + "</td>");

	var hk = "<input type='text' name='hk" + lastrow + "' id='hk" + lastrow + "' onkeyup='hanyaAngka(this)' value='0' class='ratatengah lebarhk'>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + hk + "</td>");

	var bunch = "<input type='text' name='bunch" + lastrow + "' id='bunch" + lastrow + "' onkeyup='getKGTBS(" + lastrow + ");hanyaAngka(this)' value='0' class='ratatengah lebarabsen'>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + bunch + "</td>");

	var jjgu = "<input type='text' name='jjgun" + lastrow + "' id='jjgun" + lastrow + "' onkeyup='hanyaAngka(this)' value='0' class='ratatengah lebarabsen'>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + jjgu + "</td>");

	var jjgab = "<input type='text' name='jjgab" + lastrow + "' id='jjgab" + lastrow + "' onkeyup='hanyaAngka(this)' value='0' class='ratatengah lebarabsen'>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + jjgab + "</td>");

	var abw = "<input type='text' name='abw" + lastrow + "' id='abw" + lastrow + "' onkeyup='getKGTBS(" + lastrow + ");hanyaAngka(this)' value='0' class='ratatengah lebarabsen'>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + abw + "</td>");

	var kgbrondolan = "<input type='text' name='kgbrondolan" + lastrow + "' id='kgbrondolan" + lastrow + "' onkeyup='hanyaAngka(this)' value='0' class='ratatengah lebarabsen'>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + kgbrondolan + "</td>");

	var kgtbs = "<input type='text' name='kgtbs" + lastrow + "' id='kgtbs" + lastrow + "' onkeyup='hanyaAngka(this)' value='0' class='ratatengah lebarabsen'>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + kgtbs + "</td>");

	var output = "<input type='text' name='output" + lastrow + "' id='output" + lastrow + "' onkeyup='hanyaAngka(this)' value='0' class='ratatengah lebarabsen'>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + output + "</td>");

	var ratepremi = "<input type='text' name='ratekb" + lastrow + "' id='ratekb" + lastrow + "' onkeyup='hanyaAngka(this)' value='0' class='ratatengah lebarrate'>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + ratepremi + "</td>");

	var allowedcode = "<select name='allowedcode" + lastrow + "' id='allowedcode" + lastrow + "' class='ratakiri lebarallow'></select>";
	$('#tru' + lastrow).append("<td class='ratakiri'>" + allowedcode + "</td>");

	var ratepremi = "<input type='text' name='ratepremi" + lastrow + "' id='ratepremi" + lastrow + "' onkeyup='hanyaAngka(this)' value='0' class='ratatengah lebarrate'>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + ratepremi + "</td>");

	var remark = "<textarea type='text' name='remark" + lastrow + "' id='remark" + lastrow + "' rows='3' cols='10' ></textarea>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + remark + "</td>");
	
	var tambahbutton = "<img src='css/login/images/add.jpg' onclick='addrowUpah(" + lastrow + "," + lastrow + ");' />";
	$('#tru' + lastrow).append("<td>" + tambahbutton + "</td>");

	$('#tru' + lastrow).append("<td>" + kurangbutton + "</td>");

	$('#jumldata').val(lastrow);
	

	showEmployee("fccode" + lastrow);
    showFCBA("lokasi" + lastrow);
	showAbsen("absen" + lastrow);
	showActivity("activity" + lastrow);
	//getNoRKB("activity" + lastrow);
	showAllowed("allowedcode" + lastrow);
	getNoRKBFCBA("rkbnofcba" + lastrow);

	new Spry.Widget.ValidationSelect("fccode" + lastrow);
	new Spry.Widget.ValidationSelect("absen" + lastrow);
	new Spry.Widget.ValidationSelect("activity" + lastrow);
	new Spry.Widget.ValidationSelect("tcupah" + lastrow);
	new Spry.Widget.ValidationTextField("bunch" + lastrow);
	new Spry.Widget.ValidationTextField("hk" + lastrow);


}

function deleterowUpah(i) {
	Spry.Widget.Utils.destroyWidgets('fccode' + i);
	// Spry.Widget.Utils.destroyWidgets('timein' + i);
	// Spry.Widget.Utils.destroyWidgets('timeout' + i);
	Spry.Widget.Utils.destroyWidgets('timeinx' + i);
	Spry.Widget.Utils.destroyWidgets('timeoutx' + i);
	Spry.Widget.Utils.destroyWidgets('lokasi' + i);
	Spry.Widget.Utils.destroyWidgets('absen' + i);
	Spry.Widget.Utils.destroyWidgets('activity' + i);
	Spry.Widget.Utils.destroyWidgets('tcupah' + i);
	Spry.Widget.Utils.destroyWidgets('hk' + i);
	Spry.Widget.Utils.destroyWidgets('bunch' + i);
	Spry.Widget.Utils.destroyWidgets('jjgun' + i);
	//Spry.Widget.Utils.destroyWidgets('jjgab' + i);
	Spry.Widget.Utils.destroyWidgets('abw' + i);
	Spry.Widget.Utils.destroyWidgets('kgbrondolan' + i);
	Spry.Widget.Utils.destroyWidgets('kgtbs' + i);
	Spry.Widget.Utils.destroyWidgets('output' + i);
	Spry.Widget.Utils.destroyWidgets('ratekb' + i);
	Spry.Widget.Utils.destroyWidgets('allowedcode' + i);
	Spry.Widget.Utils.destroyWidgets('ratekb' + i);
	Spry.Widget.Utils.destroyWidgets('ratepremi' + i);
	Spry.Widget.Utils.destroyWidgets('remark' + i);
	
	
	 $('#tru' + i).remove();
	
	
}



var jumlitem = $('#jumldataitem').val();
if (jumlitem=="0"){
		var StartIterationitem = 1
}else{
	var StartIterationitem =  parseInt(jumlitem) + 1
}

function addrowMutu(trid)
{

	var lstitem = this.StartIterationitem++;
	
	if (lstitem < 2) {
		var row = $("#trheader1").after("<tr id='trb" + lstitem + "'></tr>");
		var kurangbutton = "&nbsp;";
	} else {
		var row = $('#trb' + trid).after("<tr id='trb" + lstitem + "'></tr>");
		var kurangbutton = "<img src='css/login/images/del.png' onclick='deleterowMutu(" + lstitem + ");' />";
	}

	var tcmutu = "<select name='tcmutu" + lstitem + "' id='tcmutu" + lstitem + "' class='lebartc ratatengah'></select>";
	$('#trb' + lstitem).append("<td class='ratatengah'>" + tcmutu + "</td>");

	var bmateng = "<input type='text' name='bmateng" + lstitem + "' id='bmateng" + lstitem + "' value='0' onkeyup='hanyaAngka(this)' class='lebartc ratatengah'>";
	$('#trb' + lstitem).append("<td class='ratatengah'>" + bmateng + "</td>");

	var blmateng = "<input type='text' name='blmateng" + lstitem + "' id='blmateng" + lstitem + "' value='0' onkeyup='hanyaAngka(this)' class='lebartc ratatengah'>";
	$('#trb' + lstitem).append("<td class='ratatengah'>" + blmateng + "</td>");

	var bkmateng = "<input type='text' name='bkmateng" + lstitem + "' id='bkmateng" + lstitem + "' value='0' onkeyup='hanyaAngka(this)' class='lebartc ratatengah'>";
	$('#trb' + lstitem).append("<td class='ratatengah'>" + bkmateng + "</td>");

	var bmentah = "<input type='text' name='bmentah" + lstitem + "' id='bmentah" + lstitem + "' value='0' onkeyup='hanyaAngka(this)' class='lebartc ratatengah'>";
	$('#trb' + lstitem).append("<td class='ratatengah'>" + bmentah + "</td>");

	var bbusuk = "<input type='text' name='bbusuk" + lstitem + "' id='bbusuk" + lstitem + "' value='0' onkeyup='hanyaAngka(this)' class='lebartc ratatengah'>";
	$('#trb' + lstitem).append("<td class='ratatengah'>" + bbusuk + "</td>");

	var tambahbutton = "<img src='css/login/images/add.jpg' onclick='addrowMutu(" + lstitem + ");' />";
	$('#trb' + lstitem).append("<td>" + tambahbutton + "</td>");

	$('#trb' + lstitem).append("<td>" + kurangbutton + "</td>");

	$('#jumldataitem').val(lstitem);

	showBlokmutu("tcmutu" + lstitem);
	
	
}

//hapus row bahan
function deleterowMutu(i) {
	Spry.Widget.Utils.destroyWidgets('tcmutu' + i);
	Spry.Widget.Utils.destroyWidgets('bmateng' + i);
	Spry.Widget.Utils.destroyWidgets('blmateng' + i);
	Spry.Widget.Utils.destroyWidgets('bkmateng' + i);
	Spry.Widget.Utils.destroyWidgets('bmentah' + i);
	Spry.Widget.Utils.destroyWidgets('bbusuk' + i);
	 $('#trb' + i).remove();
	 
}

//function disabled input()
function getDisable(i)
{
	var absen = $('#absen' + i).val();
	if (absen=="KB" || absen == "KH" ){
		//disbale
		$( '#hk' + i ).prop( "disabled", true );
		$( '#hk' + i ).val(0);
	}else{
	    $( '#hk' + i ).prop( "disabled", false );
	}
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

//getLoaditem
function getLoaDenda() {
	
	var bkmno = $("#bkmno").val();
	var jmlitem = 1;
	$.post("bkm_employee_denda.asp?jmlitem=" + jmlitem +"&bkmno=" + bkmno, function(data) {
		var json  = JSON.parse(data);
		var jumlview  = json.results.length;
		
		for (var i = 0; i < jumlview; i++) {

			
			showEmployeedenda("empcode" + i);
			
		}
	});
	
	
}


//get item
function showItem(selid) {
	$('#ITEMCODE').empty();
	var jmlitem = 0;
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			url : "bkm_show_item.asp?jmlitem=" + jmlitem,
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


//get activity item
function showActivityitem(selid) {
	
	$('#activityitem').empty();
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getactivity.asp',
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
function getTC(i)
{
    var rkbno    = $('#rkbnofcba' + i).val();
	var activity = $('#activity' + i).val();
	var lokasi   = $('#lokasi' + i).val();
	var tglbkm   = $('#tglbkm').val();

	$('#tcupah' + i).empty();
	$('#tcupah' + i).select2({
		ajax: {
			delay: 1000,
			//url : 'bkm_getblok.asp?activity=' + activity +'&lokasi=' + lokasi + '&tglbkm=' + tglbkm,
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

//view data oth convertion test
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

//view data oth convertion
function getLineDenda(i)
{	
	var bkmno = $("#bkmno").val();
	var empcode = $("#empcode" + i).val();
	if (empcode != "") {
		$.post("bkm_employee_edit.asp?emp=" + 5 + "&empcode=" +  empcode + "&bkmno=" + bkmno, function(data) {
			$('#linedenda' + i).empty();
			var json  = JSON.parse(data);
			var linedenda = json.results[0].linedenda
			$('#linedenda' + i).val(linedenda);
						
		});
		
	}

	
}


//untuk input denda panen
var jumldenda = $('#jumldatadenda').val();
if (jumldenda=="0"){
		var StartIterationdenda = 1
}else{
	var StartIterationdenda =  parseInt(jumldenda) + 1
}
function addrowDenda(trid)
{

	var lsdenda = this.StartIterationdenda++;
	
	if  (lsdenda < 2) {
		var row = $("#trheader3").after("<tr id='trd" + lsdenda + "'></tr>");
		var kurangbutton = "<img src='css/login/images/del.png' onclick='deleterowDenda(" + lsdenda + ");' />";
	} else {
		var row = $('#trd' + trid).after("<tr id='trd" + lsdenda + "'></tr>");
		var kurangbutton = "<img src='css/login/images/del.png' onclick='deleterowDenda(" + lsdenda + ");' />";
	}

	var empcode = "<select name='empcode" + lsdenda + "' id='empcode" + lsdenda + "' onchange='getLineDenda(" + lsdenda + ")' class='lebaractivity ratatengah'></select><input type='hidden' name='linedenda" + lsdenda + "' id='linedenda" + lsdenda + "'></input>";
	$('#trd' + lsdenda).append("<td class='ratatengah'>" + empcode + "</td>");

	var ktbsa = "<input type='text' name='ktbsa" + lsdenda + "' id='ktbsa" + lsdenda + "' value='0' onkeyup='hanyaAngka(this)' class='lebardenda ratatengah'>";
	$('#trd' + lsdenda).append("<td class='ratatengah'>" + ktbsa + "</td>");

	var ktbsba = "<input type='text' name='ktbsba" + lsdenda + "' id='ktbsba" + lsdenda + "' value='0' onkeyup='hanyaAngka(this)' class='lebardenda ratatengah'>";
	$('#trd' + lsdenda).append("<td class='ratatengah'>" + ktbsba + "</td>");

	var ktbstp = "<input type='text' name='ktbstp" + lsdenda + "' id='ktbstp" + lsdenda + "' value='0' onkeyup='hanyaAngka(this)' class='lebardenda ratatengah'>";
	$('#trd' + lsdenda).append("<td class='ratatengah'>" + ktbstp + "</td>");

	var ktbsbk = "<input type='text' name='ktbsbk" + lsdenda + "' id='ktbsbk" + lsdenda + "' value='0' onkeyup='hanyaAngka(this)' class='lebardenda ratatengah'>";
	$('#trd' + lsdenda).append("<td class='ratatengah'>" + ktbsbk + "</td>");

	var ktbso = "<input type='text' name='ktbso" + lsdenda + "' id='ktbso" + lsdenda + "' value='0' onkeyup='hanyaAngka(this)' class='lebardenda ratatengah'>";
	$('#trd' + lsdenda).append("<td class='ratatengah'>" + ktbso + "</td>");

	var ktbsab = "<input type='text' name='ktbsab" + lsdenda + "' id='ktbsab" + lsdenda + "' value='0' onkeyup='hanyaAngka(this)' class='lebardenda ratatengah'>";
	$('#trd' + lsdenda).append("<td class='ratatengah'>" + ktbsab + "</td>");

	var ktbse = "<input type='text' name='ktbse" + lsdenda + "' id='ktbse" + lsdenda + "' value='0' onkeyup='hanyaAngka(this)' class='lebardenda ratatengah'>";
	$('#trd' + lsdenda).append("<td class='ratatengah'>" + ktbse + "</td>");

	var kancakbt = "<input type='text' name='kancakbt" + lsdenda + "' id='kancakbt" + lsdenda + "' value='0' onkeyup='hanyaAngka(this)' class='lebardenda ratatengah'>";
	$('#trd' + lsdenda).append("<td class='ratatengah'>" + kancakbt + "</td>");

	var kancakps = "<input type='text' name='kancakps" + lsdenda + "' id='kancakps" + lsdenda + "' value='0' onkeyup='hanyaAngka(this)' class='lebardenda ratatengah'>";
	$('#trd' + lsdenda).append("<td class='ratatengah'>" + kancakps + "</td>");

	var kancaksp = "<input type='text' name='kancaksp" + lsdenda + "' id='kancaksp" + lsdenda + "' value='0' onkeyup='hanyaAngka(this)' class='lebardenda ratatengah'>";
	$('#trd' + lsdenda).append("<td class='ratatengah'>" + kancaksp + "</td>");

	var kancakbrd = "<input type='text' name='kancakbrd" + lsdenda + "' id='kancakbrd" + lsdenda + "' value='0' onkeyup='hanyaAngka(this)' class='lebardenda ratatengah'>";
	$('#trd' + lsdenda).append("<td class='ratatengah'>" + kancakbrd + "</td>");

	var kancakop = "<input type='text' name='kancakop" + lsdenda + "' id='kancakop" + lsdenda + "' value='0' onkeyup='hanyaAngka(this)' class='lebardenda ratatengah'>";
	$('#trd' + lsdenda).append("<td class='ratatengah'>" + kancakop + "</td>");

	var kancakbm = "<input type='text' name='kancakbm" + lsdenda + "' id='kancakbm" + lsdenda + "' value='0' onkeyup='hanyaAngka(this)' class='lebardenda ratatengah'>";
	$('#trd' + lsdenda).append("<td class='ratatengah'>" + kancakbm + "</td>");

	var kancakbdt = "<input type='text' name='kancakbdt" + lsdenda + "' id='kancakbdt" + lsdenda + "' value='0' onkeyup='hanyaAngka(this)' class='lebardenda ratatengah'>";
	$('#trd' + lsdenda).append("<td class='ratatengah'>" + kancakbdt + "</td>");

	var tambahbutton = "<img src='css/login/images/add.jpg' onclick='addrowDenda(" + lsdenda + ");' />";
	$('#trd' + lsdenda).append("<td>" + tambahbutton + "</td>");

	$('#trd' + lsdenda).append("<td>" + kurangbutton + "</td>");

	$('#jumldatadenda').val(lsdenda);

	showEmployeedenda("empcode" + lsdenda);
	
	
}

//hapus row bahan
function deleterowDenda(i) {
	Spry.Widget.Utils.destroyWidgets('empcode' + i);
	Spry.Widget.Utils.destroyWidgets('ktbsa' + i);
	Spry.Widget.Utils.destroyWidgets('ktbsba' + i);
	Spry.Widget.Utils.destroyWidgets('ktbstp' + i);
	Spry.Widget.Utils.destroyWidgets('ktbsbk' + i);
	Spry.Widget.Utils.destroyWidgets('ktbso' + i);
	Spry.Widget.Utils.destroyWidgets('ktbsab' + i);
	Spry.Widget.Utils.destroyWidgets('ktbse' + i);
	Spry.Widget.Utils.destroyWidgets('kancakbt' + i);
	Spry.Widget.Utils.destroyWidgets('kancakps' + i);
	Spry.Widget.Utils.destroyWidgets('kancaksp' + i);
	Spry.Widget.Utils.destroyWidgets('kancakbrd' + i);
	Spry.Widget.Utils.destroyWidgets('kancakop' + i);
	Spry.Widget.Utils.destroyWidgets('kancakbm' + i);
	Spry.Widget.Utils.destroyWidgets('kancakbdt' + i);
	 $('#trd' + i).remove();
	
}


//hitung kgtbs
function getKGTBS(i)
{
	var jjg = parseFloat($('#bunch' + i).val());
	var abw = parseFloat($('#abw' + i).val());
	var kgtbs = jjg*abw;
	kgtbs     = Math.round(kgtbs);
	$('#kgtbs' + i).val(kgtbs)
}


function get_Allowentry()
{
	var tglbkm = $('#tglbkm').val();
	var statusbkm = $("#statusbkm").val();

	if (tglbkm != "") {
		$.post("bkm_allow_entry.asp?tglbkm=" + tglbkm, function(data) {
			$('#bodytable').empty();
			var json  = JSON.parse(data);
			var closingStatus = json.results[0].msg;
			
			
			if(closingStatus=="FALSE" || closingStatus=="TRUE" && statusbkm == "Planned"){
				$('#SAVE_UPAH').show();
				$('#DELETE_UPAH').show();
				$('#msg').html('<label color:green>OPEN PERIODE </label>')
				
				//window.open('bkm_operasional_edit.Asp?bkmno=' + bkmno);
			} else if (closingStatus == "TRUE" && statusbkm == "Review Accounting") {
				$('#SAVE_UPAH').show();
				$('#DELETE_UPAH').show();
				$('#msg').html('<label color:green>OPEN PERIODE </label>')
			}else{
				alert('PERIODE SUDAH DI CLOSE !!!');
				$('#msg').html('<label color:red>PERIODE SUDAH DI CLOSE !!!</label>')
				
				$('#SAVE_UPAH').hide();
				$('#DELETE_UPAH').hide();
				//break;
			}
			
			
						
		});
		
	} 
}

</script>
