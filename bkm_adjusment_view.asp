<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->

<%
dim btnApprove,bkmno1,sql,totalpremi,bkmdate
dim ScriptName
ScriptName	=	Request.ServerVariables("SCRIPT_NAME")


function Akses_User
Dim rsUsrPosting, strUsrPosting,Akses_Usr
    Akses_Usr  =  0
 	strUsrPosting =	" SELECT 1 result " & _
					" FROM bkmhdr hdr " & _
					" INNER JOIN authorizetemplate at ON (at.document_type = 'AA' AND at.document_classification = hdr.document_classification)" & _
					" INNER JOIN authorizestructure aus ON (aus.fcba = hdr.fcba AND aus.authorizetemplatecode = at.fccode AND aus.fcusercode = '" & Session("CodeName") & "')" & _
					" WHERE hdr.bkmno='" & trim(request("bkmno")) & "' " & _
					" AND hdr.rowstate = 'Planned'" & _
					" AND hdr.postingflg = 0" & _
					" AND at.fccode = '505'" & _
					" AND hdr.fcba='" & trim(session("fcba")) & "' " & _
					" AND authorizelevelcode = 900" & _
				    " AND hdr.sectioncode in(select SECTIONNAME  from user_section_tab where FCUSERCODE='" & Session("CodeName") & "' AND sectionname='" & SECTIONNAME & "') " 
 				
 	Set rsUsrPosting = Server.CreateObject("ADODB.Recordset")
 	rsUsrPosting.Open strUsrPosting, connstr, 3
	
 	If Not rsUsrPosting.EOF Then
 		Akses_Usr = cdbl(rsUsrPosting("result"))
 	Else
 		Akses_Usr = 0
 	End If
 	rsUsrPosting.Close
 	set rsUsrPosting = Nothing
 
	Akses_User = Akses_Usr
end function 

function Akses_Posting
Dim rsAvailable, strAvailable,Access_Post
    Access_Post  =  0
 	strAvailable =	" select authorizestructure_api.get_access_post('505','" & businessunit & "','" & Session("CodeName") & "') result" & _
 					" from dual"
		
 	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
 	rsAvailable.Open strAvailable, connstr, 3
	
 	If Not rsAvailable.EOF Then
 		Access_Post = cdbl(rsAvailable("result"))
 	Else
 		Access_Post = 0
 	End If
 	rsAvailable.Close
 	set rsAvailable = Nothing
 
	Akses_Posting = Access_Post
end function

dim businessunit,fcbaname,CompanyName,q,sqltotdata,bkmno 
bkmno		= trim(request("bkmno"))
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
dim sqlheader,rsHeader,bkmnohdr,DATECREATED,SECTIONNAME,GANGCODE,ROWSTATEID,SUPERVISI1,SUPERVISI2,SUPERVISI3,SUPERVISI4,SUPERVISINAME1,SUPERVISINAME2,SUPERVISINAME3,SUPERVISINAME4,Clasification,Clasificationname,ROWSTATE,POSTINGFLG
sqlheader = "select h.BKMNO,to_char(h.BKMDATE,'MM/DD/YYYY') BKMDATE,h.GANGCODE,h.SUPERVISI1,(SELECT s1.FCNAME FROM EMPLOYEE  s1 WHERE h.SUPERVISI1=s1.FCCODE and s1.fcba=h.fcba ) AS SUPERVISINAME1, " & _
" h.SUPERVISI2,(SELECT s2.FCNAME FROM EMPLOYEE s2 WHERE h.SUPERVISI2=s2.FCCODE and s2.fcba=h.fcba) AS SUPERVISINAME2,h.SUPERVISI3, " & _
" (SELECT s3.FCNAME FROM EMPLOYEE s3 WHERE h.SUPERVISI3=s3.FCCODE and s3.fcba=h.fcba) AS SUPERVISINAME3,h.SUPERVISI4, " & _
" (SELECT s4.FCNAME FROM EMPLOYEE s4 WHERE h.SUPERVISI4=s4.FCCODE and s4.fcba=h.fcba) AS SUPERVISINAME4,h.SECTIONCODE,h.DOCUMENT_CLASSIFICATION,c.fcname, h.ROWSTATE,h.POSTINGFLG " & _
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
	ROWSTATE = rsHeader("ROWSTATE")
	POSTINGFLG = rsHeader("POSTINGFLG")
	
end if

IF POSTINGFLG = "0" THEN 
   POSTINGFLG = "TIDAK"
ELSE
   POSTINGFLG	= "YA"
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

'proses dibawah ini digunakan untuk approved bkm dan create jurnal
btnApprove 	= request("btnsave")
bkmno1 		= request("bkmno1")
bkmdate 	= request("bkmdate")
dim strQueryStatus,modulFile,executeQry,modulFile1,executeQry1
if btnApprove = "Approve" then 
	
	strQueryStatus = "UPDATE bkmhdr SET rowstate = 'Approved', postingflg='1',fcedit='" & Session("CodeName") & "',lastupdate=sysdate WHERE fcba = '" & session("fcba") & "' AND bkmno = '" & bkmno1 & "'"
	'connStr.execute(strQueryStatus)
	'modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
	'executeQry	= "Query executed:" & strQueryStatus
	ExecuteAndLog  ScriptName,strQueryStatus
	
	SQL = "begin bkmjournal_api.Update_bkm(TO_DATE('" & bkmdate & "', 'mm/dd/yyyy'),'" & session("fcba") & "','" & Session("CodeName") & "',1); end;"
	'connStr.execute(SQL)
	'modulFile1 	= session("username") & " Execution at: " & Now() &  ScriptName 
	'executeQry1	= "Query executed:" & SQL
	ExecuteAndLog ScriptName,SQL
	
	response.redirect("bkm_adjusment_view.asp?bkmno="&bkmno1)

end if


%>

<html>
	<head>
		<title>FBPlants - View BKM Adjusment</title>
		<link href="/js/tabbedpanels/SpryTabbedPanels.css" rel="stylesheet" type="text/css"/>
		<script src="/js/tabbedpanels/SpryTabbedPanels.js" type="text/javascript"></script>
		<script src="/js/jquery-1.9.1.js"></script>
		<script src="/js/jquery-ui.js"></script>
		<link rel="stylesheet" href="css/jquery-ui.css" />
		<link rel="StyleSheet" href="style-content.css" type="text/css" />
		<link href="/js/spry/widgets/textfieldvalidation/SpryValidationTextField.css" rel="stylesheet" type="text/css"/>
		<script src="/js/spry/widgets/textfieldvalidation/SpryValidationTextField.js" type="text/javascript"></script>
	<% if Request("oldui") <> "1" then %>
<!-- Tampilan baru (css/fbp-shell.css + css/fbp-forms.css). ?oldui=1 = tampilan lama -->
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
<link rel="StyleSheet" href="/css/login/font-awesome.min.css" type="text/css" />
<link rel="StyleSheet" href="/css/fbp-shell.css?v=20261009d" type="text/css" />
<link rel="StyleSheet" href="/css/fbp-forms.css?v=20261009e" type="text/css" />
<script type="text/javascript" src="/js/fbp-shell.js?v=20261015a"></script>
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
		<strong class="posisikiri">&nbsp;&nbsp;PT. <%=CompanyName%> - &nbsp;<%=FcbaName%>   -           MODULE : BKM ADJUSMENT VIEW</strong><br>					
</div>
<br>
<br>
<div class="box_bothall" style="overflow:scroll;">			
<form name="frmEntry" method="post" action="bkm_adjusment_view.asp">
	<input type="hidden" name="jumldata" id="jumldata" value="<%=NoOfRows%>" />
	<input type="hidden" name="jumldataitem" id="jumldataitem" value="<%=NoOfRowsitem%>" />
	<input type="hidden" name="proses" id="proses" value="2" />
	<input type="hidden" name="bkmno1" id="bkmno1" value="<%=bkmno%>" />
	<input type="hidden" name="bkmdate" id="bkmdate" value="<%=DATECREATED%>" />
	<table class="data posisikiri" border="0" width="90%">			
		<tr>
			<td align="left" >BKM No:</td>
			<td align="left">
				<%=bkmno%>
			</td>


			<td align="left" >Section Name:</td>
			<td align="left">
				<%=SECTIONNAME%>
			</td>
			<td align="left" >Mandor I</td>
			<td align="left">
				<select name="supervisi1" id="supervisi1" class="lebarsection" >
					<option value="<%=SUPERVISI1%>"><%=SUPERVISINAME1%></option>
				</select>
			</td>
			<td align="left" >Kerani Panen</td>
			<td align="left">
				<select name="supervisi3" id="supervisi3" class="lebarsection" >
					<option value="<%=SUPERVISI3%>"><%=SUPERVISINAME3%></option>
				</select>
			</td>
		</tr>
		<tr>
			<td align="left" >Tanggal:</td>
			<td align="left">
				<%=DATECREATED%>
			</td>
			
			<td align="left" >Gang:</td>
			<td align="left">
				<%=GANGCODE%>
			</td>
			<td align="left" >Mandor</td>
			<td align="left">
				<select name="supervisi2" id="supervisi2" class="lebarsection" >
					<option value="<%=SUPERVISI2%>"><%=SUPERVISI2%></option>
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
			<td align="left" >Posting:</td>
			<td align="left"><%=POSTINGFLG%>-<%=ROWSTATE%></td>
		</tr>
	</table>
	
	<br>
		
	<table align=center border="0" class='data posisikiri' width='900' id='tableall'>
		<tr>
			<td>
				<div class="TabbedPanels" id="tp1" align='center'>
				
					<ul class="TabbedPanelsTabGroup">
						<li class="TabbedPanelsTab">BKM UPAH</li>
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
									<th>Absen</th>
									<th>Activity</th>
									<th>Uom</th>
									<th>TC</th>
									<th>HK</th>
									<th>Unit</th>
									<th>Output</th>
									<th>Rate</th>
									<th>Allowed Code</th>
									<th>Rate Allowed</th>
									<th>OVTH</th>
									<th>TOVTime</th>
									<th colspan="2">&nbsp;</th>
								</tr>

								<% 
								
								
								'get data detail upah
								dim sqldtlupah,rsdtlupah,no,totalamount,ttlhk,ttlunit,ttloutput,ttlrate,ttlrateall,ttlovth
								sqldtlupah = "SELECT u.BKMNO,u.fcba, u.chargeunit, u.EMPLOYEECODE,e.fcname || '-' ||  e.fccode as fcname, u.ACTCODE,a.fcname || '-' ||  u.ACTCODE as activityname, NVL(u.MANDAYS,0)MANDAYS, " & _
								" u.ALLOWANCE_CODE, nvl(u.ALLOWANCE_AMT,0)ALLOWANCE_AMT, u.ATTENDANCE,nvl(u.OVERTIME_HOURS,0)OVERTIME_HOURS,u.TYPE_OVERTIME,o.fcname type_ovname," & _ 
								" to_char(u.IN_FIN,'hh24:mi:ss')IN_FINDEPAN,to_char(u.OUT_FIN,'hh24:mi:ss')OUT_FINDEPAN, " & _
								" to_char(u.IN_FIN,'dd/mm/yyyy hh24:mi:ss')IN_FIN,to_char(u.OUT_FIN,'dd/mm/yyyy hh24:mi:ss')OUT_FIN,  " & _
								" u.LOCATIONCODE, nvl(u.OUTPUT,0)OUTPUT,nvl(u.rate,0)rate, nvl(u.UNIT,0)UNIT,al.FCCODE || ' - ' || al.fcname as ALLOWDEDNAME,nvl(u.OTHCONVERTION,0)OTHCONVERTION ,a.uom_unit" & _
								" FROM BKMUPAHDTL u " & _ 
								" INNER JOIN employee e ON (u.EMPLOYEECODE=e.fccode and u.fcba=e.fcba) " & _
								" INNER JOIN jobmaster a ON (u.ACTCODE=a.FCCODE) " & _
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
										<%=rsdtlupah("fcname")%>
									</td>
									<td class="ratatengah">
										<%=rsdtlupah("IN_FINDEPAN")%>
									</td>
									<td class="ratatengah">
										<%=rsdtlupah("OUT_FINDEPAN")%>
									</td>
									<td class="ratatengah">
										<%=rsdtlupah("chargeunit")%>
									</td>
									<td class="ratatengah">
										<%=rsdtlupah("ATTENDANCE")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlupah("activityname")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlupah("uom_unit")%>
									</td>
									
									<td class="ratatengah">
										<%=rsdtlupah("LOCATIONCODE")%>
									</td>
									<td>
										<%=rsdtlupah("MANDAYS")%>
									</td>
									
									<td>
										<%=rsdtlupah("UNIT")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlupah("OUTPUT")%>
									</td>
									
									<td class="ratakiri">
										<%=rsdtlupah("rate")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlupah("ALLOWDEDNAME")%>
									</td>
																		
									<td class="ratakiri">
										<%=FormatNumber(rsdtlupah("ALLOWANCE_AMT"),0)%>
									</td>
									<td class="ratatengah">
										<%=rsdtlupah("OVERTIME_HOURS")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlupah("TYPE_OVERTIME")%>&nbsp;<%=rsdtlupah("type_ovname")%>
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
										<td class="tebal ratakiri"><input class="ratatengah lebarhk" type="text" readonly  style="background-color:#eee;" name="ttlhk" id="ttlhk<%=no%>" value="<%=FormatNumber(ttloutput,2)%>"/></td>
										<td class="tebal ratakiri"><input class="ratatengah lebarrate" type="text" readonly  style="background-color:#eee;" name="ttlhk" id="ttlhk<%=no%>" value="<%=FormatNumber(ttlrate,2)%>"/></td>
										<td></td>
										<td class="tebal ratakiri"><input class="ratakanan lebarrate" type="text" readonly  style="background-color:#eee;" name="amount" id="amount<%=no%>" value="<%=FormatNumber(ttlrateall,0)%>"/></td>
										<td class="tebal ratatengah"><input class="ratatengah lebarhk" type="text" readonly  style="background-color:#eee;" name="amount" id="amount<%=no%>" value="<%=FormatNumber(ttlovth,0)%>"/></td>
										<td></td>
										
									</tr>
									
							</table>
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
			<th>Absen</th>
			<th>Activity</th>
			<th>Uom</th>
			<th>TC</th>
			<th>HK</th>
			<th>Unit</th>
			<th>Output</th>
			<th>Rate</th>
			<th>Allowed Code</th>
			<th>Rate Allowed</th>
			<th>OVTH</th>
			<th>Amount</th>
		</tr>
		<%
			'get data detail upah
			dim rsupah,allcode,rateallow,hkview, dtlamt
			sqldtlupah = "SELECT u.BKMNO,u.fcba, u.EMPLOYEECODE,e.fcname || '-' ||  e.fccode as fcname, u.ACTCODE,a.fcname || '-' ||  u.ACTCODE as activityname, NVL(u.MANDAYS,0)MANDAYS, " & _
			" p.ALLOWDEDCODE, nvl(u.ALLOWANCE_AMT,0)ALLOWANCE_AMT, u.ATTENDANCE,nvl(u.OVERTIME_HOURS,0)OVERTIME_HOURS,u.TYPE_OVERTIME,o.fcname type_ovname," & _ 
			" to_char(u.IN_FIN,'hh24:mi:ss')IN_FINDEPAN,to_char(u.OUT_FIN,'hh24:mi:ss')OUT_FINDEPAN, " & _
			" to_char(u.IN_FIN,'dd/mm/yyyy hh24:mi:ss')IN_FIN,to_char(u.OUT_FIN,'dd/mm/yyyy hh24:mi:ss')OUT_FIN,  " & _
			" u.LOCATIONCODE, nvl(u.OUTPUT,0)OUTPUT,nvl(u.rate,0)rate, nvl(u.UNIT,0)unit,p.ALLOWDEDCODE,al.FCCODE || ' - ' || al.fcname as ALLOWDEDNAME,nvl(p.AMOUNT,0)AMOUNT,a.uom_unit " & _
			" FROM BKMUPAHDTL u " & _ 
			" INNER JOIN employee e ON (u.EMPLOYEECODE=e.fccode and u.fcba=e.fcba) " & _
			" INNER JOIN jobmaster a ON (u.ACTCODE=a.FCCODE) " & _
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
			<td class="ratatengah"><%=rsupah("ATTENDANCE")%></td>
			<td class="ratakiri"><%=rsupah("activityname")%></td>
			<td class="ratakiri"><%=rsupah("uom_unit")%></td>
			<td class="ratatengah"><%=rsupah("LOCATIONCODE")%></td>
			<td class="ratatengah"><%=hkview%></td>
			<td class="ratatengah"><%=rsupah("UNIT")%></td>
			<td class="ratatengah"><%=rsupah("OUTPUT")%></td>
			<td class="ratatengah"><%=rsupah("RATE")%></td>
			<td class="ratakiri"><%=rsupah("ALLOWDEDNAME")%></td>
			<td class="ratakanan"><%=formatnumber(rateallow,2)%></td>
			<td class="ratatengah"><%=rsupah("OVERTIME_HOURS")%></td>
			<td class="ratakanan"><%=formatnumber(rsupah("AMOUNT"),2)%></td>
		</tr>
		<%
			totalamount = cdbl(totalamount) + cdbl(rsupah("AMOUNT"))
			rsupah.MoveNext
			Wend
		%>
		<tr>
			<td colspan="13" class="ratakanan tebal">Total</td>
			<td class="ratakanan tebal"><%=formatnumber(totalamount,2)%></td>
		</tr>
		<tr class="ratatengah">
			<td colspan="11">
			<% if Akses_Posting <> 0 and ROWSTATE="Planned" and POSTINGFLG="TIDAK" Then %>
											<input type=Submit name="btnsave" id="btnsave" value="Approve">&nbsp;&nbsp;
			<%end if%>
			<input type=button onclick="window.close();" name="btnclose" value="CLOSE">
			</td>		
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
		}
});



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


</script>
