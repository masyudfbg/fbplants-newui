<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->

<%
dim fccode,bulan,tahun,jobcode,fcba,q,fcbaname,CompanyName ,txtJenis, cmdview ,bkmfrom,bkmto,postingbkm,section ,n,sqlexec
dim modulFile,ScriptName,executeQry
bulan= trim(request("month"))
tahun= trim(request("tahun"))
bkmfrom = trim(request("bkmfrom"))
bkmto   = trim(request("bkmto"))
fcba    = trim(session("fcba"))
section    = trim(request("section"))
postingbkm = trim(request("postingbkm"))

cmdview = "TRUE"

txtJenis = request("txtJenis") '---- Untuk Membedakan Jenis Data Yang akan diolah [insert,edit,delete]

q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & fcba & "' "
	' response.write q
set rsCat = connStr.Execute(q)
if not rsCat.eof then
   fcbaname = rsCat("fcName")
   CompanyName=rsCat("FCCOMPANYNAME") 
end if

if Request("txtJenis")= "UPLOAD" then
	For n=0 to (cdbl(request("no"))-1)
	'response.write "Y-"& n & "-" & request("cb_"&TRIM(n)) & "<br>"

			if cdbl(to_nol(request("cb_"&TRIM(n)))) = 1 then
				'response.write "Y-"& n & "-" & request("cb_"&TRIM(n)) & request("bkmno_"&TRIM(n))  & "<br>"  ' dinonaktifkan: teks debug tampil di atas halaman

                'Insert header
				sqlexec     = " MERGE INTO BKMHDR a " & _ 
				              " USING ( SELECT 'M'||trim(BKMNO) BKMNO,BKMDATE,SECTIONCODE,GANGCODE,SUPERVISI1,SUPERVISI2,SUPERVISI3,SUPERVISI4,CREATEDATE,SYSDATE LASTUPDATE,FCENTRY,'" &  Session("CodeName") & "' FCEDIT,'" & Request.ServerVariables("REMOTE_ADDR") & "' FCIP,FCBA,'1' POSTINGFLG,'Approved' ROWSTATE,DOCUMENT_CLASSIFICATION,0 TOTUPAH,0 TOTBAHAN,0 TOTVRA,0 TOTWIP,VEHICLECODE,HUJAN,CONTRACTORCODE " & _
							  "         FROM MBKMHDR WHERE BKMNO = '" & request("bkmno_"&TRIM(n)) & "' " & _
							  "         AND FCBA='" & request("fcba_"&TRIM(n))  & "'" & _
							  "       )b ON (a.BKMNO=b.BKMNO and a.fcba=b.fcba) " & _ 
							  " WHEN NOT MATCHED THEN " &_ 
							  " Insert (BKMNO,BKMDATE,SECTIONCODE,GANGCODE,SUPERVISI1,SUPERVISI2,SUPERVISI3,SUPERVISI4,CREATEDATE,LASTUPDATE,FCENTRY,FCEDIT,FCIP,FCBA,POSTINGFLG, ROWSTATE,DOCUMENT_CLASSIFICATION,TOTUPAH,TOTBAHAN,TOTVRA,TOTWIP,VEHICLECODE,HUJAN,CONTRACTORCODE) " & _
							  " VALUES (b.BKMNO,b.BKMDATE,b.SECTIONCODE,b.GANGCODE,b.SUPERVISI1,b.SUPERVISI2,b.SUPERVISI3,b.SUPERVISI4,b.CREATEDATE,b.LASTUPDATE,b.FCENTRY,b.FCEDIT,b.FCIP,b.FCBA,b.POSTINGFLG,b.ROWSTATE,b.DOCUMENT_CLASSIFICATION,b.TOTUPAH,b.TOTBAHAN,b.TOTVRA,b.TOTWIP,b.VEHICLECODE,b.HUJAN,b.CONTRACTORCODE)"
				'Response.Write sqlexec
				connStr.execute(sqlexec)
				modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
				executeQry	= "Query executed:" & sqlexec
				ExecuteAndLog modulFile,executeQry
				
				'Insert upahdtl
				sqlexec     = " MERGE INTO BKMUPAHDTL a " & _ 
				              " USING ( " & _
								"	SELECT " & _
								"	'M'||trim(M.BKMNO) BKMNO,M.EMPLOYEECODE,M.ATTENDANCE,M.ACTCODE,M.LOCATIONCODE," & _
								"	M.MANDAYS,M.RATE,M.UNIT,M.OUTPUT,M.ABW," & _
								"	M.BUNCH_WEIGHT,M.KG_BRONDOLAN," & _
								"	M.OVERTIME_HOURS,M.TYPE_OVERTIME," & _
								"	M.CHARGEUNIT,M.CHARGEJOB,M.CHARGECODE," & _
								"	M.ATTENDANCE_FIN,F.TIMEIN IN_FIN,F.TIMEOUT OUT_FIN,M.MANDAYS_FIN," & _
								"	M.ALLOWANCE_CODE,M.ALLOWANCE_AMT," & _
								"	row_number() OVER (PARTITION BY M.bkmno ORDER BY M.employeecode,M.actcode,M.locationcode)  LINENOKEY," & _
								"	M.REMARKS," & _
								"	M.CREATEDATE,SYSDATE LASTUPDATE,M.FCENTRY,M.FCEDIT,'" & Request.ServerVariables("REMOTE_ADDR") & "' FCIP," & _
								"	M.FCBA," & _
								"	M.OTHCONVERTION,M.BUNCHUN,M.BUNCHBA" & _
								"	FROM mbkmupahdtl M" & _
								"	INNER JOIN MBKMHDR H ON H.BKMNO=M.BKMNO AND H.FCBA=M.FCBA " & _
								"	LEFT JOIN USERFINGER.ATTENDANCE_TAB F ON (F.EMPNO =M.EMPLOYEECODE AND F.FCBA=M.FCBA AND F.TDATE=H.BKMDATE )  " & _
								"   WHERE M.BKMNO = '" & request("bkmno_"&TRIM(n)) & "' " & _
							    "         AND M.FCBA='" & request("fcba_"&TRIM(n))  & "'" & _
								"	)b ON (a.BKMNO=b.bkmno AND a.EMPLOYEECODE=b.EMPLOYEECODE and a.ACTCODE=b.ACTCODE and a.LOCATIONCODE=b.LOCATIONCODE and a.CHARGEUNIT=b.CHARGEUNIT and a.fcba=b.fcba)" & _
								"   WHEN NOT MATCHED THEN " &_ 
								"   INSERT ( BKMNO,EMPLOYEECODE,ATTENDANCE,ACTCODE,LOCATIONCODE," & _
								"	MANDAYS,RATE,UNIT,OUTPUT,ABW," & _
								"	BUNCH_WEIGHT,KG_BRONDOLAN," & _
								"	OVERTIME_HOURS,TYPE_OVERTIME," & _
								"	CHARGEUNIT,CHARGEJOB,CHARGECODE," & _
								"	ATTENDANCE_FIN,IN_FIN,OUT_FIN,MANDAYS_FIN," & _
								"	ALLOWANCE_CODE,ALLOWANCE_AMT," & _
								"	LINENOKEY," & _
								"	REMARKS," & _
								"	CREATEDATE,LASTUPDATE,FCENTRY,FCEDIT,FCIP," & _
								"	FCBA," & _
								"	OTHCONVERTION,BUNCHUN,BUNCHBA) " & _ 
								"   VALUES ( b.BKMNO,b.EMPLOYEECODE,b.ATTENDANCE,b.ACTCODE,b.LOCATIONCODE," & _
								"	b.MANDAYS,b.RATE,b.UNIT,b.OUTPUT,b.ABW," & _
								"	b.BUNCH_WEIGHT,b.KG_BRONDOLAN," & _
								"	b.OVERTIME_HOURS,b.TYPE_OVERTIME," & _
								"	b.CHARGEUNIT,b.CHARGEJOB,b.CHARGECODE," & _
								"	b.ATTENDANCE_FIN,b.IN_FIN,b.OUT_FIN,b.MANDAYS_FIN," & _
								"	b.ALLOWANCE_CODE,b.ALLOWANCE_AMT," & _
								"	b.LINENOKEY," & _
								"	b.REMARKS," & _
								"   b.CREATEDATE,b.LASTUPDATE,b.FCENTRY,b.FCEDIT,b.FCIP," & _
								"	b.FCBA," & _
								"	b.OTHCONVERTION,b.BUNCHUN,b.BUNCHBA) " 
								'Response.Write sqlexec
								connStr.execute(sqlexec)
								modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
								executeQry	= "Query executed:" & sqlexec
								ExecuteAndLog modulFile,executeQry
				'Proses BKM
				sqlexec = "begin bkmproses_upah_api.insert_bkm('M" & trim(request("bkmno_"&TRIM(n))) & "',TO_DATE('" & request("bkmdate_"&TRIM(n)) & "', 'mm/dd/yyyy'),'" & request("fcba_"&TRIM(n)) & "','" & Session("CodeName") & "'); end;"
				connStr.execute(sqlexec)
				modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
				executeQry	= "Query executed:" & sqlexec
				ExecuteAndLog modulFile,executeQry
				
				'SET uploadflg=1
				sqlexec = "UPDATE mbkmhdr SET postingflg=1 WHERE bkmno = '" & request("bkmno_"&TRIM(n)) & "' AND fcba='" & request("fcba_"&TRIM(n)) & "'"
				connStr.execute(sqlexec)
				modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
				executeQry	= "Query executed:" & sqlexec
				ExecuteAndLog modulFile,executeQry
				
			
			end if
	next

else	
	cmdview = "TRUE"
end if

%>

<html>
	<head>
		<title>FBPlants - MBKM VIEW</title>
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
		left: 20px;
	}
	.posisiheader{
		position: relative;
		left: 100px;
	}
	.lebarsection{
		width: 160px;
	}
	.lebargangcode{
		width: 200px;
	}
	.lebarabsen{
		width: 60px;
	}
	.lebaremployee{
		width: 230px;
	}
	.lebaractivity{
		width: 250px;
	}
	.lebaracllowedcode{
		width: 210px;
	}
	.lebarupah{
		width: 120px;
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

	
</style>

<script language="JavaScript">$(function() {
$( ".datepicker" ).datepicker();
});
</script>

<script>

</script>

<body<% if Request("oldui") <> "1" then %> class="dash page-show page-legacy page-bkm"<% end if %>>
<!--#include file="top.jsp"-->

<br>
<br>
<br>
<br>
<br>


<div id="head_judul" style="float:left;font-size:10px; text-align:left;">
		<strong class="posisikiri">&nbsp;&nbsp;PT. <%=CompanyName%> - &nbsp;<%=FcbaName%>   -           MODULE : Mobile BKM</strong><br>					
</div>
<br>
<br>
<div class="box_bothall" style="overflow:scroll;">			
<form action="mbkm_data_view.Asp" method="post" name="eForm" >
	<table class="data posisiheader" border="0" width="50%">			
		<tr>
			<td align="left" >Tanggal</td>
			<td align="left">
				<input class="datepicker" type="text" maxlength="10" name="bkmfrom" id="bkmfrom" size="10" value="<%=bkmfrom%>" title="Type Date mm/dd/yyyy" >
				&nbsp;s/d&nbsp;
				<input class="datepicker" type="text" maxlength="10" name="bkmto" id="bkmto" size="10" value="<%=bkmto%>" title="Type Date mm/dd/yyyy" >
			</td>
		</tr>
		
		<tr>
			<td align="left" >FCBA</td>
			<td align="left" >
			    <input type="text" name="fcbaname" readonly size="7" value="<%=fcba%>" >	&nbsp;
				<input type="text" name="fcbaname" readonly id="fcbaname" size="50" value="<%=fcbaname%>" >				
			</td>						
		</tr>
		
		<tr>
			<td align="left" >SECTION</td>
			<td align="left" >
				<select name="section" id="section">
					<option value="">--All--</option> 
					<% 	
					    Dim rsComboAFD,usrt 
						usrt = session("CodeName")
						q = " select s.fccode as id, s.fccode as text from sectionname s " & _
						    " inner join user_section_tab u on s.fccode=u.sectionname " & _ 
							" where u.fcusercode='" & usrt & "' order by s.fccode asc"
						' response.write q
						Set rsComboAFD = Server.CreateObject("ADODB.Recordset")
						rsComboAFD.CursorType = 3
						rsComboAFD.CursorLocation = 3
						rsComboAFD.Open q, connstr
						if not rsComboAFD.eof then%>
						<%
						Do While Not rsComboAFD.eof
						%>
							<option value="<%=rsComboAFD("id")%>" <%if cstr(""&section)=cstr(""&rsComboAFD("id")) then %> selected 
								<%End If%>><%=rsComboAFD("text")%>
							</option> 
						<%rsComboAFD.movenext
						loop
						End if
						rsComboAFD.Close
						set rsComboAFD = Nothing
						q=""
					%>
				</select>
			</td>
		</tr>
		
		<tr>
			<td align="left" >UPLOAD</td>
			<td align="left" >
						<select name="postingbkm" id="postingbkm" >
							<option value="">--ALL--</option>
							<option value="1">Ya</option>
							<option value="0">Tidak</option>
						</select>
			</td>
		</tr>
				
		<tr>
		<td colspan="4" align="left">
		<!--<a class="button_menu_top" onclick="createData('PYV')" ><img src="ICON/Bank.png" border="0" ><span> New Payment </span></a>&nbsp;
		<a class="button_menu_top" onclick="createData('RCV')" ><img src="ICON/Bank.png" border="0" ><span> New Receive </span></a>&nbsp;
		<a class="button_menu_top" onclick="createData()" ><img  src="ICON/Bank.png" border="0" ><span> New Pindah Buku</span></a>&nbsp;	-->	
		<input type=Submit name="txtJenis" value="SEARCH" >&nbsp;&nbsp;
		<input type=Submit name="txtJenis" id="btnsave" value="UPLOAD" onClick="return confirm('Anda yakin untuk melakukan UPLOAD BKM yang sudah dipilih?')">&nbsp;&nbsp;
		<input type=Submit onclick="viewListing();" name="txtJenis" value="VIEW LISTING BKM MOBILE" >&nbsp;&nbsp;
		<input type=button onclick="window.close();" name="btnclose" value="CLOSE">
		
		</td>
		</tr>
		
	</table>
	
	<br>
		
	<table align=center border="0" class='data posisikiri' width='800' id='table1'>
		<tr>
			<td>
				<div class="TabbedPanels" id="tp1" align='center'>
					<ul class="TabbedPanelsTabGroup">
						<li class="TabbedPanelsTab">Mobile.BKM</li>
					</ul>

					<div class="TabbedPanelsContentGroup">

						<div class="TabbedPanelsContent">
						 	<table align=center border="0" width="1050" class='data' id='table2'>
								<tr id="trheader2">										
									<th>No.MBKM</th>
									<th>MBKM.Date</th>
									<th>BKM.Type</th>
									<th>SECTION</th>
									<th>GANG</th>
									<th>Input.Date</th>
									<th>Input.User</th>
									<th>Upload.Date</th>
									<th>Upload.User</th>
									<th><input type='checkbox' name='select-all' id='select-all'/>Upload</th>
									<th>Aksi</th>
								</tr>
								
								<%
		
								'get data detail upah
								dim sqldtlupah,rsdtlupah,no,totalamount,postingflg
								if cmdview = "TRUE" then
								sqldtlupah = " SELECT bkmno, to_char(bkmdate,'MM/DD/YYYY')bkmdate," & _
											 " case document_classification " & _ 
											 " when '501' then 'BKM-PERAWATAN' " & _ 
											 " when '503' then 'BKM-PANEN' " & _ 
											  " when '502' then 'BKM-TRAKSI' " & _ 
											" when '504' then 'BKM-WIP' End type,sectioncode, gangcode, " & _
											 " to_char(createdate,'MM/DD/YYYY') createdate , fcentry, to_char(lastupdate,'MM/DD/YYYY') lastupdate , fcedit, " & _ 
											 " postingflg,fcba " & _
											 " FROM mbkmhdr " & _
											 " where fcba='" & fcba & "'" & _ 
											 " AND bkmdate between to_date('"& bkmfrom & "','MM/DD/YYYY') and to_date('" & bkmto & "','MM/DD/YYYY') " & _ 
											 " AND sectioncode like '" & section & "%'" & _ 
											 " AND postingflg like '" & postingbkm & "%'" & _
											 " ORDER BY bkmdate DESC, sectioncode, gangcode "
								
								'response.write sqldtlupah
                           	
								set rsdtlupah = connStr.Execute(sqldtlupah)
								no = 0
								While Not rsdtlupah.EOF 
								%>
								<tr id="tru<%=no%>">
									
									<td class="ratakiri">
										<%=rsdtlupah("bkmno")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlupah("bkmdate")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlupah("type")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlupah("sectioncode")%>
									</td>
									
									<td class="ratakiri">
										<%=rsdtlupah("gangcode")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlupah("createDate")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlupah("fcentry")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlupah("lastupdate")%>
									</td>
									
									<td class="ratakiri">
										<%=rsdtlupah("fcedit")%>
									</td>
									
			                        <td class="ratakiri">
										<%if rsdtlupah("postingflg")="1" then %> 
										<%response.write "Ya" %> 
										<%else%> 
										<input type="checkbox" name="cb_<%=NO%>" id="cb_<%=NO%>" class='item' value= "" onclick='handleClick(this);'>
										<%response.write "Tidak" end if%>
										<input type="hidden" name="fcba_<%=NO%>" value="<%=rsdtlupah("fcba")%>">
										<input type="hidden" name="bkmno_<%=NO%>" value="<%=rsdtlupah("bkmno")%>">
										<input type="hidden" name="bkmdate_<%=NO%>" value="<%=rsdtlupah("bkmdate")%>">
										
									</td>
																		
									<td class="ratatengah"">
									<a href="bkm_oprupload_view.Asp?bkmno=<%=rsdtlupah("bkmno")%>&tglbkm=<%=rsdtlupah("bkmdate")%>&fcba=<%=rsdtlupah("fcba")%>" target="_blank">
									<img src="IMAGES/view.png" target="_blank"  alt="View Data : <%=rsdtlupah("bkmno")%>" border="0">
									</a>
									</td>
									
								</tr>
								<%

									no = no +1
									rsdtlupah.MoveNext
									Wend
								%>
								        <input Type="Hidden" Value="<%=no%>" Name="no">
								<%end if%>
							</table>
						</div>
					</div>
				</div>
			</td>
		</tr>
		<tr class="ratatengah">
			
		
			 
			
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
	function handleClick(cb){
		if(cb.checked == 1)
		{
			cb.value = 1;
		} else {
			cb.value = 0;
		}
		}
	

</script>
<script type="text/javascript">	
		$('#select-all').click( function () {
			$('.item').prop('checked', this.checked);
			$('.item').val(this.checked ? 1 : 0);
			
		})
</script>
	
<script type="text/javascript" LANGUAGE="VBScript">
	new Spry.Widget.TabbedPanels("tp1", {
		defaultTab: 0
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

function viewListing()
{
	
		var bkmfrom = $("#bkmfrom").val();
		var bkmto 	= $("#bkmto").val();
		var section = $("#section").val();
		
		window.open('p_listing_bkm_mobile.asp?bkmfrom=' + bkmfrom + '&bkmto=' + bkmto + '&section=' + section , '_blank');
}

</script>
