<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->

<%
dim businessunit,fcbaname,CompanyName,q,sqltotdata,bkmno,tglbkm,tglsplit,bulansplit,thnsplit
bkmno= trim(request("bkmno"))
tglbkm= trim(request("tglbkm"))
businessunit= trim(session("fcba"))

tglsplit  = left(tglbkm,2)
bulansplit  = mid(tglbkm,4,2)
thnsplit    = right(tglbkm,4)


q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & businessunit & "' "
	' response.write q
set rsCat = connStr.Execute(q)
if not rsCat.eof then
   fcbaname = rsCat("fcName")
   CompanyName=rsCat("FCCOMPANYNAME")
 
end if

' get data header  to_date('" & IssueDate & "','mm/dd/yyyy')
dim sqlheader,rsHeader,jurnalno,jurnaldate,remarks
sqlheader = " select docno, to_char(docdate,'dd/mm/yyyy')docdate,remarks from bkmjournalhdr where DOCDATE=to_date('" & tglbkm & "','dd/mm/yyyy') " & _ 
            " and docno in(select docno from bkmjournaldtl where bkmno='" & bkmno & "') AND fcba='" & businessunit & "' AND SOURCECODE='OPR-BKM'"
set rsHeader = connStr.Execute(sqlheader)
'response.write rsHeader
if not rsHeader.eof then 
	jurnalno = rsHeader("docno")
	jurnaldate = rsHeader("docdate")
	remarks = rsHeader("remarks")
end if



    Dim rsAvailable, strAvailable,Access_Post
 	strAvailable =	" select authorizestructure_api.get_access_post('ACC-JOURNAL','" & businessunit & "','" & Session("CodeName") & "') result" & _
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
	
'count detailupah
Dim SQL_Rows, NoOfRows,RSCount

SQL_Rows = "SELECT Count (*) from bkmjournaldtl where docno='" & jurnalno & "'"
set RSCount = connStr.Execute(SQL_Rows)
if not RSCount.eof then 
	NoOfRows = RSCount(0).value
end if
RSCount.close


'get value button
Dim sqlbuton,rsbuton,POSTINGFLG
sqlbuton = "SELECT POSTINGFLG FROM GLEDGERDTL WHERE JOURNALNO='" & jurnalno & "' AND FCBA='" & businessunit & "' "
set rsbuton = connStr.Execute(sqlbuton)
if not rsbuton.eof then 
	POSTINGFLG = rsbuton("POSTINGFLG")
else 
	POSTINGFLG = 0
end if
rsbuton.close

%>

<html>
	<head>
		<title>FBPlants - BKM JURNAL</title>
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
<script type="text/javascript" src="/js/fbp-shell.js?v=20261014b"></script>
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
		<strong class="posisikiri">&nbsp;&nbsp;PT. <%=CompanyName%> - &nbsp;<%=FcbaName%>   -           MODULE : BKM JURNAL</strong><br>					
</div>
<br>
<br>
<div class="box_bothall" style="overflow:scroll;">			
<form name="frmEntry">
<input type="hidden" name="jumldata" id="jumldata" value="<%=NoOfRows%>" />
<input type="hidden" name="sourcecode" id="sourcecode" value="OPR-BKM" />

	<table class="data posisiheader" border="0" width="50%">			
		<tr>
			<td align="left" >Jurnal NO</td>
			<td align="left">
				<input type="text" name="jurnalno" readonly id="jurnalno" size="20" value="<%=jurnalno%>" >
			</td>


			<td align="left" >Jurnal Date</td>
			<td align="left">
				<input type="text" name="jurnaldate" readonly id="jurnaldate" size="20" value="<%=jurnaldate%>" >
			</td>
			
		</tr>
		
		<tr>
			<td align="left" >Remark</td>
			<td align="left" colspan="">
				<textarea  name="remarks" id="remarks" cols="30" rows=30></textarea>
				
			</td>
			<td align="left" >Posting</td>
			<td align="left" colspan="">
				<%
				
					if POSTINGFLG= "0" then 
						response.write "TIDAK"
					elseif POSTINGFLG= "" then
						response.write "TIDAK"
					else 
						response.write "YA"
					end if
				%>
				
				
			</td>
			
		</tr>
		
	</table>
	
	<br>
		
	<table align=center border="0" class='data posisikiri' width='800' id='table1'>
		<tr>
			<td>
				<div class="TabbedPanels" id="tp1" align='center'>
					<ul class="TabbedPanelsTabGroup">
						<li class="TabbedPanelsTab">JURNAL BKM UPAH</li>
						<li class="TabbedPanelsTab">JURNAL BKM BAHAN</li>
						<li class="TabbedPanelsTab">JURNAL BKM VEHICLE</li>
						<li class="TabbedPanelsTab">JURNAL BKM WIP</li>
					</ul>

					<div class="TabbedPanelsContentGroup">

						<div class="TabbedPanelsContent">
							<table align=center border="0" width="900" class='data' id='table2'>
								<tr id="trheader2">
									
									<th>BKMNO</th>
									<th>DESCRIPTION</th>									
									<th>JOBNAME</th>
									<th>TARGET CODE</th>
									<th>DEBIT</th>
									<th>KREDIT</th>

								</tr>

								<% 
								
								
								'get data detail upah
								dim sqldtlupah,rsdtlupah,no,totalamount,totalamountkredit
								sqldtlupah = "select d.BKMNO,j.fccode, d.description,j.fccode || '-' ||  j.fcname as jobcode,d.locationcode, sum(d.amountdebit) amountdebit,sum(d.amountkredit)amountkredit " & _ 
								" from BKMJOURNALHDR h " & _
								" inner join BKMJOURNALDTL d on  (h.DOCNO=d.DOCNO and h.fcba=d.fcba) " & _
								" inner join jobmaster j on (j.fccode=d.JOBCODE) " & _
								" where h.DOCDATE=to_date('" & tglbkm & "','dd/mm/yyyy') AND h.sourcecode='OPR-BKM' AND d.SOURCE='BKM' and h.FCBA='" & businessunit & "' " & _
								" and h.docno='" & jurnalno & "' " & _ 
								" group by d.BKMNO,j.fccode,d.description,d.locationcode,j.fccode || '-' ||  j.fcname,lineno order by lineno" 
								set rsdtlupah = connStr.Execute(sqldtlupah)
								'response.write sqldtlupah
								

																
								no = 0
								totalamount = 0
								totalamountkredit = 0
								While Not rsdtlupah.EOF %>
								<tr id="tru<%=no%>">
									
									<td class="ratakiri">
										<input class="ratakiri lebaremployee" readonly type="text" name="bkmno<%=no%>" id="bkmno<%=no%>" value="<%=rsdtlupah("BKMNO")%>"/>
									</td>
									<td class="ratakiri">
										<input class="ratakiri lebaremployee" readonly type="text" name="description<%=no%>" id="description<%=no%>" value="<%=rsdtlupah("description")%>"/>
									</td>
									<td class="ratakiri">
										<input class="ratakiri lebaractivity" readonly type="text" name="activityname<%=no%>" id="activityname<%=no%>" value="<%=rsdtlupah("jobcode")%>"/>
										<input class="ratakiri lebaractivity" readonly type="hidden" name="activity<%=no%>" id="activity<%=no%>" value="<%=rsdtlupah("fccode")%>"/>
									</td>
									<td class="ratakiri">
										<input class="ratakiri lebaremployee" readonly type="text" name="TC<%=no%>" id="TC<%=no%>" value="<%=rsdtlupah("locationcode")%>"/>
									</td>
									
									<td>
										<input class="ratakanan lebarupah"  type="text" name="debit<%=no%>" id="debit<%=no%>" value="<%=formatnumber(rsdtlupah("amountdebit"),0)%>"/>
									</td>
									<td>
										<input class="ratakanan lebarupah" type="text" name="kredit<%=no%>" id="kredit<%=no%>" value="<%=formatnumber(rsdtlupah("amountkredit"),0)%>"/>
									</td>
								
									<td></td>
									
								</tr>

								
								<%

									

									no = no +1
									totalamount = cdbl(totalamount) + cdbl(rsdtlupah("amountdebit"))
									totalamountkredit = cdbl(totalamountkredit) + cdbl(rsdtlupah("amountkredit"))
									rsdtlupah.MoveNext
									
									Wend
								%>
									<tr class="ratakanan">
										<td colspan="4" class="tebal">Total</td>
										<td class="tebal"><input class="ratakanan lebarupah" type="text" readonly  style="background-color:#eee;" name="amount" id="amount<%=no%>" value="<%=FormatNumber(totalamount,0)%>"/></td>
										<td class="tebal"><input class="ratakanan lebarupah" type="text" readonly  style="background-color:#eee;" name="amount" id="amount<%=no%>" value="<%=FormatNumber(totalamountkredit,0)%>"/></td>
									</tr>
							</table>
						</div>

						<div class="TabbedPanelsContent">
							<!-- ============================================================================================================================= -->
							<div id="datadetail">
								<table align=center border="0" width="900" class='data' id='table1'>
									<tr id="trheader2">
										<th>BKMNO</th>
										<th>JOBNAME</th>
										<th>TARGET CODE</th>
										<th>DEBIT</th>
										<th>KREDIT</th>
										
									</tr>
								<%
										'get data detail bahan
									dim sqljrnlbhn,rsjrnlbhn,nobhn,ttlamountbhn,ttlamountbhnkredit
									sqljrnlbhn = "select d.BKMNO,j.fccode, j.fccode || '-' ||  j.fcname as jobcode,d.locationcode,sum(d.amountdebit) amountdebit,sum(d.amountkredit)amountkredit " & _ 
									" from BKMJOURNALHDR h " & _
									" inner join BKMJOURNALDTL d on  (h.DOCNO=d.DOCNO and h.fcba=d.fcba) " & _
									" inner join job j on (j.fccode=d.JOBCODE and j.fcba=d.fcba) " & _
									" where h.DOCDATE=to_date('" & tglbkm & "','dd/mm/yyyy') AND h.sourcecode='MAM-SIS' AND d.SOURCE='SIS' and h.FCBA='" & businessunit & "' " & _
									" group by d.BKMNO,j.fccode,d.locationcode,j.fccode || '-' ||  j.fcname,lineno order by lineno" 
									set rsjrnlbhn = connStr.Execute(sqljrnlbhn)
								

								nobhn = 0
								ttlamountbhn = 0
								ttlamountbhnkredit = 0 
								While Not rsjrnlbhn.EOF %>
								<tr id="trb<%=nobhn%>">
									
									<td class="ratakiri">
										<input class="ratakiri lebaremployee" readonly type="text" name="bkmnobhn<%=nobhn%>" id="bkmnobhn<%=nobhn%>" value="<%=rsjrnlbhn("BKMNO")%>"/>
									</td>
									<td class="ratakiri">
										<input class="ratakiri lebaractivity" readonly type="text" name="activitybhn<%=nobhn%>" id="activitybhn<%=nobhn%>" value="<%=rsjrnlbhn("jobcode")%>"/>
										<input class="ratakiri lebaractivity" readonly type="hidden" name="activitybhn<%=nobhn%>" id="activitybhn<%=nobhn%>" value="<%=rsjrnlbhn("fccode")%>"/>
									</td>
									<td class="ratakiri">
										<input class="ratakiri lebaremployee" readonly type="text" name="tcbhn<%=nobhn%>" id="tcbhn<%=nobhn%>" value="<%=rsjrnlbhn("locationcode")%>"/>
									</td>
									
									<td>
										<input class="ratakanan lebarupah"  type="text" name="debitbhn<%=nobhn%>" id="debitbhn<%=nobhn%>" value="<%=formatnumber(rsjrnlbhn("amountdebit"),0)%>"/>
									</td>
									<td>
										<input class="ratakanan lebarupah" type="text" name="kreditbhn<%=nobhn%>" id="kreditbhn<%=nobhn%>" value="<%=formatnumber(rsjrnlbhn("amountkredit"),0)%>"/>
									</td>
								
									<td></td>
									
								</tr>

								
								<%

									nobhn = nobhn +1
									if rsjrnlbhn("fccode") <> "1140997" then 
									ttlamountbhn = cdbl(ttlamountbhn) + cdbl(rsjrnlbhn("amountdebit"))
									ttlamountbhnkredit = cdbl(ttlamountbhnkredit) + cdbl(rsjrnlbhn("amountkredit"))
									end if
									rsjrnlbhn.MoveNext
									
									Wend
								%>
									<tr class="ratakanan">
										<td colspan="3" class="tebal">Total</td>
										<td class="tebal"><input class="ratakanan lebarupah" type="text" readonly  style="background-color:#eee;" name="amountbhn" id="amountbhn<%=nobhn%>" value="<%=FormatNumber(ttlamountbhn,0)%>"/></td>
										<td class="tebal"><input class="ratakanan lebarupah" type="text" readonly  style="background-color:#eee;" name="amountbhn" id="amountbhn<%=nobhn%>" value="<%=FormatNumber(ttlamountbhnkredit,0)%>"/></td>
									</tr>
									
								</table>
							</div>
							<!-- ============================================================================================================================= -->
						</div>

						<div class="TabbedPanelsContent">
							<!-- ============================================================================================================================= -->
							<div id="datadetail">
								<table align=center border="0" width="900" class='data' id='table1'>
									<tr id="trheader3">
										<th>BKMNO</th>
										<th>JOBNAME</th>
										<th>TARGET CODE</th>
										<th>DEBIT</th>
										<th>KREDIT</th>
										
									</tr>


										<%
										'get data detail vra
									dim sqljrnlvra,rsjrnlvra,novra,ttlamountvra,ttlamountvrakredit
									sqljrnlvra = "select d.BKMNO,j.fccode, j.fccode || '-' ||  j.fcname as jobcode,d.locationcode,sum(d.amountdebit) amountdebit,sum(d.amountkredit)amountkredit " & _ 
									" from BKMJOURNALHDR h " & _
									" inner join BKMJOURNALDTL d on  (h.DOCNO=d.DOCNO and h.fcba=d.fcba) " & _
									" inner join jobmaster j on (j.fccode=d.JOBCODE) " & _
									" where h.DOCDATE=to_date('" & tglbkm & "','dd/mm/yyyy') AND d.source='VRA' and h.FCBA='" & businessunit & "' " & _
									" group by d.BKMNO,j.fccode,d.locationcode,j.fccode || '-' ||  j.fcname,lineno order by lineno" 
									set rsjrnlvra = connStr.Execute(sqljrnlvra)
								

								novra = 0
								ttlamountvra = 0
								ttlamountvrakredit = 0
								While Not rsjrnlvra.EOF %>
								<tr id="trv<%=novra%>">
									
									<td class="ratakiri">
										<input class="ratakiri lebaremployee" readonly type="text" name="bkmnovra<%=novra%>" id="bkmnovra<%=novra%>" value="<%=rsjrnlvra("BKMNO")%>"/>
									</td>
									<td class="ratakiri">
										<input class="ratakiri lebaractivity" readonly type="text" name="activityvra<%=novra%>" id="activityvra<%=novra%>" value="<%=rsjrnlvra("jobcode")%>"/>
										<input class="ratakiri lebaractivity" readonly type="hidden" name="activityvra<%=novra%>" id="activityvra<%=novra%>" value="<%=rsjrnlvra("fccode")%>"/>
									</td>
									<td class="ratakiri">
										<input class="ratakiri lebaremployee" readonly type="text" name="tcvra<%=novra%>" id="tcvra<%=novra%>" value="<%=rsjrnlvra("locationcode")%>"/>
									</td>
									
									<td>
										<input class="ratakanan lebarupah"  type="text" name="debitvra<%=novra%>" id="debitvra<%=novra%>" value="<%=formatnumber(rsjrnlvra("amountdebit"),0)%>"/>
									</td>
									<td>
										<input class="ratakanan lebarupah" type="text" name="kreditvra<%=novra%>" id="kreditvra<%=novra%>" value="<%=formatnumber(rsjrnlvra("amountkredit"),0)%>"/>
									</td>
								
									<td></td>
									
								</tr>

								
								<%

									novra = novra +1
									ttlamountvra = cdbl(ttlamountvra) + cdbl(rsjrnlvra("amountdebit"))
									ttlamountvrakredit = cdbl(ttlamountvrakredit) + cdbl(rsjrnlvra("amountkredit"))
									rsjrnlvra.MoveNext
									
									Wend
								%>
									<tr class="ratakanan">
										<td colspan="3" class="tebal">Total</td>
										<td class="tebal"><input class="ratakanan lebarupah" type="text" readonly  style="background-color:#eee;" name="amountvra" id="amountvra<%=novra%>" value="<%=FormatNumber(ttlamountvra,0)%>"/></td>
										<td class="tebal"><input class="ratakanan lebarupah" type="text" readonly  style="background-color:#eee;" name="amountvra" id="amountvra<%=novra%>" value="<%=FormatNumber(ttlamountvrakredit,0)%>"/></td>
									</tr>
									
								</table>
							</div>
							<!-- ============================================================================================================================= -->
						</div>

						<div class="TabbedPanelsContent">
							<!-- ============================================================================================================================= -->
							<div id="datadetail">
								<table align=center border="0" width="900" class='data' id='table1'>
									<tr id="trheader4">
										<th>BKMNO</th>
										<th>JOBNAME</th>
										<th>TARGET CODE</th>
										<th>DEBIT</th>
										<th>KREDIT</th>
										
									</tr>


										<%
										'get data detail wip
									dim sqljrnlwip,rsjrnlwip,nowip,ttlamountwip,ttlamountwipkredit
									sqljrnlwip = "select d.BKMNO,j.fccode, j.fccode || '-' ||  j.fcname as jobcode,d.locationcode,sum(d.amountdebit) amountdebit,sum(d.amountkredit)amountkredit " & _ 
									" from BKMJOURNALHDR h " & _
									" inner join BKMJOURNALDTL d on  (h.DOCNO=d.DOCNO and h.fcba=d.fcba) " & _
									" inner join job j on (j.fccode=d.JOBCODE and j.fcba=d.fcba) " & _
									" where h.DOCDATE=to_date('" & tglbkm & "','dd/mm/yyyy') AND d.source='WIP' and h.FCBA='" & businessunit & "' " & _
									" group by d.BKMNO,j.fccode,d.locationcode,j.fccode || '-' ||  j.fcname,lineno order by lineno" 
									set rsjrnlwip = connStr.Execute(sqljrnlwip)
								

								nowip = 0
								ttlamountwip = 0
								ttlamountwipkredit = 0 
								While Not rsjrnlwip.EOF %>
								<tr id="trw<%=nowip%>">
									
									<td class="ratakiri">
										<input class="ratakiri lebaremployee" readonly type="text" name="bkmnowip<%=nowip%>" id="bkmnowip<%=nowip%>" value="<%=rsjrnlwip("BKMNO")%>"/>
									</td>
									<td class="ratakiri">
										<input class="ratakiri lebaractivity" readonly type="text" name="activitywip<%=nowip%>" id="activitywip<%=nowip%>" value="<%=rsjrnlwip("jobcode")%>"/>
										<input class="ratakiri lebaractivity" readonly type="hidden" name="activitywip<%=nowip%>" id="activitywip<%=nowip%>" value="<%=rsjrnlwip("fccode")%>"/>
									</td>
									<td class="ratakiri">
										<input class="ratakiri lebaremployee" readonly type="text" name="tcwip<%=nowip%>" id="tcwip<%=nowip%>" value="<%=rsjrnlwip("locationcode")%>"/>
									</td>
									
									<td>
										<input class="ratakanan lebarupah"  type="text" name="debitwip<%=nowip%>" id="debitwip<%=nowip%>" value="<%=formatnumber(rsjrnlwip("amountdebit"),0)%>"/>
									</td>
									<td>
										<input class="ratakanan lebarupah" type="text" name="kreditwip<%=nowip%>" id="kreditwip<%=nowip%>" value="<%=formatnumber(rsjrnlwip("amountkredit"),0)%>"/>
									</td>
								
									<td></td>
									
								</tr>

								
								<%

									nowip = nowip +1
									ttlamountwip = cdbl(ttlamountwip) + cdbl(rsjrnlwip("amountdebit"))
									ttlamountwipkredit = cdbl(ttlamountwipkredit) + cdbl(rsjrnlwip("amountkredit"))
									rsjrnlwip.MoveNext
									
									Wend
								%>
									<tr class="ratakanan">
										<td colspan="3" class="tebal">Total</td>
										<td class="tebal"><input class="ratakanan lebarupah" type="text" readonly  style="background-color:#eee;" name="amountwip" id="amountwip<%=nowip%>" value="<%=FormatNumber(ttlamountwip,0)%>"/></td>
										<td class="tebal"><input class="ratakanan lebarupah" type="text" readonly  style="background-color:#eee;" name="amountwip" id="amountwip<%=nowip%>" value="<%=FormatNumber(ttlamountwipkredit,0)%>"/></td>
									</tr>
									
									
								</table>
							</div>
							<!-- ============================================================================================================================= -->
						</div>




					</div>
				</div>
			</td>
		</tr>
		<tr class="ratatengah">
			<td  colspan="5">
			              <!--  
							<input type=Submit name="btnsave" id="btnsave" value="POSTING">&nbsp;&nbsp;
			                 <input type=Submit name="btnunposting" id="btnunposting" value="UNPOSTING">&nbsp;&nbsp;
							 <input type=Submit name="btncancel" value="CANCEL">&nbsp;&nbsp;
							<input type=button onclick="window.close();" name="btnclose" value="CLOSE">
							--!>
							 
							 
			</td>
			<input type="hidden" name="POSTINGFLG" id="POSTINGFLG" value="<%=POSTINGFLG%>"/>
			<input type="hidden" name="fcba" id="fcba" value="<%=businessunit%>"/>
			<input type="hidden" name="view" id="view" value="0"/>
		</tr>
	</table>

</form>
</div>
</body>
<!--#include file="bottom.jsp"-->
</html>
<link href="/js/select2/select2.min.css" rel="stylesheet" type="text/css"/>
<script src="/js/select2/select2.min.js" type="text/javascript"></script>

<script type="text/javascript" LANGUAGE="VBScript">
	new Spry.Widget.TabbedPanels("tp1", {
		defaultTab: 0
	});

$(function() {	
	$("#jurnaldate").datepicker();
	var postingflg =  $("#POSTINGFLG").val();
	if(postingflg==0){
		$("#btnsave").show();
		$("#btnunposting").hide();
	}else {
		$("#btnunposting").show();
		$("#btnsave").hide();
	}
	
		
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


function actionBtn()
	{
		var bkmfrom = $("#bkmfrom").val();
		var bkmto 	= $("#bkmto").val();
		var jumlinv = parseInt($('#jumldata').val());
		
		if (bkmfrom != "") { 
				$('#loadingjudul').show();
				$('#loading').show();
				$.post("bkm_viewdata.asp?bkmfrom=" + bkmfrom + "&bkmto=" + bkmto,function(data) {
				$('#bodytable').empty();
				var json  = JSON.parse(data);
				var jumlview  = json.results.length;
				var no = 1;
				for (var i = 0; i < jumlview; i++) {
					var x = json.results[i].fccode
					//alert(x);
					var docketlist = '';
						docketlist += '<tr id="tru'+i+'">';
						docketlist += '<td align="left">' + json.results[i].bkmno + '</td>';
						docketlist += '<td align="left">' + json.results[i].createdate + '</td>';
						docketlist += '<td align="right">' + numformat(json.results[i].totalbkmupah) + '</td>';
						docketlist += '<td align="right">' + json.results[i].totalbkmmaterial + '</td>';				
						docketlist += '<td align="right">' + json.results[i].totalbkmvehicle + '</td>';
						docketlist += '<td align="right">' + json.results[i].totalbkmwip + '</td>';
						docketlist += '<td align="center"><a target=_blank href="bkm_operasional_edit.Asp?bkmno='+ json.results[i].bkmno + '"><img src="IMAGES/edit.gif" alt="Edit Data border="0"></a></td>';
						docketlist += '<td align="center"><a target=_blank href="bkm_jurnal_input.Asp?tglbkm='+ json.results[i].createdate + '"><img src="IMAGES/view.png" alt="View Data border="0"></a></td>';
						docketlist += '</tr>';
						$('#bodytable').append(docketlist);
					no++;
					$("#jumldata").val(jumlview);
					
				}
				$('#loadingjudul').hide();
				$('#loading').hide();
				
			});
		 
		} else {
			$('#bodytable').empty();
		}
	}



</script>
