<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/ConndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
Dim MM, yyyy, mon, Yr, MinYr, MaxYr, i, rsBa, strsql,strsql1,strsql2,strsql3,strsql4,strsql5,strsql6,strsql7,strsql8
Dim ScriptName
ScriptName	=	Request.ServerVariables("SCRIPT_NAME")
%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>iPla'S Reporting - Rekap Payroll</title>
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
</head>

<body>
	<div id="wrapper">
	<div id="header-wrapper">
		<table id="form-table" class="centered-margin">
			<tr>
			<th>TAHUN</th>
			<th>BULAN</th>
			<th>&nbsp;</th>
			</tr>
			<tr>
					<form name="frmAR" method="post" action="<%=ScriptName%>">
						<td>
							<font size="2"  color="#0000FF" ></font>
			<% 
				MM=Request.Form("vMonth")
				yyyy=Request.Form("vYear")

				mon = Month(Session("v_Costbook"))
				If len(MM)=0 then MM=mon

				Yr = Year(Session("v_Costbook"))
				MinYr = Yr - 2
				MaxYr = Yr  + 2
				If len(yyyy)= 0 then yyyy = Yr
			%>
							<select name="vYear">
								<%
									For i = MinYr to MaxYr
								%>
										<Option Value="<%=i%>" <%If cint(i) = cint(yyyy) then%> Selected <%End If%>><%=i%></option>
								<%
									Next
								%>
							</select>
						</td>
						
						<td>
							<font size="2" color="#0000FF" ></font>
								<select name="vMonth">
									<%
										For i = 1 to 12
									%>
											<Option Value="<%=i%>" <%If cint(i) = cint(MM) then%> Selected <%End If%>><%=MonthName(i)%></option>
									<%
										Next
									%>
								</select>
						</td>		

						


						
			<td><input type="submit" value="..View.." /></td>
			</form>
			</tr>
		</table>

		<!--#include file="menutop.asp"-->
		
		<div id="header">
			<h4>PT. <%=Session("COMPANY")%></h4>
			<h5><%=Session("BA")%></h5>
		</div>
		<div id="subheader" class="centered">
			<% Dim ia,rekap
		        ia = request.Form("vMonth")
			%>	
			<h2>REKAP GAJI KARYAWAN</h2>
			<% IF request.Form("vMonth") <> "" then %>
			<h5>PERIODE     : <%=MonthName(cint(ia))%> &nbsp; <%=request.Form("vYear")%></h5>
			<% END IF %>
			 
		</div>
	</div>



<table border="0" width="100%" id="product-table">
     <tr>
		
		<th rowspan="2" >Afdeling</th>
		
		<th rowspan="2" >Status</th>
		
		<th rowspan="2">Gaji</th>
		<th rowspan="2" >Lembur</th>
		<th rowspan="2" >Premi</th>
		<th rowspan="2">Tunjangan</th>
		<th rowspan="2">Borongan</th>
		
		<th colspan="5">Potongan</th>
		
		<th rowspan="2">Total</th>
		<th rowspan="2">Tanda Tangan</th>
	</tr>
	<tr>
		
		
		<th>Jamsostek</th>
		<th>MK/P1</th>
		<th>Panen</th>
		<th>PPh21</th>
		<th>Lain-Lain</th>
		
		
	</tr>
	<%
	Dim rsha,strdtl , strgrp ,cdgrp1,cdgrp,cdgrpdesc  ,noline 
	Dim subAmtbsal,subAmtOT,subAmtpremi,subAmtTunj,subAmtBorong,subAmtJamsos,subAmtMKP1,subAmtPanen,subAmtPPH21,subAmtPlain, subAmtgrp
	Dim grnAmtbsal,grnAmtOT,grnAmtpremi,grnAmtTunj,grnAmtBorong,grnAmtJamsos,grnAmtMKP1,grnAmtPanen,grnAmtPPH21,grnAmtPlain, grnAmtgrp
			strdtl = " select t.YearMonth,t.fcba,t.Afdeling,t.presentgrade,  " & _
			         " sum(t.base_salary) base_salary,sum(t.Overtime)  Overtime,sum(t.Premi) Premi, sum(t.Tunjangan) Tunjangan,sum(t.Borongan) Borongan,sum(t.P_Jamsostek) P_Jamsostek,  " & _
				     " sum(t.P_MK_P1) P_MK_P1,sum(t.P_Panen) P_Panen,sum(t.P_PPH21) P_PPH21,sum(t.P_Lain) P_Lain,  " & _
					 " sum(t.base_salary + t.Overtime + t.Premi + t.Tunjangan + t.Borongan + t.P_Jamsostek + t.P_MK_P1 + t.P_Panen + t.P_PPH21 + t.P_Lain ) total " & _
					 " from VIEW_PAYROLLDETAIL t   " & _
					 " where  t.presentgrade ='KBT' and TO_NUMBER(SUBSTR(T.YEARMONTH,1,4)) = " & yyyy & _
					 " and TO_NUMBER(SUBSTR(T.YEARMONTH,5,2)) = " & MM & " and t.fcba ='" & session("fcba") & "' " & _
					 " group by t.YearMonth,t.fcba,t.Afdeling,t.presentgrade   " & _
					 " order by t.YearMonth,t.Afdeling,t.presentgrade,t.fcba   " 
			
	Set rsha = Server.CreateObject("ADODB.Recordset")
	rsha.Open strdtl, connstr, adOpenStatic
    
	  

	if Not rsha.eof then
		while Not rsha.eof  %>
		           
		  <tr>	 
			<td align="center"><%=rsha("Afdeling")%></td>
			
			<td align="center"><%=rsha("presentgrade")%></td>
			<td align="right"><%=formatnumber( cdbl(rsha("base_salary")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsha("Overtime")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsha("Premi")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsha("Tunjangan")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsha("Borongan")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsha("P_Jamsostek")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsha("P_MK_P1")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsha("P_Panen")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsha("P_PPH21")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsha("P_Lain")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsha("total")),0)%></td>
			<% if cdgrp1 = ""  then 
			 
						  Dim rs1 , sql1, r1
   						  sql1    = "  select count(Afdeling) rowsd  from " & _
								   "  (select t.YearMonth,t.fcba,t.Afdeling,t.presentgrade " & _
								   "  from VIEW_PAYROLLDETAIL t " & _
								   "  where  TO_NUMBER(SUBSTR(T.YEARMONTH,1,4)) = " & yyyy & _
								   " and TO_NUMBER(SUBSTR(T.YEARMONTH,5,2)) = " & MM & " and t.fcba ='" & session("fcba") & "' " & _
								   "  and t.presentgrade = 'KBT'  " & _ 
								   "  group by t.YearMonth,t.fcba,t.Afdeling,t.presentgrade " & _
								   "  order by t.YearMonth,t.Afdeling,decode(t.presentgrade,'KBT',1,'KHT',2,'KHL',3),t.presentgrade,t.fcba) " 
								  
								   
                        Set rs1 = Server.CreateObject("ADODB.Recordset")
	                    rs1.Open sql1, connstr, adOpenStatic
						if Not rs1.eof then
						   r1 = rs1("rowsd")
						end if
						rs1.close
						
				%>
			<td align="center" rowspan="<%=cdbl(r1)+1%>" ></td>
			<% end if %>
		  </tr>	
		      
    <%        
	        ' cari sub total
			
	          subAmtbsal = cdbl(subAmtbsal) + cdbl(rsha("base_salary")) 
	          subAmtOT = cdbl(subAmtOT) + cdbl(rsha("Overtime"))
			  subAmtpremi = cdbl(subAmtpremi) + cdbl(rsha("Premi")) 
			  subAmtTunj = cdbl(subAmtTunj) + cdbl(rsha("Tunjangan")) 
	          subAmtBorong = cdbl(subAmtBorong) + cdbl(rsha("Borongan")) 
	          subAmtJamsos = cdbl(subAmtJamsos) + cdbl(rsha("P_Jamsostek")) 
	          subAmtMKP1 = cdbl(subAmtMKP1) + cdbl(rsha("P_MK_P1")) 
			  subAmtPanen = cdbl(subAmtPanen) + cdbl(rsha("P_Panen")) 
	          subAmtPPH21 = cdbl(subAmtPPH21) + cdbl(rsha("P_PPH21")) 
	          subAmtPlain = cdbl(subAmtPlain) + cdbl(rsha("P_Lain")) 
	          subAmtgrp = cdbl(subAmtgrp) + cdbl(rsha("Total"))
			  
			 ' cari grand total 
			  
			  grnAmtbsal = cdbl(grnAmtbsal) + cdbl(rsha("base_salary")) 
	          grnAmtOT = cdbl(grnAmtOT) + cdbl(rsha("Overtime"))
			  grnAmtpremi = cdbl(grnAmtpremi) + cdbl(rsha("Premi")) 
			  grnAmtTunj = cdbl(grnAmtTunj) + cdbl(rsha("Tunjangan")) 
	          grnAmtBorong = cdbl(grnAmtBorong) + cdbl(rsha("Borongan")) 
	          grnAmtJamsos = cdbl(grnAmtJamsos) + cdbl(rsha("P_Jamsostek")) 
	          grnAmtMKP1 = cdbl(grnAmtMKP1) + cdbl(rsha("P_MK_P1")) 
			  grnAmtPanen = cdbl(grnAmtPanen) + cdbl(rsha("P_Panen")) 
	          grnAmtPPH21 = cdbl(grnAmtPPH21) + cdbl(rsha("P_PPH21")) 
	          grnAmtPlain = cdbl(grnAmtPlain) + cdbl(rsha("P_Lain")) 
	          grnAmtgrp = cdbl(grnAmtgrp) + cdbl(rsha("Total"))
			  cdgrp1 = rsha("Afdeling") 
		      rsha.MoveNext
		 wend
		%> 
		 <tr>	 
			<td align="center" colspan="2" ><b>Sub Total</td>
			<td align="right"><b><%=formatnumber(subAmtbsal,0)%></td>
			<td align="right"><b><%=formatnumber(subAmtOT,0)%></td>
			<td align="right"><b><%=formatnumber(subAmtpremi,0)%></td>
			<td align="right"><b><%=formatnumber(subAmtTunj,0)%></td>
			<td align="right"><b><%=formatnumber(subAmtBorong,0)%></td>
			<td align="right"><b><%=formatnumber(subAmtJamsos,0)%></td>
			<td align="right"><b><%=formatnumber(subAmtMKP1,0)%></td>
			<td align="right"><b><%=formatnumber(subAmtPanen,0)%></td>
			<td align="right"><b><%=formatnumber(subAmtPPH21,0)%></td>
			<td align="right"><b><%=formatnumber(subAmtPlain,0)%></td>
			<td align="right"><b><%=formatnumber(subAmtgrp,0)%></td>
		 </tr>	
		 <tr>	 
			<td align="center" colspan="14" ></td>
		 </tr>	
		 
	<%	 
	end if
					 subAmtbsal = cdbl("0") 
					 subAmtOT = cdbl("0") 
					 subAmtpremi = cdbl("0") 
					 subAmtTunj = cdbl("0") 
					 subAmtBorong = cdbl("0") 
					 subAmtJamsos = cdbl("0") 
					 subAmtMKP1 = cdbl("0") 
					 subAmtPanen = cdbl("0") 
					 subAmtPPH21 = cdbl("0") 
					 subAmtPlain = cdbl("0") 
					 subAmtgrp = cdbl("0")
	%>	
        
    <%
		
			strdtl = " select t.YearMonth,t.fcba,t.Afdeling,t.presentgrade,  " & _
			         " sum(t.base_salary) base_salary,sum(t.Overtime)  Overtime,sum(t.Premi) Premi, sum(t.Tunjangan) Tunjangan,sum(t.Borongan) Borongan,sum(t.P_Jamsostek) P_Jamsostek,  " & _
				     " sum(t.P_MK_P1) P_MK_P1,sum(t.P_Panen) P_Panen,sum(t.P_PPH21) P_PPH21,sum(t.P_Lain) P_Lain,  " & _
					 " sum(t.base_salary + t.Overtime + t.Premi + t.Tunjangan + t.Borongan + t.P_Jamsostek + t.P_MK_P1 + t.P_Panen + t.P_PPH21 + t.P_Lain ) total " & _
					 " from VIEW_PAYROLLDETAIL t   " & _
					 " where  t.presentgrade <> 'KBT' and TO_NUMBER(SUBSTR(T.YEARMONTH,1,4)) = " & yyyy & _
					 " and TO_NUMBER(SUBSTR(T.YEARMONTH,5,2)) = " & MM & " and t.fcba ='" & session("fcba") & "' " & _
					 " group by t.YearMonth,t.fcba,t.Afdeling,t.presentgrade   " & _
					 " order by t.YearMonth,t.Afdeling,decode(t.presentgrade,'KBT',1,'KHT',2,'KHL',3),t.presentgrade,t.fcba   " 
			
	
	Set rsha = Server.CreateObject("ADODB.Recordset")
	rsha.Open strdtl, connstr, adOpenStatic
    
	  

	if Not rsha.eof then
		while Not rsha.eof 
			  IF cdgrp = rsha("Afdeling")  then %>
					  <tr>	 
						
						<td align="center"><%=rsha("presentgrade")%></td>
						<td align="right"><%=formatnumber( cdbl(rsha("base_salary")),0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsha("Overtime")),0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsha("Premi")),0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsha("Tunjangan")),0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsha("Borongan")),0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsha("P_Jamsostek")),0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsha("P_MK_P1")),0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsha("P_Panen")),0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsha("P_PPH21")),0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsha("P_Lain")),0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsha("total")),0)%></td>
						
					  </tr>	
					  <%
								
					' cari sub total
						subAmtbsal = cdbl(subAmtbsal) + cdbl(rsha("base_salary")) 
						subAmtOT = cdbl(subAmtOT) + cdbl(rsha("Overtime"))
						subAmtpremi = cdbl(subAmtpremi) + cdbl(rsha("Premi")) 
						subAmtTunj = cdbl(subAmtTunj) + cdbl(rsha("Tunjangan")) 
						subAmtBorong = cdbl(subAmtBorong) + cdbl(rsha("Borongan")) 
						subAmtJamsos = cdbl(subAmtJamsos) + cdbl(rsha("P_Jamsostek")) 
						subAmtMKP1 = cdbl(subAmtMKP1) + cdbl(rsha("P_MK_P1")) 
						subAmtPanen = cdbl(subAmtPanen) + cdbl(rsha("P_Panen")) 
						subAmtPPH21 = cdbl(subAmtPPH21) + cdbl(rsha("P_PPH21")) 
						subAmtPlain = cdbl(subAmtPlain) + cdbl(rsha("P_Lain")) 
						subAmtgrp = cdbl(subAmtgrp) + cdbl(rsha("Total"))
							  
					' cari grand total 
							  
						grnAmtbsal = cdbl(grnAmtbsal) + cdbl(rsha("base_salary")) 
						grnAmtOT = cdbl(grnAmtOT) + cdbl(rsha("Overtime"))
						grnAmtpremi = cdbl(grnAmtpremi) + cdbl(rsha("Premi")) 
						grnAmtTunj = cdbl(grnAmtTunj) + cdbl(rsha("Tunjangan")) 
						grnAmtBorong = cdbl(grnAmtBorong) + cdbl(rsha("Borongan")) 
						grnAmtJamsos = cdbl(grnAmtJamsos) + cdbl(rsha("P_Jamsostek")) 
						grnAmtMKP1 = cdbl(grnAmtMKP1) + cdbl(rsha("P_MK_P1")) 
						grnAmtPanen = cdbl(grnAmtPanen) + cdbl(rsha("P_Panen")) 
						grnAmtPPH21 = cdbl(grnAmtPPH21) + cdbl(rsha("P_PPH21")) 
						grnAmtPlain = cdbl(grnAmtPlain) + cdbl(rsha("P_Lain")) 
						grnAmtgrp = cdbl(grnAmtgrp) + cdbl(rsha("Total"))
					  
					  
			 ELSE    
    			 if cdgrp <> "" then %>	  
					  <tr>	 
						<td align="center" colspan="2" ><b>Sub Total <%=cdgrp%></td>
						<td align="right"><b><%=formatnumber(subAmtbsal,0)%></td>
						<td align="right"><b><%=formatnumber(subAmtOT,0)%></td>
						<td align="right"><b><%=formatnumber(subAmtpremi,0)%></td>
						<td align="right"><b><%=formatnumber(subAmtTunj,0)%></td>
						<td align="right"><b><%=formatnumber(subAmtBorong,0)%></td>
						<td align="right"><b><%=formatnumber(subAmtJamsos,0)%></td>
						<td align="right"><b><%=formatnumber(subAmtMKP1,0)%></td>
						<td align="right"><b><%=formatnumber(subAmtPanen,0)%></td>
						<td align="right"><b><%=formatnumber(subAmtPPH21,0)%></td>
						<td align="right"><b><%=formatnumber(subAmtPlain,0)%></td>
						<td align="right"><b><%=formatnumber(subAmtgrp,0)%></td>
					 </tr>	
					 <tr>	 
						<td align="center" colspan="14" ></td>
					 </tr>
		<%
					 subAmtbsal = cdbl("0") 
					 subAmtOT = cdbl("0") 
					 subAmtpremi = cdbl("0") 
					 subAmtTunj = cdbl("0") 
					 subAmtBorong = cdbl("0") 
					 subAmtJamsos = cdbl("0") 
					 subAmtMKP1 = cdbl("0") 
					 subAmtPanen = cdbl("0") 
					 subAmtPPH21 = cdbl("0") 
					 subAmtPlain = cdbl("0") 
					 subAmtgrp = cdbl("0")
                 end if	 %>
				     <%
						  Dim rsrow , sqlrow, rows
   						  sqlrow = "  select YearMonth,fcba,Afdeling,count(Afdeling) rowsd  from " & _
								   "  (select t.YearMonth,t.fcba,t.Afdeling,t.presentgrade " & _
								   "  from VIEW_PAYROLLDETAIL t " & _
								   "  where  TO_NUMBER(SUBSTR(T.YEARMONTH,1,4)) = " & yyyy & _
								   " and TO_NUMBER(SUBSTR(T.YEARMONTH,5,2)) = " & MM & " and t.fcba ='" & session("fcba") & "' " & _
								   "  and t.presentgrade <> 'KBT' and t.Afdeling = '" & rsha("Afdeling") & "'  " & _ 
								   "  group by t.YearMonth,t.fcba,t.Afdeling,t.presentgrade " & _
								   "  order by t.YearMonth,t.Afdeling,decode(t.presentgrade,'KBT',1,'KHT',2,'KHL',3),t.presentgrade,t.fcba) " & _
								   "  group by YearMonth,fcba,Afdeling "
								   
                        Set rsrow = Server.CreateObject("ADODB.Recordset")
	                    rsrow.Open sqlrow, connstr, adOpenStatic
						if Not rsha.eof then
						   rows = rsrow("rowsd")
						end if
					 %>
						
                     <tr>	 
						<td align="center" rowspan="<%=cdbl(rows)%>" ><%=rsha("Afdeling")%></td>
						<td align="center"><%=rsha("presentgrade")%></td>
						<td align="right"><%=formatnumber( cdbl(rsha("base_salary")),0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsha("Overtime")),0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsha("Premi")),0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsha("Tunjangan")),0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsha("Borongan")),0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsha("P_Jamsostek")),0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsha("P_MK_P1")),0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsha("P_Panen")),0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsha("P_PPH21")),0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsha("P_Lain")),0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsha("total")),0)%></td>
						<td align="center" rowspan="<%=cdbl(rows)+1%>" ></td>
					  </tr>	
					  <%
								
					' cari sub total
						subAmtbsal = cdbl(subAmtbsal) + cdbl(rsha("base_salary")) 
						subAmtOT = cdbl(subAmtOT) + cdbl(rsha("Overtime"))
						subAmtpremi = cdbl(subAmtpremi) + cdbl(rsha("Premi")) 
						subAmtTunj = cdbl(subAmtTunj) + cdbl(rsha("Tunjangan")) 
						subAmtBorong = cdbl(subAmtBorong) + cdbl(rsha("Borongan")) 
						subAmtJamsos = cdbl(subAmtJamsos) + cdbl(rsha("P_Jamsostek")) 
						subAmtMKP1 = cdbl(subAmtMKP1) + cdbl(rsha("P_MK_P1")) 
						subAmtPanen = cdbl(subAmtPanen) + cdbl(rsha("P_Panen")) 
						subAmtPPH21 = cdbl(subAmtPPH21) + cdbl(rsha("P_PPH21")) 
						subAmtPlain = cdbl(subAmtPlain) + cdbl(rsha("P_Lain")) 
						subAmtgrp = cdbl(subAmtgrp) + cdbl(rsha("Total"))
							  
					' cari grand total 
							  
						grnAmtbsal = cdbl(grnAmtbsal) + cdbl(rsha("base_salary")) 
						grnAmtOT = cdbl(grnAmtOT) + cdbl(rsha("Overtime"))
						grnAmtpremi = cdbl(grnAmtpremi) + cdbl(rsha("Premi")) 
						grnAmtTunj = cdbl(grnAmtTunj) + cdbl(rsha("Tunjangan")) 
						grnAmtBorong = cdbl(grnAmtBorong) + cdbl(rsha("Borongan")) 
						grnAmtJamsos = cdbl(grnAmtJamsos) + cdbl(rsha("P_Jamsostek")) 
						grnAmtMKP1 = cdbl(grnAmtMKP1) + cdbl(rsha("P_MK_P1")) 
						grnAmtPanen = cdbl(grnAmtPanen) + cdbl(rsha("P_Panen")) 
						grnAmtPPH21 = cdbl(grnAmtPPH21) + cdbl(rsha("P_PPH21")) 
						grnAmtPlain = cdbl(grnAmtPlain) + cdbl(rsha("P_Lain")) 
						grnAmtgrp = cdbl(grnAmtgrp) + cdbl(rsha("Total"))				 
			 END IF      
			 cdgrp = rsha("Afdeling")
			 rsha.MoveNext
		wend
	end if
	%>	
	    <tr>	 
			<td align="center" colspan="2" ><b>Sub Total <%=cdgrp%></td>
			<td align="right"><b><%=formatnumber(subAmtbsal,0)%></td>
			<td align="right"><b><%=formatnumber(subAmtOT,0)%></td>
			<td align="right"><b><%=formatnumber(subAmtpremi,0)%></td>
			<td align="right"><b><%=formatnumber(subAmtTunj,0)%></td>
			<td align="right"><b><%=formatnumber(subAmtBorong,0)%></td>
			<td align="right"><b><%=formatnumber(subAmtJamsos,0)%></td>
			<td align="right"><b><%=formatnumber(subAmtMKP1,0)%></td>
			<td align="right"><b><%=formatnumber(subAmtPanen,0)%></td>
			<td align="right"><b><%=formatnumber(subAmtPPH21,0)%></td>
			<td align="right"><b><%=formatnumber(subAmtPlain,0)%></td>
			<td align="right"><b><%=formatnumber(subAmtgrp,0)%></td>
		 </tr>	
		 <tr>	 
			<td align="center" colspan="14" ></td>
		 </tr>
        <tr>	 
			<td align="center" colspan="2" ><b>Grand Total</td>
			<td align="right"><b><%=formatnumber(grnAmtbsal,0)%></td>
			<td align="right"><b><%=formatnumber(grnAmtOT,0)%></td>
			<td align="right"><b><%=formatnumber(grnAmtpremi,0)%></td>
			<td align="right"><b><%=formatnumber(grnAmtTunj,0)%></td>
			<td align="right"><b><%=formatnumber(grnAmtBorong,0)%></td>
			<td align="right"><b><%=formatnumber(grnAmtJamsos,0)%></td>
			<td align="right"><b><%=formatnumber(grnAmtMKP1,0)%></td>
			<td align="right"><b><%=formatnumber(grnAmtPanen,0)%></td>
			<td align="right"><b><%=formatnumber(grnAmtPPH21,0)%></td>
			<td align="right"><b><%=formatnumber(grnAmtPlain,0)%></td>
			<td align="right"><b><%=formatnumber(grnAmtgrp,0)%></td>
			<td align="right"></td>
		</tr>				  
</table>
<!--#include file="bottom.Asp"-->