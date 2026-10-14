<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/ConndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%

Dim v_txtDivision, v_txtPlantingDate, v_txtBlok
Dim v_txtyear,v_txtmonthto,v_txtmonthfrom
Dim txtDivision, txtBlok, txtPlantingDate
Dim v_vstatus,v_txtjob,v_txtjobgroup
Dim tempyear,tempmonth,no
Dim tempDivision, tempBlok, tempPlantingDate 
Dim SprevCode,Sprevjobname, Sprevblock
Dim sprevPlantingYear, sprevFieldCode
Dim tempjob
Dim MM,MF, yyyy, mon, Yr, MinYr, MaxYr, i, rsBa,strsql1,strsql2,strsql3,strsql4,strsql5,strsql6,strsql7,strsql8
Dim ScriptName,nSpanTotal
DIM q,FcbaName,CompanyName,MgrName,KTUName,CropName,FcbaType,rsJob

Dim rscurrentdate, strcurrentdate, varDefaultYear, varDefaultMonth
Dim str1,str2,str3,strMain,whereCons,whereCons1,selectCons,orderCons
Dim var_ColSpan : var_ColSpan = 0

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")
           
   		  
			v_txtfcba 			  = Request("txtfcba")
			v_txtPlantingDate	  = Request("txtPlantingDate")
			v_txtBlok			  = Request("txtBlok")
			v_txtyear             = Request("txtYear")
			v_txtmonthto          = Request("txtMonthto")
			v_txtmonthfrom        = Request("txtMonthfrom")
			v_vstatus             = Request("vStatus")
			v_txtjob			  = Request("txtjob")
			v_txtjobgroup		  = Request("txtjobgroup")
			
			nSpanTotal = 9
			
			strcurrentdate = " select currentdate" & _
                 " from periodcontrol" & _                 
                 " where fcba = '" & v_txtFcba & "'"

			Set rscurrentdate = Server.CreateObject("ADODB.Recordset")
			rscurrentdate.Open strcurrentdate, connstr, adOpenStatic

			If Not rscurrentdate.EOF Then
				varDefaultYear	=	Year(rscurrentdate("currentdate"))
				varDefaultMonth	=	Month(rscurrentdate("currentdate"))
			Else
				varDefaultYear	=	Year(Session("v_Employee"))
				varDefaultMonth	=	Month(Session("v_Employee"))
			End If
			rscurrentdate.Close
			set rscurrentdate = Nothing

			If (v_txtYear = "" Or Len(Trim(v_txtYear)) = 0) Then v_txtYear = varDefaultYear
			If (v_txtMonthto = "" Or Len(Trim(v_txtMonthto)) = 0) Then v_txtMonthto = varDefaultMonth
			If (v_txtMonthfrom = "" Or Len(Trim(v_txtMonthfrom)) = 0) Then v_txtMonthfrom = varDefaultMonth
	        
			
            
			if  v_txtPlantingDate <> "--SUMMARY--" and v_txtPlantingDate <> ""  then
				   if v_txtPlantingDate <>  "--ALL--" then
					  whereCons = whereCons & " AND PLANTINGYEAR = '" & v_txtplantingdate & "'"
				   end if
				   selectCons = selectCons &  " , PLANTINGYEAR "
				   orderCons  = orderCons & " , PLANTINGYEAR "
			else
			      nSpanTotal = nSpanTotal - 1 	  
			end if
				
		     if v_txtBlok <>  "--ALL--" and v_txtBlok <>  "" then
				whereCons1 = whereCons1 & " AND targetcode = '" & v_txtBlok & "'"
			end if
			
		    if v_txtjob <>  "--ALL--" and v_txtBlok <>  "" then
				whereCons1 = whereCons1 & " AND jobcode = '" & v_txtjob & "'"
			end if

strMain		=	" select " & _
				"    fcba,jobcode,block,plantingyear,status,amount from " & _
				" ( select fcba,jobcode, " & _
				" targetcode block,field_api.getValid_plantingyear(targetcode,fcba,tdate) plantingyear, " & _
				" field_api.GetValid_status(targetcode,fcba,tdate) status, amount " & _
                " from cost_activity a where to_number(to_char(tdate, 'MM')) between " & v_txtmonthfrom & " and " & v_txtmonthto & " and to_number(to_char(tdate, 'YYYY')) = " & v_txtyear & " " & _
 				" and fcba= '" & v_txtFcba & "' and targettype='FF' and jobcode not in ('1329901','1330200') " & whereCons1 & " ) " & _
				" where status= '" & v_vstatus & "' "  & whereCons  & ""
								
str1	=	"  fcba,jobcode,block,plantingyear,status"
							
str3    =   " ,job_api.get_jobcode_name(jobcode, fcba) jobname,job_api.get_jobgroupCode(jobcode, fcba) jobgroup,jobgroup_api.get_jobgroup_name(job_api.get_jobgroupCode(jobcode,fcba),fcba) jobgroupname,nvl(field_api.GetValid_hectarageplanted(block,fcba,sysdate),0) haplanted,field_api.getValid_Name(block,fcba,sysdate) org_block" 							

str2	=	" select " & str1 & " " & str3 & " , sum(amount) amount From (" & strMain & " ) Group By " & str1  & " order by plantingyear,block,jobcode"
							

	   
%>			

<html>

<head>
<title>iPla'S Reporting - Realisasi Biaya Per Block</title>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>
<script language="JavaScript" src='/js/x.js'></script>
<script language="JavaScript" src='/js/xtableheaderfixed.js'></script>
<script language="JavaScript">
xAddEventListener(window, 'load',
  function() {
    new xTableHeaderFixed('xthf-site', window);
    new xTableHeaderFixed('xthf-green', window);
    new xTableHeaderFixed('xthf-blue', window);
  }, false
);
</script>
 <script>
$(function() {
$( ".datepicker" ).datepicker();
});
</script>
<!--<script language="JavaScript" src="/js/date-picker.js"></script> 
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>-->
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
	<!--#include file="../print_excel.asp"-->
	<div id="header-wrapper">
		<div class="printhide">
		
		<form action="<%=ScriptName%>" method="post">
		<!--#include file="menutop.asp"-->
		<!--#include virtual="/REPORT/FormSelectFcbaOnly.html"-->
		<table id="form-table" class="centered-margin">
			<tr>
				<th>TAHUN</th>
				<th>DARI BULAN</th>
				<th>S/D BULAN</th>
				<th>STATUS</th>
				<th>JOBGROUP</th>
				<th>KODE</th>
				<th>TAHUN TANAM</th>
				<th>BLOCK</th>
				<th>&nbsp;</th>
			</tr>
			<tr>

					<td>
						<font size="2"  color="#0000FF" ></font>
						<% 
							MF=v_txtMonthfrom
							MM=v_txtMonthto
							yyyy=v_txtYear

							mon = Month(Session("v_Costbook"))
							If len(MM)=0 then MM=mon

							Yr = Year(Session("v_Costbook"))
							MinYr = Yr - 2
							MaxYr = Yr  + 2
							If len(yyyy)= 0 then yyyy = Yr
						%>
						<select name="txtYear">
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
							<select name="txtMonthfrom">
								<%
									For i = 1 to 12
								%>
										<Option Value="<%=i%>" <%If cint(i) = cint(MF) then%> Selected <%End If%>><%=MonthName(i)%></option>
								<%
									Next
								%>
							</select>
					</td>		

			
					<td>
						<font size="2" color="#0000FF" ></font>
							<select name="txtMonthto">
								<%
									For i = 1 to 12
								%>
										<Option Value="<%=i%>" <%If cint(i) = cint(MM) then%> Selected <%End If%>><%=MonthName(i)%></option>
								<%
									Next
								%>
							</select>
					</td>		

					<td>
						<font size="2"  color="#0000FF" ></font>
							<select name="vStatus">
								<option value="TM" <%if request.Form("vStatus")="TM" Then%> Selected <%End if%>>TM</option>
								<option value="TBM" <%if request.Form("vStatus")="TBM" Then%> Selected <%End if%>>TBM</option>
							</select>
					</td>
					<td>
						<select name="txtJobgroup">
							<option value="--ALL--" <%if request("v_txtJob")="--ALL--" then%> selected <%End if%>>--ALL--</option>
							<%
							
							Set rsJob = Server.CreateObject("ADODB.Recordset")
							rsJob.CursorLocation = adUseClient
							strsql =   " select distinct a.jobgroupcode fccode, jobgroup_api.get_jobgroup_name(a.jobgroupcode,a.fcba) fcname from job a where a.JOB_FORFIELDSTATUS in ('TM', 'TBM') and a.fccode not in '1329901' order by a.jobgroupcode "
							rsJob.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
							if not rsJob.eof then
							DO while not rsJob.eof
							  if v_txtJob= rsJob("FCCODE") then
								%>			
								<option value="<%=rsJob("FCCODE")%>" Selected><%=rsJob("FCCODE")%> - <%=rsJob("fcname")%></option>
						   <% Else %>	
								<option value="<%=rsJob("FCCODE")%>"><%=rsJob("FCCODE")%> - <%=rsJob("fcname")%></option>
						   <% End If
							  rsJob.movenext
							loop
							End if
							rsJob.Close
							set rsJob = Nothing
							%>
							
						</select>
					</td>
					<td>
						<select name="txtJob">
							<option value="--ALL--" <%if request("v_txtJob")="--ALL--" then%> selected <%End if%>>--ALL--</option>
							<%
							
							Set rsJob = Server.CreateObject("ADODB.Recordset")
							rsJob.CursorLocation = adUseClient
							strsql =  " select distinct  a.fccode,a.fcname from job a where a.JOB_FORFIELDSTATUS in ('TM', 'TBM') and a.fccode not in '1329901' order by a.fccode"
							rsJob.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
							if not rsJob.eof then
							DO while not rsJob.eof
							  if v_txtJob= rsJob("FCCODE") then
								%>		
								<option value="<%=rsJob("FCCODE")%>" Selected><%=rsJob("FCCODE")%> - <%=rsJob("fcname")%></option>
						   <% Else %>	
								<option value="<%=rsJob("FCCODE")%>"><%=rsJob("FCCODE")%> - <%=rsJob("fcname")%></option>
						   <% End If
							  rsJob.movenext
							loop
							End if
							rsJob.Close
							set rsJob = Nothing
							%>
							
						</select>
			        </td>
					
					<td align="center">
						<select name="txtPlantingDate">	
								<option value="--SUMMARY--">--SUMMARY--</option>
								<%
									Dim rsPlantingDate
									Set rsPlantingDate = Server.CreateObject("ADODB.Recordset")
									rsPlantingDate.CursorLocation = adUseClient
									
							
							
									if len(trim(v_txtPlantingDate)) > 0 then 
										strsql =  " SELECT '--ALL--' as PlantingDate from DUAL UNION SELECT DISTINCT to_char(PLANTINGDATE, 'yyyy') as PlantingDate FROM FIELDDETAIL WHERE FCBA = '" & v_txtFcba & "'"
									else
										strsql =  " SELECT '--ALL--' as PlantingDate from DUAL UNION SELECT DISTINCT to_char(PLANTINGDATE, 'yyyy') as PlantingDate FROM FIELDDETAIL WHERE FCBA = '" & Session("fcBA") & "'"
									end if
									
									
									rsPlantingDate.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
									if not rsPlantingDate.eof then
										DO while not rsPlantingDate.eof
										if v_txtPlantingDate = rsPlantingDate("PlantingDate") then
								%>			<option value="<%=rsPlantingDate("PlantingDate")%>" Selected><%=rsPlantingDate("PlantingDate")%></option>
										<%else%>
											<option value="<%=rsPlantingDate("PlantingDate")%>"><%=rsPlantingDate("PlantingDate")%></option>
										<%end if%>
								<%			rsPlantingDate.movenext
										loop
									End if
									rsPlantingDate.Close
									set rsPlantingDate = nothing
								%>
						</select>		
					</td>
					
					
					<td align="center">
						<select name="txtBlok">	
							<option value="--ALL--">--ALL--</option>
							<%
								Dim rsBlok
								Set rsBlok = Server.CreateObject("ADODB.Recordset")
								rsBlok.CursorLocation = adUseClient

						
								if len(trim(v_txtBlok)) > 0 then 
									strsql =  " SELECT DISTINCT fccode as Blok,fcname FROM FIELDDETAIL WHERE FCBA = '" & v_txtFcba & "'  order by fccode"
								else
									strsql =  " SELECT DISTINCT fccode as Blok,fcname FROM FIELDDETAIL WHERE FCBA = '" & Session("fcBA") & "'  order by fccode"
								end if
								
								rsBlok.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
								if not rsBlok.eof then
									DO while not rsBlok.eof
									if v_txtBlok = rsBlok("Blok") then
							%>			<option value="<%=rsBlok("Blok")%>" Selected><%=rsBlok("Blok")%> - <%=rsBlok("fcname")%> </option>
									<%else%>
										<option value="<%=rsBlok("Blok")%>"><%=rsBlok("Blok")%> - <%=rsBlok("fcname")%> </option>
									<%end if%>
							<%			rsBlok.movenext
									loop
								End if
								rsBlok.Close
								set rsBlok = nothing
							%>
						</select>		
					</td>
					
					<td><input type="submit" value="View" /></td>
			</tr>		
		</table>	
		</form>
		</div>
     </div>   
<table border="0" id="product-table" class="xthf-site" width="100%">
	<thead>
        <tr>
          <th colspan="19" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">	    
		<div id="header">
		   <% if  CompanyName <>  "" then %>
				<h4>PT.<%=CompanyName%></h4> 
			<% End if%>
			<% If  FcbaName <> ""  then %>
				   <h5><%=FcbaName%></h5>   
			<% End if%>
			
		</div>
		<div id="subheader" class="centered">
			<% Dim ia,ib,thntnm
			   ia = v_txtmonthto
			   if ia ="" then
				  ia="0" 
			   end if
			   ib = v_txtmonthfrom
			   if ib ="" then
				  ib="0" 
			   end if
			   
			   if   v_txtPlantingDate = "0"  then
					thntnm = "ALL"
			   else		
					thntnm = v_txtPlantingDate
			   end if
					
			%>
			<%  IF v_vstatus  = "TM" THEN %>		
			   <h3>LAPORAN REALISASI BIAYA DETAIL PER BLOCK (TM) </h3>
			<%  ELSE %>		
			   <h3>LAPORAN REALISASI BIAYA DETAIL PER BLOCK  (TBM) </h3>   
			<%  END IF %>	   
			<% IF v_txtmonthto <> "" then %>
			<h5>PERIODE     : <%=ucase(MonthName(cint(ib)))%> S/d <%=ucase(MonthName(cint(ia)))%> &nbsp; <%=v_txtyear%></h5>
			<% END IF %>
			<h2> &nbsp; &nbsp; &nbsp; </h2>
		</div>
        </th>
    </tr>
	<tr>
	    
		<th style="background-color:#00A300;color:#fff;">No</th>
		<th style="background-color:#00A300;color:#fff;">Job Group</th>
		<th style="background-color:#00A300;color:#fff;">Nama Job Group</th>
		<th style="background-color:#00A300;color:#fff;">Job Kode</th>
		<th style="background-color:#00A300;color:#fff;">Nama Job kode</th>
        <th style="background-color:#00A300;color:#fff;">Blok IPLAS</th>
		<th style="background-color:#00A300;color:#fff;">Blok Asli</th>
		<th style="background-color:#00A300;color:#fff;">Tahun Tnm</th>
		<th style="background-color:#00A300;color:#fff;">Status</th>
		<th style="background-color:#00A300;color:#fff;">Hektar Planted</th>
		<th style="background-color:#00A300;color:#fff;">Amount</th>
	</tr>
 </thead>
 
    <tbody>
		<% 
		'response.write str2
		dim rsreal,cClass,subtotal
		Set rsreal = Server.CreateObject("ADODB.Recordset")
		rsreal.Open str2, connstr, adOpenStatic

		Dim Rp_Per_Ha_RowReal : Rp_Per_Ha_RowReal = 0
		Dim Rp_Per_Ha_RowBudget : Rp_Per_Ha_RowBudget = 0
		Dim RowPercentage : RowPercentage = 0



		if rsreal.bof and rsreal.eof then
		%>
			<tr>
				<td colspan="<%=var_colspan+12%>">NO DATA AVAILABLE</td>
			</tr>
		<%
		else
				No=0
				Sprevblock=""
				while not rsreal.eof
					if no  mod 2 = 1 then
						cClass = "odd"
					else
						cClass = "default"
					end if
					No = No + 1 
		%>
						
						 <% if Sprevblock = rsreal("block") then
						       subtotal = cdbl(subtotal) +  cdbl(rsreal("amount"))
							   cClass = "default"
							   %>	
							
						 <% else
								cClass = "default" 
                                if 	Sprevblock <> "" then 						
						 %>
									<tr class="<%=cClass%>">   
									   <td colspan="10" align="center">Total Biaya Block <%=Sprevblock%></td>
									   <td align="Right"><%=FormatNumber(subtotal,0)%></td>
									</tr>   
							 <% end if
							    subtotal = 0 %>
						<% end if	%>
								
						<tr class="<%=cClass%>">	
							    <td align="center"><%=No%></td>
							    <td align="center"><%=rsreal("jobgroup")%></td>
							    <td align="left"><%=rsreal("jobgroupname")%></td>
							    <td align="center"><%=rsreal("jobcode")%></td>
							    <td align="left"><%=rsreal("jobname")%></td>
							    <td align="center"><%=rsreal("block")%></td>
							    <td align="center"><%=rsreal("org_block")%></td>
							    <td align="center"><%=rsreal("plantingyear")%></td>
							    <td align="center"><%=rsreal("status")%></td>
						        <td align="Right"><%=FormatNumber(rsreal("haplanted"),0)%></td>
							    <td align="Right"><%=FormatNumber(rsreal("amount"),0)%></td>	
						</tr>
		<%             Sprevblock= rsreal("block")
					rsreal.MoveNext
					Wend	
		rsreal.Close
		set rsreal = Nothing
		%>
		                        <tr class="<%=cClass%>">   
								   <td colspan="10" align="center">Total Biaya Block <%=Sprevblock%></td>
								   <td align="Right"><%=FormatNumber(subtotal,0)%></td>
								</tr>   
		<%End if%>
	</tbody>	
</table>
<!--#include file="bottom.Asp"-->
