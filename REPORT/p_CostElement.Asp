<!--#include file="../ConndB1.inc"-->
<!--#include virtual="txtFunction.inc"-->
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"
%>
<!--#include file="top1.Asp"-->
<%If trim("vJenisReport") <> "" then%>
  <table width="100%" border="0" cellspacing="0"  cellpadding="0">
  </table>
  <font style="font-size:12px">
  </font>
<table border=0 Cellspacing =1 Cellpadding=1 width="100%" >
    <tr><td align="left">
<!-- cmdTransaction to different of Transaction or table master 
     True = Transaction ; False = table Master
-->
<%cmdTransaction=False%>
<%vJenisReport=trim(Request.Form("vJenisReport"))%>
<span class="style2"></span></td>
<!--#include file="form_CostElement.html"-->
<table border=0 Cellspacing =0 Cellpadding=0 width="100%">
<%
  'Dim rs
  Set rs = Server.CreateObject("ADODB.Recordset")

  If trim(request.Form("vCROPTYPE")) ="1" then
     v_JobcodePanen="4119001"
  ElseIf trim(request.Form("vCROPTYPE")) ="2" then
     v_JobcodePanen="4119002"
  else
     v_JobcodePanen="4119003"
  end if
   
  rs.PageSize = 5
  rs.CacheSize = 5
  rs.CursorLocation = adUseClient
  IF request.Form("vYear")="" THEN
	  strsql =  " SELECT Division,PlantingDate,ACTIVITY_HEADING,ACCOUNTGROUP,sum(TM_AMT_ACTUAL),sum(TM_HA_KG_ACTUAL), " & _
				" max(TM_AMT_BUDGET),max(TM_HA_KG_BUDGET),sum(TY_AMT_ACTUAL),sum(TY_HA_KG_ACTUAL), " & _
				" sum(TY_AMT_BUDGET),sum(TY_HA_KG_BUDGET) from SAC_ACCOUNTREPORT_TM_TBM  " & _
				" where PERIOD_YEAR=2006 and PERIOD_MONTH=1" & _
				" and BA='xxx' and STATUS='TM' "& _
				" group by Division,PlantingDate,ACTIVITY_HEADING,ACCOUNTGROUP "
				 '& _				" order by ACTIVITY_HEADING DESC "
  ELSE			
				strsql1 = "Select " & request.Form("cfield") & _
				",sum(TM_AMT_ACTUAL) as TM_AMT_ACTUAL, sum(TY_AMT_ACTUAL) as TY_AMT_ACTUAL," & _
				"sum(TM_AMT_BUDGET) as TM_AMT_BUDGET, sum(TY_AMT_BUDGET) as TY_AMT_BUDGET," & _
				"sum(AMT_ANNUAL) AS AMT_ANNUAL" & _
				" from SAC_COSTPERSOURCE ,SAC_LOCATIONTYPE " & _
				" WHERE SAC_COSTPERSOURCE.SOURCETYPE = SAC_LOCATIONTYPE.LOCATIONTYPECODE" & _
				" AND CROPTYPE='" & request.Form("vCROPTYPE") & "'" & _
				" and OWNERSHIP in " & "(" & request.Form("vOWNERSHIP") & ")" & _
				" and ACCOUNTGROUPCODE NOT IN'" & "4124000" & "'" & _
				" and Status='" & request.Form("vStatus") & "'" & _
				" and ACCOUNTGROUPCODE NOT IN  (4124000,4123000,4122000,4121000)" & _
				" and PERIOD_YEAR=" & request.Form("vYear")

				'" AND activity_heading Not like '" & "%-%" & "'" & _

				strKG1 = " select sum(TM_KG_ACTUAL) as TM_KG_ACTUAL, sum(TM_KG_BUDGET) as TM_KG_BUDGET," & _
				"sum(TY_KG_ACTUAL) as TY_KG_ACTUAL, sum(TY_KG_BUDGET) as TY_KG_BUDGET," & _
				"sum(KG_ANNUAL) as KG_ANNUAL" & _
				" from SAC_COSTPERSOURCE where  ACCOUNTCODE='" & v_JobcodePanen & "' AND " & _
				" CROPTYPE='" & request.Form("vCROPTYPE") & "'" & _
				" and OWNERSHIP in " & "(" & request.Form("vOWNERSHIP") & ")" & _
				" and Sourcetype='" & "EM" & "'" & _   
				" and Status='" & request.Form("vStatus") & "'" & _
				" and PERIOD_YEAR=" & request.Form("vYear")

			  strHA1 = " select sum(HA) AS HA from SAC_HA_BUDGET_ACTUAL where " & _
			  " CROPTYPE='" & request.Form("vCROPTYPE") & "'" & _
			  " and OWNERSHIP in " & "(" & request.Form("vOWNERSHIP") & ")" & _
              " and Status='" & request.Form("vStatus") & "'" & _
			  " and PERIOD_YEAR=" & request.Form("vYear")

	  if  request.Form("vMonth") <>"ALL" then
	      strsql2 = " and PERIOD_MONTH=" & request.Form("vMonth") 
	  else
	      strsql2 = " "
	  end if
	  if  request.Form("vBA") <>"ALL" then
	      strsql3 = "  and BA='" & request.Form("vBA") & "'" 
	  else
	      strsql3 = " "
	  end if
	  if  request.Form("vOrganization") <>"ALL" then
	      strsql3A = "  and BA in (Select BA from SAC_Organization where GROUP_BA='" & trim(request.Form("vOrganization")) & "')" 
	  else
	      strsql3A= " "
	  end if
	  if  request.Form("vPlantingDate") <>"ALL" then
	      strsql4 = " and PlantingDate='" & request.Form("vPlantingDate") & "'"
	  else
	      strsql4 = " "
	  end if
	  if  request.Form("vDivision") <>"ALL" then
	      strsql5 = " and Division='" & request.Form("vDivision") & "'" 
	  else
	      strsql5 = " "
	  end if
      strsql6= " group by " & request.Form("cfield")
      strsql7= " order by " & request.Form("cfield")
	  strsql = strsql1 + strsql2 + strsql3 + strsql3A + strsql4 + strsql5 + strsql6 + strsql7
	  STRKG  = strKG1 + strsql2 + strsql3 + strsql3A + strsql4 + strsql5
  	  strHA  = strHA1 + strsql2 + strsql3 + strsql3A + strsql4 + strsql5

  END IF

'Response.Write "cfield " &  request.Form("cfield") 
Dim rsHA
Set rsHA = Server.CreateObject("ADODB.Recordset")
rsHA.CursorLocation = adUseClient
nHA = 0
if trim(strHA) <>"" then
   'response.Write strHA
   rsHA.Open strHA, connStr, adOpenForwardOnly, adLockReadOnly
   IF not rsha.eof then
      nHA = rsHA("HA")
   end if
End if

Dim rsKG
Set rsKG = Server.CreateObject("ADODB.Recordset")
rsKG.CursorLocation = adUseClient
nTM_KG_ACTUAL = 0
nTM_KG_BUDGET = 0
nTY_KG_ACTUAL = 0
nTY_KG_BUDGET = 0
nKG_Annual    = 0
if trim(strKG) <>"" and trim(request.Form("vStatus")) ="TM" then
   'RESPONSE.Write STRKG
   rsKG.Open strKG, connStr, adOpenForwardOnly, adLockReadOnly
   IF not rsKG.eof then
      nTM_KG_ACTUAL = rsKG("TM_KG_ACTUAL")
      nTM_KG_BUDGET = rsKG("TM_KG_BUDGET")
      nTY_KG_ACTUAL = rsKG("TY_KG_ACTUAL")
      nTY_KG_BUDGET = rsKG("TY_KG_BUDGET")
	  nKG_Annual    = rsKG("KG_ANNUAL")
	  'response.Write " TM KG actual : " & nTM_KG_ACTUAL & "ok"
   end if
End if
'response.write strsql
rs.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly

If rs.EOF and Len(trim(request("vJenisReport"))) > 0 Then%>
    <tr><td colspan="3">Data is empty</tr></td>
<%elseif not rs.eof and Len(trim(request("vJenisReport"))) > 0 Then
	rs.movefirst
	
	howmanyfields=rs.fields.count-1
	
	JumRec=1
	intRecordsPerPage = 100000000
	Rs.PageSize = intRecordsPerPage
	intNumberOfPages = 0
	intNumberOfPages =  Rs.PageCount
	intRecords = Rs.RecordCount
	' Set our AbsolutePage property.
	If isEmpty(Request("iPageNo")) Then
	   Rs.AbsolutePage = 1
	Else
	   Rs.AbsolutePage = Request("iPageNo")
	End If  
	
	'---- Khusus untuk Judul
	if trim(request.Form("vOwnership")) ="1" then
	 vTextOwnership = "Inti"
	elseif trim(request.Form("vOwnership")) ="2" then 
	 vTextOwnership = "PMBN"
	elseif trim(request.Form("vOwnership")) ="3" then
	 vTextOwnership = "PMU"
	elseif trim(request.Form("vOwnership")) ="4" then	 
	 vTextOwnership = "Plasma"
	elseif trim(request.Form("vOwnership")) ="5" then	 
	 vTextOwnership = "Plasma Pre"
	elseif trim(request.Form("vOwnership")) ="6" then	 
	 vTextOwnership = "KKPA"
	elseif trim(request.Form("vOwnership")) ="7" then	 
	 vTextOwnership = "KKPA Pre"
	elseif trim(request.Form("vOwnership")) ="8" then	 
	 vTextOwnership = "TKD"
	elseif trim(request.Form("vOwnership")) ="1,3,8" then	 
	 vTextOwnership = "Inti+TKD+PMU"
	end if	 
	
	if Trim(Request.Form("vCROPTYPE"))="1" then
	 vTextCroptype="Oil Palm"
	elseif Trim(Request.Form("vCROPTYPE"))="2" then
	 vTextCroptype="Rubber"
	elseif Trim(Request.Form("vCROPTYPE"))="3" then
	 vTextCroptype="CoCoa"
	End if
	'---- end Judul
	%>
	<tr>
	<td width="150" colspan="4"><img src="Images/Logoawal.gif" width="150" height="130"></td>
	<td width="580" colspan="8" align="center"><font style="font-size:12px"><B>
	MULTI INDO OIL<BR>
	ANALISA BIAYA PER COST ELEMENT<br>
	Group : <%=Request.Form("vOrganization")%>, Kebun : <%=Request.Form("vBA")%> <br>
	Period Year :<%=Request.Form("vYear")%> Month : <%=textMonth(trim(Request.Form("vMOnth")))%><br>
	Crop :<%=vTextCroptype%> Ownership : <%=vTextOwnership%>
	</font></div></B></td>
	<td>
    <td align="Right"></td>
	</td>
    </tr>
  	</TABLE>
	<style type="text/css">
	<!--
	.style1 {font-family: Verdana, Arial, Helvetica, sans-serif}
	-->
	</style>
	<table width="200" cellpadding="2" bgcolor="#F7F7F7" cellspacing="0">  
	<%
	if request.Form("vJenisReport")="Group" then
	 fieldnumeric = 3
	elseif request.Form("vJenisReport")="Job" then
	 fieldnumeric = 3
	else
	 fieldnumeric = 3	 
	end if
    Tot_tm_amt_actual = 0
	Tot_ty_amt_actual = 0
	Tot_tm_amt_budget = 0
    Tot_ty_amt_budget = 0
    howmanyfields=rs.fields.count-1		
	jlhcolomdipilih =  rs.fields.count-5  'khusus buat kolom kosong di kiri judul		
	jlhcolomtotal   =  jlhcolomdipilih+6   '6 adalah jumlah baris judul
    %>
	<!--#include file="HEADERPerSource.TXT"-->

	<tr style="<%=cmdstyle%>">
	<%for i = 0 to howmanyfields-5%>
        <%if i = 0 then%>
		  <td valign="top" bgcolor="#C0FBBD" align="left" style="<%=cmdstyle%>;border-Left: 1 solid #0000FF;"><%=rs.fields(i).name%></td>
		<%else%>
		  <td valign="top" bgcolor="#C0FBBD" align="left" style="<%=cmdstyle%>"><%=rs.fields(i).name%></td>
		<%end if%>
	<%next%> 
	<td colspan="7" valign="top" bgcolor="#C0FBBD" align="left" style="<%=cmdstyle%>">&nbsp;</td>
	</tr>
	<%

    While not Rs.EOF and intRecordsPerPage > 0
	  nTM_RP_HA_ACTUAL = 0
	  nTY_RP_HA_ACTUAL = 0
	  nTM_RP_HA_BUDGET = 0
	  nTY_RP_HA_BUDGET = 0
	  nRP_HA_ANNUAL    = 0
	  nPersenTM        = 0
	  nPersenTY        = 0
      nstrPanen=""
	  txtCostElement=""
	  %>
	  </font>
	  <tr style="<%=cmdstyle%>">
	     <%for i = 0 to howmanyfields-5%>
		    <%if i = 0 then%>
	          <td valign="top" bgcolor="#C0FBBD" align="left" style="<%=cmdstyle%>;border-Left: 1 solid #0000FF;"><%=rs(i)%></td>
			<%Else%>
			  <td valign="top" bgcolor="#C0FBBD" align="left" style="<%=cmdstyle%>"><%=rs(i)%></td>
			<%End if%>
		 <%next%> 
		    <%
                        if CDBL(nHA) > 0 then
			   nTMRpPerHA_Actual = cdbl(rs("tm_amt_actual"))/CDBL(nHA)
			   nTMRpPerHA_Budget = cdbl(rs("tm_amt_BUDGET"))/CDBL(nHA)
                        end if
			if nTMRpPerHA_Budget > 0 then
			   nTMPersen = (nTMRpPerHA_Budget-nTMRpPerHA_Actual) / nTMRpPerHA_Budget
			else
			   nTMPersen = 0   
			end if

                        if CDBL(nHA) > 0 then
			nTYRpPerHA_Actual = cdbl(rs("tY_amt_actual"))/CDBL(nHA)
			nTYRpPerHA_Budget = cdbl(rs("tY_amt_BUDGET"))/CDBL(nHA)
                        End if
			if nTYRpPerHA_Budget > 0 then
			   nTYPersen = (nTYRpPerHA_Budget-nTYRpPerHA_Actual) / nTYRpPerHA_Budget
			else
			   nTYPersen = 0   
			end if
                        if CDBL(nHA) > 0 then
			   nTYRpPerHA_Annual = cdbl(rs("amt_annual"))/CDBL(nHA)
                        end if
			%>
	        <td valign="top" bgcolor="#C0FBBD" align="Right" style="<%=cmdstyle%>"><%=FormatNumber(cdbl(rs("tm_amt_actual")),2)%></td>
	        <td valign="top" bgcolor="#C0FBBD" align="Right" style="<%=cmdstyle%>"><%=FormatNumber(cdbl(rs("TM_AMT_BUDGET")),2)%></td>
	        <td valign="top" bgcolor="#C0FBBD" align="Right" style="<%=cmdstyle%>"><%=FormatNumber(cdbl(rs("Ty_AMT_ACTUAL")),2)%></td>
	        <td valign="top" bgcolor="#C0FBBD" align="Right" style="<%=cmdstyle%>"><%=FormatNumber(cdbl(rs("Ty_AMT_BUDGET")),2)%></td>
	        <td valign="top" bgcolor="#C0FBBD" align="Right" style="<%=cmdstyle%>"><%=FormatNumber(cdbl(rs("amt_annual")),2)%></td>
	  </tr>
     <%
     'Headactivity     = rs.fields(1)
     'HeadJobGroup =rs.fields(0)
     'HeadJobGroupDesc =rs.fields(2)
     Tot_tm_amt_actual = Tot_tm_amt_actual + cdbl(rs("tm_amt_actual"))
	 Tot_ty_amt_actual = Tot_ty_amt_actual + cdbl(rs("ty_amt_actual"))
	 Tot_tm_amt_budget = Tot_tm_amt_budget + cdbl(rs("tm_amt_budget"))
	 Tot_ty_amt_budget = Tot_ty_amt_budget + cdbl(rs("ty_amt_budget"))

     Tot_amt_annual   = Tot_amt_annual   + cdbl(rs("amt_annual"))


     '--- End Variable for sub Total	 
	 rs.movenext
	 
     JumRec=JumRec+1
	 intRecordsPerPage = intRecordsPerPage - 1
	 
  Wend%>
	<tr>
	<td valign="top" bgcolor="#C0FBBD" align="Center" style="<%=cmdstyleLeft%>" colspan="<%=jlhcolomdipilih%>%>">
		 <font face="Verdana" size="1"><b>T o t a l</b></td>
	<td valign="top" bgcolor="#C0FBBD" align="Center" style="<%=cmdstyle%>">
		 <font face="Verdana" size="1"><b><%=FormatNumber(Tot_tm_amt_actual,2)%></td>
	<td valign="top" bgcolor="#C0FBBD" align="Center" style="<%=cmdstyle%>">
		 <font face="Verdana" size="1"><b><%=FormatNumber(Tot_tm_amt_budget,2)%></td>
	<td valign="top" bgcolor="#C0FBBD" align="Center" style="<%=cmdstyle%>">
		 <font face="Verdana" size="1"><b><%=FormatNumber(Tot_ty_amt_actual,2)%></td>
	<td valign="top" bgcolor="#C0FBBD" align="Center" style="<%=cmdstyle%>">
		 <font face="Verdana" size="1"><b><%=FormatNumber(Tot_ty_amt_budget,2)%></td>
	<td valign="top" bgcolor="#C0FBBD" align="Center" style="<%=cmdstyle%>">
		 <font face="Verdana" size="1"><b><%=FormatNumber(Tot_amt_annual,2)%></td>
	</Table></center>
	   <font size="-4">
	   <%'response.Write(strsql)%></font><%
	%>
	</font>
<%End if%>

<%End if%>
<!--#include file="bottom.Asp"-->