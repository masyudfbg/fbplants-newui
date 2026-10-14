<!--#include file="../connDb1.Inc"-->
<!--#include file="../txtFunction.inc"-->
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
  <table width="100%" border="0" cellspacing="0"  cellpadding="0">
  </table>
  <font style="font-size:12px">
  </font>
<table border=0 Cellspacing =1 Cellpadding=1 width="100%">
    <tr><td align="left">
<!-- cmdTransaction to different of Transaction or table master 
     True = Transaction ; False = table Master
-->
<%cmdTransaction=False%>
<span class="style2"></span></td>

</tr>
  <center>
  <tr align="Left" valign="middle"><td valign="middle" align="center"><form name="form1" method="post">
      <label><font size="2"  color="#0000FF" ></font>
        <select name="vYear">
          <option value="2006" selected="selected">2006</option>
          <option value="2007">2007</option>
        </select>
        </label>
      <label><font size="2" color="#0000FF" ></font>
        <select name="vMonth">
		  <%if request.Form("vMonth")="1" then%>
             <option value="1" selected="selected">January</option>
		  <%else%>
             <option value="1">January</option>
		  <%End if%>
		  <%if request.Form("vMonth")="2" then%>
             <option value="2" selected="selected">February</option>
		  <%else%>
             <option value="2">February</option>
		  <%End if%>
		  <%if request.Form("vMonth")="3" then%>
             <option value="3" selected="selected">March</option>
		  <%else%>
             <option value="3">March</option>
		  <%End if%>
		  <%if request.Form("vMonth")="4" then%>
             <option value="4" selected="selected">April</option>
		  <%else%>
             <option value="4">April</option>
		  <%End if%>
		  <%if request.Form("vMonth")="5" then%>
             <option value="5" selected="selected">May</option>
		  <%else%>
             <option value="5">May</option>
		  <%End if%>
		  <%if request.Form("vMonth")="6" then%>
             <option value="6" selected="selected">June</option>
		  <%else%>
             <option value="6">June</option>
		  <%End if%>
		  <%if request.Form("vMonth")="7" then%>
             <option value="7" selected="selected">July</option>
		  <%else%>
             <option value="7">July</option>
		  <%End if%>
		  <%if request.Form("vMonth")="8" then%>
             <option value="8" selected="selected">August</option>
		  <%else%>
             <option value="8">August</option>
		  <%End if%>
		  <%if request.Form("vMonth")="9" then%>
             <option value="9" selected="selected">September</option>
		  <%else%>
             <option value="9">September</option>
		  <%End if%>
		  <%if request.Form("vMonth")="10" then%>
             <option value="10" selected="selected">October</option>
		  <%else%>
             <option value="10">October</option>
		  <%End if%>
		  <%if request.Form("vMonth")="11" then%>
             <option value="11" selected="selected">November</option>
		  <%else%>
             <option value="11">November</option>
		  <%End if%>
		  <%if request.Form("vMonth")="12" then%>
             <option value="12" selected="selected">December</option>
		  <%else%>
             <option value="12">December</option>
		  <%End if%>
		     <option value="ALL">::ALL::</option>
        </select>
        </label>		
      <label><font size="2" color="#0000FF" ></font>
        <select name="vBA">
	     <%IF request.Form("vBA")="MILL" THEN%>
			<option value="MILL" selected="selected">::ALL Mill::</option>
		 <%ELSE%>
			<option value="MILL">::ALL Mill::</option>
		 <%END IF%>
	     <%IF request.Form("vBA")="ESTATE" THEN%>
			<option value="ESTATE" selected="selected">::ALL Estate::</option>
		 <%ELSE%>
			<option value="ESTATE">::ALL Estate::</option>
		 <%END IF%>

	     <%IF request.Form("vBA")="KPR" THEN%>
			<option value="KPR" selected="selected">KPR</option>
		 <%ELSE%>
			<option value="KPR">KPR</option>
		 <%END IF%>
	     <%IF request.Form("vBA")="KSC" THEN%>
			<option value="KSC" selected="selected">KSC</option>
		 <%ELSE%>
			<option value="KSC">CDS</option>
		 <%END IF%>


        </select>
        </label>
      <label><font size="2"  color="#0000FF" ></font>
        <select name="vJenisReport">
		 <%if request.Form("vJenisReport")="STAFF" Then%>
          <option value="ALL">Per Activity Group</option>
          <option value="ALLGOL">GC Per Group</option>
          <option value="STAFF"  Selected>GC-Staff</option>
          <option value="PEGAWAI-KANTOR">GC-Kantor</option>
          <option value="KEAMANAN">GC-Keamanan</option>
          <option value="BANGUNAN">GC-Bangunan</option>
          <option value="KONTRAK">GC-Kontraktor</option>
		 <%elseif request.Form("vJenisReport")="PEGAWAI-KANTOR" Then%>
          <option value="ALL">Per Activity Group</option>
          <option value="ALLGOL">GC Per Group</option>
          <option value="STAFF"  Selected>GC-Staff</option>
          <option value="PEGAWAI-KANTOR">GC-Kantor</option>
          <option value="KEAMANAN">GC-Keamanan</option>
          <option value="BANGUNAN">GC-Bangunan</option>
          <option value="KONTRAK">GC-Kontraktor</option>
		 <%elseif request.Form("vJenisReport")="KEAMANAN" Then%>
          <option value="ALL">Per Activity Group</option>
          <option value="ALLGOL">GC Per Group</option>
          <option value="STAFF"  Selected>GC-Staff</option>
          <option value="PEGAWAI-KANTOR">GC-Kantor</option>
          <option value="KEAMANAN">GC-Keamanan</option>
          <option value="BANGUNAN">GC-Bangunan</option>
          <option value="KONTRAK">GC-Kontraktor</option>
		 <%elseif request.Form("vJenisReport")="BANGUNAN" Then%>
          <option value="ALL">Per Activity Group</option>
          <option value="ALLGOL">GC Per Group</option>
          <option value="STAFF"  Selected>GC-Staff</option>
          <option value="PEGAWAI-KANTOR">GC-Kantor</option>
          <option value="KEAMANAN">GC-Keamanan</option>
          <option value="BANGUNAN">GC-Bangunan</option>
          <option value="KONTRAK">GC-Kontraktor</option>
		 <%elseif request.Form("vJenisReport")="ALL" Then%>
          <option value="ALL">Per Activity Group</option>
          <option value="ALLGOL">GC Per Group</option>
          <option value="STAFF"  Selected>GC-Staff</option>
          <option value="PEGAWAI-KANTOR">GC-Kantor</option>
          <option value="KEAMANAN">GC-Keamanan</option>
          <option value="BANGUNAN">GC-Bangunan</option>
          <option value="KONTRAK">GC-Kontraktor</option>
		 <%elseif request.Form("vJenisReport")="KONTRAK" Then%>
          <option value="ALL">Per Activity Group</option>
          <option value="ALLGOL">GC Per Group</option>
          <option value="STAFF"  Selected>GC-Staff</option>
          <option value="PEGAWAI-KANTOR">GC-Kantor</option>
          <option value="KEAMANAN">GC-Keamanan</option>
          <option value="BANGUNAN">GC-Bangunan</option>
          <option value="KONTRAK">GC-Kontraktor</option>
		 <%elseif request.Form("vJenisReport")="ALLGOL" Then%>
          <option value="ALL">Per Activity Group</option>
          <option value="ALLGOL">GC Per Group</option>
          <option value="STAFF"  Selected>GC-Staff</option>
          <option value="PEGAWAI-KANTOR">GC-Kantor</option>
          <option value="KEAMANAN">GC-Keamanan</option>
          <option value="BANGUNAN">GC-Bangunan</option>
          <option value="KONTRAK">GC-Kontraktor</option>
		 <%else %>
          <option value="ALL">Per Activity Group</option>
          <option value="ALLGOL">GC Per Group</option>
          <option value="STAFF"  Selected>GC-Staff</option>
          <option value="PEGAWAI-KANTOR">GC-Kantor</option>
          <option value="KEAMANAN">GC-Keamanan</option>
          <option value="BANGUNAN">GC-Bangunan</option>
          <option value="KONTRAK">GC-Kontraktor</option>
		  <%End if%>
        </select>
        </label>
		<input type="submit" value="View" >
    </form>
    </td>
  </tr>
  </table>
  
<%
response.Write "Jenis Report : " & trim(Request.Form("vJenisReport"))
IF trim(Request.Form("vJenisReport")) <> "" Then
	' Update Transfer in GC Allocation AA_BTL 
	' Dari Table AA_AccountReport_Tm_TM
	Dim vNumeric
	if request.Form("vBa") <> "MILL" and request.Form("vBa") <> "ESTATE" Then
	   for vNumeric = 1 to cdbl(request.Form("vMOnth")) 
		if mid(trim(request.Form("vBa")),1,1) = "K" then
		strUpdate = " UPDATE AA_BTL A " & _
				" SET (A.TM_AMT_ACTUAL)= nvl((SELECT -1*SUM(TM_AMT_ACTUAL) " & _
				"                         FROM AA_ACCOUNTREPORT_TM_TBM WHERE BA='" & request.Form("vBA") & "'" & _
				"                         AND PERIOD_MONTH='" & vNumeric & "' AND " & _
				"                         PERIOD_YEAR='" & request.Form("vYear") & "' AND ACCOUNTCODE='" & "4601088" & "'),0)," & _
				" (A.TY_AMT_ACTUAL)=nvl((SELECT -1*SUM(TY_AMT_ACTUAL) FROM AA_ACCOUNTREPORT_TM_TBM " & _
				"                         WHERE BA='" &request.Form("vBA") & "'" &_
				"                         AND PERIOD_MONTH='" & vNumeric & "' AND " & _
				"  					      PERIOD_YEAR='" & request.Form("vYear") &  "' AND ACCOUNTCODE='" & "4601088" & "'),0)" & _									
				" WHERE BA='" &request.Form("vBA")& "' AND PERIOD_YEAR='" & request.Form("vYear") & "' " & _
				" AND PERIOD_MONTH='" & vNumeric & "'" & _
				" AND ACCOUNTCODE='" & "4601088" & "'"
		else
		strUpdate = " UPDATE AA_BTL A " & _
				" SET (A.TM_AMT_ACTUAL)= nvl((SELECT -1*SUM(TM_AMT_ACTUAL) " & _
				"                         FROM AA_ACCOUNTREPORT_MILL WHERE BA='" & request.Form("vBA") & "'" & _
				"                         AND PERIOD_MONTH='" & vNumeric & "' AND " & _
				"                         PERIOD_YEAR='" & request.Form("vYear") & "' AND ACCOUNTCODE='" & "4601088" & "'),0)," & _
				" (A.TY_AMT_ACTUAL)=nvl((SELECT -1*SUM(TY_AMT_ACTUAL) FROM AA_ACCOUNTREPORT_MILL " & _
				"                         WHERE BA='" &request.Form("vBA") & "'" &_
				"                         AND PERIOD_MONTH='" & vNumeric & "' AND " & _
				"  					      PERIOD_YEAR='" & request.Form("vYear") &  "' AND ACCOUNTCODE='" & "4601088" & "'),0)" & _									
				" WHERE BA='" &request.Form("vBA")& "' AND PERIOD_YEAR='" & request.Form("vYear") & "' " & _
				" AND PERIOD_MONTH='" & vNumeric & "'" & _
				" AND ACCOUNTCODE='" & "4601088" & "'"
		End if		
		'response.Write strUpdate & "<br><BR>"		
		connStr.execute(strUpdate)
		next
	end if	
	connstr.execute("commit")
	' Akhir Update AA_BTL Transfer in GC
    '--------------
	%>

<%If request.Form("vJenisReport") <>"" then%>
<table Cellspacing =1 Cellpadding=1 width="100%">
<%
  'Dim rs
  Set rs = Server.CreateObject("ADODB.Recordset")

  rs.PageSize = 5
  rs.CacheSize = 5
  rs.CursorLocation = adUseClient
  IF request.Form("vYear")="" THEN
	strsql = "SELECT STATIONCODE_DESCRIPTION,MACHINECODE_DESCRIPTION,ACCOUNTCODE_DESCRIPTION," & _
		     "ACCOUNTCODE,SOURCETYPE,sum(TM_AMT_ACTUAL),max(TM_FFB_ACTUAL), " & _
			" SUM(TM_AMT_BUDGET),max(TM_FFB_BUDGET),sum(TY_AMT_ACTUAL),sum(TY_FFB_ACTUAL), " & _
			" max(TY_AMT_BUDGET),max(TY_FFB_BUDGET) from AA_ACCOUNTREPORT_MILL  WHERE " & _
			" CROPTYPE='" & "1" & "' AND " & _
			" BA='" & "xxx" & "' AND " & _
			" PERIOD_YEAR=" & "2006" & _
			" group by STATIONCODE_DESCRIPTION,MACHINECODE_DESCRIPTION,ACCOUNTCODE,ACCOUNTCODE_DESCRIPTION,SOURCETYPE"
  ELSE			
      If request.Form("vJenisReport") <>"BANGUNAN" then
	     if request.Form("vJenisReport") <>"ALLGOL"  then
		    strsql1 = "SELECT ACCOUNTCODE,DESCRIPTION," & _
					" SUM(TM_AMT_ACTUAL) AS TM_AMT_ACTUAL,SUM(TM_AMT_BUDGET) AS TM_AMT_BUDGET, " & _
					" SUM(TY_AMT_ACTUAL) AS TY_AMT_ACTUAL,SUM(TY_AMT_BUDGET) AS TY_AMT_BUDGET, " & _
					" SUM(AMT_ANNUAL) AS AMT_ANNUAL " & _
					" from AA_BTL  WHERE" & _
					" PERIOD_YEAR=" & request.Form("vYear")
            strsql5= " group by ACCOUNTCODE,DESCRIPTION" 
		    strHA1 = " select sum(HA) AS HA from AA_HA_BUDGET_ACTUAL where " & _
		           "  PERIOD_YEAR=" & request.Form("vYear")
		  else
		    strsql1 = "SELECT ACCOUNTCODE,DESCRIPTION,GCGROUP," & _
					" SUM(TM_AMT_ACTUAL) AS TM_AMT_ACTUAL,SUM(TM_AMT_BUDGET) AS TM_AMT_BUDGET, " & _
					" SUM(TY_AMT_ACTUAL) AS TY_AMT_ACTUAL,SUM(TY_AMT_BUDGET) AS TY_AMT_BUDGET, " & _
					" SUM(AMT_ANNUAL) AS AMT_ANNUAL " & _
					" from AA_BTL  WHERE" & _
					" PERIOD_YEAR=" & request.Form("vYear")
            strsql5= " group by ACCOUNTCODE,DESCRIPTION,GCGROUP" 
		    strHA1 = " select sum(HA) AS HA from AA_HA_BUDGET_ACTUAL where " & _
		           "  PERIOD_YEAR=" & request.Form("vYear")
		  end if
	  ELSE
		  strsql1 = "SELECT GCCODE,GCDESCRIPTION," & _
				" SUM(TM_AMT_ACTUAL) AS TM_AMT_ACTUAL,SUM(TM_AMT_BUDGET) AS TM_AMT_BUDGET, " & _
				" SUM(TY_AMT_ACTUAL) AS TY_AMT_ACTUAL,SUM(TY_AMT_BUDGET) AS TY_AMT_BUDGET, " & _
				" SUM(AMT_ANNUAL) AS AMT_ANNUAL " & _
				" from AA_BTL  WHERE" & _
				" PERIOD_YEAR=" & request.Form("vYear")
		  strsql5= " group by GCCODE,GCDESCRIPTION" 			
		  strHA1 = " select sum(HA) AS HA from AA_HA_BUDGET_ACTUAL where " & _
		           " PERIOD_YEAR=" & request.Form("vYear")
      End if
	  strMill1=" select max(TM_CPO_BUDGET) as TM_CPO_BUDGET,max(TM_KER_BUDGET) as TM_KER_BUDGET, " & _
	           " max(TM_FFB_BUDGET) as TM_FFB_BUDGET, " & _
			   " max(TY_CPO_BUDGET) as TY_CPO_BUDGET,max(TY_KER_BUDGET) as TY_KER_BUDGET, " & _
			   " max(TY_FFB_BUDGET) as TY_FFB_BUDGET, " & _
			   " max(TM_FFB_ACTUAL) as TM_FFB_ACTUAL,max(TM_CPO_ACTUAL) as TM_CPO_ACTUAL," & _
			   " max(TM_KER_ACTUAL) as TM_KER_ACTUAL, " & _
			   " max(TY_FFB_ACTUAL) as TY_FFB_ACTUAL,max(TY_CPO_ACTUAL) as TY_CPO_ACTUAL," & _
			   " max(TY_KER_ACTUAL) as TY_KER_ACTUAL," & _
			   " max(FFB_ANNUAL) as FFB_ANNUAL,max(CPO_ANNUAL) as CPO_ANNUAL,max(KER_ANNUAL) as KER_ANNUAL" & _
			   " from AA_ACCOUNTREPORT_MILL where " & _
			   " PERIOD_YEAR=" & request.Form("vYear")

	  if  request.Form("vJenisReport") <>"ALL" and request.Form("vJenisReport") <>"ALLGOL" then
	      strsql2 = " and GCGROUP='" & TRIM(request.Form("vJenisReport")) & "'" 
	  else
	      strsql2 = " "
	  end if

	  if  request.Form("vMonth") <>"ALL"  then
	      strsql3 = " and PERIOD_MONTH=" & request.Form("vMonth") 
	  else
	      strsql3 = " "
	  end if
	  if  request.Form("vBA") <>"MILL" and request.Form("vBA") <>"ESTATE"  then
	      strsql4 = "  and BA='" & request.Form("vBA") & "'" 
	  elseif  request.Form("vBA") ="MILL" then
	      strsql4 = "  and substr(BA,1,1)='" & "P" & "'" 
	  elseif  request.Form("vBA") ="ESTATE" then
	      strsql4 = "  and substr(BA,1,1)='" & "K" & "'" 
	  else
	      strsql4 = " "
	  end if

	  strsql = strsql1 + strsql2 + strsql3 + strsql4 + strsql5
	  strHA = strHA1 + strsql3 + strsql4 
	  strMill = strMill1 + strsql3 + strsql4
  END IF

Dim rsHA
Set rsHA = Server.CreateObject("ADODB.Recordset")
rsHA.CursorLocation = adUseClient
nHa = 1
if trim(strHA) <>"" then
   'response.Write strHA
   rsHA.Open strHA, connStr, adOpenForwardOnly, adLockReadOnly
   IF not rsha.eof then
      if trim(len(rsHA("HA"))) > 0 then
         nHa = rsHA("HA")

      else
         nHa = 1
      end if
    
   else
      nHA = 1
   end if
else
   nHA = 1
End if
  
' response.Write strsql
'	  strsql =  " SELECT Division,PlantingDate,ACTIVITY_HEADING,ACCOUNTGROUP,sum(TM_AMT_ACTUAL),sum(TM_FFB_ACTUAL), " & _
'				" sum(TM_AMT_BUDGET),sum(TM_FFB_BUDGET),sum(TY_AMT_ACTUAL),sum(TY_FFB_ACTUAL), " & _
'				" sum(TY_AMT_BUDGET),sum(TY_FFB_BUDGET) from AA_ACCOUNTREPORT_TM_TBM  " & _
'				" where PERIOD_YEAR=2006 and PERIOD_MONTH=1" & _
'				" and BA='KBT' and STATUS='TM' "& _
'				" group by Division,PlantingDate,ACTIVITY_HEADING,ACCOUNTGROUP " & _
'				" order by ACTIVITY_HEADING DESC "
   
rs.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly

If rs.EOF Then%>
  <tr><td colspan="3">Data is empty</tr></td>
<%Else
  rs.movefirst
  
  howmanyfields=rs.fields.count-1
  
  JumRec=1
  intRecordsPerPage = 10000
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
  %>
    <tr>
      <td width="150" colspan="4"><img src="Images/Logoawal.gif" width="150" height="130"></td>
      <td width="580" colspan="8" align="center"><font style="font-size:12px"><B>
	  MULTI INDO OIL<BR>
	  ANALISA BIAYA TIDAK LANGSUNG<br>
	  <%if Request.Form("vBA")="MILL" then%>
           ALL MILL<br>
	  <%elseif Request.Form("vBA")="ESTATE" then%>
           ALL ESTATE<br>
	  <%elseif Request.Form("vBA")<>"MILL" and Request.Form("vBA")<>"ESTATE" then%>
           BA : <%=Request.Form("vBA")%><br>
      <%End if%>	  
      Per Year :<%=Request.Form("vYear")%> Month : <%=textMonth(trim(Request.Form("vMOnth")))%><br>
	  <%=Request.Form("vJenisReport")%></font></div></B></td>
	  <td><a href="P_AccountReport.Asp?strsql=<%=strsql%>" target="_blank"><img src="Icon/Print.gif" border="0">
	  </a></td>
  </tr></TABLE>


  <table  border="0" bgcolor="#F7F7F7" cellpadding="0" cellspacing="0">
  <!--#include file="HEADERBTL.TXT"-->

  <%
  While not Rs.EOF and intRecordsPerPage > 0
     if Mid(request.Form("vBA"),1,1)<>"P"  and  trim(request.Form("vBA")) <> "MILL" then
        xRp_HA_TM_ACTUAL= cdbl(rs("TM_AMT_ACTUAL"))/cdbl(nHA)
		xRp_HA_TM_BUDGET= cdbl(rs("TM_AMT_BUDGET"))/cdbl(nHA)
		if xRp_HA_TM_BUDGET <> 0 then
		   nPersen_TM = (xRp_HA_TM_BUDGET-xRp_HA_TM_ACTUAL)/xRp_HA_TM_BUDGET*100
		else
		   nPersen_TM = 0   
		end if
        xRp_HA_TY_ACTUAL= cdbl(rs("TY_AMT_ACTUAL"))/cdbl(nHA)
		xRp_HA_TY_BUDGET= cdbl(rs("TY_AMT_BUDGET"))/cdbl(nHA)
		if xRp_HA_TM_BUDGET <> 0 then
		   nPersen_Ty = (xRp_HA_Ty_BUDGET-xRp_HA_Ty_ACTUAL)/xRp_HA_Ty_BUDGET*100
		else
		   nPersen_Ty = 0   
		end if
     elseif Mid(request.Form("vBA"),1,1)="P" then
        pRp_KG_TM_ACTUAL= cdbl(rs("TM_AMT_ACTUAL"))/cdbl(TM_FFB_ACTUAL)
		pRp_KG_TM_BUDGET= cdbl(rs("TM_AMT_BUDGET"))/cdbl(TM_FFB_BUDGET)
		if cdbl(pRp_KG_TM_BUDGET) <> 0 then
		   pPersen_TM = (pRp_KG_TM_BUDGET-pRp_KG_TM_ACTUAL)/pRp_KG_TM_BUDGET*100
		else
		   pPersen_TM = 0   
		end if
        pRp_KG_TY_ACTUAL= cdbl(rs("TY_AMT_ACTUAL"))/cdbl(TY_FFB_ACTUAL)
		pRp_KG_TY_BUDGET= cdbl(rs("TY_AMT_BUDGET"))/cdbl(TY_FFB_BUDGET)
		if cdbl(pRp_KG_TY_BUDGET) <> 0 then
		   pPersen_TY = (pRp_KG_TY_BUDGET-pRp_KG_TY_ACTUAL)/pRp_KG_TY_BUDGET*100
		else
		   pPersen_TY = 0   
		end if
		'response.Write "actual " & cdbl(rs("TM_AMT_ACTUAL")) & " budget : " & cdbl(rs("TM_AMT_BUDGET")) & " ffb actual " & cdbl(TM_FFB_ACTUAL)  & " ffb budget: " & cdbl(TM_FFB_BUDGET)
	 else
        xRp_HA_TM_ACTUAL= 0
		xRp_HA_TM_BUDGET= 0
		nPersen_TM = 0
        xRp_HA_TY_ACTUAL= 0
		xRp_HA_TY_BUDGET= 0
		nPersen_TY = 0
		pPersen_TM = 0
		pPersen_TY = 0
	 end if
     if Mid(request.Form("vBA"),1,1)<>"P" and  trim(request.Form("vBA")) <> "MILL"  then%>
            <TR>
			<td style="<%=cmdstyle%>;border-right: 1 solid #0000FF" valign="top" bgcolor="#C0FBBD" align="left">
			<font face="Verdana" size="1"><%=rs(0)%></td>
			<%if trim(HeadDescription)=trim(rs(1)) then%>
			     <td style="<%=cmdstyle%>;border-right: 1 solid #0000FF" valign="top" bgcolor="#C0FBBD" align="left"><font face="Verdana" size="1">&nbsp;
			     </td>
			<%else%>
			     <td style="<%=cmdstyle%>;border-right: 1 solid #0000FF"valign="top" bgcolor="#C0FBBD" align="left"><font face="Verdana" size="1">
			     <%=rs(1)%></td>
			<%End if%>
			<%if request.Form("vJenisReport") ="ALLGOL" then%>
				<td style="<%=cmdstyle%>"valign="top" bgcolor="#C0FBBD" align="left">
				<font face="Verdana" size="1"><%=rs(2)%></td>
			<%End if%>
			<td style="<%=cmdstyle%>"valign="top" bgcolor="#C0FBBD" align="right">
			<font face="Verdana" size="1"><%=FormatNumber(cdbl(rs("TM_AMT_ACTUAL")),2)%></td>
			<td style="<%=cmdstyle%>"valign="top" bgcolor="#C0FBBD" align="right">
			<font face="Verdana" size="1"><%=FormatNumber(cdbl(rs("TM_AMT_ACTUAL"))/cdbl(nHa),2)%></td>
			<td style="<%=cmdstyle%>"valign="top" bgcolor="#C0FBBD" align="right"><font face="Verdana" size="1">
			<%=FormatNumber(cdbl(rs("TM_AMT_BUDGET")),2)%></td>
			<td style="<%=cmdstyle%>"valign="top" bgcolor="#C0FBBD" align="right"><font face="Verdana" size="1">
			<%=FormatNumber(cdbl(rs("TM_AMT_BUDGET"))/cdbl(nHa),2)%></font></td>
			<%if cdbl(nPersen_TM) <0 then%>
			     <td style="<%=cmdstyle%>"valign="top" bgcolor="#C0FBBD" align="right"><font face="Verdana" size="1" color="#FF0000">
			     <%=FormatNumber(nPersen_TM,2)%></font></td>
			<%else%>
			     <td style="<%=cmdstyle%>"valign="top" bgcolor="#C0FBBD" align="right"><font face="Verdana" size="1">
			     <%=FormatNumber(nPersen_TM,2)%></font></td>
			<%End if%>
			<td style="<%=cmdstyle%>"valign="top" bgcolor="#C0FBBD" align="right"><font face="Verdana" size="1">
			<%=FormatNumber(cdbl(rs("TY_AMT_ACTUAL")),2)%></td>
			<td style="<%=cmdstyle%>"valign="top" bgcolor="#C0FBBD" align="right"><font face="Verdana" size="1">
			<%=FormatNumber(cdbl(rs("TY_AMT_ACTUAL"))/cdbl(nHa),2)%></font></td>
			<td style="<%=cmdstyle%>"valign="top" bgcolor="#C0FBBD" align="right"><font face="Verdana" size="1">
			<%=FormatNumber(cdbl(rs("TY_AMT_BUDGET")),2)%></td>
			<td style="<%=cmdstyle%>"valign="top" bgcolor="#C0FBBD" align="right"><font face="Verdana" size="1">
			<%=FormatNumber(cdbl(rs("TY_AMT_BUDGET"))/cdbl(nHa),2)%></font></td>
			<%if cdbl(nPersen_TY) < 0 then%>
			     <td style="<%=cmdstyle%>"valign="top" bgcolor="#C0FBBD" align="right"><font face="Verdana" size="1" color="#FF0000">
			     <%=FormatNumber(nPersen_TY,2)%></font></td>
			<%else%>
			     <td style="<%=cmdstyle%>"valign="top" bgcolor="#C0FBBD" align="right"><font face="Verdana" size="1">
			     <%=FormatNumber(nPersen_TY,2)%></font></td>
			<%end if%>
			<td style="<%=cmdstyle%>"valign="top" bgcolor="#C0FBBD" align="right"><font face="Verdana" size="1">
			<%=FormatNumber(cdbl(rs("AMT_ANNUAL")),2)%></td>
			<td style="<%=cmdstyle%>"valign="top" bgcolor="#C0FBBD" align="right"><font face="Verdana" size="1">
			<%=FormatNumber(cdbl(rs("AMT_ANNUAL"))/cdbl(nHa),2)%></font></td>
     <%else  '------------------------Jika Untuk Pabrik%>
			<TR>
			<td style="<%=cmdstyle%>"valign="top" bgcolor="#C0FBBD" align="left"><font face="Verdana" size="1">
			<%=rs(0)%></td>
			<%if trim(HeadDescription)=trim(rs(1)) then%>
			     <td style="<%=cmdstyle%>"valign="top" bgcolor="#C0FBBD" align="left"><font face="Verdana" size="1">&nbsp;
			     </td>
			<%else%>
			     <td style="<%=cmdstyle%>"valign="top" bgcolor="#C0FBBD" align="left"><font face="Verdana" size="1">
			     <%=rs(1)%></td>
			<%End if%>
			<%if request.Form("vJenisReport") ="ALLGOL" THEN%>
				<td style="<%=cmdstyle%>"valign="top" bgcolor="#C0FBBD" align="left">
				<font face="Verdana" size="1"><%=rs(2)%></td>
			<%End if%>
			<td style="<%=cmdstyle%>"valign="top" bgcolor="#C0FBBD" align="right"><font face="Verdana" size="1">
			<%=FormatNumber(cdbl(rs("TM_AMT_ACTUAL")),2)%></td>
			<td style="<%=cmdstyle%>"valign="top" bgcolor="#C0FBBD" align="right"><font face="Verdana" size="1">
			<%=FormatNumber(cdbl(rs("TM_AMT_ACTUAL"))/cdbl(TM_PALM_ACTUAL),2)%></td>
			<td style="<%=cmdstyle%>"valign="top" bgcolor="#C0FBBD" align="right"><font face="Verdana" size="1">
			<%=FormatNumber(cdbl(rs("TM_AMT_ACTUAL"))/cdbl(TM_FFB_ACTUAL),2)%></td>
			<td style="<%=cmdstyle%>"valign="top" bgcolor="#C0FBBD" align="right"><font face="Verdana" size="1">
			<%=FormatNumber(cdbl(rs("TM_AMT_BUDGET")),2)%></td>
			<td style="<%=cmdstyle%>"valign="top" bgcolor="#C0FBBD" align="right"><font face="Verdana" size="1">
			<%=FormatNumber(cdbl(rs("TM_AMT_BUDGET"))/cdbl(TM_PALM_BUDGET),2)%></font></td>
			<td style="<%=cmdstyle%>"valign="top" bgcolor="#C0FBBD" align="right"><font face="Verdana" size="1">
			<%=FormatNumber(cdbl(rs("TM_AMT_BUDGET"))/cdbl(TM_FFB_BUDGET),2)%></font></td>
			<%if cdbl(pPersen_TM) < 0 then%>
				<td style="<%=cmdstyle%>"valign="top" bgcolor="#C0FBBD" align="right"><font face="Verdana" size="1" color="#FF0000">
				<%=FormatNumber(pPersen_TM,2)%></font></td>
			<%else%>
				 <td style="<%=cmdstyle%>"valign="top" bgcolor="#C0FBBD" align="right"><font face="Verdana" size="1">
				 <%=FormatNumber(nPersen_TM,2)%></font></td>
			<%End if%>
			<td style="<%=cmdstyle%>"valign="top" bgcolor="#C0FBBD" align="right"><font face="Verdana" size="1">
			<%=FormatNumber(cdbl(rs("TY_AMT_ACTUAL")),2)%></td>
			<td style="<%=cmdstyle%>"valign="top" bgcolor="#C0FBBD" align="right"><font face="Verdana" size="1">
			<%=FormatNumber(cdbl(rs("TY_AMT_ACTUAL"))/cdbl(TY_PALM_ACTUAL),2)%></font></td>
			<td style="<%=cmdstyle%>"valign="top" bgcolor="#C0FBBD" align="right"><font face="Verdana" size="1">
			<%=FormatNumber(cdbl(rs("TY_AMT_ACTUAL"))/cdbl(TY_FFB_ACTUAL),2)%></font></td>
			<td style="<%=cmdstyle%>"valign="top" bgcolor="#C0FBBD" align="right"><font face="Verdana" size="1">
			<%=FormatNumber(cdbl(rs("TY_AMT_BUDGET")),2)%></td>
			<td style="<%=cmdstyle%>"valign="top" bgcolor="#C0FBBD" align="right"><font face="Verdana" size="1">
			<%=FormatNumber(cdbl(rs("TY_AMT_BUDGET"))/cdbl(TY_PALM_BUDGET),2)%></font></td>
			<td style="<%=cmdstyle%>"valign="top" bgcolor="#C0FBBD" align="right"><font face="Verdana" size="1">
			<%=FormatNumber(cdbl(rs("TY_AMT_BUDGET"))/cdbl(TY_FFB_BUDGET),2)%></font></td>
			<%if cdbl(pPersen_Ty) < 0 then%>
				<td style="<%=cmdstyle%>"valign="top" bgcolor="#C0FBBD" align="right"><font face="Verdana" size="1" color="#FF0000">
				<%=FormatNumber(pPersen_TY,2)%></font></td>
			<%else%>
				<td style="<%=cmdstyle%>"valign="top" bgcolor="#C0FBBD" align="right"><font face="Verdana" size="1">
				<%=FormatNumber(nPersen_TY,2)%></font></td>
			<%end if%>
			<td style="<%=cmdstyle%>"valign="top" bgcolor="#C0FBBD" align="right"><font face="Verdana" size="1">
			<%=FormatNumber(cdbl(rs("AMT_ANNUAL")),2)%></td>
			<td style="<%=cmdstyle%>"valign="top" bgcolor="#C0FBBD" align="right"><font face="Verdana" size="1">
			<%=FormatNumber(cdbl(rs("AMT_ANNUAL"))/cdbl(PALM_ANNUAL),2)%></font></td>
			<td style="<%=cmdstyle%>"valign="top" bgcolor="#C0FBBD" align="right"><font face="Verdana" size="1">
			<%=FormatNumber(cdbl(rs("AMT_ANNUAL"))/cdbl(FFB_ANNUAL),2)%></font></td>
     <%End If 'akhir Jika Pabrik%>
	 </TR>
	 <% 
	 HeadAfdeling = rs.fields(0)
	 HeadDescription   = rs.fields(1)
	 Headactivity     = rs.fields(2)
	 HeadJob          = rs.fields(3)

     nTotal_TM_Actual = nTotal_TM_Actual + cdbl(rs("TM_AMT_ACTUAL"))
     nTotal_TY_Actual = nTotal_TY_Actual + cdbl(rs("Ty_AMT_ACTUAL"))
     nTotal_TM_BUDGET = cdbl(nTotal_TM_BUDGET) + cdbl(rs("TM_AMT_BUDGET"))
	 'response.Write "Budghet : " & cdbl(rs("TM_AMT_BUDGET")) & "Total : " & nTotal_TM_BUDGET
     nTotal_TY_BUDGET = nTotal_TY_BUDGET + cdbl(rs("Ty_AMT_BUDGET"))
	 nTotal_AMT_Annual= nTotal_AMT_Annual + cdbl(rs("AMT_ANNUAL"))

     rs.movenext
	 
     JumRec=JumRec+1
	 intRecordsPerPage = intRecordsPerPage - 1
  Wend
End If
rs.close
set rs = nothing
IF len(trim(nHA)) > 0 and len(trim(request.Form("vJenisReport"))) > 0 THEN
   nRpPerHA_TM = nTotal_TM_Actual/CDBL(nHA)
   nRpPerHA_TM_Budget = nTotal_TM_Budget/CDBL(nHA)
   if nRpPerHA_TM_Budget > 0 then
      nRpPersen_TM = (nRpPerHA_TM_Budget - nRpPerHA_TM) / nRpPerHA_TM_Budget * 100
   else
      nRpPersen_TM = 0
   end if
   nRpPerHA_TY_Budget = nTotal_TY_budget/CDBL(nHA)
   nRpPerHA_TY = nTotal_TY_Actual/CDBL(nHA)
   if nRpPerHA_TY_Budget > 0 then
      nRpPersen_TY = (nRpPerHA_TY_Budget - nRpPerHA_TY) / nRpPerHA_TY_Budget * 100
   else
      nRpPersen_TY = 0
   end if
   if CDBL(nHA) > 0 then
      nRpPer_Annual = nTotal_AMT_Annual/CDBL(nHA)
   else
      nRpPer_Annual = 0	  
   end if
else
   nRpPerHA_TM = 0
   nRpPerHA_TY = 0
   nRpPerHA_TM_Budget = 0
   nRpPerHA_TY_Budget = 0
   nRpPer_Annual=0
   nRpPersen_TM=0
   nRpPersen_TY=0
END IF

End if

if Mid(request.Form("vBA"),1,1)="P" then
    If cdbl(TM_PALM_ACTUAL) > 0 then
	   x_PALM_TM_ACTUAL= cdbl(nTotal_TM_Actual)/cdbl(TM_PALM_ACTUAL)
	else
	   x_PALM_TM_ACTUAL= 0
	end if
	if cdbl(TM_FFB_ACTUAL) > 0 then
	   x_FFB_TM_ACTUAL= cdbl(nTotal_TM_Actual)/cdbl(TM_FFB_ACTUAL)
	else
	   x_FFB_TM_ACTUAL= 0
	end if
	if cdbl(TM_PALM_BUDGET) > 0 then
	   x_PALM_TM_BUDGET= cdbl(nTotal_TM_BUDGET)/cdbl(TM_PALM_BUDGET)
	else
	   x_PALM_TM_BUDGET= 0
	end if
	if cdbl(TM_FFB_BUDGET) > 0 then
	   x_FFB_TM_BUDGET= cdbl(nTotal_TM_BUDGET)/cdbl(TM_FFB_BUDGET)
	else
	   x_FFB_TM_BUDGET= 0
	end if

    if cdbl(Ty_PALM_ACTUAL) > 0 then
	   x_PALM_Ty_ACTUAL= cdbl(nTotal_Ty_Actual)/cdbl(Ty_PALM_ACTUAL)
	else
	   x_PALM_Ty_ACTUAL= 0
	end if
	if cdbl(Ty_FFB_ACTUAL) > 0 Then
	   x_FFB_Ty_ACTUAL= cdbl(nTotal_Ty_Actual)/cdbl(Ty_FFB_ACTUAL)
	else
	   x_FFB_Ty_ACTUAL= 0
	end if
	if cdbl(Ty_PALM_BUDGET) > 0 then
	   x_PALM_Ty_BUDGET= cdbl(nTotal_Ty_BUDGET)/cdbl(Ty_PALM_BUDGET)
	else
	   x_PALM_Ty_BUDGET= 0
	end if
	if cdbl(Ty_FFB_BUDGET) > 0 then
	   x_FFB_Ty_BUDGET= cdbl(nTotal_Ty_BUDGET)/cdbl(Ty_FFB_BUDGET)
	else
	   x_FFB_Ty_BUDGET= 0
	end if

    if cdbl(Tm_FFB_ACTUAL) > 0 then
	   xRp_KG_Tm_ACTUAL= cdbl(nTotal_TM_ACTUAL)/cdbl(Tm_FFB_ACTUAL)
	else
	   xRp_KG_Tm_ACTUAL= 0
	end if
	if cdbl(Tm_FFB_BUDGET) > 0 then
	   xRp_KG_Tm_BUDGET= cdbl(nTotal_TM_BUDGET)/cdbl(Tm_FFB_BUDGET)
	else
	   xRp_KG_Tm_BUDGET= 0
	end if
	if cdbl(xRp_KG_TM_BUDGET) <> 0 then
	   nPersen_TM = (xRp_KG_TM_BUDGET-xRp_KG_TM_ACTUAL)/xRp_KG_TM_BUDGET*100
	else
	   nPersen_TM = 0   
	end if
	if cdbl(TY_FFB_ACTUAL) > 0 then
	   xRp_KG_TY_ACTUAL= cdbl(nTotal_TY_ACTUAL)/cdbl(TY_FFB_ACTUAL)
	else
	   xRp_KG_TY_ACTUAL= 0
	end if
	if cdbl(TY_FFB_BUDGET) > 0 then
	   xRp_KG_TY_BUDGET= cdbl(nTotal_TY_BUDGET)/cdbl(TY_FFB_BUDGET)
	else
	   xRp_KG_TY_BUDGET= 0
	end if
	if cdbl(xRp_KG_TY_BUDGET) <> 0 then
	   nPersen_TY = (xRp_KG_TY_BUDGET-xRp_KG_TY_ACTUAL)/xRp_KG_TY_BUDGET*100
	else
	   nPersen_TY = 0   
	end if
    'response.Write "Tm Actual: "  & nTotal_TM_ACTUAL & " TM Budget " & nTotal_TM_BUDGET & " FFB Actual : " & Tm_FFB_ACTUAL & " FFB Budget " & Tm_FFB_BUDGET
	if cdbl(PALM_ANNUAL) > 0 then
	   x_PALM_ANNUAL= cdbl(nTotal_AMT_Annual)/cdbl(PALM_ANNUAL)
	else
	   x_PALM_ANNUAL= 0
	end if
	if cdbl(FFB_ANNUAL) > 0 then
	   x_FFB_ANNUAL= cdbl(nTotal_AMT_Annual)/cdbl(FFB_ANNUAL)
	else
	   x_FFB_ANNUAL= 0
	end if
    %>
	<tr bgcolor="#0099FF">
	<%if request.Form("vJenisReport") ="ALLGOL"  then%>
	   <td style="<%=cmdstyle%>"colspan="3" align="center"><b><font color="#FFFFFF">T O T A L</b></td>
	<%else%>   
	   <td style="<%=cmdstyle%>"colspan="2" align="center"><b><font color="#FFFFFF">T O T A L</b></td>
	<%End if%>	
	<td style="<%=cmdstyle%>"colspan="1"><b><font color="#FFFFFF"><%=FormatNumber(nTotal_TM_Actual,2)%></b></td>
	<td style="<%=cmdstyle%>"colspan="1" align="Right"><font color="#FFFFFF"><b><%=FormatNumber(x_PALM_TM_ACTUAL,2)%></td>
	<td style="<%=cmdstyle%>"colspan="1" align="Right"><font color="#FFFFFF"><b><%=FormatNumber(x_FFB_TM_ACTUAL,2)%></td>
	<td style="<%=cmdstyle%>"colspan="1"><font color="#FFFFFF"><b><%=FormatNumber(nTotal_TM_BUDGET,2)%></b></td>
	<td style="<%=cmdstyle%>"colspan="1" align="Right"><font color="#FFFFFF"><b><%=FormatNumber(x_PALM_TM_BUDGET,2)%></td>
	<td style="<%=cmdstyle%>"colspan="1" align="Right"><font color="#FFFFFF"><b><%=FormatNumber(x_FFB_TM_BUDGET,2)%></td>
	<td style="<%=cmdstyle%>"colspan="1" align="Right"><font color="#FFFFFF"><b><%=FormatNumber(nPersen_TM,2)%></b></td>
	<td style="<%=cmdstyle%>"colspan="1"><font color="#FFFFFF"><b><%=FormatNumber(nTotal_TY_Actual,2)%></b></td>
	<td style="<%=cmdstyle%>"colspan="1" align="Right"><font color="#FFFFFF"><b><%=FormatNumber(x_PALM_TY_ACTUAL,2)%></td>
	<td style="<%=cmdstyle%>"colspan="1" align="Right"><font color="#FFFFFF"><b><%=FormatNumber(x_FFB_TY_ACTUAL,2)%></td>
	<td style="<%=cmdstyle%>"colspan="1"><font color="#FFFFFF"><b><%=FormatNumber(nTotal_TY_BUDGET,2)%></b></td>
	<td style="<%=cmdstyle%>"colspan="1" align="Right"><font color="#FFFFFF"><b><%=FormatNumber(x_PALM_TY_BUDGET,2)%></td>
	<td style="<%=cmdstyle%>"colspan="1" align="Right"><font color="#FFFFFF"><b><%=FormatNumber(x_FFB_TY_BUDGET,2)%></td>
	<td style="<%=cmdstyle%>"colspan="1" align="Right"><font color="#FFFFFF"><b><%=FormatNumber(nPersen_TY,2)%></b></td>
	<td style="<%=cmdstyle%>"colspan="1"><font color="#FFFFFF"><b><%=FormatNumber(nTotal_AMT_Annual,2)%></b></td>
	<td style="<%=cmdstyle%>"colspan="1" align="Right"><font color="#FFFFFF"><b><%=FormatNumber(x_PALM_ANNUAL,2)%></b></td>
	<td style="<%=cmdstyle%>"colspan="1" align="Right"><font color="#FFFFFF"><b><%=FormatNumber(x_FFB_ANNUAL,2)%></b></td>
	</tr>
<%else%>
	<tr bgcolor="#0099FF">
	<%if request.Form("vJenisReport") ="ALLGOL"  then%>
	   <td style="<%=cmdstyle%>"colspan="3" align="center"><b><font color="#FFFFFF">T O T A L</b></td>
	<%else%>   
	   <td style="<%=cmdstyle%>"colspan="2" align="center"><b><font color="#FFFFFF">T O T A L</b></td>
	<%End if%>	
	<td style="<%=cmdstyle%>"colspan="1"><font color="#FFFFFF"><b><%=FormatNumber(nTotal_TM_Actual,2)%></b></td>
	<td style="<%=cmdstyle%>"colspan="1" align="Right"><font color="#FFFFFF"><b><%=FormatNumber(nRpPerHA_TM,2)%></b></td>
	<td style="<%=cmdstyle%>"colspan="1"><font color="#FFFFFF"><b><%=FormatNumber(nTotal_TM_BUDGET,2)%></b></td>
	<td style="<%=cmdstyle%>"colspan="1" align="Right"><font color="#FFFFFF"><b><%=FormatNumber(nRpPerHA_TM_Budget,2)%></b></td>
	<td style="<%=cmdstyle%>"colspan="1" align="Right"><font color="#FFFFFF"><b><%=FormatNumber(nRpPersen_TM,2)%></b></td>
	<td style="<%=cmdstyle%>"colspan="1"><font color="#FFFFFF"><b><%=FormatNumber(nTotal_TY_Actual,2)%></b></td>
	<td style="<%=cmdstyle%>"colspan="1" align="Right"><font color="#FFFFFF"><b><%=FormatNumber(nRpPerHA_TY,2)%></b></td>
	<td style="<%=cmdstyle%>"colspan="1"><font color="#FFFFFF"><b><%=FormatNumber(nTotal_TY_BUDGET,2)%></b></td>
	<td style="<%=cmdstyle%>"colspan="1" align="Right"><font color="#FFFFFF"><b><%=FormatNumber(nRpPerHA_TY_Budget,2)%></b></td>
	<td style="<%=cmdstyle%>"colspan="1" align="Right"><font color="#FFFFFF"><b><%=FormatNumber(nRpPersen_TY,2)%></b></td>
	<td style="<%=cmdstyle%>"colspan="1"><font color="#FFFFFF"><b><%=FormatNumber(nTotal_AMT_Annual,2)%></b></td>
	<td style="<%=cmdstyle%>"colspan="1" align="Right"><font color="#FFFFFF"><b><%=FormatNumber(nRpPer_Annual,2)%></b></td>
	</tr>
<tr>
<td style="<%=cmdstyle%>"colspan="14">
   <%if intNumberOfPages<>1 then
      For i = 1 to intNumberOfPages
           Response.Write "<font size=""2"" face=""Arial Narrow"" color=""#FFFFFF""><a href=""AccountReport.asp?&iPageNo=" &  i & "&vJenisReport=" & request.Form("vJenisReport")& "&fcValue_Organization=" & request("fcValue_Organization") & """>" & i & "  "
      Next
   End if%>
   </a><br>
   <%if Request("iPageNo") = "" then%>
        <%nPage = 1%>
   <%Else%>
        <%nPage = Request("iPageNo")%>
   <%End if%>
   <font size="1" color="#000000" face="Arial, Helvetica, sans-serif">Page : <%=nPage%>
   <font size="1" color="#000000" face="Arial, Helvetica, sans-serif">Total Records : <%=intRecords%></td>
</tr>
</Table></center>
   <font size="-4">
   <%'response.Write(strsql)%></font><%
%>
</font>
<%End if%>
<%End if%>
<!--#include file="../Bottom_Report.asp"-->