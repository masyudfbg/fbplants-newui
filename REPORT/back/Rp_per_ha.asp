<%
	  if request("vJenisReport")="Group" or request("vJenisReport")="Source" then
		  if cdbl(mid(rs("ACCOUNTGROUPCODE"),1,4))>=6502 and request.Form("vStatus") ="TM" then
		     if len(trim(nTM_HA_KG_ACTUAL)) > 0 and len(trim(nTM_HA_KG_ACTUAL)) then
			    nTM_RP_HA_ACTUAL = CDBL(rs("TM_AMT_ACTUAL"))/CDBL(nTM_HA_KG_ACTUAL)
		     else
			    nTM_RP_HA_ACTUAL = 0		
			 end if
			 if len(trim(nTY_HA_KG_ACTUAL)) > 0 and cdbl(nTY_HA_KG_ACTUAL) > 0  then
			    nTY_RP_HA_ACTUAL = CDBL(rs("TY_AMT_ACTUAL"))/CDBL(nTY_HA_KG_ACTUAL)
			 else
			    nTY_RP_HA_ACTUAL = 0
			 end if
			 if len(trim(nTM_HA_KG_BUDGET)) > 0  and cdbl(nTM_HA_KG_BUDGET) > 0 then
			    nTM_RP_HA_BUDGET = CDBL(rs("TM_AMT_BUDGET"))/CDBL(nTM_HA_KG_BUDGET)
			 else	
			    nTM_RP_HA_BUDGET = 0
			 end if
			 if len(trim(nTY_HA_KG_BUDGET)) > 0 and cdbl(nTY_HA_KG_BUDGET) > 0 then
			    nTY_RP_HA_BUDGET = CDBL(rs("TY_AMT_BUDGET"))/CDBL(nTY_HA_KG_BUDGET)
			 else
			    nTY_RP_HA_BUDGET = 0
			 end if
			 if len(trim(nHA_KG_ANNUAL)) > 0 and cdbl(nHA_KG_ANNUAL) > 0 then
			    nRP_HA_ANNUAL    = CDBL(rs("AMT_ANNUAL"))/CDBL(nHA_KG_ANNUAL)
			 else
			    nRP_HA_ANNUAL    = 0
			 end if
			 if nTM_RP_HA_BUDGET<>0 then
			    nPersenTM        = (nTM_RP_HA_BUDGET - nTM_RP_HA_ACTUAL ) / nTM_RP_HA_BUDGET * 100
			 end if
			 if nTY_RP_HA_BUDGET<>0 then
			    nPersenTY        = (nTY_RP_HA_BUDGET - nTY_RP_HA_ACTUAL ) / nTY_RP_HA_BUDGET * 100
			 end if
		  ELSE
		     if CDBL(nHA) > 0 then
			    nTM_RP_HA_ACTUAL = CDBL(rs("TM_AMT_ACTUAL"))/CDBL(nHA)
			    nTY_RP_HA_ACTUAL = CDBL(rs("TY_AMT_ACTUAL"))/CDBL(nHA)
			    nTM_RP_HA_BUDGET = CDBL(rs("TM_AMT_BUDGET"))/CDBL(nHA)
			    nTY_RP_HA_BUDGET = CDBL(rs("TY_AMT_BUDGET"))/CDBL(nHA)
			    nRP_HA_ANNUAL = CDBL(rs("AMT_ANNUAL"))/CDBL(nHA)
			 else
			    nTM_RP_HA_ACTUAL = CDBL(rs("TM_AMT_ACTUAL"))/CDBL(nHA)
			    nTY_RP_HA_ACTUAL = CDBL(rs("TY_AMT_ACTUAL"))/CDBL(nHA)
			    nTM_RP_HA_BUDGET = CDBL(rs("TM_AMT_BUDGET"))/CDBL(nHA)
			    nTY_RP_HA_BUDGET = CDBL(rs("TY_AMT_BUDGET"))/CDBL(nHA)
			    nRP_HA_ANNUAL = CDBL(rs("AMT_ANNUAL"))/CDBL(nHA)
			 end if
			 if nTM_RP_HA_BUDGET<>0 then
			    nPersenTM        = (nTM_RP_HA_BUDGET - nTM_RP_HA_ACTUAL ) / nTM_RP_HA_BUDGET * 100
			 end if
			 if nTY_RP_HA_BUDGET<>0 then
			    nPersenTY        = (nTY_RP_HA_BUDGET - nTY_RP_HA_ACTUAL ) / nTY_RP_HA_BUDGET * 100
			 end if
		  end if
	  else
 		  if cdbl(mid(rs("ACCOUNTCODE"),1,4))>=4119 then
                         IF not len(trim(nTM_HA_KG_ACTUAL)) > 0 THEN
		         if CDBL(nTM_HA_KG_ACTUAL) > 0 and cdbl(nTM_HA_KG_ACTUAL) > 0  then
			    nTM_RP_HA_ACTUAL = CDBL(rs("TM_AMT_ACTUAL"))/CDBL(nTM_HA_KG_ACTUAL)
			 else
			    nTM_RP_HA_ACTUAL = 0
			 end if
			 if CDBL(nTY_HA_KG_ACTUAL) > 0 and cdbl(nTY_HA_KG_ACTUAL) > 0 then
			    nTY_RP_HA_ACTUAL = CDBL(rs("TY_AMT_ACTUAL"))/CDBL(nTY_HA_KG_ACTUAL)
			 else
			    nTY_RP_HA_ACTUAL = 0
			 end if
			 if CDBL(nTM_HA_KG_BUDGET) > 0 and cdbl(nTM_HA_KG_BUDGET) > 0  then
			    nTM_RP_HA_BUDGET = CDBL(rs("TM_AMT_BUDGET"))/CDBL(nTM_HA_KG_BUDGET)
			 else
			    nTM_RP_HA_BUDGET = 0
			 end if
			 if CDBL(nTY_HA_KG_BUDGET) > 0 and cdbl(nTY_HA_KG_BUDGET) > 0  then
			    nTY_RP_HA_BUDGET = CDBL(rs("TY_AMT_BUDGET"))/CDBL(nTY_HA_KG_BUDGET)
			 else
			    nTY_RP_HA_BUDGET = 0
			 end if
			 if CDBL(nHA_KG_ANNUAL) > 0 and cdbl(nHA_KG_ANNUAL) > 0  then
			    nRP_HA_ANNUAL    = CDBL(rs("AMT_ANNUAL"))/CDBL(nHA_KG_ANNUAL)
			 else
			    nRP_HA_ANNUAL    = 0
			 end if
			 if nTM_RP_HA_BUDGET<>0 then
			    nPersenTM        = (nTM_RP_HA_BUDGET - nTM_RP_HA_ACTUAL ) / nTM_RP_HA_BUDGET * 100
			 end if
			 if nTY_RP_HA_BUDGET<>0 then
			    nPersenTY        = (nTY_RP_HA_BUDGET - nTY_RP_HA_ACTUAL ) / nTY_RP_HA_BUDGET * 100
			 end if
                         End if
		  ELSE
		     if CDBL(nHA) > 0 then
				 nTM_RP_HA_ACTUAL = CDBL(rs("TM_AMT_ACTUAL"))/CDBL(nHA)
				 nTY_RP_HA_ACTUAL = CDBL(rs("TY_AMT_ACTUAL"))/CDBL(nHA)
				 nTM_RP_HA_BUDGET = CDBL(rs("TM_AMT_BUDGET"))/CDBL(nHA)
				 nTY_RP_HA_BUDGET = CDBL(rs("TY_AMT_BUDGET"))/CDBL(nHA)
				 nRP_HA_ANNUAL = CDBL(rs("AMT_ANNUAL"))/CDBL(nHA)
			 else
				 nTM_RP_HA_ACTUAL = CDBL(rs("TM_AMT_ACTUAL"))/CDBL(nHA)
				 nTY_RP_HA_ACTUAL = CDBL(rs("TY_AMT_ACTUAL"))/CDBL(nHA)
				 nTM_RP_HA_BUDGET = CDBL(rs("TM_AMT_BUDGET"))/CDBL(nHA)
				 nTY_RP_HA_BUDGET = CDBL(rs("TY_AMT_BUDGET"))/CDBL(nHA)
				 nRP_HA_ANNUAL = CDBL(rs("AMT_ANNUAL"))/CDBL(nHA)
			 end if
			 if nTM_RP_HA_BUDGET<>0 then
			 nPersenTM        = (nTM_RP_HA_BUDGET - nTM_RP_HA_ACTUAL ) / nTM_RP_HA_BUDGET * 100
			 end if
			 if nTY_RP_HA_BUDGET<>0 then
			 nPersenTY        = (nTY_RP_HA_BUDGET - nTY_RP_HA_ACTUAL ) / nTY_RP_HA_BUDGET * 100
			 end if
		  end if
	  end if%>