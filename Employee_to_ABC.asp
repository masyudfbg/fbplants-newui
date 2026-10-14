<%
'----- Tranfer to Activity Costing
response.Write "Transfer to Activity Based Costing..." & "<br>"
qBasis =  " INSERT INTO COST_ACTIVITY(" & _
          " SELECT cb.transdate,cb.jobcode,'',CB.SOURCETYPE,em.presentgrade,'',cb.targettype, cb.targetcode,'',cb.referenceno, 0,0,sum(cb.amount) as Amount,cb.fcba" & _
          " FROM employee em, employeecb cb WHERE ((cb.empcode = em.fccode) AND (cb.fcba = em.fcba))" & _
		  " and to_number(to_Char(transdate,'mm'))=" & 4 &  _
		  " and to_number(to_Char(transdate,'yyyy'))=" & 2007 & _ 
          " group by " & _
          " cb.transdate,cb.jobcode,'',CB.SOURCETYPE,em.presentgrade,'',cb.targettype, cb.targetcode,'',cb.referenceno,cb.fcba)"
connstr.Execute(toABC)
%>