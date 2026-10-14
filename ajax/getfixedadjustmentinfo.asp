<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
fixedassetcode = Request("fixedassetcode")
fcba = Request("fcba")
creationdate = Request("creationdate")
jobcode = Request("jobcode")
customercode = Request("customercode")
txtYear = Request("txtYear")
txtMonth = Request("txtMonth")
txttdate = Request("txttdate")

If "" & txtYear = "" Then txtYear = Year(Session("v_Costbook"))
If "" & txtMonth = "" Then txtMonth = Month(Session("v_Costbook"))
If "" & txttdate = "" Then txttdate = textDate(date)

txttdate = txtMonth & "" & txtYear

result = ""


q =	  	" SELECT fa.fcba," & _
		"       asg.controljob," & _
		"       fa.fccode," & _
		"       fa.fcname," & _
		"       fa.assetcode," & _
		"       fa.assetname," & _
		"       fa.assetgroupcode," & _
		"       fa.assetgroupname" & _
		"  FROM fixedasset fa" & _
		"       LEFT JOIN asset ass ON (ass.fccode = fa.assetcode)" & _
		"       LEFT JOIN assetgroup asg ON (asg.fccode = fa.assetgroupcode)" & _
		" WHERE fa.fcba = '"& session("fcba") &"' AND fa.fccode = '"& fixedassetcode &"'"
 

' response.Write q
set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr

if Not rsrecurtrans.eof then
	fcba = rsrecurtrans("fcba")
	jobcode = rsrecurtrans("controljob")
	fixedassetcode = rsrecurtrans("fccode")
	fixedassetname = rsrecurtrans("fcname")
	assetcode = rsrecurtrans("assetcode")
	assetname = rsrecurtrans("assetname")
	assetgroupcode = rsrecurtrans("assetgroupcode")
	assetgroupname = rsrecurtrans("assetgroupname")
end if
rsrecurtrans.Close
set rsrecurtrans = Nothing

	strAvailable =	"  SELECT fcba," & _
					"         locationcode," & _
					"         locationtype," & _
					"         jobcode," & _
					"         SUM (AMOUNT) amount" & _
					"    FROM (  SELECT transdate," & _
					"                   locationcode," & _
					"                   locationtype," & _
					"                   jobcode," & _
					"                   SUM (AMOUNT) amount," & _
					"                   fcba" & _
					"              FROM (  SELECT pvd.locationtype," & _
					"                             pvd.locationcode," & _
					"                             pvd.jobcode," & _
					"                             SUM (pvd.amount) amount," & _
					"                             pv.datecreated AS transdate," & _
					"                             pvd.fcba" & _
					"                        FROM paymentvoucherdetail pvd" & _
					"                             INNER JOIN paymentvoucher pv" & _
					"                                ON (    pv.vouchercode = pvd.vouchercode" & _
					"                                    AND pv.fcba = pvd.fcba)" & _
					"                       WHERE     pvd.fcba = 'MKE'" & _
					"                             AND pvd.locationcode = '"& fixedassetcode &"'" & _
					"                             AND pvd.locationtype = 'FA'" & _
					"                    GROUP BY pvd.locationtype," & _
					"                             pvd.locationcode," & _
					"                             pvd.jobcode," & _
					"                             pv.datecreated," & _
					"                             pvd.fcba" & _
					"                    UNION ALL" & _
					"                      SELECT rnd.locationtype," & _
					"                             rnd.locationcode," & _
					"                             rnd.jobcode," & _
					"                             SUM (rnd.amountreceived) amount," & _
					"                             rn.rndate AS transdate," & _
					"                             rnd.fcba" & _
					"                        FROM receivenotedetail rnd" & _
					"                             INNER JOIN receivenote rn" & _
					"                                ON (    rn.receivenotecode = rnd.receivenotecode" & _
					"                                    AND rn.fcba = rnd.fcba)" & _
					"                       WHERE     rnd.fcba = 'MKE'" & _
					"                             AND rnd.locationcode = '"& fixedassetcode &"'" & _
					"                             AND rnd.locationtype = 'FA'" & _
					"                    GROUP BY rnd.locationtype," & _
					"                             rnd.locationcode," & _
					"                             rnd.jobcode," & _
					"                             rn.rndate," & _
					"                             rnd.fcba" & _
					"                    UNION ALL" & _
					"                      SELECT locationtype," & _
					"                             locationcode," & _
					"                             jobcode," & _
					"                             SUM (amount) amount," & _
					"                             wipdate AS transdate," & _
					"                             fcba" & _
					"                        FROM (  SELECT wipr.fcba," & _
					"                                       wipp.wipdate," & _
					"                                       wipr.agreementfinalno," & _
					"                                       wipr.agreementcode," & _
					"                                       wipr.agreementlineno," & _
					"                                       wipp.jobcode," & _
					"                                       wipp.locationtype locationtype," & _
					"                                       wipp.locationcode locationcode," & _
					"                                       NVL (SUM (wipr.wipamt), 0) amount" & _
					"                                  FROM workinprogressplan wipp" & _
					"                                       INNER JOIN workinprogress wipr" & _
					"                                          ON (    wipr.agreementcode =" & _
					"                                                     wipp.agreementcode" & _
					"                                              AND wipr.fcba = wipp.fcba" & _
					"                                              AND wipr.agreementlineno =" & _
					"                                                     wipp.agreementlineno" & _
					"                                              AND wipr.agreementfinalno =" & _
					"                                                     wipp.agreementfinalno)" & _
					"                                 WHERE     wipp.fcba = 'MKE'" & _
					"                                       AND wipp.locationcode = '"& fixedassetcode &"'" & _
					"                                       AND wipp.locationtype = 'FA'" & _
					"                              GROUP BY wipr.fcba," & _
					"                                       wipr.agreementfinalno," & _
					"                                       wipr.agreementcode," & _
					"                                       wipr.agreementlineno," & _
					"                                       wipp.jobcode," & _
					"                                       wipp.locationtype," & _
					"                                       wipp.locationcode," & _
					"                                       wipp.wipdate)" & _
					"                    GROUP BY locationtype," & _
					"                             locationcode," & _
					"                             jobcode," & _
					"                             wipdate," & _
					"                             fcba)" & _
					"          GROUP BY transdate," & _
					"                   locationcode," & _
					"                   locationtype," & _
					"                   jobcode," & _
					"                   fcba)" & _
					"   WHERE     locationcode = '"& fixedassetcode &"'" & _
					"         AND locationtype = 'FA'" & _
					"         AND jobcode = '"& jobcode &"'" & _
					"         AND fcba = '"& session("fcba") &"'" & _
					"         AND TO_NUMBER (TO_CHAR (transdate, 'yyyy')) = '"& txtYear &"'" & _
					"         AND TO_NUMBER (TO_CHAR (transdate, 'mm')) = '"& txtMonth &"'" & _
					" GROUP BY locationcode," & _
					"          locationtype," & _
					"          jobcode," & _
					"          fcba" 


	' response.write stravailable
	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3, 3
	
	If Not rsAvailable.EOF Then
		amount =  Cdbl(to_nol(rsAvailable("amount")))
	Else
		amount = 0
	End if
	rsAvailable.Close
	set rsAvailable = Nothing

	' if cdbl(to_nol(initialAvailable)) = 0 then
	' 	unitprice = salesprices 
	' Else
	' 	unitprice = initialAvailable
	' End If
		response.write "{"
		' response.write """quantity"" : """ & quantity & ""","
		' response.write """unitprice"" : """ & unitprice & ""","
		' response.write """productcode"" : """ & productcode & ""","
		response.write """amount"" : """ & amount & """"
		response.write "}"		

%>