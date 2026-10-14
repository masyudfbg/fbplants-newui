<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%
jobcode = Request("jobcode")
empcode = Request("empcode")
fddate = Request("fddate")
fieldcode = Request("fieldcode")

isjobpanen = 0
q = " SELECT COUNT(1) isjobpanen FROM parameterdetail WHERE parcode = 'HARDCODE' AND parheadcode = 'JOBCODE_PANEN' AND fccode = '"&jobcode&"' AND fcba = '"& session("fcba") &"'" 

'response.write q
set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr

if Not rsrecurtrans.eof then
	isjobpanen = cdbl(to_nol(rsrecurtrans("isjobpanen")))
end if
rsrecurtrans.Close : set rsrecurtrans = Nothing

response.write "{"
if isjobpanen > 0 then
	response.write """isjobpanen"" : ""1"","
	arr_json_eof = true
	q = " SELECT under_ripe" & _
	    "        , over_ripe" & _
	    "        , abnormal" & _
	    "        , long_stalk" & _
	    "        , eaten_by_rat" & _
	    "        , unharvest_ffb" & _
	    "        , uncollect_lf_circle" & _
	    "        , uncollect_lf_piece" & _
	    "        , unarrange_ffb" & _
	    "        , unprune_frond" & _
	    "        , qe_1" & _
	    "        , qe_2" & _
	    "        , qe_3" & _
	    "        , qe_4" & _
	    "        , qe_5" & _
	    "        , qe_6" & _
	    "        , qe_7" & _
	    "        , qe_8" & _
	    "        , qe_9" & _
	    "        , qe_10" & _
		" FROM harvestingquality" & _
		" WHERE empcode = '"&empcode&"'" & _
		"       AND fddate = TO_DATE('"&fddate&"', 'mm/dd/yyyy')" & _
		"       AND fieldcode = '"&fieldcode&"'" & _
		"       AND fcba = '"&session("fcba")&"'" 

	'response.write q
	set rsrecurtrans = server.CreateObject("Adodb.recordset")
	rsrecurtrans.Open q, connstr

	if Not rsrecurtrans.eof then
		arr_json = rsrecurtrans.getrows()
		arr_json_eof = false
	end if
	rsrecurtrans.Close : set rsrecurtrans = Nothing

	if not arr_json_eof then
		for xyyy=0 to ubound(arr_json,2)
			response.write """under_ripe"":""" & arr_json(0,xyyy) & ""","
			response.write """over_ripe"":""" & arr_json(1,xyyy) & ""","
			response.write """abnormal"":""" & arr_json(2,xyyy) & ""","
			response.write """long_stalk"":""" & arr_json(3,xyyy) & ""","
			response.write """eaten_by_rat"":""" & arr_json(4,xyyy) & ""","
			response.write """unharvest_ffb"":""" & arr_json(5,xyyy) & ""","
			response.write """uncollect_lf_circle"":""" & arr_json(6,xyyy) & ""","
			response.write """uncollect_lf_piece"":""" & arr_json(7,xyyy) & ""","
			response.write """unarrange_ffb"":""" & arr_json(8,xyyy) & ""","
			response.write """unprune_frond"":""" & arr_json(9,xyyy) & ""","
			response.write """qe_1"":""" & arr_json(10,xyyy) & ""","
			response.write """qe_2"":""" & arr_json(11,xyyy) & ""","
			response.write """qe_3"":""" & arr_json(12,xyyy) & ""","
			response.write """qe_4"":""" & arr_json(13,xyyy) & ""","
			response.write """qe_5"":""" & arr_json(14,xyyy) & ""","
			response.write """qe_6"":""" & arr_json(15,xyyy) & ""","
			response.write """qe_7"":""" & arr_json(16,xyyy) & ""","
			response.write """qe_8"":""" & arr_json(17,xyyy) & ""","
			response.write """qe_9"":""" & arr_json(18,xyyy) & ""","
			response.write """qe_10"":""" & arr_json(19,xyyy) & """"

		next
	else
		response.write """under_ripe"" : ""0"","
		response.write """over_ripe"" : ""0"","
		response.write """abnormal"" : ""0"","
		response.write """long_stalk"" : ""0"","
		response.write """eaten_by_rat"" : ""0"","
		response.write """unharvest_ffb"" : ""0"","
		response.write """uncollect_lf_circle"" : ""0"","
		response.write """uncollect_lf_piece"" : ""0"","
		response.write """unarrange_ffb"" : ""0"","
		response.write """unprune_frond"" : ""0"","
		response.write """qe_1"" : ""0"","
		response.write """qe_2"" : ""0"","
		response.write """qe_3"" : ""0"","
		response.write """qe_4"" : ""0"","
		response.write """qe_5"" : ""0"","
		response.write """qe_6"" : ""0"","
		response.write """qe_7"" : ""0"","
		response.write """qe_8"" : ""0"","
		response.write """qe_9"" : ""0"","
		response.write """qe_10"" : ""0"""
	end if
else
	
	response.write """isjobpanen"" : ""0"""
	
end if 'Cek Job Panen atau bukan
response.write "}"
%>