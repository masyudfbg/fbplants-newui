<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
arrdtl_eof = true
accountingyear = Request("accountingyear")
accountingperiod = Request("accountingperiod")
q = " SELECT rt.transactionno," & _
"        rt.documentno," & _
"        rt.currency," & _
"        rt.amount totalamount," & _
"        rt.status," & _
"        rd.debetcredit," & _
"        rd.jobcode," & _
"        rd.targettype," & _
"        CASE" & _
"            WHEN rt.settlementno IS NULL AND rt.transactionno = xx.transactionno THEN rd.amountendinstallment" & _
"            WHEN rt.settlementno IS NOT NULL AND rt.transactionno = xx.transactionno THEN rsd.amountendinstallment" & _
"            WHEN rt.settlementno IS NULL THEN rd.amountmonth" & _
"            WHEN rt.settlementno IS NOT NULL THEN rsd.amountmonth" & _
"        ELSE 0" & _
"        END amountmonth" & _
" FROM recurringtransaction rt" & _
" INNER JOIN recurringdtl rd ON (rt.fcba = rd.fcba AND rt.documentno = rd.documentno AND rt.accountingyear = rd.accountingyear)" & _
" LEFT JOIN (SELECT fcba, documentno, MAX(transactionno) transactionno FROM recurringtransaction WHERE accountingyear = " & accountingyear & " AND fcba = '" & Session("fcba") & "' GROUP BY fcba, documentno) xx ON (rt.fcba = xx.fcba AND rt.documentno = xx.documentno)" & _
" LEFT JOIN recurringsettlement rs ON (rt.fcba = rs.fcba AND rt.documentno = rs.documentno AND rt.accountingyear = rs.accountingyear)" & _
" LEFT JOIN recurringsettlementdtl rsd ON (rs.fcba = rsd.fcba AND rs.settlementno = rsd.settlementno AND rd.jobcode = rsd.jobcode AND rd.targettype = rsd.targettype)" & _
" WHERE rt.accountingyear = " & accountingyear & "" & _
"       AND rt.accountingperiod = " & accountingperiod & "" & _
"       AND rt.fcba = '" & Session("fcba") & "'" & _
" ORDER BY rt.transactionno, rd.jobcode"


'----OLD QUERY----'

' q = " SELECT rt.transactionno," & _
	' " rt.documentno," & _
	' " rt.currency," & _
	' " rt.amount totalamount," & _
	' " rt.status," & _
	' " rd.debetcredit," & _
	' " rd.jobcode," & _
	' " rd.targettype," & _
	' " DECODE(rt.transactionno, xx.transactionno, rd.amountendinstallment, rd.amountmonth) amountmonth" & _
	' " FROM recurringtransaction rt" & _
	' " INNER JOIN recurringdtl rd" & _
	' " ON (rt.fcba = rd.fcba AND rt.documentno = rd.documentno AND rt.accountingyear = rd.accountingyear)" & _
	' " LEFT JOIN (SELECT fcba, documentno, MAX(transactionno) transactionno FROM recurringtransaction WHERE accountingyear = " & accountingyear & " AND fcba = '" & Session("fcba") & "' GROUP BY fcba, documentno) xx ON (rt.fcba = xx.fcba AND rt.documentno = xx.documentno)" & _
	' " WHERE rt.accountingyear = " & accountingyear & "" & _
	' "       AND rt.accountingperiod = " & accountingperiod & "" & _
	' "       AND rt.fcba = '" & Session("fcba") & "'" & _
	' " ORDER BY rt.transactionno, rd.jobcode"

'response.Write q
set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr

if Not rsrecurtrans.eof then
	arrdtl_data = rsrecurtrans.getrows()
	arrdtl_eof = false
end if
rsrecurtrans.Close
set rsrecurtrans = Nothing

p_totalamount = 0
Function f_totalamount(data_array__, where__)
	result_ = 0
	arr_1 = data_array__
	for indx = 0 to UBound(arr_1, 2)
		'response.write arr_1(0, indx)
		if Cstr(arr_1(0, indx)) = Cstr(where__) then
			result_ = result_ + Cdbl(To_Nol(arr_1(8, indx)))
		end if
	next
	f_totalamount = result_
End Function

'--JSON Data Detail
IF NOT arrdtl_eof THEN
	prev_transno = ""
	row_data_dtl = ""
	For i = LBound(arrdtl_data, 2) To UBound(arrdtl_data, 2)
		new_transno = Cstr(arrdtl_data(0, i))
		IF new_transno <> prev_transno THEN
			p_totalamount = f_totalamount(arrdtl_data, new_transno)
			IF i = 0 THEN
				row_data_dtl = row_data_dtl & "{""documentno"":"""&arrdtl_data(1, i)&"""," & _
                                              """transactionno"":"""&arrdtl_data(0, i)&"""," & _
                                              """currency"":"""&arrdtl_data(2, i)&"""," & _
                                              """totalamount"":"""&p_totalamount&"""," & _
                                              """status"":"""&arrdtl_data(4, i)&"""," & _
                                              """detail"":[{" & _
                                              """debetkredit"":"""&arrdtl_data(5, i)&"""," & _
                                              """jobcode"":"""&arrdtl_data(6, i)&"""," & _
                                              """targettype"":"""&arrdtl_data(7, i)&"""," & _
                                              """amount"":"""&arrdtl_data(8, i)&"""}"
			ELSE
				row_data_dtl = row_data_dtl & "]},{""documentno"":"""&arrdtl_data(1, i)&"""," & _
					"""transactionno"":"""&arrdtl_data(0, i)&"""," & _
					"""currency"":"""&arrdtl_data(2, i)&"""," & _
					"""totalamount"":"""&p_totalamount&"""," & _
					"""status"":"""&arrdtl_data(4, i)&"""," & _
					"""detail"":[{" & _
					"""debetkredit"":"""&arrdtl_data(5, i)&"""," & _
					"""jobcode"":"""&arrdtl_data(6, i)&"""," & _
					"""targettype"":"""&arrdtl_data(7, i)&"""," & _
					"""amount"":"""&arrdtl_data(8, i)&"""}"
			END IF
		ELSE
			row_data_dtl = row_data_dtl	& ",{" & _
				"""debetkredit"":"""&arrdtl_data(5, i)&"""," & _
				"""jobcode"":"""&arrdtl_data(6, i)&"""," & _
				"""targettype"":"""&arrdtl_data(7, i)&"""," & _
				"""amount"":"""&arrdtl_data(8, i)&"""}"
		END IF
		IF i = UBound(arrdtl_data, 2) THEN
			row_data_dtl = row_data_dtl	& "]}"
		END IF
		prev_transno = new_transno
	Next			
	post_dtl = "{""data"":[" & row_data_dtl & "]}"
ELSE
	post_dtl = "{""data"":[]}"
END IF

response.write post_dtl
%>