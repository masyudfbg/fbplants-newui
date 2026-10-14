<%
v_Tahun = Cdbl( Left( Session("CurP_Costbook"), 4))
v_Bulan = Cdbl( Right( Session("CurP_Costbook"), 2))


Response.Write "Start Normalization for Fico ..., Year :" & v_Tahun & ", month : " &  v_Bulan
connstr.BeginTrans
strSQLNormalization = " delete erp_report_sap_generated " & _
                                  " where fcba = '" & Session("fcBA") & "' "  & _
                                  "          and accountingperiod = " & v_Bulan  &  _
                                  "          and accountingyear= " & v_Tahun

connStr.execute(strSQLNormalization)
If Err.Number = 0 Then
	connstr.CommitTrans
	Response.Write "(DONE)" & "<br>"
	Response.Flush
Else
	connstr.RollbackTrans
	Response.write("Error Save ["&strSQLNormalization&"] ("&Err.Description&")")
	Response.Flush
End If


Response.Write "Start Normalization for Sales and Distribution Data 01 ..., Year :" & v_Tahun & ", month : " &  v_Bulan   
connstr.BeginTrans
strSQLNormalization = " delete erp_report_sap_sd_generated " & _
                                  " where fcba = '" & Session("fcBA") & "' "  & _
                                  "          and to_number(to_Char( to_date(delivery_date,'dd.mm.yyyy'),'mm'))= " & v_Bulan  &  _
                                  "          and to_number(to_Char( to_date(delivery_date,'dd.mm.yyyy'),'yyyy'))= " & v_Tahun

connStr.execute(strSQLNormalization)
If Err.Number = 0 Then
	connstr.CommitTrans
	Response.Write "(DONE)" & "<br>"
	Response.Flush
Else
	connstr.RollbackTrans
	Response.write("Error Save ["&strSQLNormalization&"] ("&Err.Description&")")
	Response.Flush
End If


Response.Write "Start Normalization for Sales and Distribution Data 02... , Year :" & v_Tahun & ", month : " &  v_Bulan  
connstr.BeginTrans
strSQLNormalization = " delete erp_report_sap_sd01_generated " & _
                                  " where fcba = '" & Session("fcBA") & "' "  & _
                                  "          and to_number(to_Char( to_date(post_date,'dd.mm.yyyy'),'mm'))= " & v_Bulan  &  _
                                  "          and to_number(to_Char( to_date(post_date,'dd.mm.yyyy'),'yyyy'))= " & v_Tahun

connStr.execute(strSQLNormalization)
If Err.Number = 0 Then
	connstr.CommitTrans
	Response.Write "(DONE)" & "<br>"
	Response.Flush
Else
	connstr.RollbackTrans
	Response.write("Error Save ["&strSQLNormalization&"] ("&Err.Description&")")
	Response.Flush
End If


Response.Write "Start Normalization for Sales and Distribution Data 03... ,  Year :" & v_Tahun & ", month : " &  v_Bulan  
connstr.BeginTrans
strSQLNormalization = " delete erp_report_sap_sd02_generated " & _
                                  " where fcba = '" & Session("fcBA") & "' "  & _
                                  "          and to_number(to_Char( delivery_date,'MM'))= " & v_Bulan  & "" & _
                                  "          and to_number(to_Char( delivery_date,'YYYY'))= " & v_Tahun

connStr.execute(strSQLNormalization)
If Err.Number = 0 Then
	connstr.CommitTrans
	Response.Write "(DONE)" & "<br>"
	Response.Flush
Else
	connstr.RollbackTrans
	Response.write("Error Save ["&strSQLNormalization&"] ("&Err.Description&")")
	Response.Flush
End If

%>