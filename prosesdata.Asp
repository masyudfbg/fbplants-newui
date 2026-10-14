<%validatesession%>
<!--#include file="connDb1.Inc"-->
<%
	Dim sqlsimpan,prosespricetbs
	Dim ondate,supplier,product,price 
	Dim StartDate,EndDate,intStep,ts, i, tgl

	prosespricetbs 	= Request("prosesdata")
	ondate 			= Request("ondate")
	StartDate 		= Request("ondate")
	EndDate 		= DateAdd("m", 1, StartDate)
	supplier 		= Request("supplier")
	product 		= Request("product")
	price 			= Request("price")
	fcbases			= session("fcba")
	userid			= session("CodeName")
	
	if prosespricetbs = "prosespricetbs" then 
		sqlsimpan =	"MERGE INTO INVOICE_PRICETBS a " & _
		" USING   ( " & _
		"		SELECT (to_date('" & ondate & "','MM/DD/YYYY') + LEVEL - 1 ) AS DAY,'" & supplier & "' Supplier,'" & price & "' Newprice FROM dual CONNECT BY 30 >= LEVEL " & _
		"		)b " & _
		" ON (To_Char(a.ondate,'yyyymmdd')=To_Char(b.DAY,'yyyymmdd') AND a.suppliercode=b.supplier) " & _
		" WHEN MATCHED THEN " & _
		"		UPDATE SET a.price = b.Newprice  " & _
		"		WHERE a.price <> b.Newprice " & _
		" WHEN NOT MATCHED THEN " & _
		"		INSERT (ondate,suppliercode,productcode,plantingyear,price,fcentry,fcedit,fcip,fcba,lastupdate,lasttime) " & _
		"		VALUES (b.day,b.supplier,'TBS',0,b.Newprice,'DUMMY','DUMMY','0','" & fcbases & "',SYSDATE,'00')"
		connStr.execute(sqlsimpan)
	end if
	Response.redirect("pricetbs.Asp")
	
%>
<%
' function dibawah ini tidak dipakai
'	for i = 0 to 30

	'		tgl = DateAdd("d", i, StartDate)
	'		sqlsimpan = "insert into INVOICE_PRICETBSTMP2(ondate,suppliercode,productcode,plantingyear,price,fcentry,fcedit,fcip,fcba,lastupdate,lasttime) " & _
	'			" values(to_Date('" & tgl & "','mm/dd/yyyy'),'" & supplier & "','" & product & "','0', '" & price & "', " & _ 
	'			"  '" & userid & "', '" & userid & "','" & Request.ServerVariables("REMOTE_ADDR") & "','" & fcbases & "',SYSDATE,TO_CHAR(SYSDATE, 'HH24:MI')) "
	'			connStr.execute(sqlsimpan)

			'Response.Write tgl & "<br />"

	'	next
	%>