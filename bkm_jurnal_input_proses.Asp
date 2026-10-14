<%validatesession%>
<!--#include file="connDb1.Inc"-->
<html>
<body>
<%

			dim sqldtl,sqldeletehdr,sqldeletedtl,sqlhdr
			dim nojournal,tglbkm,remarks,POSTINGFLG,fcba,view,sourcecode,tipetranScriptName
			ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

			nojournal  = Request("jurnalno")
			tglbkm     = Request("jurnaldate")
			remarks    = Request("remarks")
			POSTINGFLG = Request("POSTINGFLG")
			fcba       = Request("fcba")
			view       = Request("view")
			sourcecode = Request("sourcecode")
			sqldeletedtl = ""
			tipetran   = Request("tipetran")
			if remarks = "" then 
				remarks = "-"
			end if
			
			'response.write "XXX " & nojournal & " | " & sourcecode & " |" & view & "<br>"
			
	IF len(trim(nojournal)) > 0 then
			if POSTINGFLG = "0" then

			' delete dulu data pada tabel gledgerhdr'
			'sqldeletehdr = "delete from GLEDGERHDR where docno in (select docno from gledgerdtl where JOURNALNO='" & nojournal & "')"
			'connStr.execute(sqldeletehdr)

			' delete dulu data pada tabel gledgerdtl'
			'sqldeletedtl = "delete from gledgerdtl where JOURNALNO='" & nojournal & "'"
			'connStr.execute(sqldeletedtl)


			


			' insert ke tabel gl detail'
			'sqldtl = "INSERT INTO gledgerdtl " & _
			'" SELECT 'GL/'||a.fcba||'/'||To_Char(a.docdate,'yyyymm')||'/'||jobcode, " & _
			'" a.docno,a.docdate,jobcode,0,Sum(amountdebit),Sum(amountkredit),0,SYSDATE,SYSDATE,'IT_HO1','IT_HO1','::1',a.FCBA ,'1'" & _
			'" FROM bkmjournalhdr a INNER JOIN bkmjournaldtl b ON a.docno=b.docno AND a.fcba=b.fcba " & _
			'" WHERE a.docno='" & nojournal & "' " & _
			'" GROUP BY 'GL/'||a.fcba||'/'||To_Char(a.docdate,'yyyymm'), a.docno,a.docdate,jobcode,a.fcba " & _
			'" ORDER BY jobcode,docdate"
			'connStr.execute(sqldtl)


			' insert ke tabel gl header
			'sqlhdr ="MERGE INTO GLEDGERHDR  x " & _
			'" USING (   " & _
			'" SELECT D.DOCNO,D.JOURNALNO,TO_CHAR(D.JOURNALDATE  ,'MM') AS BULAN,  " & _
			'" TO_CHAR(D.JOURNALDATE ,'YYYY') AS TAHUN,D.COA,NVL(D.OPENING_BALANCE,0) AS OPENING_BALANCE,nvl(D.DEBET,0) AS DEBET ,  " & _
			'" nvl(D.CREDIT,0) AS CREDIT ,NVL(D.OPENING_BALANCE,0) + nvl(D.DEBET,0) + nvl(D.CREDIT,0) AS ENDING_BALANCE ,D.JOURNALDATE,D.FCBA " & _
			'" FROM GLEDGERDTL  D  " & _
			'" WHERE D.JOURNALNO ='" & nojournal & "' " & _
			'" ) y   " & _
			'" ON (x.DOCNO=y.DOCNO and x.fcba=y.fcba and x.PERIOD_MONTH =to_number(y.BULAN) and x.PERIOD_YEAR =to_number(y.TAHUN))  " & _
			'" WHEN MATCHED THEN  " & _
			'" UPDATE SET x.DEBET  = y.DEBET,  " & _
			'"		x.CREDIT  = y.CREDIT  " & _
			'" WHERE x.DOCNO=y.DOCNO " & _
			'" WHEN NOT MATCHED THEN  " & _
			'" Insert (DOCNO, PERIOD_MONTH, PERIOD_YEAR, COA, OPENING_BALANCE,DEBET, CREDIT, ENDING_BALANCE, CREATEDATE, LASTUPDATE, " & _
			'" 	FCENTRY, FCEDIT, FCIP, FCBA)  " & _
			'" VALUES(y.DOCNO, to_number(y.BULAN), to_number(y.TAHUN),  y.COA, y.OPENING_BALANCE,y.DEBET,y.CREDIT,y.ENDING_BALANCE,   " & _
			'" y.JOURNALDATE,SYSDATE,'SYSTEM','SYSTEM','1:::',y.fcba)"
			
			'connStr.execute(sqlhdr)

			' update postingflg'
			'if sourcecode = "OPR-BKM" Then
			'sqldeletedtl = "update bkmhdr set postingflg=1 where fcba='" & fcba & "' AND bkmno in(select bkmno from bkmjournaldtl where fcba='" & fcba & "' and docno='" & nojournal & "')"
			'connStr.execute(sqldeletedtl)
			'end if

			'if sourcecode = "OPR-WIP" Then
			'sqldeletedtl = "update bkmhdr set postingflg=1 where fcba='" & fcba & "' AND bkmno in(select bkmno from bkmjournaldtl where fcba='" & fcba & "' and docno='" & nojournal & "')"
			'connStr.execute(sqldeletedtl)
			'end if
			
			if sourcecode = "FAC-JRV" Then
			sqldeletedtl = "update PAYMENTVOUCHER set postingflg='1',ROWSTATEID='Approved',lastupdate=sysdate,fcedit='" & session("CodeName") & "' where fcba='" & fcba & "' AND vouchercode='" & nojournal & "'"
			connStr.execute(sqldeletedtl)
			end if
			
			if sourcecode = "FAC-BNK" Then
			sqldeletedtl = "update PAYMENTVOUCHER set postingflg='1',ROWSTATEID='Approved',lastupdate=sysdate,fcedit='" & session("CodeName") & "' where fcba='" & fcba & "' AND vouchercode='" & nojournal & "'"
			connStr.execute(sqldeletedtl)
			end if
			
			if sourcecode = "FAC-IPO" Then
			sqldeletedtl = "update invoice set postingflg='1',lastupdate=sysdate,fcedit='" & session("CodeName") & "' where fcba='" & fcba & "' AND invoicecode='" & nojournal & "'"
			connStr.execute(sqldeletedtl)
			end if
			
			if sourcecode = "FAC-ISP" Then
			sqldeletedtl = "update invoiceservicehdr set postingflg='1',lastupdate=sysdate,fcedit='" & session("CodeName") & "' where fcba='" & fcba & "' AND invoiceno='" & nojournal & "'"
			connStr.execute(sqldeletedtl)
			end if
			
			if sourcecode = "FAC-IFB" Then
			sqldeletedtl = "update invoiceservicehdr set postingflg='1',lastupdate=sysdate,fcedit='" & session("CodeName") & "' where fcba='" & fcba & "' AND invoiceno='" & nojournal & "'"
			connStr.execute(sqldeletedtl)
			end if
			
			if sourcecode = "FAC-ITR" Then
			sqldeletedtl = "update transporterinvoice set postingflg='1',lastupdate=sysdate,fcedit='" & session("CodeName") & "' where fcba='" & fcba & "' AND invoicecode='" & nojournal & "'"
			connStr.execute(sqldeletedtl)
			end if
			
			if sourcecode = "FAC-ICO" Then
			sqldeletedtl = "update contractorinvoicehdr set postingflg='1',lastupdate=sysdate,fcedit='" & session("CodeName") & "' where fcba='" & fcba & "' AND invoicecode='" & nojournal & "'"
			connStr.execute(sqldeletedtl)
			end if
			
			if sourcecode = "FAC-IST" Then
			sqldeletedtl = "update invoicesto set postingflg='1',lastupdate=sysdate,fcedit='" & session("CodeName") & "' where fcba='" & fcba & "' AND invoicecode='" & nojournal & "'"
			connStr.execute(sqldeletedtl)
			end if
			
			
			'response.write "1-" & sqldeletedtl & "<br>"

			' update postingflg pada tabel bkmjournalhdr'
			sqldeletedtl = "update BKMJOURNALHDR set postingflg=1,lastupdate=sysdate,fcedit='" & session("CodeName") & "' where fcba='" & fcba & "' AND docno='" & nojournal & "'"
			'connStr.execute(sqldeletedtl)
			ExecuteAndLog ScriptName, sqldeletedtl

			sqldeletedtl = "update BKMJOURNALHDR set postingflg=1,lastupdate=sysdate,fcedit='" & session("CodeName") & "' where docno='" & nojournal & "-INT'"
			'connStr.execute(sqldeletedtl)
			ExecuteAndLog ScriptName, sqldeletedtl

            'response.write "2-" &sqldeletedtl & "<br>"
			else 

			' delete dulu data pada tabel gledgerhdr'
			'sqldeletehdr = "delete from GLEDGERHDR where docno in (select docno from gledgerdtl where JOURNALNO='" & nojournal & "')"
			'connStr.execute(sqldeletehdr)

			' delete dulu data pada tabel gledgerdtl'
			'sqldeletedtl = "delete from gledgerdtl where JOURNALNO='" & nojournal & "'"
			'connStr.execute(sqldeletedtl)

			' update postingflg'
			'if sourcecode = "OPR-BKM" Then
			'sqldeletedtl = "update bkmhdr set postingflg=0 where fcba='" & fcba & "' AND bkmno in(select bkmno from bkmjournaldtl where fcba='" & fcba & "' AND docno='" & nojournal & "')"
			'connStr.execute(sqldeletedtl)
			'end if

			'if sourcecode = "OPR-WIP" Then
			'sqldeletedtl = "update bkmhdr set postingflg=0 where fcba='" & fcba & "' AND bkmno in(select bkmno from bkmjournaldtl where fcba='" & fcba & "' AND docno='" & nojournal & "')"
			'connStr.execute(sqldeletedtl)
			'end if

            if sourcecode = "FAC-JRV" Then
			sqldeletedtl = "DELETE BKMJOURNALDTL where SOURCE='JRV' AND fcba='" & fcba & "' AND docno='" & nojournal & "'"
			connStr.execute(sqldeletedtl)

			sqldeletedtl = "DELETE BKMJOURNALHDR where SOURCECODE='FAC-JRV' AND fcba='" & fcba & "' AND docno='" & nojournal & "'"
			connStr.execute(sqldeletedtl)

			sqldeletedtl = "DELETE PAYMENTVOUCHER_HIST where fcba='" & fcba & "' AND vouchercode='" & nojournal & "'"
			connStr.execute(sqldeletedtl)

			sqldeletedtl = "update PAYMENTVOUCHER set postingflg='0',ROWSTATEID='Confirmed',lastupdate=sysdate,fcedit='" & session("CodeName") & "'  where fcba='" & fcba & "' AND vouchercode='" & nojournal & "'"
			connStr.execute(sqldeletedtl)
			end if
			
			if sourcecode = "FAC-BNK" Then
			sqldeletedtl = "DELETE BKMJOURNALDTL where SOURCE in ('PYV','RCV') AND fcba='" & fcba & "' AND docno='" & nojournal & "'"
			connStr.execute(sqldeletedtl)

			sqldeletedtl = "DELETE BKMJOURNALHDR where SOURCECODE='FAC-JRV' AND fcba='" & fcba & "' AND docno='" & nojournal & "'"
			connStr.execute(sqldeletedtl)

			sqldeletedtl = "DELETE PAYMENTVOUCHER_HIST where fcba='" & fcba & "' AND vouchercode='" & nojournal & "'"
			connStr.execute(sqldeletedtl)

			sqldeletedtl = "update PAYMENTVOUCHER set postingflg='0',ROWSTATEID='Confirmed',lastupdate=sysdate,fcedit='" & session("CodeName") & "'  where fcba='" & fcba & "' AND vouchercode='" & nojournal & "'"
			connStr.execute(sqldeletedtl)
			end if
			
			if sourcecode = "FAC-IPO" Then
			sqldeletedtl = "update invoice set postingflg='0',lastupdate=sysdate,fcedit='" & session("CodeName") & "' where fcba='" & fcba & "' AND invoicecode='" & nojournal & "'"
			connStr.execute(sqldeletedtl)
			end if
			
			if sourcecode = "FAC-ISP" Then
			sqldeletedtl = "update invoiceservicehdr set postingflg='0',lastupdate=sysdate,fcedit='" & session("CodeName") & "' where fcba='" & fcba & "' AND invoiceno='" & nojournal & "'"
			connStr.execute(sqldeletedtl)
			end if
			
			if sourcecode = "FAC-IFB" Then
			sqldeletedtl = "update invoiceservicehdr set postingflg='0',lastupdate=sysdate,fcedit='" & session("CodeName") & "' where fcba='" & fcba & "' AND invoiceno='" & nojournal & "'"
			connStr.execute(sqldeletedtl)
			end if
			
			if sourcecode = "FAC-ITR" Then
			sqldeletedtl = "update transporterinvoice set postingflg='0',lastupdate=sysdate,fcedit='" & session("CodeName") & "' where fcba='" & fcba & "' AND invoicecode='" & nojournal & "'"
			connStr.execute(sqldeletedtl)
			end if
			
			if sourcecode = "FAC-ICO" Then
			sqldeletedtl = "update contractorinvoicehdr set postingflg='0',lastupdate=sysdate,fcedit='" & session("CodeName") & "' where fcba='" & fcba & "' AND invoicecode='" & nojournal & "'"
			connStr.execute(sqldeletedtl)
			end if
			
			if sourcecode = "FAC-IST" Then
			sqldeletedtl = "update invoicesto set postingflg='0',lastupdate=sysdate,fcedit='" & session("CodeName") & "' where fcba='" & fcba & "' AND invoicecode='" & nojournal & "'"
			connStr.execute(sqldeletedtl)
			end if
			
			'response.write "1-" & sqldeletedtl & "<br>"
			
			' update postingflg pada tabel bkmjournalhdr'
			sqldeletedtl = "update BKMJOURNALHDR set postingflg=0,lastupdate=sysdate,fcedit='" & session("CodeName") & "' where fcba='" & fcba & "' AND  docno='" & nojournal & "'"
			connStr.execute(sqldeletedtl)
			
			'response.write "2-" & sqldeletedtl & "<br>"

			end if
			
			sqlhdr = "BEGIN fac_closing_accounting_api.ProsesClosing(TO_NUMBER(TO_CHAR(TO_DATE('" & tglbkm & "','DD/MM/YYYY'),'YYYY')), " & _ 
			                                                               "TO_NUMBER(TO_CHAR(TO_DATE('" & tglbkm & "','DD/MM/YYYY'),'MM')) , '" & fcba & "','System'); " & _ 
				     "END;"													 
			'response.write(	sqlhdr)																 
			connStr.execute(sqlhdr)
    END IF
	        Select case view
				case "0" : response.redirect("bkm_jurnal_input.asp?tglbkm="&tglbkm&"&fcba=" & fcba)
				case "1" : response.redirect("bkm_jurnal_view.asp?tglbkm="&tglbkm&"&fcba=" & fcba & "&bkmno=" & nojournal & "&tipetran=" & sourcecode  )
				case "2" : response.redirect("bkm_jurnalinv_view.asp?tglbkm="&tglbkm&"&fcba=" & fcba & "&bkmno=" & nojournal&"&tipetran=" & tipetran)
			End Select
%>

</body>
</html>