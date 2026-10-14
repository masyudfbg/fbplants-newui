<%@ Language=VBScript %>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<%
Response.ContentType = "application/json"
Response.Buffer = True
Response.Clear

Dim idview, transdate, kecambahqty, pnqty, mnqty
Dim persiapanamt, pnamt, mnamt, fcba, company
Dim fcentry, fcedit

idview       = Request("idview")
transdate    = Request("transdate")
kecambahqty  = Request("kecambahqty")
pnqty        = Request("pnqty")
mnqty        = Request("mnqty")
persiapanamt = Request("persiapanamt")
pnamt        = Request("pnamt")
mnamt        = Request("mnamt")
fcba         = Request("fcba")
company      = Request("company")
fcentry      = Session("CodeName")
fcedit       = Session("CodeName")

If idview = "46" Then

    Dim sqlCheck, sqlSave, cnt, rsCheck, sqlSaveAction

    sqlCheck = "SELECT COUNT(1) AS CNT FROM FAC_HPP_HARIAN_BIBIT " & _
               "WHERE FCBA      = '" & fcba & "' "                  & _
               "AND   TRANSDATE = TO_DATE('" & transdate & "','DD/MM/YYYY')"

    cnt = 0

    On Error Resume Next
    Set rsCheck = connStr.Execute(sqlCheck)
    If Err.Number <> 0 Then
        Response.Write "{""status"":""ERROR"",""message"":""Gagal cek data: " & Replace(Err.Description, """", "'") & """}"
        Response.End
    End If
    On Error GoTo 0

    If Not rsCheck.EOF Then
        cnt = CInt(rsCheck("CNT"))
    End If
    rsCheck.Close
    Set rsCheck = Nothing

    If cnt > 0 Then
        sqlSave = "UPDATE FAC_HPP_HARIAN_BIBIT SET "            & _
                  "  KECAMBAHQTY  = " & kecambahqty             & ", " & _
                  "  PNQTY        = " & pnqty                   & ", " & _
                  "  MNQTY        = " & mnqty                   & ", " & _
                  "  PERSIAPANAMT = " & persiapanamt            & ", " & _
                  "  PNAMT        = " & pnamt                   & ", " & _
                  "  MNAMT        = " & mnamt                   & ", " & _
                  "  FCEDIT       = '" & fcedit                 & "', " & _
                  "  LASTUPDATE   = SYSDATE "                   & _
                  "WHERE FCBA      = '" & fcba                  & "' " & _
                  "AND   TRANSDATE = TO_DATE('" & transdate     & "','DD/MM/YYYY')"
        sqlSaveAction = "UPDATE"
    Else
        sqlSave = "INSERT INTO FAC_HPP_HARIAN_BIBIT "            & _
                  "( TRANSDATE, KECAMBAHQTY, PNQTY, MNQTY, "    & _
                  "  PERSIAPANAMT, PNAMT, MNAMT, "               & _
                  "  FCENTRY, FCEDIT, FCBA, LASTUPDATE ) "       & _
                  "VALUES ( "                                     & _
                  "  TO_DATE('" & transdate & "','DD/MM/YYYY'), " & _
                  "  " & kecambahqty  & ", "                     & _
                  "  " & pnqty        & ", "                     & _
                  "  " & mnqty        & ", "                     & _
                  "  " & persiapanamt & ", "                     & _
                  "  " & pnamt        & ", "                     & _
                  "  " & mnamt        & ", "                     & _
                  "  '" & fcentry     & "', "                    & _
                  "  '" & fcedit      & "', "                    & _
                  "  '" & fcba        & "', "                    & _
                  "  SYSDATE )"
        sqlSaveAction = "INSERT"
    End If

    On Error Resume Next
    connStr.Execute sqlSave
    If Err.Number <> 0 Then
        Response.Write "{""status"":""ERROR"",""message"":""Gagal " & sqlSaveAction & ": " & Replace(Err.Description, """", "'") & """}"
    Else
        Response.Write "{""status"":""OK"",""action"":""" & sqlSaveAction & """}"
    End If
    On Error GoTo 0

Else
    Response.Write "{""status"":""ERROR"",""message"":""idview tidak dikenal""}"
End If
%>