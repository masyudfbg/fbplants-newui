<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>ePlaS - Laporan Harian Produksi</title>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<script language="JavaScript" src='/js/x.js'></script>
<script language="JavaScript" src='/js/xtableheaderfixed.js'></script>
<script language="JavaScript">
xAddEventListener(window, 'load',
  function() {
    new xTableHeaderFixed('xthf-site', window);
    new xTableHeaderFixed('xthf-green', window);
    new xTableHeaderFixed('xthf-blue', window);
  }, false
);
</script>
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>

</head>
<style>
	.posisikirijudul{
		position: relative;
		left: 0px;
	}
	.posisikiriisi{
		position: relative;
		left: 0px;
	}
	.lebarjamolah{
		width: 70px;
	}
	.lebartitik{
		width: 1px;
	}
	.warnagaris{
		border-color: #FFA800;
	}
	.tebal{
		font-weight: bold;
	}
	.ratakiri{
		text-align: left;
	}
	.ratakanan{
		text-align: right;
	}
	.ratatengah{
		text-align: center;
	}
    .element {
        position: relative;
        top: -75px;
        left: -6px;
    }
    .lebarkolomnodo{
		width: 340px;
	}
    .lebarkolompartydo{
		width: 90px;
	}
    .lebarkolomhi{
		width: 70px;
	}
    .lebarkolomshi{
		width: 100px;
	}
    .lebarkolomsisado{
		width: 50px;
	}
    .lebarkolomlhi{
		width: 50px;
	}
     .lebarkolomlurai{
		width: 160px;
	}
    .lebarkolomlshi{
		width: 50px;
	}
    .lebarkolomlshi{
		width: 50px;
	}
    .tabeltengah{
        margin: 20px auto;

    }
</style>
<%
Dim sqltgl, sqltbs,rsdata,sqlsimpan,sqlsimpanhk,sqlbulking,sqltankcpo
Dim tglblnawal,tglthnawal,tglconfig
Dim tglnow,tglprev,fcba,tglsplit,bulansplit,thnsplit,tglbulanprev,tahunprev
Dim tbsterimakemarin,sisatbskemarin,tbsnow,tbstersedia,tbssisanow,tbsolah,jamkerja
Dim cpoffa,cpoair,cpokotoran,cpodobi,cpoffashi,cpoairshi,cpokotoranshi,cpodobishi
Dim cpokemarin,cpotankkemarin,totalcpokemarin,pktankkemarin,totalpkkemarin,ckgkemarin,totalckgkemarin
Dim tbsdhariini,tbssdblnini,tbssdthnini,tbsolahshi,tbsolahsbi,tbsolahsti,jamolahshi,jamolahsbi,jamolahsti
Dim kapasitasjam,kapasitasshi,kapasitassbi,kapasitassti
Dim qtystockcpohi,qtystockcposhi,qtystockcposbi,qtystockcposti,cpohriini
Dim oerhi,oershi,oersbi,oersti,kerhi,kershi,kersbi,kersti   
Dim pkhi,qtypkhi,qtypkshi,qtypksbi,qtypksti
Dim pkffahi,pkairhi,pkkotoranhi,pkffashi,pkairshi,pkkotoranshi
Dim cpodo,cpodoparty,cpodohi,cpodoshi,cpodosisa
Dim pkdo,pkdoparty,pkdohi,pkdoshi,pkdosisa
Dim ckgdo,ckgdoparty,ckgdohi,ckgdoshi,ckgdosisa
Dim xbln1,xbln2,xthn1,xthn2
Dim blnsebelum,tglsebelum2,sqlcpodo,rscpodo
Dim totcpodoparty,totcpodohi,totcpodoshi,totcpodosisa
Dim totpkdoparty,totpkdohi,totpkdoshi,totpkdosisa
Dim totckgdoparty,totckgdohi,totckgdoshi,totckgdosisa
Dim ckghi,qtyckghi,qtyckgshi,qtyckgsbi,qtyckgsti,ckgviewhi
Dim cpokirimhi,cpocothi,cpocsthi,cpoothi,cpotankisihi,kernelkrmhi,kernelsilohi,pktankisihi,cangkangkrmhi,cangkangprodhi
Dim cpokirimkemarin,cpocotkemarin,cpocstkemarin,cpootkemarin,cpotkkemarin,pktkkemarin
Dim hslcpokemarin,hslcpohi,hslpkkemarin
Dim tawal,takhir,lastyear,lastyearsplit,lastmontsplit,lastdaysplit,awalthn_prev,akhir_tahunprev
Dim gettotaltbsshi,getcpokirimshi,getcpocotshi,getcpocstshi,getcpootshi,gettankisishi
Dim getpkkirimshi,getpksiloshi,getpktankshi
Dim getkirimckgshi,getkirimckgsbi,getkirimckgsti,getproduksickgshi
Dim getcpokirimld,getcpocotld,getcpocstld,getcpootld,gettankisild,getpksilold,gettankpkisild,ckgkirimld,ckgproduksild
Dim getcpokirimsbi,getcpocotsbi,getcpocstsbi,getcpootsbi,gettankisisbi
Dim getpkkirimsbi,getisipksbi,getsilopksbi
Dim getcpokirimyd,getcpocotyd,getcpocstyd,getcpootyd,gettankisiyd
Dim gettbsolahyd,pkkirimyd,pktankisiyd,pksiloyd
Dim prev_year_max_date,hslcpoblnkemarin,hslcpobulanini,hslpkbulanini,hslpkblnkemarin
Dim getcpokirimsti,getcpocotsti,getcpocststi,getcpootsti,gettankisisti
Dim getpkkirimsti,getpksilosti,getpktanksti
Dim tglresult,tbsbf,tbsterimashi,tbsterimasbi,tbsterimasti
Dim tbssisahi,gettbsolahsbi,gethasiltbsolahshi,gettbsterimasbi,gettbsterimasti,gethasiltbsolahsti
Dim jammulaitxt,jamselesaitxt,keterangantxt,airtxt,caco3txt,ttlairtxt,ttlcaco3txt
Dim ratioairshi,ratiocaco3hi,ratiocaco3shi,stockckgtxt,cpoadjhi,cpoadjshi,cpoadjsbi,cpoadjsti
Dim pottbshi,pottbsshi,pottbssbi,pottbssti


    

    tglnow      = request("lhpdate_From")
    fcba        = Session("fcba")
    bulansplit  = left(tglnow,2)
    tglsplit    = mid(tglnow,4,2)
    thnsplit    = right(tglnow,4)
    tglblnawal  = bulansplit & "/" & "01" & "/" & thnsplit
    tglthnawal  = "01" & "/" & "01" & "/" & thnsplit
  
    xbln1       = bulansplit
    xthn1       = thnsplit
    if  xbln1   = 01 then
        xbln2   = 12
        xthn2   = xthn1-1
    else   
        xbln2   = xbln1-01
        xthn2   = xthn1
    end if

    'tglbulanprev =  xbln2 & "-" & tglsplit & "-" & xthn2
   
   
     ' GET DATE CONFIG PRODUCTION  PREV'
    sqltgl = "SELECT PRODUCTIONMILLDATE FROM CONFIGMST " & _
    " WHERE FCBA='" & fcba & "'"
    'response.write sqltgl
    Set rsdata = Server.CreateObject("ADODB.Recordset")
    rsdata.CursorType = 3
    rsdata.CursorLocation = 3
    rsdata.Open sqltgl, connstr
    if not rsdata.eof then
        tglconfig = rsdata("PRODUCTIONMILLDATE")
    end if
    rsdata.close 
    set rsdata = nothing
    
    ' GET DATE PRODUCTION  PREV'
    sqltgl = "SELECT proddate FROM productiontrn " & _
    " WHERE proddate < to_date('" & tglnow & "','mm-dd-yyyy') " & _
    " ORDER BY proddate DESC FETCH FIRST 1 ROWS ONLY"
    'response.write sqltgl
    Set rsdata = Server.CreateObject("ADODB.Recordset")
    rsdata.CursorType = 3
    rsdata.CursorLocation = 3
    rsdata.Open sqltgl, connstr
    if not rsdata.eof then
        tglprev = rsdata("proddate")
    end if
    rsdata.close 
    set rsdata = nothing
    
    ' GET NETTO TBS, CPO, PK PRODUCTION PREV'
    sqltbs = "select tbsterima, tbssisa as tbssisakemarin,cpokirim as cpokirimkemarin,cpocot as cpocotkemarin,cpocst as cpocstkemarin,cpoot as cpootkemarin,cpotankisi as cpotankisikemarin,cpoot+cpocst as totcpoprev, " & _
    " kernelkirim,kernelsilo,kernelkirim+kernelsilo as totpkprev,pktankisi, cangkangkirim,cangkangproduksi,cangkangkirim+cangkangproduksi as ckgprev " & _  
    " from productiontrn  " & _
    " where proddate=to_date('" & tglprev & "','mm-dd-yyyy')"
   '' response.write sqltbs
    Set rsdata = Server.CreateObject("ADODB.Recordset")
    rsdata.CursorType = 3
    rsdata.CursorLocation = 3
    rsdata.Open sqltbs, connstr
    if not rsdata.eof then
        tbsterimakemarin    = cdbl(to_nol(rsdata("tbsterima")))
        sisatbskemarin      = cdbl(to_nol(rsdata("tbssisakemarin")))
        cpokemarin          = cdbl(to_nol(rsdata("totcpoprev")))
        pkkemarin           = cdbl(to_nol(rsdata("kernelsilo")))
        ckgkemarin          = cdbl(to_nol(rsdata("ckgprev")))
        cpokirimkemarin     = cdbl(to_nol(rsdata("cpokirimkemarin")))
        cpocotkemarin       = cdbl(to_nol(rsdata("cpocotkemarin")))
        cpocstkemarin       = cdbl(to_nol(rsdata("cpocstkemarin")))
        cpootkemarin        = cdbl(to_nol(rsdata("cpootkemarin")))
        cpotkkemarin        = cdbl(to_nol(rsdata("cpotankisikemarin")))
        pktkkemarin         = cdbl(to_nol(rsdata("pktankisi")))
    end if
    rsdata.close 
    set rsdata = nothing

    hslcpokemarin   = (cdbl(cpocotkemarin)+cdbl(cpocstkemarin)+cdbl(cpootkemarin)+cdbl(cpotkkemarin))
    hslpkkemarin    = cdbl(pkkemarin)+cdbl(pktkkemarin)

    
    ' GET NETTO CPOTANK PRODUCTION PREV'
    sqltbs = "select sum(cpoisitank) as tankcpokemarin " & _
    "from productiontankcpo where proddate=to_date('" & tglprev & "','mm-dd-yyyy')"
    'response.write sqltbs
    Set rsdata = Server.CreateObject("ADODB.Recordset")
    rsdata.CursorType = 3
    rsdata.CursorLocation = 3
    rsdata.Open sqltbs, connstr
    if not rsdata.eof then
        cpotankkemarin    = cdbl(to_nol(rsdata("tankcpokemarin")))
    end if
    rsdata.close 
    set rsdata = nothing


    ' GET NETTO PKTANK PRODUCTION PREV'
    sqltbs = "select sum(bulkingsiloisi) as tankpkkemarin " & _
    "from productionbulkingsilo where proddate=to_date('" & tglprev & "','mm-dd-yyyy')"
    'response.write sqltbs
    Set rsdata = Server.CreateObject("ADODB.Recordset")
    rsdata.CursorType = 3
    rsdata.CursorLocation = 3
    rsdata.Open sqltbs, connstr
    if not rsdata.eof then
        pktankkemarin     = cdbl(to_nol(rsdata("tankpkkemarin")))
    end if
    rsdata.close 
    set rsdata = nothing

    ' GET NETTO TBS,MUTU CPO NOW'
    sqltbs = "select tbsterima, tbssisa,jamkerja,cpoffa,cpoair,cpokotoran,cpodobi,cangkangproduksi as ckgnow" & _ 
    " from productiontrn " & _
    " where proddate=to_date('" & tglnow & "','mm-dd-yyyy') "
    'response.write sqltbs
    Set rsdata = Server.CreateObject("ADODB.Recordset")
    rsdata.CursorType = 3
    rsdata.CursorLocation = 3
    rsdata.Open sqltbs, connstr
    if not rsdata.eof then
        tbsnow      = cdbl(to_nol(rsdata("tbsterima")))
        tbssisanow  = cdbl(to_nol(rsdata("tbssisa")))
        jamkerja    = cdbl(to_nol(rsdata("jamkerja")))
        cpoffa      = cdbl(to_nol(rsdata("cpoffa")))
        cpoair      = cdbl(to_nol(rsdata("cpoair")))
        cpokotoran  = cdbl(to_nol(rsdata("cpokotoran")))
        cpodobi     = cdbl(to_nol(rsdata("cpodobi")))
        ckghi       = cdbl(to_nol(rsdata("ckgnow")))
    end if
    rsdata.close 
    set rsdata = nothing

    ' GET NETTO TBS PRODUCTION SHI,SBI,STI'
    sqltbs = "select sum(t.tbshariini) as tbshariini, sum(t.tbsbulanini) as tbsbulanini,sum(t.tbstahunini) as tbstahunini, " & _
    " sum(t.tbsolahhariini) as tbsolahhariini, sum(t.tbsolahbulanini) as tbsolahbulanini,sum(t.tbsolahtahunini) as tbsolahtahunini, " & _
    " sum(t.jamolahshi) as jamolahshi, sum(t.jamolahsbi) as jamolahsbi,sum(t.jamolahsti) as jamolahsti " & _
    " from ( " & _
    "   select sum(tbsterima)as tbshariini,0 as tbsbulanini,0 as tbstahunini, " & _
    "   sum(tbsolah)as tbsolahhariini,0 as tbsolahbulanini,0 as tbsolahtahunini, " & _
    "   sum(jamkerja)as jamolahshi,0 as jamolahsbi,0 as jamolahsti " & _
    "   from productiontrn " & _
    "   where proddate between TO_DATE ('" & tglblnawal & "', 'mm/dd/yyyy') " & _
    "   AND TO_DATE ('" & tglnow & "', 'mm/dd/yyyy') and fcba='" & fcba & "' " & _
    "        union all " & _
    "   select 0 as tbshariini,sum(tbsterima) as tbsbulanini,0 as tbstahunini, " & _
    "   0 as tbsolahhariini,sum(tbsolah) as tbsolahbulanini,0 as tbsolahtahunini, " & _
    "   0 as jamolahshi,sum(jamkerja) as jamolahsbi,0 as jamolahsti " & _
    "   from productiontrn " & _
    "   where proddate between TO_DATE ('" & tglthnawal & "', 'mm/dd/yyyy') " & _
    "   AND TO_DATE ('" & tglnow & "', 'mm/dd/yyyy') and fcba='" & fcba & "' " & _
    "        union all " & _
    "   select 0 as tbshariini,0 as tbsbulanini,sum(tbsterima) as tbstahunini, " & _
    "   0 as tbsolahhariini,0 as tbsolahbulanini,sum(tbsolah) as tbsolahtahunini, " & _
    "   0 as jamolahshi,0 as jamolahsbi,sum(jamkerja) as jamolahsti " & _
    "   from productiontrn where proddate <=to_Date('" & tglnow & "','mm/dd/yyyy') and fcba='" & fcba & "' " & _
    " ) t"
   '' response.write sqltbs
    Set rsdata = Server.CreateObject("ADODB.Recordset")
    rsdata.CursorType = 3
    rsdata.CursorLocation = 3
    rsdata.Open sqltbs, connstr
    if not rsdata.eof then
        tbsdhariini         = cdbl(to_nol(rsdata("tbshariini")))
        tbssdblnini         = cdbl(to_nol(rsdata("tbsbulanini")))
        tbssdthnini         = cdbl(to_nol(rsdata("tbstahunini")))
        jamolahshi          = cdbl(to_nol(rsdata("jamolahshi")))
        jamolahsbi          = cdbl(to_nol(rsdata("jamolahsbi")))
        jamolahsti          = cdbl(to_nol(rsdata("jamolahsti")))
    end if
    rsdata.close 
    set rsdata = nothing


    ' GET HASIL CPO,KERNEL,CKG HI
    sqltbs = "select tbssisa as tbssisahi,potongantbs, cpokirim,cpocot,cpocst,cpoot,cpotankisi, " & _
    " kernelkirim,kernelsilo,pktankisi,cangkangkirim,cangkangproduksi,jammulai,jamselesai,keterangan,air,caco3,adjcpo " & _  
    " from productiontrn  " & _
    " where proddate=to_date('" & tglnow & "','mm-dd-yyyy')"
    'response.write sqltbs
    Set rsdata = Server.CreateObject("ADODB.Recordset")
    rsdata.CursorType = 3
    rsdata.CursorLocation = 3
    rsdata.Open sqltbs, connstr
    if not rsdata.eof then
        tbssisahi           = cdbl(to_nol(rsdata("tbssisahi")))
        cpokirimhi          = cdbl(to_nol(rsdata("cpokirim")))
        cpocothi            = cdbl(to_nol(rsdata("cpocot")))
        cpocsthi            = cdbl(to_nol(rsdata("cpocst")))
        cpoothi             = cdbl(to_nol(rsdata("cpoot")))
        cpotankisihi        = cdbl(to_nol(rsdata("cpotankisi")))
        kernelkrmhi         = cdbl(to_nol(rsdata("kernelkirim")))
        kernelsilohi        = cdbl(to_nol(rsdata("kernelsilo")))
        pktankisihi         = cdbl(to_nol(rsdata("pktankisi")))
        cangkangkrmhi       = cdbl(to_nol(rsdata("cangkangkirim")))
        cangkangprodhi      = cdbl(to_nol(rsdata("cangkangproduksi")))
        jammulaitxt         = cdbl(to_nol(rsdata("jammulai")))
        jamselesaitxt       = cdbl(to_nol(rsdata("jamselesai")))
        keterangantxt       = rsdata("keterangan")
        airtxt              = cdbl(to_nol(rsdata("air")))
        caco3txt            = cdbl(to_nol(rsdata("caco3")))
        cpoadjhi            = cdbl(to_nol(rsdata("adjcpo")))
        pottbshi            = cdbl(to_nol(rsdata("potongantbs")))
    end if
    rsdata.close 
    set rsdata = nothing

    hslcpohi    = cpokirimhi + (cpocothi+cpocsthi+cpoothi+cpotankisihi+cpoadjhi)
    pkhi        = cdbl(kernelkrmhi) + (kernelsilohi+pktankisihi)


    tawal    = xbln2 & "/" & "01" & "/" & xthn2
    
     ' GET TGL AKHIR BULAN
    sqltbs = "SELECT LAST_DAY(TO_DATE('" & tawal & "', 'mm/dd/yyyy')) lasthari from dual"
    Set rsdata = Server.CreateObject("ADODB.Recordset")
    rsdata.CursorType = 3
    rsdata.CursorLocation = 3
    rsdata.Open sqltbs, connstr
    if not rsdata.eof then
       takhir = rsdata("lasthari")
    end if
    rsdata.close 
    set rsdata = nothing


     ' AMBIL VALUE TBS, CPO, PK, CKG, AIR, CACO3 SHI
    sqltbs = "select sum(tbsterima) as totaltbsterimashi,sum(potongantbs)as pottbsshi, sum(cpokirim) as cpokirimshi,sum(cpocot) as cpocotshi," & _
    " sum(cpocst) as cpocstshi,sum(cpoot)as cpootshi,sum(cpotankisi)as cpotankisishi, " & _
    " sum(kernelkirim) as kernelkrmshi,sum(kernelsilo) as kernelsiloshi,sum(pktankisi) as pktankisishi," & _
    " sum(kernelsilo) as kernelsiloshi,sum(air) as air,sum(caco3) as caco3," & _ 
    " sum(cangkangkirim) as ckgkirimshi,sum(cangkangproduksi) as ckgproduksishi,sum(adjcpo)as cpoadjshi " & _ 
    " from productiontrn " & _
    " where proddate between TO_DATE ('" & tglblnawal & "', 'mm/dd/yyyy') " & _ 
    " AND TO_DATE ('" & tglnow & "', 'mm/dd/yyyy') and fcba='" & fcba & "'"
    'response.write sqltbs
    Set rsdata = Server.CreateObject("ADODB.Recordset")
    rsdata.CursorType = 3
    rsdata.CursorLocation = 3
    rsdata.Open sqltbs, connstr
    if not rsdata.eof then
        gettotaltbsshi  = cdbl(to_nol(rsdata("totaltbsterimashi")))
        getcpokirimshi  = cdbl(to_nol(rsdata("cpokirimshi")))
        getcpocotshi    = cdbl(to_nol(rsdata("cpocotshi")))
        getcpocstshi    = cdbl(to_nol(rsdata("cpocstshi")))
        getcpootshi     = cdbl(to_nol(rsdata("cpootshi")))
        gettankisishi   = cdbl(to_nol(rsdata("cpotankisishi")))
        getpkkirimshi   = cdbl(to_nol(rsdata("kernelkrmshi")))
        getpksiloshi    = cdbl(to_nol(rsdata("kernelsiloshi")))
        getpktankshi    = cdbl(to_nol(rsdata("pktankisishi")))
        ttlairtxt       = cdbl(to_nol(rsdata("air")))
        ttlcaco3txt     = cdbl(to_nol(rsdata("caco3")))
        getkirimckgshi  = cdbl(to_nol(rsdata("ckgkirimshi")))
        getproduksickgshi  = cdbl(to_nol(rsdata("ckgproduksishi")))
        cpoadjshi       = cdbl(to_nol(rsdata("cpoadjshi")))
        pottbsshi       = cdbl(to_nol(rsdata("pottbsshi")))
        
     

    end if
    rsdata.close 
    set rsdata = nothing
   
      ' GET TGL PRODUCTION MAX
    sqltbs = "select MAX(proddate) as result " & _
    " from productiontrn " & _
    " where proddate between TO_DATE ('" & tawal & "', 'mm/dd/yyyy') " & _
    " AND TO_DATE ('" & takhir & "', 'mm/dd/yyyy')"
    'response.write sqltbs
    Set rsdata = Server.CreateObject("ADODB.Recordset")
    rsdata.CursorType = 3
    rsdata.CursorLocation = 3
    rsdata.Open sqltbs, connstr
    if not rsdata.eof then
        tglresult = rsdata("result")
    end if
    rsdata.close 
    set rsdata = nothing


   
    ' GET TBS,CPO,PK,CKG SAMPAI DENGAN TGL AKHIR BULAN KEMARIN
    sqltbs = "select tbsterima,tbssisa, (cpokirim) as cpokirimld,(cpocot) as cpocotld,(cpocst) as cpocstld,(cpoot)as cpootld,(cpotankisi)as cpotankisild, " & _
    " kernelsilo as kernelsilold,pktankisi as pktankisild,cangkangkirim as ckgkirimld,cangkangproduksi as ckgproduksild from productiontrn " & _
    " where proddate = TO_DATE ('" & tglresult & "', 'mm/dd/yyyy')  " & _
    " and fcba='" & fcba & "'"
  ''  response.write sqltbs
    Set rsdata = Server.CreateObject("ADODB.Recordset")
    rsdata.CursorType = 3
    rsdata.CursorLocation = 3
    rsdata.Open sqltbs, connstr
    if not rsdata.eof then
        tbsbf          = cdbl(to_nol(rsdata("tbssisa")))
        getcpokirimld  = cdbl(to_nol(rsdata("cpokirimld")))
        getcpocotld    = cdbl(to_nol(rsdata("cpocotld")))
        getcpocstld    = cdbl(to_nol(rsdata("cpocstld")))
        getcpootld     = cdbl(to_nol(rsdata("cpootld")))
        gettankisild   = cdbl(to_nol(rsdata("cpotankisild")))
        getpksilold    = cdbl(to_nol(rsdata("kernelsilold")))
        gettankpkisild = cdbl(to_nol(rsdata("pktankisild")))
        ckgkirimld      = cdbl(to_nol(rsdata("ckgkirimld")))
        ckgproduksild = cdbl(to_nol(rsdata("ckgproduksild")))
    end if
    rsdata.close 
    set rsdata = nothing

     ' hasil tbs, cpo, pk,ckg sampai hari ini (SHI)
    qtystockcposhi = getcpokirimshi + (cpocothi+cpocsthi+cpoothi+cpotankisihi) - (gettankisild+getcpocotld+getcpocstld+getcpootld)
    tbsterimashi   = cdbl(tbsbf) + (cdbl(gettotaltbsshi)-cdbl(tbssisahi))
   
    'Response.Write tbsbf & " / " & gettotaltbsshi & " / " & tbssisahi
    qtypkshi       = getpkkirimshi + (cdbl(pktankisihi)+cdbl(kernelsilohi))- (cdbl(gettankpkisild) + cdbl(getpksilold))

    qtyckgshi       = (getkirimckgshi+getproduksickgshi)-(ckgkirimld+ckgproduksild)
  
     ' GET LAST YEAR
    sqltbs = "select ADD_MONTHS(to_date('" & tglnow & "','mm/dd/yyyy'), 12 * 1 * -1) AS lastyear from dual"
    Set rsdata = Server.CreateObject("ADODB.Recordset")
    rsdata.CursorType = 3
    rsdata.CursorLocation = 3
    rsdata.Open sqltbs, connstr
    if not rsdata.eof then
       lastyear = rsdata("lastyear")
    end if
    rsdata.close 
    set rsdata = nothing

    lastmontsplit       = left(lastyear,2)
    lastdaysplit        = mid(lastyear,4,2)
    lastyearsplit       = right(lastyear,4)
    awalthn_prev        = "01" & "/" & "01" & "/" & lastyearsplit
    akhir_tahunprev     = "12" & "/" & "31" & "/" & lastyearsplit

     ' GET TBS,CPO,PK SBI
    sqltbs = "select sum(tbsterima) as totaltbsterimasbi,sum(tbssisa) as totaltbssisasbi,sum(potongantbs) as pottbssbi, sum(cpokirim) as cpokirimsbi, " & _ 
    " sum(cpocot) as cpocotsbi,sum(cpocst) as cpocstsbi,sum(cpoot)as cpootsbi,sum(cpotankisi)as cpotankisisbi, " & _
    " sum(kernelkirim) as kernelkrmsbi,sum(kernelsilo) as kernelsilosbi,sum(pktankisi) as pktankisisbi, " & _ 
    " sum(adjcpo) as cpoadjsbi from productiontrn " & _
    " where proddate between TO_DATE ('" & tglthnawal & "', 'mm/dd/yyyy') " & _ 
    " AND TO_DATE ('" & tglnow & "', 'mm/dd/yyyy') and fcba='" & fcba & "'"
  ''  response.write sqltbs
    Set rsdata = Server.CreateObject("ADODB.Recordset")
    rsdata.CursorType = 3
    rsdata.CursorLocation = 3
    rsdata.Open sqltbs, connstr
    if not rsdata.eof then
        gettbsterimasbi   = cdbl(to_nol(rsdata("totaltbsterimasbi")))
        gettbsolahsbi   = cdbl(to_nol(rsdata("totaltbssisasbi")))
        getcpokirimsbi  = cdbl(to_nol(rsdata("cpokirimsbi")))
        getcpocotsbi    = cdbl(to_nol(rsdata("cpocotsbi")))
        getcpocstsbi    = cdbl(to_nol(rsdata("cpocstsbi")))
        getcpootsbi     = cdbl(to_nol(rsdata("cpootsbi")))
        gettankisisbi   = cdbl(to_nol(rsdata("cpotankisisbi")))
        getpkkirimsbi   = cdbl(to_nol(rsdata("kernelkrmsbi")))
        getisipksbi     = cdbl(to_nol(rsdata("pktankisisbi")))
        getsilopksbi    = cdbl(to_nol(rsdata("kernelsilosbi")))
        cpoadjsbi       = cdbl(to_nol(rsdata("cpoadjsbi")))
        pottbssbi       = cdbl(to_nol(rsdata("pottbssbi")))

        hslcpobulanini  = getcpokirimsbi + (cpocothi+cpocsthi+cpoothi+cpotankisihi)
        hslpkbulanini   = getpkkirimsbi + (cdbl(pktankisihi)+cdbl(kernelsilohi))

      
    end if
    rsdata.close 
    set rsdata = nothing

  
     ' GET TGL PRODUCTION SBI
    sqltbs = "select MAX(proddate) as prev_year_max_date " & _
    " from productiontrn " & _
    " where proddate between TO_DATE ('" & awalthn_prev & "', 'mm/dd/yyyy') " & _
    " AND TO_DATE ('" & akhir_tahunprev & "', 'mm/dd/yyyy')"
   
    Set rsdata = Server.CreateObject("ADODB.Recordset")
    rsdata.CursorType = 3
    rsdata.CursorLocation = 3
    rsdata.Open sqltbs, connstr
    if not rsdata.eof then
        prev_year_max_date = rsdata("prev_year_max_date")
        'response.write qtystockcposbi
    end if
    rsdata.close 
    set rsdata = nothing


    ' GET TBS,CPO,PK SAMPAI DENGAN TAHUN KEMARIN
    sqltbs = "select tbssisa, (cpokirim) as cpokirimyd,(cpocot) as cpocotyd,(cpocst) as cpocstyd,(cpoot)as cpootyd,(cpotankisi)as cpotankisiyd, " & _
    " kernelkirim as kernelkrmyd,kernelsilo as kernelsiloyd, pktankisi as pktankisiyd from productiontrn " & _
    " where proddate = TO_DATE ('" & prev_year_max_date & "', 'mm/dd/yyyy')  " & _
    " and fcba='" & fcba & "'"
   '' response.write sqltbs
    Set rsdata = Server.CreateObject("ADODB.Recordset")
    rsdata.CursorType = 3
    rsdata.CursorLocation = 3
    rsdata.Open sqltbs, connstr
    if not rsdata.eof then
        gettbsolahyd        = cdbl(to_nol(rsdata("tbssisa")))
        getcpocotyd         = cdbl(to_nol(rsdata("cpocotyd")))
        getcpocstyd         = cdbl(to_nol(rsdata("cpocstyd")))
        getcpootyd          = cdbl(to_nol(rsdata("cpootyd")))
        gettankisiyd        = cdbl(to_nol(rsdata("cpotankisiyd")))
        pkkirimyd           = cdbl(to_nol(rsdata("kernelkrmyd")))
        pktankisiyd         = cdbl(to_nol(rsdata("pktankisiyd")))
        pksiloyd            = cdbl(to_nol(rsdata("kernelsiloyd")))

    end if
    rsdata.close 
    set rsdata = nothing
      hslcpoblnkemarin     = cdbl(getcpocotyd)+cdbl(getcpocstyd)+cdbl(getcpootyd)+cdbl(gettankisiyd)
      qtystockcposbi       = hslcpobulanini-hslcpoblnkemarin
      tbsolahsbi           = cdbl(gettbsolahyd) + (cdbl(gettbsterimasbi)-cdbl(tbssisahi))
      hslpkblnkemarin      = cdbl(pktankisiyd)+cdbl(pksiloyd)
      qtypksbi             = hslpkbulanini-hslpkblnkemarin
     
    

    ' GET TBS,CPO STI
    sqltbs = "select sum(tbsterima) as tbsterimasti,sum(potongantbs)as pottbssti, sum(cpokirim) as cpokirimsti,sum(cpocot) as cpocotsti,sum(cpocst) as cpocststi,sum(cpoot)as cpootsti,sum(cpotankisi)as cpotankisisti, " & _
    " sum(kernelkirim) as kernelkrmsti,sum(kernelsilo) as kernelsilosti,sum(pktankisi) as pktankisisti, " & _ 
    " sum(adjcpo) as cpoadjsti from productiontrn " & _
    " where proddate <=to_date('" & tglnow & "','mm-dd-yyyy') AND fcba='" & fcba & "'"
   '' response.write sqltbs
    Set rsdata = Server.CreateObject("ADODB.Recordset")
    rsdata.CursorType = 3
    rsdata.CursorLocation = 3
    rsdata.Open sqltbs, connstr
    if not rsdata.eof then
        getcpokirimsti       = cdbl(to_nol(rsdata("cpokirimsti")))
        getcpocotsti         = cdbl(to_nol(rsdata("cpocotsti")))
        getcpocststi         = cdbl(to_nol(rsdata("cpocststi")))
        getcpootsti          = cdbl(to_nol(rsdata("cpootsti")))
        gettankisisti        = cdbl(to_nol(rsdata("cpotankisisti")))
        gettbsterimasti      = cdbl(to_nol(rsdata("tbsterimasti")))
        getpkkirimsti        = cdbl(to_nol(rsdata("kernelkrmsti")))
        getpksilosti         = cdbl(to_nol(rsdata("kernelsilosti")))
        getpktanksti         = cdbl(to_nol(rsdata("pktankisisti")))
        cpoadjsti            = cdbl(to_nol(rsdata("cpoadjsti")))
        pottbssti            = cdbl(to_nol(rsdata("pottbssti")))
 
    end if
    rsdata.close 
    set rsdata = nothing
    qtystockcposti      = cdbl(getcpokirimsti)+(cpocothi+cpocsthi+cpoothi+cpotankisihi)
    tbsolahsti          = cdbl(gettbsterimasti)-cdbl(tbssisahi)
    qtypksti            = cdbl(getpkkirimsti)+(cdbl(pktankisihi)+cdbl(kernelsilohi))


    ' GET MUTU SHI'
    sqltbs = "select sum(t.cpoffa) as cpoffashi,sum(t.cpoair) as cpoairshi," & _
    " sum(t.cpokotoran) as cpokotoranshi,sum(t.cpodobi) as cpodobishi, " & _
    " sum(t.pkffashi) as pkffashi,sum(t.pkairshi) as pkairshi,sum(t.pkkotoranshi) as pkkotoranshi " & _
    " from( " & _
    "       select CPOFFA,CPOAIR,CPOKOTORAN,CPODOBI,0 as pkffashi,0 as pkairshi,0 as pkkotoranshi " & _
    "       FROM PRODUCTIONTRN  " & _
    "       WHERE proddate between TO_DATE ('" & tglblnawal & "', 'mm/dd/yyyy') " & _
    "       AND TO_DATE ('" & tglnow & "', 'mm/dd/yyyy') and fcba='" & fcba & "' " & _
    "           union all " & _
    "       select 0 as CPOFFA,0 as CPOAIR,0 as CPOKOTORAN,0 as CPODOBI,bulkingsiloffa as pkffashi, " & _
    "       bulkingsiloair as pkairshi,bulkingsilokotoran as pkkotoranshi " & _
    "       FROM PRODUCTIONBULKINGSILO  " & _
    "       WHERE proddate between TO_DATE ('" & tglblnawal & "', 'mm/dd/yyyy') " & _
    "       AND TO_DATE ('" & tglnow & "', 'mm/dd/yyyy') and fcba='" & fcba & "' " & _
    " )t"
   '' response.write sqltbs
    Set rsdata = Server.CreateObject("ADODB.Recordset")
    rsdata.CursorType = 3
    rsdata.CursorLocation = 3
    rsdata.Open sqltbs, connstr
    if not rsdata.eof then
        cpoffashi      = cdbl(to_nol(rsdata("cpoffashi")))
        cpoairshi      = cdbl(to_nol(rsdata("cpoairshi")))
        cpokotoranshi  = cdbl(to_nol(rsdata("cpokotoranshi")))
        cpodobishi     = cdbl(to_nol(rsdata("cpodobishi")))
        pkffashi       = cdbl(to_nol(rsdata("pkffashi")))
        pkairshi       = cdbl(to_nol(rsdata("pkairshi")))
        pkkotoranshi   = cdbl(to_nol(rsdata("pkkotoranshi")))
    end if
    rsdata.close 
    set rsdata = nothing


    ' GET MUTU PK NOW'
    sqltbs = "select BULKINGSILOFFA,BULKINGSILOAIR,BULKINGSILOKOTORAN " & _
    " from PRODUCTIONBULKINGSILO " & _
    " where proddate=to_date('" & tglnow & "','mm-dd-yyyy') AND fcba='" & fcba & "'"
    'response.write sqltbs
    Set rsdata = Server.CreateObject("ADODB.Recordset")
    rsdata.CursorType = 3
    rsdata.CursorLocation = 3
    rsdata.Open sqltbs, connstr
    if not rsdata.eof then
        pkffahi       = cdbl(to_nol(rsdata("BULKINGSILOFFA")))
        pkairhi       = cdbl(to_nol(rsdata("BULKINGSILOAIR")))
        pkkotoranhi   = cdbl(to_nol(rsdata("BULKINGSILOKOTORAN")))
    end if
    rsdata.close 
    set rsdata = nothing


    ' GET CPO DO'
    blnsebelum  = DateAdd("m", -1, tglnow)
    tglsebelum2 = DateAdd("d", -7, tglnow)
    'response.write tglsebelum2 & "--" & blnsebelum
    sqlcpodo = "select dm.DOOrderID,(SELECT SUM(dc.ContractQty) FROM LHPDOMSTDTL  dc WHERE dc.DOID=d.DOID GROUP BY d.DOID) AS Party, " & _
    " SUM ( case when h.docketdate = to_date('" & tglnow & "', 'mm/dd/yyyy') then d.NetWeight else 0 end )as HariIni," & _
    " ( " & _
    "           SELECT SUM(b.NetWeight)  " & _
    "           FROM LHPTIKETCPO b, LHPTIKETHDR a " & _
    "           WHERE a.DocketID=b.DocketID AND a.DocketDate<=to_date('" & tglnow & "', 'mm/dd/yyyy') AND b.DOID=d.DOID " & _		
    " ) AS SdHariIni " & _
    " from LHPTIKETHDR  h  " & _
    " INNER JOIN LHPTIKETCPO d ON h.DocketID=d.DocketID " & _
    " INNER JOIN LHPDOMST  dm ON d.DOID=dm.DOID " & _
    " INNER JOIN LHPCONTRACTMST  c ON dm.ContractID=c.ContractID " & _
    " WHERE h.DocketDate>=to_date('" & blnsebelum & "', 'mm/dd/yyyy') " & _
    " AND h.DocketDate<=to_date('" & tglnow & "', 'mm/dd/yyyy') AND c.ProductID='1' " & _
    " AND dm.DOToDate>=to_date('" & tglsebelum2 & "', 'mm/dd/yyyy') " & _
    " GROUP BY dm.DOOrderID, d.DOID  " & _
    " HAVING (SELECT SUM(dc.ContractQty) " & _
    " FROM LHPDOMSTDTL dc WHERE dc.DOID=d.DOID GROUP BY d.DOID) - ( SELECT SUM(b.NetWeight) FROM LHPTIKETCPO b, LHPTIKETHDR a " & _
    " WHERE a.DocketID=b.DocketID AND a.DocketDate<=to_date('" & tglnow & "', 'mm/dd/yyyy') " & _
    " AND b.DOID=d.DOID ) || SUM( case when h.docketdate = to_date('" & tglnow & "', 'mm/dd/yyyy') then d.NetWeight else 0 end )>100" 
   ' response.write sqlcpodo
    Set rscpodo = Server.CreateObject("ADODB.Recordset")
    rscpodo.CursorType = 3
    rscpodo.CursorLocation = 3
    rscpodo.Open sqlcpodo, connstr


    ' GET KERNEL DO'
    Dim sqlkerneldo,rskerneldo
    sqlkerneldo = "select dm.DOOrderID,(SELECT SUM(dc.ContractQty) FROM LHPDOMSTDTL  dc WHERE dc.DOID=d.DOID GROUP BY d.DOID) AS Party, " & _
    " SUM ( case when h.docketdate = to_date('" & tglnow & "', 'mm/dd/yyyy') then d.NetWeight else 0 end )as HariIni," & _
    " ( " & _
    "           SELECT SUM(b.NetWeight)  " & _
    "           FROM LHPTIKETKERNEL b, LHPTIKETHDR a " & _
    "           WHERE a.DocketID=b.DocketID AND a.DocketDate<=to_date('" & tglnow & "', 'mm/dd/yyyy') AND b.DOID=d.DOID " & _		
    " ) AS SdHariIni " & _
    " from LHPTIKETHDR  h  " & _
    " INNER JOIN LHPTIKETKERNEL d ON h.DocketID=d.DocketID " & _
    " INNER JOIN LHPDOMST  dm ON d.DOID=dm.DOID " & _
    " INNER JOIN LHPCONTRACTMST  c ON dm.ContractID=c.ContractID " & _
    " WHERE h.DocketDate>=to_date('" & blnsebelum & "', 'mm/dd/yyyy') " & _
    " AND h.DocketDate<=to_date('" & tglnow & "', 'mm/dd/yyyy') AND c.ProductID='3' " & _
    " AND dm.DOToDate>=to_date('" & tglsebelum2 & "', 'mm/dd/yyyy') " & _
    " GROUP BY dm.DOOrderID, d.DOID  " & _
    " HAVING (SELECT SUM(dc.ContractQty) " & _
    " FROM LHPDOMSTDTL dc WHERE dc.DOID=d.DOID GROUP BY d.DOID) - ( SELECT SUM(b.NetWeight) FROM LHPTIKETKERNEL b, LHPTIKETHDR a " & _
    " WHERE a.DocketID=b.DocketID AND a.DocketDate<=to_date('" & tglnow & "', 'mm/dd/yyyy') " & _
    " AND b.DOID=d.DOID ) || SUM( case when h.docketdate = to_date('" & tglnow & "', 'mm/dd/yyyy') then d.NetWeight else 0 end )>100" 
    'response.write sqlkerneldo
    Set rskerneldo = Server.CreateObject("ADODB.Recordset")
    rskerneldo.CursorType = 3
    rskerneldo.CursorLocation = 3
    rskerneldo.Open sqlkerneldo, connstr

    ' GET CKG DO'
    Dim sqlckgdo,rsckgdo
    sqlckgdo = "select dm.DOOrderID,(SELECT SUM(dc.ContractQty) FROM LHPDOMSTDTL  dc WHERE dc.DOID=d.DOID GROUP BY d.DOID) AS Party, " & _
    " SUM ( case when h.docketdate = to_date('" & tglnow & "', 'mm/dd/yyyy') then d.NetWeight else 0 end )as HariIni," & _
    " ( " & _
    "           SELECT SUM(b.NetWeight)  " & _
    "           FROM LHPTIKETCKG b, LHPTIKETHDR a " & _
    "           WHERE a.DocketID=b.DocketID AND a.DocketDate<=to_date('" & tglnow & "', 'mm/dd/yyyy') AND b.DOID=d.DOID " & _		
    " ) AS SdHariIni " & _
    " from LHPTIKETHDR  h  " & _
    " INNER JOIN LHPTIKETCKG d ON h.DocketID=d.DocketID " & _
    " INNER JOIN LHPDOMST  dm ON d.DOID=dm.DOID " & _
    " INNER JOIN LHPCONTRACTMST  c ON dm.ContractID=c.ContractID " & _
    " WHERE h.DocketDate>=to_date('" & blnsebelum & "', 'mm/dd/yyyy') " & _
    " AND h.DocketDate<=to_date('" & tglnow & "', 'mm/dd/yyyy') AND c.ProductID='5' " & _
    " AND dm.DOToDate>=to_date('" & tglsebelum2 & "', 'mm/dd/yyyy') " & _
    " GROUP BY dm.DOOrderID, d.DOID  " & _
    " HAVING (SELECT SUM(dc.ContractQty) " & _
    " FROM LHPDOMSTDTL dc WHERE dc.DOID=d.DOID GROUP BY d.DOID) - ( SELECT SUM(b.NetWeight) FROM LHPTIKETCKG b, LHPTIKETHDR a " & _
    " WHERE a.DocketID=b.DocketID AND a.DocketDate<=to_date('" & tglnow & "', 'mm/dd/yyyy') " & _
    " AND b.DOID=d.DOID ) || SUM( case when h.docketdate = to_date('" & tglnow & "', 'mm/dd/yyyy') then d.NetWeight else 0 end )>100" 
    'response.write sqlckgdo
    Set rsckgdo = Server.CreateObject("ADODB.Recordset")
    rsckgdo.CursorType = 3
    rsckgdo.CursorLocation = 3
    rsckgdo.Open sqlckgdo, connstr
   
   if jamolahshi = 0 then 
      jamolahshi = 0
   end if

    tbstersedia     = cdbl(tbsnow)+cdbl(sisatbskemarin)
    tbsolah         = (cdbl(tbstersedia)-cdbl(tbssisanow))-cdbl(pottbshi)
    if jamkerja = 0 then
        kapasitasjam    = 0
    else
        kapasitasjam    = cdbl(tbsolah)/cdbl(jamkerja)
    end if
    if tbsterimashi = 0 then
        kapasitasshi    = 0
    else
        kapasitasshi    = cdbl(tbsterimashi)/cdbl(jamolahshi)
    end if
    if tbsolahsbi = 0 then
        kapasitassbi    = 0
    else
        kapasitassbi    = cdbl(tbsolahsbi)/cdbl(jamolahsbi)
    end if
    if tbsolahsti = 0 then
        kapasitassti    = 0
    else
        kapasitassti    = cdbl(tbsolahsti)/cdbl(jamolahsti)
    end if
  

    totalcpokemarin = cdbl(cpokemarin) + cdbl(cpotankkemarin)
    cpohriini       = hslcpohi-hslcpokemarin
    totalpkkemarin  = cdbl(pkkemarin) + cdbl(pktankkemarin)
    pkhi            = cdbl(pkhi)-cdbl(hslpkkemarin)
    ckgviewhi       = cdbl(ckghi)
    

  

    if tbsolah = 0 then
        oerhi   = 0 
    else
        oerhi   = (cdbl(cpohriini)/cdbl(tbsolah)) *100
    end if
    if tbsterimashi = 0 then
        tbsterimashi   = 0 
    else
        oershi   = (cdbl(qtystockcposhi) / cdbl(tbsterimashi)) *100
    end if
    if tbsolahsbi = 0 then
        tbsolahsbi   = 0 
    else
        oersbi   = (cdbl(qtystockcposbi) / cdbl(tbsolahsbi)) *100
    end if
    if tbsolahsti = 0 then
        tbsolahsti   = 0 
    else
        oersti   = (cdbl(qtystockcposti) / cdbl(tbsolahsti)) *100
    end if

   
    if tbsolah = 0 then
        kerhi   = 0 
    else
        kerhi   = (cdbl(pkhi)/cdbl(tbsolah))*100
    end if
    if tbsterimashi = 0 then
        kershi   = 0 
    else
        kershi   = cdbl(qtypkshi) / cdbl(tbsterimashi)*100
        
    end if
    if tbsolahsbi = 0 then
        kersbi   = 0 
    else
        kersbi   = (cdbl(qtypksbi) / cdbl(tbsolahsbi)) *100
    end if
    if tbsolahsti = 0 then
        kersti   = 0 
    else
        kersti   = (cdbl(qtypksti) / cdbl(tbsolahsti)) *100
    end if

    'Response.Write tbsterimashi
    stockckgtxt     = cdbl(tbsolah) * (2.5/100) 

    if tbsterimashi = 0 then 
        ratioairshi = 0
       
    else
       
        ratioairshi = cdbl(ttlairtxt) / cdbl(tbsterimashi)
        
        ratiocaco3shi   = cdbl(ttlcaco3txt) / cdbl(tbsterimashi)
        
    end if
    if tbsolah = 0 then 
        ratiocaco3hi= 0
    else 
        ratiocaco3hi    = cdbl(caco3txt) / cdbl(tbsolah)
    end if
   
    tbsterimashi = cdbl(tbsterimashi)-cdbl(pottbsshi)
    tbsolahsbi = cdbl(tbsolahsbi)-cdbl(pottbssbi)
    tbsolahsti = cdbl(tbsolahsti)-cdbl(pottbssti)
   
%>

	<body>
		<div id="wrapper">
         <table border="1" width="100%" id="product-table" class="">
         <thead>
         <tr>
            <td colspan="10" style=" background-color:#FFFFFF; color:#000000; border:none">
                <div id="header-wrapper">
                    <div id="head_judul" style="float:left; text-align:left;">
                        <h4>PT. <%=Session("COMPANY")%></h4>
                        <h5><%=Session("BA")%></h5>
                    </div>
                    <div id="head_judul" style="float:right; text-align:right; position: relative; left: 250px;">
                        <h4>Kepada Yth: Bpk Mill Manager</h4>
                    </div>
                </div>
            </td>
          
		 </tr>
         <tr>
            <td colspan="11">
                 <div id="subheader" class="centered" style="width:70%; margin:auto;">
                        <h1 class="judul">LAPORAN HARIAN PRODUKSI PABRIK</h1>
                    </div>
            </td>
            <td colspan="1">
                    <table >
                        <tr>
                            <td class="">Hari</td>
                          
                            <td class="lebarjamolah" >Kamis</td>
                        </tr>
                        <tr>
                            <td class="">Tanggal</td>
                            
                            <td class="lebarjamolah" ><%=tglnow%></td>
                        </tr>
                        <tr>
                            <td class="">Jam Start</td>
                            
                            <td class="lebarjamolah" ><%=jammulaitxt%></td>
                        </tr>
                        <tr>
                            <td class="">Jam Stop</td>
                            
                            <td class="lebarjamolah" ><%=jamselesaitxt%></td>
                        </tr>
                        <tr>
                            <td class="">Jumlah Jam Olah</td>
                            
                            <td class="lebarjamolah" ><%=jamkerja%></td>
                        </tr>
                    </table> 
             </td>   
         </tr>
        <tr>
        <div id="wrapper">
            <table border="1" width="100%" id="product-table" class="xthf-site">				
				<tr>
					 
					<th rowspan="2">No</th>
					<th rowspan="2" class="lebarkolomlurai">Uraian</th>
					<th rowspan="2">Satuan</th>
					<th rowspan="2" class="lebarkolomlhi">Hari Ini</th>
					<th rowspan="2" class="lebarkolomshi">S/D Hari ini</th>
					<th rowspan="2">S/D Bulan ini</th>
					<th rowspan="2">S/D Tahun ini</th>
					<th rowspan="2" colspan="5" class="">PENGIRIMAN - LOSESS - PEMAKAIAN AIR</th>
					<!--<th rowspan="2">Party DO</th>
					<th colspan="3">Jumlah Diangkut</th>-->
				</tr>
                
                </thead>
                <tbody>
                    <tr>
                        <td align="center">I</td>
                        <td colspan="6" class="tebal">TBS MASUK DAN DIOLAH</td>
                        <td colspan="5" class="tebal">PENGIRIMAN CPO</td>
                        
                    </tr>
                    <tr>
                        <td></td>
                        <td>1. TBS sisa kemarin</td>
                        <td>Kg</td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(sisatbskemarin)),0)%></td>
                        <td class="ratakanan">0</td>
                        <td class="ratakanan">0</td>
                        <td class="ratakanan">0</td>
                        <td colspan="5" rowspan="8">
                            <table class="tabeltengah" width="600" border="0" >
                                <tr>
                                    <th rowspan="2" class="">No. DO</th>
                                    <th rowspan="2">Party DO</th>
                                    <th colspan="3">Jumlah Diangkut</th>
                                </tr>
                                 <tr>
                                    <th colspan="">Hari ini</th>
                                    <th colspan="">S/D Hari ini</th>
                                    <th colspan="">Sisa DO</th>
                                </tr>
                                <%
                
                                    While Not rscpodo.EOF
                                         cpodoparty       = cdbl(to_nol(rscpodo("PARTY")))
                                         cpodohi          = cdbl(to_nol(rscpodo("HariIni")))
                                         cpodoshi         = cdbl(to_nol(rscpodo("SdHariIni")))
                                         cpodosisa        = cdbl(cpodoparty)-cdbl(cpodoshi)

                                                                
                                %>
                                <tr >
                                    <td class=""><%=rscpodo("DOOrderID")%></td>
                                    <td class="ratakanan"><%=formatnumber(cdbl(to_nol(cpodoparty)),0)%></td>
                                    <td class="ratakanan"><%=formatnumber(cdbl(to_nol(cpodohi)),0)%></td>
                                    <td class="ratakanan"><%=formatnumber(cdbl(to_nol(cpodoshi)),0)%></td>
                                    <td class="ratakanan"><%=formatnumber(cdbl(to_nol(cpodosisa)),0)%></td>
                
                                </tr>
                                <%
                                    rscpodo.MoveNext
                                    'response.Flush()

                                         totcpodoparty =  totcpodoparty +  cpodoparty 
                                         totcpodohi    =  totcpodohi + cpodohi
                                         totcpodoshi   =  totcpodoshi + cpodoshi
                                         totcpodosisa  =  totcpodosisa + cpodosisa


                                    no=no+1
                                    Wend
                                %>
                                <tr>
                                    <td class="tebal">Jumlah</td>
                                    <td class="ratakanan tebal"><%=formatnumber(cdbl(to_nol(totcpodoparty)),0)%></td>
                                    <td class="ratakanan tebal"><%=formatnumber(cdbl(to_nol(totcpodohi)),0)%></td>
                                    <td class="ratakanan tebal"><%=formatnumber(cdbl(to_nol(totcpodoshi)),0)%></td>
                                    <td class="ratakanan tebal"><%=formatnumber(cdbl(to_nol(totcpodosisa)),0)%></td>
                                </tr>
                                <%	rscpodo.close() 
                                    set rscpodo = nothing
                                %>
                            </table>
                        </td>
                    </tr>
                    <tr>
                        <td></td>
                        <td>2. TBS masuk hari ini</td>
                        <td>Kg</td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(tbsnow)),0)%></td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(tbsdhariini)),0)%></td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(tbssdblnini)),0)%></td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(tbssdthnini)),0)%></td>
                    </tr>
                    <tr>
                        <td></td>
                        <td>3. Jumlah TBS tersedia</td>
                        <td>Kg</td>
                         <td class="ratakanan"><%=formatnumber(cdbl(to_nol(tbstersedia)),0)%></td>
                        <td class="ratakanan">0</td>
                        <td class="ratakanan">0</td>
                        <td class="ratakanan">0</td>
                    </tr>
                    <tr>
                        <td></td>
                        <td>4. TBS Keluar</td>
                        <td>Kg</td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(pottbshi)),0)%></td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(pottbsshi)),0)%></td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(pottbssbi)),0)%></td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(pottbssti)),0)%></td>
                    </tr>
                    <tr>
                        <td></td>
                        <td>5. TBS diolah hari ini</td>
                        <td>Kg</td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(tbsolah)),0)%></td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(tbsterimashi)),0)%></td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(tbsolahsbi)),0)%></td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(tbsolahsti)),0)%></td>
                    </tr>
                    <tr>
                        <td></td>
                        <td>6. Sisa TBS hari ini</td>
                        <td>Kg</td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(tbssisanow)),0)%></td>
                        <td class="ratakanan">0</td>
                        <td class="ratakanan">0</td>
                        <td class="ratakanan">0</td>
                    </tr>
                    <tr>
                        <td></td>
                        <td>7. Jumlah jam olah</td>
                        <td>Jam</td>
                        <td class="ratakanan"><%=FbpNum2(jamkerja)%></td>
                        <td class="ratakanan"><%=FbpNum2(jamolahshi)%></td>
                        <td class="ratakanan"><%=FbpNum2(jamolahsbi)%></td>
                        <td class="ratakanan"><%=FbpNum2(jamolahsti)%></td>
                    </tr>
                    <tr>
                        <td></td>
                        <td>8. Kapasitas per - jam</td>
                        <td>Kg</td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(kapasitasjam)),0)%></td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(kapasitasshi)),0)%></td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(kapasitassbi)),0)%></td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(kapasitassti)),0)%></td>
                    </tr>
                    <tr>
                        <td align="center">II</td>
                        <td colspan="6" class="tebal">HASIL DAN MUTU CPO</td>
                        <td colspan="5" class="tebal">PENGIRIMAN INTI</td>
                    </tr>
                     <tr>
                        <td></td>
                        <td>1. Hasil CPO hari ini</td>
                        <td>Kg</td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(cpohriini)),0)%></td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(qtystockcposhi)),0)%></td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(qtystockcposbi)),0)%></td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(qtystockcposti)),0)%></td>
                        <td colspan="5" rowspan="6">
                            <table class="tabeltengah" width="600" border="0">
                                <tr>
                                    <th rowspan="2" class="">No. DO</th>
                                    <th rowspan="2">Party DO</th>
                                    <th colspan="3">Jumlah Diangkut</th>
                                </tr>
                                 <tr>
                                    <th colspan="">Hari ini</th>
                                    <th colspan="">S/D Hari ini</th>
                                    <th colspan="">Sisa DO</th>
                                </tr>
                                <%
                
                                    While Not rskerneldo.EOF
                                      
                                         pkdoparty       = cdbl(to_nol(rskerneldo("PARTY")))
                                         pkdohi          = cdbl(to_nol(rskerneldo("HariIni")))
                                         pkdoshi         = cdbl(to_nol(rskerneldo("SdHariIni")))
                                         pkdosisa        = cdbl(pkdoparty)-cdbl(pkdoshi)

                                                                
                                %>
                                <tr >
                                    <td class=""><%=rskerneldo("DOOrderID")%></td>
                                    <td class="ratakanan"><%=formatnumber(cdbl(to_nol(pkdoparty)),0)%></td>
                                    <td class="ratakanan"><%=formatnumber(cdbl(to_nol(pkdohi)),0)%></td>
                                    <td class="ratakanan"><%=formatnumber(cdbl(to_nol(pkdoshi)),0)%></td>
                                    <td class="ratakanan"><%=formatnumber(cdbl(to_nol(pkdosisa)),0)%></td>
                
                                </tr>
                                <%
                                    rskerneldo.MoveNext
                                    'response.Flush()

                                         totpkdoparty =  totpkdoparty +  pkdoparty 
                                         totpkdohi    =  totpkdohi + pkdohi
                                         totpkdoshi   =  totpkdoshi + pkdoshi
                                         totpkdosisa  =  totpkdosisa + pkdosisa


                                    no=no+1
                                    Wend
                                %>
                                <tr>
                                    <td class="tebal">Jumlah</td>
                                    <td class="ratakanan tebal"><%=formatnumber(cdbl(to_nol(totpkdoparty)),0)%></td>
                                    <td class="ratakanan tebal"><%=formatnumber(cdbl(to_nol(totpkdohi)),0)%></td>
                                    <td class="ratakanan tebal"><%=formatnumber(cdbl(to_nol(totpkdoshi)),0)%></td>
                                    <td class="ratakanan tebal"><%=formatnumber(cdbl(to_nol(totpkdosisa)),0)%></td>
                                </tr>
                                <%	rskerneldo.close() 
                                    set rskerneldo = nothing
                                %>
                            </table>
                        </td>
                    </tr>
                     <tr>
                        <td></td>
                        <td>2. OER</td>
                        <td>%</td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(oerhi)),2)%></td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(oershi)),2)%></td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(oersbi)),2)%></td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(oersti)),2)%></td>
            
                    </tr>
                     <tr>
                        <td></td>
                        <td>3. FFA</td>
                        <td>%</td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(cpoffa)),2)%></td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(cpoffashi)),2)%></td>
                        <td class="ratakanan">0</td>
                        <td class="ratakanan">0</td>
                       
                    </tr>
                     <tr>
                        <td></td>
                        <td>4. Kadar Air</td>
                        <td>%</td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(cpoair)),2)%></td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(cpoairshi)),2)%></td>
                        <td class="ratakanan">0</td>
                        <td class="ratakanan">0</td>
                       
                    </tr>
                    <tr>
                        <td></td>
                        <td>5. Kadar Kotoran</td>
                        <td>%</td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(cpokotoran)),2)%></td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(cpokotoranshi)),2)%></td>
                        <td class="ratakanan">0</td>
                        <td class="ratakanan">0</td>
                        
                    </tr>
				
                     <tr>
                        <td></td>
                        <td>6. Dobi</td>
                        <td>%</td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(cpodobi)),2)%></td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(cpodobishi)),2)%></td>
                        <td class="ratakanan">0</td>
                        <td class="ratakanan">0</td>
                        
                    </tr>
                    
                     <tr>
                        <td align="center">III</td>
                        <td colspan="6" class="tebal">HASIL DAN MUTU PK</td>
                        <td colspan="5" class="tebal">PENGIRIMAN CANGKANG</td>
                    </tr>
                    <tr>
                        <td></td>
                        <td>1. Hasil PK hari ini</td>
                        <td>Kg</td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(pkhi)),0)%></td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(qtypkshi)),0)%></td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(qtypksbi)),0)%></td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(qtypksti)),0)%></td>
                        <td colspan="5" rowspan="5" >
                            <table class="tabeltengah" width="600" border="0">
                                <tr>
                                    <th rowspan="2" class="">No. DO</th>
                                    <th rowspan="2">Party DO</th>
                                    <th colspan="3">Jumlah Diangkut</th>
                                </tr>
                                 <tr>
                                    <th colspan="">Hari ini</th>
                                    <th colspan="">S/D Hari ini</th>
                                    <th colspan="">Sisa DO</th>
                                </tr>
                                <%
                
                                    While Not rsckgdo.EOF
                                         ckgdoparty       = cdbl(to_nol(rsckgdo("PARTY")))
                                         ckgdohi          = cdbl(to_nol(rsckgdo("HariIni")))
                                         ckgdoshi         = cdbl(to_nol(rsckgdo("SdHariIni")))
                                         ckgdosisa        = cdbl(ckgdoparty)-cdbl(ckgdoshi)

                                                                
                                %>
                                <tr >
                                    <td class=""><%=rsckgdo("DOOrderID")%></td>
                                    <td class="ratakanan"><%=formatnumber(cdbl(to_nol(ckgdoparty)),0)%></td>
                                    <td class="ratakanan"><%=formatnumber(cdbl(to_nol(ckgdohi)),0)%></td>
                                    <td class="ratakanan"><%=formatnumber(cdbl(to_nol(ckgdoshi)),0)%></td>
                                    <td class="ratakanan"><%=formatnumber(cdbl(to_nol(ckgdosisa)),0)%></td>
                
                                </tr>
                                <%
                                    rsckgdo.MoveNext
                                    'response.Flush()

                                         totckgdoparty =  totckgdoparty +  ckgdoparty 
                                         totckgdohi    =  totckgdohi + ckgdohi
                                         totckgdoshi   =  totckgdoshi + ckgdoshi
                                         totckgdosisa  =  totckgdosisa + ckgdosisa


                                    no=no+1
                                    Wend
                                %>
                                <tr>
                                    <td class="tebal">Jumlah</td>
                                    <td class="ratakanan tebal"><%=formatnumber(cdbl(to_nol(totckgdoparty)),0)%></td>
                                    <td class="ratakanan tebal"><%=formatnumber(cdbl(to_nol(totckgdohi)),0)%></td>
                                    <td class="ratakanan tebal"><%=formatnumber(cdbl(to_nol(totckgdoshi)),0)%></td>
                                    <td class="ratakanan tebal"><%=formatnumber(cdbl(to_nol(totckgdosisa)),0)%></td>
                                </tr>
                                <%	rsckgdo.close() 
                                    set rsckgdo = nothing
                                %>
                            </table>
                        </td>
                       
                    </tr>
                    <tr>
                        <td></td>
                        <td>2. Kernel</td>
                        <td>%</td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(kerhi)),2)%></td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(kershi)),2)%></td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(kersbi)),2)%></td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(kersti)),2)%></td>
                       
                    </tr>
                    <tr>
                        <td></td>
                        <td>3. FFA</td>
                        <td>%</td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(pkffahi)),2)%></td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(pkffashi)),2)%></td>
                        <td class="ratakanan">0</td>
                        <td class="ratakanan">0</td>
                       
                    </tr>
                    <tr>
                        <td></td>
                        <td>4. Kadar Air</td>
                        <td>%</td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(pkairhi)),2)%></td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(pkairshi)),2)%></td>
                        <td class="ratakanan">0</td>
                        <td class="ratakanan">0</td>
                        
                    </tr>
                    <tr>
                        <td></td>
                        <td>5. Kadar Kotoran</td>
                        <td>%</td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(pkkotoranhi)),2)%></td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(pkkotoranshi)),2)%></td>
                        <td class="ratakanan">0</td>
                        <td class="ratakanan">0</td>
                       
                    </tr>

                    <tr>
                        <td align="center">IV</td>
                        <td colspan="6" class="tebal">HASIL CANGKANG</td>
                        <td colspan="6" class="tebal"></td>
                    </tr>
                   
                    <tr>
                         <td></td>
                        <td>1. Hasil Cangkang hari ini</td>
                        <td>Kg</td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(ckgviewhi)),0)%></td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(qtyckgshi)),0)%></td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(qtyckgsbi)),0)%></td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(qtyckgsti)),0)%></td>
                        <td rowspan="1" colspan=6>
                            <table class="tabeltengah" width="600" border="0">
                                <tr>
                                    <th rowspan="" colspan="2" class="">Pemakaian AIR</th>
                                    <th rowspan="" colspan="2">HI</th>
                                    <th colspan="" colspan="2">SHI</th>
                                   
                                </tr>
                                <tr>
                                    <td colspan="2">Hari Ini (M3)</td>
                                    <td colspan="2" class="ratakanan"><%=FbpNum2(airtxt)%></td>
                                    <td colspan="2" class="ratakanan"><%=FbpNum2(ttlairtxt)%></td>
                                </tr>
                                <tr>
                                    <td colspan="2">Ratio (M3/Kg TBS)</td>
                                    <td colspan="2" class="ratakanan">0</td>
                                    <td colspan="2" class="ratakanan"><%=formatnumber(cdbl(to_nol(ratioairshi)),0)%></td>
                                </tr>
                                <tr>
                                    <th rowspan="" colspan="2" class="">Pemakaian CaCo3</th>
                                    <th rowspan="" colspan="2">HI</th>
                                    <th colspan="" colspan="2">SHI</th>
                                   
                                </tr>
                                <tr>
                                    <td colspan="2">Hari Ini (M3)</td>
                                    <td colspan="2" class="ratakanan"><%=FbpNum2(caco3txt)%></td>
                                    <td colspan="2" class="ratakanan"><%=FbpNum2(ttlcaco3txt)%></td>
                                </tr>
                                <tr>
                                    <td colspan="2">Ratio (Kg/Kg TBS)</td>
                                    <td colspan="2" class="ratakanan"><%=formatnumber(cdbl(to_nol(ratiocaco3hi)),0)%></td>
                                    <td colspan="2" class="ratakanan"><%=formatnumber(cdbl(to_nol(ratiocaco3shi)),0)%></td>
                                </tr>
                                                              
                            </table>
                        </td>
                    </tr>
                    
                    <tr>
                        <td align="center">V</td>
                        <td colspan="6" class="tebal">PENGIRIMAN CPO DAN PK</td>
                        <td colspan="6" class="tebal">LAPORAN LOSSES</td>
                    </tr>
                    <tr>
                        <td></td>
                        <td>1. Pengiriman CPO</td>
                        <td>Kg</td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(cpokirimhi)),0)%></td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(getcpokirimshi)),0)%></td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(getcpokirimsbi)),0)%></td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(getcpokirimsti)),0)%></td>
                        
                        <td colspan="5" rowspan="5" >

                            <%
                            
                                dim sqllosses,rslosses
                                Dim PRODDATE,EMPTYBUNCHOIL,USBOIL,PCFOIL,NPCOIL,FINALEFFLUENTOIL
                                Dim SLUDGECENTRIFUGEOIL,CYCLONEFIBERPK,LTDSIPK,LTDSIIPK,CLAYBATHPK,FLINFBPK
                                Dim TOTALOIL,TOTALPK
                                sqllosses = "select PRODDATE,EMPTYBUNCHOIL,USBOIL,PCFOIL,NPCOIL,FINALEFFLUENTOIL,SLUDGECENTRIFUGEOIL,CYCLONEFIBERPK,LTDSIPK,LTDSIIPK,CLAYBATHPK,FLINFBPK,TOTALOIL,TOTALPK " & _ 
                                " from PRODUCTIONLOSSES where proddate=to_date('" & tglnow & "','mm/dd/yyyy')"
                                Set rslosses = Server.CreateObject("ADODB.Recordset")
                                rslosses.CursorType = 3
                                rslosses.CursorLocation = 3
                                rslosses.Open sqllosses, connstr
                                if not rslosses.eof then 
                                    PRODDATE = rslosses("PRODDATE")
                                    EMPTYBUNCHOIL = rslosses("EMPTYBUNCHOIL")
                                    USBOIL = rslosses("USBOIL")
                                    PCFOIL = rslosses("PCFOIL")
                                    NPCOIL = rslosses("NPCOIL")
                                    FINALEFFLUENTOIL = rslosses("FINALEFFLUENTOIL")
                                    SLUDGECENTRIFUGEOIL = rslosses("SLUDGECENTRIFUGEOIL")
                                    CYCLONEFIBERPK = rslosses("CYCLONEFIBERPK")
                                    LTDSIPK = rslosses("LTDSIPK")
                                    LTDSIIPK = rslosses("LTDSIIPK")
                                    CLAYBATHPK = rslosses("CLAYBATHPK")
                                    FLINFBPK = rslosses("FLINFBPK")
                                    TOTALOIL = rslosses("TOTALOIL")
                                    TOTALPK = rslosses("TOTALPK")
                                end if
                                rslosses.close 
                                set rslosses = nothing 
                            
                            %>

                            <table class="tabeltengah" width="600" border="0">
                                 <tr>
                                    <th rowspan="" colspan="2" class="">OIL LOSS</th>
                                    <th rowspan="">% / TBS</th>
                                    <th colspan="">KERNEL LOSS</th>
                                    <th colspan=""></th>
                                    <th colspan="2">% / TBS</th>
                                </tr>
                                 <tr>
                                    <td></td>
                                    <td >Standard</td>
                                    <td></td>
                                    <td></td>
                                    <td>Standard</td>
                                    <td></td>
                                </tr>
                                <tr>
                                    <td colspan="">1. Empty Bunch</td>
                                    <td class="ratatengah">0.43</td>
                                    <td class="ratatengah"><%=EMPTYBUNCHOIL%></td>
                                    <td>1. Cyclone Fiber</td>
                                    <td class="ratatengah">0.23</td>
                                    <td class="ratatengah"><%=CYCLONEFIBERPK%></td>
                                </tr>
                                <tr>
                                    <td colspan="1">2. USB</td>
                                    <td class="ratatengah">0.02</td>
                                    <td class="ratatengah"><%=USBOIL%></td>
                                    <td>2. LTDS I</td>
                                    <td class="ratatengah">0.01</td>
                                    <td class="ratatengah"><%=LTDSIPK%></td>
                                </tr>

                                 <tr>
                                    <td colspan="1">3. Press Cake fibre</td>
                                    <td class="ratatengah">0.38</td>
                                    <td class="ratatengah"><%=PCFOIL%></td>
                                    <td>&nbsp;&nbsp;&nbsp;&nbsp;LTDS II</td>
                                    <td class="ratatengah">0.02</td>
                                    <td class="ratatengah"><%=LTDSIIPK%></td>
                                </tr>
                                <tr>
                                    <td colspan="1">4. Nut  Press Cake</td>
                                    <td class="ratatengah">0.10</td>
                                    <td class="ratatengah"><%=NPCOIL%></td>
                                    <td>3. Claybath</td>
                                    <td class="ratatengah">0.09</td>
                                    <td class="ratatengah"><%=CLAYBATHPK%></td>
                                </tr>
                                <tr>
                                    <td colspan="1">5. Final Effluent</td>
                                    <td class="ratatengah">0.50</td>
                                    <td class="ratatengah"><%=FINALEFFLUENTOIL%></td>
                                    <td>4. Fruit Loss in EB</td>
                                    <td class="ratatengah">0.03</td>
                                    <td class="ratatengah"><%=FLINFBPK%></td>
                                </tr>
                                <tr>
                                    <td colspan="1" class="tebal">Total</td>
                                    <td class="tebal ratatengah">1.43</td>
                                    <td class="tebal ratatengah"><%=TOTALOIL%></td>
                                    <td colspan="1" class="tebal">Total</td>
                                    <td class="tebal ratatengah">0.38</td>
                                    <td class="tebal ratatengah"><%=TOTALPK%></td>
                                </tr>
                                <tr>
                                    <td colspan="1">6. Sludge Centrifuge</td>
                                    <td class="ratatengah">0.12</td>
                                    <td class="ratatengah"><%=SLUDGECENTRIFUGEOIL%></td>
                                    <td colspan=3></td>
                                   
                                </tr>
                             
                            </table>
                        </td>
                        
                    </tr>
                    <tr>
                        <td></td>
                        <td>2. Pengiriman PK</td>
                        <td>Kg</td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(totpkdohi)),0)%></td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(getpkkirimshi)),0)%></td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(getpkkirimsbi)),0)%></td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(getpkkirimsti)),0)%></td>
                       
                    </tr>
                    <tr>
                        <td></td>
                        <td>3. Pengiriman CKG</td>
                        <td>Kg</td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(totckgdohi)),0)%></td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(0)),0)%></td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(0)),0)%></td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(0)),0)%></td>
            
                        
                    </tr>
                    <tr>
                        <td></td>
                        <td>4. Pengiriman FIBRE</td>
                        <td>Kg</td>
                        <td class="ratakanan">00</td>
                        <td class="ratakanan">00</td>
                        <td class="ratakanan">00</td>
                        <td class="ratakanan">00</td>
                        
                    </tr>
                    <tr>
                        <td></td>
                        <td>5. Adjusment CPO</td>
                        <td>Kg</td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(cpoadjhi)),0)%></td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(cpoadjshi)),0)%></td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(cpoadjsbi)),0)%></td>
                        <td class="ratakanan"><%=formatnumber(cdbl(to_nol(cpoadjsti)),0)%></td>
                        
                    </tr>

                    <tr>
                        <td align="center">VI</td>
                        <td colspan="6" class="tebal">STOCK & MUTU CPO</td>
                        <td colspan="6" class="tebal">TENAGA KERJA</td>
                    </tr>
                    
                        <%
                        
                            ' GET CPOTANK HARI INI'
                            Dim rscpotankhi,notkhi,st,st1,cpoisitankdtl,ttlstockcpo,cst_ot,totalotcpo
                            sqltbs = "select proddate,tankid,cpoisitank,ffatankcpo,airtankcpo,kotorantankcpo " & _
                            "from productiontankcpo where proddate=to_date('" & tglnow & "','mm-dd-yyyy')"
                            'response.write sqltbs
                            Set rscpotankhi = Server.CreateObject("ADODB.Recordset")
                            rscpotankhi.CursorType = 3
                            rscpotankhi.CursorLocation = 3
                            rscpotankhi.Open sqltbs, connstr
                            
                                   
                        %>
                    
                    <tr>
                        <td></td>
                        <td colspan="6"  width="150px" class="">
                            <table>
                            <tr>
                                <td colspan="" class="tebal ratatengah" style="width:195px">Storage</td>
                                <td colspan="" class="tebal ratatengah">Satuan</td>
                                <td colspan="" class="tebal ratatengah" style="width:90px">HI</td>
                                <td colspan="" class="tebal ratatengah" style="width:50px">FFA</td>
                                <td colspan="" class="tebal ratatengah" style="width:50px">AIR</td>
                                <td colspan="" class="tebal ratatengah" style="width:50px">KOTORAN</td>
                            </tr>
                            <%
                                notkhi = 1
                                While Not rscpotankhi.EOF
                                    st = rscpotankhi("tankid")
                                    cpoisitankdtl = rscpotankhi("cpoisitank")
                                    if st = "1" then
                                        st1 = "1. Stock Storage - 01"
                                    end if
                                    if st = "2" then
                                        st1 = "2. Stock Storage - 02"
                                    end if
                                    if st = "3" then
                                        st1 = "3. Stock Storage - 03"
                                    end if %>
                            <tr>
                                <td><%=st1%></td>
                                <td>Kg</td>
                                <td class="ratakanan"><%=formatnumber(rscpotankhi("cpoisitank"),0)%></td>
                                <td class="ratakanan"><%=FbpNum2(rscpotankhi("ffatankcpo"))%></td>
                                <td class="ratakanan"><%=FbpNum2(rscpotankhi("airtankcpo"))%></td>
                                <td class="ratakanan"><%=FbpNum2(rscpotankhi("kotorantankcpo"))%></td>
                            </tr>
                            <%
                                rscpotankhi.MoveNext
                                notkhi      = notkhi+1
                                ttlstockcpo = cdbl(ttlstockcpo) + cdbl(cpoisitankdtl)
                                Wend
                                rscpotankhi.close 
                                set rscpotankhi = nothing

                                cst_ot      = cdbl(cpocsthi)+cdbl(cpoothi)
                                totalotcpo  = cdbl(ttlstockcpo) + cdbl(cst_ot)
                            %>
                            <tr>
                                <td>4. Stock Klari (CST+OT)</td>
                                <td>Kg</td>
                                <td class="ratakanan"><%=formatnumber(cst_ot,0)%></td>
                                <td></td>
                                <td></td>
                                <td></td>
                            </tr>
                            <tr>
                                <td class="tebal">Total Stock CPO</td>
                                <td>Kg</td>
                                <td class="ratakanan tebal"><%=formatnumber(totalotcpo,0)%></td>
                                <td></td>
                                <td></td>
                                <td></td>
                            </tr>
                            </table>
                        </td>
                        <td colspan="5" rowspan="1">
                            <%
                                
                                ' GET HK HARI INI'
                                Dim rshkhi, ttlhk,ttlcuti,ttlsakit,ttlizin,ttlmangkir
                                Dim hk,cuti,sakit,izin,mangkir
                                sqltbs = "select shift,hk,cuti,sakit,izin,mangkir " & _
                                "from productionhk where proddate=to_date('" & tglnow & "','mm-dd-yyyy')"
                                'response.write sqltbs
                                Set rshkhi = Server.CreateObject("ADODB.Recordset")
                                rshkhi.CursorType = 3
                                rshkhi.CursorLocation = 3
                                rshkhi.Open sqltbs, connstr
                                
                                    
                            %>
                            <table class="ratatengah" width="600" >
                                <tr>
                                    <th>Shift</th>
                                    <th>HK</th>
                                    <th>Cuti</th>
                                    <th>Sakit</th>
                                    <th>Izin/P1</th>
                                    <th>Mangkir</th>
                                </tr>
                               <% While Not rshkhi.EOF

                                    hk      = rshkhi("hk")
                                    cuti    = rshkhi("cuti")
                                    sakit   = rshkhi("sakit")
                                    izin    = rshkhi("izin")
                                    mangkir = rshkhi("mangkir")
                                   
                                %>
                                <tr>
                                    <td><%=rshkhi("shift")%></td>
                                    <td><%=rshkhi("hk")%></td>
                                    <td><%=rshkhi("cuti")%></td>
                                    <td><%=rshkhi("sakit")%></td>
                                    <td><%=rshkhi("izin")%></td>
                                    <td><%=rshkhi("mangkir")%></td>
                                </tr>

                                <%
                                    rshkhi.MoveNext
                                    Wend
                                    ttlhk       = cdbl(ttlhk)+cdbl(hk)
                                    ttlcuti     = cdbl(ttlcuti)+cdbl(cuti)
                                    ttlsakit    = cdbl(ttlsakit)+cdbl(sakit)
                                    ttlizin     = cdbl(ttlsakit)+cdbl(izin)
                                    ttlmangkir  = cdbl(ttlmangkir)+cdbl(mangkir)
                                    rshkhi.close 
                                    set rshkhi = nothing
                                %>
                                <tr class="ratatengah tebal">
                                    <td>Total</td>
                                    <td><%=ttlhk%></td>
                                    <td><%=ttlcuti%></td>
                                    <td><%=ttlsakit%></td>
                                    <td><%=ttlizin%></td>
                                    <td><%=ttlmangkir%></td>
                                </tr>
                            </table>
                        
                        </td>
                    </tr>
                  
                       
                    
                    
                    
                    <%
                    ' GET PKTANK HARI INI'
                            Dim rspktankhi,totalstockpk,totalpkhitung
                            Dim bulkingsiloisihi,bulkingsiloffahi,bulkingsiloairhi,bulkingsilokotoranhi
                            sqltbs = "select proddate,siloid,bulkingsiloisi,bulkingsiloffa,bulkingsiloair,bulkingsilokotoran " & _
                            "from productionbulkingsilo where proddate=to_date('" & tglnow & "','mm-dd-yyyy')"
                            'response.write sqltbs
                            Set rspktankhi = Server.CreateObject("ADODB.Recordset")
                            rspktankhi.CursorType = 3
                            rspktankhi.CursorLocation = 3
                            rspktankhi.Open sqltbs, connstr
                            if not rspktankhi.eof then
                                totalpkhitung = cdbl(rspktankhi("bulkingsiloisi"))+cdbl(kernelsilohi)
                                totalstockpk  = totalpkhitung
                                bulkingsiloisihi = rspktankhi("bulkingsiloisi")
                                bulkingsiloffahi = rspktankhi("bulkingsiloffa")
                                bulkingsiloairhi = rspktankhi("bulkingsiloair")
                                bulkingsilokotoranhi = rspktankhi("bulkingsilokotoran")
                            end if
                    %>
                    <tr>
                        <td>VII</td>
                        <td class="tebal">STOCK & MUTU PK</td>
                        <td class="tebal ratatengah">Satuan</td>
                        <td class="tebal ratatengah">HI</td>
                        <td class="tebal ratatengah">FFA</td>
                        <td class="tebal ratatengah">Kdr Air</td>
                        <td class="tebal ratatengah">Kotoran</td>
                        <td class="tebal">STOCK CANGKANG & FIBRE</td>
                       
                    </tr>
                    <tr>
                        <td></td>
                        <td>1. Bulking Silo</td>
                        <td>Kg</td>
                        <td class="ratakanan"><%=formatnumber(bulkingsiloisihi,0)%></td>
                        <td class="ratakanan"><%=formatnumber(bulkingsiloffahi,0)%></td>
                        <td class="ratakanan"><%=formatnumber(bulkingsiloairhi,0)%></td>
                        <td class="ratakanan"><%=formatnumber(bulkingsilokotoranhi,0)%></td>
                        <td rowspan="5">
                            <table class="ratatengah" width="600">
                                <tr>
                                    <th>STOCK</th>
                                    <th>Satuan</th>
                                    <th>HI</th>
                                    <th>SHI</th>
                                </tr>
                                <tr>
                                    <td class="ratakiri">Cangkang</td>
                                    <td>Kg</td>
                                    <td><%=formatnumber(ckgviewhi,0)%></td>
                                    <td><%=formatnumber(qtyckgshi,0)%></td>
                                </tr>
                                <tr>
                                    <td class="ratakiri">Fibre</td>
                                    <td>Kg</td>
                                    <td>0</td>
                                    <td>0</td>
                                </tr>
                            </table>
                        
                        </td>
                      
                    </tr>
                    <tr>
                        <td></td>
                        <td>2. Kernel Silo</td>
                        <td>Kg</td>
                        <td class="ratakanan"><%=formatnumber(kernelsilohi,0)%></td>
                        <td></td>
                        <td></td>
                        <td></td>
                    </tr>
                    <tr>
                        <td></td>
                        <td>3. Packing Kernel</td>
                        <td>Kg</td>
                        <td class="ratakanan"></td>
                        <td></td>
                        <td></td>
                        <td></td>
                    </tr>
                    <tr>
                        <td></td>
                        <td>4. Inti Dilantai</td>
                        <td>Kg</td>
                        <td class="ratakanan"></td>
                        <td></td>
                        <td></td>
                        <td></td>
                    </tr>
                    <tr>
                        <td></td>
                        <td class="tebal">Total Stock PK</td>
                        <td>Kg</td>
                        <td class="ratakanan tebal"><%=formatnumber(totalstockpk,0)%></td>
                        <td></td>
                        <td></td>
                        <td></td>
                    </tr>
                    <tr>
                        <td>VIII</td>
                        <td class="tebal" colspan="12">TBS MASUK</td>
                    </tr>
                                        
                    <tr>
                        <td colspan="12">
                            <table class="ratatengah" width="1000">
                                <tr>
                                    <th rowspan="2">No</th>
                                    <th class="" rowspan="2" style="width: 250px;">Supplier</th>
                                    <th class="ratatengah" colspan="3">Tonase</th>
                                    <th class="ratatengah" colspan="3">Potongan KG</th>
                                  
                                    
                                </tr>
                                <tr>
                                    <th class="ratatengah">HI</th>
                                    <th class="ratatengah">SHI</th>
                                    <th class="ratatengah">SBI</th>
                                    <th class="ratatengah">HI</th>
                                    <th class="ratatengah">SHI</th>
                                    <th class="ratatengah">SBI</th>
                                  
                                </tr>

                                    <%
                                    
                                    ' GET TBSMASUK HI,SHI'
                                    Dim rstbsmasuk,rsunit,unit,unitgwe,unitpse,unitske,unitext
                                    sqltbs = "SELECT af.AreaTypeName,a.unit, a.AreaName, SUM(CASE WHEN h.DocketDate=TO_DATE('" & tglnow & "','mm/dd/yyyy') then d.FFBWeight+d.sortasiwgt else 0 end) AS TonaseHariIni,  " & _
                                    " SUM(CASE WHEN h.DocketDate=TO_DATE('" & tglnow & "','mm/dd/yyyy') then d.sortasiwgt else 0 end) AS SortasiHariIni, " & _
                                    " SUM(CASE WHEN h.DocketDate between TO_DATE('" & tglblnawal & "','mm/dd/yyyy') and TO_DATE('" & tglnow & "','mm/dd/yyyy') then d.FFBWeight+d.sortasiwgt else 0 end) AS TonaseSDHariIni, " & _
                                    " SUM(CASE WHEN h.DocketDate between TO_DATE('" & tglblnawal & "','mm/dd/yyyy') and TO_DATE('" & tglnow & "',   'mm/dd/yyyy') then d.sortasiwgt else 0 end) AS   SortasiSDHariIni, " & _
                                    " SUM(d.FFBWeight+d.sortasiwgt) AS TonaseBlnIni, SUM(d.sortasiwgt) AS SortasiBulanIni, " & _
                                    " 0 AS RpBulanIni,0 AS RpHariIni " & _
                                    " FROM LHPTIKETHDR h INNER JOIN LHPFFBWBTRN d ON h.DocketID=d.DocketID " & _
                                    " INNER JOIN LHPAREAMST a ON d.AreaID=a.AreaID " & _
                                    " INNER JOIN LHPAREATYPEREF af ON a.AreaTypeID=af.AreaTypeID " & _
                                    " WHERE h.DocketDate between TO_DATE('" & tglthnawal & "','mm/dd/yyyy') and TO_DATE('" & tglnow & "','mm/dd/yyyy') AND h.ProductID=2 " & _
                                    " GROUP BY af.AreaTypeName,a.unit, a.AreaName order by af.AreaTypeName asc"
                                    'response.write sqltbs
                                    Set rstbsmasuk = Server.CreateObject("ADODB.Recordset")
                                    rstbsmasuk.CursorType = 3
                                    rstbsmasuk.CursorLocation = 3
                                    rstbsmasuk.Open sqltbs, connstr
                                    dim notbsmasuk
                                    notbsmasuk = 1
                                    totalunitgwe    = 0
                                    totalunitpse    = 0
                                    totalunitske    = 0
                                    totalunitext    = 0
                                    While Not rstbsmasuk.EOF 
                                       
                                    %>
                                    <tr>
                                        <td><%=notbsmasuk%></td>
                                        <td class="ratakiri"><%=rstbsmasuk("AreaName")%></td>
                                        <td class="ratakanan"><%=formatnumber(rstbsmasuk("TonaseHariIni"),0)%></td>
                                        <td class="ratakanan"><%=formatnumber(rstbsmasuk("TonaseSDHariIni"),0)%></td>
                                        <td class="ratakanan"><%=formatnumber(rstbsmasuk("TonaseBlnIni"),0)%></td>
                                        <td class="ratakanan"><%=formatnumber(rstbsmasuk("SortasiHariIni"),0)%></td>
                                        <td class="ratakanan"><%=formatnumber(rstbsmasuk("SortasiSDHariIni"),0)%></td>
                                        <td class="ratakanan"><%=formatnumber(rstbsmasuk("SortasiBulanIni"),0)%></td>
                                    </tr>
                                <% 
                                    
                                    rstbsmasuk.MoveNext
                                    notbsmasuk      = notbsmasuk+1
                                    wend
                                %>
                            </table>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="12">

                            <%
                            
                                    ' GET TOTAL TBSMASUK HI,SHI'
                                    Dim rstbsmasukttl,totalunitgwe,totalunitpse,totalunitske,totalunitext
                                    sqltbs = "select t.unit,sum(t.TonaseHariIni) as TonaseHariInittl,sum(t.SortasiHariIni) as SortasiHariInittl,sum(t.TonaseSDHariIni) as TonaseSDHariIni,sum(t.SortasiSDHariIni) as SortasiSDHariIni,sum(t.TonaseBlnIni) as TonaseBlnInittl,sum(t.SortasiBulanIni) as SortasiBulanInittl " & _
                                    " from( " & _
                                    " SELECT a.unit,   " & _
                                    " SUM(CASE WHEN h.DocketDate=TO_DATE('" & tglnow & "','mm/dd/yyyy') then d.FFBWeight+d.sortasiwgt else 0 end) AS TonaseHariIni, " & _
                                    " SUM(CASE WHEN h.DocketDate=TO_DATE('" & tglnow & "',   'mm/dd/yyyy') then d.sortasiwgt else 0 end) AS   SortasiHariIni, " & _
                                    " SUM(CASE WHEN h.DocketDate between TO_DATE('" & tglblnawal & "','mm/dd/yyyy') and TO_DATE('" & tglnow & "','mm/dd/yyyy') then d.FFBWeight+d.sortasiwgt else 0 end) AS TonaseSDHariIni, " & _
                                    " SUM(CASE WHEN h.DocketDate between TO_DATE('" & tglblnawal & "','mm/dd/yyyy') and TO_DATE('" & tglnow & "',   'mm/dd/yyyy') then d.sortasiwgt else 0 end) AS   SortasiSDHariIni, " & _
                                    " SUM(d.FFBWeight+d.sortasiwgt) AS TonaseBlnIni, SUM(d.sortasiwgt) AS SortasiBulanIni " & _
                                    " FROM LHPTIKETHDR h  " & _
                                    " INNER JOIN LHPFFBWBTRN d ON h.DocketID=d.DocketID  " & _
                                    " INNER JOIN LHPAREAMST a ON d.AreaID=a.AreaID " & _
                                    " INNER JOIN LHPAREATYPEREF af ON a.AreaTypeID=af.AreaTypeID  " & _
                                    " WHERE h.DocketDate between TO_DATE('" & tglthnawal & "','mm/dd/yyyy')  " & _
                                    " and TO_DATE('" & tglnow & "','mm/dd/yyyy') AND h.ProductID=2  " & _
                                    " GROUP BY a.unit order by af.AreaTypeName asc " & _
                                    " ) t " & _
                                    " group by t.unit,t.TonaseHariIni,t.SortasiHariIni,t.TonaseSDHariIni,t.SortasiSDHariIni,t.TonaseBlnIni,t.SortasiBulanIni"
                                    'response.write sqltbs
                                    Set rstbsmasukttl = Server.CreateObject("ADODB.Recordset")
                                    rstbsmasukttl.CursorType = 3
                                    rstbsmasukttl.CursorLocation = 3
                                    rstbsmasukttl.Open sqltbs, connstr
                            
                            %>

                            <table class="ratatengah" width="1000">
                                <tr>
                                    <th rowspan="2">No</th>
                                    <th class="" rowspan="2" style="width: 250px;">Supplier</th>
                                    <th class="ratatengah" colspan="3">Tonase</th>
                                  
                                </tr>
                                <tr>
                                    <th class="ratatengah">HI</th>
                                    <th class="ratatengah">SHI</th>
                                    <th class="ratatengah">SBI</th>
                                  
                                </tr>
                                <%  
                                        dim notbstotal
                                        notbstotal =1
                                        While Not rstbsmasukttl.EOF %>
                                    

                                <tr>
                                    <td><%=notbstotal%></td>
                                    <td class="ratakiri"><%=rstbsmasukttl("unit")%></td>
                                    <td class="ratakanan"><%=formatnumber(rstbsmasukttl("TonaseHariInittl"),0)%></td>
                                    <td class="ratakanan"><%=formatnumber(rstbsmasukttl("TonaseSDHariIni"),0)%></td>
                                    <td class="ratakanan"><%=formatnumber(rstbsmasukttl("TonaseBlnInittl"),0)%></td>
                                  
                                </tr>
                               <% rstbsmasukttl.MoveNext
                                    notbstotal      = notbstotal+1
                                    wend  %>
                            </table>
                        </td>
                        
                    </tr>
                                       
                </tbody>
                <table>
                    Catatan: <span><%=keterangantxt%></span>
                </tr>
               
                </div>
                
			</table>
			
		</div>
	</body>
</html>