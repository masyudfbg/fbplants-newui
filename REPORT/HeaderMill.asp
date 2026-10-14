  <tr bgcolor="#F7F7F7">
    <td rowspan="2"><div align="center">CODE</div></td>
    <td rowspan="2"><div align="center">STATION</div></td>
    <td colspan="6"><div align="center">BULAN INI </div></td>
    <td colspan="6"><div align="center">S.D BULAN INI </div></td>
    <td colspan="3" rowspan="2"><div align="center">ANGGARAN<br />
      TAHUNAN</div></td>
  </tr>
  <tr bgcolor="#F7F7F7">
    <td colspan="3"><div align="center">REALISASI</div></td>
    <td colspan="3"><div align="center">ANGGARAN</div></td>
    <td colspan="3"><div align="center">REALISASI</div></td>
    <td colspan="3"><div align="center">ANGGARAN</div></td>
  </tr>
  <tr bgcolor="#F7F7F7">
    <td rowspan="6"><div align="center"></div></td>
    <td>TBS PROSES (KG)</td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TM_FFB_ACTUAL),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TM_FFB_BUDGET),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TY_FFB_ACTUAL),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TY_FFB_BUDGET),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_FFB_ANNUAL),2)%></div></td>
  </tr>
  <tr bgcolor="#F7F7F7">
    <td><div align="LEFT">PRODUKSI CPO (KG) </div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TM_CPO_ACTUAL),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TM_CPO_BUDGET),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TY_CPO_ACTUAL),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TY_CPO_BUDGET),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_CPO_ANNUAL),2)%></div></td>
  </tr>
  <tr bgcolor="#F7F7F7">
    <td><div align="LEFT">PRODUKSI KER (KG)</div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TM_KER_ACTUAL),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TM_KER_BUDGET),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TY_KER_ACTUAL),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TY_KER_BUDGET),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_KER_ANNUAL),2)%></div></td>
  </tr>
  <tr bgcolor="#F7F7F7">
    <td><div align="LEFT">PALM (KG) </div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TM_CPO_ACTUAL)+cdbl(v_TM_KER_ACTUAL),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TM_CPO_BUDGET)+cdbl(v_TM_KER_BUDGET),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TY_CPO_ACTUAL)+cdbl(v_TY_KER_ACTUAL),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TY_CPO_BUDGET)+cdbl(v_TY_KER_BUDGET),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_CPO_ANNUAL)+cdbl(v_KER_ANNUAL),2)%></div></td>
  </tr>
  <tr bgcolor="#F7F7F7">
    <td rowspan="2"><div align="center"></div></td>
    <td rowspan="2"><div align="center">Total</div></td>
    <td colspan="2"><div align="center">Rp./Kg.</div></td>
    <td rowspan="2"><div align="center">Total</div></td>
    <td colspan="2"><div align="center">Rp./Kg.</div></td>
    <td rowspan="2"><div align="center">Total</div></td>
    <td colspan="2"><div align="center">Rp./Kg.</div></td>
    <td rowspan="2"><div align="center">Total</div></td>
    <td colspan="2"><div align="center">Rp./Kg.</div></td>
    <td rowspan="2"><div align="center">Total</div></td>
    <td colspan="2"><div align="center">Rp./Kg.</div></td>
  </tr>
  <tr bgcolor="#F7F7F7">
    <td><div align="center">Palm</div></td>
    <td><div align="center">TBS</div></td>
    <td><div align="center">Palm</div></td>
    <td><div align="center">TBS</div></td>
    <td><div align="center">Palm</div></td>
    <td><div align="center">TBS</div></td>
    <td><div align="center">Palm</div></td>
    <td><div align="center">TBS</div></td>
    <td><div align="center">Palm</div></td>
    <td><div align="center">TBS</div></td>
  </tr>
  <%elseif request.Form("vJenisReport")="Machine" Then%>
  <tr bgcolor="#F7F7F7">
    <td width="10" rowspan="2" align="center">CODE</td>
    <td width="88" rowspan="2" align="center"><div align="center"></div>STATION</td>
    <td width="152" rowspan="2" align="center"> <div align="center"></div>ACTIVITY</td>
    <td colspan="6"><div align="center">BULAN INI </div></td>
    <td colspan="6"><div align="center">S.D BULAN INI </div></td>
    <td colspan="3" rowspan="2"><div align="center">ANGGARAN<br />
      TAHUNAN</div></td>
  </tr>
  <tr bgcolor="#F7F7F7">
    <td colspan="3"><div align="center">REALISASI</div></td>
    <td colspan="3"><div align="center">ANGGARAN</div></td>
    <td colspan="3"><div align="center">REALISASI</div></td>
    <td colspan="3"><div align="center">ANGGARAN</div></td>
  </tr>
  <tr bgcolor="#F7F7F7">
    <td rowspan="6">&nbsp;</td>
    <td>&nbsp;</td>
    <td><div align="LEFT">TBS PROSES (KG) </div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TM_FFB_ACTUAL),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TM_FFB_BUDGET),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TY_FFB_ACTUAL),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TY_FFB_BUDGET),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_FFB_ANNUAL),2)%></div></td>
  </tr>
  <tr bgcolor="#F7F7F7">
    <td>&nbsp;</td>
    <td><div align="LEFT">PRODUKSI CPO (KG)  </div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TM_CPO_ACTUAL),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TM_CPO_BUDGET),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TY_CPO_ACTUAL),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TY_CPO_BUDGET),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_CPO_ANNUAL),2)%></div></td>
  </tr>
  <tr bgcolor="#F7F7F7">
    <td>&nbsp;</td>
    <td><div align="LEFT">PRODUKSI KER (KG) </div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TM_KER_ACTUAL),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TM_KER_BUDGET),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TY_KER_ACTUAL),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TY_KER_BUDGET),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_KER_ANNUAL),2)%></div></td>
  </tr>
  <tr bgcolor="#F7F7F7">
    <td>&nbsp;</td>
    <td><div align="LEFT">PALM (KG)  </div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TM_CPO_ACTUAL)+cdbl(v_TM_KER_ACTUAL),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TM_CPO_BUDGET)+cdbl(v_TM_KER_BUDGET),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TY_CPO_ACTUAL)+cdbl(v_TY_KER_ACTUAL),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TY_CPO_BUDGET)+cdbl(v_TY_KER_BUDGET),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_CPO_ANNUAL)+cdbl(v_KER_ANNUAL),2)%></div></td>
  </tr>
  <tr bgcolor="#F7F7F7">
    <td rowspan="2">&nbsp;</td>
    <td rowspan="2"><div align="center"></div></td>
    <td width="31" rowspan="2"><div align="center">Total</div></td>
    <td colspan="2"><div align="center">Rp./Kg.</div></td>
    <td width="31" rowspan="2"><div align="center">Total</div></td>
    <td colspan="2"><div align="center">Rp./Kg.</div></td>
    <td width="31" rowspan="2"><div align="center">Total</div></td>
    <td colspan="2"><div align="center">Rp./Kg.</div></td>
    <td width="31" rowspan="2"><div align="center">Total</div></td>
    <td colspan="2"><div align="center">Rp./Kg.</div></td>
    <td width="31" rowspan="2"><div align="center">Total</div></td>
    <td colspan="2"><div align="center">Rp./Kg.</div></td>
  </tr>
  <tr bgcolor="#F7F7F7">
    <td width="30"><div align="center">Palm</div></td>
    <td width="28"><div align="center">TBS</div></td>
    <td width="30"><div align="center">Palm</div></td>
    <td width="28"><div align="center">TBS</div></td>
    <td width="30"><div align="center">Palm</div></td>
    <td width="28"><div align="center">TBS</div></td>
    <td width="30"><div align="center">Palm</div></td>
    <td width="28"><div align="center">TBS</div></td>
    <td width="30"><div align="center">Palm</div></td>
    <td width="28"><div align="center">TBS</div></td>
  </tr>
  <%elseif request.Form("vJenisReport")="StationMachine" Then%>
  <tr bgcolor="#F7F7F7">
    <td width="10" rowspan="2" align="center">CODE</td>
    <td width="88" rowspan="2" align="center"><div align="center"></div>STATION</td>
    <td width="152" rowspan="2" align="center"> <div align="center"></div>MACHINE</td>
    <td colspan="6"><div align="center">BULAN INI </div></td>
    <td colspan="6"><div align="center">S.D BULAN INI </div></td>
    <td colspan="3" rowspan="2"><div align="center">ANGGARAN<br />
      TAHUNAN</div></td>
  </tr>
  <tr bgcolor="#F7F7F7">
    <td colspan="3"><div align="center">REALISASI</div></td>
    <td colspan="3"><div align="center">ANGGARAN</div></td>
    <td colspan="3"><div align="center">REALISASI</div></td>
    <td colspan="3"><div align="center">ANGGARAN</div></td>
  </tr>
  <tr bgcolor="#F7F7F7">
    <td rowspan="6">&nbsp;</td>
    <td>&nbsp;</td>
    <td><div align="LEFT">TBS PROSES (KG) </div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TM_FFB_ACTUAL),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TM_FFB_BUDGET),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TY_FFB_ACTUAL),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TY_FFB_BUDGET),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_FFB_ANNUAL),2)%></div></td>
  </tr>
  <tr bgcolor="#F7F7F7">
    <td>&nbsp;</td>
    <td><div align="LEFT">PRODUKSI CPO (KG)  </div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TM_CPO_ACTUAL),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TM_CPO_BUDGET),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TY_CPO_ACTUAL),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TY_CPO_BUDGET),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_CPO_ANNUAL),2)%></div></td>
  </tr>
  <tr bgcolor="#F7F7F7">
    <td>&nbsp;</td>
    <td><div align="LEFT">PRODUKSI KER (KG)</div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TM_KER_ACTUAL),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TM_KER_BUDGET),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TY_KER_ACTUAL),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TY_KER_BUDGET),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_KER_ANNUAL),2)%></div></td>
  </tr>
  <tr bgcolor="#F7F7F7">
    <td>&nbsp;</td>
    <td><div align="LEFT">PALM (KG)  </div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TM_CPO_ACTUAL)+cdbl(v_TM_KER_ACTUAL),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TM_CPO_BUDGET)+cdbl(v_TM_KER_BUDGET),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TY_CPO_ACTUAL)+cdbl(v_TY_KER_ACTUAL),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TY_CPO_BUDGET)+cdbl(v_TY_KER_BUDGET),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_CPO_ANNUAL)+cdbl(v_KER_ANNUAL),2)%></div></td>
  </tr>
  <tr bgcolor="#F7F7F7">
    <td rowspan="2">&nbsp;</td>
    <td rowspan="2"><div align="center"></div></td>
    <td width="31" rowspan="2"><div align="center">Total</div></td>
    <td colspan="2"><div align="center">Rp./Kg.</div></td>
    <td width="31" rowspan="2"><div align="center">Total</div></td>
    <td colspan="2"><div align="center">Rp./Kg.</div></td>
    <td width="31" rowspan="2"><div align="center">Total</div></td>
    <td colspan="2"><div align="center">Rp./Kg.</div></td>
    <td width="31" rowspan="2"><div align="center">Total</div></td>
    <td colspan="2"><div align="center">Rp./Kg.</div></td>
    <td width="31" rowspan="2"><div align="center">Total</div></td>
    <td colspan="2"><div align="center">Rp./Kg.</div></td>
  </tr>
  <tr bgcolor="#F7F7F7">
    <td width="30"><div align="center">Palm</div></td>
    <td width="28"><div align="center">TBS</div></td>
    <td width="30"><div align="center">Palm</div></td>
    <td width="28"><div align="center">TBS</div></td>
    <td width="30"><div align="center">Palm</div></td>
    <td width="28"><div align="center">TBS</div></td>
    <td width="30"><div align="center">Palm</div></td>
    <td width="28"><div align="center">TBS</div></td>
    <td width="30"><div align="center">Palm</div></td>
    <td width="28"><div align="center">TBS</div></td>
  </tr>
  <%elseif request.Form("vJenisReport")="JobMachine" Then%>
  <tr bgcolor="#F7F7F7">
    <td width="10" rowspan="2" align="center">CODE</td>
    <td width="10" rowspan="2" align="center">STATION</td>
    <td width="88" rowspan="2" align="center"><div align="center"></div>ACTIVITY</td>
    <td width="152" rowspan="2" align="center"> <div align="center"></div>MACHINE</td>
    <td colspan="6"><div align="center">BULAN INI </div></td>
    <td colspan="6"><div align="center">S.D BULAN INI </div></td>
    <td colspan="3" rowspan="2"><div align="center">ANGGARAN<br />
      TAHUNAN</div></td>
  </tr>
  <tr bgcolor="#F7F7F7">
    <td colspan="3"><div align="center">REALISASI</div></td>
    <td colspan="3"><div align="center">ANGGARAN</div></td>
    <td colspan="3"><div align="center">REALISASI</div></td>
    <td colspan="3"><div align="center">ANGGARAN</div></td>
  </tr>
  <tr bgcolor="#F7F7F7">
    <td rowspan="6">&nbsp;</td>
    <td rowspan="6">&nbsp;</td>
    <td>&nbsp;</td>
    <td><div align="LEFT">TBS PROSES (KG) </div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TM_FFB_ACTUAL),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TM_FFB_BUDGET),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TY_FFB_ACTUAL),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TY_FFB_BUDGET),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_FFB_ANNUAL),2)%></div></td>
  </tr>
  <tr bgcolor="#F7F7F7">
    <td>&nbsp;</td>
    <td><div align="LEFT">PRODUKSI CPO (KG)  </div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TM_CPO_ACTUAL),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TM_CPO_BUDGET),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TY_CPO_ACTUAL),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TY_CPO_BUDGET),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_CPO_ANNUAL),2)%></div></td>
  </tr>
  <tr bgcolor="#F7F7F7">
    <td>&nbsp;</td>
    <td><div align="LEFT">PRODUKSI KER (KG)</div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TM_KER_ACTUAL),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TM_KER_BUDGET),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TY_KER_ACTUAL),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TY_KER_BUDGET),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_KER_ANNUAL),2)%></div></td>
  </tr>
  <tr bgcolor="#F7F7F7">
    <td>&nbsp;</td>
    <td><div align="LEFT">PALM (KG) </div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TM_CPO_ACTUAL)+cdbl(v_TM_KER_ACTUAL),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TM_CPO_BUDGET)+cdbl(v_TM_KER_BUDGET),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TY_CPO_ACTUAL)+cdbl(v_TY_KER_ACTUAL),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TY_CPO_BUDGET)+cdbl(v_TY_KER_BUDGET),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_CPO_ANNUAL)+cdbl(v_KER_ANNUAL),2)%></div></td>
  </tr>
  <tr bgcolor="#F7F7F7">
    <td rowspan="2">&nbsp;</td>
    <td rowspan="2"><div align="center"></div></td>
    <td width="31" rowspan="2"><div align="center">Total</div></td>
    <td colspan="2"><div align="center">Rp./Kg.</div></td>
    <td width="31" rowspan="2"><div align="center">Total</div></td>
    <td colspan="2"><div align="center">Rp./Kg.</div></td>
    <td width="31" rowspan="2"><div align="center">Total</div></td>
    <td colspan="2"><div align="center">Rp./Kg.</div></td>
    <td width="31" rowspan="2"><div align="center">Total</div></td>
    <td colspan="2"><div align="center">Rp./Kg.</div></td>
    <td width="31" rowspan="2"><div align="center">Total</div></td>
    <td colspan="2"><div align="center">Rp./Kg.</div></td>
  </tr>
  <tr bgcolor="#F7F7F7">
    <td width="30"><div align="center">Palm</div></td>
    <td width="28"><div align="center">TBS</div></td>
    <td width="30"><div align="center">Palm</div></td>
    <td width="28"><div align="center">TBS</div></td>
    <td width="30"><div align="center">Palm</div></td>
    <td width="28"><div align="center">TBS</div></td>
    <td width="30"><div align="center">Palm</div></td>
    <td width="28"><div align="center">TBS</div></td>
    <td width="30"><div align="center">Palm</div></td>
    <td width="28"><div align="center">TBS</div></td>
  </tr>
  <%elseif request.Form("vJenisReport")="MachineJob" Then%>
  <tr bgcolor="#F7F7F7">
    <td width="10" rowspan="2" align="center">CODE</td>
    <td width="10" rowspan="2" align="center">STATION</td>
    <td width="10" rowspan="2" align="center">MACHINE</td>
    <td width="88" rowspan="2" align="center"><div align="center"></div>ACTIVITY</td>
    <td width="152" rowspan="2" align="center"> <div align="center"></div>SOURCE</td>
    <td colspan="6"><div align="center">BULAN INI </div></td>
    <td colspan="6"><div align="center">S.D BULAN INI </div></td>
    <td colspan="3" rowspan="2"><div align="center">ANGGARAN<br />
      TAHUNAN</div></td>
  </tr>
  <tr bgcolor="#F7F7F7">
    <td colspan="3"><div align="center">REALISASI</div></td>
    <td colspan="3"><div align="center">ANGGARAN</div></td>
    <td colspan="3"><div align="center">REALISASI</div></td>
    <td colspan="3"><div align="center">ANGGARAN</div></td>
  </tr>
  <tr bgcolor="#F7F7F7">
    <td rowspan="6">&nbsp;</td>
    <td rowspan="6">&nbsp;</td>
    <td rowspan="6">&nbsp;</td>
    <td>&nbsp;</td>
    <td><div align="LEFT">TBS PROSES (KG) </div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TM_FFB_ACTUAL),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TM_FFB_BUDGET),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TY_FFB_ACTUAL),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TY_FFB_BUDGET),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_FFB_ANNUAL),2)%></div></td>
  </tr>
  <tr bgcolor="#F7F7F7">
    <td>&nbsp;</td>
    <td><div align="LEFT">PRODUKSI CPO (KG)  </div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TM_CPO_ACTUAL),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TM_CPO_BUDGET),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TY_CPO_ACTUAL),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TY_CPO_BUDGET),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_CPO_ANNUAL),2)%></div></td>
  </tr>
  <tr bgcolor="#F7F7F7">
    <td>&nbsp;</td>
    <td><div align="LEFT">PRODUKSI KER (KG)</div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TM_KER_ACTUAL),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TM_KER_BUDGET),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TY_KER_ACTUAL),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TY_KER_BUDGET),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_KER_ANNUAL),2)%></div></td>
  </tr>
  <tr bgcolor="#F7F7F7">
    <td>&nbsp;</td>
    <td><div align="LEFT">PALM (KG) </div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TM_CPO_ACTUAL)+cdbl(v_TM_KER_ACTUAL),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TM_CPO_BUDGET)+cdbl(v_TM_KER_BUDGET),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TY_CPO_ACTUAL)+cdbl(v_TY_KER_ACTUAL),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_TY_CPO_BUDGET)+cdbl(v_TY_KER_BUDGET),2)%></div></td>
    <td colspan="3"><div align="center"><%=RubahAngka(cdbl(v_CPO_ANNUAL)+cdbl(v_KER_ANNUAL),2)%></div></td>
  </tr>
  <tr bgcolor="#F7F7F7">
    <td rowspan="2">&nbsp;</td>
    <td rowspan="2"><div align="center"></div></td>
    <td width="31" rowspan="2"><div align="center">Total</div></td>
    <td colspan="2"><div align="center">Rp./Kg.</div></td>
    <td width="31" rowspan="2"><div align="center">Total</div></td>
    <td colspan="2"><div align="center">Rp./Kg.</div></td>
    <td width="31" rowspan="2"><div align="center">Total</div></td>
    <td colspan="2"><div align="center">Rp./Kg.</div></td>
    <td width="31" rowspan="2"><div align="center">Total</div></td>
    <td colspan="2"><div align="center">Rp./Kg.</div></td>
    <td width="31" rowspan="2"><div align="center">Total</div></td>
    <td colspan="2"><div align="center">Rp./Kg.</div></td>
  </tr>
  <tr bgcolor="#F7F7F7">
    <td width="30"><div align="center">Palm</div></td>
    <td width="28"><div align="center">TBS</div></td>
    <td width="30"><div align="center">Palm</div></td>
    <td width="28"><div align="center">TBS</div></td>
    <td width="30"><div align="center">Palm</div></td>
    <td width="28"><div align="center">TBS</div></td>
    <td width="30"><div align="center">Palm</div></td>
    <td width="28"><div align="center">TBS</div></td>
    <td width="30"><div align="center">Palm</div></td>
    <td width="28"><div align="center">TBS</div></td>
  </tr>
