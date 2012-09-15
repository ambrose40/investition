<Html>
 <%Set servcfg=Server.CreateObject("Scripting.FileSystemObject")%>
 <%b=Server.MapPath("\")%>
 <Head>
  <%If Request.Cookies("StyleInv")="" Then%>
   <%Set servFileStream=servcfg.OpenTextFile(b & "\Style.cfg")%>
   <%s=servFileStream.ReadLine%>
   <%servFileStream.Close%>
   <Link Rel="Stylesheet" Href="<%=s%>" Type="text/css">
  <%Else%>
   <%s=Request.Cookies("StyleInv")%>
   <Link Rel="Stylesheet" Href="<%=s%>" Type="text/css">
  <%End If%>
  <Meta http-equiv="Content-Type" content="text/Html; Charset=windows-1251">
 </Head>
 <Body Class="Report">
  <%If Request.Form("btn")="OK" Then%>
   <%ya=Request.Form("ye")%>
  <%Else%>
   <%ya=Request.QueryString("ye")%>
  <%End If%>
  <%If ya="" Then%>
   <%ya=Year(Date())%>
   <%mo=Month(Date())%>
   <%da=Day(Date())%>
   <%zz=mo-04%>
   <%If zz>=0 Then%>
    <%ya=Year(Date())%>
   <%Else%>
    <%ya=ya-1%>
   <%End If%>
  <%Else%>
   <%If Request.Form("btn")="OK" Then%>
    <%ya=Request.Form("ye")%>
   <%Else%>
    <%ya=Request.QueryString("ye")%>
   <%End If%>
  <%End If%>
  <%Set servFileStream=servcfg.OpenTextFile(b & "\server.cfg")%>
  <%s=servFileStream.ReadLine%>
  <%i=servFileStream.ReadLine%>
  <%p=servFileStream.ReadLine%>
  <%servFileStream.Close%>
  <%Set mdbo =  Server.CreateObject("ADODB.Connection")%>
  <%mdbo.ConnectionString="Driver={SQL Server};Server=" & s & ";Database=invest;Trusted_Connection=yes;"%>
  <%mdbo.Open ConnectionString%>
  <%Set mdbol1 = Server.CreateObject("ADODB.Command")%>
  <%Set mdborl1 = Server.CreateObject("ADODB.RecordSet")%>
  <%mdbol1.ActiveConnection = mdbo%>
  <%Set mdbo2 = Server.CreateObject("ADODB.Command")%>
  <%Set mdbor2 = Server.CreateObject("ADODB.RecordSet")%>
  <%mdbo2.ActiveConnection = mdbo%>
  <%Set mdbo2u = Server.CreateObject("ADODB.Command")%>
  <%Set mdbor2u = Server.CreateObject("ADODB.RecordSet")%>
  <%mdbo2u.ActiveConnection = mdbo%>
  <%Set mdbo3 = Server.CreateObject("ADODB.Command")%>
  <%Set mdbor3 = Server.CreateObject("ADODB.RecordSet")%>
  <%mdbo3.ActiveConnection = mdbo%>
  <%Set mdbo3u = Server.CreateObject("ADODB.Command")%>
  <%Set mdbor3u = Server.CreateObject("ADODB.RecordSet")%>
  <%mdbo3u.ActiveConnection = mdbo%>
  <%Set mdbo4 = Server.CreateObject("ADODB.Command")%>
  <%Set mdbor4 = Server.CreateObject("ADODB.RecordSet")%>
  <%mdbo4.ActiveConnection = mdbo%>
  <%Set mdbo5 = Server.CreateObject("ADODB.Command")%>
  <%Set mdbor5 = Server.CreateObject("ADODB.RecordSet")%>
  <%mdbo5.ActiveConnection = mdbo%>
  <%Set mdbo6 = Server.CreateObject("ADODB.Command")%>
  <%Set mdbor6 = Server.CreateObject("ADODB.RecordSet")%>
  <%mdbo6.ActiveConnection = mdbo%>
  <%Set mdbo6a = Server.CreateObject("ADODB.Command")%>
  <%Set mdbor6a = Server.CreateObject("ADODB.RecordSet")%>
  <%mdbo6a.ActiveConnection = mdbo%>
  <%Set mdbo7 = Server.CreateObject("ADODB.Command")%>
  <%Set mdbor7 = Server.CreateObject("ADODB.RecordSet")%>
  <%mdbo7.ActiveConnection = mdbo%>
      
  <Table bordercolor="0F0F0F" border="1"  Style="border-collapse: collapse">
   <tr>
    <th>Projekti nr</th>
    <th>nr</a>
    <th>Projekti nimetus</th>
    <th><%=ya%>&nbspm.a&nbspkava</th>
    <th><%=ya%>&nbspm.a&nbsptegelikkult tehtud t&ouml;&ouml;d</th>
    <th>Ettemaksed ja p&otilde;hivara laos seisuga 01.04.<%=ya%></th>
    <th>Ettemaksed ja p&otilde;hivara laos seisuga 01.07.<%=ya%></th>
    <th><%=ya%>&nbspm.a&nbspkokku investeeritud</th>
    <th><%=ya%>&nbspm.a&nbspkokku k&auml;iku antud</th>
    <th>Demontaa&#382;</th>
    <th>L&otilde;petamata ehitus seisuga 01.04.<%=ya%></th>
    <th>L&otilde;petamata ehitus seisuga 01.07.<%=ya%></th>
   </tr>
   <tr Class="Repnum">
    <%For nuu=1 to 12%>
     <td><%=nuu%></td>
    <%Next%>
   </tr>
   <%aa=0%><%ab=0%><%ac=0%>
   <%mdbol1.CommandText="SELECT DISTINCT Pid,ProjCode, OracleCode,PC, PRojName FROM inpl WHERE IDentIfier='C' AND Yearr='" & ya & "' AND SUBSTRING(PC,4,2)='00' ORDER BY ProjCode"%>
   <%mdborl1.Open mdbol1%>
   <%mdbo2.CommandText="SELECT ISNULL(ROUND((SUM(ISNULL(GP.DEBET,0))/1000),0),0) AS summi FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND ((Konto<>'18164' AND Konto<>'18151') OR ((Konto='18164' OR Konto='18151') AND SUBSTRING(PROJECT,4,3)='999') OR (Konto='18164' AND (PROJECT='EJE573' OR PROJECT='EJE516' OR PROJECT='EJB515' OR PROJECT='EJB583') AND (MES='504' OR MES='505' OR MES='506'))) AND m.Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND (KONTO<>'43350' AND SUBKONTO<>'4351') AND m.RenovBlock=0 AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND (SUBSTRING(GP.MES, 2, 2) < 07) AND m.OracleCode<>'EJB206' AND (m.IDentIfier = 'C')"%>
   <%mdbor2.Open mdbo2%>
   <%mdbo2u.CommandText="SELECT ISNULL(ROUND((SUM(ISNULL(GP.DEBET,0))/1000),0),0)-ISNULL(ROUND((SUM(ISNULL(GP.CREDIT,0))/1000),0),0) AS summi FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND m.Yearr='" & ya & "' AND m.RenovBlock=0 AND (konto NOT BETWEEN '18410' AND '18433') AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND (SUBSTRING(GP.MES, 2, 2) < 07) AND m.OracleCode='EJB206' AND (m.IDentIfier = 'C')"%>
   <%mdbor2u.Open mdbo2u%>
   <%mdbo3.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.CREDIT,0))/1000,0),0) AS summc FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE SUBSTRING(PROJECT,4,3)<>'999' AND (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') And m.Yearr='" & ya & "' AND m.RenovBlock=0 AND (konto NOT BETWEEN '18410' AND '18433') AND m.OracleCode<>'EJB206' AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND (SUBSTRING(GP.MES, 2, 2) < 07) AND (m.IDentIfier = 'C')"%>
   <%mdbor3.Open mdbo3%>
   <%mdbo3u.CommandText="SELECT ISNULL(ROUND((SUM(ISNULL(GP.DEBET,0))/1000),0),0) AS summc FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE SUBSTRING(PROJECT,4,3)='999' AND (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND m.Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND m.RenovBlock=0 AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND (SUBSTRING(GP.MES, 2, 2) < 07) AND m.OracleCode<>'EJB206' AND (m.IDentIfier = 'C')"%>
   <%mdbor3u.Open mdbo3u%>
   <%mdbo4.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0),0) AS summd FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE m.Yearr='" & ya & "' AND m.RenovBlock=0 AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND (SUBSTRING(GP.MES, 2, 2) < 07) AND (m.IDentIfier = 'C') AND KONTO='43350' AND SUBKONTO='4351'"%>
   <%mdbor4.Open mdbo4%>
   <%mdbo5.CommandText="SELECT DISTINCT ISNULL(SUM(ISNULL(SummaPlan,0)),0) AS SP FROM dbo.Delta WHERE yearr='" & ya & "' AND enn<>'00' AND RenovBlock=0"%>
   <%mdbor5.Open mdbo5%>
   <%mdbo6.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0),0)-ISNULL(ROUND(SUM(ISNULL(GP.CREDIT,0))/1000,0),0) AS summy FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE m.Yearr='" & ya-1 & "' AND (konto BETWEEN '18410' AND '18433') AND m.RenovBlock=0 AND (((SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya-1,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04)) OR (LEFT(MES,1)=" & Mid(ya,4,1) & " AND RIGHT(MES,2) <04))  AND (m.IDentIfier = 'C')"%>
   <%mdbor6.Open mdbo6%>
   <%mdbo6a.CommandText="SELECT DISTINCT ISNULL(SUM(ISNULL(Ettemaks,0)),0) AS EM,ISNULL(SUM(ISNULL(Saldo,0)),0) AS SD FROM dbo.ETTE WHERE yearr='" & ya & "' AND enn<>'00' AND RenovBlock=0"%>
   <%mdbor6a.Open mdbo6a%>
   <%mdbo7.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0),0)-ISNULL(ROUND(SUM(ISNULL(GP.CREDIT,0))/1000,0),0) AS summym FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE m.Yearr='" & ya & "' AND (konto BETWEEN '18410' AND '18433') AND m.RenovBlock=0 AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND (SUBSTRING(GP.MES, 2, 2) < 07) AND (m.IDentIfier = 'C')"%>
   <%mdbor7.Open mdbo7%>
   <tr Class="Whitetr">
    <td>EJB206</td>
    <td>01.01.02.00.</td>
    <td>IVESTEERINGUDKOKKUvar</td>
    <td><%=mdbor5("SP")%></td>
    <td>
     <%If mdbor2u.EOF=True Then%>
      <%a1=0%>
     <%Else%>
      <%a1=mdbor2u("Summi")%>
     <%End If%>
     <%If mdbor2.EOF=True Then%>
      <%a2=0%>
     <%Else%>
      <%a2=mdbor2("Summi")%>
     <%End If%>
     <%=CDbl(a2)+CDBL(a1)%>
    </td>
    <td>
     <%If mdbor6.EOF=True Then%>
      <%a6=0%>
     <%Else%>
      <%a6=mdbor6("Summy")%>
     <%End If%>
     <%If mdbor6a.EOF=True Then%>
      <%a6a=0%>
     <%Else%>
      <%a6a=mdbor6a("EM")%>
     <%End If%>
     <%If Request.Form("btn")="Kopeerimiseks" Then%>
      <%a0="a1b"%>
      <%=Request.Form(a0)%>
     <%Else%>
      <%a0="a1b"%>
      <input Type="Text" value="<%=CDBL(a6)+CDBL(a6a)%>" size="10" name="<%="a1b"%>" Style="font-family: Verdana; color: #FFFFFF; font-weight:700; background-color: #FFFFFF; border-width:0">
     <%End If%>
    </td>
    <td>
     <%If mdbor7.EOF=True Then%>
      <%a7=0%>
     <%Else%>
      <%a7=mdbor7("Summym")%>
     <%End If%>
     <%If Request.Form("btn")="Kopeerimiseks" Then%>
      <%a0="a1a"%>
      <%=Request.Form(a0)%>
     <%Else%>
      <%a0="a1a"%>
      <input Type="Text" value="<%=CDBL(a7)%>" size="10" name="<%="a1a"%>" Style="font-family: Verdana; color: #FFFFFF; font-weight:700; background-color: #FFFFFF; border-width:0">
     <%End If%>
    </td>
    <td>
     <%If mdbor2u.EOF=True Then%>
      <%a1=0%>
     <%Else%>
      <%a1=mdbor2u("Summi")%>
     <%End If%>
     <%If mdbor2.EOF=True Then%>
      <%a2=0%>
     <%Else%>
      <%a2=mdbor2("Summi")%>
     <%End If%>
     <%If mdbor6.EOF=True Then%>
      <%a6=0%>
     <%Else%>
      <%a6=mdbor6("Summy")%>
     <%End If%>
     <%If mdbor4.EOF=true Then%>
      <%a4=0%>
     <%Else%>
      <%a4=mdbor4("Summd")%>
     <%End If%>
     <%If mdbor7.EOF=True Then%>
      <%a7=0%>
     <%Else%>
      <%a7=mdbor7("Summym")%>
     <%End If%>
     <%If mdbor6a.EOF=True Then%>
      <%a3=0%>
     <%Else%>
      <%a3=mdbor6a("EM")%>
     <%End If%>
     <%=CDbl(a2)+CDBL(a1)-CDBL(a3)-CDBL(a6)+CDBL(a7)%>
    </td>
    <td>
     <%If mdbor3.EOF=FALSE Then%>
      <%a3=mdbor3("Summc")%>
     <%Else%>
      <%a3=0%>
     <%End If%>
     <%If mdbor3u.EOF=FALSE Then%>
      <%a4=mdbor3u("Summc")%>
     <%Else%>
      <%a4=0%>
     <%End If%>
     <%If mdbor2u.EOF=FALSE Then%>
      <%a2=mdbor2u("Summi")%>
     <%Else%>
      <%a2=0%>
     <%End If%>
     <%=cdbl(a3)+cdbl(a4)%>
    </td>
    <td>
     <%If mdbor4.EOF=False Then%>
      <%=mdbor4("Summd")%>
     <%Else%>
      0
     <%End If%>
    </td>
    <td>
     <%If mdbor6a.EOF=True Then%>
      <%a6=0%>
     <%Else%>
      <%a6=mdbor6a("SD")%>
     <%End If%>
     <%If Request.Form("btn")="Kopeerimiseks" Then%>
      <%a0="a1c"%>
      <%=Request.Form(a0)%>
     <%Else%>
      <%a0="a1c"%>
      <input Type="Text" value="<%=CDBL(a6)%>" size="10" name="<%="a1c"%>" Style="font-family: Verdana; color: #FFFFFF; font-weight:700; background-color: #FFFFFF; border-width:0">
     <%End If%>
    </td>
    <td>
     <%If mdbor6A.EOF=true Then%>
      <%a6=0%>
     <%Else%>
      <%a6=mdbor6a("SD")%>
     <%End If%>
     <%If mdbor3.EOF=true Then%>
      <%a3=0%>
     <%Else%>
      <%a3=mdbor3("Summc")%>
     <%End If%>
     <%If mdbor3u.EOF=true Then%>
      <%a9=0%>
     <%Else%>
      <%a9=mdbor3u("Summc")%>
     <%End If%>
     <%If mdbor2u.EOF=True Then%>
      <%a1=0%>
     <%Else%>
      <%a1=mdbor2u("Summi")%>
     <%End If%>
     <%If mdbor4.EOF=true Then%>
      <%a4=0%>
     <%Else%>
      <%a4=mdbor4("Summd")%>
     <%End If%>
     <%If mdbor2.EOF=True Then%>
      <%a2=0%>
     <%Else%>
      <%a2=mdbor2("Summi")%>
     <%End If%>
     <%a0="a1c"%>
     <%=CDbl(a6)-CDbl(a3)-CDbl(a9)+CDbl(a1)+CDbl(a2)%>
    </td>
   </tr>
   <%mdbor2.Close%>
   <%mdbor2u.Close%>
   <%mdbor3.Close%>
   <%mdbor3u.Close%>
   <%mdbor4.Close%>
   <%mdbor5.Close%>
   <%mdbor6.Close%>
   <%mdbor6a.Close%>
   <%mdbor7.Close%>
  </Table>
 </Body>
</Html>