<Html>
 <%Set servcfg=Server.CreateObject("Scripting.FileSystemObject")%>
 <Head>
  <%b= Server.MapPath("\")%>
  <%If Request.Cookies("StyleInv")="" Then%>
   <%Set servFileStream=servcfg.OpenTextFile(b & "\style.cfg")%>
   <%s=servFileStream.ReadLine%>
   <%servFileStream.Close%>
   <Link rel="stylesheet" Href="<%=s%>" Type="text/css">
  <%Else%>
   <%s=Request.Cookies("StyleInv")%>
   <Link rel="stylesheet" Href="<%=s%>" Type="text/css">
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
  <img Border="0" src="icons/report.ico" Style=float:Left><p Align="center"><a Href="Main.asp" Target="_top" Class="HeadLink"><%=ya%> majandusaasta investeeringute kava 3 kuu l&otilde;ikes</a>
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
  <%Set mdbol2 = Server.CreateObject("ADODB.Command")%>
  <%Set mdborl2 = Server.CreateObject("ADODB.RecordSet")%>
  <%mdbol2.ActiveConnection = mdbo%>
  <%Set mdbol3 = Server.CreateObject("ADODB.Command")%>
  <%Set mdborl3 = Server.CreateObject("ADODB.RecordSet")%>
  <%mdbol3.ActiveConnection = mdbo%>
  <%Set mdbo1 = Server.CreateObject("ADODB.Command")%>
  <%Set mdbor = Server.CreateObject("ADODB.RecordSet")%>
  <%mdbo1.ActiveConnection = mdbo%>
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
  <Form Method="POST" Action="Report6.asp?ye=<%=ya%>">
   <Input Type="Submit" name="btn" size="10" Value="Kopeerimiseks" Class="button">
   <Input Type="Submit" name="btn" size="10" Value="Parandamiseks"  Class="button">
   <Table Border="1" width="100%">
    <tr bgcolor="AAAAAA">
     <th>Projekti nr</th>
     <th>nr</th>
     <th>Projekti nimetus</th>
     <th><%=ya%>&nbspm.a&nbspkava</th>
     <th><%=ya%>&nbspm.a&nbsptegelikkult tehtud t&ouml;&ouml;d</th>
     <th>Ettemaksed ja p&otilde;hivara laos seisuga 01.04.<%=ya%></th>
     <th>Ettemaksed ja p&otilde;hivara laos seisuga 01.10.<%=ya%></th>
     <th><%=ya%>&nbspm.a&nbspkokku investeeritud</th>
     <th><%=ya%>&nbspm.a&nbspkokku k&auml;iku antud</th>
     <th>Demontaa&#382;</th>
     <th>L&otilde;petamata ehitus seisuga 01.04.<%=ya%></th>
     <th>L&otilde;petamata ehitus seisuga 01.10.<%=ya%></th>
    </tr>
    <tr Class="RepNum">
     <%For nuu=1 to 12%>
      <td><%=nuu%></td>
     <%Next%>
    </tr>
    <%aa=0%><%ab=0%><%ac=0%>
    <%mdbol1.CommandText="SELECT DISTINCT Pid,ProjCode, OracleCode,PC, PRojName FROM inpl WHERE IDentIfier='C' AND Yearr='" & ya & "' AND SUBSTRING(PC,4,2)='00' ORDER BY ProjCode"%>
    <%mdborl1.Open mdbol1%>
    <%mdbo3u.CommandText="SELECT ISNULL(ROUND((SUM(ISNULL(GP.DEBET,0))/1000),0),0) AS summc FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE SUBSTRING(PROJECT,4,3)='999' AND (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND m.Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND m.RenovBlock=0 AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND (SUBSTRING(GP.MES, 2, 2) < 10) AND m.OracleCode<>'EJB206' AND (m.IDentIfier = 'C')"%>
    <%mdbor3u.Open mdbo3u%>
    <%mdbo2.CommandText="SELECT ISNULL(ROUND((SUM(ISNULL(GP.DEBET,0))/1000),0),0) AS summi FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND ((Konto<>'18164' AND Konto<>'18151') OR ((Konto='18164' OR Konto='18151') AND SUBSTRING(PROJECT,4,3)='999') OR (Konto='18164' AND (PROJECT='EJE573' OR PROJECT='EJE516' OR PROJECT='EJB515' OR PROJECT='EJB583') AND (MES='504' OR MES='505' OR MES='506'))) AND m.Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND (KONTO<>'43350' AND SUBKONTO<>'4351') AND m.RenovBlock=0 AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND (SUBSTRING(GP.MES, 2, 2) < 10) AND m.OracleCode<>'EJB206' AND (m.IDentIfier = 'C')"%>
    <%mdbor2.Open mdbo2%>
    <%mdbo2u.CommandText="SELECT ISNULL(ROUND((SUM(ISNULL(GP.DEBET,0))/1000),0),0) AS summi FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND m.Yearr='" & ya & "' AND m.RenovBlock=0 AND (konto NOT BETWEEN '18410' AND '18433') AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND (SUBSTRING(GP.MES, 2, 2) < 10) AND m.OracleCode='EJB206' AND (m.IDentIfier = 'C')"%>
    <%mdbor2u.Open mdbo2u%>
    <%mdbo3.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.CREDIT,0))/1000,0),0) AS summc FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE SUBSTRING(PROJECT,4,3)<>'999' AND (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') And m.Yearr='" & ya & "' AND m.RenovBlock=0 AND (konto NOT BETWEEN '18410' AND '18433') AND m.OracleCode<>'EJB206' AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND (SUBSTRING(GP.MES, 2, 2) < 10) AND (m.IDentIfier = 'C')"%>
    <%mdbor3.Open mdbo3%>
    <%mdbo4.CommandText="SELECT ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0) AS summd FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE m.Yearr='" & ya & "' AND m.RenovBlock=0 AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND (SUBSTRING(GP.MES, 2, 2) < 10) AND (m.IDentIfier = 'C') AND KONTO='43350' AND SUBKONTO='4351'"%>
    <%mdbor4.Open mdbo4%>
    <%mdbo5.CommandText="SELECT DISTINCT SUM(SummaPlan) AS SP FROM dbo.Delta WHERE yearr='" & ya & "' AND enn<>'00' AND RenovBlock=0"%>
    <%mdbor5.Open mdbo5%>
    <%mdbo6.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0),0)-ISNULL(ROUND(SUM(ISNULL(GP.CREDIT,0))/1000,0),0) AS summy FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE m.Yearr='" & ya-1 & "' AND (konto BETWEEN '18410' AND '18433') AND m.RenovBlock=0 AND (((SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya-1,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04)) OR (LEFT(MES,1)=" & Mid(ya,4,1) & " AND RIGHT(MES,2) <04)) AND (m.IDentIfier = 'C')"%>
    <%mdbor6.Open mdbo6%>
    <%mdbo6a.CommandText="SELECT DISTINCT SUM(ISNULL(Ettemaks,0)) AS EM,SUM(ISNULL(Saldo,0)) AS SD FROM dbo.ETTE WHERE yearr='" & ya & "' AND enn<>'00' AND RenovBlock=0"%>
    <%mdbor6a.Open mdbo6a%>
    <%mdbo7.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0),0)-ISNULL(ROUND(SUM(ISNULL(GP.CREDIT,0))/1000,0),0) AS summym FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE m.Yearr='" & ya & "' AND (konto BETWEEN '18410' AND '18433') AND m.RenovBlock=0 AND GP.MES <" & Mid(ya,4,1) & "10 AND (m.IDentIfier = 'C')"%>
    <%mdbor7.Open mdbo7%>
    <tr Class="boldProjGrup">
     <td colspan=3>IVESTEERINGUD KOKKU v&auml;lja arvatud plokkide renoveerimine</td>
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
       <input Type="Text" value="<%=CDBL(a6)%>" size="10" name="<%="a1b"%>" Class="boldProjGrup">
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
       <input Type="Text" value="<%=CDBL(a7)%>" size="10" name="<%="a1a"%>"  Class="boldProjGrup">
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
      <%=CDbl(a2)+CDBL(a1)-CDBL(a6)+CDBL(a7)%>
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
       <input Type="Text" value="<%=CDBL(a6)%>" size="10" name="<%="a1c"%>"  Class="boldProjGrup">
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
      <%=CLNG(CDbl(a6)-CDbl(a3)-CDbl(a9)+CDbl(a1)+CDbl(a2))%>
     </td>
    </tr>
    <%mdbor2.Close%><%mdbor2u.Close%><%mdbor3.Close%><%mdbor3u.Close%><%mdbor4.Close%><%mdbor5.Close%><%mdbor6.Close%><%mdbor6a.Close%><%mdbor7.Close%>
    
    <%mdbo3.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.CREDIT,0))/1000,0),0) AS summc FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE SUBSTRING(PROJECT,4,3)<>'999' AND (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700')  And m.Yearr='" & ya & "' AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (konto NOT BETWEEN '18410' AND '18433') AND m.OracleCode<>'EJB206' AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND (SUBSTRING(GP.MES, 2, 2) < 10) AND (m.IDentIfier = 'C')"%>
    <%mdbor3.Open mdbo3%>
    <%mdbo3u.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0),0) AS summc FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE SUBSTRING(PROJECT,4,3)='999' AND (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') and  m.Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND (SUBSTRING(GP.MES, 2, 2) < 10) AND m.OracleCode<>'EJB206' AND (m.IDentIfier = 'C')"%>
    <%mdbor3u.Open mdbo3u%>
    <%mdbo2.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0),0) AS summi FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND ((Konto<>'18164' AND Konto<>'18151') OR ((Konto='18164' OR Konto='18151') AND SUBSTRING(PROJECT,4,3)='999') OR (Konto='18164' AND (PROJECT='EJE573' OR PROJECT='EJE516' OR PROJECT='EJB515' OR PROJECT='EJB583') AND (MES='504' OR MES='505' OR MES='506'))) AND  m.Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND (KONTO<>'43350' AND SUBKONTO<>'4351') AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND (SUBSTRING(GP.MES, 2, 2) < 10) AND m.OracleCode<>'EJB206' AND (m.IDentIfier = 'C')"%>
    <%mdbor2.Open mdbo2%>
    <%mdbo2u.CommandText="SELECT ISNULL(ROUND((SUM(ISNULL(GP.DEBET,0))/1000),0),0) AS summi FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') and m.Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND (SUBSTRING(GP.MES, 2, 2) < 10) AND m.OracleCode='EJB206' AND (m.IDentIfier = 'C')"%>
    <%mdbor2u.Open mdbo2u%>
    <%mdbo4.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0),0) AS summd FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE m.Yearr='" & ya & "' AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND (SUBSTRING(GP.MES, 2, 2) < 10) AND (m.IDentIfier = 'C') AND KONTO='43350' AND SUBKONTO='4351'"%>
    <%mdbor4.Open mdbo4%>
    <%mdbo5.CommandText="SELECT DISTINCT ISNULL(SUM(ISNULL(SummaPlan,0)),0) AS SP FROM dbo.Delta WHERE yearr='" & ya & "' AND enn<>'00'"%>
    <%mdbor5.Open mdbo5%>
    <%mdbo6.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0),0)-ISNULL(ROUND(SUM(ISNULL(GP.CREDIT,0))/1000,0),0) AS summy FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE m.Yearr='" & ya-1 & "' AND (konto BETWEEN '18410' AND '18433') AND (((SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya-1,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04)) OR (LEFT(MES,1)=" & Mid(ya,4,1) & " AND RIGHT(MES,2) <04)) AND (m.IDentIfier = 'C')"%>
    <%mdbor6.Open mdbo6%>
    <%mdbo6a.CommandText="SELECT DISTINCT ISNULL(SUM(ISNULL(Ettemaks,0)),0) AS EM, ISNULL(SUM(ISNULL(Saldo,0)),0) AS SD FROM dbo.ETTE WHERE yearr='" & ya & "' AND enn<>'00'"%>
    <%mdbor6a.Open mdbo6a%>
    <%mdbo7.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0),0)-ISNULL(ROUND(SUM(ISNULL(GP.CREDIT,0))/1000,0),0) AS summym FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE m.Yearr='" & ya & "' AND (konto BETWEEN '18410' AND '18433') AND GP.MES <" & Mid(ya,4,1) & "10 AND (m.IDentIfier = 'C')"%>
    <%mdbor7.Open mdbo7%>
    <tr Class="boldProjGrup">
     <td colspan=3>IVESTEERINGUD KOKKU koos plokkide renoveerimisega</td>
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
       <%a0="ab"%>
       <%=Request.Form(a0)%>
      <%Else%>
       <%a0="ab"%>
       <input Type="Text" value="<%=CDBL(a6)%>" size="10" name="<%="ab"%>"  Class="boldProjGrup">
      <%End If%>
     </td>
     <td>
      <%If mdbor7.EOF=True Then%>
       <%a7=0%>
      <%Else%>
       <%a7=mdbor7("Summym")%>
      <%End If%>
      <%If Request.Form("btn")="Kopeerimiseks" Then%>
       <%a0="aa"%>
       <%=Request.Form(a0)%>
      <%Else%>
       <%a0="aa"%>
       <input Type="Text" value="<%=CDBL(a7)%>" size="10" name="<%="aa"%>"  Class="boldProjGrup">
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
      <%If mdbor4.EOF=true Then%>
       <%a4=0%>
      <%Else%>
       <%a4=mdbor4("Summd")%>
      <%End If%>
      <%=CDbl(a2)+CDBL(a1)-CDBL(a6)+CDBL(a7)%>
     </td>
     <td>
      <%If mdbor3.EOF=FALSE Then%>
       <%a3=mdbor3("Summc")%>
      <%Else%>
       <%a3=0%>
      <%End If%>
      <%If mdbor3U.EOF=FALSE Then%>
       <%a4=mdbor3U("Summc")%>
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
       <%=CLNG(mdbor4("Summd"))%>
      <%Else%>
       0
      <%End If%>
     </td>
     <td>
      <%If mdbor6a.EOF=True Then%>
       <%a3=0%>
      <%Else%>
       <%a3=mdbor6a("SD")%>
      <%End If%>
      <%If Request.Form("btn")="Kopeerimiseks" Then%>
       <%a0="ac"%>
       <%=Request.Form(a0)%>
      <%Else%>
       <%a0="ac"%>
       <input Type="Text" value="<%=CDBL(a3)%>" size="10" name="<%="ac"%>"  Class="boldProjGrup">
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
      <%=CLNG(CDbl(a6)-CDbl(a3)-CDbl(a9)+CDbl(a1)+CDbl(a2))%>
     </td>
    </tr>
    <%Do Until mdborl1.EOF%>
     <%mdbor2.Close%><%mdbor2u.Close%><%mdbor3.Close%><%mdbor3u.Close%><%mdbor4.Close%><%mdbor5.Close%><%mdbor6.Close%><%mdbor6a.Close%><%mdbor7.Close%>
      
     <%mdbo3u.CommandText="SELECT ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0) AS summc, SUBSTRING(m.ProjCode, 1, 2) AS be FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE SUBSTRING(Project,4,3)='999' and (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700')  and  m.Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND SUBSTRING(m.ProjCode,1,2)='" & MId(mdborl1("PC"),1,2) & "' AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND (SUBSTRING(GP.MES, 2, 2) < 10) AND m.OracleCode<>'EJB206' AND (m.IDentIfier = 'C') GROUP BY SUBSTRING(m.ProjCode, 1, 2)"%>
     <%mdbor3u.Open mdbo3u%>
     <%mdbo2.CommandText="SELECT ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0) AS summi, SUBSTRING(m.ProjCode, 1, 2) AS be FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND ((Konto<>'18164' AND Konto<>'18151') OR ((Konto='18164' OR Konto='18151') AND SUBSTRING(PROJECT,4,3)='999') OR (Konto='18164' AND (PROJECT='EJE573' OR PROJECT='EJE516' OR PROJECT='EJB515' OR PROJECT='EJB583') AND (MES='504' OR MES='505' OR MES='506'))) AND  m.Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND (KONTO<>'43350' AND SUBKONTO<>'4351') AND SUBSTRING(m.ProjCode,1,2)='" & MId(mdborl1("PC"),1,2) & "' AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND (SUBSTRING(GP.MES, 2, 2) < 10) AND m.OracleCode<>'EJB206' AND (m.IDentIfier = 'C') GROUP BY SUBSTRING(m.ProjCode, 1, 2)"%>
     <%mdbor2.Open mdbo2%>    
     <%mdbo2u.CommandText="SELECT ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0) AS summi, SUBSTRING(m.ProjCode, 1, 2) AS be FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE  (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700')  and  m.Yearr='" & ya & "' AND SUBSTRING(m.ProjCode,1,2)='" & MId(mdborl1("PC"),1,2) & "' AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND (SUBSTRING(GP.MES, 2, 2) < 10) AND m.OracleCode='EJB206' AND (m.IDentIfier = 'C') GROUP BY SUBSTRING(m.ProjCode, 1, 2)"%>
     <%mdbor2u.Open mdbo2u%>
     <%mdbo3.CommandText="SELECT ROUND(SUM(ISNULL(GP.CREDIT,0))/1000,0) AS summc, SUBSTRING(m.ProjCode, 1, 2) AS be FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE SUBSTRING(Project,4,3)<>'999' and (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700')  And m.Yearr='" & ya & "'  AND (konto NOT BETWEEN '18410' AND '18433') AND SUBSTRING(m.ProjCode,1,2)='" & MId(mdborl1("PC"),1,2) & "' AND m.OracleCode<>'EJB206' AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND (SUBSTRING(GP.MES, 2, 2) < 10) AND (m.IDentIfier = 'C') GROUP BY SUBSTRING(m.ProjCode, 1, 2)"%>
     <%mdbor3.Open mdbo3%>
     <%mdbo4.CommandText="SELECT ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0) AS summd, SUBSTRING(m.ProjCode, 1, 2) AS be FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE m.Yearr='" & ya & "' AND SUBSTRING(m.ProjCode,1,2)='" & MId(mdborl1("PC"),1,2) & "' AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND (SUBSTRING(GP.MES, 2, 2) < 10) AND (m.IDentIfier = 'C') AND KONTO='43350' AND SUBKONTO='4351' GROUP BY SUBSTRING(m.ProjCode, 1, 2)"%>
     <%mdbor4.Open mdbo4%>
     <%mdbo5.CommandText="SELECT DISTINCT SUM(ISNULL(SummaPlan,0)) AS SP, SUM(ISNULL(SummaContract,0)) AS SC, be FROM dbo.Delta WHERE be='" & MID(mdborl1("PC"),1,2) & "' AND yearr='" & ya & "' AND enn<>'00' GROUP BY be"%>
     <%mdbor5.Open mdbo5%>
     <%mdbo6.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0),0)-ISNULL(ROUND(SUM(ISNULL(GP.CREDIT,0))/1000,0),0) AS summy FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE m.Yearr='" & ya-1 & "' AND SUBSTRING(m.ProjCode,1,2)='" & MId(mdborl1("PC"),1,2) & "' AND (konto BETWEEN '18410' AND '18433') AND (((SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya-1,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04)) OR (LEFT(MES,1)=" & Mid(ya,4,1) & " AND RIGHT(MES,2) <04)) AND (m.IDentIfier = 'C') GROUP BY SUBSTRING(m.ProjCode, 1, 2)"%>
     <%mdbor6.Open mdbo6%>
     <%mdbo6a.CommandText="SELECT DISTINCT SUM(ISNULL(Ettemaks,0)) AS EM,SUM(ISNULL(Saldo,0)) AS SD FROM dbo.ETTE WHERE be='" & MID(mdborl1("PC"),1,2) & "' AND yearr='" & ya & "' AND enn<>'00' GROUP BY be"%>
     <%mdbor6a.Open mdbo6a%>
     <%mdbo7.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0),0)-ISNULL(ROUND(SUM(ISNULL(GP.CREDIT,0))/1000,0),0) AS summym FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE m.Yearr='" & ya & "' AND SUBSTRING(m.ProjCode,1,2)='" & MId(mdborl1("PC"),1,2) & "' AND (konto BETWEEN '18410' AND '18433') AND GP.MES <" & Mid(ya,4,1) & "10 AND (m.IDentIfier = 'C') GROUP BY SUBSTRING(m.ProjCode, 1, 2)"%>
     <%mdbor7.Open mdbo7%>
     <tr Class="ProjGrup">
      <td>
       <%a=MID(mdborl1("PC"),1,3)%>
       <%=REPLACE(a, "0", "")%>
      </td>
      <td colspan=2><%=mdborl1("ProjName")%></td>
      <td>
       <%If mdbor5("be")=MId(mdborl1("PC"),1,2) Then%>
        <%=mdbor5("SP")%>
       <%Else%>
        0
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
        <%a0="ab" & mdborl1("Pid")%>
        <%=Request.Form(a0)%>
       <%Else%>
        <%a0="ab" & mdborl1("Pid")%>
        <input Type="Text" value="<%=CDBL(a6)%>" size="10" name="<%="ab" & mdborl1("Pid")%>"  Class="ProjGrup">
       <%End If%>
      </td>
      <td>
       <%If mdbor7.EOF=True Then%>
        <%a7=0%>
       <%Else%>
        <%a7=mdbor7("Summym")%>
       <%End If%>
       <%If Request.Form("btn")="Kopeerimiseks" Then%>
        <%a0="aa" & mdborl1("Pid")%>
        <%=Request.Form(a0)%>
       <%Else%>
        <%a0="aa" & mdborl1("Pid")%>
        <input Type="Text" value="<%=CDBL(a7)%>" size="10" name="<%="aa" & mdborl1("Pid")%>"  Class="ProjGrup">
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
       <%If mdbor4.EOF=true Then%>
        <%a4=0%>
       <%Else%>
        <%a4=mdbor4("Summd")%>
       <%End If%>
       <%=CDbl(a2)+CDBL(a1)-CDBL(a6)+CDBL(a7)%>
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
        <%=CLNG(mdbor4("Summd"))%>
       <%Else%>
        0
       <%End If%>
      </td>
      <td>
       <%If mdbor6a.EOF=True Then%>
        <%a3=0%>
       <%Else%>
        <%a3=mdbor6a("SD")%>
       <%End If%>
       <%If Request.Form("btn")="Kopeerimiseks" Then%>
        <%a0="ac" & mdborl1("Pid")%>
        <%=Request.Form(a0)%>
       <%Else%>
        <%a0="ac" & mdborl1("Pid")%>
        <input Type="Text" value="<%=CDBL(a3)%>" size="10" name="<%="ac" & mdborl1("Pid")%>"  Class="ProjGrup">
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
       <%If mdbor4.EOF=true Then%>
        <%a4=0%>
       <%Else%>
        <%a4=mdbor4("Summd")%>
       <%End If%>
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
       <%a0="a1c"%>
       <%=CLNG(CDbl(a6)-CDbl(a3)-CDbl(a9)+CDbl(a1)+CDbl(a2))%>
      </td>
     </tr>
     <%mdbol2.CommandText="SELECT DISTINCT Pid,ProjCode,ProjName,OracleCode,PC FROM inpl WHERE IDentIfier='C' AND Yearr=" & ya & " AND SUBSTRING(PC,4,2)<>'00' AND SUBSTRING(PC,7,2)='00' AND  SUBSTRING(PC,1,2)='" & MID(mdborl1("PC"),1,2) & "' ORDER BY ProjCode"%>
     <%mdborl2.Open mdbol2%>
     <%Do Until mdborl2.EOF%>
      <%mdbor2.Close%><%mdbor2u.Close%><%mdbor3.Close%><%mdbor3u.Close%><%mdbor4.Close%><%mdbor5.Close%><%mdbor6.Close%><%mdbor6a.Close%><%mdbor7.Close%>
         
      <%mdbo2.CommandText="SELECT ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0) AS summi, SUBSTRING(m.ProjCode, 1, 2), SUBSTRING(m.ProjCode, 4, 2) AS be FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND ((Konto<>'18164' AND Konto<>'18151') OR ((Konto='18164' OR Konto='18151') AND SUBSTRING(PROJECT,4,3)='999') OR (Konto='18164' AND (PROJECT='EJE573' OR PROJECT='EJE516' OR PROJECT='EJB515' OR PROJECT='EJB583') AND (MES='504' OR MES='505' OR MES='506'))) AND (KONTO<>'43350' AND SUBKONTO<>'4351') AND  m.Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND SUBSTRING(m.ProjCode,1,5)='" & MId(mdborl2("PC"),1,5) & "' AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND (SUBSTRING(GP.MES, 2, 2) < 10) AND m.OracleCode<>'EJB206' AND (m.IDentIfier = 'C') GROUP BY SUBSTRING(m.ProjCode, 1, 2),SUBSTRING(m.ProjCode, 4, 2)"%>
      <%mdbor2.Open mdbo2%>
      <%mdbo2u.CommandText="SELECT ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0) AS summi, SUBSTRING(m.ProjCode, 1, 2), SUBSTRING(m.ProjCode, 4, 2) AS be FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE  (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700')   and m.Yearr='" & ya & "' AND SUBSTRING(m.ProjCode,1,5)='" & MId(mdborl2("PC"),1,5) & "' AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND (SUBSTRING(GP.MES, 2, 2) < 10) AND m.OracleCode='EJB206' AND (m.IDentIfier = 'C') GROUP BY SUBSTRING(m.ProjCode, 1, 2),SUBSTRING(m.ProjCode, 4, 2)"%>
      <%mdbor2u.Open mdbo2u%>     
      <%mdbo3.CommandText="SELECT ROUND(SUM(ISNULL(GP.CREDIT,0))/1000,0) AS summc, SUBSTRING(m.ProjCode, 1, 2), SUBSTRING(m.ProjCode, 4, 2) AS be FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE SUBSTRING(PROJECT,4,3)<>'999' and (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700')  And m.Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND SUBSTRING(m.ProjCode,1,5)='" & MId(mdborl2("PC"),1,5) & "' AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND m.OracleCode<>'EJB206' AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND (SUBSTRING(GP.MES, 2, 2) < 10) AND (m.IDentIfier = 'C') GROUP BY SUBSTRING(m.ProjCode, 1, 2),SUBSTRING(m.ProjCode, 4, 2)"%>
      <%mdbor3.Open mdbo3%>
      <%mdbo3u.CommandText="SELECT ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0) AS summc, SUBSTRING(m.ProjCode, 1, 2), SUBSTRING(m.ProjCode, 4, 2) AS be FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE SUBSTRING(PROJECT,4,3)='999' and (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700')  and  m.Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND SUBSTRING(m.ProjCode,1,5)='" & MId(mdborl2("PC"),1,5) & "' AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND (SUBSTRING(GP.MES, 2, 2) < 10) AND m.OracleCode<>'EJB206' AND (m.IDentIfier = 'C') GROUP BY SUBSTRING(m.ProjCode, 1, 2),SUBSTRING(m.ProjCode, 4, 2)"%>
      <%mdbor3u.Open mdbo3u%>     
      <%mdbo4.CommandText="SELECT ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0) AS summd, SUBSTRING(m.ProjCode, 1, 2), SUBSTRING(m.ProjCode, 4, 2) AS be FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE m.Yearr='" & ya & "' AND SUBSTRING(m.ProjCode,1,5)='" & MId(mdborl2("PC"),1,5) & "' AND  (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND (SUBSTRING(GP.MES, 2, 2) < 10) AND (m.IDentIfier = 'C') AND KONTO='43350' AND SUBKONTO='4351' GROUP BY SUBSTRING(m.ProjCode, 1, 2),SUBSTRING(m.ProjCode, 4, 2)"%>
      <%mdbor4.Open mdbo4%>
      <%mdbo5.CommandText="SELECT DISTINCT SUM(ISNULL(SummaPlan,0)) AS SP, be, mi FROM dbo.Delta WHERE yearr='" & ya & "' AND mi = '" & Mid(mdborl2("PC"),4,2) & "' AND enn<>'00' AND be='" & Mid(mdborl2("PC"),1,2) & "' GROUP BY be, mi"%>
      <%mdbor5.Open mdbo5%>
      <%mdbo6.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0),0)-ISNULL(ROUND(SUM(ISNULL(GP.CREDIT,0))/1000,0),0) AS summy FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE m.Yearr='" & ya-1 & "' AND SUBSTRING(m.ProjCode,1,5)='" & MId(mdborl2("PC"),1,5) & "' AND (konto BETWEEN '18410' AND '18433') AND (((SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya-1,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04)) OR (LEFT(MES,1)=" & Mid(ya,4,1) & " AND RIGHT(MES,2) <04)) AND (m.IDentIfier = 'C') GROUP BY SUBSTRING(m.ProjCode, 1, 2),SUBSTRING(m.ProjCode, 4, 2)"%>
      <%mdbor6.Open mdbo6%>
      <%mdbo6a.CommandText="SELECT DISTINCT SUM(ISNULL(Ettemaks,0)) AS EM,SUM(ISNULL(Saldo,0)) AS SD FROM dbo.ETTE WHERE be='" & MID(mdborl1("PC"),1,2) & "' AND mi = '" & Mid(mdborl2("PC"),4,2) & "' AND yearr='" & ya & "' AND enn<>'00' GROUP BY be, mi"%>
      <%mdbor6a.Open mdbo6a%>
      <%mdbo7.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0),0)-ISNULL(ROUND(SUM(ISNULL(GP.CREDIT,0))/1000,0),0) AS summym FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE m.Yearr='" & ya & "' AND SUBSTRING(m.ProjCode,1,5)='" & MId(mdborl2("PC"),1,5) & "' AND (konto BETWEEN '18410' AND '18433') AND GP.MES <" & Mid(ya,4,1) & "10 AND (m.IDentIfier = 'C') GROUP BY SUBSTRING(m.ProjCode, 1, 2),SUBSTRING(m.ProjCode, 4, 2)"%>
      <%mdbor7.Open mdbo7%>
      <tr Class="ProjGrup">
       <td>
        <%a=MID(mdborl2("PC"),1,6)%>
        <%=REPLACE(a, "0", "")%>
       </td>
       <td colspan=2>
        <%=mdborl2("ProjName")%>
       </td>
       <td>
        <%If mdbor5("mi")=MID(mdborl2("PC"),4,2) Then%>
         <%=mdbor5("SP")%>
        <%Else%>
         0
        <%End If%>
       </td>
       <td>
        <%a0="aa" & mdborl2("Pid")%>
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
         <%a0="ab" & mdborl2("Pid")%>
         <%=Request.Form(a0)%>
        <%Else%>
         <%a0="ab" & mdborl2("Pid")%>
         <input Type="Text" value="<%=CDBL(a6)%>" size="10" name="<%="ab" & mdborl2("Pid")%>"  Class="ProjGrup">
        <%End If%>
       </td>
       <td>
        <%If mdbor7.EOF=True Then%>
         <%a7=0%>
        <%Else%>
         <%a7=mdbor7("Summym")%>
        <%End If%>
        <%If Request.Form("btn")="Kopeerimiseks" Then%>
         <%a0="aa" & mdborl2("Pid")%>
         <%=Request.Form(a0)%>
        <%Else%>
         <%a0="aa" & mdborl2("Pid")%>
         <input Type="Text" value="<%=CDBL(a7)%>" size="10" name="<%="aa" & mdborl2("Pid")%>"  Class="ProjGrup">
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
        <%If mdbor4.EOF=true Then%>
         <%a4=0%>
        <%Else%>
         <%a4=mdbor4("Summd")%>
        <%End If%>
        <%=CDbl(a2)+CDBL(a1)-CDBL(a6)+CDBL(a7)%>
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
         <%a3=0%>
        <%Else%>
         <%a3=mdbor6a("SD")%>
        <%End If%>
        <%If Request.Form("btn")="Kopeerimiseks" Then%>
         <%a0="ac" & mdborl2("Pid")%>
         <%=Request.Form(a0)%>
        <%Else%>
         <%a0="ac" & mdborl2("Pid")%>
         <input Type="Text" value="<%=CDBL(a3)%>" size="10" name="<%="ac" & mdborl2("Pid")%>"  Class="ProjGrup">
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
        <%If mdbor4.EOF=true Then%>
         <%a4=0%>
        <%Else%>
         <%a4=mdbor4("Summd")%>
        <%End If%>
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
        <%a0="a1c"%>
        <%=CDbl(a6)-CDbl(a3)-CDbl(a9)+CDbl(a1)+CDbl(a2)%>
       </td>
      </tr>
      <%mdbol3.CommandText="SELECT DISTINCT Enterprise,Edescr FROM inpl WHERE IDentIfier='C' AND Yearr=" & ya & " AND SUBSTRING(PC,4,2)='" & MID(mdborl2("PC"),4,2) & "' AND  SUBSTRING(PC,1,2)='" & MID(mdborl2("PC"),1,2) & "'"%>
      <%mdborl3.Open mdbol3%>
           
      <%Do Until mdborl3.EOF%>
       <%mdbor2.Close%><%mdbor2u.Close%><%mdbor3.Close%><%mdbor3u.Close%><%mdbor4.Close%><%mdbor5.Close%><%mdbor6.Close%><%mdbor6a.Close%><%mdbor7.Close%>
       
       <%mdbo2.CommandText="SELECT ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0) AS summi, SUBSTRING(m.ProjCode, 1, 2), SUBSTRING(m.ProjCode, 4, 2) AS be FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND ((Konto<>'18164' AND Konto<>'18151') OR ((Konto='18164' OR Konto='18151') AND  SUBSTRING(PROJECT,4,3)='999') OR (Konto='18164' AND (PROJECT='EJE573' OR PROJECT='EJE516' OR PROJECT='EJB515' OR PROJECT='EJB583') AND (MES='504' OR MES='505' OR MES='506'))) AND (KONTO<>'43350' AND SUBKONTO<>'4351') AND m.Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND SUBSTRING(m.ProjCode,1,5)='" & MId(mdborl2("PC"),1,5) & "' AND m.Enterprise='" & Mdborl3("Enterprise") & "' AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND RIGHT(MES,2) <10 AND m.OracleCode<>'EJB206' AND (m.IDentIfier = 'C') GROUP BY SUBSTRING(m.ProjCode, 1, 2),SUBSTRING(m.ProjCode, 4, 2), m.Enterprise"%>
       <%mdbor2.Open mdbo2%>
       <%mdbo2u.CommandText="SELECT ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0) AS summi, SUBSTRING(m.ProjCode, 1, 2), SUBSTRING(m.ProjCode, 4, 2) AS be FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE  (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700')   and m.Yearr='" & ya & "' AND SUBSTRING(m.ProjCode,1,5)='" & MId(mdborl2("PC"),1,5) & "' AND m.Enterprise='" & Mdborl3("Enterprise") & "' AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND RIGHT(MES,2) <10 AND m.OracleCode='EJB206' AND (m.IDentIfier = 'C') GROUP BY SUBSTRING(m.ProjCode, 1, 2),SUBSTRING(m.ProjCode, 4, 2), m.Enterprise"%>
       <%mdbor2u.Open mdbo2u%>      
       <%mdbo3.CommandText="SELECT ROUND(SUM(ISNULL(GP.CREDIT,0))/1000,0) AS summc, SUBSTRING(m.ProjCode, 1, 2), SUBSTRING(m.ProjCode, 4, 2), m.Enterprise AS be FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE SUBSTRING(Project,4,3)<>'999' and (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') and m.Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND SUBSTRING(m.ProjCode,1,5)='" & MId(mdborl2("PC"),1,5) & "' AND m.Enterprise='" & Mdborl3("Enterprise") & "' AND m.OracleCode<>'EJB206' AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND RIGHT(MES,2) <10 AND (m.IDentIfier = 'C') GROUP BY SUBSTRING(m.ProjCode, 1, 2),SUBSTRING(m.ProjCode, 4, 2), m.Enterprise"%>
       <%mdbor3.Open mdbo3%>
       <%mdbo3u.CommandText="SELECT ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0) AS summc, SUBSTRING(m.ProjCode, 1, 2), SUBSTRING(m.ProjCode, 4, 2) AS be FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE SUBSTRING(Project,4,3)='999' and (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700')  and  m.Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND SUBSTRING(m.ProjCode,1,5)='" & MId(mdborl2("PC"),1,5) & "' AND m.Enterprise='" & Mdborl3("Enterprise") & "' AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND RIGHT(MES,2) <10 AND m.OracleCode<>'EJB206' AND (m.IDentIfier = 'C') GROUP BY SUBSTRING(m.ProjCode, 1, 2),SUBSTRING(m.ProjCode, 4, 2), m.Enterprise"%>
       <%mdbor3u.Open mdbo3u%>     
       <%mdbo4.CommandText="SELECT ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0) AS summd, SUBSTRING(m.ProjCode, 1, 2), SUBSTRING(m.ProjCode, 4, 2), m.Enterprise AS be FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE m.Yearr='" & ya & "' AND SUBSTRING(m.ProjCode,1,5)='" & MId(mdborl2("PC"),1,5) & "' AND m.Enterprise='" & Mdborl3("Enterprise") & "' AND  (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND RIGHT(MES,2) <10 AND (m.IDentIfier = 'C') AND KONTO='43350' AND SUBKONTO='4351' GROUP BY SUBSTRING(m.ProjCode, 1, 2),SUBSTRING(m.ProjCode, 4, 2), m.Enterprise"%>
       <%mdbor4.Open mdbo4%>
       <%mdbo5.CommandText="SELECT DISTINCT SUM(SummaPlan) AS SP, be, mi, Enterprise FROM dbo.Delta WHERE yearr='" & ya & "' AND mi = '" & Mid(mdborl2("PC"),4,2) & "' AND Enterprise='" & Mdborl3("Enterprise") & "' AND enn<>'00' AND be='" & Mid(mdborl2("PC"),1,2) & "' GROUP BY be, mi,Enterprise"%>
       <%mdbor5.Open mdbo5%>
       <%mdbo6.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0),0)-ISNULL(ROUND(SUM(ISNULL(GP.CREDIT,0))/1000,0),0) AS summy FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE m.Yearr='" & ya-1 & "' AND SUBSTRING(m.ProjCode,1,5)='" & MId(mdborl2("PC"),1,5) & "' AND m.Enterprise='" & Mdborl3("Enterprise") & "'  AND (konto BETWEEN '18410' AND '18433') AND (((SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya-1,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04)) OR (LEFT(MES,1)=" & Mid(ya,4,1) & " AND RIGHT(MES,2) <04)) AND (m.IDentIfier = 'C') GROUP BY SUBSTRING(m.ProjCode, 1, 2),SUBSTRING(m.ProjCode, 4, 2), m.Enterprise"%>
       <%mdbor6.Open mdbo6%>
       <%mdbo6a.CommandText="SELECT DISTINCT SUM(ISNULL(Ettemaks,0)) AS EM,SUM(ISNULL(Saldo,0)) AS SD FROM dbo.ETTE WHERE be='" & MID(mdborl1("PC"),1,2) & "' AND mi = '" & Mid(mdborl2("PC"),4,2) & "' AND yearr='" & ya & "' AND Enterprise='" & Mdborl3("Enterprise") & "' AND enn<>'00' GROUP BY be, mi,Enterprise"%>
       <%mdbor6a.Open mdbo6a%>
       <%mdbo7.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0),0)-ISNULL(ROUND(SUM(ISNULL(GP.CREDIT,0))/1000,0),0) AS summym FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE m.Yearr='" & ya & "' AND SUBSTRING(m.ProjCode,1,5)='" & MId(mdborl2("PC"),1,5) & "' AND m.Enterprise='" & Mdborl3("Enterprise") & "'  AND (konto BETWEEN '18410' AND '18433') AND GP.MES <" & Mid(ya,4,1) & "10 AND (m.IDentIfier = 'C') GROUP BY SUBSTRING(m.ProjCode, 1, 2),SUBSTRING(m.ProjCode, 4, 2), m.Enterprise"%>
       <%mdbor7.Open mdbo7%>
       <tr Class="Enterp">
        <td colspan=3>
         <%=mdborl3("EDescr")%>
        </td>
        <td>
         <%If mdbor5.EOF=False Then%>
          <%=mdbor5("SP")%>
         <%Else%>
          0
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
          <%a0="ab" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>
          <%=Request.Form(a0)%>
         <%Else%>
          <%a0="ab" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>
          <input Type="Text" value="<%=CDBL(a6)%>" size="10" name="<%="ab" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>"  Class="Enterp">
         <%End If%>
        </td>
        <td>
         <%If mdbor7.EOF=True Then%>
          <%a7=0%>
         <%Else%>
          <%a7=mdbor7("Summym")%>
         <%End If%>
         <%If Request.Form("btn")="Kopeerimiseks" Then%>
          <%a0="aa" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>
          <%=Request.Form(a0)%>
         <%Else%>
          <%a0="aa" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>
          <input Type="Text" value="<%=CDBL(a7)%>" size="10" name="<%="aa" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>"  Class="Enterp">
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
         <%If mdbor4.EOF=faLSE Then%>
          <%a4=mdbor4("Summd")%>
         <%Else%>
          <%a4=0%>
         <%End If%>
         <%If mdbor6.EOF=True Then%>
          <%a6=0%>
         <%Else%>
          <%a6=mdbor6("Summy")%>
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
         <%=CDbl(a2)+CDBL(a1)-CDBL(a6)+CDBL(a7)%>
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
          <%a3=0%>
         <%Else%>
          <%a3=mdbor6a("SD")%>
         <%End If%>
         <%If Request.Form("btn")="Kopeerimiseks" Then%>
          <%a0="ac" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>
          <%=Request.Form(a0)%>
         <%Else%>
          <%a0="ac" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>
          <input Type="Text" value="<%=CDBL(a3)%>" size="10" name="<%="ac" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>"  Class="Enterp">
         <%End If%>
        </td>
        <td>
         <%If mdbor6a.EOF=true Then%>
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
         <%If mdbor4.EOF=faLSE Then%>
          <%a4=mdbor4("Summd")%>
         <%Else%>
          <%a4=0%>
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
        
       <%mdbo1.CommandText="SELECT * FROM inpl WHERE IDentIfier='C' AND Yearr=" & ya & " AND SUBSTRING(PC,1,2)='" & MID(mdborl1("PC"),1,2) & "' AND SUBSTRING(PC,4,2)='" & MID(mdborl2("PC"),4,2) & "' AND SUBSTRING(PC,7,2)<>'00' AND Enterprise='" & Mdborl3("Enterprise") & "' ORDER BY PC"%>
       <%mdbor.Open mdbo1%>
        
       <%Do Until mdbor.EOF%>
        <%mdbor2.Close%><%mdbor2u.Close%><%mdbor3.Close%><%mdbor4.Close%><%mdbor5.Close%><%mdbor6.Close%><%mdbor6a.Close%><%mdbor7.Close%>
        <%If LEN(mdbor("PC"))>9 AND RIGHT(mdbor("PC"),2)="00" Then%>
         <%mdbo2.CommandText="SELECT SUM(SummaPlan) AS SUMMAPLAN FROM delta WHERE LEFT(ProjCode,8)='" & LEFT(mdbor("PC"),8) & "' and Enterprise='" & Mdbor("Enterprise") & "' AND right(Projcode,2)<>'00' AND yearr='" & ya & "'"%>
         <%mdbor2.Open mdbo2%>
        <%Else%>
         <%mdbo2.CommandText="SELECT ProjCode,SummaPlan,OracleCode FROM delta WHERE Pid='" & mdbor("Pid") & "' AND Enterprise='" & Mdbor("Enterprise") & "' AND yearr='" & ya & "'"%>
         <%mdbor2.Open mdbo2%>
        <%End If%>
        <%If LEN(mdbor("PC"))>9 AND RIGHT(mdbor("PC"),2)="00" Then%>
         <%mdbo2u.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0),0) AS summi FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE m.Yearr='" & ya & "' AND SUBSTRING(m.ProjCode,1,8)='" & MId(mdbor("PC"),1,8) & "' and right(Projcode,2)<>'00' AND m.Enterprise='" & Mdbor("Enterprise") & "' AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND RIGHT(MES,2) <10 AND m.OracleCode='EJB206' AND (m.IDentIfier = 'C')"%>
         <%mdbor2u.Open mdbo2u%>
        <%Else%>
         <%mdbo2u.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(DEBET,0))/1000,0),0) AS summi, PROJECT FROM glav_project WHERE Project='" & mdbor("OracleCode") & "' AND  (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700')  AND Project='EJB206' AND  LEFT(MES,1)=" & Mid(ya,4,1) & " AND RIGHT(MES,2) >=04 AND RIGHT(MES,2)<10 GROUP BY PROJECT ORDER BY PROJECT DESC "%>
         <%mdbor2u.Open mdbo2u%>
        <%End If%>
        <%If LEN(mdbor("PC"))>9 AND RIGHT(mdbor("PC"),2)="00" Then%>
         <%mdbo3.CommandText="SELECT ISNULL(SUM(ROUND(ISNULL(GP.DEBET,0)/1000,0)),0) AS summd FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE m.Yearr='" & ya & "' AND SUBSTRING(m.ProjCode,1,8)='" & MId(mdbor("PC"),1,8) & "' and right(Projcode,2)<>'00' AND m.Enterprise='" & Mdbor("Enterprise") & "' AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND RIGHT(MES,2) <10 AND (m.IDentIfier = 'C') AND KONTO='43350' AND SUBKONTO='4351'"%>
         <%mdbor3.Open mdbo3%>
        <%Else%>
         <%mdbo3.CommandText="SELECT ISNULL(SUM(ROUND(ISNULL(DEBET,0)/1000,0)),0) AS summd, PROJECT FROM glav_project WHERE Project='" & mdbor("OracleCode") & "' AND LEFT(MES,1)=" & Mid(ya,4,1) & "  AND RIGHT(MES,2) >=04 AND RIGHT(MES,2) <10 AND KONTO='43350' AND SUBKONTO='4351' GROUP BY PROJECT ORDER BY PROJECT DESC "%>
         <%mdbor3.Open mdbo3%>
        <%End If%>
        <%If LEN(mdbor("PC"))>9 AND RIGHT(mdbor("PC"),2)="00" Then%>
         <%mdbo4.CommandText="SELECT ISNULL(SUM(ROUND(ISNULL(GP.CREDIT,0)/1000,0)),0) AS summc FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND m.Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND SUBSTRING(m.ProjCode,1,8)='" & MId(mdbor("PC"),1,8) & "' AND m.Enterprise='" & Mdbor("Enterprise") & "' and right(Projcode,2)<>'00' AND m.OracleCode<>'EJB206' AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND RIGHT(MES,2) <10 AND (m.IDentIfier = 'C')"%>
         <%mdbor4.Open mdbo4%>
        <%Else%>
         <%mdbo4.CommandText="SELECT ISNULL(SUM(ROUND(ISNULL(CREDIT,0)/1000,0)),0) AS summc, PROJECT FROM glav_project WHERE (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND Project='" & mdbor("OracleCode") & "' AND LEFT(MES,1)=" & Mid(ya,4,1) & " AND (konto NOT BETWEEN '18410' AND '18433') AND RIGHT(MES,2) >=04 AND Project<>'EJB206' AND RIGHT(MES,2)<10 GROUP BY PROJECT ORDER BY PROJECT DESC "%>
         <%mdbor4.Open mdbo4%>
        <%End If%>
        <%If LEN(mdbor("PC"))>9 AND RIGHT(mdbor("PC"),2)="00" Then%>
         <%mdbo5.CommandText="SELECT ISNULL(SUM(ROUND(ISNULL(GP.DEBET,0)/1000,0)),0) AS summi FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND ((Konto<>'18164' AND Konto<>'18151') OR (Konto='18164' AND (PROJECT='EJE573' OR PROJECT='EJE516' OR PROJECT='EJB515' OR PROJECT='EJB583') AND (MES='504' OR MES='505' OR MES='506'))) AND (KONTO<>'43350' AND SUBKONTO<>'4351') and m.Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433')  AND SUBSTRING(m.ProjCode,1,8)='" & MId(mdbor("PC"),1,8) & "' AND m.Enterprise='" & Mdbor("Enterprise") & "' and right(Projcode,2)<>'00' AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND RIGHT(MES,2) <10 AND m.OracleCode<>'EJB206' AND (m.IDentIfier = 'C')"%>
         <%mdbor5.Open mdbo5%>
        <%Else%>
         <%mdbo5.CommandText="SELECT ISNULL(SUM(ROUND(ISNULL(DEBET,0)/1000,0)),0) AS summi, PROJECT FROM glav_project WHERE (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND ((Konto<>'18164' AND Konto<>'18151') OR ((Konto='18164' OR Konto='18151') AND SUBSTRING(PROJECT,4,3)='999') OR (Konto='18164' AND (PROJECT='EJE573' OR PROJECT='EJE516' OR PROJECT='EJB515' OR PROJECT='EJB583') AND (MES='504' OR MES='505' OR MES='506'))) AND (KONTO<>'43350' AND SUBKONTO<>'4351') AND Project='" & mdbor("OracleCode") & "' AND LEFT(MES,1)=" & Mid(ya,4,1) & " AND (konto NOT BETWEEN '18410' AND '18433') AND RIGHT(MES,2) >=04 AND Project<>'EJB206' AND RIGHT(MES,2)<10 GROUP BY PROJECT ORDER BY PROJECT DESC "%>
         <%mdbor5.Open mdbo5%>
        <%End If%>
        <%If LEN(mdbor("PC"))>9 AND RIGHT(mdbor("PC"),2)="00" Then%>
         <%mdbo6.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0),0)-ISNULL(ROUND(SUM(ISNULL(GP.CREDIT,0))/1000,0),0) AS summy FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE m.Yearr='" & ya-1 & "' AND SUBSTRING(m.ProjCode,1,8)='" & MId(mdbor("PC"),1,8) & "' AND m.Enterprise='" & Mdbor("Enterprise") & "' and right(Projcode,2)<>'00' AND (konto BETWEEN '18410' AND '18433') AND (((SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya-1,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04)) OR (LEFT(MES,1)=" & Mid(ya,4,1) & " AND RIGHT(MES,2) <04)) AND (m.IDentIfier = 'C')"%>
         <%mdbor6.Open mdbo6%>
        <%Else%>
         <%mdbo6.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0),0)-ISNULL(ROUND(SUM(ISNULL(GP.CREDIT,0))/1000,0),0) AS summy FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE Project='" & mdbor("OracleCode") & "' AND m.Yearr='" & ya-1 & "' AND (konto BETWEEN '18410' AND '18433') AND (((SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya-1,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04)) OR (LEFT(MES,1)=" & Mid(ya,4,1) & " AND RIGHT(MES,2) <04)) AND (m.IDentIfier = 'C')"%>
         <%mdbor6.Open mdbo6%>
        <%End If%>
        <%If LEN(mdbor("PC"))>9 AND RIGHT(mdbor("PC"),2)="00" Then%>
         <%mdbo6a.CommandText="SELECT DISTINCT ISNULL(SUM(ISNULL(Ettemaks,0)),0) AS EM,ISNULL(SUM(ISNULL(Saldo,0)),0) AS SD FROM dbo.ETTE WHERE LEFT(ProjCode,8)='" & LEFT(mdbor("PC"),8) & "' and Enterprise='" & Mdbor("Enterprise") & "' AND right(Projcode,2)<>'00' AND yearr='" & ya & "'"%>
         <%mdbor6a.Open mdbo6a%>
        <%Else%>
         <%mdbo6a.CommandText="SELECT DISTINCT ISNULL(Ettemaks,0) AS EM,ISNULL(Saldo,0) AS SD FROM dbo.ETTE WHERE Pid='" & mdbor("Pid") & "' AND Enterprise='" & Mdbor("Enterprise") & "' AND yearr='" & ya & "'"%>
         <%mdbor6a.Open mdbo6a%>
        <%End If%>
        <%If LEN(mdbor("PC"))>9 AND RIGHT(mdbor("PC"),2)="00" Then%>
         <%mdbo7.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0),0)-ISNULL(ROUND(SUM(ISNULL(GP.CREDIT,0))/1000,0),0) AS summym FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE m.Yearr='" & ya & "' AND SUBSTRING(m.ProjCode,1,8)='" & MId(mdbor("PC"),1,8) & "' AND m.Enterprise='" & Mdbor("Enterprise") & "' and right(Projcode,2)<>'00' AND (konto BETWEEN '18410' AND '18433') AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) < 10) AND (m.IDentIfier = 'C')"%>
         <%mdbor7.Open mdbo7%>
        <%Else%>
         <%mdbo7.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0),0)-ISNULL(ROUND(SUM(ISNULL(GP.CREDIT,0))/1000,0),0) AS summym FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE Project='" & mdbor("OracleCode") & "' AND m.Yearr='" & ya & "' AND (konto BETWEEN '18410' AND '18433') AND GP.MES <" & Mid(ya,4,1) & "10 AND (m.IDentIfier = 'C')"%>
         <%mdbor7.Open mdbo7%>
        <%End If%>
        <tr>
         <td>
          <%If mdbor("OracleCode")="N/A" Then%>
           &nbsp
          <%Else%>
           <%=mdbor("OracleCode")%>
          <%End If%>
         </td>
         <td>
         <%if mid(mdbor("PC"),8,1)=0 and mid(mdbor("PC"),7,1)<>0 then%>
          <%a=REPLACE(MID(mdbor("PC"),1,6), "0", "") & MID(mdbor("PC"),7,2)%>
         <%Else%>
          <%a=REPLACE(mdbor("PC"), "0", "")%>
	 <%End If%>
          <%If len(a)>=7 Then%>
	  <%If Right(mdbor("PC"),2)="00" then%>
	   <%=a%>
	  <%Else%>
	   <%=a%>.
	  <%End if%>
	 <%Else%>
	  <%If Right(mdbor("PC"),2)="00" then%>
	   <%=mid(a,1,6)%>
	  <%Else%>
	   <%=mid(a,1,6)%>.
	  <%End if%>
	 <%End if%>
         </td>
         <td>
          <%=mdbor("ProjName")%>
         </td>
         <td>
          <%=mdbor2("SummaPlan")%>
         </td>
         <td>
          <%If mdbor2u.EOF=True Then%>
           <%a1=0%>
          <%Else%>
           <%a1=mdbor2u("Summi")%>
          <%End If%>
          <%If mdbor5.EOF=True Then%>
           <%a2=0%>
          <%Else%>
           <%a2=mdbor5("Summi")%>
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
           <%a0="ab" & mdborl3("Enterprise") & "_" & mdbor("Pid")%>
           <%=Request.Form(a0)%>
          <%Else%>
           <%a0="ab" & mdborl3("Enterprise") & "_" & mdbor("Pid")%>
           <input Type="Text" value="<%=CDBL(a6)%>" size="10" name="<%="ab" & mdborl3("Enterprise") & "_" & mdbor("Pid")%>" >
          <%End If%>
         </td>
         <td>
          <%If mdbor7.EOF=True Then%>
           <%a7=0%>
          <%Else%>
           <%a7=mdbor7("Summym")%>
          <%End If%>
          <%If Request.Form("btn")="Kopeerimiseks" Then%>
           <%a0="aa" & mdborl3("Enterprise") & "_" & mdbor("Pid")%>
           <%=Request.Form(a0)%>
          <%Else%>
           <%a0="aa" & mdborl3("Enterprise") & "_" & mdbor("Pid")%>
           <input Type="Text" value="<%=CDBL(a7)%>" size="10" name="<%="aa" & mdborl3("Enterprise") & "_" & mdbor("Pid")%>">
          <%End If%>
         </td>
         <td>
          <%If mdbor2u.EOF=True Then%>
           <%a1=0%>
          <%Else%>
           <%a1=mdbor2u("Summi")%>
          <%End If%>
          <%If mdbor5.EOF=True Then%>
           <%a2=0%>
          <%Else%>
           <%a2=mdbor5("Summi")%>
          <%End If%>
          <%If mdbor3.EOF=faLSE Then%>
           <%a4=mdbor3("Summd")%>
          <%Else%>
           <%a4=0%>
          <%End If%>
          <%If mdbor6.EOF=True Then%>
           <%a6=0%>
          <%Else%>
           <%a6=mdbor6("Summy")%>
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
          <%=CDbl(a2)+CDBL(a1)-CDBL(a6)+CDBL(a7)%>
         </td>
         <td>
          <%If MID(mdbor("OracleCode"),4,3)="999" Then%>
           <%If mdbor5.EOF=FALSE Then%>
            <%a3=mdbor5("Summi")%>
           <%Else%>
            <%a3=0%>
           <%End If%>
          <%Else%>
           <%If mdbor4.EOF=FALSE Then%>
            <%a3=mdbor4("Summc")%>
           <%Else%>
            <%a3=0%>
           <%End If%>
          <%End If%>
          <%If mdbor2u.EOF=FALSE Then%>
           <%a2=mdbor2u("Summi")%>
          <%Else%>
           <%a2=0%>
          <%End If%>
          <%=cdbl(a3)%>
         </td>
         <td>
          <%If mdbor3.EOF=faLSE Then%>
           <%=mdbor3("Summd")%>
          <%Else%>
           0
          <%End If%>
         </td>
         <td>
          <%If mdbor6a.EOF=True Then%>
           <%a3=0%>
          <%Else%>
           <%a3=mdbor6a("SD")%>
          <%End If%>
          <%If Request.Form("btn")="Kopeerimiseks" Then%>
           <%a0="ac" & mdborl3("Enterprise") & "_" & mdbor("Pid")%>
           <%=Request.Form(a0)%>
          <%Else%>
           <%a0="ac" & mdborl3("Enterprise") & "_" & mdbor("Pid")%>
           <input Type="Text" value="<%=CDBL(a3)%>" size="10" name="<%="ac" & mdborl3("Enterprise") & "_" & mdbor("Pid")%>">
          <%End If%>
         </td>
         <td>
          <%If mdbor6A.EOF=true Then%>
           <%a6=0%>
          <%Else%>
           <%a6=mdbor6a("SD")%>
          <%End If%>
          <%If mdbor3.EOF=faLSE Then%>
           <%a4=mdbor3("Summd")%>
          <%Else%>
           <%a4=0%>
          <%End If%>
          <%If MID(mdbor("OracleCode"),4,3)="999" Then%>
           <%If mdbor5.EOF=FALSE Then%>
            <%a3=mdbor5("Summi")%>
           <%Else%>
            <%a3=0%>
           <%End If%>
          <%Else%>
           <%If mdbor4.EOF=FALSE Then%>
            <%a3=mdbor4("Summc")%>
           <%Else%>
            <%a3=0%>
           <%End If%>
          <%End If%>
          <%If mdbor2u.EOF=True Then%>
           <%a1=0%>
          <%Else%>
           <%a1=mdbor2u("Summi")%>
          <%End If%>
          <%If mdbor5.EOF=True Then%>
           <%a2=0%>
          <%Else%>
           <%a2=mdbor5("Summi")%>
          <%End If%>
          <%a0="a1c"%>
          <%=CDbl(a6)-CDbl(a3)+CDbl(a1)+CDbl(a2)%>
         </td>
        </tr>
        <%mdbor.MoveNext%>
       <%Loop%>
       <%mdbor.Close%>
       <%mdborl3.MoveNext%>
      <%Loop%>
      <%mdborl3.Close%>
      <%mdborl2.MoveNext%>
     <%Loop%>
     <%mdborl2.Close%>
     <%mdborl1.MoveNext%>
    <%Loop%>
    <tr>
     <td colspan="12">Kokku ettev&otilde;tete kaupa</td>
    </tr>
    <%mdborl1.Close%><%mdbor2.Close%><%mdbor2u.Close%><%mdbor3.Close%><%mdbor3u.Close%><%mdbor4.Close%><%mdbor5.Close%><%mdbor6.Close%><%mdbor6a.Close%><%mdbor7.Close%>
      
    <%mdbol1.CommandText="SELECT * FROM Enterprise ORDER BY Enterprise"%>
    <%mdborl1.Open mdbol1%>
       
    <%mdbo2.CommandText="SELECT ROUND((SUM(ISNULL(GP.DEBET,0))/1000),0) AS summi, Enterprise FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND ((Konto<>'18164' AND Konto<>'18151') OR ((Konto='18164' OR Konto='18151') AND SUBSTRING(PROJECT,4,3)='999') OR (Konto='18164' AND (PROJECT='EJE573' OR PROJECT='EJE516' OR PROJECT='EJB515' OR PROJECT='EJB583') AND (MES='504' OR MES='505' OR MES='506'))) AND (KONTO<>'43350' AND SUBKONTO<>'4351') AND m.Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND (SUBSTRING(GP.MES, 2, 2) < 10) AND m.OracleCode<>'EJB206' AND (m.IDentIfier = 'C') GROUP BY Enterprise ORDER BY Enterprise"%>
    <%mdbor2.Open mdbo2%>
    <%mdbo2u.CommandText="SELECT ROUND((SUM(ISNULL(GP.DEBET,0))/1000),0)-ROUND((SUM(ISNULL(GP.CREDIT,0))/1000),0) AS summi, Enterprise FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND m.Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND (SUBSTRING(GP.MES, 2, 2) < 10) AND m.OracleCode='EJB206' AND (m.IDentIfier = 'C') GROUP BY Enterprise ORDER BY Enterprise"%>
    <%mdbor2u.Open mdbo2u%>
    <%mdbo3.CommandText="SELECT ROUND(SUM(ISNULL(GP.CREDIT,0))/1000,0) AS summc, Enterprise FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE SUBSTRING(PROJECT,4,3)<>'999' AND (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') And m.Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND m.OracleCode<>'EJB206' AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND (SUBSTRING(GP.MES, 2, 2) < 10) AND (m.IDentIfier = 'C') GROUP BY Enterprise ORDER BY Enterprise"%>
    <%mdbor3.Open mdbo3%>
    <%mdbo3u.CommandText="SELECT ROUND((SUM(ISNULL(GP.DEBET,0))/1000),0) AS summc, Enterprise FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE SUBSTRING(PROJECT,4,3)='999' AND (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND m.Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND (SUBSTRING(GP.MES, 2, 2) < 10) AND m.OracleCode<>'EJB206' AND (m.IDentIfier = 'C') GROUP BY Enterprise ORDER BY Enterprise"%>
    <%mdbor3u.Open mdbo3u%>
    <%mdbo4.CommandText="SELECT ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0) AS summd, Enterprise FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE m.Yearr='" & ya & "' AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND (SUBSTRING(GP.MES, 2, 2) < 10) AND (m.IDentIfier = 'C') AND KONTO='43350' AND SUBKONTO='4351' GROUP BY Enterprise ORDER BY Enterprise"%>
    <%mdbor4.Open mdbo4%>
    <%mdbo5.CommandText="SELECT DISTINCT ISNULL(SUM(ISNULL(SummaPlan,0)),0) AS SP, Enterprise FROM dbo.Delta WHERE yearr='" & ya & "' AND enn<>'00' AND RenovBlock=0 GROUP BY Enterprise ORDER BY Enterprise"%>
    <%mdbor5.Open mdbo5%>
    <%mdbo6.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0),0)-ISNULL(ROUND(SUM(ISNULL(GP.CREDIT,0))/1000,0),0) AS summy, Enterprise FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE m.Yearr='" & ya-1 & "' AND (konto BETWEEN '18410' AND '18433') AND (((SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya-1,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04)) OR (LEFT(MES,1)=" & Mid(ya,4,1) & " AND RIGHT(MES,2) <04))  AND (m.IDentIfier = 'C') GROUP BY Enterprise ORDER BY Enterprise"%>
    <%mdbor6.Open mdbo6%>
    <%mdbo6a.CommandText="SELECT DISTINCT SUM(ISNULL(Ettemaks,0)) AS EM,SUM(ISNULL(Saldo,0)) AS SD, Enterprise FROM dbo.ETTE WHERE yearr='" & ya & "' AND enn<>'00' AND RenovBlock=0 GROUP BY Enterprise ORDER BY Enterprise"%>
    <%mdbor6a.Open mdbo6a%>
    <%mdbo7.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0),0)-ISNULL(ROUND(SUM(ISNULL(GP.CREDIT,0))/1000,0),0) AS summym, Enterprise FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE m.Yearr='" & ya & "' AND (konto BETWEEN '18410' AND '18433') AND GP.MES <" & Mid(ya,4,1) & "10 AND (m.IDentIfier = 'C') GROUP BY Enterprise ORDER BY Enterprise"%>
    <%mdbor7.Open mdbo7%>
       
    <%Do Until mdborl1.EOF%>
     <%If mdbor7.EOF=true Then%>
      <%a7=0%>
     <%Else%>
      <%If mdbor7("Enterprise")=mdborl1("Enterprise") Then%>
       <%a7=mdbor7("Summym")%>
       <%mdbor7.MoveNext%>
      <%Else%>
       <%a7=0%>
      <%End If%>
     <%End If%>
     <%If mdbor6a.EOF=true Then%>
      <%a6a=0%>
      <%a33=0%>
     <%Else%>
      <%If mdbor6a("Enterprise")=mdborl1("Enterprise") Then%>
       <%a6a=mdbor6a("EM")%>
       <%a33=mdbor6a("SD")%>
       <%mdbor6a.MoveNext%>
      <%Else%>
       <%a6a=0%>
       <%a33=0%>
      <%End If%>
     <%End If%>
     <%If mdbor6.EOF=true Then%>
      <%a6=0%>
     <%Else%>
      <%If mdbor6("Enterprise")=mdborl1("Enterprise") Then%>
       <%a6=mdbor6("Summy")%>
       <%mdbor6.MoveNext%>
      <%Else%>
       <%a6=0%>
      <%End If%>
     <%End If%>
     <%If mdbor2u.EOF=true Then%>
      <%a1=0%>
     <%Else%>
      <%If mdbor2u("Enterprise")=mdborl1("Enterprise") Then%>
       <%a1=mdbor2u("Summi")%>
       <%mdbor2u.MoveNext%>
      <%Else%>
       <%a1=0%>
      <%End If%>
     <%End If%> 
     <%If mdbor4.EOF=true Then%>
      <%a4=0%>
     <%Else%>
      <%If mdbor4("Enterprise")=mdborl1("Enterprise") Then%>
       <%a4=mdbor4("Summd")%>
       <%mdbor4.MoveNext%>
      <%Else%>
       <%a4=0%>
      <%End If%>
     <%End If%>
     <%If mdbor2.EOF=true Then%>
      <%a2=0%>
     <%Else%>
      <%If mdbor2("Enterprise")=mdborl1("Enterprise") Then%>
       <%a2=mdbor2("Summi")%>
       <%mdbor2.MoveNext%>
      <%Else%>
       <%a2=0%>
      <%End If%>
     <%End If%>
     <%If mdbor3.EOF=true Then%>
      <%a3=0%>
     <%Else%>
      <%If mdbor3("Enterprise")=mdborl1("Enterprise") Then%>
       <%a3=mdbor3("Summc")%>
       <%mdbor3.MoveNext%>
      <%Else%>
       <%a3=0%>
      <%End If%>
     <%End If%>
     <%If mdbor3u.EOF=true Then%>
      <%a9=0%>
     <%Else%>
      <%If mdbor3u("Enterprise")=mdborl1("Enterprise") Then%>
       <%a9=mdbor3u("Summc")%>
       <%mdbor3u.MoveNext%>
      <%Else%>
       <%a9=0%>
      <%End If%>
     <%End If%>
     <tr Class="boldEnterp">
      <td colspan=3>
       <%=Mdborl1("EDescr")%>
      </td>
      <td>
       <%If mdbor5.EOF=false Then%>
        <%=mdbor5("SP")%>
       <%Else%>
        0
       <%End If%>
      </td>
      <td>
       <%=CDbl(a2)+CDBL(a1)%>
      </td>
      <td>
       <%If Request.Form("btn")="Kopeerimiseks" Then%>
        <%a0="a8b" & mdborl1("Enterprise")%>
        <%=Request.Form(a0)%>
       <%Else%>
        <%a0="a8b" & mdborl1("Enterprise")%>
        <input Type="Text" value="<%=CDBL(a6)%>" size="10" name="<%="a8b" & mdborl1("Enterprise")%>" Class="boldEnterp">
       <%End If%>
      </td>
      <td>
       <%If Request.Form("btn")="Kopeerimiseks" Then%>
        <%a0="a8a" & mdborl1("Enterprise")%>
        <%=Request.Form(a0)%>
       <%Else%>
        <%a0="a8a" & mdborl1("Enterprise")%>
        <input Type="Text" value="<%=CDBL(a7)%>" size="10" name="<%="a8a" & mdborl1("Enterprise")%>" Class="boldEnterp">
       <%End If%>
      </td>
      <td>
       <%=CDbl(a2)+CDBL(a1)-CDBL(a6a)-CDBL(a6)+CDBL(a7)%>
      </td>
      <td>
       <%=cdbl(a3)+cdbl(a9)%>
      </td>
      <td>
       <%=cdbl(a4)%>
      </td>
      <td>
       <%If Request.Form("btn")="Kopeerimiseks" Then%>
        <%a0="a8c" & mdborl1("Enterprise")%>
        <%=Request.Form(a0)%>
       <%Else%>
        <%a0="a8c" & mdborl1("Enterprise")%>
        <input Type="Text" value="<%=CDBL(a33)%>" size="10" name="<%="a8c" & mdborl1("Enterprise")%>" Class="boldEnterp">
       <%End If%>
      </td>
      <td>
       <%=CDbl(a33)-CDbl(a3)-CDbl(a9)+CDbl(a1)+CDbl(a2)%>
      </td>
     </tr>
     <%If mdbor5.EOF <> true Then%>
      <%mdbor5.MoveNext%>
     <%End If%>
     <%mdborl1.MoveNext%>
    <%Loop%>
     
    <%mdbor2.Close%><%mdbor2U.Close%><%mdbor3.Close%><%mdbor3U.Close%><%mdbor4.Close%><%mdbor5.Close%><%mdbor6.Close%><%mdbor6a.Close%><%mdbor7.Close%>
    
    <%mdbo2.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0),0) AS summi FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND ((Konto<>'18164' AND Konto<>'18151') OR ((Konto='18164' OR Konto='18151') AND SUBSTRING(PROJECT,4,3)='999') OR (Konto='18164' AND (PROJECT='EJE573' OR PROJECT='EJE516' OR PROJECT='EJB515' OR PROJECT='EJB583') AND (MES='504' OR MES='505' OR MES='506'))) AND (KONTO<>'43350' AND SUBKONTO<>'4351') AND m.Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND (SUBSTRING(GP.MES, 2, 2) < 10) AND m.OracleCode<>'EJB206' AND (m.IDentIfier = 'C')"%>
    <%mdbor2.Open mdbo2%>
    <%mdbo2u.CommandText="SELECT ISNULL(ROUND((SUM(ISNULL(GP.DEBET,0))/1000),0)-ROUND((SUM(ISNULL(GP.CREDIT,0))/1000) ,0),0) AS summi FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') and m.Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND (SUBSTRING(GP.MES, 2, 2) < 10) AND m.OracleCode='EJB206' AND (m.IDentIfier = 'C')"%>
    <%mdbor2u.Open mdbo2u%>
    <%mdbo3.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.CREDIT,0))/1000,0),0) AS summc FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE SUBSTRING(PROJECT,4,3)<>'999' AND (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700')  And m.Yearr='" & ya & "' AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (konto NOT BETWEEN '18410' AND '18433') AND m.OracleCode<>'EJB206' AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND (SUBSTRING(GP.MES, 2, 2) < 10) AND (m.IDentIfier = 'C')"%>
    <%mdbor3.Open mdbo3%>
    <%mdbo3u.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0),0) AS summc FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE SUBSTRING(PROJECT,4,3)='999' AND (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') and  m.Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND (SUBSTRING(GP.MES, 2, 2) < 10) AND m.OracleCode<>'EJB206' AND (m.IDentIfier = 'C')"%>
    <%mdbor3u.Open mdbo3u%>
    <%mdbo4.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0),0) AS summd FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE m.Yearr='" & ya & "' AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND (SUBSTRING(GP.MES, 2, 2) < 10) AND (m.IDentIfier = 'C') AND KONTO='43350' AND SUBKONTO='4351'"%>
    <%mdbor4.Open mdbo4%>
    <%mdbo5.CommandText="SELECT DISTINCT ISNULL(SUM(ISNULL(SummaPlan,0)),0) AS SP FROM dbo.Delta WHERE yearr='" & ya & "' AND enn<>'00'"%>
    <%mdbor5.Open mdbo5%>
    <%mdbo6.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0),0)-ISNULL(ROUND(SUM(ISNULL(GP.CREDIT,0))/1000,0),0) AS summy FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE m.Yearr='" & ya-1 & "' AND (konto BETWEEN '18410' AND '18433') AND (((SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya-1,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04)) OR (LEFT(MES,1)=" & Mid(ya,4,1) & " AND RIGHT(MES,2) <04))  AND (m.IDentIfier = 'C')"%>
    <%mdbor6.Open mdbo6%>
    <%mdbo6a.CommandText="SELECT DISTINCT SUM(ISNULL(Ettemaks,0)) AS EM,SUM(ISNULL(Saldo,0)) AS SD FROM dbo.ETTE WHERE yearr='" & ya & "' AND enn<>'00'"%>
    <%mdbor6a.Open mdbo6a%>
    <%mdbo7.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0),0)-ISNULL(ROUND(SUM(ISNULL(GP.CREDIT,0))/1000,0),0) AS summym FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE m.Yearr='" & ya & "' AND (konto BETWEEN '18410' AND '18433') AND GP.MES <" & Mid(ya,4,1) & "10 AND (m.IDentIfier = 'C')"%>
    <%mdbor7.Open mdbo7%>
    <tr Class="Bold">
     <td colspan=3>Kokku ettev&otildetete kaupa</td>
     <td>
      <%If mdbor5.EOF=false Then%>
       <%=mdbor5("SP")%>
      <%Else%>
       0
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
       <%a0="a4b"%>
       <%=Request.Form(a0)%>
      <%Else%>
       <%a0="a4b"%>
       <input Type="Text" value="<%=CDBL(a6)%>" size="10" name="<%="a4b"%>"  Class="bold">
      <%End If%>
     </td>
     <td>
      <%If mdbor7.EOF=True Then%>
       <%a7=0%>
      <%Else%>
       <%a7=mdbor7("Summym")%>
      <%End If%>
      <%If Request.Form("btn")="Kopeerimiseks" Then%>
       <%a0="a4a"%>
       <%=Request.Form(a0)%>
      <%Else%>
       <%a0="a4a"%>  
       <input Type="Text" value="<%=CDBL(a7)%>" size="10" name="<%="a4a"%>"  Class="bold">
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
      <%If mdbor4.EOF=true Then%>
       <%a4=0%>
      <%Else%>
       <%a4=mdbor4("Summd")%>
      <%End If%>
      <%=CDbl(a2)+CDBL(a1)-CDBL(a3)-CDBL(a6)+CDBL(a7)%>
     </td>
     <td>
      <%If mdbor3.EOF=FALSE Then%>
       <%a3=mdbor3("Summc")%>
      <%Else%>
       <%a3=0%>
      <%End If%>
      <%If mdbor3U.EOF=FALSE Then%>
       <%a4=mdbor3U("Summc")%>
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
       <%a3=0%>
      <%Else%>
       <%a3=mdbor6a("SD")%>
      <%End If%>
      <%If Request.Form("btn")="Kopeerimiseks" Then%>
       <%a0="a4c"%>
       <%=Request.Form(a0)%>
      <%Else%>
       <%a0="a4c"%>
       <input Type="Text" value="<%=CDBL(a3)%>" size="10" name="<%="a4c"%>"  Class="bold">
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
      <%a0="a4c"%>
      <%=CDbl(a6)-CDbl(a3)-CDbl(a9)+CDbl(a1)+CDbl(a2)%>
     </td>
    </tr>
    <%mdborl1.Close%><%mdbor2.Close%><%mdbor2u.Close%><%mdbor3.Close%><%mdbor3u.Close%><%mdbor4.Close%><%mdbor5.Close%><%mdbor6.Close%><%mdbor6a.Close%><%mdbor7.Close%>
      
    <%mdbol1.CommandText="SELECT * FROM Enterprise ORDER BY Enterprise"%>
    <%mdborl1.Open mdbol1%>
        
    <%mdbo2.CommandText="SELECT ISNULL(ROUND((SUM(ISNULL(GP.DEBET,0))/1000),0),0) AS summi, Enterprise FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND ((Konto<>'18164' AND Konto<>'18151') OR ((Konto='18164' OR Konto='18151') AND SUBSTRING(PROJECT,4,3)='999') OR (Konto='18164' AND (PROJECT='EJE573' OR PROJECT='EJE516' OR PROJECT='EJB515' OR PROJECT='EJB583') AND (MES='504' OR MES='505' OR MES='506'))) AND (KONTO<>'43350' AND SUBKONTO<>'4351') AND m.Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND m.RenovBlock=0 AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND (SUBSTRING(GP.MES, 2, 2) < 10) AND m.OracleCode<>'EJB206' AND (m.IDentIfier = 'C') GROUP BY Enterprise ORDER BY Enterprise"%>
    <%mdbor2.Open mdbo2%>
    <%mdbo2u.CommandText="SELECT ISNULL(ROUND((SUM(ISNULL(GP.DEBET,0))/1000),0),0)-ISNULL(ROUND((SUM(ISNULL(GP.CREDIT,0))/1000),0),0) AS summi, Enterprise FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND m.Yearr='" & ya & "' AND m.RenovBlock=0 AND (konto NOT BETWEEN '18410' AND '18433') AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND (SUBSTRING(GP.MES, 2, 2) < 10) AND m.OracleCode='EJB206' AND (m.IDentIfier = 'C') GROUP BY Enterprise ORDER BY Enterprise"%>
    <%mdbor2u.Open mdbo2u%>
    <%mdbo3.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.CREDIT,0))/1000,0),0) AS summc, Enterprise FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE SUBSTRING(PROJECT,4,3)<>'999' AND (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') And m.Yearr='" & ya & "' AND m.RenovBlock=0 AND (konto NOT BETWEEN '18410' AND '18433') AND m.OracleCode<>'EJB206' AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND (SUBSTRING(GP.MES, 2, 2) < 10) AND (m.IDentIfier = 'C') GROUP BY Enterprise ORDER BY Enterprise"%>
    <%mdbor3.Open mdbo3%>
    <%mdbo3u.CommandText="SELECT ISNULL(ROUND((SUM(ISNULL(GP.DEBET,0))/1000),0),0) AS summc, Enterprise FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE SUBSTRING(PROJECT,4,3)='999' AND (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND m.Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND m.RenovBlock=0 AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND (SUBSTRING(GP.MES, 2, 2) < 10) AND m.OracleCode<>'EJB206' AND (m.IDentIfier = 'C') GROUP BY Enterprise ORDER BY Enterprise"%>
    <%mdbor3u.Open mdbo3u%>
    <%mdbo4.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0),0) AS summd, Enterprise FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE m.Yearr='" & ya & "' AND m.RenovBlock=0 AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND (SUBSTRING(GP.MES, 2, 2) < 10) AND (m.IDentIfier = 'C') AND KONTO='43350' AND SUBKONTO='4351' GROUP BY Enterprise ORDER BY Enterprise"%>
    <%mdbor4.Open mdbo4%>
    <%mdbo5.CommandText="SELECT DISTINCT ISNULL(SUM(ISNULL(SummaPlan,0)),0) AS SP, Enterprise FROM dbo.Delta WHERE yearr='" & ya & "' AND enn<>'00' AND RenovBlock=0 GROUP BY Enterprise ORDER BY Enterprise"%>
    <%mdbor5.Open mdbo5%>
    <%mdbo6.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0),0)-ISNULL(ROUND(SUM(ISNULL(GP.CREDIT,0))/1000,0),0) AS summy, Enterprise FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE m.Yearr='" & ya-1 & "' AND (konto BETWEEN '18410' AND '18433') AND m.RenovBlock=0 AND (((SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya-1,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04)) OR (LEFT(MES,1)=" & Mid(ya,4,1) & " AND RIGHT(MES,2) <04))  AND (m.IDentIfier = 'C') GROUP BY Enterprise ORDER BY Enterprise"%>
    <%mdbor6.Open mdbo6%>
    <%mdbo6a.CommandText="SELECT DISTINCT ISNULL(SUM(ISNULL(Ettemaks,0)),0) AS EM,ISNULL(SUM(ISNULL(Saldo,0)),0) AS SD, Enterprise FROM dbo.ETTE WHERE yearr='" & ya & "' AND enn<>'00' AND RenovBlock=0 GROUP BY Enterprise ORDER BY Enterprise"%>
    <%mdbor6a.Open mdbo6a%>
    <%mdbo7.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0),0)-ISNULL(ROUND(SUM(ISNULL(GP.CREDIT,0))/1000,0),0) AS summym, Enterprise FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE m.Yearr='" & ya & "' AND (konto BETWEEN '18410' AND '18433') AND m.RenovBlock=0 AND GP.MES <" & Mid(ya,4,1) & "10 AND (m.IDentIfier = 'C') GROUP BY Enterprise ORDER BY Enterprise"%>
    <%mdbor7.Open mdbo7%>
      
    <%Do Until mdborl1.EOF%>
     <%If mdbor7.EOF=true Then%>
      <%a7=0%>
     <%Else%>
      <%If mdbor7("Enterprise")=mdborl1("Enterprise") Then%>
       <%a7=mdbor7("Summym")%>
       <%mdbor7.MoveNext%>
      <%Else%>
       <%a7=0%>
      <%End If%>
     <%End If%>     
     <%If mdbor6a.EOF=true Then%>
      <%a6a=0%>
      <%a33=0%>
     <%Else%>
      <%If mdbor6a("Enterprise")=mdborl1("Enterprise") Then%>
       <%a6a=mdbor6a("EM")%>
       <%a33=mdbor6a("SD")%>
       <%mdbor6a.MoveNext%>
      <%Else%>
       <%a6a=0%>
       <%a33=0%>
      <%End If%>
     <%End If%>
     <%If mdbor6.EOF=true Then%>
      <%a6=0%>
     <%Else%>
      <%If mdbor6("Enterprise")=mdborl1("Enterprise") Then%>
       <%a6=mdbor6("Summy")%>
       <%mdbor6.MoveNext%>
      <%Else%>
       <%a6=0%>
      <%End If%>
     <%End If%>
     <%If mdbor2u.EOF=true Then%>
      <%a1=0%>
     <%Else%>
      <%If mdbor2u("Enterprise")=mdborl1("Enterprise") Then%>
       <%a1=mdbor2u("Summi")%>
       <%mdbor2u.MoveNext%>
      <%Else%>
       <%a1=0%>
      <%End If%>
     <%End If%> 
     <%If mdbor4.EOF=true Then%>
      <%a4=0%>
     <%Else%>
      <%If mdbor4("Enterprise")=mdborl1("Enterprise") Then%>
       <%a4=mdbor4("Summd")%>
       <%mdbor4.MoveNext%>
      <%Else%>
       <%a4=0%>
      <%End If%>
     <%End If%>
     <%If mdbor2.EOF=true Then%>
      <%a2=0%>
     <%Else%>
      <%If mdbor2("Enterprise")=mdborl1("Enterprise") Then%>
       <%a2=mdbor2("Summi")%>
       <%mdbor2.MoveNext%>
      <%Else%>
       <%a2=0%>
      <%End If%>
     <%End If%>
     <%If mdbor3.EOF=true Then%>
      <%a3=0%>
     <%Else%>
      <%If mdbor3("Enterprise")=mdborl1("Enterprise") Then%>
       <%a3=mdbor3("Summc")%>
       <%mdbor3.MoveNext%>
      <%Else%>
       <%a3=0%>
      <%End If%>
     <%End If%>
     <%If mdbor3u.EOF=true Then%>
      <%a9=0%>
     <%Else%>
      <%If mdbor3u("Enterprise")=mdborl1("Enterprise") Then%>
       <%a9=mdbor3u("Summc")%>
       <%mdbor3u.MoveNext%>
      <%Else%>
       <%a9=0%>
      <%End If%>
     <%End If%>
       
     <tr Class="boldEnterp">
      <td colspan=3><%=Mdborl1("EDescr")%></td>
      <td>
       <%If mdbor5.EOF=false Then%>
        <%=mdbor5("SP")%>
       <%Else%>
        0
       <%End If%>
      </td>
      <td>
       <%=CDbl(a2)+CDBL(a1)%>
      </td>
      <td>
       <%If Request.Form("btn")="Kopeerimiseks" Then%>
        <%a0="a5b" & mdborl1("Enterprise")%>
        <%=Request.Form(a0)%>
       <%Else%>
        <%a0="a5b" & mdborl1("Enterprise")%>
        <input Type="Text" value="<%=CDBL(a6)%>" size="10" name="<%="a5b" & mdborl1("Enterprise")%>" Class="boldEnterp">
       <%End If%>
      </td>
      <td>
       <%If Request.Form("btn")="Kopeerimiseks" Then%>
        <%a0="a5a" & mdborl1("Enterprise")%>
        <%=Request.Form(a0)%>
       <%Else%>
        <%a0="a5a" & mdborl1("Enterprise")%>
        <input Type="Text" value="<%=CDBL(a7)%>" size="10" name="<%="a5a" & mdborl1("Enterprise")%>" Class="boldEnterp">
       <%End If%>
      </td>
      <td>
       <%=CDbl(a2)+CDBL(a1)-CDBL(a6a)-CDBL(a6)+CDBL(a7)%>
      </td>
      <td>
       <%=cdbl(a3)+cdbl(a9)%>
      </td>
      <td>
       <%=cdbl(a4)%>
      </td>
      <td>
       <%If Request.Form("btn")="Kopeerimiseks" Then%>
        <%a0="a5c" & mdborl1("Enterprise")%>
        <%=Request.Form(a0)%>
       <%Else%>
        <%a0="a5c" & mdborl1("Enterprise")%>
        <input Type="Text" value="<%=CDBL(a33)%>" size="10" name="<%="a5c" & mdborl1("Enterprise")%>" Class="boldEnterp">
       <%End If%>
      </td>
      <td>
       <%a0="a5c"%>
       <%=CDbl(a33)-CDbl(a3)-CDbl(a9)+CDbl(a1)+CDbl(a2)%>
      </td>
     </tr>
     <%If mdbor5.EOF <> true Then%>
      <%mdbor5.MoveNext%>
     <%End If%>
     <%mdborl1.MoveNext%>
    <%Loop%>
      
    <%mdbor2.Close%><%mdbor2U.Close%><%mdbor3.Close%><%mdbor3u.Close%><%mdbor4.Close%><%mdbor5.Close%><%mdbor6.Close%><%mdbor6a.Close%><%mdbor7.Close%>
       
    <%mdbo2.CommandText="SELECT ISNULL(ROUND((SUM(ISNULL(GP.DEBET,0))/1000),0),0) AS summi FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND((Konto<>'18164' AND Konto<>'18151') OR ((Konto='18164' OR Konto='18151') AND SUBSTRING(PROJECT,4,3)='999') OR (Konto='18164' AND (PROJECT='EJE573' OR PROJECT='EJE516' OR PROJECT='EJB515' OR PROJECT='EJB583') AND (MES='504' OR MES='505' OR MES='506'))) AND (KONTO<>'43350' AND SUBKONTO<>'4351') AND m.Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND m.RenovBlock=0 AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND (SUBSTRING(GP.MES, 2, 2) < 10) AND m.OracleCode<>'EJB206' AND (m.IDentIfier = 'C')"%>
    <%mdbor2.Open mdbo2%>
    <%mdbo2u.CommandText="SELECT ISNULL(ROUND((SUM(ISNULL(GP.DEBET,0))/1000),0),0)-ISNULL(ROUND((SUM(ISNULL(GP.CREDIT,0))/1000),0),0) AS summi FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND m.Yearr='" & ya & "' AND m.RenovBlock=0 AND (konto NOT BETWEEN '18410' AND '18433') AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND (SUBSTRING(GP.MES, 2, 2) < 10) AND m.OracleCode='EJB206' AND (m.IDentIfier = 'C')"%>
    <%mdbor2u.Open mdbo2u%>
    <%mdbo3.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.CREDIT,0))/1000,0),0) AS summc FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE SUBSTRING(PROJECT,4,3)<>'999' AND (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') And m.Yearr='" & ya & "' AND m.RenovBlock=0 AND (konto NOT BETWEEN '18410' AND '18433') AND m.OracleCode<>'EJB206' AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND (SUBSTRING(GP.MES, 2, 2) < 10) AND (m.IDentIfier = 'C')"%>
    <%mdbor3.Open mdbo3%>
    <%mdbo3u.CommandText="SELECT ISNULL(ROUND((SUM(ISNULL(GP.DEBET,0))/1000),0),0) AS summc FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE SUBSTRING(PROJECT,4,3)='999' AND (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND m.Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND m.RenovBlock=0 AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND (SUBSTRING(GP.MES, 2, 2) < 10) AND m.OracleCode<>'EJB206' AND (m.IDentIfier = 'C')"%>
    <%mdbor3u.Open mdbo3u%>
    <%mdbo4.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0),0) AS summd FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE m.Yearr='" & ya & "' AND m.RenovBlock=0 AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND (SUBSTRING(GP.MES, 2, 2) < 10) AND (m.IDentIfier = 'C') AND KONTO='43350' AND SUBKONTO='4351'"%>
    <%mdbor4.Open mdbo4%>
    <%mdbo5.CommandText="SELECT DISTINCT ISNULL(SUM(ISNULL(SummaPlan,0)),0) AS SP FROM dbo.Delta WHERE yearr='" & ya & "' AND enn<>'00' AND RenovBlock=0"%>
    <%mdbor5.Open mdbo5%>
    <%mdbo6.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0),0)-ISNULL(ROUND(SUM(ISNULL(GP.CREDIT,0))/1000,0),0) AS summy FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE m.Yearr='" & ya-1 & "' AND (konto BETWEEN '18410' AND '18433') AND m.RenovBlock=0 AND (((SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya-1,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04)) OR (LEFT(MES,1)=" & Mid(ya,4,1) & " AND RIGHT(MES,2) <04))  AND (m.IDentIfier = 'C')"%>
    <%mdbor6.Open mdbo6%>
    <%mdbo6a.CommandText="SELECT DISTINCT ISNULL(SUM(ISNULL(Ettemaks,0)),0) AS EM, ISNULL(SUM(ISNULL(Saldo,0)),0) AS SD FROM dbo.ETTE WHERE yearr='" & ya & "' AND enn<>'00' AND RenovBlock=0"%>
    <%mdbor6a.Open mdbo6a%>
    <%mdbo7.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0),0)-ISNULL(ROUND(SUM(ISNULL(GP.CREDIT,0))/1000,0),0) AS summym FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE m.Yearr='" & ya & "' AND (konto BETWEEN '18410' AND '18433') AND m.RenovBlock=0 AND GP.MES <" & Mid(ya,4,1) & "10 AND (m.IDentIfier = 'C')"%>
    <%mdbor7.Open mdbo7%>
    <tr Class="bold">
     <td colspan=3>Kokku ettev&otilde;tete kaupa, v&auml;lja arvatud plokkide renoveerimine </td>
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
       <%a0="a7b"%>
       <%=Request.Form(a0)%>
      <%Else%>
       <%a0="a7b"%>
       <input Type="Text" value="<%=CDBL(a6)%>" size="10" name="<%="a7b"%>"  Class="bold">
      <%End If%>
     </td>
     <td>
      <%If mdbor7.EOF=True Then%>
       <%a7=0%>
      <%Else%>
       <%a7=mdbor7("Summym")%>
      <%End If%>
      <%If Request.Form("btn")="Kopeerimiseks" Then%>
       <%a0="a7a"%>
       <%=Request.Form(a0)%>
      <%Else%>
       <%a0="a7a"%>
       <input Type="Text" value="<%=CDBL(a7)%>" size="10" name="<%="a7a"%>"  Class="bold">
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
       <%a0="a7c"%>
       <%=Request.Form(a0)%>
      <%Else%>
       <%a0="a7c"%>
       <input Type="Text" value="<%=CDBL(a6)%>" size="10" name="<%="a7c"%>" Class="bold">
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
      <%=CDbl(a6)-CDbl(a3)-CDbl(a9)+CDbl(a1)+CDbl(a2)%>
     </td>
    </tr>
    <%mdbor2.Close%><%mdbor2u.Close%><%mdbor3.Close%><%mdbor3u.Close%><%mdbor4.Close%><%mdbor5.Close%><%mdbor6.Close%><%mdbor6a.Close%><%mdbor7.Close%>
   </Table> 
  </Form>
 </Body>
</Html>