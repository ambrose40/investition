<html>
 <%set servcfg=Server.CreateObject("Scripting.FileSystemObject")%>
 <Head>
  <%b= Server.MapPath("\")%>
  <%if request.Cookies("StyleInv")="" then%>
   <%set servFileStream=servcfg.OpenTextFile(b & "\style.cfg")%>
   <%s=servFileStream.ReadLine%>
   <%servFileStream.Close%>
   <link rel="stylesheet" href="<%=s%>" type="text/css">
  <%else%>
   <%s=request.Cookies("StyleInv")%>
   <link rel="stylesheet" href="<%=s%>" type="text/css">
  <%End if%>
 </Head>
 <body class="report">
  <img border="0" src="icons/report.ico" Style=float:Left><p align="center"><a href="Main.asp"  target="_top" class="Headlink">MAJANDUSAASTA INVESTEERINGUTE KAVA igakuine investeerimine</a></p>
  <%If Request.Form("btn")="OK" Then%>
   <%ya=Request.Form("ye")%>
  <%Else%>
   <%ya=Request.QueryString("ye")%>
  <%End if%>
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
   <%End if%>
  <%End If%>
  <Form Method="POST" Action="Report1.asp?ye=<%=ya%>">
   <Input type="Submit" name="btn" size="10" Value="Kopeerimiseks" class="button">
   <Input type="Submit" name="btn" size="10" Value="Parandamiseks" class="button">
   <%set servFileStream=servcfg.OpenTextFile(b & "\server.cfg")%>
   <%s=servFileStream.ReadLine%>
   <%i=servFileStream.ReadLine%>
   <%p=servFileStream.ReadLine%>
   <%servFileStream.Close%>
   <%set mdbo =  Server.CreateObject("ADODB.Connection")%>
   <%mdbo.ConnectionString="Driver={SQL Server};Server=" & s & ";Database=invest;Trusted_Connection=yes;"%>
   <%mdbo.Open ConnectionString%>
   <%set mdbol1 = Server.CreateObject("ADODB.Command")%>
   <%set mdborl1 = Server.CreateObject("ADODB.Recordset")%>
   <%mdbol1.ActiveConnection = mdbo%>
   <%set mdbol2 = Server.CreateObject("ADODB.Command")%>
   <%set mdborl2 = Server.CreateObject("ADODB.Recordset")%>
   <%mdbol2.ActiveConnection = mdbo%>
   <%set mdbol3 = Server.CreateObject("ADODB.Command")%>
   <%set mdborl3 = Server.CreateObject("ADODB.Recordset")%>
   <%mdbol3.ActiveConnection = mdbo%>
   <%set mdbo1 = Server.CreateObject("ADODB.Command")%>
   <%set mdbor = Server.CreateObject("ADODB.Recordset")%>
   <%mdbo1.ActiveConnection = mdbo%>
   <%set mdbo2 = Server.CreateObject("ADODB.Command")%>
   <%set mdbor2 = Server.CreateObject("ADODB.Recordset")%>
   <%mdbo2.ActiveConnection = mdbo%>
   <%set mdbo2u = Server.CreateObject("ADODB.Command")%>
   <%set mdbor2u = Server.CreateObject("ADODB.Recordset")%>
   <%mdbo2u.ActiveConnection = mdbo%>
   <%set mdbo5 = Server.CreateObject("ADODB.Command")%>
   <%set mdbor5 = Server.CreateObject("ADODB.Recordset")%>
   <%mdbo5.ActiveConnection = mdbo%>
   <%set mdbo6 = Server.CreateObject("ADODB.Command")%>
   <%set mdbor6 = Server.CreateObject("ADODB.Recordset")%>
   <%mdbo6.ActiveConnection = mdbo%>
   <%set mdbo6a = Server.CreateObject("ADODB.Command")%>
   <%set mdbor6a = Server.CreateObject("ADODB.Recordset")%>
   <%mdbo6a.ActiveConnection = mdbo%>
   <%set mdbo7 = Server.CreateObject("ADODB.Command")%>
   <%set mdbor7 = Server.CreateObject("ADODB.Recordset")%>
   <%mdbo7.ActiveConnection = mdbo%>
   <%set mdbosl = Server.CreateObject("ADODB.Command")%>
   <%set mdborsl = Server.CreateObject("ADODB.Recordset")%>
   <%mdbosl.ActiveConnection = mdbo%>

   <table border=1  width="100%">
    <tr>
     <th rowspan="2">Projekti nr.</th>
     <th rowspan="2">Nr</th>
     <th rowspan="2">Projekti nimetus</th>
     <th rowspan="2">Ehitus aastad (m.a.)</th>
     <th rowspan="2">L&otilde;petatud seisuga&nbsp;1.4.<%=ya%></th>
     <th rowspan="2"><%=Mid(ya,3,2)%>&nbspm.a</th>
     <th colspan="15" align="Center">Investeeritud</th> 
    </tr>
    <tr>
     <th>Aprill</th>
     <th>Mai</th>
     <th>Juuni</th>
     <th>Juuli</th>
     <th>August</th>
     <th>September</th>
     <th>Kokku</th>
     <th>Oktoober</th>
     <th>November</th>
     <th>Detsember</th>
     <th>Jaanuar</th>
     <th>Veebruar</th>
     <th>M&auml;rts</th>
     <th>Kokku</th>
     <th>Aastad kokku</th>
    </tr>
    <tr class="repnum">
     <%For nuu=1 to 21%>
      <td><%=nuu%></td>
     <%Next%>
    </tr>

    <%d=Month(Date()) & "." & Day(Date()) & "." & Year(Date())%>
    <%aa=0%><%ab=0%><%ac=0%>
    <%Dim entt(10,23)%>
    <%Dim ent2(10,23)%>

    <%mdbol1.CommandText="SELECT DISTINCT pid, PC, OracleCode, PRojName FROM inpl WHERE IDentifier='C' AND Yearr=" & ya & " AND SUBSTRING(PC,4,2)='00' ORDER BY PC"%>
    <%mdborl1.Open mdbol1%>

    <%mdbo2.CommandText="SELECT ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0) AS summi, GP.MES FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE m.RenovBlock=0 AND Project<>'EJB206' AND (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND ((Konto<>'18164' AND Konto<>'18151') OR ((Konto='18164' OR Konto='18151') AND SUBSTRING(PROJECT,4,3)='999') OR (Konto='18164' AND (PROJECT='EJE573' OR PROJECT='EJE516' OR PROJECT='EJB515' OR PROJECT='EJB583') AND (MES='504' OR MES='505' OR MES='506' OR MES='507'))) AND yearr='" & ya & "' AND (m.IDentifier = 'C') AND (konto NOT BETWEEN '18410' AND '18433') AND KONTO<>'43350' AND SUBKONTO<>'4351' GROUP BY GP.MES HAVING (RIGHT(MES,2)>=04 AND LEFT(MES,1)='" & MID(ya,4,1) & "') OR (RIGHT(MES,2)<04 AND LEFT(MES,1)='" & MID(ya+1,4,1) & "') ORDER BY MES"%>
    <%mdbor2.Open mdbo2%>
    <%mdbo2u.CommandText="SELECT ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0) AS summi, GP.MES FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE m.RenovBlock=0 AND Project='EJB206' AND yearr='" & ya & "' AND (m.IDentifier = 'C') and GP.description<>'maagaas' AND KONTO<>'43350' AND SUBKONTO<>'4351' AND (konto NOT BETWEEN '18410' AND '18433') GROUP BY GP.MES HAVING (RIGHT(MES,2)>=04 AND LEFT(MES,1)='" & MID(ya,4,1) & "') OR (RIGHT(MES,2)<04 AND LEFT(MES,1)='" & MID(ya+1,4,1) & "') ORDER BY MES"%>
    <%mdbor2u.Open mdbo2u%>
    <%mdbo5.CommandText="SELECT DISTINCT SUM(ISNULL(SummaPlan,0)) AS SP,SUM(ISNULL(PastSum,0)) as PastSum FROM dbo.Delta WHERE RenovBlock=0 AND yearr='" & ya & "' AND enn<>'00'"%>
    <%mdbor5.Open mdbo5%>
    <%mdbo6a.CommandText="SELECT DISTINCT SUM(ISNULL(Ettemaks,0)) AS EM,MES FROM dbo.ETE WHERE yearr='" & ya & "' AND enn<>'00' AND RenovBlock=0 AND MES IS NOT NULL group by MES ORDER BY MES"%>
    <%mdbor6a.Open mdbo6a%>
    <%mdbo6.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0),0) AS summy, GP.MES FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE m.Yearr='" & ya & "' AND (konto BETWEEN '18410' AND '18433') AND m.RenovBlock=0 AND (m.IDentifier = 'C') GROUP BY GP.MES HAVING (RIGHT(MES,2)>=04 AND LEFT(MES,1)='" & MID(ya,4,1) & "') OR (RIGHT(MES,2)<04 AND LEFT(MES,1)='" & MID(ya+1,4,1) & "') ORDER BY MES"%>
    <%mdbor6.Open mdbo6%>
    <%mdbo7.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.CREDIT,0))/1000,0),0) AS summy, GP.MES FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE m.Yearr='" & ya & "' AND (konto BETWEEN '18410' AND '18433') AND m.RenovBlock=0 AND (m.IDentifier = 'C') GROUP BY GP.MES HAVING (RIGHT(MES,2)>=04 AND LEFT(MES,1)='" & MID(ya,4,1) & "') OR (RIGHT(MES,2)<04 AND LEFT(MES,1)='" & MID(ya+1,4,1) & "') ORDER BY MES"%>
    <%mdbor7.Open mdbo7%>
    <%suu=0%><%suu2=0%>
    <tr class="boldProjGrup">
     <td colspan=4>INVESTEERINGUD KOKKU v&auml;lja arvatud plokkide renoveerimine</td>
     <td>
      <%If mdbor5.EOF=False Then%>
       <%sim=mdbor5("PastSum")%>
      <%Else%>
       <%sim=0%>
      <%End If%>
      <%If Request.Form("btn")="Kopeerimiseks" Then%>
       <%a0="a1d"%>  
       <%=Request.Form(a0)%>
       <%If Request.Form(a0)="" Then%>
        <input type="hidden" value="<%=Sim%>" name="a1d">
       <%Else%>
        <input type="hidden" value="<%=Request.Form(a0)%>" name="<%="a1d"%>">
       <%End If%>
      <%Else%>
       <%a0="a1d"%>
       <%If Request.Form(a0)="" Then%>
        <input type="Text" value="<%=sim%>" name="<%="a1d"%>" size="10" class="boldProjGrup">
       <%Else%>
        <input type="Text" value="<%=Request.Form(a0)%>" name="<%="a1d"%>" size="10" class="boldProjGrup">
       <%End If%>
      <%End If%>
     </td>
     <td>
      <%If mdbor5.EOF=False Then%>
       <%=mdbor5("SP")%>
      <%End If%>
     </td>
     <%Jj=4%>
     <%For Jj=4 to 9%>
      <td>
       <%If mdbor2.EOF=False Then%>
        <%If Mid(mdbor2("MES"),2,2)="0" & Jj Then%>
         <%a1=mdbor2("Summi")%>
         <%mdbor2.MoveNext%>
        <%Else%>
         <%a1=0%>
        <%End If%>
       <%Else%>
        <%a1=0%>
       <%End If%>       
       <%If mdbor2u.EOF=False Then%>
        <%If Mid(mdbor2u("MES"),2,2)="0" & Jj Then%>
         <%a2=mdbor2u("Summi")%>
         <%mdbor2u.MoveNext%>
        <%Else%>
         <%a2=0%>
        <%End If%>
       <%Else%>
        <%a2=0%>
       <%End If%>
       <%If mdbor7.EOF=False Then%>
        <%If Mid(mdbor7("MES"),2,2)="0" & Jj Then%>
         <%a7=mdbor7("Summy")%>
         <%mdbor7.MoveNext%>
        <%Else%>
         <%a7=0%>
        <%End If%>
       <%Else%>
        <%a7=0%>
       <%End If%>
       <%If mdbor6a.EOF=False Then%>
        <%If Mid(mdbor6a("MES"),2,2)="0" & Jj Then%>
         <%a6=mdbor6a("EM")%>
         <%mdbor6a.MoveNext%>
        <%Else%>
         <%a6=0%>
        <%End If%>
       <%Else%>
        <%a6=0%>
       <%End If%>
       <%If mdbor6.EOF=False Then%>
        <%If Mid(mdbor6("MES"),2,2)="0" & Jj Then%>
         <%a4=mdbor6("Summy")%>
         <%mdbor6.MoveNext%>
        <%Else%>
         <%a4=0%>
        <%End If%>
       <%Else%>
        <%a4=0%>
       <%End If%>

       <%suu=suu+CDbl(a2)+CDBL(a1)-CDBL(a6)+CDBL(a4)-CDBL(a7)%>
       <%=CDbl(a2)+CDBL(a1)-CDBL(a6)+CDBL(a4)-CDBL(a7)%>
      </td>
     <%Next%>
     <td class="thickbord" width="28"><%=suu%></td>
     <%Jj=10%>
     <%For Jj=10 to 12%>
      <td>
       <%If mdbor2.EOF=False Then%>
        <%If Mid(mdbor2("MES"),2,2)=Jj & "" Then%>
         <%a1=mdbor2("Summi")%>
         <%mdbor2.MoveNext%>
        <%Else%>
         <%a1=0%>
        <%End If%>
       <%Else%>
        <%a1=0%>
       <%End If%>
       <%If mdbor2u.EOF=False Then%>
        <%If Mid(mdbor2u("MES"),2,2)=Jj & "" Then%>
         <%a2=mdbor2u("Summi")%>
         <%mdbor2u.MoveNext%>
        <%Else%>
         <%a2=0%>
        <%End If%>
       <%Else%>
        <%a2=0%>
       <%End If%>
       <%If mdbor7.EOF=False Then%>
        <%If Mid(mdbor7("MES"),2,2)=Jj & "" Then%>
         <%a7=mdbor7("Summy")%>
         <%mdbor7.MoveNext%>
        <%Else%>
         <%a7=0%>
        <%End If%>
       <%Else%>
        <%a7=0%>
       <%End If%>
       <%If mdbor6a.EOF=False Then%>
        <%If Mid(mdbor6a("MES"),2,2)=Jj & "" Then%>
         <%a6=mdbor6a("EM")%>
         <%mdbor6a.MoveNext%>
        <%Else%>
         <%a6=0%>
        <%End If%>
       <%Else%>
        <%a6=0%>
       <%End If%>
       <%If mdbor6.EOF=False Then%>
        <%If Mid(mdbor6("MES"),2,2)=Jj & "" Then%>
         <%a4=mdbor6("Summy")%>
         <%mdbor6.MoveNext%>
        <%Else%>
         <%a4=0%>
        <%End If%>
       <%Else%>
        <%a4=0%>
       <%End If%>

       <%suu2=suu2+CDbl(a2)+CDBL(a1)-CDBL(a6)+CDBL(a4)-CDBL(a7)%>
       <%=CDbl(a2)+CDBL(a1)-CDBL(a6)+CDBL(a4)-CDBL(a7)%>
       
      </td>
     <%Next%>
     <%Jj=1%>
     <%For Jj=1 to 3%>
      <td>
       <%If mdbor2.EOF=False Then%>
        <%If Mid(mdbor2("MES"),2,2)="0" & Jj Then%>
         <%a1=mdbor2("Summi")%>
         <%mdbor2.MoveNext%>
        <%Else%>
         <%a1=0%>
        <%End If%>
       <%Else%>
        <%a1=0%>
       <%End If%>
       <%If mdbor2u.EOF=False Then%>
        <%If Mid(mdbor2u("MES"),2,2)="0" & Jj Then%>
         <%a2=mdbor2u("Summi")%>
         <%mdbor2u.MoveNext%>
        <%Else%>
         <%a2=0%>
        <%End If%>
       <%Else%>
        <%a2=0%>
       <%End If%>
       <%If mdbor7.EOF=False Then%>
        <%If Mid(mdbor7("MES"),2,2)="0" & Jj Then%>
         <%a7=mdbor7("Summy")%>
         <%mdbor7.MoveNext%>
        <%Else%>
         <%a7=0%>
        <%End If%>
       <%Else%>
        <%a7=0%>
       <%End If%>
       <%If mdbor6a.EOF=False Then%>
        <%If Mid(mdbor6a("MES"),2,2)="0" & Jj Then%>
         <%a6=mdbor6a("EM")%>
         <%mdbor6a.MoveNext%>
        <%Else%>
         <%a6=0%>
        <%End If%>
       <%Else%>
        <%a6=0%>
       <%End If%>
              <%If mdbor6.EOF=False Then%>
        <%If Mid(mdbor6("MES"),2,2)="0" & Jj Then%>
         <%a4=mdbor6("Summy")%>
         <%mdbor6.MoveNext%>
        <%Else%>
         <%a4=0%>
        <%End If%>
       <%Else%>
        <%a4=0%>
       <%End If%>
       <%suu2=suu2+CDbl(a2)+CDBL(a1)-CDBL(a6)+CDBL(a4)-CDBL(a7)%>
       <%=CDbl(a2)+CDBL(a1)-CDBL(a6)+CDBL(a4)-CDBL(a7)%>
      </td>
     <%Next%>
     <td class="thickbord" width="45">
      <%=suu2%>
     </td>
     <td class="thickbord" width="70">
      <%=suu+suu2%>
     </td>
    </tr>
    <%mdbor6.Close%><%mdbor2.Close%><%mdbor2u.Close%><%mdbor5.Close%><%mdbor6a.Close%><%mdbor7.Close%>
    <%mdbo2.CommandText="SELECT ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0) AS summi, GP.MES FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE Project<>'EJB206' AND (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND ((Konto<>'18164' AND Konto<>'18151') OR ((Konto='18164' OR Konto='18151') AND SUBSTRING(PROJECT,4,3)='999') OR (Konto='18164' AND (PROJECT='EJE573' OR PROJECT='EJE516' OR PROJECT='EJB515' OR PROJECT='EJB583') AND (MES='504' OR MES='505' OR MES='506' OR MES='507'))) AND yearr='" & ya & "' AND (m.IDentifier = 'C') AND (konto NOT BETWEEN '18410' AND '18433') AND KONTO<>'43350' AND SUBKONTO<>'4351' GROUP BY GP.MES HAVING (RIGHT(MES,2)>=04 AND LEFT(MES,1)='" & MID(ya,4,1) & "') OR (RIGHT(MES,2)<04 AND LEFT(MES,1)='" & MID(ya+1,4,1) & "') ORDER BY MES"%>
    <%mdbor2.Open mdbo2%>
    <%mdbo2u.CommandText="SELECT ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0) AS summi, GP.MES FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE Project='EJB206' AND yearr='" & ya & "' AND (m.IDentifier = 'C') and GP.description<>'maagaas' AND KONTO<>'43350' AND SUBKONTO<>'4351' AND (konto NOT BETWEEN '18410' AND '18433') GROUP BY GP.MES HAVING (RIGHT(MES,2)>=04 AND LEFT(MES,1)='" & MID(ya,4,1) & "') OR (RIGHT(MES,2)<04 AND LEFT(MES,1)='" & MID(ya+1,4,1) & "') ORDER BY MES"%>
    <%mdbor2u.Open mdbo2u%>
    <%mdbo5.CommandText="SELECT DISTINCT SUM(ISNULL(SummaPlan,0)) AS SP,SUM(ISNULL(PastSum,0)) as PastSum FROM dbo.Delta WHERE yearr='" & ya & "' AND enn<>'00'"%>
    <%mdbor5.Open mdbo5%>
    <%mdbo6a.CommandText="SELECT DISTINCT SUM(ISNULL(Ettemaks,0)) AS EM,MES FROM dbo.ETE WHERE yearr='" & ya & "' AND enn<>'00' AND MES IS NOT NULL group by MES ORDER BY MES"%>
    <%mdbor6a.Open mdbo6a%>
    <%mdbo6.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0),0) AS summy, GP.MES FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE m.Yearr='" & ya & "' AND (konto BETWEEN '18410' AND '18433') AND (m.IDentifier = 'C') GROUP BY GP.MES HAVING (RIGHT(MES,2)>=04 AND LEFT(MES,1)='" & MID(ya,4,1) & "') OR (RIGHT(MES,2)<04 AND LEFT(MES,1)='" & MID(ya+1,4,1) & "') ORDER BY MES"%>
    <%mdbor6.Open mdbo6%>
    <%mdbo7.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.CREDIT,0))/1000,0),0) AS summy, GP.MES FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE m.Yearr='" & ya & "' AND (konto BETWEEN '18410' AND '18433') AND (m.IDentifier = 'C') GROUP BY GP.MES HAVING (RIGHT(MES,2)>=04 AND LEFT(MES,1)='" & MID(ya,4,1) & "') OR (RIGHT(MES,2)<04 AND LEFT(MES,1)='" & MID(ya+1,4,1) & "') ORDER BY MES"%>
    <%mdbor7.Open mdbo7%>
    <%suu=0%>
    <%suu2=0%>
    <tr class="boldProjGrup">
     <td colspan=4>IVESTEERINGUD KOKKU koos plokkide renoveerimisega</td>
     <td>
      <%If mdbor5.EOF=False Then%>
       <%sim=mdbor5("PastSum")%>
      <%Else%>
       <%sim=0%>
      <%End If%>
      <%If Request.Form("btn")="Kopeerimiseks" Then%>
       <%a0="ad"%>
       <%=Request.Form(a0)%>
       <%If Request.Form(a0)="" Then%>
        <input type="hidden" value="<%=Sim%>" name="ad">
       <%Else%>
        <input type="hidden" value="<%=Request.Form(a0)%>" name="<%="ad"%>">
       <%End If%>
      <%Else%>
       <%a0="ad"%>
       <%If Request.Form(a0)="" Then%>
        <input type="Text" value="<%=sim%>" name="<%="ad"%>" size="10" class="boldProjGrup">
       <%Else%>
        <input type="Text" value="<%=Request.Form(a0)%>" name="<%="ad"%>" size="10" class="boldProjGrup">
       <%End If%>
      <%End If%>
     </td>
     <td>
      <%If mdbor5.EOF=False Then%>
       <%=mdbor5("SP")%>
      <%End If%>
     </td>
     <%Jj=4%>
     <%For Jj=4 to 9%>
      <td>
       <%If mdbor2.EOF=False Then%>
        <%If Mid(mdbor2("MES"),2,2)="0" & Jj Then%>
         <%a1=mdbor2("Summi")%>
         <%mdbor2.MoveNext%>
        <%Else%>
         <%a1=0%>
        <%End If%>
       <%Else%>
        <%a1=0%>
       <%End If%>
       <%If mdbor2u.EOF=False Then%>
        <%If Mid(mdbor2u("MES"),2,2)="0" & Jj Then%>
         <%a2=mdbor2u("Summi")%>
         <%mdbor2u.MoveNext%>
        <%Else%>
         <%a2=0%>
        <%End If%>
       <%Else%>
        <%a2=0%>
       <%End If%>
       <%If mdbor7.EOF=False Then%>
        <%If Mid(mdbor7("MES"),2,2)="0" & Jj Then%>
         <%a7=mdbor7("Summy")%>
         <%mdbor7.MoveNext%>
        <%Else%>
         <%a7=0%>
        <%End If%>
       <%Else%>
        <%a7=0%>
       <%End If%>
       <%If mdbor6a.EOF=False Then%>
        <%If Mid(mdbor6a("MES"),2,2)="0" & Jj Then%>
         <%a6=mdbor6a("EM")%>
         <%mdbor6a.MoveNext%>
        <%Else%>
         <%a6=0%>
        <%End If%>
       <%Else%>
        <%a6=0%>
       <%End If%>       
       <%If mdbor6.EOF=False Then%>
        <%If Mid(mdbor6("MES"),2,2)="0" & Jj Then%>
         <%a4=mdbor6("Summy")%>
         <%mdbor6.MoveNext%>
        <%Else%>
         <%a4=0%>
        <%End If%>
       <%Else%>
        <%a4=0%>
       <%End If%>

       <%suu=suu+CDbl(a2)+CDBL(a1)-CDBL(a6)+CDBL(a4)-CDBL(a7)%>
       <%=CDbl(a2)+CDBL(a1)-CDBL(a6)+CDBL(a4)-CDBL(a7)%>
      </td>
     <%Next%>
     <td class="thickbord" width="28">
      <%=suu%>
     </td>
     <%Jj=10%>
     <%For Jj=10 to 12%>
      <td>
       <%If mdbor2.EOF=False Then%>
        <%If Mid(mdbor2("MES"),2,2)=Jj & "" Then%>
         <%a1=mdbor2("Summi")%>
         <%mdbor2.MoveNext%>
        <%Else%>
         <%a1=0%>
        <%End If%>
       <%Else%>
        <%a1=0%>
       <%End If%>
       <%If mdbor2u.EOF=False Then%>
        <%If Mid(mdbor2u("MES"),2,2)=Jj & "" Then%>
         <%a2=mdbor2u("Summi")%>
         <%mdbor2u.MoveNext%>
        <%Else%>
         <%a2=0%>
        <%End If%>
       <%Else%>
        <%a2=0%>
       <%End If%>
       <%If mdbor7.EOF=False Then%>
        <%If Mid(mdbor7("MES"),2,2)=Jj & "" Then%>
         <%a7=mdbor7("Summy")%>
         <%mdbor7.MoveNext%>
        <%Else%>
         <%a7=0%>
        <%End If%>
       <%Else%>
        <%a7=0%>
       <%End If%>
       <%If mdbor6a.EOF=False Then%>
        <%If Mid(mdbor6a("MES"),2,2)=Jj & "" Then%>
         <%a6=mdbor6a("EM")%>
         <%mdbor6a.MoveNext%>
        <%Else%>
         <%a6=0%>
        <%End If%>
       <%Else%>
        <%a6=0%>
       <%End If%>
       <%If mdbor6.EOF=False Then%>
        <%If Mid(mdbor6("MES"),2,2)=Jj & "" Then%>
         <%a4=mdbor6("Summy")%>
         <%mdbor6.MoveNext%>
        <%Else%>
         <%a4=0%>
        <%End If%>
       <%Else%>
        <%a4=0%>
       <%End If%>

       <%suu2=suu2+CDbl(a2)+CDBL(a1)-CDBL(a6)+CDBL(a4)-CDBL(a7)%>
       <%=CDbl(a2)+CDBL(a1)-CDBL(a6)+CDBL(a4)-CDBL(a7)%>
      </td>
     <%Next%>
     <%Jj=1%>
     <%For Jj=1 to 3%>
      <td>
       <%If mdbor2.EOF=False Then%>
        <%If Mid(mdbor2("MES"),2,2)="0" & Jj Then%>
         <%a1=mdbor2("Summi")%>
         <%mdbor2.MoveNext%>
        <%Else%>
         <%a1=0%>
        <%End If%>
       <%Else%>
        <%a1=0%>
       <%End If%>
       <%If mdbor2u.EOF=False Then%>
        <%If Mid(mdbor2u("MES"),2,2)="0" & Jj Then%>
         <%a2=mdbor2u("Summi")%>
         <%mdbor2u.MoveNext%>
        <%Else%>
         <%a2=0%>
        <%End If%>
       <%Else%>
        <%a2=0%>
       <%End If%>
       <%If mdbor7.EOF=False Then%>
        <%If Mid(mdbor7("MES"),2,2)="0" & Jj Then%>
         <%a7=mdbor7("Summy")%>
         <%mdbor7.MoveNext%>
        <%Else%>
         <%a7=0%>
        <%End If%>
       <%Else%>
        <%a7=0%>
       <%End If%>
       <%If mdbor6a.EOF=False Then%>
        <%If Mid(mdbor6a("MES"),2,2)="0" & Jj Then%>
         <%a6=mdbor6a("EM")%>
         <%mdbor6a.MoveNext%>
        <%Else%>
         <%a6=0%>
        <%End If%>
       <%Else%>
        <%a6=0%>
       <%End If%>
       <%If mdbor6.EOF=False Then%>
        <%If Mid(mdbor6("MES"),2,2)="0" & Jj Then%>
         <%a4=mdbor6("Summy")%>
         <%mdbor6.MoveNext%>
        <%Else%>
         <%a4=0%>
        <%End If%>
       <%Else%>
        <%a4=0%>
       <%End If%>

       <%suu2=suu2+CDbl(a2)+CDBL(a1)-CDBL(a6)+CDBL(a4)-CDBL(a7)%>
       <%=CDbl(a2)+CDBL(a1)-CDBL(a6)+CDBL(a4)-CDBL(a7)%>      </td>
     <%Next%>
     <td class="thickbord" width="45">
      <%=suu2%>
     </td>
     <td class="thickbord" width="70">
      <%=suu+suu2%>
     </td>
    </tr>

    <%Do until mdborl1.EOF%>
     <%mdbor2.Close%><%mdbor6.Close%><%mdbor2u.Close%><%mdbor5.Close%><%'mdbor6.Close%><%mdbor6a.Close%><%mdbor7.Close%>
     <%mdbo2.CommandText="SELECT ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0) AS summi,SUBSTRING(m.ProjCode, 1, 2) AS be, GP.MES FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE SUBSTRING(m.ProjCode,1,2)='" & MID(Mdborl1("PC"),1,2) & "' AND project<>'EJB206' AND (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND ((Konto<>'18164' AND Konto<>'18151') OR ((Konto='18164' OR Konto='18151') AND SUBSTRING(PROJECT,4,3)='999') OR (Konto='18164' AND (PROJECT='EJE573' OR PROJECT='EJE516' OR PROJECT='EJB515' OR PROJECT='EJB583') AND (MES='504' OR MES='505' OR MES='506' OR MES='507'))) AND yearr='" & ya & "' AND (m.IDentifier = 'C') AND (konto NOT BETWEEN '18410' AND '18433') AND KONTO<>'43350' AND SUBKONTO<>'4351' GROUP BY SUBSTRING(m.ProjCode, 1, 2),GP.MES HAVING (RIGHT(MES,2)>=04 AND LEFT(MES,1)='" & MID(ya,4,1) & "') OR (RIGHT(MES,2)<04 AND LEFT(MES,1)='" & MID(ya+1,4,1) & "') ORDER BY MES"%>
     <%mdbor2.Open mdbo2%>
     <%mdbo2u.CommandText="SELECT ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0) AS summi,SUBSTRING(m.ProjCode, 1, 2) AS be, GP.MES FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE SUBSTRING(m.ProjCode,1,2)='" & MID(Mdborl1("PC"),1,2) & "' AND Project='EJB206' AND yearr='" & ya & "' AND (m.IDentifier = 'C') and GP.description<>'maagaas' AND KONTO<>'43350' AND SUBKONTO<>'4351' AND (konto NOT BETWEEN '18410' AND '18433') GROUP BY SUBSTRING(m.ProjCode, 1, 2),GP.MES HAVING (RIGHT(MES,2)>=04 AND LEFT(MES,1)='" & MID(ya,4,1) & "') OR (RIGHT(MES,2)<04 AND LEFT(MES,1)='" & MID(ya+1,4,1) & "') ORDER BY MES"%>
     <%mdbor2u.Open mdbo2u%>
     <%mdbo5.CommandText="SELECT DISTINCT SUM(ISNULL(SummaPlan,0)) AS SP, SUM(ISNULL(PastSum,0)) AS PastSum, SUM(ISNULL(SummaContract,0)) AS SC, be FROM dbo.Delta WHERE be='" & MID(mdborl1("PC"),1,2) & "' AND yearr='" & ya & "' AND enn<>'00' GROUP BY be"%>
     <%mdbor5.Open mdbo5%>
     <%mdbo6a.CommandText="SELECT DISTINCT SUM(ISNULL(Ettemaks,0)) AS EM,MES, be FROM dbo.ETE WHERE be='" & MID(mdborl1("PC"),1,2) & "' AND yearr='" & ya & "' AND enn<>'00' AND MES IS NOT NULL group by be,MES ORDER BY MES"%>
     <%mdbor6a.Open mdbo6a%>
     <%mdbo6.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0),0) AS summy,SUBSTRING(m.ProjCode, 1, 2) as be, GP.MES FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE SUBSTRING(m.ProjCode,1,2)='" & MID(Mdborl1("PC"),1,2) & "' AND m.Yearr='" & ya & "' AND (konto BETWEEN '18410' AND '18433') AND (m.IDentifier = 'C') GROUP BY SUBSTRING(m.ProjCode, 1, 2), GP.MES HAVING (RIGHT(MES,2)>=04 AND LEFT(MES,1)='" & MID(ya,4,1) & "') OR (RIGHT(MES,2)<04 AND LEFT(MES,1)='" & MID(ya+1,4,1) & "') ORDER BY MES"%>
     <%mdbor6.Open mdbo6%>
     <%mdbo7.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.CREDIT,0))/1000,0),0) AS summy,SUBSTRING(m.ProjCode, 1, 2) as be, GP.MES FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE SUBSTRING(m.ProjCode,1,2)='" & MID(Mdborl1("PC"),1,2) & "' AND m.Yearr='" & ya & "' AND (konto BETWEEN '18410' AND '18433') AND (m.IDentifier = 'C') GROUP BY SUBSTRING(m.ProjCode, 1, 2), GP.MES HAVING (RIGHT(MES,2)>=04 AND LEFT(MES,1)='" & MID(ya,4,1) & "') OR (RIGHT(MES,2)<04 AND LEFT(MES,1)='" & MID(ya+1,4,1) & "') ORDER BY MES"%>
     <%mdbor7.Open mdbo7%>
     <%suu=0%><%suu2=0%>
     <tr class="ProjGrup">
      <td>
       <%a=MID(mdborl1("PC"),1,3)%>
        <%=REPLACE(a, "0", "")%>
      </td>
      <td colspan=3>
       <%=mdborl1("ProjName")%>
      </td>
      <td>
       <%If mdbor5.EOF=False Then%>
        <%sim=mdbor5("PastSum")%>
       <%Else%>
        <%sim=0%>
       <%End If%>
       <%If Request.Form("btn")="Kopeerimiseks" Then%>
        <%a0="ad" & mdborl1("PC")%>
        <%=Request.Form(a0)%>
        <%If Request.Form(a0)="" Then%>
         <input type="hidden" value="<%=Sim%>" name="<%="ad" & mdborl1("PC")%>">
        <%Else%>
         <input type="hidden" value="<%=Request.Form(a0)%>" name="<%="ad" & mdborl1("PC")%>">
        <%End If%>
       <%Else%>
        <%a0="ad" & mdborl1("PC")%>
        <%If Request.Form(a0)="" Then%>
         <input type="Text" value="<%=sim%>" name="<%="ad" & mdborl1("PC")%>" size="10" class="ProjGrup">
        <%Else%>
         <input type="Text" value="<%=Request.Form(a0)%>" name="<%="ad" & mdborl1("PC")%>" size="10" class="Projgrup">
        <%End If%>
       <%End If%>
      </td>
      <td>
       <%If mdbor5.EOF=False Then%>
        <%=mdbor5("SP")%>
       <%End If%>
      </td>
      <%Jj=4%>
      <%For Jj=4 to 9%>
       <td>
        <%If mdbor2.EOF=False Then%>
         <%If Mid(mdbor2("MES"),2,2)="0" & Jj Then%>
          <%a1=mdbor2("Summi")%>
          <%mdbor2.MoveNext%>
         <%Else%>
          <%a1=0%>
         <%End If%>
        <%Else%>
         <%a1=0%>
        <%End If%>
        <%If mdbor2u.EOF=False Then%>
         <%If Mid(mdbor2u("MES"),2,2)="0" & Jj Then%>
          <%a2=mdbor2u("Summi")%>
          <%mdbor2u.MoveNext%>
         <%Else%>
          <%a2=0%>
         <%End If%>
        <%Else%>
         <%a2=0%>
        <%End If%>
        <%If mdbor7.EOF=False Then%>
         <%If Mid(mdbor7("MES"),2,2)="0" & Jj Then%>
          <%a7=mdbor7("Summy")%>
          <%mdbor7.MoveNext%>
         <%Else%>
          <%a7=0%>
         <%End If%>
        <%Else%>
         <%a7=0%>
        <%End If%>
        <%If mdbor6a.EOF=False Then%>
         <%If Mid(mdbor6a("MES"),2,2)="0" & Jj Then%>
          <%a6=mdbor6a("EM")%>
          <%mdbor6a.MoveNext%>
         <%Else%>
          <%a6=0%>
         <%End If%>
        <%Else%>
         <%a6=0%>
        <%End If%>
       <%If mdbor6.EOF=False Then%>
        <%If Mid(mdbor6("MES"),2,2)="0" & Jj Then%>
         <%a4=mdbor6("Summy")%>
         <%mdbor6.MoveNext%>
        <%Else%>
         <%a4=0%>
        <%End If%>
       <%Else%>
        <%a4=0%>
       <%End If%>

       <%suu=suu+CDbl(a2)+CDBL(a1)-CDBL(a6)+CDBL(a4)-CDBL(a7)%>
       <%=CDbl(a2)+CDBL(a1)-CDBL(a6)+CDBL(a4)-CDBL(a7)%>       </td>
      <%Next%>
      <td class="thickbord" width="28">
       <%=suu%>
      </td>
      <%Jj=10%>
      <%For Jj=10 to 12%>
       <td>
        <%If mdbor2.EOF=False Then%>
         <%If Mid(mdbor2("MES"),2,2)=Jj & "" Then%>
          <%a1=mdbor2("Summi")%>
          <%mdbor2.MoveNext%>
         <%Else%>
          <%a1=0%>
         <%End If%>
        <%Else%>
         <%a1=0%>
        <%End If%>
        <%If mdbor2u.EOF=False Then%>
         <%If Mid(mdbor2u("MES"),2,2)=Jj & "" Then%>
          <%a2=mdbor2u("Summi")%>
          <%mdbor2u.MoveNext%>
         <%Else%>
          <%a2=0%>
         <%End If%>
        <%Else%>
         <%a2=0%>
        <%End If%>
        <%If mdbor7.EOF=False Then%>
         <%If Mid(mdbor7("MES"),2,2)=Jj & "" Then%>
          <%a7=mdbor7("Summy")%>
          <%mdbor7.MoveNext%>
         <%Else%>
          <%a7=0%>
         <%End If%>
        <%Else%>
         <%a7=0%>
        <%End If%>
        <%If mdbor6a.EOF=False Then%>
         <%If Mid(mdbor6a("MES"),2,2)=Jj & "" Then%>
          <%a6=mdbor6a("EM")%>
          <%mdbor6a.MoveNext%>
         <%Else%>
          <%a6=0%>
         <%End If%>
        <%Else%>
         <%a6=0%>
        <%End If%>
        <%If mdbor6.EOF=False Then%>
         <%If Mid(mdbor6("MES"),2,2)=Jj & "" Then%>
          <%a4=mdbor6("Summy")%>
          <%mdbor6.MoveNext%>
         <%Else%>
          <%a4=0%>
         <%End If%>
        <%Else%>
         <%a4=0%>
        <%End If%>
        <%suu2=suu2+CDbl(a2)+CDBL(a1)-CDBL(a6)+CDBL(a4)-CDBL(a7)%>
        <%=CDbl(a2)+CDBL(a1)-CDBL(a6)+CDBL(a4)-CDBL(a7)%>       
       </td>
      <%Next%>
      <%Jj=1%>
      <%For Jj=1 to 3%>
       <td>
        <%If mdbor2.EOF=False Then%>
         <%If Mid(mdbor2("MES"),2,2)="0" & Jj Then%>
          <%a1=mdbor2("Summi")%>
          <%mdbor2.MoveNext%>
         <%Else%>
          <%a1=0%>
         <%End If%>
        <%Else%>
         <%a1=0%>
        <%End If%>
        <%If mdbor2u.EOF=False Then%>
         <%If Mid(mdbor2u("MES"),2,2)="0" & Jj Then%>
          <%a2=mdbor2u("Summi")%>
          <%mdbor2u.MoveNext%>
         <%Else%>
          <%a2=0%>
         <%End If%>
        <%Else%>
         <%a2=0%>
        <%End If%>
        <%If mdbor7.EOF=False Then%>
         <%If Mid(mdbor7("MES"),2,2)="0" & Jj Then%>
          <%a7=mdbor7("Summy")%>
          <%mdbor7.MoveNext%>
         <%Else%>
          <%a7=0%>
         <%End If%>
        <%Else%>
         <%a7=0%>
        <%End If%>
        <%If mdbor6a.EOF=False Then%>
         <%If Mid(mdbor6a("MES"),2,2)="0" & Jj Then%>
          <%a6=mdbor6a("EM")%>
          <%mdbor6a.MoveNext%>
         <%Else%>
          <%a6=0%>
         <%End If%>
        <%Else%>
         <%a6=0%>
        <%End If%>
        <%If mdbor6.EOF=False Then%>
         <%If Mid(mdbor6("MES"),2,2)="0" & Jj Then%>
          <%a4=mdbor6("Summy")%>
          <%mdbor6.MoveNext%>
         <%Else%>
          <%a4=0%>
         <%End If%>
        <%Else%>
         <%a4=0%>
        <%End If%>
        <%suu2=suu2+CDbl(a2)+CDBL(a1)-CDBL(a6)+CDBL(a4)-CDBL(a7)%>
        <%=CDbl(a2)+CDBL(a1)-CDBL(a6)+CDBL(a4)-CDBL(a7)%>
       </td>
      <%Next%>
      <td class="thickbord" width="45"><%=suu2%></td>
      <td  class="thickbord" width="70"><%=suu+suu2%></td>
     </tr>
        
     <%mdbol2.CommandText="SELECT DISTINCT PC,ProjName,OracleCode FROM inpl WHERE IDentifier='C' AND Yearr=" & ya & " AND SUBSTRING(PC,4,2)<>'00' AND SUBSTRING(PC,7,2)='00' AND  SUBSTRING(PC,1,2)='" & MID(mdborl1("PC"),1,2) & "' ORDER BY PC"%>
     <%mdborl2.Open mdbol2%>
        
     <%Do until mdborl2.EOF%>
      <%mdbor2.Close%><%mdbor6.Close%><%mdbor2u.Close%><%mdbor5.Close%><%mdbor6a.Close%><%mdbor7.Close%>

      <%mdbo2.CommandText="SELECT ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0) AS summi,SUBSTRING(m.ProjCode, 1, 2) AS be, SUBSTRING(m.ProjCode, 4, 2) AS mi, GP.MES FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE SUBSTRING(m.ProjCode,1,5)='" & MID(Mdborl2("PC"),1,5) & "' AND OracleCOde<>'EJB206' AND (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND ((Konto<>'18164' AND Konto<>'18151') OR ((Konto='18164' OR Konto='18151') AND SUBSTRING(PROJECT,4,3)='999') OR (Konto='18164' AND (PROJECT='EJE573' OR PROJECT='EJE516' OR PROJECT='EJB515' OR PROJECT='EJB583') AND (MES='504' OR MES='505' OR MES='506' OR MES='507'))) AND yearr='" & ya & "' AND (m.IDentifier = 'C') AND KONTO<>'43350' AND SUBKONTO<>'4351' AND (konto NOT BETWEEN '18410' AND '18433') GROUP BY SUBSTRING(m.ProjCode, 1, 2), SUBSTRING(m.ProjCode, 4, 2),GP.MES HAVING (RIGHT(MES,2)>=04 AND LEFT(MES,1)='" & MID(ya,4,1) & "') OR (RIGHT(MES,2)<04 AND LEFT(MES,1)='" & MID(ya+1,4,1) & "') ORDER BY MES,be,mi"%>
      <%mdbor2.Open mdbo2%>
      <%mdbo2u.CommandText="SELECT ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0) AS summi,SUBSTRING(m.ProjCode, 1, 2) AS be, SUBSTRING(m.ProjCode, 4, 2) AS mi, GP.MES FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE SUBSTRING(m.ProjCode,1,5)='" & MID(Mdborl2("PC"),1,5) & "' AND Project='EJB206' AND yearr='" & ya & "' AND (m.IDentifier = 'C') and GP.description<>'maagaas' AND KONTO<>'43350' AND SUBKONTO<>'4351' AND (konto NOT BETWEEN '18410' AND '18433') GROUP BY SUBSTRING(m.ProjCode, 1, 2), SUBSTRING(m.ProjCode, 4, 2), GP.MES HAVING (RIGHT(MES,2)>=04 AND LEFT(MES,1)='" & MID(ya,4,1) & "') OR (RIGHT(MES,2)<04 AND LEFT(MES,1)='" & MID(ya+1,4,1) & "') ORDER BY MES,be,mi"%>
      <%mdbor2u.Open mdbo2u%>
      <%mdbo5.CommandText="SELECT DISTINCT SUM(ISNULL(SummaPlan,0)) AS SP, SUM(ISNULL(PastSum,0)) AS PastSum, SUM(ISNULL(SummaContract,0)) AS SC, be, mi FROM dbo.Delta WHERE mi = '" & Mid(mdborl2("PC"),4,2) & "' AND be='" & MID(mdborl1("PC"),1,2) & "' AND yearr='" & ya & "' AND enn<>'00' GROUP BY be, mi"%>
      <%mdbor5.Open mdbo5%>
      <%mdbo6a.CommandText="SELECT DISTINCT SUM(ISNULL(Ettemaks,0)) AS EM,MES, be, mi FROM dbo.ETE WHERE mi = '" & Mid(mdborl2("PC"),4,2) & "' AND be='" & MID(mdborl1("PC"),1,2) & "' AND yearr='" & ya & "' AND enn<>'00' AND MES IS NOT NULL group by be,mi,MES ORDER BY MES"%>
      <%mdbor6a.Open mdbo6a%>
      <%mdbo6.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0),0) AS summy,SUBSTRING(m.ProjCode, 1, 2) as be,SUBSTRING(m.ProjCode, 4, 2) AS mi, GP.MES FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE SUBSTRING(m.ProjCode,1,5)='" & MID(Mdborl2("PC"),1,5) & "' AND m.Yearr='" & ya & "' AND (konto BETWEEN '18410' AND '18433') AND (m.IDentifier = 'C') GROUP BY SUBSTRING(m.ProjCode, 1, 2),SUBSTRING(m.ProjCode, 4, 2), GP.MES HAVING (RIGHT(MES,2)>=04 AND LEFT(MES,1)='" & MID(ya,4,1) & "') OR (RIGHT(MES,2)<04 AND LEFT(MES,1)='" & MID(ya+1,4,1) & "') ORDER BY MES"%>
      <%mdbor6.Open mdbo6%>
      <%mdbo7.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.CREDIT,0))/1000,0),0) AS summy,SUBSTRING(m.ProjCode, 1, 2) as be,SUBSTRING(m.ProjCode, 4, 2) AS mi, GP.MES FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE SUBSTRING(m.ProjCode,1,5)='" & MID(Mdborl2("PC"),1,5) & "' AND m.Yearr='" & ya & "' AND (konto BETWEEN '18410' AND '18433') AND (m.IDentifier = 'C') GROUP BY SUBSTRING(m.ProjCode, 1, 2),SUBSTRING(m.ProjCode, 4, 2), GP.MES HAVING (RIGHT(MES,2)>=04 AND LEFT(MES,1)='" & MID(ya,4,1) & "') OR (RIGHT(MES,2)<04 AND LEFT(MES,1)='" & MID(ya+1,4,1) & "') ORDER BY MES"%>
      <%mdbor7.Open mdbo7%>
      <%suu=0%><%suu2=0%>
      <tr class="ProjGrup">
       <td>
        <%a=MID(mdborl2("PC"),1,6)%>
        <%=REPLACE(a, "0", "")%>
       </td>
       <td colspan=3>
        <%=mdborl2("ProjName")%>
       </td>
       <td>
        <%If mdbor5.EOF=False Then%>
         <%sim=mdbor5("PastSum")%>
        <%Else%>
         <%sim=0%>
        <%End If%>
        <%If Request.Form("btn")="Kopeerimiseks" Then%>
         <%a0="ad" & mdborl2("PC")%>
         <%=Request.Form(a0)%>
         <%If Request.Form(a0)="" Then%>
          <input type="hidden" value="<%=Sim%>" name="<%="ad" & mdborl2("PC")%>">
         <%Else%>
          <input type="hidden" value="<%=Request.Form(a0)%>" name="<%="ad" & mdborl2("PC")%>">
         <%End If%>
        <%Else%>
         <%a0="ad" & mdborl2("PC")%>
         <%If Request.Form(a0)="" Then%>
          <input type="Text" value="<%=sim%>" name="<%="ad" & mdborl2("PC")%>" size="10" class="ProJGrup">
         <%Else%>
          <input type="Text" value="<%=Request.Form(a0)%>" name="<%="ad" & mdborl2("PC")%>" size="10" class="ProjGrup">
         <%End If%>
        <%End If%>
       </td>
       <td>
        <%If mdbor5.EOF=False Then%>
         <%=mdbor5("SP")%>
        <%End If%>
       </td>
       <%Jj=4%>
       <%For Jj=4 to 9%>
        <td>
         <%If mdbor2.EOF=False Then%>
          <%If Mid(mdbor2("MES"),2,2)="0" & Jj Then%>
           <%a1=mdbor2("Summi")%>
           <%mdbor2.MoveNext%>
          <%Else%>
           <%a1=0%>
          <%End If%>
         <%Else%>
          <%a1=0%>
         <%End If%>
         <%If mdbor2u.EOF=False Then%>
          <%If Mid(mdbor2u("MES"),2,2)="0" & Jj Then%>
           <%a2=mdbor2u("Summi")%>
           <%mdbor2u.MoveNext%>
          <%Else%>
           <%a2=0%>
          <%End If%>
         <%Else%>
          <%a2=0%>
         <%End If%>
         <%If mdbor7.EOF=False Then%>
          <%If Mid(mdbor7("MES"),2,2)="0" & Jj Then%>
           <%a7=mdbor7("Summy")%>
           <%mdbor7.MoveNext%>
          <%Else%>
           <%a7=0%>
          <%End If%>
         <%Else%>
          <%a7=0%>
         <%End If%>
         <%If mdbor6a.EOF=False Then%>
          <%If Mid(mdbor6a("MES"),2,2)="0" & Jj Then%>
           <%a6=mdbor6a("EM")%>
           <%mdbor6a.MoveNext%>
          <%Else%>
           <%a6=0%>
          <%End If%>
         <%Else%>
          <%a6=0%>
         <%End If%>
         <%If mdbor6.EOF=False Then%>
          <%If Mid(mdbor6("MES"),2,2)="0" & Jj Then%>
           <%a4=mdbor6("Summy")%>
           <%mdbor6.MoveNext%>
          <%Else%>
           <%a4=0%>
          <%End If%>
         <%Else%>
          <%a4=0%>
         <%End If%>
         <%suu=suu+CDbl(a2)+CDBL(a1)-CDBL(a6)+CDBL(a4)-CDBL(a7)%>
         <%=CDbl(a2)+CDBL(a1)-CDBL(a6)+CDBL(a4)-CDBL(a7)%>
        </td>
       <%Next%>
       <td class="thickbord" width="28">
        <%=suu%>
       </td>
       <%Jj=10%>
       <%For Jj=10 to 12%>
        <td>
         <%If mdbor2.EOF=False Then%>
          <%If Mid(mdbor2("MES"),2,2)=Jj & "" Then%>
           <%a1=mdbor2("Summi")%>
           <%mdbor2.MoveNext%>
          <%Else%>
           <%a1=0%>
          <%End If%>
         <%Else%>
          <%a1=0%>
         <%End If%>
         <%If mdbor2u.EOF=False Then%>
          <%If Mid(mdbor2u("MES"),2,2)=Jj & "" Then%>
           <%a2=mdbor2u("Summi")%>
           <%mdbor2u.MoveNext%>
          <%Else%>
           <%a2=0%>
          <%End If%>
         <%Else%>
          <%a2=0%>
         <%End If%>
         <%If mdbor7.EOF=False Then%>
          <%If Mid(mdbor7("MES"),2,2)=Jj & "" Then%>
           <%a7=mdbor7("Summy")%>
           <%mdbor7.MoveNext%>
          <%Else%>
           <%a7=0%>
          <%End If%>
         <%Else%>
          <%a7=0%>
         <%End If%>
         <%If mdbor6a.EOF=False Then%>
          <%If Mid(mdbor6a("MES"),2,2)=Jj & "" Then%>
           <%a6=mdbor6a("EM")%>
           <%mdbor6a.MoveNext%>
          <%Else%>
           <%a6=0%>
          <%End If%>
         <%Else%>
          <%a6=0%>
         <%End If%>
         <%If mdbor6.EOF=False Then%>
          <%If Mid(mdbor6("MES"),2,2)=Jj & "" Then%>
           <%a4=mdbor6("Summy")%>
           <%mdbor6.MoveNext%>
          <%Else%>
           <%a4=0%>
          <%End If%>
         <%Else%>
          <%a4=0%>
         <%End If%>
         <%suu2=suu2+CDbl(a2)+CDBL(a1)-CDBL(a6)+CDBL(a4)-CDBL(a7)%>
         <%=CDbl(a2)+CDBL(a1)-CDBL(a6)+CDBL(a4)-CDBL(a7)%>
        </td>
       <%Next%>
       <%Jj=1%>
       <%For Jj=1 to 3%>
        <td>
         <%If mdbor2.EOF=False Then%>
          <%If Mid(mdbor2("MES"),2,2)="0" & Jj Then%>
           <%a1=mdbor2("Summi")%>
           <%mdbor2.MoveNext%>
          <%Else%>
           <%a1=0%>
          <%End If%>
         <%Else%>
          <%a1=0%>
         <%End If%>
         <%If mdbor2u.EOF=False Then%>
          <%If Mid(mdbor2u("MES"),2,2)="0" & Jj Then%>
           <%a2=mdbor2u("Summi")%>
           <%mdbor2u.MoveNext%>
          <%Else%>
           <%a2=0%>
          <%End If%>
         <%Else%>
          <%a2=0%>
         <%End If%>
         <%If mdbor7.EOF=False Then%>
          <%If Mid(mdbor7("MES"),2,2)="0" & Jj Then%>
           <%a7=mdbor7("Summy")%>
           <%mdbor7.MoveNext%>
          <%Else%>
           <%a7=0%>
          <%End If%>
         <%Else%>
          <%a7=0%>
         <%End If%>
         <%If mdbor6a.EOF=False Then%>
          <%If Mid(mdbor6a("MES"),2,2)="0" & Jj Then%>
           <%a6=mdbor6a("EM")%>
           <%mdbor6a.MoveNext%>
          <%Else%>
           <%a6=0%>
          <%End If%>
         <%Else%>
          <%a6=0%>
         <%End If%>
         <%If mdbor6.EOF=False Then%>
          <%If Mid(mdbor6("MES"),2,2)="0" & Jj Then%>
           <%a4=mdbor6("Summy")%>
           <%mdbor6.MoveNext%>
          <%Else%>
           <%a4=0%>
          <%End If%>
         <%Else%>
          <%a4=0%>
         <%End If%>
         <%suu2=suu2+CDbl(a2)+CDBL(a1)-CDBL(a6)+CDBL(a4)-CDBL(a7)%>
         <%=CDbl(a2)+CDBL(a1)-CDBL(a6)+CDBL(a4)-CDBL(a7)%>
        </td>
       <%Next%>
       <td class="thickbord" width="45">
        <%=suu2%>
       </td>
       <td class="thickbord" width="70">
        <%=suu+suu2%>
       </td>
      </tr>
      <%mdbol3.CommandText="SELECT DISTINCT Enterprise,Edescr FROM inpl WHERE IDentifier='C' AND Yearr=" & ya & " AND SUBSTRING(PC,4,2)='" & MID(mdborl2("PC"),4,2) & "' AND  SUBSTRING(PC,1,2)='" & MID(mdborl2("PC"),1,2) & "'"%>
      <%mdborl3.Open mdbol3%>
      <%Do until mdborl3.EOF%>
       <%jo=1%>
       <%mdbor2.Close%><%mdbor6.Close%><%mdbor2U.Close%><%mdbor5.Close%><%mdbor6a.Close%><%mdbor7.Close%>

       <%mdbo2.CommandText="SELECT ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0) AS summi, SUBSTRING(m.ProjCode, 1, 2) AS be, SUBSTRING(m.ProjCode, 4, 2) AS mi, m.Enterprise, MES FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE SUBSTRING(m.ProjCode,1,5)='" & MID(Mdborl2("PC"),1,5) & "' AND OracleCOde<>'EJB206' AND m.Enterprise='" & Mdborl3("Enterprise") & "' AND (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND ((Konto<>'18164' AND Konto<>'18151') OR ((Konto='18164' OR Konto='18151') AND SUBSTRING(PROJECT,4,3)='999') OR (Konto='18164' AND (PROJECT='EJE573' OR PROJECT='EJE516' OR PROJECT='EJB515' OR PROJECT='EJB583') AND (MES='504' OR MES='505' OR MES='506' OR MES='507'))) AND yearr='" & ya & "' AND (m.IDentifier = 'C') AND KONTO<>'43350' AND SUBKONTO<>'4351'  AND (konto NOT BETWEEN '18410' AND '18433') GROUP BY SUBSTRING(m.ProjCode, 1, 2), SUBSTRING(m.ProjCode, 4, 2), m.Enterprise, MES HAVING (RIGHT(MES,2)>=04 AND LEFT(MES,1)='" & MID(ya,4,1) & "') OR (RIGHT(MES,2)<04 AND LEFT(MES,1)='" & MID(ya+1,4,1) & "') ORDER BY MES,be,mi"%>
       <%mdbor2.Open mdbo2%>
       <%mdbo2u.CommandText="SELECT ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0) AS summi, SUBSTRING(m.ProjCode, 1, 2) AS be, SUBSTRING(m.ProjCode, 4, 2) AS mi, m.Enterprise, MES FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE SUBSTRING(m.ProjCode,1,5)='" & MID(Mdborl2("PC"),1,5) & "' AND m.Enterprise='" & Mdborl3("Enterprise") & "' AND yearr='" & ya & "' AND OracleCOde='EJB206' AND (m.IDentifier = 'C') And gp.description<>'maagaas' AND (konto NOT BETWEEN '18410' AND '18433') AND KONTO<>'43350' AND SUBKONTO<>'4351' AND (konto NOT BETWEEN '18410' AND '18433') GROUP BY SUBSTRING(m.ProjCode, 1, 2), SUBSTRING(m.ProjCode, 4, 2), m.Enterprise, MES HAVING (RIGHT(MES,2)>=04 AND LEFT(MES,1)='" & MID(ya,4,1) & "') OR (RIGHT(MES,2)<04 AND LEFT(MES,1)='" & MID(ya+1,4,1) & "') ORDER BY MES,be,mi"%>
       <%mdbor2u.Open mdbo2u%>
       <%mdbo5.CommandText="SELECT DISTINCT SUM(ISNULL(SummaPlan,0)) AS SP,SUM(ISNULL(PastSum,0)) AS PastSum, be, mi, Enterprise FROM dbo.Delta WHERE yearr='" & ya & "' AND mi = '" & Mid(mdborl2("PC"),4,2) & "' AND Enterprise='" & Mdborl3("Enterprise") & "' AND be='" & Mid(mdborl2("PC"),1,2) & "' AND enn<>'00' GROUP BY be, mi,Enterprise"%>
       <%mdbor5.Open mdbo5%>
       <%mdbo6a.CommandText="SELECT DISTINCT SUM(ISNULL(Ettemaks,0)) AS EM, be, mi, Enterprise, MES FROM dbo.ETE WHERE mi = '" & Mid(mdborl2("PC"),4,2) & "' AND Enterprise='" & Mdborl3("Enterprise") & "' AND be='" & MID(mdborl1("PC"),1,2) & "' AND yearr='" & ya & "' AND enn<>'00' AND MES IS NOT NULL group by be, mi,Enterprise,MES ORDER BY MES,be,mi"%>
       <%mdbor6a.Open mdbo6a%>
       <%mdbo6.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0),0) AS summy,SUBSTRING(m.ProjCode, 1, 2) as be,SUBSTRING(m.ProjCode, 4, 2) AS mi, m.Enterprise, GP.MES FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE SUBSTRING(m.ProjCode,1,5)='" & MID(Mdborl2("PC"),1,5) & "' AND m.Enterprise='" & Mdborl3("Enterprise") & "' AND m.Yearr='" & ya & "' AND (konto BETWEEN '18410' AND '18433') AND (m.IDentifier = 'C') GROUP BY SUBSTRING(m.ProjCode, 1, 2), SUBSTRING(m.ProjCode, 4, 2), m.Enterprise, GP.MES HAVING (RIGHT(MES,2)>=04 AND LEFT(MES,1)='" & MID(ya,4,1) & "') OR (RIGHT(MES,2)<04 AND LEFT(MES,1)='" & MID(ya+1,4,1) & "') ORDER BY MES,be,mi"%>
       <%mdbor6.Open mdbo6%>
       <%mdbo7.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.CREDIT,0))/1000,0),0) AS summy,SUBSTRING(m.ProjCode, 1, 2) as be,SUBSTRING(m.ProjCode, 4, 2) AS mi, m.Enterprise, GP.MES FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE SUBSTRING(m.ProjCode,1,5)='" & MID(Mdborl2("PC"),1,5) & "' AND m.Enterprise='" & Mdborl3("Enterprise") & "' AND m.Yearr='" & ya & "' AND (konto BETWEEN '18410' AND '18433') AND (m.IDentifier = 'C') GROUP BY SUBSTRING(m.ProjCode, 1, 2), SUBSTRING(m.ProjCode, 4, 2), m.Enterprise, GP.MES HAVING (RIGHT(MES,2)>=04 AND LEFT(MES,1)='" & MID(ya,4,1) & "') OR (RIGHT(MES,2)<04 AND LEFT(MES,1)='" & MID(ya+1,4,1) & "') ORDER BY MES,be,mi"%>
       <%mdbor7.Open mdbo7%>
       <%suu=0%><%suu2=0%>
       <tr class="Enterp">
        <td colspan=4>
         <%=mdborl3("EDescr")%>
        </td>
        <td>
         <%If mdbor5.EOF=False Then%>
          <%sim=mdbor5("PastSum")%>
         <%Else%>
          <%sim=0%>
         <%End If%>
         <%If Request.Form("btn")="Kopeerimiseks" Then%>
          <%a0="ad" & mdborl3("Enterprise") & "_" & mdborl2("PC")%>
          <%=Request.Form(a0)%>
          <%If Request.Form(a0)="" Then%>
           <input type="hidden" value="<%=Sim%>" name="<%="ad" & mdborl3("Enterprise") & "_" & mdborl2("PC")%>">
          <%Else%>
           <input type="hidden" value="<%=Request.Form(a0)%>" name="<%="ad" & mdborl3("Enterprise") & "_" & mdborl2("PC")%>">
          <%End If%>
         <%Else%>
          <%a0="ad" & mdborl3("Enterprise") & "_" & mdborl2("PC")%>
          <%If Request.Form(a0)="" Then%>
           <input type="Text" value="<%=sim%>" name="<%="ad" & mdborl3("Enterprise") & "_" & mdborl2("PC")%>" size="10" class="Enterp">
          <%Else%>
           <input type="Text" value="<%=Request.Form(a0)%>" name="<%="ad" & mdborl3("Enterprise") & "_" & mdborl2("PC")%>" size="10" class="Enterp">
          <%End If%>
         <%End If%>
	 <%entt(mdborl3("Enterprise"),jo)=entt(mdborl3("Enterprise"),jo)+CDbl(sim)%>
         <%jo=jo+1%>
        </td>
        <td>
         <%If mdbor5.EOF=False Then%>
          <%sim=mdbor5("SP")%>
          <%=mdbor5("SP")%>
         <%Else%>
          <%sim=0%>
          <%=0%>
         <%End If%>
         <%entt(mdborl3("Enterprise"),jo)=entt(mdborl3("Enterprise"),jo)+CDbl(sim)%>
         <%jo=jo+1%>
        </td>
        <%Jj=4%>
        <%For Jj=4 to 9%>
         <td>
          <%If mdbor2.EOF=False Then%>
           <%If Mid(mdbor2("MES"),2,2)="0" & Jj Then%>
            <%a1=mdbor2("Summi")%>
            <%mdbor2.MoveNext%>
           <%Else%>
            <%a1=0%>
           <%End If%>
          <%Else%>
           <%a1=0%>
          <%End If%>
          <%If mdbor2u.EOF=False Then%>
           <%If Mid(mdbor2u("MES"),2,2)="0" & Jj Then%>
            <%a2=mdbor2u("Summi")%>
            <%mdbor2u.MoveNext%>
           <%Else%>
            <%a2=0%>
           <%End If%>
          <%Else%>
           <%a2=0%>
          <%End If%>
          <%If mdbor7.EOF=False Then%>
           <%If Mid(mdbor7("MES"),2,2)="0" & Jj Then%>
            <%a7=mdbor7("Summy")%>
            <%mdbor7.MoveNext%>
           <%Else%>
            <%a7=0%>
           <%End If%>
          <%Else%>
           <%a7=0%>
          <%End If%>
          <%If mdbor6a.EOF=False Then%>
           <%If Mid(mdbor6a("MES"),2,2)="0" & Jj Then%>
            <%a6=mdbor6a("EM")%>
            <%mdbor6a.MoveNext%>
           <%Else%>
            <%a6=0%>
           <%End If%>
          <%Else%>
           <%a6=0%>
          <%End If%>
          <%If mdbor6.EOF=False Then%>
           <%If Mid(mdbor6("MES"),2,2)="0" & Jj Then%>
            <%a4=mdbor6("Summy")%>
            <%mdbor6.MoveNext%>
           <%Else%>
            <%a4=0%>
           <%End If%>
          <%Else%>
           <%a4=0%>
          <%End If%>
          <%suu=suu+CDbl(a2)+CDBL(a4)-CDBL(a7)+CDBL(a1)-CDBL(a6)%>
          <%sim=CDbl(a2)+CDBL(a4)-CDBL(a7)+CDBL(a1)-CDBL(a6)%>
          <%=Sim%>
          <%entt(mdborl3("Enterprise"),jo)=entt(mdborl3("Enterprise"),jo)+CDbl(sim)%>
          <%jo=jo+1%>
         </td>
        <%Next%>
        <td class="thickbord" width="28">
         <%=suu%><%sim=suu%>
         <%entt(mdborl3("Enterprise"),jo)=entt(mdborl3("Enterprise"),jo)+CDbl(sim)%>
         <%jo=jo+1%>
        </td>
        <%Jj=10%>
        <%For Jj=10 to 12%>
         <td>
          <%If mdbor2.EOF=False Then%>
           <%If Mid(mdbor2("MES"),2,2)=Jj & "" Then%>
            <%a1=mdbor2("Summi")%>
            <%mdbor2.MoveNext%>
           <%Else%>
            <%a1=0%>
           <%End If%>
          <%Else%>
           <%a1=0%>
          <%End If%>
          <%If mdbor2u.EOF=False Then%>
           <%If Mid(mdbor2u("MES"),2,2)=Jj & "" Then%>
            <%a2=mdbor2u("Summi")%>
            <%mdbor2u.MoveNext%>
           <%Else%>
            <%a2=0%>
           <%End If%>
          <%Else%>
           <%a2=0%>
          <%End If%>
          <%If mdbor7.EOF=False Then%>
           <%If Mid(mdbor7("MES"),2,2)=Jj & "" Then%>
            <%a7=mdbor7("Summy")%>
            <%mdbor7.MoveNext%>
           <%Else%>
            <%a7=0%>
           <%End If%>
          <%Else%>
           <%a7=0%>
          <%End If%>
          <%If mdbor6a.EOF=False Then%>
           <%If Mid(mdbor6a("MES"),2,2)=Jj & "" Then%>
            <%a6=mdbor6a("EM")%>
            <%mdbor6a.MoveNext%>
           <%Else%>
            <%a6=0%>
           <%End If%>
          <%Else%>
           <%a6=0%>
          <%End If%>
          <%If mdbor6.EOF=False Then%>
           <%If Mid(mdbor6("MES"),2,2)=Jj & "" Then%>
            <%a4=mdbor6("Summy")%>
            <%mdbor6.MoveNext%>
           <%Else%>
            <%a4=0%>
           <%End If%>
          <%Else%>
           <%a4=0%>
          <%End If%>
          <%suu2=suu2+CDbl(a2)+CDBL(a4)-CDBL(a7)+CDBL(a1)-CDBL(a6)%>
          <%sim=CDbl(a2)+CDBL(a4)-CDBL(a7)+CDBL(a1)-CDBL(a6)%>
          <%=Sim%>
          <%entt(mdborl3("Enterprise"),jo)=entt(mdborl3("Enterprise"),jo)+CDbl(sim)%>
          <%jo=jo+1%>
         </td>
        <%Next%>
        <%Jj=1%>
        <%For Jj=1 to 3%>
         <td>
          <%If mdbor2.EOF=False Then%>
           <%If Mid(mdbor2("MES"),2,2)="0" & Jj Then%>
            <%a1=mdbor2("Summi")%>
            <%mdbor2.MoveNext%>
           <%Else%>
            <%a1=0%>
           <%End If%>
          <%Else%>
           <%a1=0%>
          <%End If%>
          <%If mdbor2u.EOF=False Then%>
           <%If Mid(mdbor2u("MES"),2,2)="0" & Jj Then%>
            <%a2=mdbor2u("Summi")%>
            <%mdbor2u.MoveNext%>
           <%Else%>
            <%a2=0%>
           <%End If%>
          <%Else%>
           <%a2=0%>
          <%End If%>
          <%If mdbor7.EOF=False Then%>
           <%If Mid(mdbor7("MES"),2,2)="0" & Jj Then%>
            <%a7=mdbor7("Summy")%>
            <%mdbor7.MoveNext%>
           <%Else%>
            <%a7=0%>
           <%End If%>
          <%Else%>
           <%a7=0%>
          <%End If%>
          <%If mdbor6a.EOF=False Then%>
           <%If Mid(mdbor6a("MES"),2,2)="0" & Jj Then%>
            <%a6=mdbor6a("EM")%>
            <%mdbor6a.MoveNext%>
           <%Else%>
            <%a6=0%>
           <%End If%>
          <%Else%>
           <%a6=0%>
          <%End If%>
          <%If mdbor6.EOF=False Then%>
           <%If Mid(mdbor6("MES"),2,2)="0" & Jj Then%>
            <%a4=mdbor6("Summy")%>
            <%mdbor6.MoveNext%>
           <%Else%>
            <%a4=0%>
           <%End If%>
          <%Else%>
           <%a4=0%>
          <%End If%>
          <%suu2=suu2+CDbl(a2)+CDBL(a4)-CDBL(a7)+CDBL(a1)-CDBL(a6)%>
          <%sim=CDbl(a2)+CDBL(a4)-CDBL(a7)+CDBL(a1)-CDBL(a6)%>
          <%=Sim%>
          <%entt(mdborl3("Enterprise"),jo)=entt(mdborl3("Enterprise"),jo)+CDbl(sim)%>
          <%jo=jo+1%>
         </td>
        <%Next%>
        <td class="thickbord" width="45">
         <%=suu2%><%sim=suu2%>
         <%entt(mdborl3("Enterprise"),jo)=entt(mdborl3("Enterprise"),jo)+CDbl(sim)%>
         <%jo=jo+1%>
        </td>
        <td class="thickbord" width="70">
         <%=suu+suu2%><%sim=suu+suu2%>
         <%entt(mdborl3("Enterprise"),jo)=entt(mdborl3("Enterprise"),jo)+CDbl(sim)%>
         <%jo=jo+1%>
        </td>
       </tr>
       <%fotnum=1%>
       <%mdbo1.CommandText="SELECT DISTINCT Pid,RenovBlock,PC,ProjName,OracleCode,Enterprise,Footnote FROM inpl WHERE IDentifier='C' AND Yearr=" & ya & " AND SUBSTRING(PC,1,2)='" & MID(mdborl1("PC"),1,2) & "' AND SUBSTRING(PC,4,2)='" & MID(mdborl2("PC"),4,2) & "' AND SUBSTRING(PC,7,2)<>'00' AND Enterprise='" & Mdborl3("Enterprise") & "' ORDER BY PC"%>
       <%mdbor.Open mdbo1%>
       <%mdbosl.CommandText="SELECT ROUND(SUM(ISNULL(DEBET,0))/1000,0) AS summi, m.ProjCode FROM dbo.Main AS m INNER JOIN dbo.glav_project AS GP ON m.OracleCode = GP.PROJECT WHERE m.Yearr='" & ya & "' AND LEFT(m.ProjCode,5)='" & MID(mdborl2("PC"),1,5) & "' AND m.Enterprise='" & mdborl3("Enterprise") & "' AND ((LEFT(MES,1)<='" & MID(ya-1,4,1) & "') OR (LEFT(MES,1)='" & MID(ya,4,1) & "' AND RIGHT(MES,1)<04) OR (LEFT(MES,1)='9')) AND (m.IDentifier = 'C') GROUP BY m.ProjCode ORDER BY M.ProjCode"%>
       <%mdborsl.Open mdbosl%>
       <%Do Until mdbor.EOF%>
        <%jo=1%>
        <%mdbor2.Close%><%mdbor2u.Close%><%mdbor5.Close%><%mdbor6.Close%><%mdbor6a.Close%><%mdbor7.Close%>
        <%IF LEN(mdbor("PC"))>9 AND RIGHT(mdbor("PC"),2)="00" then%>
         <%mdbo2.CommandText="SELECT ROUND(SUM(ISNULL(Gp.DEBET,0))/1000,0) AS summi, MES FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE SUBSTRING(m.ProjCode,1,8)='" & MID(Mdbor("PC"),1,8) & "' AND OracleCOde<>'EJB206' AND m.Enterprise='" & Mdbor("Enterprise") & "' AND (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND ((Konto<>'18164' AND Konto<>'18151') OR ((Konto='18164' OR Konto='18151') AND SUBSTRING(PROJECT,4,3)='999') OR (Konto='18164' AND (PROJECT='EJE573' OR PROJECT='EJE516' OR PROJECT='EJB515' OR PROJECT='EJB583') AND (MES='504' OR MES='505' OR MES='506' OR MES='507'))) AND KONTO<>'43350' AND SUBKONTO<>'4351' AND yearr='" & ya & "' AND (m.IDentifier = 'C') AND (konto NOT BETWEEN '18410' AND '18433') GROUP BY MES HAVING (RIGHT(MES,2)>=04 AND LEFT(MES,1)='" & MID(ya,4,1) & "') OR (RIGHT(MES,2)<04 AND LEFT(MES,1)='" & MID(ya+1,4,1) & "') ORDER BY MES"%>
         <%mdbor2.Open mdbo2%>
         <%mdbo2u.CommandText="SELECT ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0) AS summi, MES FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE SUBSTRING(m.ProjCode,1,8)='" & MID(Mdbor("PC"),1,8) & "' AND m.Enterprise='" & Mdbor("Enterprise") & "' AND right(M.ProjCode,2)<>'00' AND yearr='" & ya & "' AND OracleCOde='EJB206' AND (m.IDentifier = 'C') And gp.description<>'maagaas' AND (konto NOT BETWEEN '18410' AND '18433') GROUP BY MES HAVING (RIGHT(MES,2)>=04 AND LEFT(MES,1)='" & MID(ya,4,1) & "') OR (RIGHT(MES,2)<04 AND LEFT(MES,1)='" & MID(ya+1,4,1) & "') ORDER BY MES"%>
         <%mdbor2u.Open mdbo2u%>
         <%mdbo5.CommandText="SELECT ISNULL(SUM(ISNULL(SummaPlan,0)),0) as Summaplan ,ISNULL(SUM(ISNULL(PastSum,0)),0) as PastSum FROM delta WHERE yearr='" & ya & "' AND LEFT(ProjCode,8)='" & LEFT(mdbor("PC"),8) & "' and Enterprise='" & Mdbor("Enterprise") & "' AND right(ProjCode,2)<>'00' AND enn<>'00'"%>
         <%mdbor5.Open mdbo5%>
	 <%mdbo6.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0),0) AS summy, GP.MES FROM glav_project GP WHERE Project='" & mdbor("OracleCode") & "' AND (konto BETWEEN '18410' AND '18433') GROUP BY GP.MES HAVING (RIGHT(MES,2)>=04 AND LEFT(MES,1)='" & MID(ya,4,1) & "') OR (RIGHT(MES,2)<04 AND LEFT(MES,1)='" & MID(ya+1,4,1) & "') ORDER BY MES"%>
         <%mdbor6.Open mdbo6%>                 
	 <%mdbo6a.CommandText="SELECT DISTINCT SUM(ISNULL(Ettemaks,0)) AS EM, MES FROM dbo.ETE WHERE LEFT(ProjCode,8)='" & LEFT(mdbor("PC"),8) & "' AND Enterprise='" & Mdbor("Enterprise") & "' AND yearr='" & ya & "' AND enn<>'00' AND MES IS NOT NULL group by MES ORDER BY MES"%>
         <%mdbor6a.Open mdbo6a%>        
         <%mdbo7.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.CREDIT,0))/1000,0),0) AS summy, GP.MES FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE SUBSTRING(m.ProjCode,1,8)='" & MID(Mdbor("PC"),1,8) & "' AND m.Enterprise='" & Mdbor("Enterprise") & "' AND m.Yearr='" & ya & "' AND (konto BETWEEN '18410' AND '18433') AND (m.IDentifier = 'C') GROUP BY GP.MES HAVING (RIGHT(MES,2)>=04 AND LEFT(MES,1)='" & MID(ya,4,1) & "') OR (RIGHT(MES,2)<04 AND LEFT(MES,1)='" & MID(ya+1,4,1) & "') ORDER BY MES"%>
         <%mdbor7.Open mdbo7%>
        <%ELSE%>
         <%mdbo2.CommandText="SELECT ROUND(SUM(ISNULL(DEBET,0))/1000,0) AS summi, MES FROM glav_project WHERE Project='" & mdbor("OracleCode") &"' AND (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND ((Konto<>'18164' AND Konto<>'18151') OR ((Konto='18164' OR Konto='18151') AND SUBSTRING(PROJECT,4,3)='999') OR (Konto='18164' AND (PROJECT='EJE573' OR PROJECT='EJE516' OR PROJECT='EJB515' OR PROJECT='EJB583') AND (MES='504' OR MES='505' OR MES='506' OR MES='507'))) AND Project<>'EJB206' AND KONTO<>'43350' AND SUBKONTO<>'4351' AND (konto NOT BETWEEN '18410' AND '18433') GROUP BY MES HAVING (RIGHT(MES,2)>=04 AND LEFT(MES,1)='" & MID(ya,4,1) & "') OR (RIGHT(MES,2)<04 AND LEFT(MES,1)='" & MID(ya+1,4,1) & "') ORDER BY MES"%>
         <%mdbor2.Open mdbo2%>
         <%'="SELECT ROUND(SUM(ISNULL(DEBET,0))/1000,0) AS summi, MES FROM glav_project WHERE Project='" & mdbor("OracleCode") &"' AND (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND ((Konto<>'18164' AND Konto<>'18151') OR ((Konto='18164' OR Konto='18151') AND SUBSTRING(PROJECT,4,3)='999') OR (Konto='18164' AND (PROJECT='EJE573' OR PROJECT='EJE516' OR PROJECT='EJB515' OR PROJECT='EJB583') AND (MES='504' OR MES='505' OR MES='506' OR MES='507'))) AND Project<>'EJB206' AND KONTO<>'43350' AND SUBKONTO<>'4351' AND (konto NOT BETWEEN '18410' AND '18433') GROUP BY MES HAVING (RIGHT(MES,2)>=04 AND LEFT(MES,1)='" & MID(ya,4,1) & "') OR (RIGHT(MES,2)<04 AND LEFT(MES,1)='" & MID(ya+1,4,1) & "') ORDER BY MES"%>
         <%mdbo2u.CommandText="SELECT ROUND(SUM(ISNULL(DEBET,0))/1000,0) AS summi, MES FROM glav_project WHERE Project='" & mdbor("OracleCode") &"' AND Project='EJB206' And description<>'maagaas' AND KONTO<>'43350' AND SUBKONTO<>'4351' AND (konto NOT BETWEEN '18410' AND '18433') GROUP BY MES HAVING (RIGHT(MES,2)>=04 AND LEFT(MES,1)='" & MID(ya,4,1) & "') OR (RIGHT(MES,2)<04 AND LEFT(MES,1)='" & MID(ya+1,4,1) & "') ORDER BY MES"%>
         <%mdbor2u.Open mdbo2u%>
         <%mdbo5.CommandText="SELECT ProjCode,SummaPlan,OracleCode,PastSum FROM delta WHERE Pid='" & mdbor("Pid") & "' AND yearr='" & ya & "' AND enn<>'00'"%>
         <%mdbor5.Open mdbo5%>
         <%mdbo6a.CommandText="SELECT DISTINCT SUM(ISNULL(Ettemaks,0)) AS EM, MES FROM dbo.ETE WHERE Pid='" & mdbor("Pid") & "' AND yearr='" & ya & "' AND enn<>'00' AND MES IS NOT NULL GROUP BY MES"%>
         <%mdbor6a.Open mdbo6a%>         
         <%mdbo6.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0),0) AS summy, GP.MES FROM glav_project GP WHERE Project='" & mdbor("OracleCode") & "' AND (konto BETWEEN '18410' AND '18433') GROUP BY GP.MES HAVING (RIGHT(MES,2)>=04 AND LEFT(MES,1)='" & MID(ya,4,1) & "') OR (RIGHT(MES,2)<04 AND LEFT(MES,1)='" & MID(ya+1,4,1) & "') ORDER BY MES"%>        
         <%mdbor6.Open mdbo6%>
         <%mdbo7.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.CREDIT,0))/1000,0),0) AS summy, GP.MES FROM glav_project GP WHERE Project='" & mdbor("OracleCode") & "' AND (konto BETWEEN '18410' AND '18433') GROUP BY GP.MES HAVING (RIGHT(MES,2)>=04 AND LEFT(MES,1)='" & MID(ya,4,1) & "') OR (RIGHT(MES,2)<04 AND LEFT(MES,1)='" & MID(ya+1,4,1) & "') ORDER BY MES"%>
         <%mdbor7.Open mdbo7%>
        <%End If%>
        <%suu=0%>
        <%suu2=0%>
       <tr>
        <td>
         <%=mdbor("OracleCode")%>
        </td>
        <td>
         <%if mid(mdbor("PC"),8,1)=0 and mid(mdbor("PC"),7,1)<>0 then%>
          <%If len(mdbor("PC"))>=9 then%>
           <%a=REPLACE(MID(mdbor("PC"),1,6), "0", "") & MID(mdbor("PC"),7,2) & REPLACE(MID(mdbor("PC"),9,4), "0", "")%>
          <%Else%>         
           <%a=REPLACE(MID(mdbor("PC"),1,6), "0", "") & MID(mdbor("PC"),7,2)%>
          <%End if%>
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
         <%If Request.Form("btn")="Kopeerimiseks" Then%>
          <%a0="aa" & mdborl3("Enterprise") & "_" & mdbor("PC")%>
          <%=Request.Form(a0)%>
         <%Else%>
          <%a0="aa" & mdborl3("Enterprise") & "_" & mdbor("PC")%>
          <input type="Text" value="<%=Request.Form(a0)%>" size="10" name="<%="aa" & mdborl3("Enterprise") & "_" & mdbor("PC")%>">
         <%End If%>
        </td>
        <td>
         <%If mdbor5.EOF=False Then%>
          <%sim=mdbor5("PastSum")%>
         <%Else%>
          <%sim=0%>
         <%End If%>
         <%If Request.Form("btn")="Kopeerimiseks" Then%>
          <%a0="ad" & mdborl3("Enterprise") & "_" & mdbor("PC")%>
          <%=Request.Form(a0)%>
          <%If Request.Form(a0)="" Then%>
           <input type="hidden" value="<%=Sim%>" name="<%="ad" & mdborl3("Enterprise") & "_" & mdbor("PC")%>">
          <%Else%>
           <input type="hidden" value="<%=Request.Form(a0)%>" name="<%="ad" & mdborl3("Enterprise") & "_" & mdbor("PC")%>">
          <%End If%>
         <%Else%>
          <%a0="ad" & mdborl3("Enterprise") & "_" & mdbor("PC")%>
          <%If Request.Form(a0)="" Then%>
           <input type="Text" value="<%=sim%>" name="<%="ad" & mdborl3("Enterprise") & "_" & mdbor("PC")%>" size="10">
          <%Else%>
           <input type="Text" value="<%=Request.Form(a0)%>" name="<%="ad" & mdborl3("Enterprise") & "_" & mdbor("PC")%>" size="10">
          <%End If%>
         <%End If%>
         <%If mdbor("RenovBlock")=0 AND (MID(mdbor("PC"),10,2)<>"00") Then%>
          <%ent2(mdborl3("Enterprise"),jo)=ent2(mdborl3("Enterprise"),jo)+CDbl(sim)%>
         <%End If%>
         <%jo=jo+1%>
        </td>
        <td>
         <%If mdbor5.EOF=False Then%>
          <%sim=mdbor5("SummaPlan")%>
          <%=mdbor5("SummaPlan")%>
         <%else%>
          <%sim=0%>
          0
         <%End If%>
         <%If mdbor("RenovBlock")=0 AND (MID(mdbor("PC"),10,2)<>"00") Then%>
          <%ent2(mdborl3("Enterprise"),jo)=ent2(mdborl3("Enterprise"),jo)+CDbl(sim)%>
         <%End If%>
         <%jo=jo+1%>
        </td>
        <%Jj=4%>
        <%For Jj=4 to 9%>
         <td>
          <%If mdbor2.EOF=False Then%>
           <%If Mid(mdbor2("MES"),2,2)="0" & Jj Then%>
            <%a1=mdbor2("Summi")%>
            <%mdbor2.MoveNext%>
           <%Else%>
            <%a1=0%>
           <%End If%>
          <%Else%>
           <%a1=0%>
          <%End If%>
          <%If mdbor2u.EOF=False Then%>
           <%If Mid(mdbor2u("MES"),2,2)="0" & Jj Then%>
            <%a2=mdbor2u("Summi")%>
            <%mdbor2u.MoveNext%>
           <%Else%>
            <%a2=0%>
           <%End If%>
          <%Else%>
           <%a2=0%>
          <%End If%>
          <%If mdbor7.EOF=False Then%>
           <%If Mid(mdbor7("MES"),2,2)="0" & Jj Then%>
            <%a7=mdbor7("Summy")%>
            <%mdbor7.MoveNext%>
           <%Else%>
            <%a7=0%>
           <%End If%>
          <%Else%>
           <%a7=0%>
          <%End If%>
	  <%If mdbor6.EOF=False Then%>
           <%If Mid(mdbor6("MES"),2,2)="0" & Jj Then%>
            <%a4=mdbor6("Summy")%>
            <%mdbor6.MoveNext%>
           <%Else%>
            <%a4=0%>
           <%End If%>
          <%Else%>
           <%a4=0%>
          <%End If%>
          <%If mdbor6a.EOF=False Then%>
           <%If Mid(mdbor6a("MES"),2,2)="0" & Jj Then%>
            <%a6=mdbor6a("EM")%>
            <%mdbor6a.MoveNext%>
           <%Else%>
            <%a6=0%>
           <%End If%>
          <%Else%>
           <%a6=0%>
          <%End If%>


          <%suu=suu+CDbl(a2)+CDBL(a4)-CDBL(a7)+CDBL(a1)-CDBL(a6)%>
          <%sim=CDbl(a2)+CDBL(a4)-CDBL(a7)+CDBL(a1)-CDBL(a6)%>
          <%=Sim%>
          <%If mdbor("RenovBlock")=0 AND (MID(mdbor("PC"),10,2)<>"00") Then%>
           <%ent2(mdborl3("Enterprise"),jo)=ent2(mdborl3("Enterprise"),jo)+CDbl(sim)%>
          <%End If%>
          <%jo=jo+1%>
         </td>
        <%Next%>
        <td  class="thickbord" width="28">
         <%If suu<>1 or suu<>-1 then%>
          <%=suu%>
          <%sim=suu%>
         <%Else%>
          0
          <%sim=0%>
         <%END iF%>
         <%If mdbor("RenovBlock")=0 AND (MID(mdbor("PC"),10,2)<>"00") Then%>
          <%ent2(mdborl3("Enterprise"),jo)=ent2(mdborl3("Enterprise"),jo)+CDbl(sim)%>
         <%End If%>
         <%jo=jo+1%>
        </td>
        <%Jj=10%>
        <%For Jj=10 to 12%>
         <td>
          <%If mdbor2.EOF=False Then%>
           <%If Mid(mdbor2("MES"),2,2)=Jj & "" Then%>
            <%a1=mdbor2("Summi")%>
            <%mdbor2.MoveNext%>
           <%Else%>
            <%a1=0%>
           <%End If%>
          <%Else%>
           <%a1=0%>
          <%End If%>
          <%If mdbor2u.EOF=False Then%>
           <%If Mid(mdbor2u("MES"),2,2)=Jj & "" Then%>
            <%a2=mdbor2u("Summi")%>
            <%mdbor2u.MoveNext%>
           <%Else%>
            <%a2=0%>
           <%End If%>
          <%Else%>
           <%a2=0%>
          <%End If%>
          <%If mdbor7.EOF=False Then%>
           <%If Mid(mdbor7("MES"),2,2)=Jj & "" Then%>
            <%a7=mdbor7("Summy")%>
            <%mdbor7.MoveNext%>
           <%Else%>
            <%a7=0%>
           <%End If%>
          <%Else%>
           <%a7=0%>
          <%End If%>
          <%If mdbor6a.EOF=False Then%>
           <%If Mid(mdbor6a("MES"),2,2)=Jj & "" Then%>
            <%a6=mdbor6a("EM")%>
            <%mdbor6a.MoveNext%>
           <%Else%>
            <%a6=0%>
           <%End If%>
          <%Else%>
           <%a6=0%>
          <%End If%>
          <%If mdbor6.EOF=False Then%>
           <%If Mid(mdbor6("MES"),2,2)=Jj & "" Then%>
            <%a4=mdbor6("Summy")%>
            <%mdbor6.MoveNext%>
           <%Else%>
            <%a4=0%>
           <%End If%>
          <%Else%>
           <%a4=0%>
          <%End If%>



          <%suu2=suu2+CDbl(a2)+CDBL(a4)-CDBL(a7)+CDBL(a1)-CDBL(a6)%>
          <%sim=CDbl(a2)+CDBL(a4)-CDBL(a7)+CDBL(a1)-CDBL(a6)%>
          <%=Sim%>
          <%If mdbor("RenovBlock")=0 AND (MID(mdbor("PC"),10,2)<>"00") Then%>
           <%ent2(mdborl3("Enterprise"),jo)=ent2(mdborl3("Enterprise"),jo)+CDbl(sim)%>
          <%End If%>
          <%jo=jo+1%>
         </td>
        <%Next%>
        <%Jj=1%>
        <%For Jj=1 to 3%>
         <td>
          <%If mdbor2.EOF=False Then%>
           <%If Mid(mdbor2("MES"),2,2)="0" & Jj Then%>
            <%a1=mdbor2("Summi")%>
            <%mdbor2.MoveNext%>
           <%Else%>
            <%a1=0%>
           <%End If%>
          <%Else%>
           <%a1=0%>
          <%End If%>
          <%If mdbor2u.EOF=False Then%>
           <%If Mid(mdbor2u("MES"),2,2)="0" & Jj Then%>
            <%a2=mdbor2u("Summi")%>
            <%mdbor2u.MoveNext%>
           <%Else%>
            <%a2=0%>
           <%End If%>
          <%Else%>
           <%a2=0%>
          <%End If%>
          <%If mdbor7.EOF=False Then%>
           <%If Mid(mdbor7("MES"),2,2)="0" & Jj Then%>
            <%a7=mdbor7("Summy")%>
            <%mdbor7.MoveNext%>
           <%Else%>
            <%a7=0%>
           <%End If%>
          <%Else%>
           <%a7=0%>
          <%End If%>
          <%If mdbor6a.EOF=False Then%>
           <%If Mid(mdbor6a("MES"),2,2)="0" & Jj Then%>
            <%a6=mdbor6a("EM")%>
            <%mdbor6a.MoveNext%>
           <%Else%>
            <%a6=0%>
           <%End If%>
          <%Else%>
           <%a6=0%>
          <%End If%>
          <%If mdbor6.EOF=False Then%>
           <%If Mid(mdbor6("MES"),2,2)="0" & Jj Then%>
            <%a4=mdbor6("Summy")%>
            <%mdbor6.MoveNext%>
           <%Else%>
            <%a4=0%>
           <%End If%>
          <%Else%>
           <%a4=0%>
          <%End If%>
          <%suu2=suu2+CDbl(a2)+CDBL(a4)-CDBL(a7)+CDBL(a1)-CDBL(a6)%>
          <%sim=CDbl(a2)+CDBL(a4)-CDBL(a7)+CDBL(a1)-CDBL(a6)%>
          <%=Sim%>
          <%If mdbor("RenovBlock")=0 AND (MID(mdbor("PC"),10,2)<>"00") Then%>
           <%ent2(mdborl3("Enterprise"),jo)=ent2(mdborl3("Enterprise"),jo)+CDbl(sim)%>
          <%End If%>
          <%jo=jo+1%>
         </td>
        <%Next%>
        <td class="thickbord" width="45">
         <%If suu2<>1 or suu2<>-1 then%>
          <%=suu2%>
          <%sim=suu2%>
         <%Else%>
          0
          <%sim=0%>
         <%END iF%>
         <%If mdbor("RenovBlock")=0 AND (MID(mdbor("PC"),10,2)<>"00") Then%>
          <%ent2(mdborl3("Enterprise"),jo)=ent2(mdborl3("Enterprise"),jo)+CDbl(sim)%>
         <%End If%>
         <%jo=jo+1%>
        </td>
        <td class="thickbord" width="70">
         <%If (suu+suu2)<>1 or (suu+suu2)<>-1 then%>
          <%=suu+suu2%>
          <%sim=suu+suu2%>
         <%Else%>
          0
          <%sim=0%>
         <%END iF%>
         <%If mdbor("RenovBlock")=0 AND (MID(mdbor("PC"),10,2)<>"00") Then%>
          <%ent2(mdborl3("Enterprise"),jo)=ent2(mdborl3("Enterprise"),jo)+CDbl(sim)%>
         <%End If%>
         <%jo=jo+1%>
        </td>
       </tr>
       <%If mdborsl.EOF=True Then%>
       <%Else%>
        <%mdborsl.Movenext%>
       <%End IF%>
       <%mdbor.Movenext%>
      <%loop%>
      <%mdbor.Close%>
      <%mdborsl.Close%>
      <%mdborl3.Movenext%>
     <%loop%>
     <%mdborl3.Close%>
     <%mdborl2.Movenext%>
    <%loop%>
    <%mdborl2.Close%>
    <%mdborl1.Movenext%>
   <%Loop%>
   <%mdbor2.Close%>
   <%mdbor5.Close%>
   <%mdbor2u.Close%>
   <%Dim koku(17)%><%Dim kok2(17)%>
   <tr>
    <td colspan="21" width="1301">Kokku ettev&otildette kaupa</td>
   </tr>
   <%mdbo5.CommandText="SELECT * FROM Enterprise ORDER BY ENTERPRISE"%>
   <%mdbor5.Open mdbo5%>
   <%Do until mdbor5.EOF%>
    <tr class="boldEnterp">
     <td colspan=4>
      <%=mdbor5("EDescr")%>
     </td>
     <%For nuu=5 to 21%>
      <td>
       <%=entt(Mdbor5("Enterprise"),nuu-4)%>
      </td>
      <%koku(nuu-4)=koku(nuu-4)+entt(Mdbor5("Enterprise"),nuu-4)%>
     <%Next%>
    </tr>
    <%mdbor5.Movenext%>
   <%Loop%>
   <tr>
    <td colspan=4 class="bold">Kokku</td>
    <%For nuu=5 to 21%>
     <td class="bold"><%=koku(nuu-4)%></td>
    <%Next%>
   </tr>
   <tr>
    <td class="bold" colspan="21" width="1301">
     Kokku ettev&otildette kaupa, v&auml;lja arvatud plokkide renoveerimine
    </td>
   </tr>
   <%mdbor5.MoveFirst%>
   <%Do until mdbor5.EOF%>
    <tr class="boldEnterp">
     <td colspan=4>
      <%=mdbor5("EDescr")%>
     </td>
     <%For nuu=5 to 21%>
      <td>
       <%=ent2(Mdbor5("Enterprise"),nuu-4)%>
      </td>
      <%kok2(nuu-4)=kok2(nuu-4)+ent2(Mdbor5("Enterprise"),nuu-4)%>
     <%Next%>
    </tr>
    <%mdbor5.Movenext%>
   <%Loop%>
   <tr>
    <td class="bold" colspan=4>Kokku</td>
    <%For nuu=5 to 21%>
     <td class="bold">
      <%=kok2(nuu-4)%>
     </td>
    <%Next%>
   </tr>
  </table>
 </body>
</html>