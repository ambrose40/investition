<Html>
 <%Set servcfg=Server.CreateObject("Scripting.FileSystemObject")%>
 <%b= Server.MapPath("\")%>
 <Head>
  <%if request.Cookies("StyleInv")="" then%>
   <%Set servFileStream=servcfg.OpenTextFile(b & "\Style.cfg")%>
   <%s=servFileStream.ReadLine%>
   <%servFileStream.Close%>
   <link rel="Stylesheet" href="<%=s%>" type="text/css">
  <%else%>
   <%s=request.Cookies("StyleInv")%>
   <link rel="Stylesheet" href="<%=s%>" type="text/css">
  <%End If%>
 </Head>
 <Body Class="report">
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
  <%Set mdbo1 = Server.CreateObject("ADODB.Command")%>
  <%Set mdbor = Server.CreateObject("ADODB.Recordset")%>
  <%mdbo1.ActiveConnection = mdbo%>
  <%Set mdbo2 = Server.CreateObject("ADODB.Command")%>
  <%Set mdbor2 = Server.CreateObject("ADODB.Recordset")%>
  <%mdbo2.ActiveConnection = mdbo%>
  <%Set mdbo2u = Server.CreateObject("ADODB.Command")%>
  <%Set mdbor2u = Server.CreateObject("ADODB.Recordset")%>
  <%mdbo2u.ActiveConnection = mdbo%>
  <%Set mdbo5 = Server.CreateObject("ADODB.Command")%>
  <%Set mdbor5 = Server.CreateObject("ADODB.Recordset")%>
  <%mdbo5.ActiveConnection = mdbo%>
  <%Set mdbo6 = Server.CreateObject("ADODB.Command")%>
  <%Set mdbor6 = Server.CreateObject("ADODB.Recordset")%>
  <%mdbo6.ActiveConnection = mdbo%>
  <%Set mdbo6a = Server.CreateObject("ADODB.Command")%>
  <%Set mdbor6a = Server.CreateObject("ADODB.Recordset")%>
  <%mdbo6a.ActiveConnection = mdbo%>
  <%Set mdbo7 = Server.CreateObject("ADODB.Command")%>
  <%Set mdbor7 = Server.CreateObject("ADODB.Recordset")%>
  <%mdbo7.ActiveConnection = mdbo%>
  <Table Border=1>
   <tr>
    <th rowspan="2">Projekti nr.</th>
    <th rowspan="2">Nr</th>
    <th rowspan="2">Projekti nimetus</th>
    <th rowspan="2">Ehitusaastad (m.a.)</th>
    <th rowspan="2">L&otilde;petatud seisuga&nbsp1.4.<%=ya%></th>
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
   <tr Class="Repnum">
    <%For nuu=1 to 21%>
     <td><%=nuu%></td>
    <%Next%>
   </tr>
   <%d=Month(Date()) & "." & Day(Date()) & "." & Year(Date())%>
   <%aa=0%><%ab=0%><%ac=0%>
   <%mdbo2.CommandText="SELECT ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0) AS summi, GP.MES FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE m.RenovBlock=0 AND Project<>'EJB206' AND (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND ((Konto<>'18164' AND Konto<>'18151') OR ((Konto='18164' OR Konto='18151') AND SUBSTRING(PROJECT,4,3)='999') OR (Konto='18164' AND (PROJECT='EJE573' OR PROJECT='EJE516' OR PROJECT='EJB515' OR PROJECT='EJB583') AND (MES='504' OR MES='505' OR MES='506' OR MES='507'))) AND yearr='" & ya & "' AND (m.IDentifier = 'C') AND (konto NOT BETWEEN '18410' AND '18433') AND KONTO<>'43350' AND SUBKONTO<>'4351' GROUP BY GP.MES HAVING (RIGHT(MES,2)>=04 AND LEFT(MES,1)='" & MID(ya,4,1) & "') OR (RIGHT(MES,2)<04 AND LEFT(MES,1)='" & MID(ya+1,4,1) & "') ORDER BY MES"%>
   <%mdbor2.Open mdbo2%>
   <%mdbo2u.CommandText="SELECT ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0) AS summi, GP.MES FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE m.RenovBlock=0 AND Project='EJB206' AND yearr='" & ya & "' AND (m.IDentifier = 'C') and GP.description<>'maagaas' AND KONTO<>'43350' AND SUBKONTO<>'4351' AND (konto NOT BETWEEN '18410' AND '18433') GROUP BY GP.MES HAVING (RIGHT(MES,2)>=04 AND LEFT(MES,1)='" & MID(ya,4,1) & "') OR (RIGHT(MES,2)<04 AND LEFT(MES,1)='" & MID(ya+1,4,1) & "') ORDER BY MES"%>
   <%mdbor2u.Open mdbo2u%>
   <%mdbo5.CommandText="SELECT DISTINCT SUM(ISNULL(SummaPlan,0)) AS SP,SUM(ISNULL(PastSum,0)) as PastSum FROM dbo.Delta WHERE RenovBlock=0 AND yearr='" & ya & "' AND enn<>'00'"%>
   <%mdbor5.Open mdbo5%>
   <%mdbo6a.CommandText="SELECT DISTINCT SUM(ISNULL(Ettemaks,0)) AS EM,MES FROM dbo.ETE WHERE yearr='" & ya & "' AND enn<>'00' AND RenovBlock=0 AND MES IS NOT NULL group by MES ORDER BY MES"%>
   <%mdbor6a.Open mdbo6a%>
   <%mdbo6.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0),0)-ISNULL(ROUND(SUM(ISNULL(GP.CREDIT,0))/1000,0),0) AS summy, GP.MES FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE m.Yearr='" & ya & "' AND (konto BETWEEN '18410' AND '18433') AND m.RenovBlock=0 AND (m.IDentifier = 'C') GROUP BY GP.MES HAVING (RIGHT(MES,2)>=04 AND LEFT(MES,1)='" & MID(ya-1,4,1) & "') OR (RIGHT(MES,2)<04 AND LEFT(MES,1)='" & MID(ya,4,1) & "') ORDER BY MES"%>
   <%mdbor6.Open mdbo6%>
   <%mdbo7.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0),0)-ISNULL(ROUND(SUM(ISNULL(GP.CREDIT,0))/1000,0),0) AS summy, GP.MES FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE m.Yearr='" & ya & "' AND (konto BETWEEN '18410' AND '18433') AND m.RenovBlock=0 AND (m.IDentifier = 'C') GROUP BY GP.MES HAVING (RIGHT(MES,2)>=04 AND LEFT(MES,1)='" & MID(ya,4,1) & "') OR (RIGHT(MES,2)<04 AND LEFT(MES,1)='" & MID(ya+1,4,1) & "') ORDER BY MES"%>
   <%mdbor7.Open mdbo7%>
   <%suu=0%><%suu2=0%>
   <tr Class="whitetr">
    <td></td>
    <td>01.01.02.00.</td>
    <td>INVESTEERINGUDaKOKKUv</td>
    <td>
     <%If Request.Form("btn")="Kopeerimiseks" Then%>
      <%a0="a1a"%>
      <%=Request.Form(a0)%>
      <input type="hidden" value="<%=Request.Form(a0)%>" name="<%="a1a"%>">
     <%Else%>
      <%a0="a1a"%>
      <input type="Text" value="<%=Request.Form(a0)%>" size="10" name="<%="a1a"%>" Style="font-family: Verdana; font-weight:700; color: #FFFFFF; background-color: #FFFFFF; Border-width:0">
     <%End If%>
    </td>
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
       <input type="Text" value="<%=sim%>" name="<%="a1d"%>" size="10" Style="font-family: Verdana; color: #FFFFFF; background-color: #FFFFFF; Border-width:0">
      <%Else%>
       <input type="Text" value="<%=Request.Form(a0)%>" name="<%="a1d"%>" size="10" Style="font-family: Verdana; color: #FFFFFF; background-color: #FFFFFF; Border-width:0">
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
        <%a3=mdbor6("Summy")%>
        <%mdbor6.MoveNext%>
       <%Else%>
        <%a3=0%>
       <%End If%>
      <%Else%>
       <%a3=0%>
      <%End If%>
      <%suu=suu+CDbl(a2)+CDBL(a1)-CDBL(a3)-CDBL(a6)+CDBL(a7)%>
      <%=CDbl(a2)+CDBL(a1)-CDBL(a3)-CDBL(a6)+CDBL(a7)%>
     </td>
    <%Next%>
    <td>
     <%=suu%>
    </td>
    <%Jj=10%>
    <%For Jj=10 to 12%>
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
        <%a3=mdbor6("Summy")%>
        <%mdbor6.MoveNext%>
       <%Else%>
        <%a3=0%>
       <%End If%>
      <%Else%>
       <%a3=0%>
      <%End If%>
      <%suu=suu+CDbl(a2)+CDBL(a1)-CDBL(a3)-CDBL(a6)+CDBL(a7)%>
      <%=CDbl(a2)+CDBL(a1)-CDBL(a3)-CDBL(a6)+CDBL(a7)%>
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
        <%a3=mdbor6("Summy")%>
        <%mdbor6.MoveNext%>
       <%Else%>
        <%a3=0%>
       <%End If%>
      <%Else%>
       <%a3=0%>
      <%End If%>
      <%suu=suu+CDbl(a2)+CDBL(a1)-CDBL(a3)-CDBL(a6)+CDBL(a7)%>
      <%=CDbl(a2)+CDBL(a1)-CDBL(a3)-CDBL(a6)+CDBL(a7)%>
     </td>
    <%Next%>
    <td>
     <%=suu2%>&nbsp
    </td>
    <td  Style="Border-left-width:3; Border-right-width:3">
     <%=suu+suu2%>&nbsp
    </td>
   </tr>
   <%mdbor2.Close%><%mdbor2u.Close%><%mdbor5.Close%><%mdbor6.Close%><%mdbor6a.Close%><%mdbor7.Close%>
  </Table>
 </Body>
</Html>