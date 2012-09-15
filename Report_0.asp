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
  <%If Request.Form("btn")="OK" Then%>
   <%ya=Request.Form("ye")%>
  <%Else%>
   <%ya=Request.QueryString("ye")%>
   <%=Request.QueryString("ye")%>
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
   <%set servFileStream=servcfg.OpenTextFile(b & "\server.cfg")%>
   <%s=servFileStream.ReadLine%>
   <%i=servFileStream.ReadLine%>
   <%p=servFileStream.ReadLine%>
   <%servFileStream.Close%>
   <%set mdbo =  Server.CreateObject("ADODB.Connection")%>
   <%mdbo.ConnectionString="Driver={SQL Server};Server=" & s & ";Database=invest;Trusted_Connection=yes;"%>
   <%mdbo.Open ConnectionString%>
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
     <%For nuu=1 to 20%>
      <td><%=nuu%></td>
     <%Next%>
    </tr>

    <%d=Month(Date()) & "." & Day(Date()) & "." & Year(Date())%>
    <%aa=0%><%ab=0%><%ac=0%>
    <%Dim entt(10,23)%>
    <%Dim ent2(10,23)%>

<%mdbo2.CommandText="SELECT ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0) AS summi, GP.MES FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE Project<>'EJB206' AND (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND ((Konto<>'18164' AND Konto<>'18151') OR ((Konto='18164' OR Konto='18151') AND SUBSTRING(PROJECT,4,3)='999') OR (Konto='18164' AND (PROJECT='EJE573' OR PROJECT='EJE516' OR PROJECT='EJB515' OR PROJECT='EJB583') AND (MES='504' OR MES='505' OR MES='506' OR MES='507'))) AND yearr='" & ya & "' AND (m.IDentifier = 'C') AND (konto NOT BETWEEN '18410' AND '18433') AND KONTO<>'43350' AND SUBKONTO<>'4351' GROUP BY GP.MES HAVING (RIGHT(MES,2)>=04 AND LEFT(MES,1)='" & MID(ya,4,1) & "') OR (RIGHT(MES,2)<04 AND LEFT(MES,1)='" & MID(ya+1,4,1) & "') ORDER BY MES"%>
    <%mdbor2.Open mdbo2%>
    <%mdbo2u.CommandText="SELECT ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0) AS summi, GP.MES FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE Project='EJB206' AND yearr='" & ya & "' AND (m.IDentifier = 'C') and GP.description<>'maagaas' AND KONTO<>'43350' AND SUBKONTO<>'4351' AND (konto NOT BETWEEN '18410' AND '18433') GROUP BY GP.MES HAVING (RIGHT(MES,2)>=04 AND LEFT(MES,1)='" & MID(ya,4,1) & "') OR (RIGHT(MES,2)<04 AND LEFT(MES,1)='" & MID(ya+1,4,1) & "') ORDER BY MES"%>
    <%mdbor2u.Open mdbo2u%>
    <%mdbo5.CommandText="SELECT DISTINCT SUM(ISNULL(SummaPlan,0)) AS SP,SUM(ISNULL(PastSum,0)) as PastSum FROM dbo.Delta WHERE  yearr='" & ya & "' AND enn<>'00'"%>
    <%mdbor5.Open mdbo5%>
    <%mdbo6a.CommandText="SELECT DISTINCT SUM(ISNULL(Ettemaks,0)) AS EM,MES FROM dbo.ETE WHERE yearr='" & ya & "' AND enn<>'00' AND  MES IS NOT NULL group by MES ORDER BY MES"%>
    <%mdbor6a.Open mdbo6a%>
    <%mdbo6.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0),0) AS summy, GP.MES FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE m.Yearr='" & ya & "' AND (konto BETWEEN '18410' AND '18433') AND (m.IDentifier = 'C') GROUP BY GP.MES HAVING (RIGHT(MES,2)>=04 AND LEFT(MES,1)='" & MID(ya,4,1) & "') OR (RIGHT(MES,2)<04 AND LEFT(MES,1)='" & MID(ya+1,4,1) & "') ORDER BY MES"%>
    <%mdbor6.Open mdbo6%>
    <%mdbo7.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.CREDIT,0))/1000,0),0) AS summy, GP.MES FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE m.Yearr='" & ya & "' AND (konto BETWEEN '18410' AND '18433') AND (m.IDentifier = 'C') GROUP BY GP.MES HAVING (RIGHT(MES,2)>=04 AND LEFT(MES,1)='" & MID(ya,4,1) & "') OR (RIGHT(MES,2)<04 AND LEFT(MES,1)='" & MID(ya+1,4,1) & "') ORDER BY MES"%>
    <%mdbor7.Open mdbo7%>
    <%suu=0%><%suu2=0%>
    <tr class="boldProjGrup">
     <td colspan=3>INVESTEERINGUD KOKKU</td>
     <td>
      <%If mdbor5.EOF=False Then%>
       <%=mdbor5("PastSum")%>
      <%Else%>
       <%=0%>
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

       <%mdbo1.CommandText="SELECT DISTINCT Pid,RenovBlock,PC,ProjName,OracleCode,Enterprise,Footnote FROM inpl WHERE IDentifier='C' AND Yearr=" & ya & " and SUBSTRING(PC,7,2)<>'00' and (oraclecode like '%999%' or oraclecode like '%EJ%') ORDER BY PC"%>
       <%mdbor.Open mdbo1%>
       <%mdbosl.CommandText="SELECT ROUND(SUM(ISNULL(DEBET,0))/1000,0) AS summi, m.ProjCode FROM dbo.Main AS m INNER JOIN dbo.glav_project AS GP ON m.OracleCode = GP.PROJECT WHERE m.Yearr='" & ya & "' AND ((LEFT(MES,1)<='" & MID(ya-1,4,1) & "') OR (LEFT(MES,1)='" & MID(ya,4,1) & "' AND RIGHT(MES,1)<04) OR (LEFT(MES,1)='9')) AND (m.IDentifier = 'C') GROUP BY m.ProjCode ORDER BY M.ProjCode"%>
       <%mdborsl.Open mdbosl%>
       <%Do Until mdbor.EOF%>
        <%jo=1%>
        <%IF LEN(mdbor("PC"))>9 AND RIGHT(mdbor("PC"),2)="00" then%>
        <%ELSE%>
         <%mdbo2.CommandText="SELECT ROUND(SUM(ISNULL(DEBET,0))/1000,0) AS summi, MES FROM glav_project WHERE Project='" & mdbor("OracleCode") &"' AND (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND ((Konto<>'18164' AND Konto<>'18151') OR ((Konto='18164' OR Konto='18151') AND SUBSTRING(PROJECT,4,3)='999') OR (Konto='18164' AND (PROJECT='EJE573' OR PROJECT='EJE516' OR PROJECT='EJB515' OR PROJECT='EJB583') AND (MES='504' OR MES='505' OR MES='506' OR MES='507'))) AND Project<>'EJB206' AND KONTO<>'43350' AND SUBKONTO<>'4351' AND (konto NOT BETWEEN '18410' AND '18433') GROUP BY MES HAVING (RIGHT(MES,2)>=04 AND LEFT(MES,1)='" & MID(ya,4,1) & "') OR (RIGHT(MES,2)<04 AND LEFT(MES,1)='" & MID(ya+1,4,1) & "') ORDER BY MES"%>
         <%mdbor2.Open mdbo2%>
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
         <%If mdbor5.EOF=False Then%>
          <%=mdbor5("PastSum")%>
         <%Else%>
          <%=0%>
         <%End If%>
        </td>
        <td>
         <%If mdbor5.EOF=False Then%>
          <%=mdbor5("SummaPlan")%>
         <%else%>
          0
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


          <%suu=suu+CDbl(a2)+CDbl(a4)+CDBL(a1)-CDBL(a6)-CDBL(a7)%>
          <%sim=CDbl(a2)+CDbl(a4)+CDBL(a1)-CDBL(a6)-CDBL(a7)%>
          <%=Sim%>
         </td>
        <%Next%>
        <td  class="thickbord" width="28">
         <%If suu<>1 or suu<>-1 then%>
          <%=suu%>
         <%Else%>
          0
         <%END iF%>
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

          <%suu2=suu2+CDbl(a2)+CDbl(a4)+CDBL(a1)-CDBL(a6)-CDBL(a7)%>
          <%=CDbl(a2)+CDbl(a4)+CDBL(a1)-CDBL(a6)-CDBL(a7)%>
          <%'=Sim%>
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
          <%suu2=suu2+CDbl(a2)+CDbl(a4)+CDBL(a1)-CDBL(a6)-CDBL(a7)%>
          <%sim=CDbl(a2)+CDbl(a4)+CDBL(a1)-CDBL(a6)-CDBL(a7)%>
          <%=Sim%>
         </td>
        <%Next%>
        <td class="thickbord" width="45">
         <%If suu2<>1 or suu2<>-1 then%>
          <%=suu2%>
         <%Else%>
          0
         <%END iF%>
        </td>
        <td class="thickbord" width="70">
         <%If (suu+suu2)<>1 or (suu+suu2)<>-1 then%>
          <%=suu+suu2%>
         <%Else%>
          0
         <%END iF%>
        </td>
       </tr>
       <%If mdborsl.EOF=True Then%>
       <%Else%>
        <%mdborsl.Movenext%>
       <%End IF%>
   <%mdbor2.Close%>
   <%mdbor5.Close%>
   <%mdbor2u.Close%>
   <%mdbor6.Close%>
   <%mdbor7.Close%>
   <%mdbor6a.Close%>
 <%End If%>
       <%mdbor.Movenext%>
      <%loop%>
      <%mdbor.Close%>
      <%mdborsl.Close%>

   <%Dim koku(17)%><%Dim kok2(17)%>
  </table>
 </body>
</html>