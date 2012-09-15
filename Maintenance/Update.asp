<%
Response.Expires = 0
Response.AddHeader "pragma", "no-cache"
%>
<Html>
<Head>
<meta http-equiv="Content-Type" content="text/html; charset=ibm852">
<title>
InformatsiooniSusteem Investeerimise Kava Teostamise Kontrollimiseks. Invest-IT!on
</title>
</Head>

<%b= Server.MapPath("\inv")%>
<%set mdbo =  Server.CreateObject("ADODB.Connection")%>
<%set servcfg=Server.CreateObject("Scripting.FileSystemObject")
  set servFileStream=servcfg.OpenTextFile(b & "\server.cfg")
  s=servFileStream.ReadLine
  i=servFileStream.ReadLine
  p=servFileStream.ReadLine
  servFileStream.Close%>
<%mdbo.ConnectionString="Driver={SQL Server};Server=" & s & ";Database=invest;Trusted_Connection=yes;"%>
<%mdbo.Open ConnectionString%>
<%set mdbo1 = Server.CreateObject("ADODB.Command")%>
<%set mdbor = Server.CreateObject("ADODB.Recordset")%>
<%mdbo1.ActiveConnection = mdbo%>
<%set mdbo6 = Server.CreateObject("ADODB.Command")%>
<%set mdbor6 = Server.CreateObject("ADODB.Recordset")%>
<%mdbo6.ActiveConnection = mdbo%>
<%set mdbo7 = Server.CreateObject("ADODB.Command")%>
<%set mdbor7 = Server.CreateObject("ADODB.Recordset")%>
<%mdbo7.ActiveConnection = mdbo%>
<%set mdbo6a = Server.CreateObject("ADODB.Command")%>
<%set mdbor6a = Server.CreateObject("ADODB.Recordset")%>
<%mdbo6a.ActiveConnection = mdbo%>
<%set mdbo2 = Server.CreateObject("ADODB.Command")%>
<%set mdbor2 = Server.CreateObject("ADODB.Recordset")%>
<%mdbo2.ActiveConnection = mdbo%>
<%set mdbo2u = Server.CreateObject("ADODB.Command")%>
<%set mdbor2u = Server.CreateObject("ADODB.Recordset")%>
<%mdbo2u.ActiveConnection = mdbo%>

<%set mdbou = Server.CreateObject("ADODB.Command")%>
<%set mdboru = Server.CreateObject("ADODB.Recordset")%>
<%mdbou.ActiveConnection = mdbo%>
<%ya=Year(Date())%>
<%mo=Month(Date())%>
<%da=Day(Date())%>
<%zz=mo-04%>

<%If zz>=0 Then%>
<%ya=Year(Date())%>
<%Else%>
<%ya=ya-1%>
<%End If%>





<%mdbo1.CommandText="SELECT DISTINCT Pid,RenovBlock,PC,ProjName,OracleCode,Enterprise,Footnote FROM inpl WHERE IDentifier='C' AND Yearr=" & ya & " ORDER BY PC"%>
<%mdbor.Open mdbo1%>
<%Do Until mdbor.EOF%>
<%jo=1%>
<%IF LEN(mdbor("PC"))>9 AND RIGHT(mdbor("PC"),2)="00" then%>
<%mdbo2.CommandText="SELECT ROUND(SUM(ISNULL(Gp.DEBET,0))/1000,0) AS summi, MES FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE SUBSTRING(m.ProjCode,1,8)='" & MID(Mdbor("PC"),1,8) & "' AND OracleCOde<>'EJB206' AND m.Enterprise='" & Mdbor("Enterprise") & "' AND (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND ((Konto<>'18164' AND Konto<>'18151') OR ((Konto='18164' OR Konto='18151') AND SUBSTRING(PROJECT,4,3)='999') OR (Konto='18164' AND (PROJECT='EJE573' OR PROJECT='EJE516' OR PROJECT='EJB515' OR PROJECT='EJB583') AND (MES='504' OR MES='505' OR MES='506' OR MES='507'))) AND KONTO<>'43350' AND SUBKONTO<>'4351' AND yearr='" & ya & "' AND (m.IDentifier = 'C') AND (konto NOT BETWEEN '18410' AND '18433') GROUP BY MES HAVING ((RIGHT(MES,2)>=04 AND LEFT(MES,1)='" & MID(ya,4,1) & "') OR (RIGHT(MES,2)<04 AND LEFT(MES,1)='" & MID(ya+1,4,1) & "')) ORDER BY MES"%>
<%mdbor2.Open mdbo2%>
<%mdbo2u.CommandText="SELECT ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0) AS summi, MES FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE SUBSTRING(m.ProjCode,1,8)='" & MID(Mdbor("PC"),1,8) & "' AND m.Enterprise='" & Mdbor("Enterprise") & "' AND right(M.ProjCode,2)<>'00' AND yearr='" & ya & "' AND OracleCOde='EJB206' AND (m.IDentifier = 'C') And gp.description<>'maagaas' AND (konto NOT BETWEEN '18410' AND '18433') GROUP BY MES HAVING ((RIGHT(MES,2)>=04 AND LEFT(MES,1)='" & MID(ya,4,1) & "') OR (RIGHT(MES,2)<04 AND LEFT(MES,1)='" & MID(ya+1,4,1) & "')) ORDER BY MES"%>
<%mdbor2u.Open mdbo2u%>
<%mdbo6a.CommandText="SELECT DISTINCT SUM(ISNULL(Ettemaks,0)) AS EM, MES FROM dbo.ETE WHERE LEFT(ProjCode,8)='" & LEFT(mdbor("PC"),8) & "' AND Enterprise='" & Mdbor("Enterprise") & "' AND yearr='" & ya & "' AND enn<>'00' AND MES IS NOT NULL group by MES ORDER BY MES"%>
<%mdbor6a.Open mdbo6a%>
<%mdbo6.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0),0)-ISNULL(ROUND(SUM(ISNULL(GP.CREDIT,0))/1000,0),0) AS summy, GP.MES FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE SUBSTRING(m.ProjCode,1,8)='" & MID(Mdbor("PC"),1,8) & "' AND m.Enterprise='" & Mdbor("Enterprise") & "' AND m.Yearr='" & ya & "' AND (konto BETWEEN '18410' AND '18433') AND (m.IDentifier = 'C') GROUP BY GP.MES HAVING ((RIGHT(MES,2)>=04 AND LEFT(MES,1)='" & MID(ya-1,4,1) & "') OR (RIGHT(MES,2)<04 AND LEFT(MES,1)='" & MID(ya,4,1) & "')) ORDER BY MES"%>
<%mdbor6.Open mdbo6%>
<%mdbo7.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0),0)-ISNULL(ROUND(SUM(ISNULL(GP.CREDIT,0))/1000,0),0) AS summy, GP.MES FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE SUBSTRING(m.ProjCode,1,8)='" & MID(Mdbor("PC"),1,8) & "' AND m.Enterprise='" & Mdbor("Enterprise") & "' AND m.Yearr='" & ya & "' AND (konto BETWEEN '18410' AND '18433') AND (m.IDentifier = 'C') GROUP BY GP.MES HAVING ((RIGHT(MES,2)>=04 AND LEFT(MES,1)='" & MID(ya,4,1) & "') OR (RIGHT(MES,2)<04 AND LEFT(MES,1)='" & MID(ya+1,4,1) & "')) ORDER BY MES"%>
<%mdbor7.Open mdbo7%>
<%ELSE%>
<%mdbo2.CommandText="SELECT ROUND(SUM(ISNULL(DEBET,0))/1000,0) AS summi, MES FROM glav_project WHERE Project='" & mdbor("OracleCode") &"' AND (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND ((Konto<>'18164' AND Konto<>'18151') OR ((Konto='18164' OR Konto='18151') AND SUBSTRING(PROJECT,4,3)='999') OR (Konto='18164' AND (PROJECT='EJE573' OR PROJECT='EJE516' OR PROJECT='EJB515' OR PROJECT='EJB583') AND (MES='504' OR MES='505' OR MES='506' OR MES='507'))) AND Project<>'EJB206' AND KONTO<>'43350' AND SUBKONTO<>'4351' AND (konto NOT BETWEEN '18410' AND '18433') GROUP BY MES HAVING ((RIGHT(MES,2)>=04 AND LEFT(MES,1)='" & MID(ya,4,1) & "') OR (RIGHT(MES,2)<04 AND LEFT(MES,1)='" & MID(ya+1,4,1) & "')) ORDER BY MES"%>
<%mdbor2.Open mdbo2%>
<%mdbo2u.CommandText="SELECT ROUND(SUM(ISNULL(DEBET,0))/1000,0) AS summi, MES FROM glav_project WHERE Project='" & mdbor("OracleCode") &"' AND Project='EJB206' And description<>'maagaas' AND KONTO<>'43350' AND SUBKONTO<>'4351' AND (konto NOT BETWEEN '18410' AND '18433') GROUP BY MES HAVING ((RIGHT(MES,2)>=04 AND LEFT(MES,1)='" & MID(ya,4,1) & "') OR (RIGHT(MES,2)<04 AND LEFT(MES,1)='" & MID(ya+1,4,1) & "')) ORDER BY MES"%>
<%mdbor2u.Open mdbo2u%>
<%mdbo6a.CommandText="SELECT DISTINCT SUM(ISNULL(Ettemaks,0)) AS EM, MES FROM dbo.ETE WHERE Pid='" & mdbor("Pid") & "' AND yearr='" & ya & "' AND enn<>'00' AND MES IS NOT NULL GROUP BY MES"%>
<%mdbor6a.Open mdbo6a%>
<%mdbo6.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0),0)-ISNULL(ROUND(SUM(ISNULL(GP.CREDIT,0))/1000,0),0) AS summy, GP.MES FROM glav_project GP WHERE Project='" & mdbor("OracleCode") & "' AND (konto BETWEEN '18410' AND '18433') GROUP BY GP.MES HAVING ((RIGHT(MES,2)>=04 AND LEFT(MES,1)='" & MID(ya-1,4,1) & "') OR (RIGHT(MES,2)<04 AND LEFT(MES,1)='" & MID(ya,4,1) & "')) ORDER BY MES"%>
<%mdbor6.Open mdbo6%>
<%mdbo7.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0),0)-ISNULL(ROUND(SUM(ISNULL(GP.CREDIT,0))/1000,0),0) AS summy, GP.MES FROM glav_project GP WHERE Project='" & mdbor("OracleCode") & "' AND (konto BETWEEN '18410' AND '18433') GROUP BY GP.MES HAVING ((RIGHT(MES,2)>=04 AND LEFT(MES,1)='" & MID(ya,4,1) & "') OR (RIGHT(MES,2)<04 AND LEFT(MES,1)='" & MID(ya+1,4,1) & "')) ORDER BY MES"%>
<%mdbor7.Open mdbo7%>
<%End If%>

<%a1_1=0%>
<%a1_2=0%>
<%a1_3=0%>
<%a1_4=0%>

<%a2_1=0%>
<%a2_2=0%>
<%a2_3=0%>
<%a2_4=0%>

<%a3_1=0%>
<%a3_2=0%>
<%a3_3=0%>
<%a3_4=0%>

<%a6_1=0%>
<%a6_2=0%>
<%a6_3=0%>
<%a6_4=0%>

<%a7_1=0%>
<%a7_2=0%>
<%a7_3=0%>
<%a7_4=0%>


<%For i=4 to 6%>
<%If mdbor2.EOF=False Then%>
<%y=i%>
<%IF mdbor2("MES")=MID(ya,4,1) & "0" & i then%>
<%a1_1=a1_1+CDBL(mdbor2("Summi"))%>
<%mdbor2.Movenext%>
<%ELSE%>
<%a1_1=a1_1+0%>
<%End If%>
<%Else%>
<%a1_1=a1_1+0%>
<%End If%>
<%Next%>
<%For i=7 to 9%>
<%If mdbor2.EOF=False Then%>
<%IF mdbor2("MES")=MID(ya,4,1) & "0" & i then%>
<%a1_2=a1_2+CDBL(mdbor2("Summi"))%>
<%mdbor2.Movenext%>
<%ELSE%>
<%a1_2=a1_2+0%>
<%End If%>
<%Else%>
<%a1_2=a1_2+0%>
<%End If%>
<%Next%>
<%For i=10 to 12%>
<%If mdbor2.EOF=False Then%>
<%IF mdbor2("MES")=MID(ya,4,1) & i then%>
<%a1_3=a1_3+CDBL(mdbor2("Summi"))%>
<%mdbor2.Movenext%>
<%ELSE%>
<%a1_3=a1_3+0%>
<%End If%>
<%Else%>
<%a1_3=a1_3+0%>
<%End If%>
<%Next%>
<%For i=1 to 3%>
<%If mdbor2.EOF=False Then%>
<%IF mdbor2("MES")=MID(ya+1,4,1) & "0" & i then%>
<%a1_4=a1_4+CDBL(mdbor2("Summi"))%>
<%mdbor2.Movenext%>
<%ELSE%>
<%a1_4=a1_4+0%>
<%End If%>
<%Else%>
<%a1_4=0%>
<%End If%>
<%Next%>

<%For i=4 to 6%>
<%If mdbor2u.EOF=False Then%>
<%IF mdbor2u("MES")=MID(ya,4,1) & "0" & i then%>
<%a2_1=a2_1+CDBL(mdbor2u("Summi"))%>
<%mdbor2u.Movenext%>
<%ELSE%>
<%a2_1=a2_1+0%>
<%End If%>
<%Else%>
<%a2_1=a2_1+0%>
<%End If%>
<%Next%>
<%For i=7 to 9%>
<%If mdbor2u.EOF=False Then%>
<%IF mdbor2u("MES")=MID(ya,4,1) & "0" & i then%>
<%a2_2=a2_2+CDBL(mdbor2u("Summi"))%>
<%mdbor2u.Movenext%>
<%ELSE%>
<%a2_2=a2_2+0%>
<%End If%>
<%Else%>
<%a2_2=a2_2+0%>
<%End If%>
<%Next%>
<%For i=10 to 12%>
<%If mdbor2u.EOF=False Then%>
<%IF mdbor2u("MES")=MID(ya,4,1) & i then%>
<%a2_3=a2_3+CDBL(mdbor2u("Summi"))%>
<%mdbor2u.Movenext%>
<%ELSE%>
<%a2_3=a2_3+0%>
<%End If%>
<%Else%>
<%a2_3=a2_3+0%>
<%End If%>
<%Next%>
<%For i=1 to 3%>
<%If mdbor2u.EOF=False Then%>
<%IF mdbor2u("MES")=MID(ya+1,4,1) & "0" & i then%>
<%a2_4=a2_4+CDBL(mdbor2u("Summi"))%>
<%mdbor2u.Movenext%>
<%ELSE%>
<%a2_4=a2_4+0%>
<%End If%>
<%Else%>
<%a2_4=a2_4+0%>
<%End If%>
<%Next%>

<%For i=4 to 6%>
<%If mdbor6.EOF=False Then%>
<%IF mdbor6("MES")=MID(ya,4,1) & "0" & i then%>
<%a3_1=a3_1+CDBL(mdbor6("Summy"))%>
<%mdbor6.Movenext%>
<%ELSE%>
<%a3_1=a3_1+0%>
<%End If%>
<%Else%>
<%a3_1=a3_1+0%>
<%End If%>
<%Next%>
<%For i=7 to 9%>
<%If mdbor6.EOF=False Then%>
<%IF mdbor6("MES")=MID(ya,4,1) & "0" &i then%>
<%a3_2=a3_2+CDBL(mdbor6("Summy"))%>
<%mdbor6.Movenext%>
<%ELSE%>
<%a3_2=a3_2+0%>
<%End If%>
<%Else%>
<%a3_2=a3_2+0%>
<%End If%>
<%Next%>
<%For i=10 to 12%>
<%If mdbor6.EOF=False Then%>
<%IF mdbor6("MES")=MID(ya,4,1) & i then%>
<%a3_3=a3_3+CDBL(mdbor6("Summy"))%>
<%mdbor6.Movenext%>
<%ELSE%>
<%a3_3=a3_3+0%>
<%End If%>
<%Else%>
<%a3_3=a3_3+0%>
<%End If%>
<%Next%>
<%For i=1 to 3%>
<%If mdbor6.EOF=False Then%>
<%IF mdbor6("MES")=MID(ya+1,4,1) & "0" & i then%>
<%a3_4=a3_4+CDBL(mdbor6("Summy"))%>
<%mdbor6.Movenext%>
<%ELSE%>
<%a3_4=a3_4+0%>
<%End If%>
<%Else%>
<%a3_4=a3_4+0%>
<%End If%>
<%Next%>

<%For i=4 to 6%>
<%If mdbor6a.EOF=False Then%>
<%IF mdbor6a("MES")=MID(ya,4,1) & "0" & i then%>
<%a6_1=a6_1+CDBL(mdbor6a("EM"))%>
<%mdbor6a.Movenext%>
<%ELSE%>
<%a6_1=a6_1+0%>
<%End If%>
<%Else%>
<%a6_1=a6_1+0%>
<%End If%>
<%Next%>
<%For i=7 to 9%>
<%If mdbor6a.EOF=False Then%>
<%IF mdbor6a("MES")=MID(ya,4,1) & "0" & i then%>
<%a6_2=a6_2+CDBL(mdbor6a("EM"))%>
<%mdbor6a.Movenext%>
<%ELSE%>
<%a6_2=a6_2+0%>
<%End If%>
<%Else%>
<%a6_2=a6_2+0%>
<%End If%>
<%Next%>
<%For i=10 to 12%>
<%If mdbor6a.EOF=False Then%>
<%IF mdbor6a("MES")=MID(ya,4,1) & i then%>
<%a6_3=a6_3+CDBL(mdbor6a("EM"))%>
<%mdbor6a.Movenext%>
<%ELSE%>
<%a6_3=a6_3+0%>
<%End If%>
<%Else%>
<%a6_3=a6_3+0%>
<%End If%>
<%Next%>
<%For i=1 to 3%>
<%If mdbor6a.EOF=False Then%>
<%IF mdbor6a("MES")=MID(ya+1,4,1) & "0" & i then%>
<%a6_4=a6_4+CDBL(mdbor6a("EM"))%>
<%mdbor6a.Movenext%>
<%ELSE%>
<%a6_4=a6_4+0%>
<%End If%>
<%Else%>
<%a6_4=a6_4+0%>
<%End If%>
<%Next%>


<%For i=4 to 6%>
<%If mdbor7.EOF=False Then%>
<%IF mdbor7("MES")=MID(ya,4,1) & "0" & i then%>
<%a7_1=a7_1+CDBL(mdbor7("Summy"))%>
<%mdbor7.Movenext%>
<%ELSE%>
<%a7_1=a7_1+0%>
<%End If%>
<%Else%>
<%a7_1=a7_1+0%>
<%End If%>
<%Next%>
<%For i=7 to 9%>
<%If mdbor7.EOF=False Then%>
<%IF mdbor7("MES")=MID(ya,4,1) & "0" & i then%>
<%a7_2=a7_2+CDBL(mdbor7("Summy"))%>
<%mdbor7.Movenext%>
<%ELSE%>
<%a7_2=a7_2+0%>
<%End If%>
<%Else%>
<%a7_2=a7_2+0%>
<%End If%>
<%Next%>
<%For i=10 to 12%>
<%If mdbor7.EOF=False Then%>
<%IF mdbor7("MES")=MID(ya,4,1) & i then%>
<%a7_3=a7_3+CDBL(mdbor7("Summy"))%>
<%mdbor7.Movenext%>
<%ELSE%>
<%a7_3=a7_3+0%>
<%End If%>
<%Else%>
<%a7_3=a7_3+0%>
<%End If%>
<%Next%>
<%For i=1 to 3%>
<%If mdbor7.EOF=False Then%>
<%IF mdbor7("MES")=MID(ya+1,4,1) & "0" & i then%>
<%a7_4=a7_4+CDBL(mdbor7("Summy"))%>
<%mdbor7.Movenext%>
<%ELSE%>
<%a7_4=a7_4+0%>
<%End If%>
<%Else%>
<%a7_4=a7_4+0%>
<%End If%>
<%Next%>

<%va1=CDbl(a2_1)+CDBL(a1_1)-CDBL(a3_1)-CDBL(a6_1)+CDBL(a7_1)%>
<%va2=CDbl(a2_2)+CDBL(a1_2)-CDBL(a3_2)-CDBL(a6_2)+CDBL(a7_2)%>
<%va3=CDbl(a2_3)+CDBL(a1_3)-CDBL(a3_3)-CDBL(a6_3)+CDBL(a7_3)%>
<%va4=CDbl(a2_4)+CDBL(a1_4)-CDBL(a3_4)-CDBL(a6_4)+CDBL(a7_4)%>
<%mdbor6.Close%>
<%mdbor2.Close%>
<%mdbor2u.Close%>
<%mdbor6a.Close%>
<%mdbor7.Close%>
<%=va1%>
<%=va2%>
<%=va3%>
<%=va4%>

<%mdbou.Commandtext="UPDATE Main SET Ikvartal=" & va1 & " WHERE Pid='" & mdbor("Pid") & "' AND Enterprise='" & mdbor("Enterprise") & "' AND Yearr='" & ya & "' AND Identifier='F'"%>
<%mdboru.Open mdbou%>
<%mdbou.Commandtext="UPDATE Main SET IIkvartal=" & va2 & " WHERE Pid='" & mdbor("Pid") & "' AND Enterprise='" & mdbor("Enterprise") & "' AND Yearr='" & ya & "' AND Identifier='F'"%>
<%mdboru.Open mdbou%>
<%mdbou.Commandtext="UPDATE Main SET IIIkvartal=" & va3 & " WHERE Pid='" & mdbor("Pid") & "' AND Enterprise='" & mdbor("Enterprise") & "' AND Yearr='" & ya & "' AND Identifier='F'"%>
<%mdboru.Open mdbou%>
<%mdbou.Commandtext="UPDATE Main SET IVkvartal=" & va4 & " WHERE Pid='" & mdbor("Pid") & "' AND Enterprise='" & mdbor("Enterprise") & "' AND Yearr='" & ya & "' AND Identifier='F'"%>
<%mdboru.Open mdbou%>

<%Mdbor.MoveNext%>

<%Loop%>
<%mdbor.Close%>
Tegelike summade kohta k&auml;ivate andmete uuendamine k&otilde;ikides projektides on toimunud edukalt. Tagasi <a href="https://intranet/inv/">pealehek&uuml;ljele!</a>.
</body>
</html>