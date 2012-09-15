<Html>
 <Head>
  <meta http-equiv="Content-Type" content="text/Html; charSet=windows-1251">
  <Title>
   InFORmatsiooniSusteem Investeerimise Kava Teostamise Kontrollimiseks. Invest-IT!on
  </Title>
 </Head>
 <Body>
  <%If Request.FORm("btn")="OK" Then%>
   <%ya=Request.FORm("ye")%>
  <%Else%>
   <%ya=Request.QueryString("ya")%>
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
   <%If Request.FORm("btn")="OK" Then%>
    <%ya=Request.FORm("ye")%>
   <%Else%>
    <%ya=Request.QueryString("ya")%>
   <%End If%>
  <%End If%>
  <a Href="ps.asp?ya=<%=zzz%>"><Img BORDER="0" src="icons/p.ico" Style=float:left></a><a Href="ps.asp?ya=<%=zzz2%>"><Img BORDER="0" src="icons/n.ico" Style=float:right></a>
  <%b= Server.MapPath("\")%>
  <%Set mdbo =  Server.CreateObject("ADODB.Connection")%>
  <%Set servcfg=Server.CreateObject("Scripting.FileSystemObject")%>
  <%Set servFileStream=servcfg.OpenTextFile(b & "\server.cfg")%>
  <%s=servFileStream.ReadLine%>
  <%i=servFileStream.ReadLine%>
  <%p=servFileStream.ReadLine%>
  <%servFileStream.Close%>
  <%mdbo.ConnectionString="Driver={SQL Server};Server=" & s & ";Database=invest;Trusted_Connection=yes;"%>
  <%mdbo.Open ConnectionString%>
  <%Set mdbol1 = Server.CreateObject("ADODB.Command")%>
  <%Set mdborl1 = Server.CreateObject("ADODB.RecORdSet")%>
  <%mdbol1.ActiveConnection = mdbo%>
  <%Set mdbol1b = Server.CreateObject("ADODB.Command")%>
  <%Set mdborl1b = Server.CreateObject("ADODB.RecORdSet")%>
  <%mdbol1b.ActiveConnection = mdbo%>
  <%Set mdbol2 = Server.CreateObject("ADODB.Command")%>
  <%Set mdborl2 = Server.CreateObject("ADODB.RecORdSet")%>
  <%mdbol2.ActiveConnection = mdbo%>
  <%Set mdbol2b = Server.CreateObject("ADODB.Command")%>
  <%Set mdborl2b = Server.CreateObject("ADODB.RecORdSet")%>
  <%mdbol2b.ActiveConnection = mdbo%>
  <%Set mdbo3 = Server.CreateObject("ADODB.Command")%>
  <%Set mdbor3 = Server.CreateObject("ADODB.RecORdSet")%>
  <%mdbo3.ActiveConnection = mdbo%>
  <%Set mdbo1 = Server.CreateObject("ADODB.Command")%>
  <%Set mdbor = Server.CreateObject("ADODB.RecORdSet")%>
  <%mdbo1.ActiveConnection = mdbo%>
  <%Set mdbolc = Server.CreateObject("ADODB.Command")%>
  <%Set mdborlc = Server.CreateObject("ADODB.RecORdSet")%>
  <%mdbolc.ActiveConnection = mdbo%>
  <%d=Month(Date()) & "." & Day(Date()) & "." & Year(Date())%>
  <%yy= zo & np%>
  <%SELECT Case yy%>
   <%Case zo%>
    <%yr=zo%>
   <%Case np%>
    <%yr=np%>
   <%Case zo & np%>
    <%yr=zo & " AND " & np%>
  <%End SELECT%>
  <%aa=0%><%ab=0%><%ac=0%>

  <%Set servFileStream=servcfg.CreateTextFile(b & "\investps.csv")%>
  <%uid=uid+1%>
  <%servFileStream.writeline("Nr|ProjName|ProjMangr|Levell|GROUPe|DatBeg|DatEnd|PlanSumm|FaktSumm|ContrSumm|Enterprise|UID|")%>

  <%mdbol1.CommandText="SELECT * FROM projserv WHERE Yearr=" & ya & " AND SUBSTRING(ProjCode,4,2)='00' ORDER BY ProjCode"%>
  <%mdborl1.Open mdbol1%>
  <%mdbol1b.CommandText="SELECT SUM(summtot) as st,SUBSTRING(ProjCode,1,2) FROM projserv3 WHERE Yearr='" & ya & "' GROUP BY IdentIfier,SUBSTRING(ProjCode,1,2)"%>
  <%mdborl1b.Open mdbol1b%>
  <%Do Until mdborl1.EOF%>
   <%mdbolc.CommandText="SELECT * FROM projserv WHERE Yearr=" & ya & " AND SUBSTRING(ProjCode,1,2)='" & Mid(mdborl1("ProjCode"),1,2) & "' AND ProjCode<>'" & mdborl1("ProjCode") & "' ORDER BY ProjCode"%>
   <%mdborlc.Open mdbolc%>
   <%If mdborlc.EOF Then%>
    <%grp="Нет"%>
   <%Else%>
    <%grp="Да"%>
   <%End If%>
   <%mdborlc.Close%>
   <%FOR i=1 to len(Mid(mdborl1("ProjCode"),1,2))%>
    <%a=Mid(Mid(mdborl1("ProjCode"),1,2),i,1)%>
    <%If i=1 Then%>
     <%a2="."%>
    <%Else%>
     <%a2=Mid(mdborl1("ProjCode"),i-1,1)%>
    <%End If%>
    <%If a="0" AND (a2="0" OR a2=".") Then%>
    <%Else%>
     <%pc0=pc0 & a%>
    <%End If%>
   <%Next%>
   <%cc=0%>
   <%mdborl1b.MoveNext%>
   <%ff=0%>
   <%mdborl1b.MoveNext%>
   <%pp=0%>
   <%mdborl1b.MoveNext%>
   <%sdr=sdr+1%>
   <%uid=uid+1%>
   <%a=(pc0 & "|" & mdborl1("ProjName") & "||1|" & grp & "|||" & pp & "|" & ff & "|" & cc & "||" & uid & "|")%>
   <%a=REPLACE(a, "&auml;", "a")%>
   <%a=REPLACE(a, "&ouml;", "o")%>
   <%a=REPLACE(a, "&uuml;", "u")%>
   <%a=REPLACE(a, "&otilde;", "y")%>
   <%a=REPLACE(a, "&Auml;", "A")%>
   <%a=REPLACE(a, "&Ouml;", "O")%>
   <%a=REPLACE(a, "&Uuml;", "U")%>
   <%a=REPLACE(a, "&Otilde;", "Y")%>
   <%a=REPLACE(a, "&#353;", "Sh")%>
   <%a=REPLACE(a, "&#245;", "y")%>
   <%a=REPLACE(a, "&#228;", "a")%>
   <%a=REPLACE(a, "&#246;", "o")%>
   <%a=REPLACE(a, "&#252;", "u")%>
   <%a=REPLACE(a, "&#213;", "O")%>
   <%a=REPLACE(a, "&#214;", "O")%>
   <%a=REPLACE(a, "&#220;", "U")%>
   <%cc=0%><%ff=0%><%pp=0%>
   <%ServFileStream.Writeline a%>
   <%pc0=""%>
   <%mdbol2.CommandText="SELECT DISTINCT ProjName,StatusName,ProjCode,DateBegin,DateEnd,EmplName,EmplFname,Yearr,SummTot,IDentIfier FROM projserv WHERE Yearr=" & ya & " AND SUBSTRING(ProjCode,4,2)<>'00' AND SUBSTRING(ProjCode,7,2)='00' AND SUBSTRING(ProjCode,10,2)<>'00' AND SUBSTRING(ProjCode,1,2)='" & MID(mdborl1("ProjCode"),1,2) & "' ORDER BY ProjCode"%>
   <%mdborl2.Open mdbol2%>
   <%mdbol2b.CommandText="SELECT SUBSTRING(ProjCode,4,2),SUM(SummTot) as st FROM projserv3 WHERE Yearr='" & ya & "' AND SUBSTRING(ProjCode,4,2)<>'00' AND SUBSTRING(ProjCode,10,2)<>'00' AND SUBSTRING(ProjCode,1,2)='" & MID(mdborl1("ProjCode"),1,2) & "' GROUP BY IdentIfier,SUBSTRING(ProjCode,4,2)"%>
   <%mdborl2b.Open mdbol2b%>
   <%Do Until mdborl2.EOF%>
    <%mdbolc.CommandText="SELECT * FROM projserv WHERE Yearr=" & ya & " AND SUBSTRING(ProjCode,10,2)<>'00' AND SUBSTRING(ProjCode,1,5)='" & MID(mdborl2("ProjCode"),1,5) & "' AND ProjCode<>'" & mdborl2("ProjCode") & "' ORDER BY ProjCode"%>
    <%mdborlc.Open mdbolc%>
    <%If mdborlc.EOF Then%>
     <%grp="Нет"%>
    <%Else%>
     <%grp="Да"%>
    <%End If%>
    <%mdborlc.Close%>
    <%FOR i=1 to len(Mid(mdborl2("ProjCode"),1,5))%>
     <%a=Mid(Mid(mdborl2("ProjCode"),1,5),i,1)%>
     <%If i=1 Then%>
      <%a2="."%>
     <%Else%>
      <%a2=Mid(mdborl2("ProjCode"),i-1,1)%>
     <%End If%>
     <%If a="0" AND (a2="0" OR a2=".") Then%>
     <%Else%>
      <%pc0=pc0 & a%>
     <%End If%>
    <%Next%>
    <%cc=0%>
    <%mdborl2b.MoveNext%>
    <%ff=0%>
    <%mdborl2b.MoveNext%>
    <%pp=0%>
    <%mdborl2b.MoveNext%>
    <%sdr2=sdr2+1%>
    <%str2=sdr & "." & sdr2%>
    <%uid=uid+1%>
    <%a=(pc0 & "|" & mdborl2("ProjName") & "||2|" & grp & "|||" & pp & "|" & ff & "|" & cc & "||" & uid & "|")%>
    <%a=REPLACE(a, "&auml;", "a")%>
    <%a=REPLACE(a, "&ouml;", "o")%>
    <%a=REPLACE(a, "&uuml;", "u")%>
    <%a=REPLACE(a, "&otilde;", "y")%>
    <%a=REPLACE(a, "&Auml;", "A")%>
    <%a=REPLACE(a, "&Ouml;", "O")%>
    <%a=REPLACE(a, "&Uuml;", "U")%>
    <%a=REPLACE(a, "&Otilde;", "Y")%>
    <%a=REPLACE(a, "&#353;", "Sh")%>
    <%a=REPLACE(a, "&#245;", "y")%>
    <%a=REPLACE(a, "&#228;", "a")%>
    <%a=REPLACE(a, "&#246;", "o")%>
    <%a=REPLACE(a, "&#252;", "u")%>
    <%a=REPLACE(a, "&#213;", "O")%>
    <%a=REPLACE(a, "&#214;", "O")%>
    <%a=REPLACE(a, "&#220;", "U")%>
    <%cc=0%><%ff=0%><%pp=0%>
    <%ServFileStream.Writeline a%>
    <%pc0=""%>
    <%'="SELECT distinct pid, ISNULL(SUBSTRING(PROJCODE,10,2),'00'), SUBSTRING(PROJCODE,1,8),projcode,EDescr,rusname as Projname,emplfname,SummTot,emplname FROM projserv WHERE Yearr='" & ya+1 & "' AND  SUBSTRING(projcode,1,5)='" & MID(mdborl2("ProjCode"),1,5) & "' AND SUBSTRING(projcode,7,2)<>'00' ORDER BY SUBSTRING(PROJCODE,1,8),ISNULL(SUBSTRING(PROJCODE,10,2),'00')"%>
    <%mdbo1.CommandText="SELECT distinct pid, ISNULL(SUBSTRING(PROJCODE,10,2),'00'), SUBSTRING(PROJCODE,1,8),projcode,EDescr,rusname as Projname,emplfname,SummTot,emplname FROM projserv WHERE Yearr='" & ya+1 & "' AND  SUBSTRING(projcode,1,5)='" & MID(mdborl2("ProjCode"),1,5) & "' AND SUBSTRING(projcode,7,2)<>'00' ORDER BY SUBSTRING(PROJCODE,1,8),ISNULL(SUBSTRING(PROJCODE,10,2),'00')"%>
    <%mdbor.Open mdbo1%>
    <%Do Until mdbor.EOF%>
     <%mdbo3.CommandText="SELECT ISNULL(SUM(ISNULL(SummaPlan,0)),0) as SP ,ISNULL(SUM(ISNULL(SummaContract,0)),0) as SC,ISNULL(ROUND(SUM(ISNULL(SummaFact,0))/1000,0),0) as SF FROM delta WHERE Yearr=" & ya+1 & " AND EDescr='" & mdbor("EDEscr") & "' AND Pid='" & mdbor("Pid") & "' AND SUBSTRING(ProjCode,10,2)<>'00'"%>
     <%mdbor3.Open mdbo3%>
     <%mdbolc.CommandText="SELECT projcode,projname,emplfname,emplname FROM projserv WHERE Yearr='" & ya+1 & "' AND PID='" & mdbor("PID") & "'"%>
     <%mdborlc.Open mdbolc%>
     <%If LEN(mdbor("ProjCode"))<9 Then%>
      <%grp="Нет"%>
      <%lv="3"%>
     <%Else%>
      <%If mid(mdbor("ProjCode"),10,2) ="00" Then%>
       <%grp="Да"%>
       <%lv="3"%>
      <%Else%>
       <%grp="Нет"%>
       <%lv="4"%>
      <%End If%>
     <%End If%>
     <%FOR i=1 to len(mdbor("ProjCode"))%>
      <%a=Mid(mdbor("ProjCode"),i,1)%>
      <%If i=1 Then%>
       <%a2="."%>
      <%Else%>
       <%a2=Mid(mdbor("ProjCode"),i-1,1)%>
      <%End If%>
      <%If a="0" AND (a2="0" OR a2=".") Then%>
      <%Else%>
       <%pc0=pc0 & a%>
      <%End If%>
     <%Next%>
     <%If mdbor3.EOF="False" Then%>
      <%cc=mdbor3("sc")%>
      <%ff=mdbor3("sf")%>
      <%pp=mdbor3("sp")%>
     <%End If%>
     <%sdr3=sdr3+1%><%str3=str2 & "." & sdr3%><%uid=uid+1%>
     <%a=(pc0 & "|" & mdbor("ProjName") & "|" & mdborlc("emplfname") & " " & mdborlc("emplname") & "|" & lv & "|" & grp & "|||" & pp & "|" & ff & "|" & cc & "|" & mdbor("EDEscr") & "|" & uid & "|")%>
     <%a=REPLACE(a, "&auml;", "a")%>
     <%a=REPLACE(a, "&ouml;", "o")%>
     <%a=REPLACE(a, "&uuml;", "u")%>
     <%a=REPLACE(a, "&otilde;", "y")%>
     <%a=REPLACE(a, "&Auml;", "A")%>
     <%a=REPLACE(a, "&Ouml;", "O")%>
     <%a=REPLACE(a, "&Uuml;", "U")%>
     <%a=REPLACE(a, "&Otilde;", "Y")%>
     <%a=REPLACE(a, "&#353;", "Sh")%>
     <%a=REPLACE(a, "&#245;", "y")%>
     <%a=REPLACE(a, "&#228;", "a")%>
     <%a=REPLACE(a, "&#246;", "o")%>
     <%a=REPLACE(a, "&#252;", "u")%>
     <%a=REPLACE(a, "&#213;", "O")%>
     <%a=REPLACE(a, "&#214;", "O")%>
     <%a=REPLACE(a, "&#220;", "U")%>
     <%a=REPLACE(a, "&#352;", "Sh")%>
     <%cc=0%><%ff=0%><%pp=0%>
<%=a%>
     <%ServFileStream.Writeline a%>
     <%pc0=""%>
     <%mdborlc.Close%>
     <%mdbor.MoveNext%>
     <%mdbor3.Close%>
    <%Loop%>
    <%mdbor.Close%>
    <%mdborl2.MoveNext%>
   <%Loop%>
   <%mdborl2.Close%><%mdborl2b.Close%>
   <%mdborl1.MoveNext%>
  <%Loop%>
  <%mdborl1.Close%><%mdborl1b.Close%>
  <%servFileStream.Close%>
  <hr>
  <p Align="Left"><img src="icons\1a.ico" ></P>CSV-vORming on edukalt loodud!<p Align="Left">
  <img src="icons\2a.ico" ></P>
  N&uuml;&uuml;d v&otilde;ite ANDa makrok&auml;su <a Href="iu.vbs">INVESTUPDATE</a> globaalses moodulis (t&ouml;&ouml;keskkonnas) MS Project Server, v&otilde;i siis ANDa makrok&auml;su <a Href="ii.vbs">INVESTIMPORT</a>, kui teete ANDmev&otilde;ttu teisest s&uuml;steemist, rakEndusest v&otilde;i failivORmingust esimest kORda.
 </Body>
</Html>