<%
Response.Expires = 0
Response.AddHeader "pragma", "no-cache"
%>
<Html>
<Head>
<meta http-equiv="Content-Type" content="text/html; charset=ibm852">
<title>
InformatsiooniSusteem Investeerimise Kava Teostamise Kontrollimiseks. Invest-IT!on
</title></Head>

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
<%set mdbo8 = Server.CreateObject("ADODB.Command")%>
<%set mdbor8 = Server.CreateObject("ADODB.Recordset")%>
<%mdbo8.ActiveConnection = mdbo%>
<%set mdbou = Server.CreateObject("ADODB.Command")%>
<%set mdboru = Server.CreateObject("ADODB.Recordset")%>
<%mdbou.ActiveConnection = mdbo%>

<%mdbo8.CommandText="SELECT DISTINCT PID, YEARR FROM MAIN ORDER BY PID,YEARR"%>
<%mdbor8.Open mdbo8%>
<%Do until mdbor8.EOF%>
<%If MDBOR8("Pid")=abcde THEN%>
<%mdbor8.MoveNExt%>
<%ELSE%>
<%Abcde=MDBOR8("Pid")%>
<%mdbou.CommandText="UPDATE MAIN SET YEARBG='" & MDBOR8("Yearr") & "' WHERE PID='" & MDBOR8("Pid") & "'"%>
<%mdboru.Open mdbou%>
<%mdbor8.MoveNExt%>
<%END IF%>
<%Loop%>
</html>
</body>