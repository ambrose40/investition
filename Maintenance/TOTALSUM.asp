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
<%set mdbo9 = Server.CreateObject("ADODB.Command")%>
<%set mdbor9 = Server.CreateObject("ADODB.Recordset")%>
<%mdbo9.ActiveConnection = mdbo%>
<%set mdbou = Server.CreateObject("ADODB.Command")%>
<%set mdboru = Server.CreateObject("ADODB.Recordset")%>
<%mdbou.ActiveConnection = mdbo%>
<%mdbo8.CommandText="SELECT PROJCODE,PID,Yearr,Enterprise,Identifier FROM MAIN WHERE SUBSTRING(PROJCODE,10,2)='00'"%>
<%mdbor8.Open mdbo8%>
<%Do until mdbor8.EOF%>
<%mdbo9.CommandText="SELECT ISNULL(SUM(ISNULL(PASTSUM,0)),0) as PS FROM MAIN WHERE LEN(PROJCODE)>9  AND SUBSTRING(PROJCODE,10,2)<>'00' AND SUBSTRING(PROJCODE,1,8)='" & MID(mdbor8("PROJCODE"),1,8) & "' and Yearr='" & MDBOR8("Yearr") & "' AND Enterprise='" & MDBOR8("Enterprise") & "' AND IDEntifier='" & MDBOR8("IDentifier") & "'"%>
<%mdbor9.Open mdbo9%>
<%mdbou.CommandText="UPDATE MAIN SET PASTSUM='" & MDBOR9("PS") & "' WHERE PID='" & MDBOR8("PID") & "' and Yearr='" & MDBOR8("Yearr") & "' AND Enterprise='" & MDBOR8("Enterprise") & "' AND IDEntifier='" & MDBOR8("IDentifier") & "'"%>
<%mdboru.Open mdbou%>
<%mdbor8.MoveNExt%>
<%mdbor9.Close%>
<%Loop%>
<%mdbor8.Movefirst%>
<%Do until mdbor8.EOF%>
<%mdbo9.CommandText="SELECT ISNULL(SUM(ISNULL(PROGNTEH,0)),0) as PS FROM MAIN WHERE LEN(PROJCODE)>9  AND SUBSTRING(PROJCODE,10,2)<>'00' AND SUBSTRING(PROJCODE,1,8)='" & MID(mdbor8("PROJCODE"),1,8) & "' and Yearr='" & MDBOR8("Yearr") & "' AND Enterprise='" & MDBOR8("Enterprise") & "' AND IDEntifier='" & MDBOR8("IDentifier") & "'"%>
<%mdbor9.Open mdbo9%>
<%mdbou.CommandText="UPDATE MAIN SET PROGNTEH='" & MDBOR9("PS") & "' WHERE PID='" & MDBOR8("PID") & "' and Yearr='" & MDBOR8("Yearr") & "' AND Enterprise='" & MDBOR8("Enterprise") & "' AND IDEntifier='" & MDBOR8("IDentifier") & "'"%>
<%mdboru.Open mdbou%>
<%mdbor8.MoveNExt%>
<%mdbor9.Close%>
<%Loop%>
<%mdbor8.Movefirst%>
<%Do until mdbor8.EOF%>
<%mdbo9.CommandText="SELECT ISNULL(SUM(ISNULL(IKVARTAL,0)),0) as PS FROM MAIN WHERE LEN(PROJCODE)>9  AND SUBSTRING(PROJCODE,10,2)<>'00' AND SUBSTRING(PROJCODE,1,8)='" & MID(mdbor8("PROJCODE"),1,8) & "' and Yearr='" & MDBOR8("Yearr") & "' AND Enterprise='" & MDBOR8("Enterprise") & "' AND IDEntifier='" & MDBOR8("IDentifier") & "'"%>
<%mdbor9.Open mdbo9%>
<%mdbou.CommandText="UPDATE MAIN SET IKVARTAL='" & MDBOR9("PS") & "' WHERE PID='" & MDBOR8("PID") & "' and Yearr='" & MDBOR8("Yearr") & "' AND Enterprise='" & MDBOR8("Enterprise") & "' AND IDEntifier='" & MDBOR8("IDentifier") & "'"%>
<%mdboru.Open mdbou%>
<%mdbor8.MoveNExt%>
<%mdbor9.Close%>
<%Loop%>
<%mdbor8.Movefirst%>
<%Do until mdbor8.EOF%>
<%mdbo9.CommandText="SELECT ISNULL(SUM(ISNULL(IIKVARTAL,0)),0) as PS FROM MAIN WHERE LEN(PROJCODE)>9  AND SUBSTRING(PROJCODE,10,2)<>'00' AND SUBSTRING(PROJCODE,1,8)='" & MID(mdbor8("PROJCODE"),1,8) & "' and Yearr='" & MDBOR8("Yearr") & "' AND Enterprise='" & MDBOR8("Enterprise") & "' AND IDEntifier='" & MDBOR8("IDentifier") & "'"%>
<%mdbor9.Open mdbo9%>
<%mdbou.CommandText="UPDATE MAIN SET IIKVARTAL='" & MDBOR9("PS") & "' WHERE PID='" & MDBOR8("PID") & "' and Yearr='" & MDBOR8("Yearr") & "' AND Enterprise='" & MDBOR8("Enterprise") & "' AND IDEntifier='" & MDBOR8("IDentifier") & "'"%>
<%mdboru.Open mdbou%>
<%mdbor8.MoveNExt%>
<%mdbor9.Close%>
<%Loop%>
<%mdbor8.Movefirst%>
<%Do until mdbor8.EOF%>
<%mdbo9.CommandText="SELECT ISNULL(SUM(ISNULL(IIIKVARTAL,0)),0) as PS FROM MAIN WHERE LEN(PROJCODE)>9  AND SUBSTRING(PROJCODE,10,2)<>'00' AND SUBSTRING(PROJCODE,1,8)='" & MID(mdbor8("PROJCODE"),1,8) & "' and Yearr='" & MDBOR8("Yearr") & "' AND Enterprise='" & MDBOR8("Enterprise") & "' AND IDEntifier='" & MDBOR8("IDentifier") & "'"%>
<%mdbor9.Open mdbo9%>
<%mdbou.CommandText="UPDATE MAIN SET IIIKVARTAL='" & MDBOR9("PS") & "' WHERE PID='" & MDBOR8("PID") & "' and Yearr='" & MDBOR8("Yearr") & "' AND Enterprise='" & MDBOR8("Enterprise") & "' AND IDEntifier='" & MDBOR8("IDentifier") & "'"%>
<%mdboru.Open mdbou%>
<%mdbor8.MoveNExt%>
<%mdbor9.Close%>
<%Loop%>
<%mdbor8.Movefirst%>
<%Do until mdbor8.EOF%>
<%mdbo9.CommandText="SELECT ISNULL(SUM(ISNULL(IVKVARTAL,0)),0) as PS FROM MAIN WHERE LEN(PROJCODE)>9  AND SUBSTRING(PROJCODE,10,2)<>'00' AND SUBSTRING(PROJCODE,1,8)='" & MID(mdbor8("PROJCODE"),1,8) & "' and Yearr='" & MDBOR8("Yearr") & "' AND Enterprise='" & MDBOR8("Enterprise") & "' AND IDEntifier='" & MDBOR8("IDentifier") & "'"%>
<%mdbor9.Open mdbo9%>
<%mdbou.CommandText="UPDATE MAIN SET IVKVARTAL='" & MDBOR9("PS") & "' WHERE PID='" & MDBOR8("PID") & "' and Yearr='" & MDBOR8("Yearr") & "' AND Enterprise='" & MDBOR8("Enterprise") & "' AND IDEntifier='" & MDBOR8("IDentifier") & "'"%>
<%mdboru.Open mdbou%>
<%mdbor8.MoveNExt%>
<%mdbor9.Close%>
<%Loop%>
<%mdbor8.Close%>
</html>
</body>