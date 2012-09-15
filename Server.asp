<html>
<Head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<title>
Information System for Control of Investition Plan. Invest-IT!on
</title>
</Head>
<body bgcolor="222222">

<META HTTP-EQUIV='Content-Type' Content='text/html; charset=cyrillic-windows'>
<%If Request.Form("sbm")="" Then%>
<img border="0" src="icons/disk.ico" Style=float:Left><p align="center"><a href="Main.asp"><font face="Verdana" Size="5" color="00FF00"><b><u>ADMINISTRATOR LOGIN</font></u></b></a></p><p>
<hr color="00EED00">
<Font face="Verdana" size="4" color="00FF00"> Please enter the system administrator id and password:
 <P>
 <Form method=POST action="server.asp">
  System Administrator ID:
  <input type="text", name="uid1", size="25" style="font-size: 12pt; font-weight: bold; color: #008000; background-color: #FFFFFF" value="<%=Request.ServerVariables("LOGON_USER")%>">
  <p>
  System Administrator password: 
  <input type="password", name="pwd1", size="25" style="font-size: 12pt; font-weight: bold; color: #008000; background-color: #FFFFFF" value="">
  <p>
  <input type="submit", name="sbm", value="Proceed">
 </form>
</Font>
<%Else%>
<%If Request.Form("sbm")="Proceed" Then%>
<%b= Server.MapPath("\")%>
<%set servcfg=Server.CreateObject("Scripting.FileSystemObject")%>
<%set servFileStream=servcfg.OpenTextFile(b & "\server.cfg")
  s=servFileStream.ReadLine
  i=servFileStream.ReadLine
  p=servFileStream.ReadLine
  servFileStream.Close%><p>
<%if Request.Form("uid1")=i and Request.Form("pwd1")=p then%>
<hR>
<font face="Verdana" size="4" color="#00FF00"><b>Please specify the information for your SQL Server:</b>
<Form method="POST" action="server.asp">
<p>
</font>
<font face="Verdana" size="4" color="#00FF00">SQL server name:
<input type="text" name="txt1" size="30" value=<%=s%>>
<p>
User ID:
<input type="text" name="txti" size="30" value=<%=i%>>
<p>
User password:
<input type="password" name="txtp" size="30">
<p>
<input type="Submit" value="Submit and Save" name="sbm">
</form>
<%else%>
<font face="Verdana" size="4" color="#00FF00"><B>Wrong Password or ID!!!</B>
<P>
<%Set Nol=Server.CreateObject("MSWC.NextLink")%>
<a href="<%=Nol.GetNthUrl("Links.cfg", 1)%>">
<font face="Verdana" size="4" color="#00FF00">Return to Main Page</a></font>
<%end if%>
</font>
<%Else%>
<%If  Request.Form("sbm")="Submit and Save" Then%>
<%Set Nol=Server.CreateObject("MSWC.NextLink")%>
<img border="0" src="icons/disk.ico" Style=float:Left><p align="center"><a href="<%=Nol.GetNthURL("Links.cfg", 1)%>"><font face="Verdana" Size="5" color="00FF00"><b><u>SAVING CONFIGURATION</font></u></b></a></p><p>
<hr color="00EE00">
<font face="Verdana" Size="4" color="00FF00">
<b>Saved succesfully!  </b>
<%s=Request.Form("txt1")%>
<%i=Request.Form("txti")%>
<%p=Request.Form("txtp")%>
<%b= Server.MapPath("/")%>
<%set servcfg=Server.CreateObject("Scripting.FileSystemObject")%>
<%set servFileStream=servcfg.createTextFile(b & "/server.cfg")
  servFileStream.WriteLine s
  servFileStream.WriteLine i
  servFileStream.WriteLine p
  servFileStream.Close%><p>
</font>
<%End If%>
<%End If%>
<%End If%>
<hr color="00EED00">
<img src="icons/WRENCH.ICO" Style=float:Left><p align="Center"><font face="Verdana" Size="5" color="00FF00"><b>MAINTENANCE</b></font></p></p>
<%Set Nol=Server.CreateObject("MSWC.NextLink")%>
<a href="<%=Nol.GetNthURL("Links.cfg", 21)%>"><font face="Verdana" size="4" color="#00FF00"><%=Nol.GetNthDescription("Links.cfg",21)%></font></a><br>
<a href="<%=Nol.GetNthURL("Links.cfg", 22)%>"><font face="Verdana" size="4" color="#00FF00"><%=Nol.GetNthDescription("Links.cfg",22)%></font></a>

</body>
</html>