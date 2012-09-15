<html>
<body>
  <%b=Server.MapPath("/")%>
  <%Set servcfg=Server.CreateObject("Scripting.FileSystemObject")%>
  <%Set servFileStream=servcfg.OpenTextFile(b & "\310.ttx")%>
  <%Set servcfg2=Server.CreateObject("Scripting.FileSystemObject")%>
  <%Set servFileStream2=servcfg2.CreateTextFile(b & "\410.ttx")%>
<%For i=1 to 368138%>
<%s=servFileStream.ReadLine%>
<%If Instr(1,s,"Yes")=0 and Instr(1,s,"No")=0 and Instr(1,s,chr(34))<>0 then%>
<%servFileStream2.WriteLine s%>
<%End if%>
<%Next%>

  <%servFileStream.Close%>
</body>
</html>