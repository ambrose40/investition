<html>
<body>
  <%b=Server.MapPath("/")%>
  <%Set servcfg=Server.CreateObject("Scripting.FileSystemObject")%>
  <%Set servFileStream=servcfg.OpenTextFile(b & "\410.ttx")%>
  <%Set servcfg2=Server.CreateObject("Scripting.FileSystemObject")%>
  <%Set servFileStream2=servcfg2.CreateTextFile(b & "\510.ttx")%>
<%DIm M(1000,1000)%>
<%k=1%>
<%j=1%>
<%st=0%>
<%For i=1 to 92390%>
  <%s=servFileStream.ReadLine%>
  <%If Instr(1,s,"/")<>0 then%>
    <%For i2=1 to len(s)%>
      <%a=Mid(s,i2,1)%>
      <%If Asc(a)=34 Then%>
        <%For i3=i2+1 to len(s)%>
          <%a2=Mid(s,i3,1)%>
          <%If Asc(a2)=34 Then%>
            <%pr=Mid(s,i2,i3-i2)%>
            <%For i4=1 to 1000%>
            
            <%if pr=M(j,i4) then%>
              <%st=1%>
              <%Exit for%>
              <%End if%>
            <%Next%>
<%If st=0 then%>
            <%M(j,k)=pr%>
            <%k=k+1%>
            <%servFileStream2.WriteLine pr%>
<%'=Pr%>
<%Exit for%>
<%Else%>
<%st=0%>
<%Exit for%>
<%end if%>

          <%End If%>
        <%Next%>
<%Exit for%>
      <%End If%>
    <%Next%>
<%Else%>
<%For i2=1 to len(s)%>
      <%a=Mid(s,i2,1)%>
      <%If Asc(a)=34 Then%>
        <%For i3=i2+1 to len(s)%>
          <%a2=Mid(s,i3,1)%>
          <%If Asc(a2)=34 Then%>
            <%pr=Mid(s,i2,i3-i2)%>
            <%j=j+1%>
            <%servFileStream2.WriteLine pr%>
            <%ent=Chr(13) & Chr(10)%>
            <%servFileStream2.WriteLine ent%>
            <%Exit for%>
          <%End If%>
        <%Next%>
        <%Exit for%>
      <%End If%>
    <%Next%>
  <%End If%>
<%Next%>
<%servFileStream.Close%>
</body>
</html>