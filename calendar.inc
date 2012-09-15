<!--Скрипт создания календаря-->

      <%
      If Request.QueryString("Me")="" then
        mes=Month(Date())
      Else
        mes=Request.QueryString("Me")
      End if

      If Request.QueryString("Yee")="" then
        ya=Year(Date())
      Else
        ya=Request.QueryString("Yee")
      End if

      mdbo1.CommandText="SELECT * from Mesjaca WHERE MONTHH=" & mes
      mdbor.Open mdbo1
      mn=mdbor("MonthName")
      mdbor.Close
      mdbo1.CommandText="SELECT * from Calendar WHERE MONTHH=" & mes & " AND YEARR=" & ya
      mdbor.Open mdbo1
      %>
      <Table border=1 class="card">
        <tr class="card">
          <th class="card" colspan="7">
            <a href="ProjCard.asp?pc=<%=proc%>&d=<%=Request.QueryString("d")%>&me=<%=CDBL(mes)%>&yee=<%=CDBL(ya)-1%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>" class="th"><<</a>
            <%=mn%>&nbsp<%=ya%>
            <a href="ProjCard.asp?pc=<%=proc%>&d=<%=Request.QueryString("d")%>&me=<%=CDBL(mes)%>&yee=<%=CDBL(ya)+1%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>" class="th">>></a>
          </th>
        </tr>
        <tr class="card">
          <th class="card">ES</tH>
          <th class="card">TE</tH>
          <th class="card">KO</tH>
          <th class="card">NE</tH>
          <th class="card">RE</tH>
          <th class="card">LA</tH>
          <th class="card">P&Uuml;</tH>
        </tr>
          
        <%For i=1 to 5%>
          <tr class="card">
            <%
            j=1
            For j=1 to 7
              If mdbor.EOF=false then
                If CDBL(mdbor("WeekD"))=j then
                  If mdbor("Dayy")=DAY(DATE()) and CDBL(mes)=CDBL(MONTH(DATE())) and CDBL(ya)=CDBL(YEAR(DATE())) then
                  %>
                    <th class="card">
                  <%Else%>
                    <td class="card">
                  <%
                  End If
                  mdd=mdbor("Dayy")
                  If LEN(mdd)<2 then
                    mdd="0" & mdd
                  End if
                  mdm=mdbor("Monthh")
                  If LEN(mdm)<2 then
                    mdm="0" & mdm
                  End if
                  If RIGHT(request.QueryString("d"),1)="e" then
                  %>
                    <a href="ProjCard.asp?pc=<%=proc%>&d2b=<%=Request.QueryString("d2b")%>&m2b=<%=Request.QueryString("m2b")%>&y2b=<%=Request.QueryString("y2b")%>&d=nec&d2e=<%=mdbor("Dayy")%>&m2e=<%=CDBL(mes)%>&y2e=<%=CDBL(ya)%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>">
                      <%If mdbor("Dayy")=DAY(DATE()) and CDBL(mes)=CDBL(MONTH(DATE())) and CDBL(ya)=CDBL(YEAR(DATE())) then%>
                        <Font color="FF0000"><%=mdbor("Dayy")%></font>
                      <%Else%>                  
                        <%=mdbor("Dayy")%>
                      <%End If%>
                    </a>
                  <%
                  End if
                  If RIGHT(request.QueryString("d"),1)="b" then
                  %>
                    <a href="ProjCard.asp?pc=<%=proc%>&d2e=<%=Request.QueryString("d2e")%>&m2e=<%=Request.QueryString("m2e")%>&y2e=<%=Request.QueryString("y2e")%>&d=nec&d2b=<%=mdd%>&m2b=<%=mdm%>&y2b=<%=CDBL(ya)%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>">
                      <%If mdbor("Dayy")=DAY(DATE()) and CDBL(mes)=CDBL(MONTH(DATE())) and CDBL(ya)=CDBL(YEAR(DATE())) then%>
                        <Font color="FF0000"><%=mdbor("Dayy")%></font>
                      <%Else%>                  
                        <%=mdbor("Dayy")%>
                      <%End If%>
                    </a>
                  <%
                  End if
                  mdbor.Movenext
                Else
                %>
                  <td class="card"></td>
                <%
                End if
              Else
              %>
                <td class="card"></td>
              <%
              End if           
            Next
            %>
          </tr>
        <%
        Next
        mdbor.close
        If Request.QueryString("Me")="" then
          mes2=Month(Date())
        Else
          mes2=Request.QueryString("Me")
        End if
        If CDBL(mes)-1<=0 then
          mes=13
        Else
          If CDBL(mes)+1>=13 then
            mes2=0
          End if
        End if
        If mes2-mes<0 then
          mdbo1.CommandText="SELECT * from Mesjaca WHERE MONTHH=" & mes-1 & " OR  MONTHH=" & mes2+1 & " ORDER BY MONTHH DESC"
        Else
          mdbo1.CommandText="SELECT * from Mesjaca WHERE MONTHH=" & mes-1 & " OR  MONTHH=" & mes2+1 & " ORDER BY MONTHH ASC"
        End if
        mdbor.Open mdbo1
        %>
        <tr class="card">
          <%If CDBL(mes)-1<>12 Then%>
            <td class="card" colspan="3">
              <a href="ProjCard.asp?pc=<%=proc%>&d=<%=Request.QueryString("d")%>&me=<%=CDBL(mes)-1%>&yee=<%=CDBL(ya)%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>"><%=mdbor("MonthName")%></a>
            </td>
          <%Else%>
            <td class="card" colspan="3">
              <a href="ProjCard.asp?pc=<%=proc%>&d=<%=Request.QueryString("d")%>&me=<%=CDBL(mes)-1%>&yee=<%=CDBL(ya)-1%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>"><%=mdbor("MonthName")%></a>
            </td>
          <%End If%>       
          <td class="card"><%mdbor.MoveNext%></td>
          <%If CDBL(mes2)+1<>1 Then%>
            <td class="card" colspan="3">
              <a href="ProjCard.asp?pc=<%=proc%>&d=<%=Request.QueryString("d")%>&me=<%=CDBL(mes2)+1%>&yee=<%=CDBL(ya)%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>"><%=mdbor("MonthName")%></a>
            </td>
          <%Else%>
            <td class="card" colspan="3">
              <a href="ProjCard.asp?pc=<%=proc%>&d=<%=Request.QueryString("d")%>&me=<%=CDBL(mes2)+1%>&yee=<%=CDBL(ya)+1%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>"><%=mdbor("MonthName")%></a>
            </td>
          <%End If%>
        </tr>
      </table>