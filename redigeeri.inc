<!--Если была нажата кнопка Muuda Seisund (изменить состояние)-->
    <%
    If request.Form("btn")="Muuda seisund" Then
      g=Request.Form("idc")
      a0="st" & g
      a=Request.Form("stchg")
      sl2=correct_est(a)
      mdboen.CommandText="UPDATE StatCode SET StatusName= '" & sl2 & "' WHERE StatusID = '" & request.Form(a0) & "'"
      mdboren.Open mdboen
    End If

'<!--Если была нажата кнопка Muuda kirje (изменить запись)-->
    If request.Form("btn")="Muuda kirje" Then
      g=Request.Form("idc2")
      a0="em" & g
      a=Request.Form("nchg")
      sl=correct_est(a)
      mdboen.CommandText="UPDATE Worker SET EmplName= '" & sl & "' WHERE EmployeeID = '" & request.Form(a0) & "'"
      mdboren.Open mdboen
      a=Request.Form("fnchg")
      sl=correct_est(a)
      mdboen.CommandText="UPDATE Worker SET EmplFName= '" & sl & "' WHERE EmployeeID = '" & request.Form(a0) & "'"
      mdboren.Open mdboen
      a=Request.Form("tchg")
      sl=correct_est(a)
      mdboden.CommandText="UPDATE Worker SET TitleName= '" & sl & "' WHERE EmployeeID = '" & request.Form(a0) & "'"
      mdborden.Open mdboden
    End If
'<!--Если было выбрано и подтверждено удаление статуса из истории состояний (таблица Status), то произвести удалении соответствующей записи для выбранного проекта-->
    If Request.QueryString("deh")<>"" then
      mdboen.CommandText="DELETE Status WHERE HistID = '" & request.QueryString("deh") & "'"
      mdboren.Open mdboen
    End if
'<!--Если была нажата кнопка Annuleeri (Аннулировать)-->
    If request.Form("btn")="Annuleeri" Then
'<!--Выборка всех записей из таблицы контрактов-->
      mdboen.CommandText="SELECT * from Contracts"
      mdboren.Open mdboen
      i=1
      Do until mdboren.EOF
        a1="dl2" & i
        If request.Form(a1)="on" Then
          a2="del2" & i
          a3="coi" & i
'<!--Удаляется записи выбранные к удалению из таблицы перечня заключенных контрактов и из таблицы-связки контрактов между проектами-->
          mdboco.CommandText="DELETE CompData WHERE ContractNo = '" & request.Form(a2) & "' AND CompanyID='" & request.Form(a3) & "' AND Pid='" & Mid(proc,6,5) & "'"
          mdborco.Open mdboco
          mdboco.CommandText="DELETE Contracts WHERE ContractNo = '" & request.Form(a2) & "'"
          mdborco.Open mdboco
        End If
        mdboren.MoveNext
        i=i+1
      Loop
      mdboren.Close
    End If
'<!--Вычисление финансового года-->
    ya=fin_year()
    set mdbo1 = Server.CreateObject("ADODB.Command")
    set mdbor = Server.CreateObject("ADODB.Recordset")
    mdbo1.ActiveConnection = mdbo
    mdbo1.CommandText="SELECT * from InvPlan WHERE Pid = '" & Mid(proc,6,5) & "' AND Yearr='" & Mid(proc,1,4) & "' AND Enterprise='" & Mid(proc,5,1) & "' ORDER BY IDentifier"
    mdbor.Open mdbo1
    %>
    <input type="hidden" value="<%=mdbor("Enterprise")%>" name="entt" class="card">
    <%
    mdboen.CommandText="SELECT * from Codes WHERE Pid = '" & Mid(proc,6,5) & "'"
    mdboren.Open mdboen
    %>
<!--Основные параметры-->
    <table border=1 class="Card">
      <tr class="Card">
        <td class="card">
          <b>Project Code: <Font color="0000FF"><%=mdboren("ProjCode")%></font></b>
        </td>
      </tr>
      <tr class="Card">
        <td class="card">
          <b>Projekti Nimestus: </b><Font color="0000FF"><%=mdboren("ProjName")%></font>
        </td>
      </tr>
    </table>
<!--Прорисовка формы управления редактированием карточки проекта-->
    <Form method="POST" action="ProjCard.asp?pc=<%=proc%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>">
      <input type="submit" value="Vaata" name="btn"  class="card"  onmouseover='window.status="Vaatas Projekti kaart.";'onmouseout='window.status="";'>
      <input type="submit" value="Redigeeri" name="btn"  class="card"  onmouseover='window.status="Muudas Projekti kaart.";'onmouseout='window.status="";'>
      <input type="submit" value="Salvesta" name="btn"   class="card" onmouseover='window.status="Salvestas muutused projekti kaardis.";'onmouseout='window.status="";'> 
      <br>
      <%i=1%>
      <table border=1 class="Card">
<!--Заголовок таблицы с финансовыми данными проекта, открытыми для редактирования-->
        <tr class="card">
          <th class="card">
            Aasta
          </th>
          <th class="card">
            <%If request.QueryString("e")="ne4" Then%>
              <a href="ProjCard.asp?e=<%="ol4"%>&pc=<%=proc%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>"  class="th">Eelmise aastade summa</a>
            <%Else%>
              <a href="ProjCard.asp?e=<%="ne4"%>&pc=<%=proc%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>"  class="th">Eelmise aastade summa</a>
            <%End If%>
          </th>
          <th class="card">
            I Kvartal
          </th>
          <th class="card">
            II Kvartal
          </th>
          <th class="card">
            III Kvartal
          </th>
          <th class="card">
            IV Kvartal
          </th>
          <th class="card">
            Seda aasta Summa
          </th>
          <th class="card">
            Kogutud Summa
          </th>
          <th class="card">
            Projekti paralleel
          </th>
        </tr>
        <%
        i=1
'<!--Проверка на наличие записей в базе данных-->
        If mdbor.EOF=False Then
        %>
          <tr class="card">
            <td class="card" rowspan="55" bordercolor=#CCCCCC>
              <%=mdbor("Yearr")%>
            </td>
          </tr>
        <%Else%>
          <tr class="card">
            <td class="card" colspan="8">
              No records in Database
            </td>
          </tr>
        <%
        End If
'<!--Основная часть (тело) таблицы с финансовыми данными проекта, открытыми для редактирования-->
        Do until mdbor.EOF
        %>
          <tr class="card">
            <td class="card">
              <%If request.QueryString("e")="ne4" Then%>
                <input type="text" value="<%=mdbor("PastSum")%>" name="<%="ps" & i%>"  class="card">
              <%Else%>
                <%=mdbor("PastSum")%>
                <input type="hidden" value="<%=mdbor("PastSum")%>" name="<%="ps" & i%>" class="card">
              <%End If%>
            </td>
            <td class="card">
              <input type="text" value="<%=mdbor("Ikvartal")%>" name="<%="1k" & i%>" size="10" class="card">
            </td>
            <td class="card">
              <input type="text" value="<%=mdbor("IIkvartal")%>" name="<%="2k" & i%>" size="10"  class="card">
            </td>
            <td class="card">
              <input type="text" value="<%=mdbor("IIIkvartal")%>" name="<%="3k" & i%>" size="10" class="card">
            </td>
            <td class="card">
              <input type="text" value="<%=mdbor("IVkvartal")%>" name="<%="4k" & i%>" size="10"  class="card">
            </td>
            <td class="card">
              <%=mdbor("SummYe")%>
            </td>
            <td class="card">
              <%=mdbor("SummTot")%>
            </td>
            <td class="card">
<!--Если прорисовывается сумма по контракту, то добавить ссылке для открытия расширенного списка всех заключенных контрактов-->
              <%
              If mdbor("IDentifier")="C" Then
                If Request.QueryString("nm")="neo" Then
                %>
                  <a href="ProjCard.asp?pc=<%=proc%>&nm=<%="oli"%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>"><%=mdbor("IDtransl")%></a>
                <%Else%>
                  <a href="ProjCard.asp?pc=<%=proc%>&nm=<%="neo"%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>"><%=mdbor("IDtransl")%></a>
                <%
                End If
              Else
              %>
                <%=mdbor("IDtransl")%>
              <%End If%>
            </td>
          </tr>
<!--Список контрактов, открытых по ссылке на сумму по контракту в режиме редактирования-->
          <%
          If mdbor("IDentifier")="C" AND Mid(Request.QueryString("nm"),1,3)="neo" Then
            mdboco.CommandText="SELECT Distinct CompanyId,ContractNo,CompanyName,DateOfConcl,DateOfEnding,EmplFName,EmployeeID,EmplName, SummOfContr from contra WHERE Pid = '" & Mid(proc,6,5) & "' AND ((YEAR(DateofConcl)='" & Mid(proc,1,4) & "' AND MONTH(DateofConcl) BETWEEN 04 AND 12) OR (YEAR(DateofConcl)='" & CDbl(Mid(proc,1,4))+1 & "' AND MONTH(DateofConcl) BETWEEN 01 AND 03))"
            mdborco.Open mdboco
            %>
<!--Заголовок подтаблицы со списком заключенных контрактов для выбранного проекта-->
            <tr class="card">
              <th class="card">
                Kontrakti Number
              </th>
              <th class="card">
                Firma nimetus
              </th>
              <th class="card">
                S&otilde;lmimise kuup&auml;ev
              </th>
              <th class="card">
                L&otilde;ppimise kuup&auml;ev
              </th>
              <th class="card">
                T&ouml;&ouml;taja
              </th>
              <th class="card">
                Summa
              </th>
<!--Кнопка аннулирования помеченных галочкой контрактов-->
              <th class="card">
                <input type="submit" value="Annuleeri"   class="card" name="btn" onmouseover='window.status="Kustutas valitud seisundid Seda Projektist.";'onmouseout='window.status="";'> 
              </th>
            </tr>
<!--Тело подтаблицы со списком заключенных контрактов для выбранного проекта-->
            <%
            mdboren.Close
            jj=1
            mdboen.CommandText="SELECT * from CompID"
            mdboren.Open mdboen
            Do until mdborco.EOF
            %>
              <tr class="card">
                <td class="card">
                  <input type="Text" value="<%=mdborco("ContractNo")%>" name="<%="cnt" & jj%>" class="card">
                  <input type="hidden" value="<%=mdborco("ContractNo")%>" name="<%="chi" & jj%>" class="card">
                </td>
<!--Поле со списком компаний. Выбранной оказывается компания, с которой был заключен контракт-->
                <td class="card">
                  <input type="hidden" value="<%=mdborco("CompanyId")%>" name="<%="coi" & jj%>" class="card">
                  <select Name="<%="cmp" & jj%>" size="1">
                    <option value="<%=mdborco("CompanyId")%>"><%=mdborco("CompanyName")%></option>
                    <%
                    Do until mdboren.EOF
                      If mdborco("CompanyID") <> mdboren("CompanyID") Then
                      %>
                        <option value="<%=mdboren("CompanyId")%>"><%=mdboren("CompanyName")%></option>
                      <%
                      End If
                      mdboren.MoveNext
                    Loop
                    %>
                  </select>
                </td>
                <td class="card">
                  <input type="Text" value="<%=mdborco("DateofConcl")%>" name="<%="dco" & jj%>" class="card">
                </td>
                <td class="card">
                  <input type="Text" value="<%=mdborco("DateOfending")%>" name="<%="dcn" & jj%>" class="card">
                </td>
<!--Поле со списком ответственных работников. Выбранной оказывается работник, который ответственен за контракт-->
                <td class="card">
                  <%           
                  mdboen.CommandText="SELECT * from Worker ORDER BY EmplFName"
                  mdboren.Open mdboen
                  %>
                  <select Name="<%="empl" & jj%>" size="1" class="card">
                    <option value="<%=mdborco("EmployeeId")%>"><%=mdborco("EmplFName")%><%=mdborco("EmplName")%></option>
                    <%Do until mdboren.EOF
                      If mdborco("EmployeeID") <> mdboren("EmployeeID") Then%>
                        <option value="<%=mdboren("EmployeeId")%>"><%=mdboren("EmplFName")%>&nbsp<%=mdboren("EmplName")%></option>
                      <%
                      End If
                      mdboren.MoveNext
                    Loop
                    %>
                  </select>
                </td>
<!--Поле с суммой по контракту-->
                <td class="card">
                  <input type="Text" value="<%=mdborco("SummOfContr")%>" name="<%="suc" & jj%>" class="card">
                </td>
<!--Галка для маркироваки на удаление-->
                <td class="card">
                  <input type="CheckBox" name="<%="dl2" & jj%>" class="card"><input type="Hidden" name="<%="del2" & jj%>" value="<%=Mdborco("ContractNO")%>">
                </td>
              </tr>
              <%
              mdborco.Movenext
              jj=jj+1
            Loop
            mdborco.Close
          End If
          i=i+1
          mdbor.Movenext
        Loop
        %>
      </table>
<!--Таблица со списком состояний открытых для редактирования-->
      <%
      mdbor.Close
'<!--SQL запрос для осуществления выборки состояний для выбранного проекта-->
      mdbo1.CommandText="SELECT * from sta WHERE Pid = '" & Mid(proc,6,5) & "'"
      mdbor.Open mdbo1
      %>
      <br>
      <table border=1 class="card">
<!--Заголовок указанной выше таблицы-->
        <tr class="card">
          <th class="card">
            Seisundi nimetus
          </th>
          <th class="card">
<!--Формирование системы ссылок на динамический календарь, для выбора даты в полуавтоматическом режиме-->
            <%If LEFT(request.QueryString("d"),2)<>"de" Then%>
              <a href="ProjCard.asp?pc=<%=proc%>&d2b=<%=Request.QueryString("d2b")%>&m2b=<%=Request.QueryString("m2b")%>&y2b=<%=Request.QueryString("y2b")%>&d2e=<%=Request.QueryString("d2e")%>&m2e=<%=Request.QueryString("m2e")%>&y2e=<%=Request.QueryString("y2e")%>&d=deb&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>" class="th">Alguse kuup&auml;ev</a>
            <%Else%>
              <a href="ProjCard.asp?pc=<%=proc%>&d2b=<%=Request.QueryString("d2b")%>&m2b=<%=Request.QueryString("m2b")%>&y2b=<%=Request.QueryString("y2b")%>&d2e=<%=Request.QueryString("d2e")%>&m2e=<%=Request.QueryString("m2e")%>&y2e=<%=Request.QueryString("y2e")%>&d=nec&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>" class="th">Alguse kuup&auml;ev</a>
            <%End if%>
          </th>
<!--Формирование системы ссылок на динамический календарь, для выбора даты в полуавтоматическом режиме-->
          <th class="card">
            <%If LEFT(request.QueryString("d"),2)<>"de" Then%>
              <a href="ProjCard.asp?pc=<%=proc%>&d2b=<%=Request.QueryString("d2b")%>&m2b=<%=Request.QueryString("m2b")%>&y2b=<%=Request.QueryString("y2b")%>&d2e=<%=Request.QueryString("d2e")%>&m2e=<%=Request.QueryString("m2e")%>&y2e=<%=Request.QueryString("y2e")%>&d=dee&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>" class="th">L&otilde;ppuse kuup&auml;ev</a>
            <%Else%>
              <a href="ProjCard.asp?pc=<%=proc%>&d2b=<%=Request.QueryString("d2b")%>&m2b=<%=Request.QueryString("m2b")%>&y2b=<%=Request.QueryString("y2b")%>&d2e=<%=Request.QueryString("d2e")%>&m2e=<%=Request.QueryString("m2e")%>&y2e=<%=Request.QueryString("y2e")%>&d=nec&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>" class="th">L&otilde;ppuse kuup&auml;ev</a>
            <%End if%>
          </th>
          <th class="card">
            Viide failile
          </th>
          <th class="card">
            Vastutav T&ouml;&ouml;taja
          </th>
          <th class="card">
            <%a2="del" & i%>
            Kustuta?
          </th>
        </tr>
        <%
        i=1
        mdboren.Close
        Do until mdbor.EOF
        %>
<!--Тело (основная часть) указанной выше таблицы-->
          <tr class="card">
            <td class="card">
              <% 
              mdboen.CommandText="SELECT * from StatCode"
              mdboren.Open mdboen
              %>
              <select size="1" name="<%="st" & i%>" class="card">
                <option value=<%=mdbor("StatusID")%>><%=mdbor("StatusName")%></option>
                <%
                Do until mdboren.EOF
                  If mdboren("StatusID")<>mdbor("StatusID") Then
                  %>
                    <option value=<%=mdboren("StatusID")%>><%=mdboren("StatusName")%></option>
                  <%
                  End If
                  mdboren.movenext
                Loop
                mdboren.Close
                %>
              </select>
                
              <%If request.QueryString("nn")="chg" Then%>
                <a href="ProjCard.asp?pc=<%=proc%>&i=<%=i%>&nn=<%="lev"%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>">#</a>
              <%Else%>
                <a href="ProjCard.asp?pc=<%=proc%>&i=<%=i%>&nn=<%="chg"%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>">#</a>
              <%End If%>
            </td>
            <td class="card">
<!--Проверка была ли выбрана дата из календаря-->
              <%
              If request.QueryString("d2b")<>"" then
'<!--Если да, то вывести выбранную в календаре дату, сфоримировав ее из трех разных переменных, указанных в адресной строке-->
                detb=Request.QueryString("d2b") & "." & Request.QueryString("m2b") & "." & Request.QueryString("y2b")
              Else
'<!--Если нет, то вывести дату сохраненную в базе данных-->
                detb=Mdbor("DateBegin")
              End if
              %>
              <input type="text" value="<%=detb%>" name="<%="datb" & i%>" class="card" size="12">
              <input type="Hidden" name="<%="hih" & i%>" value="<%=Mdbor("HistId")%>">
            </td>
<!--Проверка была ли выбрана дата из календаря-->
            <td class="card">
              <%
              If request.QueryString("d2e")<>"" then
'<!--Если да, то вывести выбранную в календаре дату, сфоримировав ее из трех разных переменных, указанных в адресной строке-->
                dete=Request.QueryString("d2e") & "." & Request.QueryString("m2e") & "." & Request.QueryString("y2e")
              Else
'<!--Если нет, то вывести дату сохраненную в базе данных-->
                dete=Mdbor("DateEnd")
              End if
              %>
              <input type="text" value="<%=dete%>" name="<%="date" & i%>" class="card" size="12">
            </td>
            <td class="card">
              <input type="text" value="<%=Mdbor("LinkToFile")%>" name="<%="ltf" & i%>" class="card" size="13">
            </td>
<!--Вывести список работников, в поле с раксрывающемся списком указанной выше формы-->
            <td class="card">
              <%           
              mdboen.CommandText="SELECT * from Worker ORDER BY EmplFName"
              mdboren.Open mdboen
              %>
              <select size="1" name="<%="em" & i%>" class="card">
                <option value="<%=mdbor("EmployeeID")%>"><%=mdbor("EmplFName")%>&nbsp<%=mdbor("EmplName")%></option>
                <%
                Do until mdboren.EOF
                  If mdboren("EmployeeID")<>mdbor("EmployeeID") Then
                  %>
                    <option value="<%=mdboren("EmployeeID")%>"><%=mdboren("EmplFName")%>&nbsp<%=mdboren("EmplName")%></option>
                  <%
                  End if
                  mdboren.movenext
                Loop
                mdboren.Close                
                %>
              </select>
              <%If request.QueryString("nn")="chg2" Then%>
                <a href="ProjCard.asp?pc=<%=proc%>&i=<%=i%>&nn=<%="lev2"%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>">#</a>
              <%Else%>
                <a href="ProjCard.asp?pc=<%=proc%>&i=<%=i%>&nn=<%="chg2"%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>">#</a>
              <%End If%>
            </td>
<!--Ссылка для удаления записи о состоянии проекта-->
            <td class="card">
              <a href="ProjCard.asp?pc=<%=proc%>&deh=<%=Mdbor("HistID")%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>">Kustuta</a>
            </td>
          </tr>
          <%
          mdbor.Movenext
          h=Request.QueryString("i")
          det=h-i
'<!--Элементы управления (кнопки) формой-->
          If request.QueryString("nn")="chg" and det=0 Then
          %>
            <tr class="card">
              <td class="card">
                <%a0="st" & h%>
                <input type="text" name="stchg" value="Muuda Seisundi nimetus" class="card" ><input type="hidden" name="idc" value="<%=h%>">
              </td>
              <td class="card">
                <input type="submit" name="btn" value="Muuda seisund" class="card">
              </td>
            </tr>
          <%End If%>
          <%If request.QueryString("nn")="chg2" and det=0 Then%>
            <tr class="card">
              <td class="card" class='altActive' id='1' title='T&ouml;&ouml;taja nimi siin'>
                <%a0="st" & h%>
                <input type="text" name="nchg" value="" class="card"><input type="hidden" name="idc2" value="<%=h%>">
              </td>
              <td class="card" class='altActive' id='1' title='T&ouml;&ouml;taja perekonnanimi siin'>
                <%a0="st" & h%>
                <input type="text" name="fnchg" value="" class="card">
              </td>
              <td class="card" class='altActive' id='1' title='T&ouml;&ouml;taja ametikoht siin'>
                <input type="text" name="tchg" value="" class="card">
              </td>
              <td class="card">
                <input type="submit" name="btn" value="Muuda kirje" class="card">
              </td> 
            </tr>
          <%
          End If
          i=i+1
        Loop
        %>
      </table>
    </form>
    <%mdbor.Close%>