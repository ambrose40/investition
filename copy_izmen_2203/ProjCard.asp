<html>
<!--#include File="server_func.inc"-->
<Head>
<!--Подключение стандартыных процедур соединения с SQL сервером, клиентские функции и настройка заголовка страницы-->
<!--#include File="header.inc"-->
<title>Invest-IT!on: PROJEKTI KAART</title>
<!--#include File="client_func.inc"-->
</Head>
<body class="Card" background="icons/back.gif">
<!--#include File="connection.inc"-->
<!--Проерка параметров адресной строки при (пере)загрузке карточки, и назначение переменным необходимым для работы соответствующих значений-->
<%
  If Request.QueryString("ye")<>"" Then
'<!--Если в адресной строке указан год (ye), то проектная строка (proc) равна "год+предприятие+проект" данные извлекаются из адресной строки-->
    proc=request.QueryString("ye") & request.QueryString("entt") & request.QueryString("pco")
  Else
'<!--Если год не указан в адресной строке-->
    If Request.Form("btn")="Ava" Then
'<!--Если нажата кнопка Ava (открыть), то проектная строка (proc) равна "год+предприятие+проект", данные извлекаются из введенных в форму значений-->
      proc=request.Form("ye") & request.Form("entt") & request.Form("pco")
    Else
'<!--Если кнопка Ava (открыть) не нажата-->
      If Request.Form("btn")="*" Then
'<!--Если кнопка * нажата, то проектная строка (proc) равна "год из формы"+10 символов начиная с пятого из адресной строки (переменная pc)-->
        proc=Request.Form("yir") & MID(request.QueryString("pc"),5,10)
      Else
'<!--Если кнопка * не нажата-->
        proc=request.QueryString("pc")
        srt=Request.QueryString("sr")
        zo=Request.QueryString("y")
        np=Request.QueryString("e3")
        pb=Request.QueryString("em")
        so=Request.QueryString("so")
        co=Request.QueryString("s")
        n=Request.QueryString("no")
'<!--Если нажата кнопка Sisteus (ввод), проектная строка равна пяти первым собственным символам и переменной формы prc2-->
        If request.Form("btn")="Sisestus" Then
          proc= Mid(proc,1,5) & request.Form("prc2")
        End IF
      End IF
    End If
  End IF
%>

<!--Проверка был ли выбран проект при общращении на страницу проектной карточки. Проверяется длина сформированной ранее проектной строки.-->
<!--Если длина проектной строки равна 5, то вывести на экран сообщение о невыбарнности проекта-->
<%If len(proc)=5 Then%>
Projekt ei olnud valitud!
<%Else%>
<!--Если длина проектной строки не равна 5, проводяться основные алгоритмы страницы-->
<p align="Center"><a href="#null" class="Headlink" onClick="confirmClose()">PROJEKTI KAART</a>

<!--Если была нажата одна из кнопок Lisa kirje (Добавить запись), то добавляеться запись о работнике-->
<%
If request.Form("btn")="Lisa kirje" or request.Form("btn2")="Lisa kirje" Then
  set mdboae = Server.CreateObject("ADODB.Command")
  set mdborae = Server.CreateObject("ADODB.Recordset")
  mdboae.ActiveConnection = mdbo
'<!--Исправление эстонских букв-->
  a=Request.Form("en")
  sl=correct_est(a)
  a=Request.Form("efn")
  sl1=correct_est(a)
  a=Request.Form("tn")
  sl2=correct_est(a)
'<!--Собственно добавление-->
  mdboae.CommandText="INSERT INTO Worker (EmplName,EmplFname,TitleName) VALUES ('" & sl & "', '" & sl1 & "', '" & sl2 & "')"
  mdborae.Open mdboae
End If
%>

<!--Запуск хранимой процедуры выборки списка финансовых годов для записей проекта, по уникальному коду-->
<%
set mdboen = Server.CreateObject("ADODB.Command")
set mdboren = Server.CreateObject("ADODB.Recordset")
mdboen.ActiveConnection = mdbo

mdboen.CommandText="EXEC years @PI=" & Mid(proc,6,5)
mdboren.Open mdboen
%>

<!--Прорисовка сложной формы управления проектной карточкой проектной-->
<form method="POST" action="ProjCard.asp?pc=<%=proc%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>">

<!--Формирование списка финансовых годов для выбранного проекта-->
<Select size=1 name="yir">
<%Do Until mdboren.EOF%>
<option value="<%=mdboren("Yearr")%>"><%=mdboren("Yearr")%> m.a.</option>
<%mdboren.Movenext
Loop%>
</select>
<input type="Submit" size="4" name="btn" value="*" class="card">
</form>
<%mdboren.Close%>

</p>

<!--Загрузка названия предприятия по его коду из проектной строки-->
<%
mdboen.CommandText="SELECT * from Enterprise WHERE Enterprise='" & Mid(proc,5,1) & "'"
mdboren.Open mdboen
If Request.Form("btn")<>"Muuda" Then
%>
  <td class="card">
  <%If mdboren.EOF<>"True" Then%>
    <b>Ettev&otilde;te: <font color="0000FF"><%=mdboren("EDescr")%></font></b>
  <%Else%>
    Rohkem Ettev&otilde;te pole.
  <%End If%>
  </td>
<%End If%>
<%mdboren.Close%>

<!--Загрузка справочника предприятий в массив записей mdboren-->
<%

'<!--Формирование проектных строк для ссылок, реализующих проекрутку назад и вперед по проектным карточкам-->
nex=Mid(proc,6,5)
nex=nex+1
pron=Mid(proc,1,5) & nex
nex=nex-2
prov=Mid(proc,1,5) & nex
net=Mid(proc,5,1)
net=net+1
prot=Mid(proc,1,4) & net & Mid(proc,6,5)
If net=1 Then
  net=1
else
  net=net-2
End If
prod=Mid(proc,1,4) & net & Mid(proc,6,5)


'<!--Создание объектов для работы с данными-->
set mdboco = Server.CreateObject("ADODB.Command")
set mdborco = Server.CreateObject("ADODB.Recordset")
mdboco.ActiveConnection = mdbo

'<!--Формирование дат в формате SQL Server(YYYY-MM-DD), на основе полей формы dea и dba, для корректной работы запроса-->
dea=Year(Request.Form("dea")) & "-" & Month(Request.Form("dea")) & "-" & Day(Request.Form("dea"))
dba=Year(Request.Form("dba")) & "-" & Month(Request.Form("dba")) & "-" & Day(Request.Form("dba"))

'<!--Если переменная del в строке адреса равна 1 (выбрано и подтверждено удаление проекта), то вычислив финансовый год удаляем выбранный проект-->
If Request.QueryString("del")="1" Then
  ya=fin_year()
  mdboen.CommandText="DELETE Main WHERE Pid = '" & Mid(proc,6,5) & "' AND yearr='" & Mid(proc,1,4) & "'"
  mdboren.Open mdboen
End If

'<!--Если страница загружает карточку в режиме просмотра, то выполняется прорисовка базовой (на просмотр) карточки проекта-->
If request.QueryString("fd")<>""  or request.Form("btn")="*" or request.Form("btn")="Lisa eritingimus" or request.Form("btn")="Sisestus" or request.Form("btn")="Submit" or request.Form("btn")="Muuda" or request.Form("btn")="Vaata" or (request.Form("btn")="" and request.QueryString("n")="" and request.Form("btn2")="" AND request.QueryString("deh")="" AND request.QueryString("d")="" and Request.Querystring("e")="" and request.QueryString("nn")="" AND request.QueryString("ne")="" and (request.QueryString("nm")="" or request.QueryString("nm")<>"neo" and request.QueryString("nm")<>"oli" )) or request.Form("btn")="Salvesta" or request.Form("btn")="Kohaldama muutused" or request.Form("btn")="Kohaldama" or request.Form("btn")="Ava" Then

'<!--Если при этом была нажата кнопка Kohaldama (подтвердить), то добавить новую запись в таблицу status (Этапы или статусы проекта)-->
  If request.Form("btn")="Kohaldama" Then
    mdboen.CommandText="INSERT INTO Status (StatusID, DateBegin, DateEnd, LinkToFile, EmployeeID) VALUES ('" & Request.Form("sta") & "', '" & dba & "', '" & dea & "', '" & Request.Form("ltfa") & "', '" & Request.Form("ema") & "')"
    mdboren.Open mdboen
'<!--Нахождение последнего кода добавленной записи в таблицу Status-->  
    mdboen.CommandText="SELECT Max(HistId) HistId FROM Status"
    mdboren.Open mdboen
    sch=Mdboren("HistID")
    mdboren.Close
'<!--Добавление связи между проектом и статусом в таблицу StatProj-->  
    mdboen.CommandText="INSERT INTO statProj (HistId, Pid) VALUES ('" & sch & "', '" & Mid(Proc,6,5) & "')"
    mdboren.Open mdboen
  End If

'<!--Если при этом была нажата кнопка Kohaldama muutused (подтвердить изменения), то добавить новую запись в таблицу contracts (Заключенные договора по проекту)-->
  If request.Form("btn")="Kohaldama muutused" Then
'<!--Формирование дат в формате SQL Server(YYYY-MM-DD), на основе полей формы dea и dba, для корректной работы запроса-->
    dc=Month(Request.Form("dcol")) & "." & Day(Request.Form("dcol")) & "." & Year(Request.Form("dcol"))
    de=Month(Request.Form("dcnl")) & "." & Day(Request.Form("dcnl")) & "." & Year(Request.Form("dcnl"))
    mdboen.CommandText="INSERT INTO Contracts (ContractNo, DateOfConcl, DateOfEnding, EmployeeID, SummOfContr) VALUES ('" & Request.Form("cntl") & "', '" & dc & "', '" & de & "', '" & Request.Form("empll") & "', '" & Request.Form("sucl") & "')"
    mdboren.Open mdboen
'<!--Добавление связи между компаниями, контрактами и проектами-->  
    mdboen.CommandText="INSERT INTO CompData (ContractNo, CompanyId, Pid) VALUES ('" & Request.Form("cntl") & "', '" & Request.Form("cmpl") & "', '" & Mid(proc,6,5) & "')"
    mdboren.Open mdboen
  End If

'<!--Если при этом была нажата кнопка Lisa eritingimus (добавить особое условие), то добавить новую запись в таблицу faktproj (Факторы задействованные в проекте)-->
  If request.Form("btn")="Lisa eritingimus" Then
    mdboen.CommandText="INSERT INTO FaktProj (FaktorID,PID,Basis) VALUES(" & Request.Form("eri") & ", " & Mid(proc,6,5) &  ", '" & Request.Form("bas") & "')"
    mdboren.Open mdboen
  End If

'<!--Если подтверждаеться удаление особого условия для проекта (в переменной fd адресной строки присутствует значение), 
'     то удаляется соответствующая запись из таблицы FaktProj. В переменной fd, указан код удаляемого фактора-->
  If request.QueryString("fd")<>"" Then
    mdboen.CommandText="DELETE FROM FaktProj WHERE PID=" & Mid(proc,6,5) &  " AND FaktorID=" & request.QueryString("fd")
    mdboren.Open mdboen
  End If
'<!--Если была нажата кнопка Salvesta (сохранить), то выполнить сохранение изменений в проектной карточке-->  
  If request.Form("btn")="Salvesta" Then

'<!--Обновление главной таблицы (main), данными из проектной карточки. См. функцию в server_func.inc-->  
    Call Update_main(1,"C")
    Call Update_main(2,"F")
    Call Update_main(3,"P")

    mdboen.CommandText="SELECT StatusID from sta WHERE Pid = '" & Mid(proc,6,5) & "'"
    mdboren.Open mdboen


'<!--Обновление вспомогательной таблицы (состояний проекта) при сохранение изменений в проектной карточке-->  
    i=1
    Do until mdboren.EOF
      a0="st" & i
      az="hih" & i
      mdboco.CommandText="UPDATE Status SET StatusID='" & Request.Form(a0) & "' WHERE HistID='" & Request.Form(az) & "'"
      mdborco.Open mdboco
      w=Request.Form("datb" & i)
      dbg=MID(w,4,2) & "/" & MID(w,1,2) & "/" & MID(w,7,4)
      mdboco.CommandText="UPDATE Status SET DateBegin='" & dbg & "' WHERE HistID='" & Request.Form(az) & "'"
      mdborco.Open mdboco
      w=Request.Form("date" & i)
      den=MID(w,4,2) & "/" & MID(w,1,2) & "/" & MID(w,7,4)
      mdboco.CommandText="UPDATE Status SET DateEnd='" & den & "' WHERE HistID='" & Request.Form(az) & "'"
      mdborco.Open mdboco
      a0="ltf" & i
      mdboco.CommandText="UPDATE Status SET LinkToFile='" & Request.Form(a0) & "' WHERE HistID='" & Request.Form(az) & "'"
      mdborco.Open mdboco
      a0="em" & i
      mdboco.CommandText="UPDATE Status SET EmployeeID='" & Request.Form(a0) & "' WHERE HistID='" & Request.Form(az) & "'"
      mdborco.Open mdboco
      mdboren.MoveNext
      i=i+1
    Loop
  
'<!--Обновление вспомогательной таблицы (заключенные договоры) при сохранение изменений в проектной карточке-->  
    i=1

    mdboren.Close
    mdboen.CommandText="SELECT ContractNo from con WHERE Pid = '" & Mid(proc,6,5) & "'"
    mdboren.Open mdboen
'<!--Обновление вспомогательной таблицы (заключенные договоры) при сохранение изменений в проектной карточке-->  
    Do until mdboren.EOF
      a1="cnt" & i
      az="chi" & i
      mdboco.CommandText="UPDATE Contracts SET ContractNo='" & Request.Form(a1) & "' WHERE ContractNo='" & Request.Form(az) & "'"
      mdborco.Open mdboco
      a0="cmp" & i
      a2="coi" & i
      mdboco.CommandText="UPDATE CompData SET CompanyID='" & Request.Form(a0) & "' WHERE CompanyID='" & Request.Form(a2) & "' AND Pid='" &  Mid(proc,5,6) & "' AND ContractNo='" & Request.Form(a1) & "'"
      mdborco.Open mdboco
      a3="dcn" & i
      a4="dco" & i
      dn=Month(Request.Form(a3)) & "." & Day(Request.Form(a3)) & "." & Year(Request.Form(a3))
      mdboco.CommandText="UPDATE Contracts SET DateOfEnding='" & dn & "' WHERE ContractNo='" & Request.Form(a1) & "'"
      mdborco.Open mdboco
      db=Month(Request.Form(a4)) & "." & Day(Request.Form(a4)) & "." & Year(Request.Form(a4))
      mdboco.CommandText="UPDATE Contracts SET DateOfConcl='" & db & "' WHERE ContractNo='" & Request.Form(a1) & "'"
      mdborco.Open mdboco
      a5="empl" & i
      mdboco.CommandText="UPDATE Contracts SET EmployeeId='" & Request.Form(a5) & "' WHERE ContractNo='" & Request.Form(a1) & "'"
      mdborco.Open mdbosv
      a6="suc" & i
      mdboco.CommandText="UPDATE Contracts SET SummOfContr='" & Request.Form(a6) & "' WHERE ContractNo='" & Request.Form(a1) & "'"
      mdborco.Open mdboco
      mdboren.MoveNext
      i=i+1
    Loop
    mdboren.Close
  End If

'<!--Если была нажата кнопка Sisestus (ввод), то сохранить изменные общие данные о проекте-->  
  If request.Form("btn")="Sisestus" Then
    a=Request.Form("prn")
'<!--Кооректировка эстонских символов. См функцию в server_func.inc-->  
    sl=correct_est(a)

'<!--Сохраниеие проектного имени и проектного кода отдела развития-->  
    mdboen.CommandText="UPDATE Codes SET ProjName ='" & sl & "' WHERE Pid='" & Request.Form("prc2") & "'"
    mdboren.Open mdboen
    mdboen.CommandText="UPDATE Main SET ProjCode ='" & Request.Form("prc") & "' WHERE Pid='" & Request.Form("prc2") & "' AND Yearr>='" & Request.Form("pry2") & "' AND Enterprise='" & Request.Form("pre2") & "'"
    mdboren.Open mdboen

'<!--Проверка на существование введеного (измененного) кода Oracle у другого проекта в этот год в базе-->
    mdboen.CommandText="SELECT OracleCode From Main WHERE Yearr='" & Request.Form("pry2") & "' and pid <> '" & Request.Form("prc2") & "'"
    mdboren.Open mdboen
    sta=0
    Do until mdboren.EOF
      If mdboren("OracleCode")=Request.Form("orc") Then
        sta=1
      End If
      mdboren.Movenext
    Loop
    mdboren.Close

'<!--Присвоение кода Oracle на основе результатов сличения или вывод сообщения об ошибке-->
    If sta=0 or Request.Form("orc")="" Then
      If sta=0 then
        oc=Left(Request.Form("orc"),3) & Right(Request.Form("orc"),3)
      Else
        oc="N/A"
      End If
      mdboen.CommandText="UPDATE Main SET OracleCode ='" & oc & "' WHERE Pid='" & Request.Form("prc2") & "' AND Yearr>='" & Request.Form("pry2") & "' AND Enterprise='" & Request.Form("pre2") & "'"
      mdboren.Open mdboen
    Else
%>
      <b>On vaja sisestama jalle unikaalne Oracle kood, sest kood mis te oli sisestanud on sama mis juba oli s&uuml;steemis!</b>
<%
    End If
'<!--Обновление предприятия за которым закреплен проект-->
    mdboen.CommandText="UPDATE Main SET Enterprise ='" & Request.Form("ent") & "' WHERE Pid='" & Request.Form("prc2") & "' AND Yearr='" & Request.Form("pry2") & "' AND Enterprise='" & Request.Form("pre2") & "'"
    mdboren.Open mdboen
    If Request.Form("Ren") = "on" then
      renov=1
    Else
      renov=0
    End If
'<!--Обновление остальных полей (являеться ли проект частью реновации блока, русское наименование, комментарий, примечание(сноска))-->
    mdboen.CommandText="UPDATE Main SET RenovBlock = " & renov & " WHERE Pid='" & Request.Form("prc2") & "' AND Yearr='" & Request.Form("pry2") & "' AND Enterprise='" & Request.Form("pre2") & "'"
    mdboren.Open mdboen
    mdboen.CommandText="UPDATE Main SET RusName = '" & Request.Form("dsc") & "' WHERE Pid='" & Request.Form("prc2") & "' AND Yearr>='" & Request.Form("pry2") & "' AND Enterprise='" & Request.Form("pre2") & "'"
    mdboren.Open mdboen
    mdboen.CommandText="UPDATE Main SET Comment = '" & Request.Form("dsc2") & "' WHERE Pid='" & Request.Form("prc2") & "' AND Yearr>='" & Request.Form("pry2") & "' AND Enterprise='" & Request.Form("pre2") & "'"
    mdboren.Open mdboen
    mdboen.CommandText="UPDATE Main SET FootNote = '" & Request.Form("dscf") & "' WHERE Pid='" & Request.Form("prc2") & "' AND Yearr>='" & Request.Form("pry2") & "' AND Enterprise='" & Request.Form("pre2") & "'"
    mdboren.Open mdboen
  End If



'<!--Алгоритмы формирования карточки на этапе просмотра-->
  set mdbo1 = Server.CreateObject("ADODB.Command")
  set mdbor = Server.CreateObject("ADODB.Recordset")
  mdbo1.ActiveConnection = mdbo
  mdbo1.CommandText="SELECT * from InvPlan WHERE Pid = '" & Mid(proc,6,5) & "' AND Yearr='" & Mid(proc,1,4) & "' AND Enterprise='" & Mid(proc,5,1) & "' ORDER BY IDentifier"
  mdbor.Open mdbo1

'<!--Обновление фактических значений в данном проекте, производится каждый раз при запуске карточки на просмотр-->
  If mdbor.EOF=False Then
    If Len(mdbor("ProjCode"))>9 and MID(mdbor("ProjCode"),10,2)="00" Then
      call Update_main_composite(i, "F")
      call Update_main_composite(i, "C")
      call Update_main_composite(i, "P")
    Else
'<!--Обновление обновлять проектные фактические суммы следует только для карточек проектов существующих после 2004 года-->
      If CDBL(Mid(proc,1,4))>2004 then
        va1=Select_sum("DEBET", "04 AND 06", 0)
        va2=Select_sum("DEBET", "07 AND 09", 0)
        va3=Select_sum("DEBET", "10 AND 12", 0)
        va4=Select_sum("DEBET", "01 AND 03", 1)

        vc1=Select_sum("CREDIT", "04 AND 06", 0)
        vc2=Select_sum("CREDIT", "07 AND 09", 0)
        vc3=Select_sum("CREDIT", "10 AND 12", 0)
        vc4=Select_sum("CREDIT", "01 AND 03", 1)

        mdboen.CommandText="UPDATE Main SET Ikvartal=" & CLNG(CDBL(va1)-CDBL(vc1)) & " WHERE Pid='" & mdbor("Pid") & "' AND Enterprise='" & mdbor("Enterprise") & "' AND Yearr='" & Mid(proc,1,4) & "' AND Identifier='F'"
        mdboren.Open mdboen
        mdboen.CommandText="UPDATE Main SET IIkvartal=" & CLNG(CDBL(va2)-CDBL(vc2)) & " WHERE Pid='" & mdbor("Pid") & "' AND Enterprise='" & mdbor("Enterprise") & "' AND Yearr='" & Mid(proc,1,4) & "' AND Identifier='F'"
        mdboren.Open mdboen
        mdboen.CommandText="UPDATE Main SET IIIkvartal=" & CLNG(CDBL(va3)-CDBL(vc3)) & " WHERE Pid='" & mdbor("Pid") & "' AND Enterprise='" & mdbor("Enterprise") & "' AND Yearr='" & Mid(proc,1,4) & "' AND Identifier='F'"
        mdboren.Open mdboen
        mdboen.CommandText="UPDATE Main SET IVkvartal=" & CLNG(CDBL(va4)-CDBL(vc4)) & " WHERE Pid='" & mdbor("Pid") & "' AND Enterprise='" & mdbor("Enterprise") & "' AND Yearr='" & Mid(proc,1,4) & "' AND Identifier='F'"
        mdboren.Open mdboen
      End If
'<!--Подбивка контрактов для обновления суммы по договору-->
      Call Contract_sum_update("I", "04", "06", 0)
      Call Contract_sum_update("II", "07", "09", 0)
      Call Contract_sum_update("III", "10", "12", 0)
      Call Contract_sum_update("IV", "01", "03", 1)
    End If
  End If

'<!--Выборка записи о проекте из справочника проектов (таблица Codes)-->
  mdboen.CommandText="SELECT * from Codes WHERE Pid = '" & Mid(proc,6,5) & "'"
  mdboren.Open mdboen

'<!--Проверка наличия записи о проекте-->
  If mdboren.EOF="True" Then
  %>
<!--Если нет записи с таким кодом вывести соответствующее сообщение-->
    Seda kirjepaneku Andmebaasis pole.
  <%
  Else
'<!--Если запись найдена то рисовать карточку дальше-->
'<!--В переменную fc заноситься последняя часть кода проекта-->
    fc=Mid(mdboren("ProjCode"),7,6)%>
<!--Прорисовка основной формы (общие данные о проекте)-->
    <Form id="ValidForm" method="POST" action="ProjCard.asp?pc=<%=proc%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>">
      <table border=1 class="card">
<!--Если была нажата кнопка Muuda (изменить), то вывести форму в виде формы для изменения параметров-->
        <%If request.Form("btn")="Muuda" Then%>
          <tr class="Card">
            <td class="card">
              <a href="Projcard.asp?pc=<%=prov%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>"><Img border="0" src="icons/p.ico"></a>
            </td>
            <td class="card">
              <b>Projekti kood: </b>
              <Input Type="text" size="25" name="prc" value="<%=mdbor("ProjCode")%>"  class="card"><Input Type="hidden" name="prc2" value="<%=mdbor("Pid")%>">
              <Input Type="hidden" name="pry2" value="<%=mdbor("Yearr")%>">
            </td>
            <td class="card">
              <b>Oracle kood: </b>
              <Input Type="text" size="25" name="orc" value="<%=mdbor("OracleCode")%>"  class="card"><Input Type="hidden" name="orc2" value="<%=mdbor("OracleCode")%>">
            </td>
            <td class="card">
              <a href="Projcard.asp?pc=<%=pron%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>"><Img border="0" src="icons/n.ico"></a>
            </td>
          </tr>
          <tr class="Card">
            <td class="card">
              <a href="Projcard.asp?pc=<%=prot%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>"><Img border="0" src="icons/u.ico"></a>
            </td>
            <td class="card">
              <b>
              <%If mdbor("RenovBlock")=true Then%>
                Plokkide renoveerimine <input type="checkbox" name="ren" checked="true"  class="card">
              <%Else%>
                  Plokkide renoveerimine <input type="checkbox" name="ren"  class="card">
              <%End If%>
              </b>
            </td>
<!--Записать в массив записей список предприятий-->
            <%
            mdboen.CommandText="SELECT * from Enterprise"
            mdboren.Open mdboen
            %>
<!--Отобразить список предприятий в поле с раскрывающимся списком-->
            <td class="card">
              <b>Ettev&otilde;te: </b>
              <Input Type="hidden" name="pre2" value="<%=mdbor("Enterprise")%>" class="card">
              <select size="1" name="ent" class="card">
<!--Первым в списке поставить предприятие за которым закреплен проект-->
                <option value="<%=mdbor("Enterprise")%>"><%=mdbor("EDescr")%></option>
                <%
                Do Until mdboren.EOF
                  If mdbor("Enterprise")<>mdboren("Enterprise") Then%>
                    <option value="<%=mdboren("Enterprise")%>"><%=mdboren("EDescr")%></option>
                  <%
                  End If
                  mdboren.movenext
                Loop%>
              </select>
            </td>
            <td class="card">
              <a href="Projcard.asp?pc=<%=prod%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>"><Img border="0" src="icons/d.ico"></a>
            </td>
          </tr>
<!--Прорисовка остальных полей формы (Название проекта, русское название, комментарий, примечание)-->
          <tr class="Card">
            <td class="card"></td>
            <td class="card">
              <b>Projekti Nimetus: </b>
              <Input Type="text" size="45" name="prn" value="<%=mdborc("ProjName")%>" class="card"></b><Input Type="hidden" name="prn2" value="<%=mdborc("ProjName")%>">
            </td>
            <td class="card">
              <b>Название: </b>
              <br><Input Type="text" size="45" name="dsc" value="<%=mdbor("RusName")%>" class="card">
            </td>
            <td class="card"></td>
          </tr>
          <tr class="Card">
            <td class="card"></td>
            <td class="card">
              <b>Comment: </b>
              <br><Input Type="text" size="45" name="dsc2" value="<%=mdbor("Comment")%>" class="card">
            </td>
            <td class="card">
              <b>Footnote: </b>
              <br><Input Type="text" size="45" name="dscf" value="<%=mdbor("FootNote")%>" class="card">
            </td>
            <td class="card"></td>
          </tr>
          <input type="submit" value="Sisestus" name="btn" class="card">
<!--Если кнопка muuda не нажата, то запускается прорисовка основной части формы для просмотра-->
        <%
        Else
          If mdbor.EOF=True Then%>
            <tr class="Card">
              <td class="card">
                <a href="Projcard.asp?pc=<%=prov%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>"><Img border="0" src="icons/p.ico"></a>
              </td>
              <td class="card">
                <b>Investeeringute kavas kanne puudub.</b>
              </td>
              <td class="card"></td>
              <td class="card">
                <a href="Projcard.asp?pc=<%=pron%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>"><Img border="0" src="icons/n.ico"></a>
              </td>
            </tr>
            <tr class="Card">
              <td class="card">
                <a href="Projcard.asp?pc=<%=prot%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>"><Img border="0" src="icons/u.ico"></a>
              </td>
              <td class="card">
                <b>Projekti Nimetus: <%=mdborc("ProjName")%></b>
              </td>
              <td class="card">
                <input type="submit" value="Muuda" name="btn" disabled="true" class="card">
              </td>
              <td class="card">
                <a href="Projcard.asp?pc=<%=prod%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>"><Img border="0" src="icons/d.ico"></a>
              </td>
            </tr>
          <%Else%>
            <tr class="Card">
              <td class="card">
                <a href="Projcard.asp?pc=<%=prov%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>"><Img border="0" src="icons/p.ico"></a>
              </td>
              <td class="card">
                <b>Projekti Kood:&nbsp;&nbsp;<FOnt color="0000FF"><%=Mid(proc,6,5)%></font>&nbsp;|&nbsp;<%=mdbor("ProjCode")%></b>
              </td>
              <td class="card">
                <b>Oracle kood: <Font color="0000AA"><%=mdbor("OracleCode")%></b></font>
              </td>
              <td class="card">
                <a href="Projcard.asp?pc=<%=pron%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>"><Img border="0" src="icons/n.ico"></a>
              </td>
            </tr>
            <tr class="Card">
              <td class="card">
                <a href="Projcard.asp?pc=<%=prot%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>"><Img border="0" src="icons/u.ico"></a>
              </td>
              <td class="card">
                <b>Projekti Nimetus: </b><Font color="0000AA"><%=mdboren("ProjName")%></font>
              </td>
              <td class="card">
                <input type="submit" value="Muuda" name="btn"  class="card"  onmouseover='window.status="Muuda Projekti kood, Proecti Nimetus, Oracle kood ja Ettev&otilde;te.";'onmouseout='window.status="";'>
              </td>
              <td class="card">
                <a href="Projcard.asp?pc=<%=prod%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>"><Img border="0" src="icons/d.ico"></a>
              </td>
            </tr>
            <tr class="Card">
              <td class="card"></td>
<!--Указание являеться ли проект частью реновации блока-->
              <td class="card">
                <b>
                <%If mdbor("RenovBlock")=true Then%>
                  <Font color="BB0000">Plokkide renoveerimine</font>
                <%Else%>
                  <Font color="0000BB">Tavaline</font>
                <%End If%>
                </b>
              </td>
              <td class="card">
                <b>Название: </b><Font color="0000AA"><%=mdbor("RusName")%></font>
              </td>
              <td class="card"></td>
            </tr>
            <tr class="Card">
              <td class="card"></td>
              <td class="card">
                <b>Comment: </b><Font color="0000AA"><%=mdbor("Comment")%></font>
              </td>
              <td class="card">
                <b>FootNote: </b><Font color="0000AA"><%=mdbor("Footnote")%></font>
              </td>
              <td class="card"></td>
            </tr>
          <%End If%>
        <%End If%>
      </table>
<!--Далее прорисовка управляющей части карточки проекта (кнопки Vaata (просмотр), Redigeeri (редактирование))-->
      <table border=1 class="Card">
        <tr class="Card">
          <td class="card">
            <%If mdbor.EOF=True Then%>
              <input type="submit" value="Vaata" name="btn" disabled="true" class="card"  onmouseover='window.status="Vaatas Projekti kaart.";'onmouseout='window.status="";'>
              <input type="submit" value="Redigeeri" name="btn" disabled="true" class="card"  onmouseover='window.status="Muudas Projekti kaart.";'onmouseout='window.status="";'>
            <%Else%>
              <input type="submit" value="Vaata" name="btn"  class="card" onmouseover='window.status="Vaatas Projekti kaart.";'onmouseout='window.status="";'>
              <input type="submit" value="Redigeeri" name="btn"  class="card" onmouseover='window.status="Muudas Projekti kaart.";'onmouseout='window.status="";'>
            <%End If%>
          </td>
    </form>
<!--Далее прорисовка управляющей части карточки проекта (кнопки Projekti kustutamine (удаление проекта), Projekti sisestamine (добавление проекта))-->
          <td class="card">
            <Form method="POST" action="insert.asp?pc=<%=proc%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>"  target="_blank">
              <input type="submit" size="10" value="Projekti sisestamine" name="btn"  class="card" onmouseover='window.status="Panes sisse uus Proekt Investetimisse kavasse.";'onmouseout='window.status="";'>
            </Form>
          </td>
          <td class="card">
            <Form id="ValidForm2" method="POST" action="ProjCard.asp?pc=<%=proc%>&del=1&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>">
              <%If mdbor.EOF=True Then%>
                <input type="button" value="Projekti kustutamine" name="btnk" class="card"  onmouseover='window.status="Kustutas projekt investeerimis kavast.";'onmouseout='window.status="";'>
              <%Else%>
                <input type="button" value="Projekti kustutamine" name="btnk" class="card"  onmouseover='window.status="Kustutas projekt investeerimis kavast.";'onmouseout='window.status="";'>
              <%End If%>
            </form>
          </td>
        </tr>
      </table>
  <%End If%>
<!--Далее прорисовка финансовой части проектной карточки-->
  <b>Finantsiline osa</b>
  <br>
<!--Проверка, является ли проект проектной группой-->
  <%If fc<>"00" or (Len(fc)>2 and Mid(fc,4,2)<>"00") Then%>
    <table border=1 class="Card">
      <tr class="card">
        <th class="card">
          Aasta
        </th>
        <th class="card">
          Eelmise aastade summa
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
      <%If mdbor.EOF=False Then%>
        <tr class="Card">
          <td class="card" rowspan="55">
            <%=mdbor("Yearr")%>
          </td>
        </tr>
      <%End If%>
      <%Do until mdbor.EOF%>
        <tr class="Card">
          <td class="card">
            <%=mdbor("PastSum")%>
          </td>
<!--В случае прорисовки значений по контракту к значению в каждом квартале приписать ссылку для просмотра составляющих сумму значений-->
          <td class="card">
            <%If mdbor("IDentifier")="C" Then%>
              <%If Request.QueryString("nm")="neu1" Then%>
                <a href="ProjCard.asp?pc=<%=proc%>&i=<%=i%>&nm=<%="olt1"%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>"><%=mdbor("Ikvartal")%></a>
              <%Else%>
                <a href="ProjCard.asp?pc=<%=proc%>&i=<%=i%>&nm=<%="neu1"%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>"><%=mdbor("Ikvartal")%></a>
              <%End If%>
            <%Else%>
              <%=mdbor("Ikvartal")%>
            <%End If%>
          </td>
          <td class="card">
            <%If mdbor("IDentifier")="C" Then%>
              <%If Request.QueryString("nm")="neu2" Then%>
                <a href="ProjCard.asp?pc=<%=proc%>&i=<%=i%>&nm=<%="olt2"%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>"><%=mdbor("IIkvartal")%></a>
              <%Else%>
                <a href="ProjCard.asp?pc=<%=proc%>&i=<%=i%>&nm=<%="neu2"%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>"><%=mdbor("IIkvartal")%></a>
              <%End If%>
            <%Else%>
              <%=mdbor("IIkvartal")%>
            <%End If%>
          </td>
          <td class="card">
            <%If mdbor("IDentifier")="C" Then%>
              <%If Request.QueryString("nm")="neu3" Then%>
                <a href="ProjCard.asp?pc=<%=proc%>&i=<%=i%>&nm=<%="olt3"%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>"><%=mdbor("IIIkvartal")%></a>
              <%Else%>
                <a href="ProjCard.asp?pc=<%=proc%>&i=<%=i%>&nm=<%="neu3"%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>"><%=mdbor("IIIkvartal")%></a>
              <%End If%>
            <%Else%>
              <%=mdbor("IIIkvartal")%>
            <%End If%>
          </td>
          <td class="card">
            <%If mdbor("IDentifier")="C" Then%>
              <%If Request.QueryString("nm")="neu4" Then%>
                <a href="ProjCard.asp?pc=<%=proc%>&i=<%=i%>&nm=<%="olt4"%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>"><%=mdbor("IVkvartal")%></a>
              <%Else%>
                <a href="ProjCard.asp?pc=<%=proc%>&i=<%=i%>&nm=<%="neu4"%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>"><%=mdbor("IVkvartal")%></a>
              <%End If%>
            <%Else%>
              <%=mdbor("IVkvartal")%>
            <%End If%>
          </td>
          <td class="card">
            <%=mdbor("SummYe")%>
          </td>
          <td class="card">
            <%=mdbor("SummTot")%>
          </td>
          <td class="card">
            <%=mdbor("IDtransl")%>
          </td>
        </tr>
<!--Если был произведен переход по ссылке из значения по контракту, то вывести составляющие сумму контракты за выбранный квартал-->
        <%
'<!--Создание необходимого для указанной выше выборки SQL запроса-->
        If mdbor("IDentifier")="C" AND Mid(Request.QueryString("nm"),1,3)="neu" Then
          If Mid(Request.QueryString("nm"),4,1)="1" Then
            mdboco.CommandText="SELECT Distinct ContractNo,CompanyName,DateOfConcl,DateOfEnding,EmplFName, EmplName, SummOfContr from contra WHERE Pid = '" & Mid(proc,6,5) & "' AND YEAR(DateofConcl)='" & Mid(proc,1,4) & "' AND MONTH(DateOfConcl)>='04' AND MONTH(DateOfConcl)<'07'"
          End If
          If Mid(Request.QueryString("nm"),4,1)="2" Then
            mdboco.CommandText="SELECT Distinct ContractNo,CompanyName,DateOfConcl,DateOfEnding,EmplFName, EmplName, SummOfContr from contra WHERE Pid = '" & Mid(proc,6,5) & "' AND YEAR(DateofConcl)='" & Mid(proc,1,4) & "' AND MONTH(DateOfConcl)>='07' AND MONTH(DateOfConcl)<'10'"
          End If
          If Mid(Request.QueryString("nm"),4,1)="3" Then
            mdboco.CommandText="SELECT Distinct ContractNo,CompanyName,DateOfConcl,DateOfEnding,EmplFName, EmplName, SummOfContr from contra WHERE Pid = '" & Mid(proc,6,5) & "' AND YEAR(DateofConcl)='" & Mid(proc,1,4) & "' AND MONTH(DateOfConcl)>='10' AND MONTH(DateOfConcl)<'12'"
          End If
          If Mid(Request.QueryString("nm"),4,1)="4" Then
            mdboco.CommandText="SELECT Distinct ContractNo,CompanyName,DateOfConcl,DateOfEnding,EmplFName, EmplName, SummOfContr from contra WHERE Pid = '" & Mid(proc,6,5) & "' AND YEAR(DateofConcl)='" & Cdbl(Mid(proc,1,4))+1 & "' AND MONTH(DateOfConcl)>='01' AND MONTH(DateOfConcl)<'04'"
          End If
          mdborco.Open mdboco
          %>
<!--Прорисовка подтаблицы со суммами и другими данными контрактов-->
          <tr class="Card">
<!--Загловок таблицы-->
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
          </tr>
'<!--Прокрутка массива с записями и прорисовка строчек указанной выше таблицы-->
          <%Do until mdborco.EOF%>
            <tr class="Card">
              <td class="card">
                <%=mdborco("ContractNo")%>
              </td>
              <td class="card">
                <%=mdborco("CompanyName")%>
              </td>
              <td class="card">
                <%=mdborco("DateofConcl")%>
              </td>
              <td class="card">
                <%=mdborco("DateOfending")%>
              </td>
              <td class="card">
                <%=mdborco("EmplFName")%>&nbsp<%=mdborco("EmplName")%>
              </td>
              <td class="card">
                <%=mdborco("SummOfContr")%>
              </td>
            </tr>
            <%
            mdborco.Movenext
          Loop
          mdborco.Close
        End If
        mdbor.Movenext
      Loop
      Mdboren.Close
      mdboen.CommandText="SELECT * from sta WHERE Pid = '" & Mid(proc,6,5) & "'"
      mdboren.Open mdboen%>
    </table>
<!--Далее идет прорисовка вспомогательной части проектной карточки-->
<!--Прорисовка таблицы состояний проекта-->
    <b>Seisundid</b>
    <Form method="POST" action="ProjCard.asp?pc=<%=proc%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>">
      <input type="submit" value="Lisa Kontrakt" name="btn"  class="card" onmouseover='window.status="Lisa Kontrakt Andmebaasile.";'onmouseout='window.status="";'>
      <input type="submit" value="Lisa" name="btn"   class="card" onmouseover='window.status="Lisa seisundid Projektile.";'onmouseout='window.status="";'>
    </form>
    <table border=1 class="Card">
<!--Заголовок указанной выше таблицы-->
      <tr class="card">
        <th class="card">
          Seisundi Nimetus
        </th>
        <th class="card">
          Alguse kuup&auml;ev
        </th>
        <th class="card">
          L&otilde;ppimise kuup&auml;ev
        </th>
        <th class="card">
          Viide Failile
        </th>
        <th class="card">
          T&ouml;&ouml;taja perekonnanimi
        </th>
        <th class="card">
          T&ouml;&ouml;taja eesnimi
        </th>
        <th class="card">
          T&ouml;&ouml;taja ametikoht
        </th>
      </tr>
<!--Записи состояний проекта-->
      <%Do until mdboren.EOF%>
        <tr class="Card">
          <td class="card">
            <%If mdboren("StatusID")="6" or mdboren("StatusID")="7" Then%>
              !
            <%End If%>
            <%=mdboren("StatusName")%>
          </td>
          <td class="card">
            <%=mdboren("DateBegin")%>
          </td>
          <td class="card">
            <%=mdboren("DateEnd")%>
          </td>
          <td class="card">
            <%=mdboren("LinktoFile")%>
          </td>
          <td class="card">
            <%=mdboren("EmplFname")%>
          </td>
          <td class="card">
            <%=mdboren("EmplName")%>
          </td>
          <td class="card">
            <%=mdboren("Titlename")%>
          </td>
        </tr>
        <%
        mdboren.Movenext
      Loop
      mdboren.Close
      %>
    </table>
    <br>
<!--Прорисовка таблицы с особыми условиями для проекта-->
    <table border=1 class="Card">
        <%
        jed=1     
        mdboen.CommandText="SELECT * FROM tingimused WHERE PID='" & Mid(proc,6,5) & "'"
        mdboren.Open mdboen
        %>
<!--Заголовок указанной выше таблицы-->
        <tr class="Card">
          <th class="card">
            Eritingimused
          </th>
          <th class="card">
            P&otilde;hjendus
          </th>
          <th class="card"></th>
        </tr>
<!--Основная часть указанной выше таблицы-->
        <%Do until mdboren.EOF%>
          <tr class="card">
            <td class="card">
              <%=mdboren("FaktorName")%>
            </td>
            <td class="card">
              <%=mdboren("Basis")%>
            </td>
            <td class="card">
              <a href="ProjCard.asp?pc=<%=proc%>&fd=<%=mdboren("FaktorID")%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>">Kustuta</a>
            </td>
            <%
            jed=jed+1
            mdboren.Movenext
            %>
          </tr>
        <%
        Loop
        mdboren.Close
        %>
<!--Форма для добавления новых особый условий касательно выбранного проекта-->      
      <form action="ProjCard.asp?pc=<%=proc%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>" method="POST">
        <tr class="Card">
          <td class="card">
            <select size="1" name="eri" class="Card">
              <%
              mdboen.CommandText="SELECT * FROM Faktory"
              mdboren.Open mdboen
              Do until mdboren.EOF
              %>
                <option value="<%=mdboren("FaktorID")%>"><%=mdboren("FaktorName")%></option>
                <%
                mdboren.MovenExt
              Loop
              %>
            </select>
          </td>
          <td class="card" rowspan="1">
            <input type="Text" name="bas" value="" class="card">
          </td>
          <td class="card" colspan="1">
            <input type="Submit" name="btn" value="Lisa eritingimus" class="card">
          </td>
        </tr>
    </Form>
      </table>
  <%Else%>
<!--Если проект является группой вывести об этом соответствующее сообщение-->      
    Puudub juurdep&auml;&auml;s Projekti Grupi projekti kaardile.
  <%
  End If
  mdboren.Close
Else
'<!--Если страница загружает карточку в режиме редактирования, то выполняется прорисовка карточки проекта для редактирования-->
  If request.Querystring("deh")<>"" or request.Querystring("nn")<>"" or LEFT(request.Querystring("d"),2)="de" or request.Querystring("d")="nec" or request.QueryString("nm")="neo" or request.Querystring("nm")="oli" Or request.Form("btn")="Muuda seisund" or request.Form("btn")="Muuda kirje" or request.Form("btn")="Redigeeri" or request.Form("btn")="Kustuta" Or request.Form("btn")="Annuleeri" Or Request.QueryString("e")<>"" Then
  %>
<!--#include File="redigeeri.inc"-->  
<!--Кнопка удаления проекта в инвестиционном плане-->
    <Form id="ValidForm3" method="POST" action="ProjCard.asp?pc=<%=proc%>&del=1&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>">
      <input type="button" value="Projekti kustutamine" name="btnk2"  class="card"  onmouseover='window.status="Kustutas projekt investeerimis kavast.";'onmouseout='window.status="";'>
    </form>
    <%
<!--Если был вызван календарь по соответствующей ссылке, то отобразить календарь из скрипта-->  
    If LEFT(request.QueryString("d"),1)="d" then
    %>
<!--#include File="calendar.inc"-->  
    <%
    End if
  Else
    If Request.QueryString("del")="1" then%>
<!--Если было выбрано и подтверждено удаление проекта из инвестиционного плана, то выпонить процедуру удаления-->  
      <Form method="POST" action="ProjCard.asp?pc=<%=proc%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&e=<%=pb%>&e3=<%=np%>">
        <input type="submit" value="Vaata" name="btn"  class="card">
        <input type="submit" value="Redigeeri" name="btn"  class="card">
      </Form>
      <%
'<!--Запуск команды SQL для удаления соответствующей проекту и году записи из базы (таблица Main)-->  
      mdboen.CommandText="DELETE Main WHERE Pid = '" & Mid(proc,6,5) & "' AND yearr='" & Mid(proc,1,4) & "'"
      mdboren.Open mdboen
    Else
'<!--Обработка дополнительного модуля (Добавление состояния) проектной карточки-->  
      If request.QueryString("n")="ol2" or request.QueryString("n")="ne2" or request.QueryString("n")="old" or request.QueryString("n")="new" or LEFT(request.QueryString("d"),2)="da" or request.QueryString("d")="noc" or request.Form("btn")="Lisa" or request.Form("btn")="Lisa kirje" or request.Form("btn")="Lisa Seisund" Then
      %>

        <Form method="POST" action="ProjCard.asp?pc=<%=proc%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&e=<%=pb%>&e3=<%=np%>">
          <input type="submit" value="Vaata" name="btn" class="card">
          <input type="submit" value="Redigeeri" name="btn" class="card">
          <input type="submit" value="Kohaldama" name="btn" class="card">
<!--Если была нажата кнопка Lisa Seisund (Добавить состояние в справочник состояний)-->  
          <%
          If request.Form("btn")="Lisa Seisund" Then
            a=Request.Form("sn")
            sl2=crrect_est(a)
            mdboen.CommandText="INSERT INTO StatCode (StatusID,StatusName) VALUES ('" & request.Form("sc") & "', '" & sl2 & "')"
            mdboras.Open mdboas
          End If
          %>
<%set mdbos = Server.CreateObject("ADODB.Command")%>
<%set mdbors = Server.CreateObject("ADODB.Recordset")%>
<%mdbos.ActiveConnection = mdbo%>
<%mdbos.CommandText="SELECT * FROM StatCode"%>
<%mdbors.Open mdbos%>
              
<%set mdbow = Server.CreateObject("ADODB.Command")%>
<%set mdborw = Server.CreateObject("ADODB.Recordset")%>
<%mdbow.ActiveConnection = mdbo%>
<%mdbow.CommandText="SELECT * FROM Worker ORDER BY EmplFname"%>
<%mdborw.Open mdbow%>
             
          <table border=1 class="card">
            <tr class="card">
<!--Прорисовка заголовка таблицы добавления состояний-->  
<!--В заголовках формируются ссылки по сложной системе сохранения введенных данных в ячейках, при внесении изменений в справочники или при использовании календаря-->  
              <th class="card">
                <%
                nn1=request.QueryString("n1")
                If nn1 & "e" ="e" then
                  nn1=0
                End if
                If LEFT(request.QueryString("d"),2)<>"da" Then
                %>
                  <a id="link1" href="ProjCard.asp?pc=<%=proc%>&d1b=<%=Request.QueryString("d1b")%>&m1b=<%=Request.QueryString("m1b")%>&y1b=<%=Request.QueryString("y1b")%>&d1e=<%=Request.QueryString("d1e")%>&m1e=<%=Request.QueryString("m1e")%>&y1e=<%=Request.QueryString("y1e")%>&d=dab&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&s2=<%=request.QueryString("s2")%>&n1=<%=nn1%>" class="th">Alguse kuup&auml;ev</a>
                <%Else%>
                  <a id="link1" href="ProjCard.asp?pc=<%=proc%>&d1b=<%=Request.QueryString("d1b")%>&m1b=<%=Request.QueryString("m1b")%>&y1b=<%=Request.QueryString("y1b")%>&d1e=<%=Request.QueryString("d1e")%>&m1e=<%=Request.QueryString("m1e")%>&y1e=<%=Request.QueryString("y1e")%>&d=noc&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&s2=<%=request.QueryString("s2")%>&n1=<%=nn1%>" class="th">Alguse kuup&auml;ev</a>
                <%End if%>
              </th>
              <th class="card">
                <%If LEFT(request.QueryString("d"),2)<>"da" Then%>
                  <a href="ProjCard.asp?pc=<%=proc%>&d1b=<%=Request.QueryString("d1b")%>&m1b=<%=Request.QueryString("m1b")%>&y1b=<%=Request.QueryString("y1b")%>&d1e=<%=Request.QueryString("d1e")%>&m1e=<%=Request.QueryString("m1e")%>&y1e=<%=Request.QueryString("y1e")%>&d=dae&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&s2=<%=request.QueryString("s2")%>&n1=<%=nn1%>" class="th">L&otilde;ppuse kuup&auml;ev</a>
                <%Else%>
                  <a href="ProjCard.asp?pc=<%=proc%>&d1b=<%=Request.QueryString("d1b")%>&m1b=<%=Request.QueryString("m1b")%>&y1b=<%=Request.QueryString("y1b")%>&d1e=<%=Request.QueryString("d1e")%>&m1e=<%=Request.QueryString("m1e")%>&y1e=<%=Request.QueryString("y1e")%>&d=noc&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&s2=<%=request.QueryString("s2")%>&n1=<%=nn1%>" class="th">L&otilde;ppuse kuup&auml;ev</a>
                <%End if%>
              </th>
              <th class="card">
                <%If request.QueryString("n")="ne2" Then%>
                  <a href="ProjCard.asp?pc=<%=proc%>&d1b=<%=Request.QueryString("d1b")%>&m1b=<%=Request.QueryString("m1b")%>&y1b=<%=Request.QueryString("y1b")%>&d1e=<%=Request.QueryString("d1e")%>&m1e=<%=Request.QueryString("m1e")%>&y1e=<%=Request.QueryString("y1e")%>&n=<%="ol2"%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&s2=<%=request.QueryString("s2")%>&n1=<%=nn1%>" class="th">Seisundi Nimetus</a>
                <%Else%>
                  <a href="ProjCard.asp?pc=<%=proc%>&d1b=<%=Request.QueryString("d1b")%>&m1b=<%=Request.QueryString("m1b")%>&y1b=<%=Request.QueryString("y1b")%>&d1e=<%=Request.QueryString("d1e")%>&m1e=<%=Request.QueryString("m1e")%>&y1e=<%=Request.QueryString("y1e")%>&n=<%="ne2"%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&s2=<%=request.QueryString("s2")%>&n1=<%=nn1%>" class="th">Seisundi Nimetus</a>
                <%End If%>
              </th>
              <th class="card">
                <%If request.QueryString("n")="new" Then%>
                  <a href="ProjCard.asp?pc=<%=proc%>&d1b=<%=Request.QueryString("d1b")%>&m1b=<%=Request.QueryString("m1b")%>&y1b=<%=Request.QueryString("y1b")%>&d1e=<%=Request.QueryString("d1e")%>&m1e=<%=Request.QueryString("m1e")%>&y1e=<%=Request.QueryString("y1e")%>&n=<%="old"%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&s2=<%=request.QueryString("s2")%>&n1=<%=nn1%>" class="th">Vastutav t&ouml;&ouml;taja</a>
                <%Else%>
                  <a href="ProjCard.asp?pc=<%=proc%>&d1b=<%=Request.QueryString("d1b")%>&m1b=<%=Request.QueryString("m1b")%>&y1b=<%=Request.QueryString("y1b")%>&d1e=<%=Request.QueryString("d1e")%>&m1e=<%=Request.QueryString("m1e")%>&y1e=<%=Request.QueryString("y1e")%>&n=<%="new"%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&s2=<%=request.QueryString("s2")%>&n1=<%=nn1%>" class="th">Vastutav t&ouml;&ouml;taja</a>
                <%End If%>
              </th>
<!--Необязательное поле, ссылки на связанный с состоянием документ-->  
              <th class="card">
                Viide failile
              </th>
            </tr>
            <tr class="card">
<!--Проверка была ли дата выбрана из календаря. Если да, то сшить дату из трех переменных передаваемых календарем-->  
              <td class="card">
                <%
                If Request.QueryString("d1b")="" then
                  Dat1b=Date()
                Else
                  Dat1b=Request.QueryString("d1b") & "." & Request.QueryString("m1b") & "." & Request.QueryString("y1b")
                End if
                %>
                <input type="text" value="<%=Dat1b%>" name="dba" class="card" size=12>
              </td>
<!--Проверка была ли дата выбрана из календаря. Если да, то сшить дату из трех переменных передаваемых календарем-->  
              <td class="card">
                <%
                If Request.QueryString("d1e")="" then
                  Dat1e=Date()
                Else
                  Dat1e=Request.QueryString("d1e") & "." & Request.QueryString("m1e") & "." & Request.QueryString("y1e")
                End if%>
                <input type="text" value="<%=Dat1e%>" name="dea" class="card" size=12>
              </td>
              <td class="card">
                <%
                If LEFT(request.QueryString("d"),2)<>"da" Then
                  onch="change('ProjCard.asp?pc=" & proc & "&d1b=" & Request.QueryString("d1b") & "&m1b=" & Request.QueryString("m1b") & "&y1b=" & Request.QueryString("y1b") & "&d1e=" & Request.QueryString("d1e") & "&m1e=" & Request.QueryString("m1e") & "&y1e=" & Request.QueryString("y1e") & "&d=dab&sr=" & srt & dd & "&no=" & n & "&y=" & zo & " &s=" & co & "&em=" & pb & "&e3=" & np & "&s2=' + sta.options[sta.selectedIndex].value + '&n1=' + ema.options[ema.selectedIndex].value,'ProjCard.asp?pc=" & proc & "&d1b=" & Request.QueryString("d1b") & "&m1b=" & Request.QueryString("m1b") & "&y1b=" & Request.QueryString("y1b") & "&d1e=" & Request.QueryString("d1e") & "&m1e=" & Request.QueryString("m1e") & "&y1e=" & Request.QueryString("y1e") & "&d=dae&sr=" & srt & dd & "&no=" & n & "&y=" & zo & " &s=" & co & "&em=" & pb & "&e3=" & np & "&s2=' + sta.options[sta.selectedIndex].value + '&n1=' + ema.options[ema.selectedIndex].value"
                Else
                  onch="change('ProjCard.asp?pc=" & proc & "&d1b=" & Request.QueryString("d1b") & "&m1b=" & Request.QueryString("m1b") & "&y1b=" & Request.QueryString("y1b") & "&d1e=" & Request.QueryString("d1e") & "&m1e=" & Request.QueryString("m1e") & "&y1e=" & Request.QueryString("y1e") & "&d=noc&sr=" & srt & dd & "&no=" & n & "&y=" & zo & " &s=" & co & "&em=" & pb & "&e3=" & np & "&s2=' + sta.options[sta.selectedIndex].value + '&n1=' + ema.options[ema.selectedIndex].value,'ProjCard.asp?pc=" & proc & "&d1b=" & Request.QueryString("d1b") & "&m1b=" & Request.QueryString("m1b") & "&y1b=" & Request.QueryString("y1b") & "&d1e=" & Request.QueryString("d1e") & "&m1e=" & Request.QueryString("m1e") & "&y1e=" & Request.QueryString("y1e") & "&d=noc&sr=" & srt & dd & "&no=" & n & "&y=" & zo & " &s=" & co & "&em=" & pb & "&e3=" & np & "&s2=' + sta.options[sta.selectedIndex].value + '&n1=' + ema.options[ema.selectedIndex].value"
                End if
                If request.QueryString("n")="ne2" Then
                  onch=onch & ",'ProjCard.asp?pc=" & proc & "&d1b=" & Request.QueryString("d1b") & "&m1b=" & Request.QueryString("m1b") & "&y1b=" & Request.QueryString("y1b") & "&d1e=" & Request.QueryString("d1e") & "&m1e=" & Request.QueryString("m1e") & "&y1e=" & Request.QueryString("y1e") & "&n=ol2&sr=" & srt & dd & "&no=" & n & "&y=" & zo & " &s=" & co & "&em=" & pb & "&e3=" & np & "&s2=' + sta.options[sta.selectedIndex].value + '&n1=' + ema.options[ema.selectedIndex].value"
                Else
<%onch=onch & ",'ProjCard.asp?pc=" & proc & "&d1b=" & Request.QueryString("d1b") & "&m1b=" & Request.QueryString("m1b") & "&y1b=" & Request.QueryString("y1b") & "&d1e=" & Request.QueryString("d1e") & "&m1e=" & Request.QueryString("m1e") & "&y1e=" & Request.QueryString("y1e") & "&n=ne2&sr=" & srt & dd & "&no=" & n & "&y=" & zo & " &s=" & co & "&em=" & pb & "&e3=" & np & "&s2=' + sta.options[sta.selectedIndex].value + '&n1=' + ema.options[ema.selectedIndex].value"%>
<%End If%>

<%If request.QueryString("n")="new" Then%>
<%onch=onch & ",'ProjCard.asp?pc=" & proc & "&d1b=" & Request.QueryString("d1b") & "&m1b=" & Request.QueryString("m1b") & "&y1b=" & Request.QueryString("y1b") & "&d1e=" & Request.QueryString("d1e") & "&m1e=" & Request.QueryString("m1e") & "&y1e=" & Request.QueryString("y1e") & "&n=old&sr=" & srt & dd & "&no=" & n & "&y=" & zo & " &s=" & co & "&em=" & pb & "&e3=" & np & "&s2=' + sta.options[sta.selectedIndex].value + '&n1=' + ema.options[ema.selectedIndex].value)"%>
<%Else%>
<%onch=onch & ",'ProjCard.asp?pc=" & proc & "&d1b=" & Request.QueryString("d1b") & "&m1b=" & Request.QueryString("m1b") & "&y1b=" & Request.QueryString("y1b") & "&d1e=" & Request.QueryString("d1e") & "&m1e=" & Request.QueryString("m1e") & "&y1e=" & Request.QueryString("y1e") & "&n=new&sr=" & srt & dd & "&no=" & n & "&y=" & zo & " &s=" & co & "&em=" & pb & "&e3=" & np & "&s2=' + sta.options[sta.selectedIndex].value + '&n1=' + ema.options[ema.selectedIndex].value)"%>
<%End If%>

<%'=onch%>

<select size="1" name="sta" class="card" onChange="<%=onch%>">
<%Do until mdbors.EOF%>
<%if request.QueryString("s2")=mdbors("StatusID") then%>
<option value=<%=mdbors("StatusID")%> selected="true"><%=mdbors("StatusName")%></option>
<%Else%>
<option value=<%=mdbors("StatusID")%>><%=mdbors("StatusName")%></option>
<%End if%>
<%mdbors.movenext%>
<%Loop%>
</select>
</td>
<td class="card">
<select size="1" name="ema" class="card" onChange="<%=onch%>">
<%Do until mdborw.EOF%>
<%if CDBL(nn1)<>CDBL(mdborw("EmployeeId")) then%>
<option value="<%=mdborw("EmployeeID")%>"><%=mdborw("EmplFName")%>&nbsp<%=mdborw("EmplName")%></option>
<%Else%>
<option value="<%=mdborw("EmployeeID")%>" selected="true"><%=mdborw("EmplFName")%>&nbsp<%=mdborw("EmplName")%></option>
<%End if%>
<%mdborw.movenext%>
<%Loop%>
</select>
</td>
<td class="card">
<input type="text" value="" name="ltfa" class="card" size=15>
</td>
</tr>
<a name="totj"></a>
<%If request.QueryString("n")="new" Then%>
<tr class="card">
<th class="card">
Isikukood
</th>
<th class="card">
Tootaja nimi
</th>
<th class="card">
Tootaja perekonnanimi
</th>
<th class="card">
Tootaja ametikoht
</th>
<th class="card">
</th>
</tr>

<tr class="card">
<td class="card">
</td>
<td class="card">
<input type="text" value="" name="En" class="card">
</td>
<td class="card">
<input type="text" value="" name="Efn" class="card">
</td>
<td class="card">
<input type="text" value="" name="tn" class="card">
</td>
<td class="card">
<input type="submit" value="Lisa kirje" name="btn" class="card">
</td>
</tr>



<%End If%>

<%If request.QueryString("n")="ne2" Then%>
<tr class="card">
<th class="card">
Seisundi kood
</th>
<th class="card">
Seisundi nimetus
</th>
<th class="card">
</th>
<tr/>
<tr class="card">
<td class="card">
<input type="text" value="" name="sc" class="card">
</td>
<td class="card">
<input type="text" value="" name="sn" class="card">
</td>
<td class="card">
<input type="submit" value="Lisa Seisund" name="btn" class="card">
</td>
</tr>
<%End If%>
</Form>
</Table>

<%If LEFT(request.QueryString("d"),1)="d" then%>
<%set mdboe = Server.CreateObject("ADODB.Command")%>
<%set mdbor = Server.CreateObject("ADODB.Recordset")%>
<%mdboe.ActiveConnection = mdbo%>

<%If Request.QueryString("Me")="" then%>
<%mes=Month(Date())%>
<%Else%>
<%mes=Request.QueryString("Me")%>
<%End if%>

<%If Request.QueryString("Yee")="" then%>
<%ya=Year(Date())%>
<%Else%>
<%ya=Request.QueryString("Yee")%>
<%End if%>

<%mdboe.CommandText="SELECT * from Mesjaca WHERE MONTHH=" & mes%>
<%mdbor.Open mdboe%>
<%mn=mdbor("MonthName")%>
<%mdbor.Close%>
<%mdboe.CommandText="SELECT * from Calendar WHERE MONTHH=" & mes & " AND YEARR=" & ya%>
<%mdbor.Open mdboe%>
<Table border=1 class="card">
<tr class="card"">
<th class="card" colspan="7">
<a href="ProjCard.asp?pc=<%=proc%>&d=<%=Request.QueryString("d")%>&me=<%=CDBL(mes)%>&yee=<%=CDBL(ya)-1%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>" class="th"><<</a>
<%=mn%>&nbsp<%=ya%>
<a href="ProjCard.asp?pc=<%=proc%>&d=<%=Request.QueryString("d")%>&me=<%=CDBL(mes)%>&yee=<%=CDBL(ya)+1%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>"  class="th">>></a>
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
<%j=1%>
<%For j=1 to 7%>
<%If mdbor.EOF=false then%>
<%If CDBL(mdbor("WeekD"))=j then%>

<%IF mdbor("Dayy")=DAY(DATE()) and CDBL(mes)=CDBL(MONTH(DATE())) and CDBL(ya)=CDBL(YEAR(DATE())) then%>
<th class="card">
<Font color="FF0000">
<%ELSE%>
<td class="card">
<%END IF%>
<%If RIGHT(request.QueryString("d"),1)="e" then%>
<a href="ProjCard.asp?pc=<%=proc%>&d1b=<%=Request.QueryString("d1b")%>&m1b=<%=Request.QueryString("m1b")%>&y1b=<%=Request.QueryString("y1b")%>&d=noc&d1e=<%=mdbor("Dayy")%>&m1e=<%=CDBL(mes)%>&y1e=<%=CDBL(ya)%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>">
<%IF mdbor("Dayy")=DAY(DATE()) and CDBL(mes)=CDBL(MONTH(DATE())) and CDBL(ya)=CDBL(YEAR(DATE())) then%>
<Font color="FF0000">
<%ELSE%>
<%END IF%>
<%=mdbor("Dayy")%>
</font>
</a>
<%End if%>
<%If RIGHT(request.QueryString("d"),1)="b" then%>
<a href="ProjCard.asp?pc=<%=proc%>&d1e=<%=Request.QueryString("d1e")%>&m1e=<%=Request.QueryString("m1e")%>&y1e=<%=Request.QueryString("y1e")%>&d=noc&d1b=<%=mdbor("Dayy")%>&m1b=<%=CDBL(mes)%>&y1b=<%=CDBL(ya)%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>">
<%IF mdbor("Dayy")=DAY(DATE()) and CDBL(mes)=CDBL(MONTH(DATE())) and CDBL(ya)=CDBL(YEAR(DATE())) then%>
<Font color="FF0000">
<%ELSE%>
<%END IF%>
<%=mdbor("Dayy")%>
</font>
</a>
<%End if%>
<%mdbor.Movenext%>
<%Else%>
<td class="card">
<%End if%>
<%Else%>
<td class="card">
<%End if%>
</td>
<%Next%>
</tr>
<%Next%>
<%mdbor.close%>
<%If Request.QueryString("Me")="" then%>
<%mes2=Month(Date())%>
<%Else%>
<%mes2=Request.QueryString("Me")%>
<%End if%>
<%If CDBL(mes)-1<=0 then%>
<%mes=13%>
<%Else%>
<%If CDBL(mes)+1>=13 then%>
<%mes2=0%>
<%End if%>
<%End if%>
<%If mes2-mes<0 then%>
<%mdboe.CommandText="SELECT * from Mesjaca WHERE MONTHH=" & mes-1 & " OR  MONTHH=" & mes2+1 & " ORDER BY MONTHH DESC"%>
<%Else%>
<%mdboe.CommandText="SELECT * from Mesjaca WHERE MONTHH=" & mes-1 & " OR  MONTHH=" & mes2+1 & " ORDER BY MONTHH ASC"%>
<%End if%>
<%mdbor.Open mdboe%>
<tr class="card">
<%If CDBL(mes)-1<>12 Then%>
<td class="card" colspan="3"><a href="ProjCard.asp?pc=<%=proc%>&d=<%=Request.QueryString("d")%>&me=<%=CDBL(mes)-1%>&yee=<%=CDBL(ya)%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>"><%=mdbor("MonthName")%></a>
<%Else%>
<td class="card" colspan="3"><a href="ProjCard.asp?pc=<%=proc%>&d=<%=Request.QueryString("d")%>&me=<%=CDBL(mes)-1%>&yee=<%=CDBL(ya)-1%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>"><%=mdbor("MonthName")%></a>
<%End If%>
</td>
<td class="card"><%mdbor.MoveNext%></td>
<%If CDBL(mes2)+1<>1 Then%>
<td class="card" colspan="3"><a href="ProjCard.asp?pc=<%=proc%>&d=<%=Request.QueryString("d")%>&me=<%=CDBL(mes2)+1%>&yee=<%=CDBL(ya)%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>"><%=mdbor("MonthName")%></a>
<%Else%>
<td class="card" colspan="3"><a href="ProjCard.asp?pc=<%=proc%>&d=<%=Request.QueryString("d")%>&me=<%=CDBL(mes2)+1%>&yee=<%=CDBL(ya)+1%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>"><%=mdbor("MonthName")%></a>
<%End If%>

</td>
</tr>
</table>

<%End if%>

<%Else%>
<%If request.Form("btn")="Lisa Kontrakt" or request.Form("btn2")="Lisa kirje" or request.QueryString("ne")<>"" or request.QueryString("n")="olz" or request.QueryString("n")="nez" or LEFT(request.QueryString("d"),2)="ns" or LEFT(request.QueryString("d"),2)="ds" or request.Form("btn")="Lisa Firma" Then%>
<%If request.Form("btn")="Lisa Firma" Then%>
<%set mdboaf = Server.CreateObject("ADODB.Command")%>
<%set mdboraf = Server.CreateObject("ADODB.Recordset")%>
<%mdboaf.ActiveConnection = mdbo%>
<%a=Request.Form("fs")%>
<%l=len(a)%>
<%sl2=""%>
<%For i=1 To l%>
<%c=Mid(a,i,1)%>
<%v=asc(c)%>
<%SELECT CASE v%>
<%Case 245%>
<%sl2=sl2 & "&otilde;"%>
<%Case 228%>
<%sl2=sl2 & "&auml;"%>
<%Case 246%>
<%sl2=sl2 & "&ouml;"%>
<%Case 252%>
<%sl2=sl2 & "&uuml;"%>
<%Case 213%>
<%sl2=sl2 & "&Otilde;"%>
<%Case 196%>
<%sl2=sl2 & "&Auml;"%>
<%Case 214%>
<%sl2=sl2 & "&Ouml;"%>
<%Case 220%>
<%sl2=sl2 & "&Uuml;"%>
<%Case Else%>
<%sl2=sl2 & c%>
<%END SELECT%>
<%Next%>
<%mdboaf.CommandText="INSERT INTO CompID (CompanyID,CompanyName) VALUES ('" & request.Form("fd") & "', '" & sl2 & "')"%>
<%mdboraf.Open mdboaf%>
<%End If%>
<Form method="POST" action="ProjCard.asp?pc=<%=proc%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&e=<%=pb%>&e3=<%=np%>">
<input type="submit" value="Vaata" name="btn" class="card">
<input type="submit" value="Redigeeri" name="btn" class="card">
<input type="submit" value="Kohaldama muutused" name="btn" class="card">
<br>
<table>
<tr class="card">
<th class="card">
Kontrakti Number
</th>
<th class="card">
<%If request.QueryString("ne")="nev" Then%>
<a href="ProjCard.asp?pc=<%=proc%>&d2b=<%=Request.QueryString("d2b")%>&m2b=<%=Request.QueryString("m2b")%>&y2b=<%=Request.QueryString("y2b")%>&d2e=<%=Request.QueryString("d2e")%>&m2e=<%=Request.QueryString("m2e")%>&y2e=<%=Request.QueryString("y2e")%>&ne=<%="oltt"%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&s2=<%=Request.QueryString("s2")%>&n1=<%=Request.QueryString("n1")%>" class="th">Firma Nimetus</a>
<%Else%>
<a href="ProjCard.asp?pc=<%=proc%>&d2b=<%=Request.QueryString("d2b")%>&m2b=<%=Request.QueryString("m2b")%>&y2b=<%=Request.QueryString("y2b")%>&d2e=<%=Request.QueryString("d2e")%>&m2e=<%=Request.QueryString("m2e")%>&y2e=<%=Request.QueryString("y2e")%>&ne=<%="nev"%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&s2=<%=Request.QueryString("s2")%>&n1=<%=Request.QueryString("n1")%>" class="th">Firma Nimetus</a>
<%End If%>

</th>
<th class="card">
<%If LEFT(request.QueryString("d"),2)<>"ds" Then%>
<a href="ProjCard.asp?pc=<%=proc%>&d2b=<%=Request.QueryString("d2b")%>&m2b=<%=Request.QueryString("m2b")%>&y2b=<%=Request.QueryString("y2b")%>&d2e=<%=Request.QueryString("d2e")%>&m2e=<%=Request.QueryString("m2e")%>&y2e=<%=Request.QueryString("y2e")%>&d=dsb&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&s2=<%=Request.QueryString("s2")%>&n1=<%=Request.QueryString("n1")%>" class="th">S&otilde;lmimise kuup&auml;ev</a>
<%Else%>
<a href="ProjCard.asp?pc=<%=proc%>&d2b=<%=Request.QueryString("d2b")%>&m2b=<%=Request.QueryString("m2b")%>&y2b=<%=Request.QueryString("y2b")%>&d2e=<%=Request.QueryString("d2e")%>&m2e=<%=Request.QueryString("m2e")%>&y2e=<%=Request.QueryString("y2e")%>&d=nsc&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&s2=<%=Request.QueryString("s2")%>&n1=<%=Request.QueryString("n1")%>" class="th">S&otilde;lmimise kuup&auml;ev</a>
<%End if%>
</th>
<th class="card">
<%If LEFT(request.QueryString("d"),2)<>"ds" Then%>
<a href="ProjCard.asp?pc=<%=proc%>&d2b=<%=Request.QueryString("d2b")%>&m2b=<%=Request.QueryString("m2b")%>&y2b=<%=Request.QueryString("y2b")%>&d2e=<%=Request.QueryString("d2e")%>&m2e=<%=Request.QueryString("m2e")%>&y2e=<%=Request.QueryString("y2e")%>&d=dse&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&s2=<%=Request.QueryString("s2")%>&n1=<%=Request.QueryString("n1")%>" class="th">L&otilde;ppimise kuup&auml;ev</a>
<%Else%>
<a href="ProjCard.asp?pc=<%=proc%>&d2b=<%=Request.QueryString("d2b")%>&m2b=<%=Request.QueryString("m2b")%>&y2b=<%=Request.QueryString("y2b")%>&d2e=<%=Request.QueryString("d2e")%>&m2e=<%=Request.QueryString("m2e")%>&y2e=<%=Request.QueryString("y2e")%>&d=nsc&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&s2=<%=Request.QueryString("s2")%>&n1=<%=Request.QueryString("n1")%>" class="th">L&otilde;ppimise kuup&auml;ev</a>
<%End if%>
</th>
<th class="card">
<%if request.QueryString("n")="nez" then%>
<a href="ProjCard.asp?pc=<%=proc%>&d2b=<%=Request.QueryString("d2b")%>&m2b=<%=Request.QueryString("m2b")%>&y2b=<%=Request.QueryString("y2b")%>&d2e=<%=Request.QueryString("d2e")%>&m2e=<%=Request.QueryString("m2e")%>&y2e=<%=Request.QueryString("y2e")%>&n=<%="olz"%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&s2=<%=Request.QueryString("s2")%>&n1=<%=Request.QueryString("n1")%>" class="th">T&ouml;&ouml;taja</a>
<%Else%>
<a href="ProjCard.asp?pc=<%=proc%>&d2b=<%=Request.QueryString("d2b")%>&m2b=<%=Request.QueryString("m2b")%>&y2b=<%=Request.QueryString("y2b")%>&d2e=<%=Request.QueryString("d2e")%>&m2e=<%=Request.QueryString("m2e")%>&y2e=<%=Request.QueryString("y2e")%>&n=<%="nez"%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&s2=<%=Request.QueryString("s2")%>&n1=<%=Request.QueryString("n1")%>" class="th">T&ouml;&ouml;taja</a>
<%End if%>
</th>
<th class="card">
Summa
</th>
</tr>
<%set mdbow = Server.CreateObject("ADODB.Command")%>
<%set mdborw = Server.CreateObject("ADODB.Recordset")%>
<%mdbow.ActiveConnection = mdbo%>
<%mdbow.CommandText="SELECT * from Worker ORDER BY EmplFname"%>
<%mdborw.Open mdbow%>
<%set mdbocy = Server.CreateObject("ADODB.Command")%>
<%set mdborcy = Server.CreateObject("ADODB.Recordset")%>
<%mdbocy.ActiveConnection = mdbo%>
<%mdbocy.CommandText="SELECT * from CompID ORDER BY CompanyName"%>
<%mdborcy.Open mdbocy%>

<tr class="card">
<td class="card">

<input type="Text" value="" name="<%="cntl"%>" class="card" size=11>

</td>
<td class="card">
<%If request.QueryString("ne")="nev" Then%>
<%onch="change('ProjCard.asp?pc=" & proc & "&d2b=" & Request.QueryString("d2b") & "&m2b=" & Request.QueryString("m2b") & "&y2b=" & Request.QueryString("y2b") & "&d2e=" & Request.QueryString("d2e") & "&m2e=" & Request.QueryString("m2e") & "&y2e=" & Request.QueryString("y2e") & "&ne=oltt&sr=" & srt & dd & "&no=" & n & "&y=" & zo & " &s=" & co & "&em=" & pb & "&e3=" & np & "&s2=' + cmpl.options[cmpl.selectedIndex].value + '&n1=' + empll.options[empll.selectedIndex].value"%>
<%Else%>
<%onch="change('ProjCard.asp?pc=" & proc & "&d2b=" & Request.QueryString("d2b") & "&m2b=" & Request.QueryString("m2b") & "&y2b=" & Request.QueryString("y2b") & "&d2e=" & Request.QueryString("d2e") & "&m2e=" & Request.QueryString("m2e") & "&y2e=" & Request.QueryString("y2e") & "&ne=nev&sr=" & srt & dd & "&no=" & n & "&y=" & zo & " &s=" & co & "&em=" & pb & "&e3=" & np & "&s2=' + cmpl.options[cmpl.selectedIndex].value + '&n1=' + empll.options[empll.selectedIndex].value"%>
<%End if%>

<%If LEFT(request.QueryString("d"),2)<>"ds" Then%>
<%onch=onch & ",'ProjCard.asp?pc=" & proc & "&d2b=" & Request.QueryString("d2b") & "&m2b=" & Request.QueryString("m2b") & "&y2b=" & Request.QueryString("y2b") & "&d2e=" & Request.QueryString("d2e") & "&m2e=" & Request.QueryString("m2e") & "&y2e=" & Request.QueryString("y2e") & "&d=dsb&sr=" & srt & dd & "&no=" & n & "&y=" & zo & " &s=" & co & "&em=" & pb & "&e3=" & np & "&s2=' + cmpl.options[cmpl.selectedIndex].value + '&n1=' + empll.options[empll.selectedIndex].value,'ProjCard.asp?pc=" & proc & "&d2b=" & Request.QueryString("d2b") & "&m2b=" & Request.QueryString("m2b") & "&y2b=" & Request.QueryString("y2b") & "&d2e=" & Request.QueryString("d2e") & "&m2e=" & Request.QueryString("m2e") & "&y2e=" & Request.QueryString("y2e") & "&d=dsb&sr=" & srt & dd & "&no=" & n & "&y=" & zo & " &s=" & co & "&em=" & pb & "&e3=" & np & "&s2=' + cmpl.options[cmpl.selectedIndex].value + '&n1=' + empll.options[empll.selectedIndex].value"%>
<%Else%>
<%onch=onch & ",'ProjCard.asp?pc=" & proc & "&d2b=" & Request.QueryString("d2b") & "&m2b=" & Request.QueryString("m2b") & "&y2b=" & Request.QueryString("y2b") & "&d2e=" & Request.QueryString("d2e") & "&m2e=" & Request.QueryString("m2e") & "&y2e=" & Request.QueryString("y2e") & "&d=nsc&sr=" & srt & dd & "&no=" & n & "&y=" & zo & " &s=" & co & "&em=" & pb & "&e3=" & np & "&s2=' + cmpl.options[cmpl.selectedIndex].value + '&n1=' + empll.options[empll.selectedIndex].value,'ProjCard.asp?pc=" & proc & "&d2b=" & Request.QueryString("d2b") & "&m2b=" & Request.QueryString("m2b") & "&y2b=" & Request.QueryString("y2b") & "&d2e=" & Request.QueryString("d2e") & "&m2e=" & Request.QueryString("m2e") & "&y2e=" & Request.QueryString("y2e") & "&d=nsc&sr=" & srt & dd & "&no=" & n & "&y=" & zo & " &s=" & co & "&em=" & pb & "&e3=" & np & "&s2=' + cmpl.options[cmpl.selectedIndex].value + '&n1=' + empll.options[empll.selectedIndex].value"%>
<%End if%>

<%if request.QueryString("n")="nez" then%>
<%onch=onch & ",'ProjCard.asp?pc=" & proc & "&d2b=" & Request.QueryString("d2b") & "&m2b=" & Request.QueryString("m2b") & "&y2b=" & Request.QueryString("y2b") & "&d2e=" & Request.QueryString("d2e") & "&m2e=" & Request.QueryString("m2e") & "&y2e=" & Request.QueryString("y2e") & "&n=olz&sr=" & srt & dd & "&no=" & n & "&y=" & zo & " &s=" & co & "&em=" & pb & "&e3=" & np & "&s2=' + cmpl.options[cmpl.selectedIndex].value + '&n1=' + empll.options[empll.selectedIndex].value)"%>
<%Else%>
<%onch=onch & ",'ProjCard.asp?pc=" & proc & "&d2b=" & Request.QueryString("d2b") & "&m2b=" & Request.QueryString("m2b") & "&y2b=" & Request.QueryString("y2b") & "&d2e=" & Request.QueryString("d2e") & "&m2e=" & Request.QueryString("m2e") & "&y2e=" & Request.QueryString("y2e") & "&n=nez&sr=" & srt & dd & "&no=" & n & "&y=" & zo & " &s=" & co & "&em=" & pb & "&e3=" & np & "&s2=' + cmpl.options[cmpl.selectedIndex].value + '&n1=' + empll.options[empll.selectedIndex].value)"%>
<%End if%>
<select Name="cmpl" size="1" class="card" style="width:150" onChange="<%=onch%>"> 
<%'mdborcy.MoveFirst%>
<%Do until mdborcy.EOF%>
<%if request.QueryString("s2")=mdborcy("CompanyId") then%>
<option value=<%=mdborcy("CompanyId")%> selected="true"><%=mdborcy("CompanyName")%></option>
<%Else%>
<option value="<%=mdborcy("CompanyId")%>" ><%=mdborcy("CompanyName")%></option>
<%End If%>
<%mdborcy.MoveNext%>
<%Loop%>
</select>
</td>
<td class="card">
<%If Request.QueryString("d2b")="" then%>
<%Dat1b=Date()%>
<%Else%>
<%Dat1b=Request.QueryString("d2b") & "." & Request.QueryString("m2b") & "." & Request.QueryString("y2b")%>
<%End if%>
<input type="Text" value="<%=Dat1b%>" name="<%="dcol"%>" class="card" size=12>
</td>
<td class="card">
<%If Request.QueryString("d2e")="" then%>
<%Dat1e=Date()%>
<%Else%>
<%Dat1e=Request.QueryString("d2e") & "." & Request.QueryString("m2e") & "." & Request.QueryString("y2e")%>
<%End if%>
<input type="Text" value="<%=Dat1e%>" name="<%="dcnl"%>" class="card" size=12>
</td>
<td class="card">
<select Name="empll" class="card" size="1" onChange="<%=onch%>">
<%'mdborw.MoveFirst%>
<%Do until mdborw.EOF%>
<%if request.QueryString("n1")="" then%>
<%ene=0%>
<%Else%>
<%ene=request.QueryString("n1")%>
<%End if%>
<%if CDBL(ene)=CDBL(mdborw("EmployeeId")) then%>
<option value=<%=mdborw("EmployeeId")%> selected="true"><%=mdborw("EmplFName")%>&nbsp<%=mdborw("EmplName")%></option>
<%Else%>
<option value="<%=mdborw("EmployeeId")%>"><%=mdborw("EmplFName")%>&nbsp<%=mdborw("EmplName")%></option>
<%End if%>
<%mdborw.MoveNext%>
<%Loop%>
</select>
</td>
<td class="card">
<input type="Text" value="0" name="<%="sucl"%>" class="card" size=8> 
</td>
</tr>

<%If LEFT(request.QueryString("d"),1)="d" then%>
<%set mdboe = Server.CreateObject("ADODB.Command")%>
<%set mdbor = Server.CreateObject("ADODB.Recordset")%>
<%mdboe.ActiveConnection = mdbo%>

<%If Request.QueryString("Me")="" then%>
<%mes=Month(Date())%>
<%Else%>
<%mes=Request.QueryString("Me")%>
<%End if%>

<%If Request.QueryString("Yee")="" then%>
<%ya=Year(Date())%>
<%Else%>
<%ya=Request.QueryString("Yee")%>
<%End if%>

<%mdboe.CommandText="SELECT * from Mesjaca WHERE MONTHH=" & mes%>
<%mdbor.Open mdboe%>
<%mn=mdbor("MonthName")%>
<%mdbor.Close%>
<%mdboe.CommandText="SELECT * from Calendar WHERE MONTHH=" & mes & " AND YEARR=" & ya%>
<%mdbor.Open mdboe%>
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
<%j=1%>
<%For j=1 to 7%>
<%If mdbor.EOF=false then%>
<%If CDBL(mdbor("WeekD"))=j then%>

<%IF mdbor("Dayy")=DAY(DATE()) and CDBL(mes)=CDBL(MONTH(DATE())) and CDBL(ya)=CDBL(YEAR(DATE())) then%>
<th class="card">
<%ELSE%>
<td class="card">
<%END IF%>
<%mdd=mdbor("Dayy")%>
<%If LEN(mdd)<2 then%>
<%mdd="0" & mdd%>
<%End if%>
<%mdm=mdbor("Monthh")%>
<%If LEN(mdm)<2 then%>
<%mdm="0" & mdm%>
<%End if%>
<%If RIGHT(request.QueryString("d"),1)="e" then%>
<a href="ProjCard.asp?pc=<%=proc%>&d2b=<%=Request.QueryString("d2b")%>&m2b=<%=Request.QueryString("m2b")%>&y2b=<%=Request.QueryString("y2b")%>&d=nsc&d2e=<%=mdd%>&m2e=<%=mdm%>&y2e=<%=CDBL(ya)%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>">
<%IF mdbor("Dayy")=DAY(DATE()) and CDBL(mes)=CDBL(MONTH(DATE())) and CDBL(ya)=CDBL(YEAR(DATE())) then%>
<Font color="FF0000">
<%ELSE%>
<%END IF%>
<%=mdbor("Dayy")%>
</font>
</a>
<%End if%>
<%If RIGHT(request.QueryString("d"),1)="b" then%>
<a href="ProjCard.asp?pc=<%=proc%>&d2e=<%=Request.QueryString("d2e")%>&m2e=<%=Request.QueryString("m2e")%>&y2e=<%=Request.QueryString("y2e")%>&d=nsc&d2b=<%=mdd%>&m2b=<%=mdm%>&y2b=<%=CDBL(ya)%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>">
<%IF mdbor("Dayy")=DAY(DATE()) and CDBL(mes)=CDBL(MONTH(DATE())) and CDBL(ya)=CDBL(YEAR(DATE())) then%>
<Font color="FF0000">
<%ELSE%>
<%END IF%>
<%=mdbor("Dayy")%>
</font>
</a>
<%End if%>
<%mdbor.Movenext%>
<%Else%>
<td class="card">

<%End if%>
<%Else%>
<td class="card">

<%End if%>

</td>
<%Next%>
</tr>
<%Next%>
<%mdbor.close%>
<%If Request.QueryString("Me")="" then%>
<%mes2=Month(Date())%>
<%Else%>
<%mes2=Request.QueryString("Me")%>
<%End if%>
<%If CDBL(mes)-1<=0 then%>
<%mes=13%>
<%Else%>
<%If CDBL(mes)+1>=13 then%>
<%mes2=0%>
<%End if%>
<%End if%>
<%If mes2-mes<0 then%>
<%mdboe.CommandText="SELECT * from Mesjaca WHERE MONTHH=" & mes-1 & " OR  MONTHH=" & mes2+1 & " ORDER BY MONTHH DESC"%>
<%Else%>
<%mdboe.CommandText="SELECT * from Mesjaca WHERE MONTHH=" & mes-1 & " OR  MONTHH=" & mes2+1 & " ORDER BY MONTHH ASC"%>
<%End if%>
<%mdbor.Open mdboe%>
<tr class="card">
<%If CDBL(mes)-1<>12 Then%>
<td class="card" colspan="3"><a href="ProjCard.asp?pc=<%=proc%>&d=<%=Request.QueryString("d")%>&me=<%=CDBL(mes)-1%>&yee=<%=CDBL(ya)%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>"><%=mdbor("MonthName")%></a>
<%Else%>
<td class="card" colspan="3"><a href="ProjCard.asp?pc=<%=proc%>&d=<%=Request.QueryString("d")%>&me=<%=CDBL(mes)-1%>&yee=<%=CDBL(ya)-1%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>"><%=mdbor("MonthName")%></a>
<%End If%>
</td>
<td class="card"><%mdbor.MoveNext%></td>
<%If CDBL(mes2)+1<>1 Then%>
<td class="card" colspan="3"><a href="ProjCard.asp?pc=<%=proc%>&d=<%=Request.QueryString("d")%>&me=<%=CDBL(mes2)+1%>&yee=<%=CDBL(ya)%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>"><%=mdbor("MonthName")%></a>
<%Else%>
<td class="card" colspan="3"><a href="ProjCard.asp?pc=<%=proc%>&d=<%=Request.QueryString("d")%>&me=<%=CDBL(mes2)+1%>&yee=<%=CDBL(ya)+1%>&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>"><%=mdbor("MonthName")%></a>
<%End If%>

</td>
</tr>
</table>

<%End if%>


<%If request.QueryString("ne")="nev" Then%>
<tr class="card>
<th class="card">
Firma kood
</th>
<th class="card">
Firma nimetus
</th>
<th class="card">
</th>
</tr>
<tr class="card">
<td class="card">
<input type="text" value="" name="fd" class="card">
</td>
<td class="card">
<input type="text" value="" name="fs" class="card">
</td>
<td class="card">
<input type="submit" value="Lisa Firma" name="btn"  class="card">
</td>
</tr>
<%End If%>
<%If request.QueryString("n")="nez" Then%>
<tr class="card">
<th class="card">
Isikukood
</th>
<th class="card">
Tootaja nimi
</th>
<th class="card">
Tootaja perekonnanimi
</th>
<th class="card">
Tootaja ametikoht
</th>
<th class="card">
</th>
<tr/>
<tr class="card">
<td class="card">
</td>
<td class="card">
<input type="text" value="" name="En" class="card">
</td>
<td class="card">
<input type="text" value="" name="Efn" class="card">
</td>
<td class="card">
<input type="text" value="" name="tn" class="card">
</td>
<td class="card">
<input type="submit" value="Lisa kirje" name="btn2" class="card">
</td>
</tr>
<%End If%>

</table>
</form>
<%End If%>
<%End If%>
<%End if%>
<%End if%>
<%End if%>
<%End if%>
</body>
</html>