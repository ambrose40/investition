<html>
<!--Клиентский скрипт для изменения содержания ссылок на странице-->
<script type="text/javascript">
window.onload = function() {
  document.onselectstart = function() {return false;} // ie
  document.onmousedown = function() {return false;} // mozilla
}

function onKeyPress () {
var keycode;
if (window.event) keycode = window.event.keyCode;
else if (e) keycode = e.which;
else return true;
if (keycode == 13) {
newWindow ('Help.mht','','','','scrollbars');
return false
}
return true 
}
document.onkeypress = onKeyPress;

function change(name,name2,name3,name4)
{
document.links[1].href=name;
document.links[2].href=name2;
document.links[3].href=name3;
document.links[4].href=name4;
}
<!--Клиентский скрипт запрашивающий подтверждения о закрытии страницы и закрывающий её при утвердительном ответе-->
function confirmClose() {
    if (confirm("Kas tahate panema see aken kinni?")) {
      parent.close();
    }

}
</script>

<Head>

<!--Серверный скрипт выбора стиля оформления страницы на основе выбора пользователя, путём чтения объекта Cookie-->
<%b= Server.MapPath("\")%>
<%if request.Cookies("StyleInv")="" then%>
<%set servcfg=Server.CreateObject("Scripting.FileSystemObject")%>
<%set servFileStream=servcfg.OpenTextFile(b & "\style.cfg")
  s=servFileStream.ReadLine
  servFileStream.Close%>
<link rel="stylesheet" href="<%=s%>" type="text/css">
<%else%>
<%s=request.Cookies("StyleInv")%>
<link rel="stylesheet" href="<%=s%>" type="text/css">
<%End if%>

<!--Выбор русскоязычной кодировки-->
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<title>
Invest-IT!on: PROJEKTI ANDMESISESTUS
</title>

<!--Скрипт подтверждения удаления -->
<SCRIPT LANGUAGE="VBScript" FOR="btnk2"> 
Sub btnk2_OnClick
 Dim CEForm
 Set CEForm = Document.forms("ValidForm3")
 MyVar2=MsgBox("Kas tahate kustuta see projekt???",VbYesNo)
 If myVar2=6 then     
 CEForm.action="Insert.asp?del=1"
 CEForm.Submit
 End if
End Sub
</SCRIPT>
</Head>
<body class="card">
<!--В переменную b закидываем физический путь к каталогу интранета inv-->
<%b=Server.MapPath("\")%>
<!--Имя сервера вынимается из конфигурационного файла server.cfg и подставляется к строке подключения-->
<%set mdbo =  Server.CreateObject("ADODB.Connection")%>
<%set servcfg=Server.CreateObject("Scripting.FileSystemObject")
  set servFileStream=servcfg.OpenTextFile(b & "\server.cfg")
  s=servFileStream.ReadLine
  i=servFileStream.ReadLine
  p=servFileStream.ReadLine
  servFileStream.Close%>
<!--Инициализация строки подключения-->
<%mdbo.ConnectionString="Driver={SQL Server};Server=" & s & ";Database=invest;Trusted_Connection=yes;"%>
<%mdbo.Open ConnectionString%>
<!--Создание объектов для работы с базой данных-->
<%set mdboo = Server.CreateObject("ADODB.Command")%>
<%set mdboro = Server.CreateObject("ADODB.Recordset")%>
<%set mdboe = Server.CreateObject("ADODB.Command")%>
<%set mdbore = Server.CreateObject("ADODB.Recordset")%>
<%mdboo.ActiveConnection = mdbo%>
<%mdboe.ActiveConnection = mdbo%>

<!--Если утвердительно ответили на вопрос об удалении, то удаляем выбранный проект-->
<%If Request.QueryString("del") = "1" Then%>
<%If MID(Request.Form("pid"),9,1)="." Then%>
 <%piid=Mid(Request.Form("pid"),12,5)%>
 <%pood=Mid(Request.Form("pid"),1,11)%>
<%Else%>
 <%piid=MID(Request.Form("pid"),9,5)%>
 <%pood=Mid(Request.Form("pid"),1,8)%>
<%End If%>
<%mdboo.Commandtext="DELETE FROM CODES WHERE Pid=" & piid & ""%>
<%mdboro.Open mdboo%>
<%End if%>

<!--Спрашиваем нажата ли кнопка подтверждающая вставку проекта в инвестиционный план-->
<%If Request.Form("btn") = "     Kinnita" Then%>
<!--Запись в файл последнего выбранного значения года для упрощения ввода большого числа проектов-->
<%set servcfg=Server.CreateObject("Scripting.FileSystemObject")%>
<%set servFileStream=servcfg.createTextFile(b & "/cookie.cfg")%>
<%yas=Request.Form("yra")%>
<%servFileStream.WriteLine yas%>
<%servFileStream.Close%>

<%ftor=1%>
<!--Выделяем проектный код от уникального идентификатора в зависимости от длины проектного кода-->
<%If MID(Request.Form("pid"),9,1)="." Then%>
 <%piid=Mid(Request.Form("pid"),12,5)%>
 <%pood=Mid(Request.Form("pid"),1,11)%>
<%Else%>
 <%piid=MID(Request.Form("pid"),9,5)%>
 <%pood=Mid(Request.Form("pid"),1,8)%>
<%End If%>

<!--Выбираем последний год для которого есть запись с указанным идентификатором Pid-->
<%mdboo.CommandText="SELECT MAX(Yearr) as my FROM Main WHERE Pid ='" & piid & "'"%>
<%mdboro.Open mdboo%>

<!--Проверяем Массив записей на пустоту и в зависимости от результата устанавливаем переменную ftor в определенное значение-->
<%If mdboro.EOF or (mdboro("my") & "e")="e" then%>
<%ftor=1%>
<%Else%>
<%ftor=5%>
<%iii=mdboro("my")%>
<%End if%>
<%mdboro.Close%>

<!--Если переменная ftor переключена в значение 5 значит такой проект уже был в прошлом и надо его сумму на конец перебросить на начало года выбранного проекта-->
<%If ftor="5" Then%>
<!--Создаём объект работы с базой данных-->

<!--Выбираем общую сумму по контракту из инвестиционного плана для выбранного года, проекта и предприятия-->
<%mdboo.CommandText="SELECT SummTot FROM Main WHERE Identifier='C' AND Yearr='" & iii & "' AND Pid ='" & piid & "' AND Enterprise='" & Request.Form("ena") & "'"%>
<%mdboro.Open mdboo%>
<!--Если предыдущая выборка не дала результата, то выбираем снова но абстрагируясь от выбранного предприятия-->
<%If mdboro.EOF="True" Then%>
<%mdboro.Close%>
<%mdboo.CommandText="SELECT SummTot FROM Main WHERE Identifier='C' AND Yearr='" & iii & "' AND Pid ='" & piid & "'"%>
<%mdboro.Open mdboo%>
<%End If%>
<!--Присваиваем переменной aa значение выборки и закрываем Массив записей-->
<%aa=mdboro("SummTot")%>
<%mdboro.Close%>

<!--Аналогичная последовательность операций только для плановых значений-->
<%mdboo.CommandText="SELECT SummTot FROM Main WHERE Identifier='P' AND Yearr='" & iii & "' AND Pid='" & piid & "' AND Enterprise='" & Request.Form("ena") & "'"%>
<%mdboro.Open mdboo%>
<%If mdboro.EOF="True" Then%>
<%mdboro.Close%>
<%mdboo.CommandText="SELECT SummTot FROM Main WHERE Identifier='P' AND Yearr='" & iii & "' AND Pid ='" & piid & "'"%>
<%mdboro.Open mdboo%>
<%End If%>
<%bb=mdboro("SummTot")%>
<%mdboro.Close%>

<!--Аналогичная последовательность операций только для Фактических значений-->
<%mdboo.CommandText="SELECT SummTot,OracleCode,RusName FROM Main WHERE Identifier='F' AND Yearr='" & iii & "' AND Pid='" & piid & "' AND Enterprise='" & Request.Form("ena") & "'"%>
<%mdboro.Open mdboo%>
<%If mdboro.EOF="True" Then%>
<%mdboro.Close%>
<%mdboo.CommandText="SELECT SummTot,OracleCode,RusName FROM Main WHERE Identifier='F' AND Yearr='" & iii & "' AND Pid ='" & piid & "'"%>
<%mdboro.Open mdboo%>
<%End If%>
<%cc=mdboro("SummTot")%>

<!--Дополнительно записываем значение кода Oracle и Русское название проекта если он уже существует-->
<%oco=mdboro("OracleCode")%>
<%dcp=mdboro("RusName")%>
<%mdboro.Close%>
<%Else%>

<!--Если переменная ftor переключена в другое значение, значит такой проект ещё не был введен в инвестиционный план-->
<!--Обнуляем переменные aa,bb,cc, и присваем умолчательные значения коду Oracle и русскоязычному названию-->
<%aa=0%>
<%bb=0%>
<%cc=0%>
<%oco="N/A"%>
<%dcp="---"%>
<%End If%>

<!--Далее определившись с подставляемыми значениями осуществляем добавление записей в инвестиционный план.Всего три записи: по плану, по факту и по контракту-->
<%mdboo.CommandText="INSERT INTO Main (ProjCode,Pid,Yearr,Enterprise,PastSum,IKvartal,IIkvartal,IIIKvartal,IVKvartal,Identifier,OracleCode,RusName) VALUES ('" & pood & "','" & piid & "', '" & Request.Form("yra") & "', '" & Request.Form("ena") & "'," & aa & ",0,0,0,0,'C','" & oco & "','" & dcp & "')"%>
<%mdboro.Open mdboo%>
<%mdboo.CommandText="INSERT INTO Main (ProjCode,Pid,Yearr,Enterprise,PastSum,IKvartal,IIkvartal,IIIKvartal,IVKvartal,Identifier,OracleCode,RusName) VALUES ('" & pood & "','" & piid & "', '" & Request.Form("yra") & "', '" & Request.Form("ena") & "'," & cc & ",0,0,0,0,'F','" & oco & "','" & dcp & "')"%>
<%mdboro.Open mdboo%>
<%mdboo.CommandText="INSERT INTO Main (ProjCode,Pid,Yearr,Enterprise,PastSum,IKvartal,IIkvartal,IIIKvartal,IVKvartal,Identifier,OracleCode,RusName) VALUES ('" & pood & "','" & piid & "', '" & Request.Form("yra") & "', '" & Request.Form("ena") & "'," & bb & ",0,0,0,0,'P','" & oco & "','" & dcp & "')"%>
<%mdboro.Open mdboo%>

<!--Делаем обновление всех записей год начала проекта-->
<%mdboo.CommandText="EXEC yearBeg"%>
<%mdboro.Open mdboo%>
<%End if%>

<!--Спрашиваем была ли нажата кнопка добавления нового проекта в справочник-->
<%If request.Form("btn")="Lisa Proekt" Then%>
<!--Обработка введенной строки на наличие эстонских букв-->
<%a=Request.Form("pn")%>
<%l=len(a)%>
<%sl=""%>
<%For i=1 To l%>
<%c=Mid(a,i,1)%>
<%v=asc(c)%>
<%SELECT CASE v%>
<%Case 245%>
<%sl=sl & "&otilde;"%>
<%Case 228%>
<%sl=sl & "&auml;"%>
<%Case 246%>
<%sl=sl & "&ouml;"%>
<%Case 252%>
<%sl=sl & "&uuml;"%>
<%Case 213%>
<%sl=sl & "&Otilde;"%>
<%Case 196%>
<%sl=sl & "&Auml;"%>
<%Case 214%>
<%sl=sl & "&Ouml;"%>
<%Case 220%>
<%sl=sl & "&Uuml;"%>
<%Case Else%>
<%sl=sl & c%>
<%END SELECT%>
<%Next%>
<!--Вставляем обработанные значения в виде новой записи в справочнике проектов-->
<%mdboo.CommandText="INSERT INTO Codes (ProjCode,ProjName) VALUES ('" & request.Form("pc") & "', '" & sl & "')"%>
<%mdboro.Open mdboo%>
<%End If%>

<!--Спрашиваем была ли нажата кнопка добавления предприятия-->
<%If request.Form("btn")="Lisa Ettevotte" Then%>
<!--Обработка введенной строки на наличие эстонских букв-->
<%a=Request.Form("en")%>
<%l=len(a)%>
<%sl=""%>
<%For i=1 To l%>
<%c=Mid(a,i,1)%>
<%v=asc(c)%>
<%SELECT CASE v%>
<%Case 245%>
<%sl=sl & "&otilde;"%>
<%Case 228%>
<%sl=sl & "&auml;"%>
<%Case 246%>
<%sl=sl & "&ouml;"%>
<%Case 252%>
<%sl=sl & "&uuml;"%>
<%Case 213%>
<%sl=sl & "&Otilde;"%>
<%Case 196%>
<%sl=sl & "&Auml;"%>
<%Case 214%>
<%sl=sl & "&Ouml;"%>
<%Case 220%>
<%sl=sl & "&Uuml;"%>
<%Case Else%>
<%sl=sl & c%>
<%END SELECT%>
<%Next%>
<!--Вставляем обработанные значения в виде новой записи в справочнике предприятий-->
<%mdboo.CommandText="INSERT INTO Enterprise (Enterprise,EDescr) VALUES ('" & request.Form("ec") & "', '" & sl & "')"%>
<%mdboro.Open mdboo%>
<%End If%>

<!--Заполняем Массивы всеми записями проектов и предприятий из справочников-->
<%mdboo.CommandText="SELECT * FROM Codes"%>
<%mdboro.Open mdboo%>
<%mdboe.CommandText="SELECT * FROM Enterprise"%>
<%mdbore.Open mdboe%>

<!--Прорисовываем заголовок и картинку-->
<img border="0" src="icons/ins.ico" Style=float:Left><p align="center"><a href="Main.asp"  target="_top" class="Headlink" onClick="confirmClose()">PROJEKTI ANDMESISESTUS</a></p>
<br>
<!--Объявляем форму и рисуем две кнопки добавить и удалить проект-->
<Form id="ValidForm3" method="POST" action="Insert.asp">
<input type="Submit" class="button" value="     Kinnita" name="btn"  style="height=35;background-image:url('icons/kinnita.ico'); background-repeat: no-repeat; background-position: LEft;">&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp
<input type="button" class="button" value="     Kustuta" name="btnk2" style="height=35;background-image:url('icons/destroy.ico'); background-repeat: no-repeat; background-position: LEft;">
<br>
<br>

<!--Рисуем таблицу для оптимального отображения-->
<table border=1 class="card">
 <tr class="Card">
  <th>
   <%If request.QueryString("n")="new" Then%>
    <a href="Insert.asp?n=<%="old"%>" class="th">Projekti kood ja nimetus</a>
   <%Else%>
    <a href="Insert.asp?n=<%="new"%>" class="th">Projekti kood ja nimetus</a>
   <%End If%>
  </th>
  <th>
   Aasta
  </th>
  <th>
   <%If request.QueryString("n")="ne2" Then%>
    <a href="Insert.asp?n=<%="ol2"%>" class="th">Ettev&otilde;tte</a>
   <%Else%>
    <a href="Insert.asp?n=<%="ne2"%>" class="th">Ettev&otilde;tte</a>
   <%End If%>
  </th>
 </tr>
 <tr class="Card">
  <td class="Card">
   <!--Выводим список проектов на экран в объект Select-->
   <select size="25" name="pid" class="card" style="width:700" style="font-family:Lucida Console">
    <%Do until mdboro.EOF%>
     <%l=LEN(mdboro("Pid"))%>
     <%l2=LEN(mdboro("ProjCode"))%>
     <%If l=1 then%>
      <%s="&nbsp;&nbsp;"%>
     <%End if%>
     <%If l=2 then%>
      <%s="&nbsp;"%>
     <%End if%>
     <%If l=3 then%>
      <%s=""%>
     <%End if%>
     <%If l2=8 then%>
      <%s2="&nbsp;&nbsp;&nbsp;"%>
     <%else%>
      <%s2=""%>
     <%End if%>
     <option value="<%=mdboro("ProjCode")%><%=mdboro("Pid")%>"><%=mdboro("Pid")%><%=s%>| <%=mdboro("ProjCode")%><%=s2%>| <%=mdboro("ProjName")%></option>
     <%mdboro.movenext%>
    <%Loop%>
   </select>
  </td>
  <td class="Card">
   <!--Выводим список финансовых лет, в зависимости от текущего календарного года-->
   <!--Выбранным оказывается тот год, который был выбран в прошлый раз при добавлении проекта в инвестиционный план.
       Он считывается из кнофигурационного файла cookie.cfg-->
   <%set servcfg=Server.CreateObject("Scripting.FileSystemObject")%>
   <%set servFileStream=servcfg.OpenTextFile(b & "\cookie.cfg")
   yas=servFileStream.ReadLine
   servFileStream.Close%><p>
   <select size="1" name="yra" class="card">
    <%yar=Year(Date())%>
    <%yr=yar-5%>
    <%Do Until yr>yar+5%>
     <%If CDBL(yas)=yr then%>
      <option selected=true value="<%=yr%>"><%=yr%></option>
     <%Else%>
      <option value="<%=yr%>"><%=yr%></option>
     <%End if%>
     <%yr=yr+1%>
    <%Loop%>
   </select>
  </td>
  <td class="Card">
   <%mdbore.movefirst%>
   <select size="6" name="ena" class="card">
    <%Do until mdbore.EOF%>
     <option value="<%=mdbore("Enterprise")%>" ><%=mdbore("Edescr")%></option>
     <%mdbore.movenext%>
    <%Loop%>
   </select>
  </td>
 </tr>
</Form>
<Form method="POST" action="Insert.asp">
 <!--Запрашиваем строку запроса, если пользователь захотел добавить новый проект в справочник,
 то вывести на экран интерфейс добавления проекта.-->
 <%If Request.QueryString("n")="new" Then%>
  <tr class="Card">
   <th>
    Projekti Nimetus
   </th>
   <th colspan=2>
    Projekti Kood
   </th>
  </tr>
  <tr class="Card">
   <td class="Card">
    <input type="text" class="card" value="" name="pn" size="60" >
   </td>
   <td class="Card">
    <input type="text" class="card" value="" name="pc" size="10">
   </td>
   <td class="Card">
    <input type="Submit" class="button" value="Lisa Proekt" name="btn">
   </td>
  </tr>
 <%End If%>
 <!--Запрашиваем строку запроса, если пользователь захотел добавить новое предприятие в справочник,
 то вывести на экран интерфейс добавления предприятия.-->
 <%If Request.QueryString("n")="ne2" Then%>
  <tr class="Card">
   <th>
    Ettev&otilde;tte Nimetus
   </th>
   <th colspan=2>
    Ettev&otilde;tte Kood
   </th>
  </tr>
  <tr class="Card">
   <td class="Card">
    <input type="text" class="card" value="" name="en" size="60" >
   </td>
   <td class="Card">
    <input type="text" class="card" value="" name="ec" size="10">
   </td>
   <td class="Card">
    <input type="Submit" class="button" value="Lisa Ettevotte" name="btn">
   </td>
  </tr>
 <%End If%>
</form>
</Table>
</body>
</html>