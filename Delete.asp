<html>
 <%set servcfg=Server.CreateObject("Scripting.FileSystemObject")%> 
 <Head>
<!--Серверный скрипт выбора стиля оформления страницы на основе выбора пользователя, путём чтения объекта Cookie-->
  <%b= Server.MapPath("\")%>
  <%if request.Cookies("StyleInv")="" then%>
   <%set servFileStream=servcfg.OpenTextFile(b & "\style.cfg")%>
   <%s=servFileStream.ReadLine%>
   <%servFileStream.Close%>
   <link rel="stylesheet" href="<%=s%>" type="text/css">
  <%else%>
   <%s=request.Cookies("StyleInv")%>
   <link rel="stylesheet" href="<%=s%>" type="text/css">
  <%End if%>
<!--Выбираем русскую кодировку-->
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
  <title>Invest-IT!on: ANDMETE KUSTUTAMINE</title>
<!--Скрипты подтверждения удаления-->
  <SCRIPT LANGUAGE="VBScript"> 
   Sub del_OnClick
    Dim TheForm
    Set TheForm = Document.forms("ValidForm")
    MyVar=MsgBox("Kas tahate kustuta see kirje???",VbYesNo,"Kustutamine")
    If myVar=6 then     
     TheForm.Submit
    End if
   End Sub
   Sub del2_OnClick
    Dim TheForm
    Set TheForm = Document.forms("ValidForm2")
    MyVar=MsgBox("Kas tahate kustuta see kirje???",VbYesNo,"Kustutamine")
    If myVar=6 then     
     TheForm.Submit
    End if
   End Sub
   Sub del3_OnClick
    Dim TheForm
    Set TheForm = Document.forms("ValidForm3")
    MyVar=MsgBox("Kas tahate kustuta see kirje???",VbYesNo,"Kustutamine")
    If myVar=6 then     
     TheForm.Submit
    End if
   End Sub
   Sub del4_OnClick
    Dim TheForm
    Set TheForm = Document.forms("ValidForm4")
    MyVar=MsgBox("Kas tahate kustuta see kirje???",VbYesNo,"Kustutamine")
    If myVar=6 then     
     TheForm.Submit
    End if
   End Sub
  </SCRIPT>
 </Head>
 <body class="main">
<!--Создаём подключение к серверу баз данных-->
  <%set mdbo =  Server.CreateObject("ADODB.Connection")%>
  <%set servFileStream=servcfg.OpenTextFile(b & "\server.cfg")%>
  <%s=servFileStream.ReadLine%>
  <%i=servFileStream.ReadLine%>
  <%p=servFileStream.ReadLine%>
  <%servFileStream.Close%>
  <%mdbo.ConnectionString="Driver={SQL Server};Server=" & s & ";Trusted_Connection=yes;Database=invest;"%>
  <%mdbo.Open ConnectionString%>
<!--Рисуем заголовок с картинкой в углу-->
  <img border="0" src="icons/delete.ico" Style=float:Left><p align="center"><a href="Main.asp" class="Headlink">Andmete Kustutamine</a></p>
<!--Создаём серверные объекты для работы с базой данных-->
  <%set mdboe = Server.CreateObject("ADODB.Command")%>
  <%set mdbor = Server.CreateObject("ADODB.Recordset")%>
  <%mdboe.ActiveConnection = mdbo%>
<!--Запрашиваем всех внесённых в базу данных работников-->
  <%mdboe.CommandText="SELECT * from Worker"%>
  <%mdbor.Open mdboe%>
<!--Создаём внешнюю таблицу для размещения других внутренних таблиц-->
  <table>
   <tr>
    <td>
     <table bordercolor="5F5F5F" border="1" Style="border-collapse: collapse">
      <Form action="Delete.asp?did=1" method="POST" ID="ValidForm">
       <tr>
        <th colspan="2">
         Kustuta T&ouml;&ouml;taja kirje
        </th>
       </tr>
       <tr>
        <td>
         Vali t&ouml;&ouml;taja
        </td>
        <td>
         Kustuta?
        </td>
       </tr>
       <tr>
        <td>
         <Select name="emp" class="Main" size="10" style="font-size:smaller;"  onmouseover='window.status="Vali Aasta siin.";'onmouseout='window.status="";'  style="margin:0; padding:0">
          <%do until mdbor.EOF%>
           <option value="<%=mdbor("EmployeeID")%>"><%=mdbor("Emplname")%>&nbsp<%=mdbor("EmplFname")%></option>
           <%Mdbor.MoveNext%>
          <%Loop%>
         </select>
        </td>
        <td>
         <input name="del" class="Main" size="10" type="Submit" value="Kustuta!"  style="font-size:smaller;"  onmouseover='window.status="Vali Aasta siin.";'onmouseout='window.status="";'  style="margin:0; padding:0">
        </td>
       </tr>
      </form>
     </table>
    </td>
    <td>
<!--Закрываем ненужный более Массив с записями и задаём новый запрос. Запрашиваем все фирмы из справочника.-->
     <%mdbor.close%>
     <%mdboe.CommandText="SELECT * from CompID"%>
     <%mdbor.Open mdboe%>
<!--Вторая внутренняя таблица-->
     <table bordercolor="5F5F5F" border="1" Style="border-collapse: collapse">
      <Form action="Delete.asp?deed=1" method="POST" ID="ValidForm2">
       <tr>
        <th colspan="2">
         Kustuta firma kirje
        </th>
       </tr>
       <tr>
        <td>
         Vali firma
        </td>
        <td>
         Kustuta?
        </td>
       </tr>
       <tr>
        <td>
         <Select name="cmp" size="10" class="Main" style="font-size:smaller;"  onmouseover='window.status="Vali Aasta siin.";'onmouseout='window.status="";'  style="margin:0; padding:0">
          <%do until mdbor.EOF%>
           <option value="<%=mdbor("CompanyID")%>"><%=mdbor("Companyname")%></option>
           <%Mdbor.MoveNext%>
          <%Loop%>
         </select>
        </td>
        <td>
         <input name="del2" size="10" class="Main" type="Submit" value="Kustuta!"  style="font-size:smaller;"  onmouseover='window.status="Vali Aasta siin.";'onmouseout='window.status="";'  style="margin:0; padding:0">
        </td>
       </tr>
      </form>
     </table>
    </td>
    <td>
<!--Закрываем прошлый Массив. Запрашиваем все внесённые в базу данных предприятия-->
     <%mdbor.close%>
     <%mdboe.CommandText="SELECT * from Enterprise"%>
     <%mdbor.Open mdboe%>
     <table bordercolor="5F5F5F" border="1" Style="border-collapse: collapse">
      <Form action="Delete.asp?ded=1" method="POST" ID="ValidForm3">
       <tr>
        <th colspan="2">
         Kustuta ettev&otilde;tte kirje
        </th>
       </tr>
       <tr>
        <td>
         Vali ettev&otilde;tte
        </td>
        <td>
         Kustuta?
        </td>
       </tr>
       <tr>
        <td>
         <Select name="ent" size="10" class="Main" style="font-size:smaller;"  onmouseover='window.status="Vali Aasta siin.";'onmouseout='window.status="";'  style="margin:0; padding:0">
          <%do until mdbor.EOF%>
           <option value="<%=mdbor("Enterprise")%>"><%=mdbor("EDescr")%></option>
           <%Mdbor.MoveNext%>
          <%Loop%>
         </select>
        </td>
        <td>
         <input name="del3" size="10" class="Main" type="Submit" value="Kustuta!"  style="font-size:smaller;"  onmouseover='window.status="Vali Aasta siin.";'onmouseout='window.status="";'  style="margin:0; padding:0">
        </td>
       </tr>
      </form>
     </table>
    </td>
    <td>
<!--Закрываем прошлый Массив. Запрашиваем все внесённые в базу данных виды статусов-->
     <%mdbor.close%>
     <%mdboe.CommandText="SELECT * from StatCode"%>
     <%mdbor.Open mdboe%>
     <table bordercolor="5F5F5F" border="1" Style="border-collapse: collapse">
      <Form action="Delete.asp?diid=1" method="POST" ID="ValidForm4">
       <tr>
        <th colspan="2">
         Kustuta seisundi kirje
        </th>
       </tr>
       <tr>
        <td>
         Vali seisund
        </td>
        <td>
         Kustuta?
        </td>
       </tr>
       <tr>
        <td>
         <Select name="sei" size="10" class="Main" style="font-size:smaller;"  onmouseover='window.status="Vali Aasta siin.";'onmouseout='window.status="";'  style="margin:0; padding:0">
          <%do until mdbor.EOF%>
           <option value="<%=mdbor("StatusID")%>"><%=mdbor("StatusName")%></option>
           <%Mdbor.MoveNext%>
          <%Loop%>
         </select>
        </td>
        <td>
         <input name="del4" size="10" class="Main" type="Submit" value="Kustuta!"  style="font-size:smaller;"  onmouseover='window.status="Vali Aasta siin.";'onmouseout='window.status="";'  style="margin:0; padding:0">
        </td>
       </tr>
      </form>
     </table>
    </td>
   </tr>
  </table>
<!--Удаление при подтверждении удаления, указанного в строке URL запроса-->
  <%mdbor.Close%>
  <%IF Request.QueryString("did")=1 then%>
   <%mdboe.CommandText="DELETE FROM Worker WHERE EmployeeID='" & request.Form("emp") & "'"%>
   <%mdbor.Open mdboe%>
  <%End if%>
  <%IF Request.QueryString("diid")=1 then%>
   <%mdboe.CommandText="DELETE FROM StatCode WHERE StatusID='" & request.Form("sei") & "'"%>
   <%mdbor.Open mdboe%>
  <%End if%>
  <%IF Request.QueryString("ded")=1 then%>
   <%mdboe.CommandText="DELETE FROM Enterprise WHERE Enterprise='" & request.Form("ent") & "'"%>
   <%mdbor.Open mdboe%>
  <%End if%>
  <%IF Request.QueryString("deed")=1 then%>
   <%mdboe.CommandText="DELETE FROM CompID WHERE CompanyID='" & request.Form("cmp") & "'"%>
   <%mdbor.Open mdboe%>
  <%End if%>
 </body>
</html>