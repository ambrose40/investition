<Html>
 <Head>
<!--������ ��� ������������� �������� ���� ��������� ����� ����������� �������������-->
  <Script Type="text/javaScript">
   function confirmClose() 
    {
    if (confirm("Kas tahate panema see aken kinni?")) 
     {
     parent.close();
     }
    }
  </Script>
<!--���������� ���������� ���� ������������ ���� Inv-->
  <%b= Server.MapPath("\inv")%>
<!--������ ���� �������� ���� StyleInv ������, �� ��������� ����������� ����� ���������� � ��������������� ����� �� �������-->
  <%If Request.Cookies("StyleInv")="" Then%>
   <%Set servcfg=Server.CreateObject("Scripting.FileSystemObject")%>
   <%Set servFileStream=servcfg.OpenTextFile(b & "\style.cfg")%>
   <%s=servFileStream.ReadLine%>
   <%servFileStream.Close%>
   <Link Rel="stylesheet" href="<%=s%>" Type="text/css">
  <%Else%>
<!--������ ���� �������� ���� StyleInv �� ������, �� ��������� ����� ��������� � ���� ��������-->
   <%s=Request.Cookies("StyleInv")%>
   <Link Rel="stylesheet" href="<%=s%>" Type="text/css">
  <%End If%>
<!--������������� ���������-->
  <Meta http-equiv="Content-Type" Content="text/Html; Charset=windows-1251">
  <Title>Invest-IT!on: N&Auml;GEMUSE VALIMINE</Title>
 </Head>
 <Body Class="Main">
<!--��� ������ ����� �� ��������� ���������� ������ ������������� �������� � ����� ������� ������ ��������� �����-->
  <p align="center"><a href="Main.asp"  target="_top" Class="HeadLink" onClick="confirmClose()">N&Auml;GEMUSE VALIMINE</a></p><br><br><br>
<!--���� ���� ������ ������ MUUTA �� ������������ ����� �������� ���� StyleInv �� ������ �������� � ����-->
  <%If Request.Form("btn")="MUUTA" Then%>
   <%s=Request.Form("styl")%>
   <%response.Cookies("StyleInv")=s%> 
   <%Response.Cookies("StyleInv").path="/"%>
   <%Response.Cookies("StyleInv").expires="01/01/2010"%>
  <%End If%>
<!--��������� ����� ��� ������ ������������� ������ � �������� �������������� ���������-->
  <Form Action="stylec.asp" Method="POST">
   <Select  Name="styl"  Class="Main">
    <Option Value="STYLE.CSS">Kollane N&auml;gemus</Option>
    <Option Value="STYLE2.CSS">Sinine N&auml;gemus</Option>
    <Option Value="STYLE3.CSS">Roheline N&auml;gemus</Option>
    <Option Value="STYLE4.CSS">Halline N&auml;gemus</Option>
    <Input Type="Submit" Name="btn" Value="MUUTA" Class="Main">
   </select>
  </Form>
 </Body>
</Html>