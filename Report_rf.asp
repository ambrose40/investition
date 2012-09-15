<!--Страничка для слияния в разных фреймах, страниц заголовка отчета, списка проектов и собственно отчёта -->
<Html>
 <Head>
  <Meta http-equiv="Content-Type" Content="text/Html; Charset=windows-1251">
  <Title>Invest-IT!on: SISEARUANNE</Title>
 </Head>
 <Frameset Rows="12%, 88%">
   <Frame  Scrolling="yes" Src=<%="report_rh.asp?" & Request.QueryString%>></Frame>
  <Frameset Cols="25%, 75%">
   <Frame  Scrolling="auto" Src=<%="report_rl.asp?" & Request.QueryString%>></Frame>
   <Frame  Scrolling="yes" Src=<%="report_r.asp?" & Request.QueryString%>></Frame>
  </Frameset>
 </Frameset>
</Html>