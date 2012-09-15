<!--Страничка для слияния в разных фреймах, страниц заголовка отчета, списка проектов и собственно отчёта -->
<Html>
 <Head>
  <Title>Invest-IT!on: 10 AASTA INVESTEERIMISKAVA AASTATE KAUPA</Title>
 </Head>
 <Frameset rows="12%, 88%">
   <Frame Scrolling="yes" Src=<%="report10h.asp?" & Request.QueryString%>>
  <Frameset cols="25%, 75%">
   <Frame Scrolling="auto" Src=<%="report10l.asp?" & Request.QueryString%>>
   <Frame Scrolling="yes" Src=<%="report10.asp?" & Request.QueryString%>>
  </Frameset>
 </Frameset>
</Html>