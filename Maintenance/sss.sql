EXEC Oracle.dbo.INVESTINVOICECOPY
SELECT * FROM Invest.dbo.glav_project WHERE LEFT(ROWID,1)='A'
SELECT * FROM Oracle.dbo.glav_project WHERE LEFT(ROWID,1)='A'
use invest
BEGIN TRANSACTION
DECLARE @DSK as VARchar(50),@II as int,@JJ as int
SET @DSK=''
SET @II=0
DECLARE All_proj CURSOR FOR SELECT DISTINCT a.DESCRIPTION FROM Invest.dbo.glav_project a 
FULL OUTER JOIN oracle.dbo.glav_project b 
ON a.ZAVOD + a.PROJECT + a.TSEHH + a.KONTO + a.SUBKONTO + a.MES + a.EDATE = b.ZAVOD + b.PROJECT + b.TSEHH + b.KONTO + b.SUBKONTO + b.MES + b.EDATE AND 
                      ISNULL(a.DESCRIPTION, 0) = ISNULL(b.DESCRIPTION, 0) AND ISNULL(a.DOCS_REGNR, 0) = ISNULL(b.DOCS_REGNR, 0) 
WHERE     (b.ROWID IS NULL)

OPEN All_proj
SET @JJ=@@CURSOR_ROWS
WHILE @II < @JJ
BEGIN
SET @II=@II+1
PRINT @II
FETCH NEXT FROM All_proj INTO @DSK
PRINT @DSK
DELETE GLAV_PROJECT WHERE DESCRIPTION=@DSK
END

CLOSE all_proj
DEALLOCATE all_proj

COMMIT TRANSACTION