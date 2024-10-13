ALTER TABLE tbl_ExpenditureDet ALTER COLUMN tmp INTEGER

BEGIN TRAN
GO
BEGIN
UPDATE X
SET dAmount = Y.dAmount,
	sNotes = Y.tr,
	tmp = Y.RowNum
FROM tbl_ExpenditureDet X
INNER JOIN (SELECT	ROW_NUMBER() OVER (PARTITION BY d.dtDate, d.iCategory ORDER BY d.dtDate, d.iCategory) AS RowNum,
		D.hKey, 
		D.dtDate, 
		C.sCategory,  
		D.sNotes, 
		D1.dAmount,
		D1.tr
  FROM	tbl_ExpenditureDet D
		INNER JOIN tbl_CategoryList C ON C.hKey = D.iCategory
		INNER JOIN (SELECT	d.dtDate, 
							d.iCategory, 
							SUM(d.dAmount) AS dAmount,  
							STRING_AGG(CASE 
											WHEN ISNULL(d.sNotes, '') = '' THEN NULL 
											ELSE ISNULL(d.sNotes, '') + ' (' + CONVERT(VARCHAR, d.dAmount) + ')' 
										END, ',') AS tr 
					  FROM tbl_ExpenditureDet d 
					 GROUP BY d.dtDate, d.iCategory 
					HAVING COUNT(*) > 1) D1 ON D.dtDate = D1.dtDate AND D.iCategory = D1.iCategory) Y ON X.hKey = Y.hKey
WHERE Y.RowNum = 1

UPDATE X
SET tmp = -1
FROM tbl_ExpenditureDet X
INNER JOIN (SELECT	ROW_NUMBER() OVER (PARTITION BY d.dtDate, d.iCategory ORDER BY d.dtDate, d.iCategory) AS RowNum,
		D.hKey, 
		D.dtDate, 
		C.sCategory,  
		D.sNotes, 
		D1.dAmount,
		D1.tr
  FROM	tbl_ExpenditureDet D
		INNER JOIN tbl_CategoryList C ON C.hKey = D.iCategory
		INNER JOIN (SELECT	d.dtDate, 
							d.iCategory, 
							SUM(d.dAmount) AS dAmount,  
							STRING_AGG(CASE 
											WHEN ISNULL(d.sNotes, '') = '' THEN NULL 
											ELSE ISNULL(d.sNotes, '') + ' (' + CONVERT(VARCHAR, d.dAmount) + ')' 
										END, ',') AS tr 
					  FROM tbl_ExpenditureDet d 
					 GROUP BY d.dtDate, d.iCategory 
					HAVING COUNT(*) > 1) D1 ON D.dtDate = D1.dtDate AND D.iCategory = D1.iCategory) Y ON X.hKey = Y.hKey
WHERE Y.RowNum > 1
END
GO

--select * from tbl_ExpenditureDet where tmp is not null order by dtDate, iCategory 

--ROLLBACK

--Commit

--select * from tbl_ExpenditureDet where hKey in (398755,398776) --, 315, 160


--select d.dtDate, C.sCategory, count(*) from tbl_ExpenditureDet d
--inner join tbl_CategoryList c on c.hKey = d.iCategory
--group by d.dtDate, c.sCategory
--having count(*) > 1
--order by dtDate

--BEGIN TRAN
--GO
--delete from tbl_ExpenditureDet where ISNULL(tmp, 0) = -1