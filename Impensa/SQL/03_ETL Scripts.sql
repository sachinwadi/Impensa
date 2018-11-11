UPDATE Tmp_Expenses
SET sNotes = REPLACE(sNotes, '"','')
GO

insert into tbl_Notes (sNotes)
Select distinct sNotes From Tmp_Expenses Order By sNotes
GO

Insert Into tbl_CategoryList(sCategory)
Select Distinct Category From Tmp_Expenses Order By Category
GO

Insert into tbl_ExpenditureDet (dtDate, iCategory, dAmount, sNotes)
SELECT T.dtDate, C.hKey, T.dAmount, T.sNotes FROM Tmp_Expenses T
INNER JOIN tbl_CategoryList C ON C.sCategory = T.Category
ORDER BY T.dtDate
GO