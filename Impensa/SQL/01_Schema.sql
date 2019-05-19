IF OBJECT_ID('tbl_CategoryList') IS NULL 
	CREATE TABLE tbl_CategoryList (hKey Numeric(18, 0) Primary Key Identity(1,1), sCategory VARCHAR(100), sNotes VARCHAR(500), IsObsolete BIT DEFAULT (0))
GO

IF OBJECT_ID('tbl_ExpenditureDet') IS NULL 
BEGIN
	CREATE TABLE tbl_ExpenditureDet
	(	
		hKey Numeric(18,0) Primary Key Identity(1,1),
		dtDate DATE,
		iCategory Numeric(18,0) Foreign Key References tbl_CategoryList(hKey),
		dAmount Money,
		sNotes VARCHAR (500),
		iImportedRecord BIT
	)
END
GO
IF OBJECT_ID('UQ_Dt_Cat_Amt') IS NULL 
BEGIN
ALTER TABLE tbl_ExpenditureDet
ADD CONSTRAINT UQ_Dt_Cat_Amt UNIQUE (dtDate, iCategory, dAmount)
END
GO

IF OBJECT_ID('Temp_ImportData') IS NULL 
BEGIN
	CREATE TABLE Temp_ImportData
	(
		hKey Numeric(18,0) Primary Key Identity(1,1),
		dtDate VARCHAR(20),
		sCategory VARCHAR(100),
		dAmount VARCHAR(20),
		sNotes VARCHAR(500),
		sImportComments VARCHAR(500),
		iDelete BIT,
		iRowNumber INTEGER
	)
END
GO

IF EXISTS(SELECT * FROM sys.indexes WHERE OBJECT_ID = OBJECT_ID('tbl_ExpenditureDet') AND NAME ='I_DateCat')
	DROP INDEX I_DateCat ON tbl_ExpenditureDet
GO

CREATE NONCLUSTERED INDEX I_DateCat ON tbl_ExpenditureDet (dtDate, iCategory)
GO

IF EXISTS(SELECT * FROM sys.indexes WHERE OBJECT_ID = OBJECT_ID('tbl_ExpenditureDet') AND NAME ='I_DateCatNote')
	DROP INDEX I_DateCatNote ON tbl_ExpenditureDet
GO

CREATE NONCLUSTERED INDEX I_DateCatNote ON tbl_ExpenditureDet (dtDate, iCategory, sNotes)
GO

IF EXISTS(SELECT * FROM sys.indexes WHERE OBJECT_ID = OBJECT_ID('tbl_ExpenditureDet') AND NAME ='I_DateNote')
	DROP INDEX I_DateNote ON tbl_ExpenditureDet
GO

CREATE NONCLUSTERED INDEX I_DateNote ON tbl_ExpenditureDet (dtDate, sNotes)
GO

IF OBJECT_ID('Tmp_Expenses') IS NULL 
	CREATE TABLE Tmp_Expenses (dtDate DATE, Category VARCHAR (100), dAmount MONEY, sNotes VARCHAR (500))
GO

IF OBJECT_ID('tbl_EOY') IS NULL 
	CREATE TABLE tbl_EOY (Year# INT, IsYrClosed Bit)				
GO

IF OBJECT_ID('tbl_Notes') IS NULL 
	CREATE TABLE tbl_Notes (sNotes VARCHAR(500), dtLastUsed DATE DEFAULT GETDATE())
GO

IF EXISTS(SELECT * FROM sys.indexes WHERE OBJECT_ID = OBJECT_ID('tbl_Notes') AND NAME ='I_sNotes')
	DROP INDEX I_sNotes ON tbl_Notes
GO

CREATE CLUSTERED INDEX I_sNotes ON tbl_Notes(sNotes)
GO

IF OBJECT_ID('tbl_ExpThresholds') IS NULL 
BEGIN
	CREATE TABLE tbl_ExpThresholds
	(	
		hKey NUMERIC(18,0) PRIMARY KEY IDENTITY(1,1),
		dtMonth DATE,
		iCategory NUMERIC(18,0) FOREIGN KEY REFERENCES tbl_CategoryList(hKey),
		TAmount MONEY
	)
END
GO

IF NOT EXISTS(SELECT 1 
			    FROM sys.objects O 
					 INNER JOIN sys.columns C ON C.object_id = O.object_id
			   WHERE O.name = 'tbl_CategoryList'
				 AND C.name = 'bNotify'
				 AND O.type = 'U')

ALTER TABLE tbl_CategoryList ADD bNotify BIT DEFAULT 0 WITH VALUES
GO 

IF NOT EXISTS(SELECT 1 
				FROM sys.objects O 
					 INNER JOIN sys.columns C ON C.object_id = O.object_id
			   WHERE O.name = 'tbl_CategoryList'
			     AND C.name = 'iMonthlyOccurrences'
			     AND O.type = 'U')

ALTER TABLE tbl_CategoryList ADD iMonthlyOccurrences SMALLINT
GO

IF OBJECT_ID('tbl_CategoryPrevMonthOccurrences') IS NULL

CREATE TABLE tbl_CategoryPrevMonthOccurrences
(
	hKey BIGINT IDENTITY(1,1),
	dtMonth DATE,
	iCategory NUMERIC(18, 0),
	iMonthlyOccurrences SMALLINT,
	CONSTRAINT PK_tbl_CategoryPrevMonthOccurrences_hKey PRIMARY KEY CLUSTERED (hKey ASC),
	CONSTRAINT FK_tbl_CategoryPrevMonthOccurrences_iCategory FOREIGN KEY (iCategory) REFERENCES tbl_CategoryList(hKey)	
)
GO