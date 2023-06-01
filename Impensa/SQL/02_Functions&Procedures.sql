IF OBJECT_ID('Fn_SearchStrSplit') IS NOT NULL 
DROP FUNCTION Fn_SearchStrSplit
GO

CREATE FUNCTION Fn_SearchStrSplit
(
    @String VARCHAR(500), 
    @Delimiter CHAR(1)
)        
RETURNS @temptable TABLE (items VARCHAR(500))       
AS       
BEGIN
    DECLARE @idx INT,
			@slice VARCHAR(500)       

    SELECT @idx = 1       
        IF LEN(ISNULL(@String, '')) = 0 RETURN

    WHILE @idx!= 0       
    BEGIN       
        SET @idx = CHARINDEX(@Delimiter,@String)       
        IF @idx!=0       
            SET @slice = left(@String,@idx - 1)       
        ELSE 
            SET @slice = @String       

        IF(LEN(@slice)>0)  
            INSERT INTO @temptable(Items) VALUES(RTRIM(LTRIM(@slice)))       

        SET @String = RIGHT(@String,LEN(@String) - @idx)       
        IF LEN(@String) = 0 BREAK
    END   
	RETURN       
END 
GO
/*############################################################################################################################################################*/

IF OBJECT_ID('Fn_ListObsoleteCategories') IS NOT NULL 
DROP FUNCTION Fn_ListObsoleteCategories
GO

CREATE FUNCTION Fn_ListObsoleteCategories(@P_FromDate DATE, @P_ToDate DATE)       
RETURNS TABLE 
AS       
	RETURN (SELECT DISTINCT e.iCategory 
	          FROM tbl_ExpenditureDet e 
			       INNER JOIN tbl_CategoryList c ON c.hKey = e.iCategory
		     WHERE NOT EXISTS (SELECT 1 
								 FROM tbl_ExpenditureDet e1 
								WHERE e.iCategory = e1.iCategory 
								  AND e1.dtDate BETWEEN @P_FromDate AND @P_ToDate)
			    -- e.dtDate NOT BETWEEN @P_FromDate AND @P_ToDate
		       AND c.IsObsolete = 1)
GO
/*############################################################################################################################################################*/

IF OBJECT_ID('Fn_GetAllMonthsList') IS NOT NULL 
DROP FUNCTION Fn_GetAllMonthsList
GO

CREATE FUNCTION Fn_GetAllMonthsList(@FromDate DATE, @ToDate DATE)
RETURNS @MonthsList Table(sItemsList VARCHAR(50), dtFirstDay DATE)
AS
BEGIN
	Declare @Months Integer,
	@Years Integer,
	@CounterMonths Integer =1,
	@CounterYears Integer = 0,
	@StartingYear Integer,
	@BeginMonth CHAR(3),
	@EndMonth CHAR(3)
	
	Set @Months = DATEDIFF(mm, @FromDate, @ToDate)
	Set @CounterMonths = MONTH(@FromDate)
	Set @Years = DATEDIFF(YYYY, @FromDate, @ToDate)
	Set @StartingYear = YEAR(@FromDate)
	
	WHILE @CounterMonths <= (@Months + MONTH(@FromDate))
		Begin		
			Insert @MonthsList (sItemsList, dtFirstDay)
			SELECT DateName(mm,DATEADD(mm,@CounterMonths,-1)) + '-' + CONVERT(VARCHAR, @StartingYear + CASE WHEN @CounterMonths%12 = 0 THEN  ((@CounterMonths/12) - 1) ELSE (@CounterMonths/12) END ),
				   CONVERT(DATE, CONVERT(VARCHAR, @StartingYear + CASE WHEN @CounterMonths%12 = 0 THEN  ((@CounterMonths/12) - 1) ELSE (@CounterMonths/12) END ) +'-'+
				   CONVERT(VARCHAR, DATEPART(MM,DateName(mm,DATEADD(mm,@CounterMonths,-1))+' 01 1900'))+'-1')
			
			SET @CounterMonths = @CounterMonths + 1
			Continue
		End	
	
	WHILE @CounterYears <= @Years
		Begin	
			IF @CounterYears = 0	
				SET @BeginMonth = LEFT(DATENAME(M, @FromDate),3)
			ELSE
				SET @BeginMonth = 'Jan'
			
			IF (YEAR(@FromDate) = YEAR(@ToDate) OR (@StartingYear + @CounterYears) = YEAR(@ToDate))
				SET @EndMonth = LEFT(DATENAME(M, @ToDate),3)
			ELSE
				SET @EndMonth = 'Dec'
			
			INSERT @MonthsList (sItemsList, dtFirstDay)
			SELECT 'Year ' + Convert(VARCHAR,(@StartingYear + @CounterYears))+ ' (' + @BeginMonth + '-'+ @EndMonth +')', CONVERT(DATE, Convert(VARCHAR,(@StartingYear + @CounterYears)) + '-12-31')
			
			SET @CounterYears = @CounterYears + 1
		Continue
	End
	
	INSERT @MonthsList (sItemsList, dtFirstDay)
	SELECT 'Entire Selected Period', '2100-01-01'
			
	RETURN
END
GO
/*############################################################################################################################################################*/

IF OBJECT_ID('Fn_dAmount') IS NOT NULL 
	DROP FUNCTION Fn_dAmount
GO

CREATE FUNCTION Fn_dAmount(@P_sNotes VARCHAR(500), @P_SearchStr VARCHAR(500), @P_dAmount MONEY)
RETURNS MONEY
AS
BEGIN
	DECLARE @SubString VARCHAR(10),
			@Amount MONEY

	IF (ISNULL(@P_SearchStr, '') = '' OR CHARINDEX('(', @P_sNotes) = 0)
		RETURN @P_dAmount
	ELSE 
	BEGIN
		SET @SubString = SUBSTRING(LTRIM(SUBSTRING(@P_sNotes, CHARINDEX(@P_SearchStr, @P_sNotes), ( LEN(@P_sNotes) - CHARINDEX(@P_SearchStr, @P_sNotes) ) + 1)), 
																		 CHARINDEX('(', LTRIM(SUBSTRING(@P_sNotes, CHARINDEX(@P_SearchStr, @P_sNotes), (LEN(@P_sNotes)- CHARINDEX(@P_SearchStr, @P_sNotes))+1))) + 1, 
																		(CHARINDEX(')', LTRIM(SUBSTRING(@P_sNotes, CHARINDEX(@P_SearchStr, @P_sNotes), ( LEN(@P_sNotes) - CHARINDEX(@P_SearchStr, @P_sNotes) ) + 1))) - 
																		 CHARINDEX('(', LTRIM(SUBSTRING(@P_sNotes, CHARINDEX(@P_SearchStr, @P_sNotes), ( LEN(@P_sNotes) - CHARINDEX(@P_SearchStr, @P_sNotes) ) + 1))) ) - 1)
		IF ISNUMERIC(@SubString) = 1															 
			SET @Amount = CONVERT(MONEY, @SubString)
		ELSE
			SET @Amount = @P_dAmount
	END	
	RETURN @Amount
END
GO
/*############################################################################################################################################################*/

IF OBJECT_ID('Fn_GetTDTotal') IS NOT NULL 
	DROP FUNCTION Fn_GetTDTotal
GO

CREATE FUNCTION Fn_GetTDTotal(@P_BKStartDate DATE)
RETURNS @TDValues TABLE (MTD MONEY, YTD MONEY, BKSTD MONEY)
AS
BEGIN
	INSERT INTO @TDValues(MTD, YTD, BKSTD)
	SELECT 
	(SELECT ISNULL(SUM(dAmount),0) FROM tbl_ExpenditureDet WHERE dtDate >= CONVERT(DATE, DATEADD(dd,-(DAY(GETDATE())-1),GETDATE()))),
	(SELECT ISNULL(SUM(dAmount),0) FROM tbl_ExpenditureDet WHERE dtDate >= CONVERT(DATE, DATEADD(yy, DATEDIFF(yy,0,GETDATE()), 0))),
	(SELECT ISNULL(SUM(dAmount),0) FROM tbl_ExpenditureDet WHERE dtDate >= @P_BKStartDate)
	RETURN 
END
GO
/*############################################################################################################################################################*/

IF Object_id('Fn_DataStore') IS NOT NULL
  DROP FUNCTION Fn_DataStore
GO

CREATE FUNCTION Fn_DataStore(@P_FromDate DATE, @P_ToDate DATE, @P_iCategory INTEGER, @P_SearchStr VARCHAR(500), @P_PeriodLimit BIT = 0)
RETURNS TABLE
AS
    RETURN
      (SELECT E.hKey, E.dtDate, E.iCategory, E.dAmount, E.sNotes, C.sCategory, C.IsObsolete, C.sNotes [CatNotes], Fn.items
         FROM tbl_ExpenditureDet E
              INNER JOIN tbl_CategoryList C ON C.hKey = E.iCategory
              LEFT JOIN dbo.Fn_searchstrsplit(@P_SearchStr, ',') Fn ON 1 = 1
       WHERE  1 = CASE WHEN @P_PeriodLimit = 1 AND E.dtDate BETWEEN CONVERT(DATE, CONVERT(VARCHAR, YEAR(E.dtDate)) + '-' + CONVERT(VARCHAR, DATEPART(M,@P_FromDate)) + '-'+ CONVERT(VARCHAR, DATEPART(D,@P_FromDate))) AND 
												                    CONVERT(DATE, CONVERT(VARCHAR, YEAR(E.dtDate)) + '-' + CONVERT(VARCHAR, DATEPART(M,@P_ToDate)) + '-'+ CONVERT(VARCHAR, DATEPART(D,@P_ToDate))) THEN 1
				       WHEN @P_PeriodLimit = 0 AND E.dtDate BETWEEN @P_FromDate AND @P_ToDate THEN 1
				       ELSE 0
			       END
              AND E.iCategory = CASE WHEN @P_iCategory = 0 THEN e.iCategory ELSE @P_iCategory END
              AND ISNULL(E.sNotes, '') LIKE '%' + CASE WHEN ISNULL (@P_SearchStr, '') = '' THEN ISNULL(E.sNotes, '') ELSE Fn.items END + '%')
GO
/*############################################################################################################################################################*/

IF OBJECT_ID('sp_GetExpenditureDetails') IS NOT NULL 
	DROP PROCEDURE sp_GetExpenditureDetails
GO
  
CREATE PROCEDURE sp_GetExpenditureDetails (@P_FromDate DATE, @P_ToDate DATE, @P_Category VARCHAR(MAX))  
AS  
BEGIN  
	SELECT	NULL [bDelete],
			CASE WHEN X.hKey = 999999999 THEN NULL ELSE X.hKey END hKey,
			CASE WHEN (X.hKey = 999999999 AND X.dtDate != '2100-01-01') THEN 'TOTAL' 
				 WHEN (X.hKey = 999999999 AND X.dtDate = '2100-01-01') THEN 'GRAND TOTAL'
				 ELSE CONVERT(VARCHAR,X.dtDate,103) 
			END [Date], 
			X.iCategory, 
			X.sCategory,
			X.dAmount Amount, 			
			X.sNotes Notes, 
			X.IsReadOnly, 
			X.IsDummy, 
			0 [IsDummyRowAdded], 
			X.sCategory [CategoryName], 
			CONVERT(VARCHAR,X.dtDate,103) DateOriginal  
	  FROM (SELECT 1 [Sort], E.hKey, E.dtDate, C.sCategory, E.dAmount, E.sNotes, 
				   E.iCategory, ISNULL(Y.IsYrClosed, 0) [IsReadOnly], 0 [IsDummy] 
			  FROM tbl_ExpenditureDet E 
			       INNER JOIN tbl_CategoryList C ON C.hKey = E.iCategory
				   LEFT OUTER JOIN tbl_EOY Y ON Y.Year# = YEAR(E.dtDate)  
			 WHERE dtDate BETWEEN ISNULL(@P_FromDate,'1900-01-01') AND ISNULL(@P_ToDate,'2100-01-01')  
			   AND 1= CASE WHEN LEN(ISNULL(@P_Category,'')) > 0 AND CHARINDEX(C.sCategory, ISNULL(@P_Category,'')) > 0 THEN 1  
						   WHEN LEN(ISNULL(@P_Category,'')) = 0 THEN 1  
						   ELSE 0 
					   END
	
			UNION ALL  
			
			SELECT 2, 999999999, ISNULL(E.dtDate,'2100-01-01') dtDate, NULL, SUM(E.dAmount), NULL, NULL, 1, 0 
			  FROM	tbl_ExpenditureDet E 
					INNER JOIN tbl_CategoryList C ON C.hKey = E.iCategory  
			 WHERE dtDate BETWEEN ISNULL(@P_FromDate,'1900-01-01') AND ISNULL(@P_ToDate,'2100-01-01')  
			   AND 1 = CASE WHEN LEN(ISNULL(@P_Category,'')) > 0 AND CHARINDEX(C.sCategory, ISNULL(@P_Category,'')) > 0 THEN 1  
							WHEN LEN(ISNULL(@P_Category,'')) = 0 THEN 1  
							ELSE 0 
						END
			 GROUP BY dtDate WITH ROLLUP
			
			UNION ALL
			
			SELECT TOP(5) 3, NULL,CONVERT(DATE,GETDATE(),105),NULL,NULL,NULL,NULL, 0, 1 FROM SYSOBJECTS  
			) X
	 ORDER BY ISNULL(dtDate,'2100-01-01') ,ISNULL(X.hKey,999999999),X.Sort
END
GO
/*############################################################################################################################################################*/

IF Object_id('sp_GetExpenditureSummary_Monthly') IS NOT NULL
  DROP PROCEDURE sp_GetExpenditureSummary_Monthly
GO

CREATE PROCEDURE sp_GetExpenditureSummary_Monthly(@P_FromDate  DATE, @P_ToDate DATE, @P_Category  VARCHAR(MAX), @P_SearchStr VARCHAR(500))
AS
  BEGIN
      DECLARE @col1 VARCHAR(MAX), 
			  @col2 VARCHAR(MAX), 
			  @col3 VARCHAR(MAX),
			  @col4 VARCHAR(MAX), 
			  @col5 VARCHAR(MAX), 
			  @col6 VARCHAR(MAX),
			  @col7 VARCHAR(MAX), 
			  @col8 VARCHAR(MAX), 
			  @col9 VARCHAR(MAX),
			  @sql VARCHAR(MAX)

	  SELECT DISTINCT DATEADD(dd, -( DAY(E.dtDate) - 1 ), E.dtDate) dtDate INTO #Temp
        FROM dbo.Fn_DataStore(@P_FromDate, @P_ToDate, 0, @P_SearchStr, 0) E
       WHERE 1 = CASE
                      WHEN LEN(ISNULL(@P_Category, '')) > 0 AND CHARINDEX(E.sCategory, ISNULL(@P_Category, '')) > 0 THEN 1
                      WHEN LEN(ISNULL(@P_Category, '')) = 0 THEN 1
                      ELSE 0
                  END
                  	
      SET @col1 = STUFF((SELECT ', ISNULL([' + DATENAME(MONTH, dtDate) + '-' + CONVERT(VARCHAR, YEAR(dtDate)) + '],0) [' + DATENAME(MONTH, dtDate) + '-' + CONVERT(VARCHAR, YEAR(dtDate)) + ']' FROM #Temp FOR XML PATH('')), 1, 1, '')
      SET @col1 = @col1 + ', [TOTAL]'
      
      SET @col2 = STUFF((SELECT ', [' + DATENAME(MONTH, dtDate) + '-' + CONVERT(VARCHAR, YEAR(dtDate)) + ']' FROM #Temp FOR XML PATH('')), 1, 1, '')
      SET @col2 = @col2 + ', [TOTAL]'
      
      SET @col3 = STUFF((SELECT ', SUM([' + DATENAME(MONTH, dtDate) + '-' + CONVERT(VARCHAR, YEAR(dtDate)) + '])' FROM #Temp FOR XML PATH('')), 1, 1, '')
      SET @col3 = @col3 + ', SUM([TOTAL])'

	  SET @col4 = STUFF((SELECT ', ISNULL([' + DATENAME(MONTH, dtDate) + '-' + CONVERT(VARCHAR, YEAR(dtDate)) + '_CNT],0) [' + DATENAME(MONTH, dtDate) + '-' + CONVERT(VARCHAR, YEAR(dtDate)) + '_CNT]' FROM #Temp FOR XML PATH('')), 1, 1, '')
      SET @col5 = STUFF((SELECT ', [' + DATENAME(MONTH, dtDate) + '-' + CONVERT(VARCHAR, YEAR(dtDate)) + '_CNT]' FROM #Temp FOR XML PATH('')), 1, 1, '')
      SET @col6 = STUFF((SELECT ', NULL' FROM #Temp FOR XML PATH('')), 1, 1, '')

	  SET @col7 = STUFF((SELECT ', ISNULL([' + DATENAME(MONTH, dtDate) + '-' + CONVERT(VARCHAR, YEAR(dtDate)) + '_BDG],0) [' + DATENAME(MONTH, dtDate) + '-' + CONVERT(VARCHAR, YEAR(dtDate)) + '_BDG]' FROM #Temp FOR XML PATH('')), 1, 1, '')
      SET @col8 = STUFF((SELECT ', [' + DATENAME(MONTH, dtDate) + '-' + CONVERT(VARCHAR, YEAR(dtDate)) + '_BDG]' FROM #Temp FOR XML PATH('')), 1, 1, '')
      SET @col9 = STUFF((SELECT ', NULL' FROM #Temp FOR XML PATH('')), 1, 1, '')

	  SET @sql = 'SELECT * INTO #Temp_Data
				    FROM (SELECT DATENAME(MONTH,dtDate) +''-''+  CONVERT(VARCHAR,YEAR(E.dtDate)) dtdate, 
								 E.sCategory, 
								 E.iCategory hKey, 
								 dbo.Fn_dAmount(E.sNotes, E.items, E.dAmount) dAmount 
							FROM dbo.Fn_DataStore(''<@P_FromDate>'', ''<@P_ToDate>'', 0, ''<@P_SearchStr>'', 0) E
						   WHERE 1 = CASE WHEN LEN(ISNULL(''<@P_Category>'','''')) > 0 
						     AND CHARINDEX(E.sCategory, ISNULL(''<@P_Category>'','''')) > 0 THEN 1 WHEN LEN(ISNULL(''<@P_Category>'','''')) = 0 THEN 1 ELSE 0 END) AS X
						 RIGHT JOIN (SELECT c1.sCategory Category, 
											c1.hKey iCategory 
									   FROM tbl_CategoryList c1 
											LEFT JOIN dbo.Fn_ListObsoleteCategories(''<@P_FromDate>'', ''<@P_ToDate>'') Fn ON c1.hKey = Fn.iCategory 
									  WHERE Fn.iCategory IS NULL
										AND 1 = CASE WHEN LEN(ISNULL(''<@P_Category>'','''')) > 0 
										AND CHARINDEX(C1.sCategory, ISNULL(''<@P_Category>'','''')) > 0 THEN 1 WHEN LEN(ISNULL(''<@P_Category>'','''')) = 0 THEN 1 ELSE 0 END)c ON c.iCategory = X.hKey; '
    
	SET @sql = @sql + 'SELECT * INTO #Temp_Data_Bdg FROM (SELECT DATENAME(MONTH,T.dtMonth) +''-''+  CONVERT(VARCHAR,YEAR(T.dtMonth)) + ''_BDG'' dtdate, 
																 C1.sCategory, 
																 C1.hKey, 
																 T.TAmount 
														    FROM tbl_ExpThresholds T 
																 INNER JOIN tbl_CategoryList C1 ON C1.hKey = T.iCategory
														   WHERE T.dtMonth BETWEEN DATEADD(dd,-(DAY(''<@P_FromDate>'') -1), ''<@P_FromDate>'') 
															 AND DATEADD(dd,-(DAY(''<@P_ToDate>'') -1), ''<@P_ToDate>'')
															 AND 1 = CASE WHEN LEN(ISNULL(''<@P_Category>'','''')) > 0 AND CHARINDEX(C1.sCategory, ISNULL(''<@P_Category>'','''')) > 0 THEN 1 WHEN LEN(ISNULL(''<@P_Category>'','''')) = 0 THEN 1 ELSE 0 END) AS X
														 RIGHT JOIN (SELECT c1.sCategory Category, c1.hKey iCategory 
																	   FROM tbl_CategoryList c1 
																			LEFT JOIN dbo.Fn_ListObsoleteCategories(''<@P_FromDate>'', ''<@P_ToDate>'') Fn ON c1.hKey = Fn.iCategory 
																	  WHERE Fn.iCategory IS NULL
																		AND 1 = CASE WHEN LEN(ISNULL(''<@P_Category>'','''')) > 0 AND CHARINDEX(C1.sCategory, ISNULL(''<@P_Category>'','''')) > 0 THEN 1 WHEN LEN(ISNULL(''<@P_Category>'','''')) = 0 THEN 1 ELSE 0 END)c ON c.iCategory = X.hKey; '

	SET @sql = @sql + 'SELECT X.Sort, X.iCategory, X.Category, <@col1>, <@col4>, <@col7>
					     FROM (SELECT 1 Sort, iCategory, Category, <@col1> 
							     FROM #Temp_Data X
							    PIVOT (SUM(x.dAmount) FOR X.dtDate IN (<@col2>)) AS pvt
					   
							   UNION ALL

							   SELECT 2 Sort, 999, ''TOTAL'', <@col3> 
							     FROM #Temp_Data X
							    PIVOT (SUM(x.dAmount) FOR X.dtDate IN (<@col2>)) AS pvt) X /*SUM*/
							  INNER JOIN (SELECT 1 Sort, iCategory, Category, <@col5> 
										    FROM (SELECT iCategory, Category, dtDate + ''_CNT'' dtDate, dAmount FROM #Temp_Data) X
											PIVOT (COUNT(x.dAmount) FOR X.dtDate IN (<@col5>)) AS pvt
					   
										  UNION ALL

										  SELECT 2 Sort, 999, ''TOTAL'', <@col6>) Y ON X.iCategory = Y.iCategory /*COUNT*/
							  INNER JOIN (SELECT 1 Sort, iCategory, Category, <@col7> 
										    FROM (SELECT * FROM #Temp_Data_Bdg) X
										   PIVOT (SUM(x.TAmount) FOR X.dtDate in (<@col8>)) AS pvt
					   
										  UNION ALL

										  SELECT 2 Sort, 999, ''TOTAL'', <@col9>) Z ON X.iCategory = Z.iCategory /*BUDGET*/
					  ORDER BY X.Sort, X.Category'


	  SET @sql = REPLACE(@sql, '<@col1>',        @col1)
	  SET @sql = REPLACE(@sql, '<@col2>',        @col2)
	  SET @sql = REPLACE(@sql, '<@col3>',        @col3)
	  SET @sql = REPLACE(@sql, '<@col4>',        @col4)
	  SET @sql = REPLACE(@sql, '<@col5>',        @col5)
	  SET @sql = REPLACE(@sql, '<@col6>',        @col6)
	  SET @sql = REPLACE(@sql, '<@col7>',        @col7)
	  SET @sql = REPLACE(@sql, '<@col8>',        @col8)
	  SET @sql = REPLACE(@sql, '<@col9>',        @col9)
	  SET @sql = REPLACE(@sql, '<@P_FromDate>',  @P_FromDate)
	  SET @sql = REPLACE(@sql, '<@P_ToDate>',    @P_ToDate)
	  SET @sql = REPLACE(@sql, '<@P_Category>',  @P_Category)
	  SET @sql = REPLACE(@sql, '<@P_SearchStr>', @P_SearchStr)
	  
      EXEC (@sql)
  END
GO
/*############################################################################################################################################################*/

IF OBJECT_ID('sp_GetExpenditureSummary_Yearly') IS NOT NULL
  DROP PROCEDURE sp_GetExpenditureSummary_Yearly
GO

CREATE PROCEDURE sp_GetExpenditureSummary_Yearly(@P_FromDate DATE, @P_ToDate DATE, @P_Category VARCHAR(MAX), @P_SearchStr VARCHAR(500), @P_Years VARCHAR(MAX) = NULL)
AS
  BEGIN
      DECLARE @col1 VARCHAR(MAX), 
			  @col2 VARCHAR(MAX), 
			  @col3 VARCHAR(MAX),
			  @col4 VARCHAR(MAX), 
			  @col5 VARCHAR(MAX), 
			  @col6 VARCHAR(MAX),
			  @col7 VARCHAR(MAX), 
			  @col8 VARCHAR(MAX), 
			  @col9 VARCHAR(MAX),
			  @sql VARCHAR(MAX)

	  SELECT DISTINCT CONVERT(VARCHAR, YEAR(E.dtDate)) dtDate INTO #Temp
        FROM dbo.Fn_DataStore(@P_FromDate, @P_ToDate, 0, @P_SearchStr, 0) E
       WHERE 1 = CASE WHEN ISNULL(@P_Years,'') = '' THEN 1 
                      WHEN (ISNULL(@P_Years,'') != '' AND CHARINDEX(CONVERT(VARCHAR, YEAR(E.dtDate)), @P_Years) > 0) THEN 1 
                      ELSE 0 
                  END
			 AND 1 = CASE WHEN LEN(ISNULL(@P_Category, '')) > 0 AND CHARINDEX(E.sCategory, ISNULL(@P_Category, '')) > 0 THEN 1
                          WHEN LEN(ISNULL(@P_Category, '')) = 0 THEN 1
                          ELSE 0 
                     END
                  	
      SET @col1 = STUFF((SELECT ', ISNULL([' + dtDate + '],0) [' + dtDate + ']' FROM #Temp FOR XML PATH('')), 1, 1, '')
      SET @col1 = @col1 + ', [TOTAL]'
      
      SET @col2 = STUFF((SELECT ', [' + dtDate + ']' FROM #Temp FOR XML PATH('')), 1, 1, '')
      SET @col2 = @col2 + ', [TOTAL]'
      
      SET @col3 = STUFF((SELECT ', SUM([' + dtDate + '])' FROM #Temp FOR XML PATH('')), 1, 1, '')
      SET @col3 = @col3 + ', SUM([TOTAL])'

	  SET @col4 = STUFF((SELECT ', ISNULL([' + dtDate + '_CNT],0) [' + dtDate + '_CNT]' FROM #Temp FOR XML PATH('')), 1, 1, '')
      SET @col5 = STUFF((SELECT ', [' + dtDate + '_CNT]' FROM #Temp FOR XML PATH('')), 1, 1, '') 
      SET @col6 = STUFF((SELECT ', NULL' FROM #Temp FOR XML PATH('')), 1, 1, '')
     

	  SET @col7 = STUFF((SELECT ', ISNULL([' + dtDate + '_BDG],0) [' + dtDate + '_BDG]' FROM #Temp FOR XML PATH('')), 1, 1, '')
      SET @col8 = STUFF((SELECT ', [' + dtDate + '_BDG]' FROM #Temp FOR XML PATH('')), 1, 1, '')
      SET @col9 = STUFF((SELECT ', NULL' FROM #Temp FOR XML PATH('')), 1, 1, '')
      

	  SET @sql = 'SELECT * INTO #Temp_Data
				    FROM (SELECT CONVERT(VARCHAR, YEAR(dtDate)) dtdate, 
								 E.sCategory, 
								 E.iCategory hKey, 
								 dbo.Fn_dAmount(E.sNotes, E.items, E.dAmount) dAmount 
							FROM dbo.Fn_DataStore(''<@P_FromDate>'', ''<@P_ToDate>'', 0, ''<@P_SearchStr>'', 0) E
						   WHERE 1 = CASE WHEN LEN(ISNULL(''<@P_Category>'','''')) > 0 
						     AND CHARINDEX(E.sCategory, ISNULL(''<@P_Category>'','''')) > 0 THEN 1 WHEN LEN(ISNULL(''<@P_Category>'','''')) = 0 THEN 1 ELSE 0 END) AS X
						 RIGHT JOIN (SELECT c1.sCategory Category, 
											c1.hKey iCategory 
									   FROM tbl_CategoryList c1 
											LEFT JOIN dbo.Fn_ListObsoleteCategories(''<@P_FromDate>'', ''<@P_ToDate>'') Fn ON c1.hKey = Fn.iCategory 
									  WHERE Fn.iCategory IS NULL
										AND 1 = CASE WHEN LEN(ISNULL(''<@P_Category>'','''')) > 0 
										AND CHARINDEX(C1.sCategory, ISNULL(''<@P_Category>'','''')) > 0 THEN 1 WHEN LEN(ISNULL(''<@P_Category>'','''')) = 0 THEN 1 ELSE 0 END)c ON c.iCategory = X.hKey; '
    
	SET @sql = @sql + 'SELECT * INTO #Temp_Data_Bdg FROM (SELECT CONVERT(VARCHAR, YEAR(T.dtMonth)) + ''_BDG'' dtdate, 
																 C1.sCategory, 
																 C1.hKey, 
																 T.TAmount 
														    FROM tbl_ExpThresholds T 
																 INNER JOIN tbl_CategoryList C1 ON C1.hKey = T.iCategory
														   WHERE T.dtMonth BETWEEN DATEADD(dd,-(DAY(''<@P_FromDate>'') -1), ''<@P_FromDate>'') 
															 AND DATEADD(dd,-(DAY(''<@P_ToDate>'') -1), ''<@P_ToDate>'')
															 AND 1 = CASE WHEN LEN(ISNULL(''<@P_Category>'','''')) > 0 AND CHARINDEX(C1.sCategory, ISNULL(''<@P_Category>'','''')) > 0 THEN 1 WHEN LEN(ISNULL(''<@P_Category>'','''')) = 0 THEN 1 ELSE 0 END) AS X
														 RIGHT JOIN (SELECT c1.sCategory Category, c1.hKey iCategory 
																	   FROM tbl_CategoryList c1 
																			LEFT JOIN dbo.Fn_ListObsoleteCategories(''<@P_FromDate>'', ''<@P_ToDate>'') Fn ON c1.hKey = Fn.iCategory 
																	  WHERE Fn.iCategory IS NULL
																		AND 1 = CASE WHEN LEN(ISNULL(''<@P_Category>'','''')) > 0 AND CHARINDEX(C1.sCategory, ISNULL(''<@P_Category>'','''')) > 0 THEN 1 WHEN LEN(ISNULL(''<@P_Category>'','''')) = 0 THEN 1 ELSE 0 END)c ON c.iCategory = X.hKey; '

	SET @sql = @sql + 'SELECT X.Sort, X.iCategory, X.Category, <@col1>, <@col4>, <@col7>
					     FROM (SELECT 1 Sort, iCategory, Category, <@col1> 
							     FROM #Temp_Data X
							    PIVOT (SUM(x.dAmount) FOR X.dtDate IN (<@col2>)) AS pvt
					   
							   UNION ALL

							   SELECT 2 Sort, 999, ''TOTAL'', <@col3> 
							     FROM #Temp_Data X
							    PIVOT (SUM(x.dAmount) FOR X.dtDate IN (<@col2>)) AS pvt) X /*SUM*/
							  INNER JOIN (SELECT 1 Sort, iCategory, Category, <@col5> 
										    FROM (SELECT iCategory, Category, dtDate + ''_CNT'' dtDate, dAmount FROM #Temp_Data) X
											PIVOT (COUNT(x.dAmount) FOR X.dtDate IN (<@col5>)) AS pvt
					   
										  UNION ALL

										  SELECT 2 Sort, 999, ''TOTAL'', <@col6>) Y ON X.iCategory = Y.iCategory /*COUNT*/
							  INNER JOIN (SELECT 1 Sort, iCategory, Category, <@col7> 
										    FROM (SELECT * FROM #Temp_Data_Bdg) X
										   PIVOT (SUM(x.TAmount) FOR X.dtDate in (<@col8>)) AS pvt
					   
										  UNION ALL

										  SELECT 2 Sort, 999, ''TOTAL'', <@col9>) Z ON X.iCategory = Z.iCategory /*BUDGET*/
					  ORDER BY X.Sort, X.Category'


	  SET @sql = REPLACE(@sql, '<@col1>',        @col1)
	  SET @sql = REPLACE(@sql, '<@col2>',        @col2)
	  SET @sql = REPLACE(@sql, '<@col3>',        @col3)
	  SET @sql = REPLACE(@sql, '<@col4>',        @col4)
	  SET @sql = REPLACE(@sql, '<@col5>',        @col5)
	  SET @sql = REPLACE(@sql, '<@col6>',        @col6)
	  SET @sql = REPLACE(@sql, '<@col7>',        @col7)
	  SET @sql = REPLACE(@sql, '<@col8>',        @col8)
	  SET @sql = REPLACE(@sql, '<@col9>',        @col9)
	  SET @sql = REPLACE(@sql, '<@P_FromDate>',  @P_FromDate)
	  SET @sql = REPLACE(@sql, '<@P_ToDate>',    @P_ToDate)
	  SET @sql = REPLACE(@sql, '<@P_Category>',  @P_Category)
	  SET @sql = REPLACE(@sql, '<@P_SearchStr>', @P_SearchStr)
	  
      EXEC (@sql)
  END
GO
/*############################################################################################################################################################*/

IF OBJECT_ID('sp_GetExpenditureSummary_RunningTotals') IS NOT NULL
  DROP PROCEDURE sp_GetExpenditureSummary_RunningTotals
GO

CREATE PROCEDURE sp_GetExpenditureSummary_RunningTotals(@P_FromDate DATE, @P_ToDate DATE, @P_iCategory INTEGER, @P_iMonth INTEGER, @P_SearchStr VARCHAR(500), @P_ShowOnlyRunningTotal BIT = 0, @P_SummaryType AS CHAR(3) = 'SUM')
AS
  BEGIN
      DECLARE @col1 VARCHAR(MAX), @sql VARCHAR(MAX), @col2 VARCHAR(MAX), @col3 VARCHAR(MAX), @col4 VARCHAR(MAX), @RowCount INTEGER

	  SELECT DISTINCT CONVERT(VARCHAR, YEAR(dtDate)) dtDate INTO #Temp FROM dbo.Fn_DataStore(@P_FromDate, @P_ToDate, @P_iCategory, @P_SearchStr, 0) E

      SET @col1 = STUFF((SELECT ',ISNULL([' + dtDate + '],0) [' + dtDate + ']' FROM #Temp FOR XML PATH('')), 1, 1, '')
      SET @col1 = @col1 + ',[TOTAL]'
      
      SET @col2 = STUFF((SELECT ',[' + dtDate + ']' FROM #Temp FOR XML PATH('')), 1, 1, '')
      SET @col2 = @col2 + ',[TOTAL]'
      
      SET @col3 = STUFF((SELECT ', SUM([' + dtDate + ']) [' + dtDate + ']' FROM #Temp FOR XML PATH('')), 1, 1, '')
      SET @col3 = @col3 + ',SUM([TOTAL]) [TOTAL]'
      
      SET @RowCount = (SELECT COUNT(DISTINCT CONVERT(VARCHAR, YEAR(dtDate))) cnt FROM dbo.Fn_DataStore(@P_FromDate, @P_ToDate, @P_iCategory, @P_SearchStr, 0))
      
      SET @col4 = LEFT(REPLICATE('NULL, ', @RowCount + 1), LEN(REPLICATE('NULL, ', @RowCount + 1)) - 1)
      
      SET @sql = 'SELECT 1 Sort, iMonth [iCategory], dtMonth [Category], <@col1>'
                 + ' FROM ( SELECT CONVERT(VARCHAR, YEAR(dtDate)) dtYear, Month(dtDate) #Month, dbo.Fn_damount(E.sNotes, E.items, E.dAmount) dAmount FROM dbo.Fn_DataStore(''<@P_FromDate>'', ''<@P_ToDate>'', <@P_iCategory>, ''<@P_SearchStr>'', 0) E ) AS X'
                 + ' RIGHT JOIN (SELECT Datename(M, CONVERT(DATE, ''1900-'' + CONVERT(VARCHAR, Y.iMonth) + ''-01'')) dtMonth, Y.iMonth FROM (SELECT TOP(12) ROW_NUMBER() OVER(ORDER BY NAME) iMonth FROM sys.objects)Y)Z ON X.#Month = Z.iMonth'
			     + ' PIVOT(' + CASE @P_SummaryType WHEN 'CNT' THEN 'COUNT' ELSE 'SUM' END + '(X.dAmount) FOR X.dtYear IN (<@col2>)) AS PVT WHERE <@P_ShowOnlyRunningTotal> = 0 

				UNION ALL
				
				SELECT X.* FROM (
				SELECT 2 Sort, NULL [iCategory], ''TOTAL'' [Category], <@col3>' 
                 + ' FROM (SELECT CONVERT(VARCHAR, YEAR(dtDate)) dtYear, Month(dtDate) #Month, dbo.Fn_damount(E.sNotes, E.items, E.dAmount) dAmount'
				 + ' FROM dbo.Fn_DataStore(''<@P_FromDate>'', ''<@P_ToDate>'',<@P_iCategory>, ''<@P_SearchStr>'', 0) E ) AS X'
                 + ' RIGHT JOIN (SELECT Datename(M, CONVERT(DATE, ''1900-'' + CONVERT(VARCHAR, Y.iMonth) + ''-01'')) dtMonth, Y.iMonth FROM (SELECT TOP(12) ROW_NUMBER() OVER(ORDER BY NAME) iMonth FROM sys.objects)Y)Z ON X.#Month = Z.iMonth'
				 + ' PIVOT(' + CASE @P_SummaryType WHEN 'CNT' THEN 'COUNT' ELSE 'SUM' END + '(X.dAmount) FOR X.dtYear IN (<@col2>)) AS PVT ) X WHERE <@P_ShowOnlyRunningTotal> = 0 
				
				UNION ALL
				
				SELECT 3 Sort, NULL, NULL, <@col4> FROM (SELECT 1 [X]) X WHERE <@P_ShowOnlyRunningTotal> = 0
				
				UNION ALL
				
				SELECT 4 Sort, <@P_iMonth>, Datename(M, CONVERT(DATE, ''1900-<@P_iMonth>-01'')), <@col3>'
                 + ' FROM (SELECT CONVERT(VARCHAR, YEAR(dtDate)) dtYear, Month(dtDate) #Month, CASE WHEN Month(dtDate) <= <@P_iMonth>'
                 + ' THEN ' + CASE @P_SummaryType WHEN 'CNT' THEN '1' ELSE 'dbo.Fn_damount(E.sNotes, E.items, E.dAmount)' END + ' ELSE 0 END dAmount'
				 + ' FROM dbo.Fn_DataStore(''<@P_FromDate>'', ''<@P_ToDate>'',<@P_iCategory>, ''<@P_SearchStr>'', 0) E ) AS X'
                 + ' RIGHT JOIN (SELECT Datename(M, CONVERT(DATE, ''1900-'' + CONVERT(VARCHAR, Y.iMonth) + ''-01'')) dtMonth, Y.iMonth FROM (SELECT TOP(12) ROW_NUMBER() OVER(ORDER BY NAME) iMonth FROM sys.objects)Y)Z ON X.#Month = Z.iMonth'
			     + ' PIVOT(SUM(X.dAmount) FOR X.dtYear IN (<@col2>)) AS PVT'
				 + ' ORDER BY Sort, iMonth'
	 
	 SET @sql = REPLACE(@sql, '<@col1>', @col1)
	 SET @sql = REPLACE(@sql, '<@col2>', @col2)
	 SET @sql = REPLACE(@sql, '<@col3>', @col3)
	 SET @sql = REPLACE(@sql, '<@col4>', @col4)
	 SET @sql = REPLACE(@sql, '<@P_FromDate>', @P_FromDate)
	 SET @sql = REPLACE(@sql, '<@P_ToDate>', @P_ToDate)
	 SET @sql = REPLACE(@sql, '<@P_iCategory>', @P_iCategory)
	 SET @sql = REPLACE(@sql, '<@P_iMonth>', @P_iMonth)
	 SET @sql = REPLACE(@sql, '<@P_SearchStr>', @P_SearchStr)
	 SET @sql = REPLACE(@sql, '<@P_ShowOnlyRunningTotal>', @P_ShowOnlyRunningTotal)
	 
     EXEC(@sql)
  END
GO
/*############################################################################################################################################################*/

IF Object_id('sp_GetExpenditureSummaryBudget_Monthly') IS NOT NULL
  DROP PROCEDURE sp_GetExpenditureSummaryBudget_Monthly
GO

CREATE PROCEDURE sp_GetExpenditureSummaryBudget_Monthly(@P_FromDate  DATE, @P_ToDate DATE, @P_Category  VARCHAR(MAX))
AS
  BEGIN
      DECLARE @sql VARCHAR(MAX), @col1 VARCHAR(MAX), @col2 VARCHAR(MAX), @col3 VARCHAR(MAX), @col4 VARCHAR(MAX), @col5 VARCHAR(MAX), @col6 VARCHAR(MAX)
	
	  SELECT DISTINCT DATEADD(dd, -( DAY(E.dtDate) - 1 ), E.dtDate) dtDate INTO #Temp
        FROM dbo.Fn_DataStore(@P_FromDate, @P_ToDate, 0, '', 0) E
       WHERE 1 = CASE
                      WHEN LEN(ISNULL(@P_Category, '')) > 0 AND CHARINDEX(E.sCategory, ISNULL(@P_Category, '')) > 0 THEN 1
                      WHEN LEN(ISNULL(@P_Category, '')) = 0 THEN 1
                      ELSE 0
                  END
                  	
      SET @col1 = STUFF((SELECT ',ISNULL([' + DATENAME(MONTH, dtDate) + '-' + CONVERT(VARCHAR, YEAR(dtDate)) + '],0) [' + DATENAME(MONTH, dtDate) + '-' + CONVERT(VARCHAR, YEAR(dtDate)) + ']' FROM #Temp FOR XML PATH('')), 1, 1, '')
      SET @col1 = @col1 + ',[TOTAL]'
      
      SET @col2 = STUFF((SELECT ',[' + DATENAME(MONTH, dtDate) + '-' + CONVERT(VARCHAR, YEAR(dtDate)) + ']' FROM #Temp FOR XML PATH('')), 1, 1, '')
      SET @col2 = @col2 + ',[TOTAL]'
      
      SET @col3 = STUFF((SELECT ',SUM([' + DATENAME(MONTH, dtDate) + '-' + CONVERT(VARCHAR, YEAR(dtDate)) + '])' FROM #Temp FOR XML PATH('')), 1, 1, '')
      SET @col3 = @col3 + ',SUM([TOTAL])'
      
      SELECT DISTINCT DATEADD(dd, -( DAY(T.dtMonth) - 1 ), T.dtMonth) dtDate INTO #Temp1
        FROM tbl_ExpThresholds T 
			 INNER JOIN tbl_CategoryList C1 ON C1.hKey = T.iCategory
       WHERE T.dtMonth BETWEEN DATEADD(dd,-(DAY(@P_FromDate)- 1), @P_FromDate) AND DATEADD(dd,-(DAY(@P_ToDate)- 1), @P_ToDate)
			 AND 1 = CASE
                      WHEN LEN(ISNULL(@P_Category, '')) > 0 AND CHARINDEX(C1.sCategory, ISNULL(@P_Category, '')) > 0 THEN 1
                      WHEN LEN(ISNULL(@P_Category, '')) = 0 THEN 1
                      ELSE 0
                  END
             AND EXISTS(SELECT 1 FROM tbl_ExpenditureDet E WHERE E.iCategory = T.iCategory AND DATEADD(dd,-(DAY(E.dtDate)- 1),E.dtDate) = T.dtMonth)
                  	
      SET @col4 = STUFF((SELECT ',ISNULL([' + DATENAME(MONTH, dtDate) + '-' + CONVERT(VARCHAR, YEAR(dtDate)) + '],0) [' + DATENAME(MONTH, dtDate) + '-' + CONVERT(VARCHAR, YEAR(dtDate)) + ']' FROM #Temp1 FOR XML PATH('')), 1, 1, '')
      SET @col4 = @col4 + ',[TOTAL]'
      SET @col5 = STUFF((SELECT ',[' + DATENAME(MONTH, dtDate) + '-' + CONVERT(VARCHAR, YEAR(dtDate)) + ']' FROM #Temp1 FOR XML PATH('')), 1, 1, '')
      SET @col5 = @col5 + ',[TOTAL]'
	  SET @col6 = STUFF((SELECT ',SUM([' + DATENAME(MONTH, dtDate) + '-' + CONVERT(VARCHAR, YEAR(dtDate)) + '])' FROM #Temp1 FOR XML PATH('')), 1, 1, '')
      SET @col6 = @col6 + ',SUM([TOTAL])'
      
      SET @sql = 'SELECT 1 Sort, iCategory, Category, <@col2> INTO #Temp2 FROM (SELECT DATENAME(MONTH,dtDate) +''-''+  CONVERT(VARCHAR,YEAR(E.dtDate)) dtdate, E.sCategory, E.iCategory hKey, dbo.Fn_dAmount(E.sNotes, E.items, E.dAmount) dAmount FROM dbo.Fn_DataStore(''<@P_FromDate>'', ''<@P_ToDate>'', 0, '''', 0) E'
                 + ' WHERE 1 = CASE WHEN LEN(ISNULL(''<@P_Category>'','''')) > 0 AND CHARINDEX(E.sCategory, ISNULL(''<@P_Category>'','''')) > 0 THEN 1 WHEN LEN(ISNULL(''<@P_Category>'','''')) = 0 THEN 1 ELSE 0 END) AS X'
                 + ' RIGHT JOIN (SELECT c1.sCategory Category, c1.hKey iCategory FROM tbl_CategoryList c1 LEFT JOIN dbo.Fn_ListObsoleteCategories(''<@P_FromDate>'', ''<@P_ToDate>'') Fn ON c1.hKey = Fn.iCategory WHERE Fn.iCategory IS NULL'
                 + ' AND 1 = CASE WHEN LEN(ISNULL(''<@P_Category>'','''')) > 0 AND CHARINDEX(C1.sCategory, ISNULL(''<@P_Category>'','''')) > 0 THEN 1 WHEN LEN(ISNULL(''<@P_Category>'','''')) = 0 THEN 1 ELSE 0 END)c ON c.iCategory = X.hKey'
                 + ' PIVOT (SUM(x.dAmount) FOR X.dtDate in (<@col2>)) AS pvt WHERE 1=2' +
                 + ' UNION ALL'
	             + ' SELECT 2 Sort, 999, ''TOTAL'', <@col3>'
                 + ' FROM (SELECT DATENAME(MONTH, E.dtDate) +''-''+  CONVERT(VARCHAR,YEAR(E.dtDate)) dtdate, E.sCategory, E.iCategory hKey, dbo.Fn_dAmount(E.sNotes, E.items, E.dAmount) dAmount FROM dbo.Fn_DataStore(''<@P_FromDate>'', ''<@P_ToDate>'', 0, '''', 0) E' 
                 + ' WHERE 1 = CASE WHEN LEN(ISNULL(''<@P_Category>'','''')) > 0 AND CHARINDEX(E.sCategory, ISNULL(''<@P_Category>'','''')) > 0 THEN 1 WHEN LEN(ISNULL(''<@P_Category>'','''')) = 0 THEN 1 ELSE 0 END) AS X'
				 + ' PIVOT (SUM(x.dAmount) FOR X.dtDate in (<@col2>)) AS pvt WHERE 1=2' +
	             + ' ORDER BY sort,category;'
	  
	             + ' TRUNCATE TABLE #Temp2;'
	  
	             + ' INSERT INTO #Temp2 (Sort, iCategory, Category, <@col5>) SELECT 1 Sort, iCategory, Category, <@col5> FROM (SELECT DATENAME(MONTH,T.dtMonth) +''-''+  CONVERT(VARCHAR,YEAR(T.dtMonth)) dtdate, C1.sCategory, C1.hKey, T.TAmount FROM tbl_ExpThresholds T INNER JOIN tbl_CategoryList C1 ON C1.hKey = T.iCategory'
                 + ' WHERE T.dtMonth BETWEEN DATEADD(dd,-(DAY(''<@P_FromDate>'') -1), ''<@P_FromDate>'') AND DATEADD(dd,-(DAY(''<@P_ToDate>'') -1), ''<@P_ToDate>'') '
                 + ' AND 1 = CASE WHEN LEN(ISNULL(''<@P_Category>'','''')) > 0 AND CHARINDEX(C1.sCategory, ISNULL(''<@P_Category>'','''')) > 0 THEN 1 WHEN LEN(ISNULL(''<@P_Category>'','''')) = 0 THEN 1 ELSE 0 END) AS X'
                 + ' RIGHT JOIN (SELECT c1.sCategory Category, c1.hKey iCategory FROM tbl_CategoryList c1 LEFT JOIN dbo.Fn_ListObsoleteCategories(''<@P_FromDate>'', ''<@P_ToDate>'') Fn ON c1.hKey = Fn.iCategory WHERE Fn.iCategory IS NULL'
                 + ' AND 1 = CASE WHEN LEN(ISNULL(''<@P_Category>'','''')) > 0 AND CHARINDEX(C1.sCategory, ISNULL(''<@P_Category>'','''')) > 0 THEN 1 WHEN LEN(ISNULL(''<@P_Category>'','''')) = 0 THEN 1 ELSE 0 END)c ON c.iCategory = X.hKey'
                 + ' PIVOT (SUM(x.TAmount) FOR X.dtDate in (<@col5>)) AS pvt'
	             + ' UNION ALL'
	             + ' SELECT 2 Sort, 999, ''TOTAL'', <@col6>'
                 + ' FROM (SELECT DATENAME(MONTH,T.dtMonth) +''-''+  CONVERT(VARCHAR,YEAR(T.dtMonth)) dtdate, C1.sCategory, C1.hKey, T.TAmount FROM tbl_ExpThresholds T INNER JOIN tbl_CategoryList C1 ON C1.hKey = T.iCategory'
                 + ' WHERE T.dtMonth BETWEEN DATEADD(dd,-(DAY(''<@P_FromDate>'') -1), ''<@P_FromDate>'') AND DATEADD(dd,-(DAY(''<@P_ToDate>'') -1), ''<@P_ToDate>'') '
                 + ' AND 1 = CASE WHEN LEN(ISNULL(''<@P_Category>'','''')) > 0 AND CHARINDEX(C1.sCategory, ISNULL(''<@P_Category>'','''')) > 0 THEN 1 WHEN LEN(ISNULL(''<@P_Category>'','''')) = 0 THEN 1 ELSE 0 END) AS X'
                 + ' RIGHT JOIN (SELECT c1.sCategory Category, c1.hKey iCategory FROM tbl_CategoryList c1 LEFT JOIN dbo.Fn_ListObsoleteCategories(''<@P_FromDate>'', ''<@P_ToDate>'') Fn ON c1.hKey = Fn.iCategory WHERE Fn.iCategory IS NULL'
                 + ' AND 1 = CASE WHEN LEN(ISNULL(''<@P_Category>'','''')) > 0 AND CHARINDEX(C1.sCategory, ISNULL(''<@P_Category>'','''')) > 0 THEN 1 WHEN LEN(ISNULL(''<@P_Category>'','''')) = 0 THEN 1 ELSE 0 END)c ON c.iCategory = X.hKey'
                 + ' PIVOT (SUM(x.TAmount) FOR X.dtDate in (<@col5>)) AS pvt' +
	             + ' ORDER BY sort,category;'
	  
	  SET @sql = REPLACE(@sql, '<@col1>', @col1)
	  SET @sql = REPLACE(@sql, '<@col2>', @col2)
	  SET @sql = REPLACE(@sql, '<@col3>', @col3)
	  SET @sql = REPLACE(@sql, '<@col4>', @col4)
	  SET @sql = REPLACE(@sql, '<@col5>', @col5)
	  SET @sql = REPLACE(@sql, '<@col6>', @col6)
	  SET @sql = REPLACE(@sql, '<@P_FromDate>', @P_FromDate)
	  SET @sql = REPLACE(@sql, '<@P_ToDate>', @P_ToDate)
	  SET @sql = REPLACE(@sql, '<@P_Category>', @P_Category)

      EXEC (@sql +' SELECT * FROM #Temp2;')
  END
GO
/*############################################################################################################################################################*/

IF OBJECT_ID('sp_GetExpenditureSummaryBudget_Yearly') IS NOT NULL
  DROP PROCEDURE sp_GetExpenditureSummaryBudget_Yearly
GO

CREATE PROCEDURE sp_GetExpenditureSummaryBudget_Yearly(@P_FromDate DATE, @P_ToDate DATE, @P_Category VARCHAR(MAX), @P_Years VARCHAR(MAX) = NULL)
AS
  BEGIN
      DECLARE @sql VARCHAR(MAX), @col1 VARCHAR(MAX), @col2 VARCHAR(MAX), @col3 VARCHAR(MAX), @col4 VARCHAR(MAX), @col5 VARCHAR(MAX), @col6 VARCHAR(MAX)
      
      SELECT DISTINCT CONVERT(VARCHAR, YEAR(E.dtDate)) dtDate INTO #Temp
        FROM dbo.Fn_DataStore(@P_FromDate, @P_ToDate, 0, '', 0) E
       WHERE 1 = CASE WHEN ISNULL(@P_Years,'') = '' THEN 1 
                      WHEN (ISNULL(@P_Years,'') != '' AND CHARINDEX(CONVERT(VARCHAR, YEAR(E.dtDate)), @P_Years) > 0) THEN 1 
                      ELSE 0 
                  END
             AND 1 = CASE WHEN LEN(ISNULL(@P_Category, '')) > 0 AND CHARINDEX(E.sCategory, ISNULL(@P_Category, '')) > 0 THEN 1
                          WHEN LEN(ISNULL(@P_Category, '')) = 0 THEN 1
                          ELSE 0 
                      END

      SET @col1 = STUFF((SELECT ',ISNULL([' + dtDate + '],0) [' + dtDate + ']' FROM #Temp FOR XML PATH('')), 1, 1, '')
      SET @col1 = @col1 + ',[TOTAL]'
      
      SET @col2 = STUFF((SELECT ',[' + dtDate + ']' FROM #Temp FOR XML PATH('')), 1, 1, '')
      SET @col2 = @col2 + ',[TOTAL]'
      
      SET @col3 = STUFF((SELECT ',SUM([' + dtDate + '])' FROM #Temp FOR XML PATH('')), 1, 1, '')
      SET @col3 = @col3 + ',SUM([TOTAL])'
      
      
      SELECT DISTINCT CONVERT(VARCHAR, YEAR(T.dtMonth)) dtDate INTO #Temp1
        FROM tbl_ExpThresholds T 
			 INNER JOIN tbl_CategoryList C1 ON C1.hKey = T.iCategory
       WHERE T.dtMonth BETWEEN DATEADD(dd,-(DAY(@P_FromDate)- 1), @P_FromDate) AND DATEADD(dd,-(DAY(@P_ToDate)- 1), @P_ToDate)
             AND 1 = CASE WHEN ISNULL(@P_Years,'') = '' THEN 1 
                          WHEN (ISNULL(@P_Years,'') != '' AND CHARINDEX(CONVERT(VARCHAR, YEAR(T.dtMonth)), @P_Years) > 0) THEN 1 
                          ELSE 0 
                      END
			 AND 1 = CASE WHEN LEN(ISNULL(@P_Category, '')) > 0 AND CHARINDEX(C1.sCategory, ISNULL(@P_Category, '')) > 0 THEN 1
                          WHEN LEN(ISNULL(@P_Category, '')) = 0 THEN 1
                          ELSE 0 
                      END
             AND EXISTS(SELECT 1 FROM tbl_ExpenditureDet E WHERE E.iCategory = T.iCategory AND DATEADD(dd,-(DAY(E.dtDate)- 1),E.dtDate) = T.dtMonth)

      SET @col4 = STUFF((SELECT ',ISNULL([' + dtDate + '],0) [' + dtDate + ']' FROM #Temp1 FOR XML PATH('')), 1, 1, '')
      SET @col4 = @col4 + ',[TOTAL]'
      SET @col5 = STUFF((SELECT ',[' + dtDate + ']' FROM #Temp1 FOR XML PATH('')), 1, 1, '')
      SET @col5 = @col5 + ',[TOTAL]'
      SET @col6 = STUFF((SELECT ',SUM([' + dtDate + '])' FROM #Temp1 FOR XML PATH('')), 1, 1, '')
      SET @col6 = @col6 + ',SUM([TOTAL])'
      
      SET @sql = 'SELECT 1 Sort, iCategory, Category, <@col2> INTO #Temp2 '
                 + ' FROM (SELECT CONVERT(VARCHAR, YEAR(dtDate)) dtdate, E.sCategory, E.iCategory hKey, dbo.Fn_dAmount(E.sNotes, E.items, E.dAmount) dAmount FROM dbo.Fn_DataStore(''<@P_FromDate>'', ''<@P_ToDate>'', 0, '''', 0) E'
                 + ' WHERE 1 = CASE WHEN LEN(ISNULL(''<@P_Category>'','''')) > 0 AND CHARINDEX(E.sCategory, ISNULL(''<@P_Category>'','''')) > 0 THEN 1 WHEN LEN(ISNULL(''<@P_Category>'','''')) = 0 THEN 1 ELSE 0 END) AS X'
                 + ' RIGHT JOIN (SELECT c1.sCategory Category, c1.hKey iCategory FROM tbl_CategoryList c1 LEFT JOIN dbo.Fn_ListObsoleteCategories(''<@P_FromDate>'', ''<@P_ToDate>'') Fn ON c1.hKey = Fn.iCategory WHERE Fn.iCategory IS NULL'
                 + ' AND 1 = CASE WHEN LEN(ISNULL(''<@P_Category>'','''')) > 0 AND CHARINDEX(C1.sCategory, ISNULL(''<@P_Category>'','''')) > 0 THEN 1 WHEN LEN(ISNULL(''<@P_Category>'','''')) = 0 THEN 1 ELSE 0 END)c ON c.iCategory = X.hKey'
                 + ' PIVOT (SUM(x.dAmount) FOR X.dtDate in (<@col2>)) AS pvt WHERE 1=2'
	             + ' UNION ALL'
	             + ' SELECT 2 Sort, NULL, ''TOTAL'', <@col3>'
                 + ' FROM (SELECT CONVERT(VARCHAR, YEAR(dtDate)) dtdate, E.sCategory Category ,dbo.Fn_dAmount(E.sNotes, E.items, E.dAmount) dAmount FROM dbo.Fn_DataStore(''<@P_FromDate>'', ''<@P_ToDate>'', 0, '''', 0) E'
                 + ' WHERE 1 = CASE WHEN LEN(ISNULL(''<@P_Category>'','''')) > 0 AND CHARINDEX(E.sCategory, ISNULL(''<@P_Category>'','''')) > 0 THEN 1 WHEN LEN(ISNULL(''<@P_Category>'','''')) = 0 THEN 1 ELSE 0 END) AS X'
	             + ' PIVOT (SUM(x.dAmount) FOR X.dtDate in (<@col2>)) AS pvt WHERE 1=2' 
	             + ' ORDER BY sort,category;'

				 + ' TRUNCATE TABLE #Temp2'

				 + ' INSERT INTO #Temp2 (Sort, iCategory, Category, <@col5>) SELECT 1 Sort, iCategory, Category, <@col5>'
                 + ' FROM (SELECT CONVERT(VARCHAR, YEAR(T.dtMonth)) dtdate, C1.sCategory, C1.hKey, T.TAmount FROM tbl_ExpThresholds T INNER JOIN tbl_CategoryList C1 ON C1.hKey = T.iCategory'
                 + ' WHERE T.dtMonth BETWEEN DATEADD(dd,-(DAY(''<@P_FromDate>'') -1), ''<@P_FromDate>'') AND DATEADD(dd,-(DAY(''<@P_ToDate>'') -1), ''<@P_ToDate>'') '
                 + ' AND 1 = CASE WHEN LEN(ISNULL(''<@P_Category>'','''')) > 0 AND CHARINDEX(C1.sCategory, ISNULL(''<@P_Category>'','''')) > 0 THEN 1 WHEN LEN(ISNULL(''<@P_Category>'','''')) = 0 THEN 1 ELSE 0 END) AS X'
                 + ' RIGHT JOIN (SELECT c1.sCategory Category, c1.hKey iCategory FROM tbl_CategoryList c1 LEFT JOIN dbo.Fn_ListObsoleteCategories(''<@P_FromDate>'', ''<@P_ToDate>'') Fn ON c1.hKey = Fn.iCategory WHERE Fn.iCategory IS NULL'
                 + ' AND 1 = CASE WHEN LEN(ISNULL(''<@P_Category>'','''')) > 0 AND CHARINDEX(C1.sCategory, ISNULL(''<@P_Category>'','''')) > 0 THEN 1 WHEN LEN(ISNULL(''<@P_Category>'','''')) = 0 THEN 1 ELSE 0 END)c ON c.iCategory = X.hKey'
                 + ' PIVOT (SUM(x.TAmount) FOR X.dtDate in (<@col5>)) AS pvt'
                 + ' UNION ALL'
	             + ' SELECT 2 Sort, NULL, ''TOTAL'', <@col6>'
                 + ' FROM (SELECT CONVERT(VARCHAR, YEAR(T.dtMonth)) dtdate, C1.sCategory, C1.hKey, T.TAmount FROM tbl_ExpThresholds T INNER JOIN tbl_CategoryList C1 ON C1.hKey = T.iCategory'
                 + ' WHERE T.dtMonth BETWEEN DATEADD(dd,-(DAY(''<@P_FromDate>'') -1), ''<@P_FromDate>'') AND DATEADD(dd,-(DAY(''<@P_ToDate>'') -1), ''<@P_ToDate>'') '
                 + ' AND 1 = CASE WHEN LEN(ISNULL(''<@P_Category>'','''')) > 0 AND CHARINDEX(C1.sCategory, ISNULL(''<@P_Category>'','''')) > 0 THEN 1 WHEN LEN(ISNULL(''<@P_Category>'','''')) = 0 THEN 1 ELSE 0 END) AS X'
                 + ' RIGHT JOIN (SELECT c1.sCategory Category, c1.hKey iCategory FROM tbl_CategoryList c1 LEFT JOIN dbo.Fn_ListObsoleteCategories(''<@P_FromDate>'', ''<@P_ToDate>'') Fn ON c1.hKey = Fn.iCategory WHERE Fn.iCategory IS NULL'
                 + ' AND 1 = CASE WHEN LEN(ISNULL(''<@P_Category>'','''')) > 0 AND CHARINDEX(C1.sCategory, ISNULL(''<@P_Category>'','''')) > 0 THEN 1 WHEN LEN(ISNULL(''<@P_Category>'','''')) = 0 THEN 1 ELSE 0 END)c ON c.iCategory = X.hKey'
                 + ' PIVOT (SUM(x.TAmount) FOR X.dtDate in (<@col5>)) AS pvt'
	             + ' ORDER BY sort,category;'

	  SET @sql = REPLACE(@sql, '<@col1>', @col1)
	  SET @sql = REPLACE(@sql, '<@col2>', @col2)
	  SET @sql = REPLACE(@sql, '<@col3>', @col3)
	  SET @sql = REPLACE(@sql, '<@col4>', @col4)
	  SET @sql = REPLACE(@sql, '<@col5>', @col5)
	  SET @sql = REPLACE(@sql, '<@col6>', @col6)
	  SET @sql = REPLACE(@sql, '<@P_FromDate>', @P_FromDate)
	  SET @sql = REPLACE(@sql, '<@P_ToDate>', @P_ToDate)
	  SET @sql = REPLACE(@sql, '<@P_Category>', @P_Category)
      
      EXEC(@sql + 'SELECT * FROM #Temp2;')
  END
GO
/*############################################################################################################################################################*/

/*IF OBJECT_ID('Fn_GetChartData_Categorywise') IS NOT NULL 
DROP FUNCTION Fn_GetChartData_Categorywise
GO

CREATE FUNCTION Fn_GetChartData_Categorywise(@P_Category AS VARCHAR(50))
RETURNS TABLE
AS
RETURN (SELECT CASE WHEN X.[Month] IS NULL THEN 'TOTAL' ELSE X.[Month] END [Month], X.Amount FROM (
SELECT DATENAME(mm,dtdate) [Month], SUM(dAmount) Amount FROM tbl_ExpenditureDet E INNER JOIN tbl_CategoryList C ON C.hKey = E.iCategory WHERE C.sCategory = @P_Category
group BY DATENAME(mm,dtdate) WITH ROLLUP ) X)
GO*/
/*############################################################################################################################################################*/

IF OBJECT_ID('sp_GetChartData_1') IS NOT NULL 
	DROP PROCEDURE sp_GetChartData_1
GO

CREATE PROCEDURE sp_GetChartData_1(@P_FromDate DATE, @P_ToDate DATE, @P_iCategory INTEGER, @P_SearchStr VARCHAR(500), @P_Years VARCHAR(MAX), @P_ExcludeZeroEntry BIT = 0)
AS
BEGIN	
	SELECT Y.Month, ISNULL(X.Amount, 0) Amount, ISNULL(X.Count, 0) Count
      FROM (SELECT LEFT(DATENAME(M, E.dtDate), 3) + '-' + CONVERT(CHAR(4), RIGHT(YEAR(E.dtDate), 2)) 'Month', 
				  SUM(dbo.Fn_dAmount(E.sNotes, E.items, E.dAmount)) 'Amount', 
				  COUNT(dbo.Fn_dAmount(E.sNotes, E.items, E.dAmount)) 'Count', 
			MONTH(E.dtDate) iMonth
              FROM dbo.Fn_DataStore(@P_FromDate, @P_ToDate, @P_iCategory, @P_SearchStr, 0) E
             WHERE (@P_Years != '' AND CHARINDEX(CONVERT(VARCHAR, YEAR(E.dtDate)), @P_Years) > 0)
             GROUP BY LEFT(DATENAME(M, e.dtDate), 3) + '-' + CONVERT(CHAR(4), RIGHT(YEAR(e.dtDate), 2)), MONTH(e.dtDate))X
            RIGHT JOIN (SELECT LEFT(DATENAME(M, CONVERT(DATE, '1900-' + CONVERT(VARCHAR, Z.iMonth) + '-01')), 3) + '-' + CONVERT(CHAR(4), RIGHT(Z.iYear, 2)) Month, Z.iMonth, Z.iYear
                          FROM  (SELECT X.iMonth, X.iYear
                                   FROM   (SELECT ROW_NUMBER() OVER(PARTITION BY YEAR(dtDate) ORDER BY dtDate) iMonth, YEAR(dtDate) iYear
											 FROM tbl_ExpenditureDet e
										    WHERE e.dtDate BETWEEN @P_FromDate AND @P_ToDate
												  AND @P_Years != '' AND CHARINDEX(CONVERT(VARCHAR, YEAR(E.dtDate)), @P_Years) > 0)X
                                  WHERE  X.iMonth <= 12)Z) Y ON Y.Month = X.Month
	 WHERE 1 = CASE WHEN ISNULL(@P_ExcludeZeroEntry, 0) = 0 THEN 1
					WHEN ISNULL(@P_ExcludeZeroEntry, 0) = 1 AND ISNULL(X.Amount, 0) != 0 THEN 1
					ELSE 0
				END
     ORDER BY Y.iYear, Y.iMonth
END
GO
/*############################################################################################################################################################*/

IF OBJECT_ID('sp_GetChartData_2') IS NOT NULL 
	DROP PROCEDURE sp_GetChartData_2
GO

CREATE PROCEDURE sp_GetChartData_2(@P_FromDate DATE, @P_ToDate DATE, @P_SearchStr VARCHAR(500), @P_Years VARCHAR(MAX), @P_ExcludeZeroEntry BIT = 0)
AS
BEGIN
	SELECT C.sCategory [Category], ISNULL(X.Amount, 0) Amount, ISNULL(X.Count, 0) Count
	  FROM (SELECT e.iCategory, 
				   SUM(dbo.Fn_dAmount(E.sNotes, E.items, E.dAmount)) 'Amount',
				   COUNT(dbo.Fn_dAmount(E.sNotes, E.items, E.dAmount)) 'Count' 
	          FROM dbo.Fn_DataStore(@P_FromDate, @P_ToDate, 0, @P_SearchStr, 0) E 
	         WHERE (@P_Years != '' AND CHARINDEX(CONVERT(VARCHAR, YEAR(E.dtDate)), @P_Years) > 0)
	          GROUP BY e.iCategory)X
			       RIGHT JOIN (SELECT c1.sCategory, c1.hKey, Fn.iCategory 
						         FROM tbl_CategoryList c1 
								      LEFT JOIN dbo.Fn_ListObsoleteCategories(@P_FromDate, @P_ToDate) Fn ON c1.hKey = Fn.iCategory 
						        WHERE Fn.iCategory IS NULL)c ON c.hKey = X.iCategory
	 WHERE 1 = CASE WHEN ISNULL(@P_ExcludeZeroEntry, 0) = 0 THEN 1
					WHEN ISNULL(@P_ExcludeZeroEntry, 0) = 1 AND ISNULL(X.Amount, 0) != 0 THEN 1
					ELSE 0
				END
	 ORDER BY C.sCategory
END
GO
/*############################################################################################################################################################*/

IF OBJECT_ID('sp_GetChartData_3') IS NOT NULL 
	DROP PROCEDURE sp_GetChartData_3
GO

CREATE PROCEDURE sp_GetChartData_3(@P_FromDate DATE, @P_ToDate DATE, @P_iCategory INTEGER, @P_SearchStr VARCHAR(500), @P_Years VARCHAR(MAX), @P_PeriodLimit BIT = 0)
AS
BEGIN
	DECLARE @Col1 VARCHAR(MAX) = '', 
			@Col2 VARCHAR(MAX) = '',
			@Col3 VARCHAR(MAX) = '',
			@Col4 VARCHAR(MAX) = '',
			@Sql  VARCHAR(MAX), 
			@Years INTEGER, 
			@CounterYears INTEGER = 0, 
			@StartingYear INTEGER
		
	SET @Years = DATEDIFF(YYYY, @P_FromDate, @P_ToDate)
	SET @StartingYear = YEAR(@P_FromDate)
		
	WHILE @CounterYears <= @Years
	BEGIN
		IF CHARINDEX(CONVERT(VARCHAR, (@StartingYear + @CounterYears)), @P_Years) > 0
		BEGIN
			SET @Col1 = @Col1 + 'ISNULL(['  + CONVERT(VARCHAR, (@StartingYear + @CounterYears)) + '],0) ['  + CONVERT(VARCHAR, (@StartingYear + @CounterYears)) + '], '
			SET @Col2 = @Col2 + '['  + CONVERT(VARCHAR, (@StartingYear + @CounterYears)) + '], '
			SET @Col3 = @Col3 + 'ISNULL(['  + CONVERT(VARCHAR, (@StartingYear + @CounterYears)) + '_CNT],0) ['  + CONVERT(VARCHAR, (@StartingYear + @CounterYears)) + '_CNT], '
			SET @Col4 = @Col4 + '['  + CONVERT(VARCHAR, (@StartingYear + @CounterYears)) + '_CNT], '
		END
		SET @CounterYears = @CounterYears + 1
		CONTINUE
	END
	
	SET @Col1 = SUBSTRING (@Col1,1, Len(@Col1)- 1)
	SET @Col2 = SUBSTRING (@Col2,1, Len(@Col2)- 1)
	SET @Col3 = SUBSTRING (@Col3,1, Len(@Col3)- 1)
	SET @Col4 = SUBSTRING (@Col4,1, Len(@Col4)- 1)
	

	SET @Sql = 'SELECT * INTO #Temp_Data 
				  FROM (SELECT Y.Month, L.Amount, L.iYear, Y.iMonth 
			             FROM (SELECT DATENAME(MM,dtDate) [Month], dbo.Fn_dAmount(E.sNotes, E.items, E.dAmount) Amount, YEAR(dtDate) iYear, Month(dtDate) MonthNum
						         FROM dbo.Fn_DataStore(''<@P_FromDate>'', ''<@P_ToDate>'', <@P_iCategory>, ''<@P_SearchStr>'', <@P_PeriodLimit>) E
								WHERE (''<@P_Years>'' != '''' AND CHARINDEX(CONVERT(VARCHAR, YEAR(E.dtDate)), ''<@P_Years>'') > 0)) L
							   RIGHT JOIN (SELECT DATENAME(M, CONVERT(DATE, ''1900-'' + CONVERT(VARCHAR, Z.iMonth) + ''-01'')) Month, Z.iMonth, Z.iYear  
					                         FROM (SELECT X.iMonth, X.iYear 
													FROM (SELECT Row_number() OVER (PARTITION BY YEAR(dtDate) ORDER BY dtDate) iMonth, YEAR(dtDate) iYear 
														    FROM tbl_ExpenditureDet e 
														   WHERE e.dtDate BETWEEN ''<@P_FromDate>'' AND ''<@P_ToDate>''
															 AND (''<@P_Years>'' != '''' 
															 AND CHARINDEX(CONVERT(VARCHAR, YEAR(E.dtDate)), ''<@P_Years>'') > 0))X 
					                               WHERE X.iMonth <= 12)Z) Y ON Y.Month = L.Month AND Y.iYear = L.iYear) Z; ' + 

				'SELECT X.Month, <@Col1>, <@Col3> 
				   FROM (SELECT Month, iMonth, <@Col1>
			               FROM (SELECT Month, Amount, iYear, iMonth FROM #Temp_Data) T
			              PIVOT (SUM(T.Amount) FOR T.iYear IN (<@Col2>)) AS pvt) X
			            INNER JOIN (SELECT Month, iMonth, <@Col3>
			                         FROM (SELECT Month, Amount, CONVERT(VARCHAR, iYear) + ''_CNT'' iYear, iMonth FROM #Temp_Data) T
			                        PIVOT (COUNT(T.Amount) FOR T.iYear IN (<@Col4>)) AS pvt) Y ON X.Month = Y.Month
			   ORDER BY X.iMonth'
	
	  SET @sql = REPLACE(@sql, '<@col1>', @col1)
	  SET @sql = REPLACE(@sql, '<@col2>', @col2)
	  SET @sql = REPLACE(@sql, '<@col3>', @col3)
	  SET @sql = REPLACE(@sql, '<@col4>', @col4)
	  SET @sql = REPLACE(@sql, '<@P_FromDate>', @P_FromDate)
	  SET @sql = REPLACE(@sql, '<@P_ToDate>', @P_ToDate)
	  SET @sql = REPLACE(@sql, '<@P_iCategory>', @P_iCategory)
	  SET @sql = REPLACE(@sql, '<@P_SearchStr>', @P_SearchStr)
	  SET @sql = REPLACE(@sql, '<@P_Years>', @P_Years)
	  SET @sql = REPLACE(@sql, '<@P_PeriodLimit>', @P_PeriodLimit)
	  
	EXEC(@Sql)
END
GO
/*############################################################################################################################################################*/

IF OBJECT_ID('sp_GetChartData_3B') IS NOT NULL 
	DROP PROCEDURE sp_GetChartData_3B
GO

CREATE PROCEDURE sp_GetChartData_3B(@P_FromDate DATE, @P_ToDate DATE, @P_iMonth INTEGER, @P_SearchStr VARCHAR(500), @P_Years VARCHAR(MAX), @P_PeriodLimit BIT = 0)
AS
BEGIN
	DECLARE @col1 VARCHAR(MAX),  
			@col2 VARCHAR(MAX),
			@col3 VARCHAR(MAX),  
			@col4 VARCHAR(MAX),
			@sql VARCHAR(MAX)
	
	SELECT DISTINCT CONVERT(VARCHAR, YEAR(E.dtDate)) dtDate INTO #Temp 
	  FROM dbo.Fn_DataStore(@P_FromDate, @P_ToDate, 0, @P_SearchStr, @P_PeriodLimit) E
	 WHERE (@P_Years != '' AND CHARINDEX(CONVERT(VARCHAR, YEAR(E.dtDate)), @P_Years) > 0)
				
	SET @col1  = STUFF((SELECT ',ISNULL([' + dtDate + '],0) [' + dtDate + ']' FROM #Temp FOR XML PATH('')),1,1,'')
	SET @col2  = STUFF((SELECT ',[' + dtDate + ']' FROM #Temp FOR XML PATH('')),1,1,'')
	SET @col3  = STUFF((SELECT ',ISNULL([' + dtDate + '_CNT],0) [' + dtDate + '_CNT]' FROM #Temp FOR XML PATH('')),1,1,'')
	SET @col4  = STUFF((SELECT ',[' + dtDate + '_CNT]' FROM #Temp FOR XML PATH('')),1,1,'')
	
	SET @sql = 'SELECT * INTO #Temp_Data
				  FROM (SELECT E.sCategory, E.iCategory, dbo.Fn_dAmount(E.sNotes, E.items, E.dAmount) Amount, YEAR(dtDate) [iYear]
						  FROM dbo.Fn_DataStore(''<@P_FromDate>'', ''<@P_ToDate>'', 0, ''<@P_SearchStr>'', <@P_PeriodLimit>) E
						 WHERE 1 = CASE WHEN <@P_iMonth> = 0 THEN 1 WHEN DATEPART(M, E.dtDate) = <@P_iMonth> THEN 1 ELSE 0 END 
						   AND (''<@P_Years>'' != '''' AND CHARINDEX(CONVERT(VARCHAR, YEAR(E.dtDate)), ''<@P_Years>'') > 0))X
					   RIGHT JOIN (SELECT c1.sCategory Category, c1.hKey 
									 FROM tbl_CategoryList c1 
										  LEFT JOIN dbo.Fn_ListObsoleteCategories(''<@P_FromDate>'', ''<@P_ToDate>'') Fn ON c1.hKey = Fn.iCategory 
									WHERE Fn.iCategory IS NULL)c ON c.hKey = X.iCategory;

				SELECT X.Category, <@Col1>, <@Col3>
				FROM (SELECT Category, <@Col1> 
						FROM (SELECT T.sCategory Category, T.iCategory, T.Amount, T.iYear FROM #Temp_Data T) T
					   PIVOT (SUM(T.Amount) FOR T.iYear IN (<@Col2>)) AS PVT) X
					 INNER JOIN (SELECT Category, <@Col3> 
								   FROM (SELECT T.sCategory Category, T.iCategory, T.Amount, CONVERT(VARCHAR, T.iYear) + ''_CNT'' iYear FROM #Temp_Data T) T
							      PIVOT (COUNT(T.Amount) FOR T.iYear IN (<@Col4>)) AS PVT) Y ON X.Category = Y.Category
			   ORDER BY X.Category;'

    SET @sql = REPLACE(@sql, '<@col1>', @col1)
	SET @sql = REPLACE(@sql, '<@col2>', @col2)
	SET @sql = REPLACE(@sql, '<@col3>', @col3)
	SET @sql = REPLACE(@sql, '<@col4>', @col4)
	SET @sql = REPLACE(@sql, '<@P_FromDate>', @P_FromDate)
	SET @sql = REPLACE(@sql, '<@P_ToDate>', @P_ToDate)
	SET @sql = REPLACE(@sql, '<@P_iMonth>', @P_iMonth)
	SET @sql = REPLACE(@sql, '<@P_SearchStr>', @P_SearchStr)
	SET @sql = REPLACE(@sql, '<@P_Years>', @P_Years)
	SET @sql = REPLACE(@sql, '<@P_PeriodLimit>', @P_PeriodLimit)
	  
	EXEC (@sql)
END
GO
/*############################################################################################################################################################*/

IF OBJECT_ID('sp_GetChartData_4') IS NOT NULL 
	DROP PROCEDURE sp_GetChartData_4
GO

CREATE PROCEDURE sp_GetChartData_4(	@P_FromDate DATE, 
									@P_ToDate DATE, 
									@P_iCategory INTEGER, 
									@P_SearchStr VARCHAR(500), 
									@P_Years VARCHAR(MAX), 
									@P_PeriodLimit BIT = 0,
									@P_ExcludeZeroEntry BIT = 0)
AS
BEGIN
	SELECT X.[Year], X.Amount, X.[Count]
	  FROM (SELECT CONVERT(VARCHAR, YEAR(E.dtDate)) [Year], 
				   SUM(dbo.Fn_dAmount(E.sNotes, E.items, E.dAmount)) [Amount],
				   COUNT(dbo.Fn_dAmount(E.sNotes, E.items, E.dAmount)) [Count]
	          FROM dbo.Fn_DataStore(@P_FromDate, @P_ToDate, @P_iCategory, @P_SearchStr, @P_PeriodLimit) E
	         WHERE @P_Years  != '' 
			   AND CHARINDEX(CONVERT(VARCHAR, YEAR(E.dtDate)), @P_Years ) > 0
			 GROUP BY CONVERT(VARCHAR, YEAR(E.dtDate))) X
	 WHERE 1 = CASE WHEN ISNULL(@P_ExcludeZeroEntry, 0) = 0 THEN 1
					WHEN ISNULL(@P_ExcludeZeroEntry, 0) = 1 AND ISNULL(X.Amount, 0) != 0 THEN 1
					ELSE 0
				END
END
GO
/*############################################################################################################################################################*/

IF OBJECT_ID('sp_PopulateSelectYearCombo') IS NOT NULL 
	DROP PROCEDURE sp_PopulateSelectYearCombo
GO

CREATE PROCEDURE sp_PopulateSelectYearCombo(@P_StartDate DATE)
AS
BEGIN
	 DECLARE @YrDiff INT = YEAR(GETDATE()) - YEAR(@P_StartDate), @Counter INT = 0
	 
	 WHILE @Counter < @YrDiff
	 BEGIN
		INSERT INTO tbl_EOY(Year#, IsYrClosed) 
		SELECT YEAR(@P_StartDate) + @Counter, 0 FROM (SELECT 1 Dummy) X
		WHERE NOT EXISTS(SELECT 1 FROM tbl_EOY E2 WHERE E2.Year# = (YEAR(@P_StartDate) + @Counter))
		      AND EXISTS (SELECT DISTINCT YEAR(dtDate) FROM tbl_ExpenditureDet WHERE YEAR(dtDate) = YEAR(@P_StartDate) + @Counter)
		
		SET @Counter = @Counter + 1
	
		CONTINUE
	 END

	DELETE FROM tbl_EOY WHERE Year# < YEAR(@P_StartDate) AND IsYrClosed = 0

	SELECT Year#, IsYrClosed FROM tbl_EOY WHERE Year# >= YEAR(@P_StartDate) ORDER BY Year#
END
GO
/*############################################################################################################################################################*/

IF OBJECT_ID('sp_SearchExpenses') IS NOT NULL 
	DROP PROCEDURE sp_SearchExpenses
GO

CREATE PROCEDURE sp_SearchExpenses(@P_FromDate DATE, @P_ToDate DATE, @P_Category VARCHAR(MAX), @P_SearchStr VARCHAR(500))  
AS
BEGIN
WITH X (Sort, RowNum, hKey, dtDate, sCategory, dAmount, sNotes, iCategory, SearchAmt, items, dtDateOrderBy, IsReadOnly, CategoryName )
AS
	(SELECT 1 [Sort], NULL RowNum, NULL [hKey], 'Keyword:- '+ Fn.items [dtDate], NULL [sCategory], NULL [dAmount], NULL [sNotes], NULL [iCategory], NULL [dSearchAmt], Fn.items [items], NULL [dtDateOrderBy], 1 [IsReadOnly], NULL [CategoryName]
			  FROM  dbo.Fn_searchstrsplit(@P_SearchStr, ',') Fn
			
	 UNION ALL

	SELECT 2,
		   ROW_NUMBER () OVER (PARTITION BY E.items ORDER BY E.items, E.sCategory, E.dtDate),
		   E.hKey,
		   CONVERT(VARCHAR, E.dtDate, 103),
		   E.sCategory,
		   E.dAmount,
		   E.sNotes,
		   E.iCategory,
		   dbo.Fn_dAmount(E.sNotes, E.items, E.dAmount),
		   E.items,
		   E.dtDate,
		   0,
		   E.sCategory [CategoryName]
	FROM   dbo.Fn_DataStore(@P_FromDate, @P_ToDate, 0, @P_SearchStr, 0) E
		   LEFT OUTER JOIN tbl_EOY Y ON Y.YEAR# = YEAR(E.dtDate)
	WHERE  1 = CASE WHEN LEN(ISNULL(@P_Category, '')) > 0 AND CHARINDEX(E.sCategory, ISNULL(@P_Category, '')) > 0 THEN 1
					WHEN LEN(ISNULL(@P_Category, '')) = 0 THEN 1
					WHEN CHARINDEX('ALL', ISNULL(@P_Category, '')) > 0 THEN 1
					ELSE 0
			    END 
	UNION ALL
    
	SELECT DISTINCT 
		   3,
		   NULL,
		   NULL,
		   'TOTAL',
		   NULL,
		   SUM(E.dAmount),
		   NULL,
		   NULL,
		   SUM(dbo.Fn_dAmount(E.sNotes, E.items, E.dAmount)),
		   E.items,
		   NULL,
		   1,
		   NULL
	FROM   dbo.Fn_DataStore(@P_FromDate, @P_ToDate, 0, @P_SearchStr, 0) E
	WHERE  1 = CASE WHEN LEN(ISNULL(@P_Category, '')) > 0 AND CHARINDEX(E.sCategory, ISNULL(@P_Category, '')) > 0 THEN 1
					WHEN LEN(ISNULL(@P_Category, '')) = 0 THEN 1
					WHEN CHARINDEX('ALL', ISNULL(@P_Category, '')) > 0 THEN 1
					ELSE 0
			    END
	GROUP  BY E.items WITH ROLLUP)


	SELECT	Y.Sort, Y.hKey, 
			CASE WHEN (Y.Sort = 3 AND Y.hKey IS NULL AND Y.dtDate = 'TOTAL' AND Y.Notes IS NULL) THEN 'GRAND TOTAL' 
				 ELSE Y.dtDate 
			 END [Date], Y.iCategory, Y.sCategory, Y.SearchAmt [Amount], Y.Notes, Y.IsReadOnly, NULL [bDelete], 0 [IsDummy], 0 [IsDummyRowAdded], Y.CategoryName [CategoryName], NULL [DateOriginal]
	  FROM (SELECT X.Sort, X.hKey, X.dtDate, X.iCategory, X.SearchAmt,
				   CASE WHEN X.Sort = 3 THEN (SELECT 'Count = '+ CONVERT(VARCHAR,MAX(Y.RowNum)) FROM X Y WHERE Y.items  = X.items GROUP BY Y.items) 
						ELSE X.sNotes 
					END [Notes],
				   X.IsReadOnly, X.sCategory, X.dtDateOrderBy, X.items, X.CategoryName
			  FROM X) Y
      ORDER  BY CASE WHEN Y.items IS NULL THEN 'zzz' 
					 ELSE Y.items 
				 END, Y.Sort, Y.dtDateOrderBy, Y.sCategory
END
GO
/*############################################################################################################################################################*/

IF OBJECT_ID('sp_GetExpensesThresholds') IS NOT NULL 
	DROP PROCEDURE sp_GetExpensesThresholds
GO

CREATE PROCEDURE sp_GetExpensesThresholds(@P_FromDate DATE, @P_ToDate DATE)
AS
BEGIN
	
	DECLARE @SingleMonth AS BIT = 0
	
	IF MONTH(@P_FromDate) = MONTH(@P_ToDate) AND YEAR(@P_FromDate) = YEAR(@P_ToDate) 
		SET @SingleMonth = 1
		
	INSERT INTO tbl_ExpThresholds (dtMonth, iCategory, TAmount)
	SELECT X.dtMonth, X.iCategory, X.TAmount 
	  FROM(SELECT @P_FromDate [dtMonth], T.iCategory, T.TAmount, C.sCategory
			 FROM tbl_ExpThresholds T
				  INNER JOIN tbl_CategoryList C ON C.hKey = T.iCategory AND C.IsObsolete = 0
			WHERE T.dtMonth = DATEADD(M, -1, @P_FromDate)
			  AND NOT EXISTS (SELECT 1 FROM tbl_ExpThresholds T1 WHERE T1.dtMonth = @P_FromDate)
			  AND @SingleMonth = 1
			   
			UNION ALL
	
			SELECT @P_FromDate, C.hKey, 0, C.sCategory 
			  FROM tbl_CategoryList C
			 WHERE (SELECT COUNT(*) FROM tbl_ExpThresholds) = 0
			   AND @SingleMonth = 1 
			   AND C.IsObsolete = 0
			   ) X
	 ORDER BY X.sCategory
	
	INSERT INTO tbl_ExpThresholds (dtMonth, iCategory, TAmount)
	SELECT @P_FromDate, C.hKey, 0 FROM tbl_CategoryList C WHERE NOT EXISTS( SELECT 1 FROM tbl_ExpThresholds T WHERE T.iCategory = C.hKey)
	
	SELECT	CASE WHEN @SingleMonth = 1 THEN X.hKey ELSE 0 END [hKey], X.iCategory, X.sCategory [Category],
			CASE WHEN @SingleMonth = 1 THEN X.dtMonth ELSE '1900-01-01' END [dtMonth],
			SUM(X.TAmount) [TAmount], SUM(X.SAmount) [SAmount], SUM(X.Diff) [Difference], SUM(X.Diff_Sign) [DifferenceSign],  
			CASE WHEN @SingleMonth = 1 THEN 0 ELSE 1 END [IsReadOnly]
	 FROM  (SELECT	T.hKey, T.iCategory, T.dtMonth, C.sCategory, T.TAmount, 
					SUM(ISNULL(E.dAmount, 0)) [SAmount], ABS(SUM(ISNULL(E.dAmount, 0)) - T.TAmount) [Diff], (SUM(ISNULL(E.dAmount, 0)) - T.TAmount) [Diff_Sign]
			  FROM tbl_ExpThresholds T 
				   INNER JOIN tbl_CategoryList C ON C.hKey = T.iCategory AND C.IsObsolete = 0
				   LEFT JOIN tbl_ExpenditureDet E ON T.iCategory = E.iCategory  AND T.dtMonth = DATEADD(dd,-(DAY(E.dtDate)- 1),E.dtDate)
			 WHERE T.dtMonth BETWEEN @P_FromDate AND @P_ToDate
			 GROUP BY T.hKey, T.iCategory, T.dtMonth, C.sCategory, T.TAmount)X
	GROUP BY CASE WHEN @SingleMonth = 1 THEN X.hKey ELSE 0 END, X.iCategory, 
			 CASE WHEN @SingleMonth = 1 THEN X.dtMonth ELSE '1900-01-01' END, X.sCategory
	ORDER BY X.sCategory
END
GO
/*############################################################################################################################################################*/

IF OBJECT_ID('sp_GetExpensesTicker') IS NOT NULL 
	DROP PROCEDURE sp_GetExpensesTicker
GO

CREATE PROCEDURE sp_GetExpensesTicker(@P_FromDate DATE, @P_ToDate DATE)
AS
BEGIN
	DECLARE @FromMonthBudgetEntryCnt AS NUMERIC(18,0) = 0
	SET @FromMonthBudgetEntryCnt = (SELECT COUNT(*) FROM tbl_ExpThresholds T1 WHERE  dtMonth = @P_FromDate)
	
	SELECT C.sCategory [Category], ISNULL(X.Amount, 0) Amount, T.TAmount,
		   CASE
			 WHEN @FromMonthBudgetEntryCnt = 0 THEN -99
			 WHEN ISNULL(X.Amount, 0) > T.TAmount THEN 1
			 WHEN ISNULL(X.Amount, 0) = T.TAmount THEN 0
			 ELSE -1
		   END [IsOverBudget]
	FROM   (SELECT e.iCategory, SUM(E.dAmount) [Amount]
			  FROM tbl_ExpenditureDet e
			 WHERE e.dtDate BETWEEN @P_FromDate AND @P_ToDate
			 GROUP BY e.iCategory)X
		   RIGHT JOIN (SELECT T1.iCategory, SUM(T1.TAmount) TAmount FROM tbl_ExpThresholds T1 WHERE  dtMonth BETWEEN @P_FromDate AND @P_ToDate GROUP BY T1.iCategory)T ON T.iCategory = X.iCategory
		   RIGHT JOIN (SELECT c1.sCategory, c1.hKey
						 FROM tbl_CategoryList c1
							  LEFT JOIN dbo.Fn_listobsoletecategories(@P_FromDate, @P_ToDate) Fn ON c1.hKey = Fn.iCategory 
						WHERE Fn.iCategory IS NULL)c ON c.hKey = T.iCategory 
   ORDER BY C.sCategory
END
GO
/*############################################################################################################################################################*/

IF OBJECT_ID('sp_ImportData') IS NOT NULL 
	DROP PROCEDURE sp_ImportData
GO

CREATE PROCEDURE sp_ImportData
AS
BEGIN
	SET DATEFORMAT DMY
	
	DECLARE @dtDate VARCHAR(10),
	@sCategory VARCHAR(100),
	@dAmount VARCHAR(20),
	@sNotes VARCHAR(500),
	@sImportComments VARCHAR(500),
	@Counter INTEGER = 1
		
	DELETE FROM Temp_ImportData WHERE iDelete = 1	
		
	DECLARE ImportCursor CURSOR FOR SELECT TMP.dtDate, TMP.sCategory, TMP.dAmount, TMP.sNotes, TMP.sImportComments FROM Temp_ImportData TMP

	OPEN ImportCursor
	
	FETCH NEXT FROM ImportCursor INTO @dtDate, @sCategory, @dAmount, @sNotes, @sImportComments
	
	WHILE @@FETCH_STATUS = 0   
	BEGIN
	
		IF  @dtDate IS NULL OR @sCategory IS NULL OR @dAmount IS NULL
		BEGIN
			UPDATE Temp_ImportData
			   SET sImportComments = 'Record No. #' + CONVERT(VARCHAR, @Counter) + ' failed to import. Missing required field(s).',
				   @sImportComments = 'Record No. #' + CONVERT(VARCHAR, @Counter) + ' failed to import. Missing required field(s).',
				   iRowNumber = @Counter
		     WHERE CURRENT OF ImportCursor
		END
	
		IF ISDATE(@dtDate) = 0
		BEGIN
			UPDATE Temp_ImportData
			   SET sImportComments = 'Record No. #' + CONVERT(VARCHAR, @Counter) + ' failed to import. Invalid DATE format.',
				   @sImportComments = 'Record No. #' + CONVERT(VARCHAR, @Counter) + ' failed to import. Invalid DATE format.',
				   iRowNumber = @Counter
		     WHERE CURRENT OF ImportCursor
		END
		
		IF ISNUMERIC(@dAmount) = 0
		BEGIN
			UPDATE Temp_ImportData
			   SET sImportComments = 'Record No. #' + @Counter + ' failed to import. Invalid AMOUNT format.',
			       @sImportComments = 'Record No. #' + @Counter + ' failed to import. Invalid AMOUNT format.',
				   iRowNumber = @Counter
			 WHERE CURRENT OF ImportCursor
		END
		
		IF (SELECT COUNT(*) FROM tbl_CategoryList C WHERE Lower(C.sCategory) = Lower(@sCategory)) = 0
		BEGIN
			UPDATE Temp_ImportData
			   SET sImportComments = 'Record No. #' + CONVERT(VARCHAR, @Counter) + ' failed to import. Invalid CATEGORY.',
			   @sImportComments = 'Record No. #' + CONVERT(VARCHAR, @Counter) + ' failed to import. Invalid CATEGORY.',
			   iRowNumber = @Counter
			 WHERE CURRENT OF ImportCursor
		END
		
		IF ISDATE(@dtDate) = 1 AND ISNUMERIC(@dAmount) = 1
		BEGIN     
		  IF (SELECT COUNT(*)   
			   FROM tbl_ExpenditureDet E INNER JOIN tbl_CategoryList C ON C.hKey = E.iCategory   
			  WHERE E.dtDate = @dtDate AND C.sCategory = @sCategory AND E.dAmount =  @dAmount /*AND ISNULL(E.sNotes,'') =  ISNULL(@sNotes,'') */) > 0  
		  BEGIN  
		   UPDATE Temp_ImportData  
			  SET sImportComments = 'Record No. #' + CONVERT(VARCHAR, @Counter) + ' skipped. Record was already imported.',  
				  @sImportComments = 'Record No. #' + CONVERT(VARCHAR, @Counter) + ' skipped. Record was already imported.',
				  iRowNumber = @Counter			  
			WHERE CURRENT OF ImportCursor  
		  END  
		END
	
		IF @sImportComments IS NULL
		BEGIN
			UPDATE Temp_ImportData
			   SET sImportComments = 'Record No. #' + CONVERT(VARCHAR, @Counter) + ' imported successfully.',
			       iRowNumber = @Counter
			 WHERE CURRENT OF ImportCursor
		END
	
		SET @Counter  += 1
		
		FETCH NEXT FROM ImportCursor INTO @dtDate, @sCategory, @dAmount, @sNotes, @sImportComments
	END

	CLOSE ImportCursor
	DEALLOCATE ImportCursor

	INSERT INTO tbl_ExpenditureDet(dtDate, iCategory, dAmount, sNotes, iImportedRecord)
	SELECT CONVERT(DATE, TMP.dtDate, 105), C.hKey, TMP.dAmount, TMP.sNotes, 1 
	  FROM Temp_ImportData TMP
			INNER JOIN tbl_CategoryList C ON C.sCategory = TMP.sCategory 
			LEFT JOIN  tbl_ExpenditureDet T ON T.dtDate = CONVERT(DATE, TMP.dtDate, 105) AND T.iCategory = C.hKey AND T.dAmount = TMP.dAmount AND T.sNotes = TMP.sNotes
	 WHERE	CHARINDEX('success', ISNULL(TMP.sImportComments,'')) > 0
			AND T.dtDate IS NULL
			AND T.iCategory IS NULL
			AND T.dAmount IS NULL
			AND T.sNotes IS NULL
	 ORDER BY CONVERT(DATE, TMP.dtDate, 105)	

	UPDATE Temp_ImportData SET iDelete = 1

	SELECT 	CONVERT(VARCHAR, CONVERT(DATE, dtDate), 103) [Date], 
			TMP.sCategory [Category], 
			CONVERT(MONEY, TMP.dAmount) [Amount], 
			TMP.sNotes [Notes], 
			TMP.sImportComments,
		    C.hKey [iCategory]
	FROM Temp_ImportData TMP
		 LEFT JOIN tbl_CategoryList C ON C.sCategory = TMP.sCategory
   ORDER BY iRowNumber
END
GO
/*############################################################################################################################################################*/

IF OBJECT_ID('sp_GetUnpaidBillsCurrentMonth') IS NOT NULL 
	DROP PROCEDURE sp_GetUnpaidBillsCurrentMonth
GO

CREATE PROCEDURE sp_GetUnpaidBillsCurrentMonth
AS
BEGIN
	DECLARE @StartOfMonth DATE,
			@EndOfMonth DATE

	SET @StartOfMonth = CONVERT(DATE, DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()), 0))
	SET @EndOfMonth = CONVERT(DATE, DATEADD(MONTH,1+DATEDIFF(MONTH,0,GETDATE()),-1))

	SELECT C.sCategory + ' (' + CONVERT(VARCHAR, COUNT(C.iMonthlyOccurrences - E.iCategory)) + '/' + CONVERT(VARCHAR, C.iMonthlyOccurrences) + ')' UnpaidCategory
	  FROM tbl_CategoryList C 
		   LEFT JOIN tbl_ExpenditureDet E ON C.hKey = E.iCategory AND E.dtDate BETWEEN  @StartOfMonth AND @EndOfMonth
     WHERE C.IsObsolete = 0
	   AND C.bNotify = 1
     GROUP BY C.sCategory, C.iMonthlyOccurrences
    HAVING C.iMonthlyOccurrences > COUNT(E.iCategory)
END
GO
/*############################################################################################################################################################*/

IF OBJECT_ID('sp_GetUnpaidBillsPreviousMonth') IS NOT NULL 
	DROP PROCEDURE sp_GetUnpaidBillsPreviousMonth
GO

CREATE PROCEDURE sp_GetUnpaidBillsPreviousMonth
AS
BEGIN
	DECLARE @StartOfPrevMonth DATE,
			@EndOfPrevMonth DATE

	SET @StartOfPrevMonth = CONVERT(DATE,DATEADD(MM, DATEDIFF(MM, 0, GETDATE())-1, 0))
	SET @EndOfPrevMonth = CONVERT(DATE,DATEADD(MS, -3, DATEADD(MM, DATEDIFF(MM, 0, GETDATE()) , 0)))

	SELECT C.sCategory + ' (' + CONVERT(VARCHAR, COUNT(C1.iMonthlyOccurrences - E.iCategory)) + '/' + CONVERT(VARCHAR, C1.iMonthlyOccurrences) + ')' UnpaidCategory
	  FROM tbl_CategoryList C
		   INNER JOIN  tbl_CategoryPrevMonthOccurrences C1 ON C.hKey = C1.iCategory
		   LEFT JOIN tbl_ExpenditureDet E ON C.hKey = E.iCategory AND E.dtDate BETWEEN  @StartOfPrevMonth AND @EndOfPrevMonth
     WHERE C.IsObsolete = 0
	   AND C.bNotify = 1
     GROUP BY C.sCategory, C1.iMonthlyOccurrences
    HAVING C1.iMonthlyOccurrences > COUNT(E.iCategory)
END
GO
/*############################################################################################################################################################*/

IF OBJECT_ID('sp_InsertCategoryPrevMonthOccurrences') IS NOT NULL
	DROP PROCEDURE sp_InsertCategoryPrevMonthOccurrences
GO

CREATE PROCEDURE sp_InsertCategoryPrevMonthOccurrences
AS
BEGIN
	DECLARE @PrevMonth DATE
	
	SET @PrevMonth = CONVERT(DATE,DATEADD(MM, DATEDIFF(MM, 0, GETDATE())-1, 0))

	TRUNCATE TABLE tbl_CategoryPrevMonthOccurrences
	
	INSERT INTO tbl_CategoryPrevMonthOccurrences(dtMonth, iCategory, iMonthlyOccurrences)
	SELECT @PrevMonth, C.hKey, C.iMonthlyOccurrences
	  FROM tbl_CategoryList C
	 WHERE C.IsObsolete = 0
	   AND C.bNotify = 1
END
GO
/*############################################################################################################################################################*/

IF Object_id('sp_GetExpenditureSummary_AllInOne') IS NOT NULL
  DROP PROCEDURE sp_GetExpenditureSummary_AllInOne
GO

CREATE PROCEDURE sp_GetExpenditureSummary_AllInOne(	@P_InceptionDate DATE, 
													@P_Category  VARCHAR(MAX), 
													@P_SearchStr VARCHAR(500), 
													@P_FromDate DATE = NULL,
													@P_ToDate DATE = NULL)
AS
  BEGIN
	DECLARE	@sql VARCHAR(MAX),
			@FromDate DATE,
			@ToDate DATE
	
	CREATE TABLE #MTD(iCategory NUMERIC(18,0), MTD MONEY)
	CREATE TABLE #YTD(iCategory NUMERIC(18,0), YTD MONEY)
	CREATE TABLE #ITD(iCategory NUMERIC(18,0), ITD MONEY)
	CREATE TABLE #Temp_AllInOne(Sort INT, iCategory NUMERIC(18,0), sCategory VARCHAR(100), dAmount MONEY, Cnt INT, Bdg MONEY, Total MONEY)
	
	--MTD
	IF @P_FromDate IS NULL
		SET	@FromDate = DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()), 0) --First Day of Current Month
	ELSE
		SET	@FromDate = @P_FromDate
	
	IF @P_ToDate IS NULL
		SET @ToDate = GETDATE()
	ELSE
		SET	@ToDate = @P_ToDate
	
	INSERT INTO #Temp_AllInOne EXEC sp_GetExpenditureSummary_Monthly @FromDate, @ToDate, @P_Category, @P_SearchStr
	INSERT INTO #MTD(iCategory, MTD) SELECT iCategory, dAmount FROM #Temp_AllInOne
	
	--YTD
	SET	@FromDate = CONVERT(DATE, '01-01-'+CONVERT(VARCHAR, YEAR(GETDATE()))) --First Day of current year
	
	TRUNCATE TABLE #Temp_AllInOne
	
	INSERT INTO #Temp_AllInOne EXEC sp_GetExpenditureSummary_Yearly @FromDate, @ToDate, @P_Category, @P_SearchStr, NULL
	INSERT INTO #YTD(iCategory, YTD) SELECT iCategory, dAmount FROM #Temp_AllInOne

	TRUNCATE TABLE #Temp_AllInOne

	--ITD
    SET @sql = 'SELECT 1 Sort, iCategory, Category, ITD, NULL, NULL, [TOTAL] 
				  FROM (SELECT ''ITD'' dtdate, E.sCategory, E.iCategory hKey, dbo.Fn_dAmount(E.sNotes, E.items, E.dAmount) dAmount 
						  FROM dbo.Fn_DataStore(''<@P_InceptionDate>'', ''<@ToDate>'', 0, ''<@P_SearchStr>'', 0) E
						 WHERE 1 = CASE WHEN LEN(ISNULL(''<@P_Category>'','''')) > 0 AND CHARINDEX(E.sCategory, ISNULL(''<@P_Category>'','''')) > 0 THEN 1 WHEN LEN(ISNULL(''<@P_Category>'','''')) = 0 THEN 1 ELSE 0 END) AS X
					   RIGHT JOIN (SELECT c1.sCategory Category, c1.hKey iCategory 
								     FROM tbl_CategoryList c1 
										  LEFT JOIN dbo.Fn_ListObsoleteCategories(''<@P_InceptionDate>'', ''<@ToDate>'') Fn ON c1.hKey = Fn.iCategory 
								    WHERE Fn.iCategory IS NULL
									  AND 1 = CASE WHEN LEN(ISNULL(''<@P_Category>'','''')) > 0 AND CHARINDEX(C1.sCategory, ISNULL(''<@P_Category>'','''')) > 0 THEN 1 WHEN LEN(ISNULL(''<@P_Category>'','''')) = 0 THEN 1 ELSE 0 END)c ON c.iCategory = X.hKey
					   PIVOT (SUM(x.dAmount) FOR X.dtDate in (ITD, [TOTAL])) AS pvt
	           UNION ALL
	           SELECT 2 Sort, 999, ''TOTAL'', SUM(ITD), NULL, NULL, SUM([TOTAL]) 
				 FROM (SELECT ''ITD'' dtdate, E.sCategory, E.iCategory hKey, dbo.Fn_dAmount(E.sNotes, E.items, E.dAmount) dAmount 
						 FROM dbo.Fn_DataStore(''<@P_InceptionDate>'', ''<@ToDate>'', 0, ''<@P_SearchStr>'', 0) E 
                WHERE 1 = CASE WHEN LEN(ISNULL(''<@P_Category>'','''')) > 0 AND CHARINDEX(E.sCategory, ISNULL(''<@P_Category>'','''')) > 0 THEN 1 WHEN LEN(ISNULL(''<@P_Category>'','''')) = 0 THEN 1 ELSE 0 END) AS X
			    PIVOT (SUM(x.dAmount) FOR X.dtDate in (ITD, [TOTAL])) AS pvt
	            ORDER BY sort,category'

	  SET @sql = REPLACE(@sql, '<@P_InceptionDate>', @P_InceptionDate)
	  SET @sql = REPLACE(@sql, '<@ToDate>', @ToDate)
	  SET @sql = REPLACE(@sql, '<@P_Category>', @P_Category)
	  SET @sql = REPLACE(@sql, '<@P_SearchStr>', @P_SearchStr)
	
	  INSERT INTO #Temp_AllInOne EXEC (@sql)
	  INSERT INTO #ITD(iCategory, ITD) SELECT iCategory, dAmount FROM #Temp_AllInOne
	  
	  SELECT CASE I.iCategory WHEN 999 THEN 2 ELSE 1 END Sort,
			 I.iCategory,
			 CASE I.iCategory WHEN 999 THEN 'TOTAL' ELSE C.sCategory END Category,
			 ISNULL(M.MTD, 0) MTD,
			 ISNULL(Y.YTD, 0) YTD,
			 ISNULL(I.ITD, 0) ITD,
			 NULL [TOTAL]
	    FROM #ITD I 
			 LEFT JOIN #YTD Y ON I.iCategory = Y.iCategory
			 LEFT JOIN #MTD M ON M.iCategory = Y.iCategory
			 LEFT JOIN tbl_CategoryList C ON C.hKey = I.iCategory 
  END
GO
/*############################################################################################################################################################*/