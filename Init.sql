USE [dev]
GO
/****** Object:  StoredProcedure [DQ].[Run]    Script Date: 25/12/2025 4:54:10 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- init script
--
-- Create schema if not exists
IF (NOT EXISTS (SELECT * FROM sys.schemas WHERE name = 'DQ'))
BEGIN
    EXEC('CREATE SCHEMA [DQ] AUTHORIZATION [dbo]');
END

-- Create DQ.SessionResult table
-- This table stores the results of the tests runs
drop table if exists DQ.SessionResult;
CREATE TABLE DQ.SessionResult
(RunDate Datetime2(7) NOT NULL,
TestID int NOT NULL,
IsValidTest bit NOT NULL,						-- 1: The test is valid and has been performed
												-- 0: The test is not valid (missing parameters, etc) and has not been performed
ReasonIsNotValidTest nvarchar(255),				-- The reason the test has been invalidated
RecordsTotal int NOT NULL,						-- The total number of records in the table referenced by the test
RecordsPassed int NOT NULL,						-- The total number of records passing the test (TestType = R) or 1 if the test condition has been satisfied
IsPassedTest bit NOT NULL,						-- 1: The test is a pass
												-- 0: The test is a fail
SQLStatement nvarchar(4000) NULL);				-- The executed SQL statement (the SQL the test consists of)
CREATE CLUSTERED INDEX [PK_DQSessionResult] ON [DQ].[SessionResult] ([RunDate] ASC, [TestID] ASC);

-- Create DQ.TestMeta table
-- This table contains pre-coded tests, and metadata allowing to validate them when used in the DQ.Test table
drop table if exists DQ.TestMeta;
CREATE TABLE DQ.TestMeta
(TestName nvarchar(255) NOT NULL,
TestDescription nvarchar(255) NULL,
IsColumnNameRequired bit NOT NULL,
IsValue1Required bit NOT NULL,
IsValue2Required bit NOT NULL,
IsValue1AColumnName bit NOT NULL,
IsValue1AFullyQualifiedTableName bit NOT NULL,
RecordOrCondition char(1) NULL);				-- 'R': The test is on records (aka if table total records = Records Passed the test is a pass)
												-- 'C': The test asserts a condition. That is independent from the number of records passing the test
CREATE CLUSTERED INDEX [PK_DQTestMeta] ON [DQ].[TestMeta] ([TestName] ASC);

INSERT INTO DQ.TestMeta (TestName, TestDescription, IsColumnNameRequired, IsValue1Required, IsValue2Required, IsValue1AColumnName, IsValue1AFullyQualifiedTableName, RecordOrCondition)
VALUES
-- column tests
('ColumnEqualToValue', 'ColumnName must be equal to Value1', 1, 1, 0, 0, 0, 'R')
,('ColumnNotEqualToValue', 'ColumnName must not be equal to Value1', 1, 1, 0, 0, 0, 'R')
,('ColumnLessThanValue', 'ColumnName must have a value less than Value1', 1, 1, 0, 0, 0, 'R')
,('ColumnGreaterThanValue', 'ColumnName must have a value greater than Value1', 1, 1, 0, 0, 0, 'R')
,('ColumnEqualToColumn', 'ColumnName must be equal to the column name in Value1', 1, 1, 0, 0, 0, 'R')
,('ColumnNotEqualToColumn', 'ColumnName must not be equal to the column name in Value1', 1, 1, 0, 1, 0, 'R')
,('ColumnLessThanColumn', 'ColumnName must have a value less than the column name in Value1', 1, 1, 0, 1, 0, 'R')
,('ColumnGreaterThanColumn', 'ColumnName must have a value greater than the column name in Value1', 1, 1, 0, 1, 0, 'R')
,('ColumnNULL', 'ColumnName must be NULL', 1, 0, 0, 0, 0, 'R')
,('ColumnNotNULL', 'ColumnName must not be NULL', 1, 0, 0, 0, 0, 'R')
,('ColumnEmpty', 'ColumnName must be empty', 1, 0, 0, 0, 0, 'R')
,('ColumnNotEmpty', 'ColumnName must not be empty', 1, 0, 0, 0, 0, 'R')
,('ColumnNULLOrEmpty', 'ColumnName must be NULL or empty', 1, 0, 0, 0, 0, 'R')
,('ColumnNotNULLOrEmpty', 'ColumnName must not be NULL or empty', 1, 0, 0, 0, 0, 'R')
,('ColumnWithinRange', 'ColumnName must be between Value1 and Value2', 1, 1, 1, 0, 0, 'R')
,('ColumnOutsideRange', 'ColumnName must not be between Value1 and Value2', 1, 1, 1, 0, 0, 'R')
,('ColumnIsDate', 'ColumnName contents must be a valid (convertible) date value', 1, 0, 0, 0, 0, 'R')
,('ColumnIsInt', 'The column must have a valid int value', 1, 0, 0, 0, 0, 'R')
,('ColumnIsNumeric', 'The column must have a valid numeric value (NB: in this case int is valid as well)', 1, 0, 0, 0, 0, 'R')
,('ColumnWithinSet', 'The column must have one of the values of the Value1 (SetList.ListID)', 1, 1, 0, 0, 0, 'R')
,('ColumnNotWithinSet', 'The column must not have one of the values of the Value1 (SetList.ListID)', 1, 1, 0, 0, 0, 'R')
,('ColumnLengthMin', 'The column must be at least Value1 chars long', 1, 1, 0, 0, 0, 'R')
,('ColumnLengthMax', 'The column must be no more than Value1 chars long', 1, 1, 0, 0, 0, 'R')
,('ColumnLengthEqual', 'The column must be Value1 long', 1, 1, 0, 0, 0, 'R')
,('ColumnLengthWithinRange', 'ColumnName length must be between Value1 and Value2', 1, 1, 1, 0, 0, 'R')
,('ColumnInParentTable', 'The column value must exist in the table described by Value1 (DBName.SchemaName.TableName). When NOT NULL, the where clause (column(s)) applies to the parent table.', 1, 1, 0, 0, 1, 'R')
,('ColumnIsLike', 'The column value must be like the like expression in Value1', 1, 1, 0, 0, 0, 'R')
-- set tests...
,('ColumnDistinctValuesGreaterThanValue', 'ColumnName must have more than Value1 distinct values', 1, 1, 0, 0, 0, 'C')
,('ColumnDistinctValuesLessThanValue', 'ColumnName must not have more than Value1 distinct values', 1, 1, 0, 0, 0, 'C')
,('ColumnDistinctValuesEqualToValue', 'ColumnName must have Value1 distinct values', 1, 1, 0, 0, 0, 'C')
,('ColumnDistinctValuesEqualToRecordsNumber', 'ColumnName must have as many distinct values as of total records in the table', 1, 0, 0, 0, 0, 'C')
,('ColumnAverageWithinRange', 'The average value of ColumnName values must be between Value1 and Value2', 1, 1, 1, 0, 0, 'C')
,('ColumnSumWithinRange', 'The sum of the values of ColumnName must be between Value1 and Value2', 1, 1, 1, 0, 0, 'C')
,('RecordsNumberGreaterThanValue', 'The number of the records of TableName must be more than Value1', 0, 1, 0, 0, 0, 'C')
,('RecordsNumberLessThanValue', 'The number of the records of TableName must be less than Value1', 0, 1, 0, 0, 0, 'C')
,('RecordsNumberEqualToValue', 'The number of the records of TableName must be equal than Value1', 0, 1, 0, 0, 0, 'C')
,('RecordsNumberGreaterThanLast', 'The number of the records of TableName must be more than during previous run', 0, 0, 0, 0, 0, 'C')
,('RecordsNumberGreaterOrEqualThanLast', 'The number of the records of TableName must be more than during previous run', 0, 0, 0, 0, 0, 'C')
,('RecordsNumberLessThanLast', 'The number of the records of TableName must be less than during previous run', 0, 0, 0, 0, 0, 'C')
,('RecordsNumberLessOrEqualThanLast', 'The number of the records of TableName must be less than during previous run', 0, 0, 0, 0, 0, 'C')
,('RecordsNumberEqualToLast', 'The number of the records of TableName must be equal to the one observed during previous run', 0, 0, 0, 0, 0, 'C')
,('RecordsNumberGreaterThanTable', 'The number of the records of TableName must be more than the table described by Value1 (DBName.SchemaName.TableName)', 0, 1, 0, 0, 1, 'C')
,('RecordsNumberGreaterOrEqualThanTable', 'The number of the records of TableName must be more than the table described by Value1 (DBName.SchemaName.TableName)', 0, 1, 0, 0, 1, 'C')
,('RecordsNumberLessThanTable', 'The number of the records of TableName must be less than the table described by Value1 (DBName.SchemaName.TableName)', 0, 1, 0, 0, 1, 'C')
,('RecordsNumberLessOrEqualThanTable', 'The number of the records of TableName must be less than the table described by Value1 (DBName.SchemaName.TableName)', 0, 1, 0, 0, 1, 'C')
,('RecordsNumberEqualToTable', 'The number of the records of TableName must be equal than in the table described by Value1 (DBName.SchemaName.TableName)', 0, 1, 0, 0, 1, 'C')
,('RecordsUniqueGroupBy', 'The records of the table must be unique when grouped by the list of columns in Value1', 0, 1, 0, 0, 0, 'C') 
,('SQLEXPRESSION', 'The test consists of a dynamic sql string: select count(*) from Databasename.TableName where [WhereClause]. Value1 gives a custom name to the test.', 0, 1, 0, 0, 0, NULL);

-- Create the DQ.Test table
-- This is the table containing all the tests a user wants to perform
drop table if exists DQ.Test;
CREATE TABLE DQ.Test
(TestID int IDENTITY(1,1) NOT NULL,
DatabaseName nvarchar(255) NOT NULL,
TableName nvarchar(255) NOT NULL,		-- schemaname.tablename
ColumnName nvarchar(255) NULL,
TestName nvarchar(255) NOT NULL,		-- must be one of DQ.TestMeta
Value1 nvarchar(255) NULL,
Value2 nvarchar(255) NULL,
WhereClause nvarchar(1024) NULL,		-- reserved for future use
SQLExpression nvarchar(4000) NULL,		-- when the testname is 'SQLEXPRESSION' (user custom test)
										-- The SQL statement must return a @RecordsCount column (e.g. 'Select @RecordsCount from ...) 
RecordOrCondition char(1) NULL			-- For SQL Expression only.
										-- It tells the engine if the result of the 'select @RecordsCount...' is a number of records satisfying the test ('R')
										-- or if it's a 1/0 result ('C').  
);
CREATE CLUSTERED INDEX [PK_DQTest] ON [DQ].[Test] ([TestID] ASC);

-- Create the DQ.SetList table
-- This table can be addressed by the engine for certain tests consisting of testing a 'user'/'static'/... foreign key 
-- It contains user populated records 
drop table if exists DQ.SetList;
CREATE TABLE DQ.SetList
(ListName nvarchar(255) NOT NULL,
Ord int NoT NULL,
[Value] nvarchar(255) NULL);
CREATE CLUSTERED INDEX [PK_DQSetList] ON [DQ].[SetList] ([ListName] ASC, [Ord] ASC);
-- e.g.
-- insert into DQ.SetList (ListName, Ord, [Value]) VALUES ('Gender', 1, 'Female');
-- insert into DQ.SetList (ListName, Ord, [Value]) VALUES ('Gender', 2, 'Male');
-- Some tests allow to check a column value is into the 'Gender' list...
go

;CREATE OR ALTER proc [DQ].[Run]  
@DatabaseName nvarchar(255)=NULL,
@TableName nvarchar(255)=NULL	-- schemaname.tablename
as
-- Copyright (c) 2025 Dominique Beneteau (dombeneteau@yahoo.com)
-- Utilisation under the MIT License

-- When		Who			What
-- 20251222	dom			refactoring condition tests + SQL + testing
-- 20251221	dom			end testing column based tests
-- 20251220	dom			testing column based tests
-- 20251219	dom			full refactoring pre-github - Column based tests
-- 

begin
	declare @message nvarchar(255);
	declare @Now datetime2(7) = getdate();
	declare @TestID int, @IsValid bit;
	declare  @TestName nvarchar(255),
			@TestDescription nvarchar(255),
			@IsColumnNameRequired bit,
			@IsValue1Required bit,
			@IsValue2Required bit,
			@IsValue1AColumnName bit,
			@IsValue1AFullyQualifiedTableName bit;
	declare @TestDatabaseName nvarchar(255),
			@TestTableName nvarchar(255),	-- schemaname.tablename
			@TestColumnName nvarchar(255),
			@Value1 nvarchar(255),
			@Value2 nvarchar(255),
			@WhereClause nvarchar(1024);
	declare @FullyQualifiedTableName nvarchar(255);
	declare @RecordsCount int, @TableRecordsCount int;
	declare @SQL nvarchar(4000), @ParamDef nvarchar(255);
	declare @SingleQuote nvarchar(2) = '''';
	declare @RecordOrCondition char(1), @SQLExpression nvarchar(4000);

	-- Check databasename
	if @DatabaseName IS NOT NULL
	begin
		if DB_ID(@DatabaseName) IS NULL
		begin
			set @Message = 'Database ' + @DatabaseName + ' does not exist or offline or access denied, process aborted.';
			throw 50000, @Message, 1
			return -1
		end
	end

	-- Check Tablename
	if @TableName IS NOT NULL
	begin
		if object_id(@TableName) IS NULL
		begin
			set @Message = 'Table ' + @TableName + ' does not exist or access denied, process aborted.';
			throw 50000, @Message, 1
			return -1
		end
	end
		
	-- Build tests table
	drop table if exists #test
	select	TestID, 
			convert(bit, 0) as Checked,
			convert(nvarchar(255), NULL) as ReasonIsNotValidTest
	into #test
	from DQ.Test 
	where DatabaseName = ISNULL(@DatabaseName, DatabaseName)
	and TableName = ISNULL(@TableName, TableName)

	-- Check tests validity
	while exists (select top 1 1 from #test where Checked = 0)
	begin
		select top 1 @TestID = TestID
		from #test
		where Checked = 0

		set @IsValid = 1;

		select  @TestDatabaseName = DatabaseName,
				@TestTableName = TableName,
				@TestColumnName = ColumnName,
				@TestName = TestName,
				@Value1 = Value1,
				@Value2 = Value2,
				@WhereClause = WhereClause,
				@RecordOrCondition = RecordOrCondition,
				@SQLExpression = SQLExpression
		from DQ.Test 
		where TestID = @TestID

		set @FullyQualifiedTableName = @TestDatabaseName + '.' + @TestTableName;

		-- Check testname exists

		if @TestName NOT IN (select TestName from DQ.TestMeta)
		begin
			set @Message = 'Invalid TestName.';
			set @IsValid = 0;
		end

		-- Check DB exists

		if @IsValid = 1
		begin
			if DB_ID(@TestDatabaseName) IS NULL
			begin
				set @Message = 'Database ' + ISNULL(@TestDatabaseName, '?') + ' does not exist or offline or access denied.';
				set @IsValid = 0;
			end
		end

		-- Check table exists

		if @IsValid = 1
		begin
			if object_id(@TestTableName) IS NULL
			begin
				set @Message = 'Table ' + ISNULL(@TestTableName, '?') + ' does not exist or access denied.';
				set @IsValid = 0;
			end
		end

		-- Check RecordOrCondition exists for SQLEXPRESSION test

		if @IsValid = 1
		begin
			if @TestName = 'SQLEXPRESSION' and ISNULL(@RecordOrCondition, '') NOT IN ('R', 'C')
			begin
				set @Message = 'A SQLEXPRESSION test must have its RecordOrCondition populated with either R or C.';
				set @IsValid = 0;
			end
		end

		-- Check SQLExpression exists for SQLEXPRESSION test

		if @IsValid = 1
		begin
			if @TestName = 'SQLEXPRESSION' and ISNULL(@SQLExpression, '?') = '?'
			begin
				set @Message = 'A SQLEXPRESSION test must have a SQLExpression statement.';
				set @IsValid = 0;
			end
		end

		-- Test meta 

		if @IsValid = 1
		begin
			select	@IsColumnNameRequired = IsColumnNameRequired,
					@IsValue1Required = IsValue1Required,
					@IsValue2Required = IsValue2Required,
					@IsValue1AColumnName = IsValue1AColumnName,
					@IsValue1AFullyQualifiedTableName = IsValue1AFullyQualifiedTableName
			from DQ.TestMeta
			where TestName = @TestName;

--			select 'meta read', @IsColumnNameRequired, @FullyQualifiedTableName, @TestColumnName; 

			if @IsColumnNameRequired = 1
			begin
				-- Check column exists if required
				if COL_LENGTH(@FullyQualifiedTableName, @TestColumnName) IS NULL
				begin
					set @Message = 'Column ' + ISNULL(@TestColumnName, '?') + ' does not exist or access denied.';
					set @IsValid = 0;
				end
			end
		end
		
		if @IsValid = 1
		begin
			if @IsValue1Required = 1
			begin
				if ltrim(rtrim(@Value1)) < ' '
				begin
					set @Message = 'Value1 must be provided.';
					set @IsValid = 0;
				end
			end
		end
		
		if @IsValid = 1
		begin
			if @IsValue2Required = 1
			begin
				if ltrim(rtrim(@Value2)) < ' '
				begin
					set @Message = 'Value2 must be provided.';
					set @IsValid = 0;
				end
			end
		end
				
		if @IsValid = 1
		begin
			if @IsValue1AColumnName = 1
			begin
				-- Check column exists if required
				if COL_LENGTH(@FullyQualifiedTableName, @Value1) IS NULL
				begin
					set @Message = 'Column ' + @Value1 + ' does not exist or access denied.';
					set @IsValid = 0;
				end
			end
		end

		if @IsValid = 1
		begin
			if @IsValue1AFullyQualifiedTableName = 1
			begin
				-- Check column exists if required
				if OBJECT_ID(@Value1) IS NULL
				begin
					set @Message = @Value1 + ' must be a fully qualified table name.';
					set @IsValid = 0;
				end
			end
		end

		-- test where clause
		if @IsValid = 1
		begin
			if @WhereClause IS NOT NULL
			begin
				SET @SQL = N'SELECT @TableRecordsCount = count(*) from ' +@FullyQualifiedTableName + ' where ' + @WhereClause;
				SET @ParamDef = N'@TableRecordsCount int OUTPUT';
				begin try
					EXECUTE sp_executesql @SQL, @ParamDef, @TableRecordsCount = @TableRecordsCount OUTPUT;
				end try
				begin catch
					set @IsValid = 0;
					set @Message = 'Invalid WhereClause';
				end CATCH
			end
		end

		-- epilog

		if @IsValid = 0
		begin
			insert into DQ.SessionResult (RunDate, TestID, IsValidTest, ReasonIsNotValidTest, RecordsTotal, RecordsPassed, IsPassedTest)
			Values (@Now, @TestID, 0, @Message, 0, 0, 0);

			delete #test
			where TestID = @TestID;	-- we don't want this test to be part of the session
		end
		else
		begin
			update #test
			set Checked = 1
			where TestID = @TestID;	-- we don't want this test to be part of the session
		end
	end

	-- In #test we now have the full list of tests to run

	while exists (select top 1 1 from #test)
	begin
		set @IsValid = 1;

		select top 1 @TestID = TestID
		from #test

		select  @TestDatabaseName = DatabaseName,
				@TestTableName = TableName,
				@TestColumnName = ColumnName,
				@TestName = TestName,
				@Value1 = Value1,
				@Value2 = Value2,
				@WhereClause = WhereClause,
				@RecordOrCondition = RecordOrCondition,
				@SQLExpression = SQLExpression
		from DQ.Test 
		where TestID = @TestID

		if @TestName <> 'SQLEXPRESSION'
		begin
			select @RecordOrCondition = RecordOrCondition
			from DQ.TestMeta
			where TestName = @TestName
		end

		set @FullyQualifiedTableName = @TestDatabaseName + '.' + @TestTableName;

		-- records total

		SET @SQL = N'SELECT @TableRecordsCount = count(*) FROM ' +@FullyQualifiedTableName;
		SET @ParamDef = N'@TableRecordsCount int OUTPUT';
		begin try
			EXECUTE sp_executesql @SQL, @ParamDef, @TableRecordsCount = @TableRecordsCount OUTPUT;
		end try
		begin catch
			set @IsValid = 0;
			set @Message = 'Total count: ' + ERROR_MESSAGE();
			insert into DQ.SessionResult (RunDate, TestID, IsValidTest, ReasonIsNotValidTest, RecordsTotal, RecordsPassed, IsPassedTest, SQLStatement)
			Values (@Now, @TestID, 0, @message, 0, 0, 0, @SQL);

			delete #test
			where TestID = @TestID;	-- abort this test
		end CATCH

		-- records pass

		if @IsValid = 1
		begin

			SET @ParamDef = N'@RecordsCount int OUTPUT';
			set @SQL = NULL;

			if @TestName = 'ColumnEqualToValue'
				SET @SQL = N'SELECT @RecordsCount = count(*) FROM ' +@FullyQualifiedTableName + ' where ' + @TestColumnName + ' = ' + @SingleQuote + @Value1 + @SingleQuote;
			if @TestName = 'ColumnNotEqualToValue'
				SET @SQL = N'SELECT @RecordsCount = count(*) FROM ' +@FullyQualifiedTableName + ' where ' + @TestColumnName + ' <> ' + @SingleQuote + @Value1 + @SingleQuote;
			if @TestName = 'ColumnLessThanValue'
				SET @SQL = N'SELECT @RecordsCount = count(*) FROM ' +@FullyQualifiedTableName + ' where ' + @TestColumnName + ' < ' + @SingleQuote + @Value1 + @SingleQuote;
			if @TestName = 'ColumnGreaterThanValue'
				SET @SQL = N'SELECT @RecordsCount = count(*) FROM ' +@FullyQualifiedTableName + ' where ' + @TestColumnName + ' > ' + @SingleQuote + @Value1 + @SingleQuote;
			if @TestName = 'ColumnEqualToColumn'
				SET @SQL = N'SELECT @RecordsCount = count(*) FROM ' +@FullyQualifiedTableName + ' where ' + @TestColumnName + ' = ' + @Value1;
			if @TestName = 'ColumnNotEqualToColumn'
				SET @SQL = N'SELECT @RecordsCount = count(*) FROM ' +@FullyQualifiedTableName + ' where ' + @TestColumnName + ' <> ' + @Value1;
			if @TestName = 'ColumnLessThanColumn'
				SET @SQL = N'SELECT @RecordsCount = count(*) FROM ' +@FullyQualifiedTableName + ' where ' + @TestColumnName + ' < ' + @Value1;
			if @TestName = 'ColumnGreaterThanColumn'
				SET @SQL = N'SELECT @RecordsCount = count(*) FROM ' +@FullyQualifiedTableName + ' where ' + @TestColumnName + ' > ' + @Value1;
			if @TestName = 'ColumnNULL'
				SET @SQL = N'SELECT @RecordsCount = count(*) FROM ' +@FullyQualifiedTableName + ' where ' + @TestColumnName + ' IS NULL';
			if @TestName = 'ColumnNotNULL'
				SET @SQL = N'SELECT @RecordsCount = count(*) FROM ' +@FullyQualifiedTableName + ' where ' + @TestColumnName + ' IS NOT NULL';
			if @TestName = 'ColumnEmpty'
				SET @SQL = N'SELECT @RecordsCount = count(*) FROM ' +@FullyQualifiedTableName + ' where ' + @TestColumnName + ' = ' + @SingleQuote + @SingleQuote;
			if @TestName = 'ColumnNotEmpty'
				SET @SQL = N'SELECT @RecordsCount = count(*) FROM ' +@FullyQualifiedTableName + ' where ' + @TestColumnName + ' <> ' + @SingleQuote + @SingleQuote;
			if @TestName = 'ColumnNULLOrEmpty'
				SET @SQL = N'SELECT @RecordsCount = count(*) FROM ' +@FullyQualifiedTableName + ' where ' + @TestColumnName + ' IN (NULL, ' + @SingleQuote + @SingleQuote + ')';
			if @TestName = 'ColumnNotNULLOrEmpty'
				SET @SQL = N'SELECT @RecordsCount = count(*) FROM ' +@FullyQualifiedTableName + ' where ' + @TestColumnName + ' NOT IN (NULL, ' + @SingleQuote + @SingleQuote + ')';
			if @TestName = 'ColumnWithinRange'
				SET @SQL = N'SELECT @RecordsCount = count(*) FROM ' +@FullyQualifiedTableName + ' where ' + @TestColumnName + ' between ' + @SingleQuote + @Value1 + @SingleQuote + ' and ' + @SingleQuote + @Value2 + @SingleQuote;
			if @TestName = 'ColumnOutsideRange'
				SET @SQL = N'SELECT @RecordsCount = count(*) FROM ' +@FullyQualifiedTableName + ' where ' + @TestColumnName + ' NOT between ' + @SingleQuote + @Value1 + @SingleQuote + ' and ' + @SingleQuote + @Value2 + @SingleQuote;
			if @TestName = 'ColumnIsDate'
				SET @SQL = N'SELECT @RecordsCount = count(*) FROM ' +@FullyQualifiedTableName + ' where ISDATE(' + @TestColumnName + ') = 1';
			if @TestName = 'ColumnIsInt'
				SET @SQL = N';WITH s as (SELECT ' +  @TestColumnName + ' FROM ' +@FullyQualifiedTableName + ' where ISNUMERIC(' + @TestColumnName + ') = 1) 
								SELECT @RecordsCount = count(*) FROM s where ' + @TestColumnName + ' = convert(int, ' + @TestColumnName + ')';
			if @TestName = 'ColumnIsNumeric'
				SET @SQL = N'SELECT @RecordsCount = count(*) FROM ' +@FullyQualifiedTableName + ' where ISNUMERIC(' + @TestColumnName + ') = 1';
			if @TestName = 'ColumnWithinSet'
				SET @SQL = N'SELECT @RecordsCount = count(*) FROM ' +@FullyQualifiedTableName + ' where ' + @TestColumnName + ' IN (select [Value] from DQ.SetList where ListName = ' + @SingleQuote + @Value1 + @Singlequote + ')';
			if @TestName = 'ColumnNotWithinSet'
				SET @SQL = N'SELECT @RecordsCount = count(*) FROM ' +@FullyQualifiedTableName + ' where ' + @TestColumnName + ' NOT IN (select [Value] from DQ.SetList where ListName = ' + @SingleQuote + @Value1 + @Singlequote + ')';
			if @TestName = 'ColumnLengthMin'
				SET @SQL = N'SELECT @RecordsCount = count(*) FROM ' +@FullyQualifiedTableName + ' where len(convert(nvarchar(1048), ' + @TestColumnName + ')) >= ' + @Value1;
			if @TestName = 'ColumnLengthMax'
				SET @SQL = N'SELECT @RecordsCount = count(*) FROM ' +@FullyQualifiedTableName + ' where len(convert(nvarchar(1048), ' + @TestColumnName + ')) <= ' + @Value1;
			if @TestName = 'ColumnLengthEqual'
				SET @SQL = N'SELECT @RecordsCount = count(*) FROM ' +@FullyQualifiedTableName + ' where len(convert(nvarchar(1048), ' + @TestColumnName + ')) = ' + @Value1;
			if @TestName = 'ColumnLengthWithinRange'
				SET @SQL = N'SELECT @RecordsCount = count(*) FROM ' +@FullyQualifiedTableName + ' where len(convert(nvarchar(1048), ' + @TestColumnName + ')) between ' + @Value1 + ' and ' + @Value2;
			if @TestName = 'ColumnInParentTable'
				SET @SQL = N'SELECT @RecordsCount = count(*) FROM ' +@FullyQualifiedTableName + ' where ' + @TestColumnName + ' IN (select ' + @Value2 + ' from ' + @Value1 + ')';
			if @TestName = 'ColumnIsLike'
				SET @SQL = N'SELECT @RecordsCount = count(*) FROM ' +@FullyQualifiedTableName + ' where ' + @TestColumnName + ' like ' + @SingleQuote + @Value1 + @SingleQuote;

			if @TestName = 'ColumnDistinctValuesGreaterThanValue'
				SET @SQL = N';with s as (select ' + @TestColumnName + ' FROM ' + @FullyQualifiedTableName + ' group by ' + @TestColumnName + ') 
								SELECT @RecordsCount = case when (select count(*) FROM s) > ' + @Value1 + ' then 1 else 0 end';
			if @TestName = 'ColumnDistinctValuesLessThanValue'
				SET @SQL = N';with s as (select ' + @TestColumnName + ' FROM ' + @FullyQualifiedTableName + ' group by ' + @TestColumnName + ') 
								SELECT @RecordsCount = case when (select count(*) FROM s) < ' + @Value1 + ' then 1 else 0 end';
			if @TestName = 'ColumnDistinctValuesEqualToValue'
				SET @SQL = N';with s as (select ' + @TestColumnName + ' FROM ' + @FullyQualifiedTableName + ' group by ' + @TestColumnName + ') 
								SELECT @RecordsCount = case when (select count(*) FROM s) = ' + @Value1 + ' then 1 else 0 end';
			if @TestName = 'ColumnDistinctValuesEqualToRecordsNumber'
				SET @SQL = N';with s as (select ' + @TestColumnName + ' FROM ' + @FullyQualifiedTableName + ' group by ' + @TestColumnName + ') 
								SELECT @RecordsCount = case when (select count(*) FROM s) = (select count(*) from ' + @FullyQualifiedTableName + ') then 1 else 0 end';
			if @TestName = 'ColumnAverageWithinRange'
				SET @SQL = N';with s as (select AVG(' + @TestColumnName + ') as a FROM ' + @FullyQualifiedTableName + ') 
								SELECT @RecordsCount = case when (select a FROM s) between ' + @Value1 + ' and ' + @Value2 + ' then 1 else 0 end';
			if @TestName = 'ColumnSumWithinRange'
				SET @SQL = N';with s as (select SUM(' + @TestColumnName + ') as a FROM ' + @FullyQualifiedTableName + ') 
								SELECT @RecordsCount = case when (select a FROM s) between ' + @Value1 + ' and ' + @Value2 + ' then 1 else 0 end';
			if @TestName = 'RecordsNumberGreaterThanValue'
				SET @SQL = N'SELECT @RecordsCount = case when (select count(*) FROM ' + @FullyQualifiedTableName + ') > ' + @Value1 + ' then 1 else 0 end';
			if @TestName = 'RecordsNumberLessThanValue'
				SET @SQL = N'SELECT @RecordsCount = case when (select count(*) FROM ' + @FullyQualifiedTableName + ') < ' + @Value1 + ' then 1 else 0 end';
			if @TestName = 'RecordsNumberEqualToValue'
				SET @SQL = N'SELECT @RecordsCount = case when (select count(*) FROM ' + @FullyQualifiedTableName + ') = ' + @Value1 + ' then 1 else 0 end';
			if @TestName = 'RecordsNumberGreaterThanLast'
				SET @SQL = N'declare @lastCount int=0;
							SELECT top 1 @LastCount = sr.RecordsTotal
							from DQ.SessionResult sr
							inner join DQ.Test t on t.TestID = sr.TestID 
							where t.DatabaseName + ' + @SingleQuote + '.' + @SingleQuote + ' + t.TableName = ' + @SingleQuote + @FullyQualifiedTableName + @SingleQuote + ' 
							order by sr.RunDate Desc;
							SELECT @RecordsCount = case when ' + convert(nvarchar(32), @TableRecordsCount) + ' > @lastCount then 1 else 0 end';
			if @TestName = 'RecordsNumberGreaterOrEqualThanLast'
				SET @SQL = N'declare @lastCount int=0;
							SELECT top 1 @LastCount = sr.RecordsTotal
							from DQ.SessionResult sr
							inner join DQ.Test t on t.TestID = sr.TestID 
							where t.DatabaseName + ' + @SingleQuote + '.' + @SingleQuote + ' + t.TableName = ' + @SingleQuote + @FullyQualifiedTableName + @SingleQuote + ' 
							order by sr.RunDate Desc;
							SELECT @RecordsCount = case when ' + convert(nvarchar(32), @TableRecordsCount) + ' >= @lastCount then 1 else 0 end';
			if @TestName = 'RecordsNumberLessThanLast'
				SET @SQL = N'declare @lastCount int=0;
							SELECT top 1 @LastCount = sr.RecordsTotal
							from DQ.SessionResult sr
							inner join DQ.Test t on t.TestID = sr.TestID 
							where t.DatabaseName + ' + @SingleQuote + '.' + @SingleQuote + ' + t.TableName = ' + @SingleQuote + @FullyQualifiedTableName + @SingleQuote + ' 
							order by sr.RunDate Desc;
							SELECT @RecordsCount = case when ' + convert(nvarchar(32), @TableRecordsCount) + ' < @lastCount then 1 else 0 end';
			if @TestName = 'RecordsNumberLessOrEqualThanLast'
				SET @SQL = N'declare @lastCount int=0;
							SELECT top 1 @LastCount = sr.RecordsTotal
							from DQ.SessionResult sr
							inner join DQ.Test t on t.TestID = sr.TestID 
							where t.DatabaseName + ' + @SingleQuote + '.' + @SingleQuote + ' + t.TableName = ' + @SingleQuote + @FullyQualifiedTableName + @SingleQuote + ' 
							order by sr.RunDate Desc;
							SELECT @RecordsCount = case when ' + convert(nvarchar(32), @TableRecordsCount) + ' <= @lastCount then 1 else 0 end';
			if @TestName = 'RecordsNumberEqualToLast'
				SET @SQL = N'declare @lastCount int=0;
							SELECT top 1 @LastCount = sr.RecordsTotal
							from DQ.SessionResult sr
							inner join DQ.Test t on t.TestID = sr.TestID 
							where t.DatabaseName + ' + @SingleQuote + '.' + @SingleQuote + ' + t.TableName = ' + @SingleQuote + @FullyQualifiedTableName + @SingleQuote + ' 
							order by sr.RunDate Desc;
							SELECT @RecordsCount = case when ' + convert(nvarchar(32), @TableRecordsCount) + ' = @lastCount then 1 else 0 end';
			if @TestName = 'RecordsNumberGreaterThanTable'
				SET @SQL = N'declare @TableCount int=0;
							SELECT @TableCount = count(*) from ' + @Value1 + '; 
							SELECT @RecordsCount = case when ' + convert(nvarchar(32), @TableRecordsCount) + ' > @TableCount then 1 else 0 end';
			if @TestName = 'RecordsNumberGreaterOrEqualThanTable'
				SET @SQL = N'declare @TableCount int=0;
							SELECT @TableCount = count(*) from ' + @Value1 + '; 
							SELECT @RecordsCount = case when ' + convert(nvarchar(32), @TableRecordsCount) + ' >= @TableCount then 1 else 0 end';
			if @TestName = 'RecordsNumberLessThanTable'
				SET @SQL = N'declare @TableCount int=0;
							SELECT @TableCount = count(*) from ' + @Value1 + '; 
							SELECT @RecordsCount = case when ' + convert(nvarchar(32), @TableRecordsCount) + ' < @TableCount then 1 else 0 end';
			if @TestName = 'RecordsNumberLessOrEqualThanTable'
				SET @SQL = N'declare @TableCount int=0;
							SELECT @TableCount = count(*) from ' + @Value1 + '; 
							SELECT @RecordsCount = case when ' + convert(nvarchar(32), @TableRecordsCount) + ' <= @TableCount then 1 else 0 end';
			if @TestName = 'RecordsNumberEqualToTable'
				SET @SQL = N'declare @TableCount int=0;
							SELECT @TableCount = count(*) from ' + @Value1 + '; 
							SELECT @RecordsCount = case when ' + convert(nvarchar(32), @TableRecordsCount) + ' = @TableCount then 1 else 0 end';
			if @TestName = 'RecordsUniqueGroupBy'
				SET @SQL = N';with s as (select ' + @Value1 + ' from ' + @FullyQualifiedTableName + ' group by ' + @Value1 + ' having count(*) > 1)  
							SELECT @RecordsCount = case when not exists (select 1 from s) then 1 else 0 end';

			if @TestName = 'SQLEXPRESSION'
				SET @SQL = @SQLExpression;

			if @SQL IS NOT NULL
			begin
				begin try
--					select @SQL;

					EXECUTE sp_executesql @SQL, @ParamDef, @RecordsCount = @RecordsCount OUTPUT;
				end try
				begin catch
					set @IsValid = 0;
					set @message = 'Test run: ' + ERROR_MESSAGE();
				end CATCH

				if @IsValid = 0
				begin
					insert into DQ.SessionResult (RunDate, TestID, IsValidTest, ReasonIsNotValidTest, RecordsTotal, RecordsPassed, IsPassedTest, SQLStatement)
					Values (@Now, @TestID, 0, @message, 0, 0, 0, @SQL);
				end
				else
				begin
					If @RecordOrCondition = 'R'
					begin
						-- tests on columns
						-- These tests count the number of records satisfying the test. 
						-- If this number equals the total number of records, the test passed otherwise it failed. 
						insert into DQ.SessionResult (RunDate, TestID, IsValidTest, ReasonIsNotValidTest, RecordsTotal, RecordsPassed, IsPassedTest, SQLStatement)
						Values (@Now, @TestID, 1, NULL, @TableRecordsCount, @RecordsCount, case when @TableRecordsCount = @RecordsCount then 1 else 0 end, @SQL);
					end
					else
					begin
						-- tests of condition
						-- These tests are assessing conditions without necessarily counting a number of records satisfying these.
						-- These conditions can be based on groups, hence the number of records would become a number of sets (groups), etc.
						-- The RecordsPassed column value is irrelevant here. 1 means pass, 0 means fail.
						insert into DQ.SessionResult (RunDate, TestID, IsValidTest, ReasonIsNotValidTest, RecordsTotal, RecordsPassed, IsPassedTest, SQLStatement)
						Values (@Now, @TestID, 1, NULL, @TableRecordsCount, @RecordsCount, @RecordsCount, @SQL);
					end
				end
			end

			delete #test
			where TestID = @TestID;

		end
	end
end

go

