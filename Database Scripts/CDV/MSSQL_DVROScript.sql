/*
------------------------------------------------------------------
Copyright:   InferMed Ltd. 2001. All Rights Reserved
File:        MACRO_DVROTRIGGER (DVROScript.sql)
Author:      Richard Weare
Purpose:     Script to create a trigger that fills 'Response Only' 
             dataviews when DataItemResponse is inserted into or 
             updated.
------------------------------------------------------------------
Revisions	Mo Morris 23/4/2002, Changes around the option to use category codes or values.
		Changes stemming from the typing of dataview question columns.
            Mo Morris 30/4/2002, Changes around the handling of NULL ResponseValues
            and NULL ValueCodes
            Mo Morris 27/6/2002, Changes around Date Response data being placed in DateTime fields
            Mo Morris (and ASH) 29/8/2002 changes around the handling of RQGs
		TA 11/02/2004: changed tablename variable to 255 length
		Mo Morris 1/12/2004 Bug 2446 - Adding Special Values facilities to CDV
		MLM 20/04/05: Modified for study-specific CDV.
		Mo Morris 3/11/2005 COD0030 - Changes around the new Thesaurus Data Item Type
		TA 03/02/06 - CDB 2679 - Upped size of varchar2 respone datatype to 500 to allow for single quotes to be doubled
					- QUOTENAME function now replace by REPLACE function - quotename was losing characters
		Mo Morris 25/10/2006 - Bug 2824 - Make MACRO Create Data Views Module comply with Partial Dates.
			When formating a date response the Partial Dates flag (DataItemCase) is now checked.
			When the flag = 1, it is a partial date question that will be written to the Data View tables as a string.
		Mo Morris 6/7/2007 - Bug 2936 - Make the Trigger correctly handle missing response data with a status of Success (0)
------------------------------------------------------------------
*/

--IF EXISTS (SELECT name FROM sysobjects WHERE name = 'Macro_DVROtrigger' AND type = 'TR')
--    DROP TRIGGER Macro_DVROtrigger
--GO

CREATE TRIGGER dbo.Macro_DVROtrigger 
ON             dbo.DataItemResponse
FOR            INSERT, UPDATE
AS
IF UPDATE(ResponseValue)
BEGIN            
  SET NOCOUNT ON

  DECLARE        @TableName          VARCHAR(255),
                 @ColumnName         VARCHAR(30),
                 @ColumnType         INT,
                 @CurrentRows        VARCHAR(1),
                 @SeparateVisits     INT,
                 @OutputCatValues    INT,
                 @SVMissing          VARCHAR(2),
                 @SVUnobtainable     VARCHAR(2),
                 @SVNotApplicable    VARCHAR(2),
                 @strSQL             VARCHAR(1024),
                 @err                INT,
                 @rows               INT,
                 @Response           VARCHAR(512),
                 @DateFS             VARCHAR(30),
                 @DateResp           VARCHAR(30),
                 @D1                 INT,
                 @D2                 INT,
                 @D3                 INT,
                 @D4                 INT,
                 @D5                 INT,
                 @D6                 INT,
                 @P1                 VARCHAR(4),
                 @P2                 VARCHAR(4),
                 @P3                 VARCHAR(4),
                 @P4                 VARCHAR(4),
                 @P5                 VARCHAR(4),
                 @P6                 VARCHAR(4),
                 @PartialDateFlag    INT,
                 
                 @ClinicalTrialId    INT,
                 @VisitId            INT,
                 @CRFPageId          INT,
                 @TrialSite          VARCHAR(8),
                 @PersonId           INT,
                 @VisitCycleNumber   INT,
                 @CRFPageCycleNumber SMALLINT,
                 @ResponseValue      VARCHAR(255),
                 @ResponseStatus     SMALLINT,
                 @DataItemId         INT,
                 @ValueCode          VARCHAR(15),
                 @OWQGroupID         INT,
                 @RepeatNumber       INT
 

  IF NOT EXISTS (SELECT * FROM sysobjects WHERE id = object_id(N'[dbo].[DataViewDetails]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
  RETURN

  IF NOT EXISTS (SELECT * FROM sysobjects WHERE id = object_id(N'[dbo].[DataViewTables]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
  RETURN

--~~~~~~~~~~~~~~~~~~~~
--cursor wrapper
--~~~~~~~~~~~~~~~~~~~~

  DECLARE         Macro_DVROcursor CURSOR
  FOR             SELECT       ClinicalTrialId, VisitId, CRFPageId, TrialSite, PersonId,
                               VisitCycleNumber, CRFPageCycleNumber, ResponseValue, ResponseStatus, DataItemId, ValueCode, RepeatNumber
	          FROM         INSERTED

  OPEN            Macro_DVROcursor 
  FETCH NEXT FROM Macro_DVROcursor INTO @ClinicalTrialId, @VisitId, @CRFPageId, @TrialSite, @PersonId,
                               @VisitCycleNumber, @CRFPageCycleNumber, @ResponseValue, @ResponseStatus, @DataItemId, @ValueCode, @RepeatNumber
    WHILE @@FETCH_STATUS = 0
    BEGIN
--~~~~~~~~~~~~~~~~~~~
-- ATO 20/08/2002 To get OwnerQGroupID */
  SELECT          @OWQGroupID = OwnerQGroupID
  FROM            CRFElement
  WHERE           ClinicalTrialId = @ClinicalTrialId 
  AND             CRFPageId = @CRFPageId
  AND             DataItemId = @DataItemId

/*
  SELECT          @SeparateVisits = DataViewSeparateVisits,
		  @OutputCatValues = OutputCategoryValues,
                  @SVMissing = SpecialValueMissing,
                  @SVUnobtainable = SpecialValueUnobtainable,
                  @SVNotApplicable = SpecialValueNotApplicable
  FROM            DataViewDetails
  WHERE           ClinicalTrialId = @ClinicalTrialId
*/

  SELECT          @SeparateVisits = DataViewSeparateVisits 
  FROM            DataViewDetails
  WHERE           ClinicalTrialId = @ClinicalTrialId

  SELECT          @OutputCatValues = OutputCategoryValues 
  FROM            DataViewDetails
  WHERE           ClinicalTrialId = @ClinicalTrialId

  SELECT          @SVMissing = SpecialValueMissing
  FROM            DataViewDetails
  WHERE           ClinicalTrialId = @ClinicalTrialId

  SELECT          @SVUnobtainable  = SpecialValueUnobtainable
  FROM            DataViewDetails
  WHERE           ClinicalTrialId = @ClinicalTrialId

  SELECT          @SVNotApplicable = SpecialValueNotApplicable 
  FROM            DataViewDetails
  WHERE           ClinicalTrialId = @ClinicalTrialId

    
  IF              @SeparateVisits = 1
      BEGIN
          SELECT        @TableName = DataViewName
          FROM          DataViewTables
          WHERE         ClinicalTrialId = @ClinicalTrialId
          AND           VisitId = @VisitId
          AND           CRFPageId = @CRFPageId
          AND           QGroupID = @OWQGroupID
          AND           DataViewType = 'RO'
      END
  ELSE
      BEGIN	
        SELECT        @TableName = DataViewName
        FROM          DataViewTables
        WHERE         ClinicalTrialId = @ClinicalTrialId
        AND           CRFPageId = @CRFPageId
        AND           QGroupID = @OWQGroupID
        AND           DataViewType = 'RO'
      END

  IF @TableName IS NULL
      BEGIN
          CLOSE Macro_DVROcursor
          DEALLOCATE Macro_DVROcursor
          RETURN
      END
  ELSE
      BEGIN
          IF @OWQGroupID = 0
              BEGIN
                  SELECT @strSQL =           'UPDATE DataViewDetails SET DataViewTrigCalc = ('  
                  SELECT @strSQL = @strSQL + ' SELECT CASE COUNT(*) '
                  SELECT @strSQL = @strSQL + ' WHEN 0 THEN ' + char(39) + 'I' + char(39) + '' 
                  SELECT @strSQL = @strSQL + ' ELSE ' + char(39) + 'U' + char(39) + '' 
                  SELECT @strSQL = @strSQL + ' END'
                  SELECT @strSQL = @strSQL + ' FROM [' + @TableName + ']'
                  SELECT @strSQL = @strSQL + ' WHERE ClinicalTrialId = ' + CONVERT(VARCHAR(20),@ClinicalTrialId) + ''
                  SELECT @strSQL = @strSQL + ' AND Site = ' + char(39) + @TrialSite + char(39) + ''
                  SELECT @strSQL = @strSQL + ' AND PersonId = ' + CONVERT(VARCHAR(20),@PersonId) + ''
                  SELECT @strSQL = @strSQL + ' AND VisitId = ' + CONVERT(VARCHAR(20),@VisitId) + ''
                  SELECT @strSQL = @strSQL + ' AND VisitCycleNumber = ' + CONVERT(VARCHAR(20),@VisitCycleNumber) + ''
                  SELECT @strSQL = @strSQL + ' AND CRFPageId = ' + CONVERT(VARCHAR(20),@CRFPageId) + ''
                  SELECT @strSQL = @strSQL + ' AND CRFPageCycleNumber = ' + CONVERT(VARCHAR(20),@CRFPageCycleNumber) + ')'
                  SELECT @strSQL = @strSQL + ' WHERE ClinicalTrialId = ' + CONVERT(VARCHAR(20),@ClinicalTrialId) + ''
              END
          ELSE
              BEGIN
                  SELECT @strSQL =           'UPDATE DataViewDetails SET DataViewTrigCalc = ('  
                  SELECT @strSQL = @strSQL + ' SELECT CASE COUNT(*) '
                  SELECT @strSQL = @strSQL + ' WHEN 0 THEN ' + char(39) + 'I' + char(39) + '' 
                  SELECT @strSQL = @strSQL + ' ELSE ' + char(39) + 'U' + char(39) + '' 
                  SELECT @strSQL = @strSQL + ' END'
                  SELECT @strSQL = @strSQL + ' FROM [' + @TableName + ']'
                  SELECT @strSQL = @strSQL + ' WHERE ClinicalTrialId = ' + CONVERT(VARCHAR(20),@ClinicalTrialId) + ''
                  SELECT @strSQL = @strSQL + ' AND Site = ' + char(39) + @TrialSite + char(39) + ''
                  SELECT @strSQL = @strSQL + ' AND PersonId = ' + CONVERT(VARCHAR(20),@PersonId) + ''
                  SELECT @strSQL = @strSQL + ' AND VisitId = ' + CONVERT(VARCHAR(20),@VisitId) + ''
                  SELECT @strSQL = @strSQL + ' AND VisitCycleNumber = ' + CONVERT(VARCHAR(20),@VisitCycleNumber) + ''
                  SELECT @strSQL = @strSQL + ' AND CRFPageId = ' + CONVERT(VARCHAR(20),@CRFPageId) + ''
                  SELECT @strSQL = @strSQL + ' AND CRFPageCycleNumber = ' + CONVERT(VARCHAR(20),@CRFPageCycleNumber) + ''
                  SELECT @strSQL = @strSQL + ' AND OwnerQGroupID = ' + CONVERT(VARCHAR(20),@OWQGroupID) + ''
                  SELECT @strSQL = @strSQL + ' AND RepeatNumber = ' + CONVERT(VARCHAR(20),@RepeatNumber) + ')'
                  SELECT @strSQL = @strSQL + ' WHERE ClinicalTrialId = ' + CONVERT(VARCHAR(20),@ClinicalTrialId) + ''
              END
      END

    EXEC (@strSQL)

    SELECT @err = @@ERROR, @rows = @@ROWCOUNT

    IF @err <> 0
      BEGIN
          CLOSE Macro_DVROcursor
          DEALLOCATE Macro_DVROcursor
          RETURN
      END
    IF @rows <> 1
      BEGIN
          CLOSE Macro_DVROcursor
          DEALLOCATE Macro_DVROcursor
          RETURN
      END

    SELECT        @ColumnName = DataItemCode
    FROM          DataItem
    WHERE         DataItem.ClinicalTrialId = @ClinicalTrialId
    AND           DataItem.DataItemId = @DataItemId

    SELECT        @ColumnType = DataType
    FROM          DataItem
    WHERE         DataItem.ClinicalTrialId = @ClinicalTrialId
    AND           DataItem.DataItemId = @DataItemId

    SELECT        @DateFS = DataItemFormat
    FROM          DataItem
    WHERE         DataItem.ClinicalTrialId = @ClinicalTrialId
    AND           DataItem.DataItemId = @DataItemId
    
    SELECT        @PartialDateFlag = DataItemCase
    FROM          DataItem
    WHERE         DataItem.ClinicalTrialId = @ClinicalTrialId
    AND           DataItem.DataItemId = @DataItemId
    
    If @PartialDateFlag IS NULL
		BEGIN
			SELECT @PartialDateFlag = 0
		END


    -- Format the ResponseValue ready for inserting into Data Views table DataViewName 
    -- its Text or Mulimedia or Thesaurus
    IF (@ColumnType = 0) OR (@ColumnType = 5) OR (@ColumnType = 8)
        BEGIN
            IF @ResponseValue IS NULL
                BEGIN
                    IF (@ResponseStatus = 10)
                        BEGIN
                            IF @SVMissing <> ''
                                SELECT @Response = @SVMissing
                            ELSE
                                SELECT @Response = 'null'
                        END
                    ELSE IF (@ResponseStatus = -8)
                        BEGIN
                            IF @SVNotApplicable <> ''
                                SELECT @Response = @SVNotApplicable
                            ELSE
                                SELECT @Response = 'null'
                        END
                    ELSE IF (@ResponseStatus = -5)
                        BEGIN
                            IF @SVUnobtainable <> ''
                                SELECT @Response = @SVUnobtainable
                            ELSE
                                SELECT @Response = 'null'
                        END
                    ELSE
                        SELECT @Response = 'null'
                END
            ELSE
	        SELECT @Response = '''' + replace(@responsevalue,'''','''''') + ''''
        END

    -- its a Date
    IF (@ColumnType = 4)	
        BEGIN
            -- Standardize the date format string
            SELECT @DateFS = REPLACE(@DateFS,'dd','d')
            SELECT @DateFS = REPLACE(@DateFS,'mm','m')
            SELECT @DateFS = REPLACE(@DateFS,'hh','h')
            SELECT @DateFS = REPLACE(@DateFS,'ss','s')
            SELECT @DateFS = REPLACE(@DateFS,'yyyy','y')
            SELECT @DateFS = REPLACE(@DateFS,':','/')
            SELECT @DateFS = REPLACE(@DateFS,'.','/')
            SELECT @DateFS = REPLACE(@DateFS,'-','/')
            SELECT @DateFS = REPLACE(@DateFS,' ','/')
            -- is it a date question that will be converted into a date field */
            IF @PartialDateFlag = 0 AND ((@DateFS = 'd/m/y') OR (@DateFS = 'm/d/y') OR (@DateFS = 'y/m/d') OR (@DateFS = 'h/m') OR (@DateFS = 'h/m/s') OR (@DateFS = 'd/m/y/h/m') OR (@DateFS = 'm/d/y/h/m') OR (@DateFS = 'y/m/d/h/m') OR (@DateFS = 'd/m/y/h/m/s') OR (@DateFS = 'm/d/y/h/m/s') OR (@DateFS = 'y/m/d/h/m/s'))
                BEGIN
                    IF @ResponseValue IS NULL
                        BEGIN
                            IF (@ResponseStatus = 10)
                                BEGIN
                                    IF @SVMissing <> ''
                                        BEGIN
                                            IF @SVMissing = '-1'
                                                SELECT @Response = 'CONVERT(DATETIME,''' +  '12/29/1899' + ''',101)'
                                            ELSE IF @SVMissing = '-2'
                                                SELECT @Response = 'CONVERT(DATETIME,''' +  '12/28/1899' + ''',101)'
                                            ELSE IF @SVMissing = '-3'
                                                SELECT @Response = 'CONVERT(DATETIME,''' +  '12/27/1899' + ''',101)'
                                            ELSE IF @SVMissing = '-4'
                                                SELECT @Response = 'CONVERT(DATETIME,''' +  '12/26/1899' + ''',101)'
                                            ELSE IF @SVMissing = '-5'
                                                SELECT @Response = 'CONVERT(DATETIME,''' +  '12/25/1899' + ''',101)'
                                            ELSE IF @SVMissing = '-6'
                                                SELECT @Response = 'CONVERT(DATETIME,''' +  '12/24/1899' + ''',101)'
                                            ELSE IF @SVMissing = '-7'
                                                SELECT @Response = 'CONVERT(DATETIME,''' +  '12/23/1899' + ''',101)'
                                            ELSE IF @SVMissing = '-8'
                                                SELECT @Response = 'CONVERT(DATETIME,''' +  '12/22/1899' + ''',101)'
                                            ELSE IF @SVMissing = '-9'
                                                SELECT @Response = 'CONVERT(DATETIME,''' +  '12/21/1899' + ''',101)'
                                        END
                                    ELSE
                                        SELECT @Response = 'null'
                                END
                            ELSE IF (@ResponseStatus = -8)
                                BEGIN
                                    IF @SVNotApplicable <> ''
                                        BEGIN
                                            IF @SVNotApplicable = '-1'
                                                SELECT @Response = 'CONVERT(DATETIME,''' +  '12/29/1899' + ''',101)'
                                            ELSE IF @SVNotApplicable = '-2'
                                                SELECT @Response = 'CONVERT(DATETIME,''' +  '12/28/1899' + ''',101)'
                                            ELSE IF @SVNotApplicable = '-3'
                                                SELECT @Response = 'CONVERT(DATETIME,''' +  '12/27/1899' + ''',101)'
                                            ELSE IF @SVNotApplicable = '-4'
                                                SELECT @Response = 'CONVERT(DATETIME,''' +  '12/26/1899' + ''',101)'
                                            ELSE IF @SVNotApplicable = '-5'
                                                SELECT @Response = 'CONVERT(DATETIME,''' +  '12/25/1899' + ''',101)'
                                            ELSE IF @SVNotApplicable = '-6'
                                                SELECT @Response = 'CONVERT(DATETIME,''' +  '12/24/1899' + ''',101)'
                                            ELSE IF @SVNotApplicable = '-7'
                                                SELECT @Response = 'CONVERT(DATETIME,''' +  '12/23/1899' + ''',101)'
                                            ELSE IF @SVNotApplicable = '-8'
                                                SELECT @Response = 'CONVERT(DATETIME,''' +  '12/22/1899' + ''',101)'
                                            ELSE IF @SVNotApplicable = '-9'
                                                SELECT @Response = 'CONVERT(DATETIME,''' +  '12/21/1899' + ''',101)'
                                        END
                                    ELSE
                                        SELECT @Response = 'null'
                                END
                            ELSE IF (@ResponseStatus = -5)
                                BEGIN
                                    IF @SVUnobtainable <> ''
                                        BEGIN
                                            IF @SVUnobtainable = '-1'
                                                SELECT @Response = 'CONVERT(DATETIME,''' +  '12/29/1899' + ''',101)'
                                            ELSE IF @SVUnobtainable = '-2'
                                                SELECT @Response = 'CONVERT(DATETIME,''' +  '12/28/1899' + ''',101)'
                                            ELSE IF @SVUnobtainable = '-3'
                                                SELECT @Response = 'CONVERT(DATETIME,''' +  '12/27/1899' + ''',101)'
                                            ELSE IF @SVUnobtainable = '-4'
                                                SELECT @Response = 'CONVERT(DATETIME,''' +  '12/26/1899' + ''',101)'
                                            ELSE IF @SVUnobtainable = '-5'
                                                SELECT @Response = 'CONVERT(DATETIME,''' +  '12/25/1899' + ''',101)'
                                            ELSE IF @SVUnobtainable = '-6'
                                                SELECT @Response = 'CONVERT(DATETIME,''' +  '12/24/1899' + ''',101)'
                                            ELSE IF @SVUnobtainable = '-7'
                                                SELECT @Response = 'CONVERT(DATETIME,''' +  '12/23/1899' + ''',101)'
                                            ELSE IF @SVUnobtainable = '-8'
                                                SELECT @Response = 'CONVERT(DATETIME,''' +  '12/22/1899' + ''',101)'
                                            ELSE IF @SVUnobtainable = '-9'
                                                SELECT @Response = 'CONVERT(DATETIME,''' +  '12/21/1899' + ''',101)'
                                        END
                                    ELSE
                                        SELECT @Response = 'null'
                                END
                            ELSE
                                SELECT @Response = 'null'
                        END
                    ELSE
                        BEGIN
                            -- Standardize the separators within the response
                            SELECT @DateResp = @ResponseValue
                            SELECT @DateResp = REPLACE(@DateResp,':','/')
                            SELECT @DateResp = REPLACE(@DateResp,'.','/')
                            SELECT @DateResp = REPLACE(@DateResp,'-','/')
                            SELECT @DateResp = REPLACE(@DateResp,' ','/')
                            -- Add a trailing slash to DateResp
                            SELECT @DateResp = @DateResp + '/'
                            -- Using the standardized format string the response is separated out into its individual parts and put together as a Universal Date Format string
                            SELECT @D1 = CHARINDEX('/',@DateResp)
                            SELECT @D2 = CHARINDEX('/',@DateResp,@D1+1)
                            SELECT @D3 = CHARINDEX('/',@DateResp,@D2+1)
                            SELECT @D4 = CHARINDEX('/',@DateResp,@D3+1)
                            SELECT @D5 = CHARINDEX('/',@DateResp,@D4+1)
                            SELECT @D6 = CHARINDEX('/',@DateResp,@D5+1)
                            -- Using DateFS as a template build a Universal Date Format string
                            IF (@DateFs = 'd/m/y')
                                BEGIN
                                    SELECT @P1 = SUBSTRING(@DateResp,1,(@D1-1))
                                    SELECT @P2 = SUBSTRING(@DateResp,(@D1+1),(@D2-@D1-1))
                                    SELECT @P3 = SUBSTRING(@DateResp,(@D2+1),(@D3-@D2-1))
                                    SELECT @DateResp = @P2 + '/' + @P1 + '/' + @P3
                                END
                            ELSE IF (@DateFs = 'm/d/y')
                                BEGIN
                                    SELECT @P1 = SUBSTRING(@DateResp,1,(@D1-1))
                                    SELECT @P2 = SUBSTRING(@DateResp,(@D1+1),(@D2-@D1-1))
                                    SELECT @P3 = SUBSTRING(@DateResp,(@D2+1),(@D3-@D2-1))
                                    SELECT @DateResp = @P1 + '/' + @P2 + '/' + @P3
                                END
                            ELSE IF (@DateFs = 'y/m/d')
                                BEGIN
                                    SELECT @P1 = SUBSTRING(@DateResp,1,(@D1-1))
                                    SELECT @P2 = SUBSTRING(@DateResp,(@D1+1),(@D2-@D1-1))
                                    SELECT @P3 = SUBSTRING(@DateResp,(@D2+1),(@D3-@D2-1))
                                    SELECT @DateResp = @P2 + '/' + @P3 + '/' + @P1
                                END
                            ELSE IF (@DateFs = 'h/m')
                                BEGIN
                                    SELECT @P1 = SUBSTRING(@DateResp,1,(@D1-1))
                                    SELECT @P2 = SUBSTRING(@DateResp,(@D1+1),(@D2-@D1-1))
                                    SELECT @DateResp = @P1 + ':' + @P2
                                END
                            ELSE IF (@DateFs = 'h/m/s')
                                BEGIN
                                    SELECT @P1 = SUBSTRING(@DateResp,1,(@D1-1))
                                    SELECT @P2 = SUBSTRING(@DateResp,(@D1+1),(@D2-@D1-1))
                                    SELECT @P3 = SUBSTRING(@DateResp,(@D2+1),(@D3-@D2-1))
                                    SELECT @DateResp = @P1 + ':' + @P2 + ':' + @P3
                                END
                            ELSE IF (@DateFs = 'd/m/y/h/m')
                                BEGIN
                                    SELECT @P1 = SUBSTRING(@DateResp,1,(@D1-1))
                                    SELECT @P2 = SUBSTRING(@DateResp,(@D1+1),(@D2-@D1-1))
                                    SELECT @P3 = SUBSTRING(@DateResp,(@D2+1),(@D3-@D2-1))
                                    SELECT @P4 = SUBSTRING(@DateResp,(@D3+1),(@D4-@D3-1))
                                    SELECT @P5 = SUBSTRING(@DateResp,(@D4+1),(@D5-@D4-1))
                                    SELECT @DateResp = @P2 + '/' + @P1 + '/' + @P3 + ' ' + @P4 + ':' + @P5
                                END
                            ELSE IF (@DateFs = 'm/d/y/h/m')
                                BEGIN
                                    SELECT @P1 = SUBSTRING(@DateResp,1,(@D1-1))
                                    SELECT @P2 = SUBSTRING(@DateResp,(@D1+1),(@D2-@D1-1))
                                    SELECT @P3 = SUBSTRING(@DateResp,(@D2+1),(@D3-@D2-1))
                                    SELECT @P4 = SUBSTRING(@DateResp,(@D3+1),(@D4-@D3-1))
                                    SELECT @P5 = SUBSTRING(@DateResp,(@D4+1),(@D5-@D4-1))
                                    SELECT @DateResp = @P1 + '/' + @P2 + '/' + @P3 + ' ' + @P4 + ':' + @P5
                                END
                            ELSE IF (@DateFs = 'y/m/d/h/m')
                                BEGIN
                                    SELECT @P1 = SUBSTRING(@DateResp,1,(@D1-1))
                                    SELECT @P2 = SUBSTRING(@DateResp,(@D1+1),(@D2-@D1-1))
                                    SELECT @P3 = SUBSTRING(@DateResp,(@D2+1),(@D3-@D2-1))
                                    SELECT @P4 = SUBSTRING(@DateResp,(@D3+1),(@D4-@D3-1))
                                    SELECT @P5 = SUBSTRING(@DateResp,(@D4+1),(@D5-@D4-1))
                                    SELECT @DateResp = @P2 + '/' + @P3 + '/' + @P1 + ' ' + @P4 + ':' + @P5
                                END
                            ELSE IF (@DateFs = 'd/m/y/h/m/s')
                                BEGIN
                                    SELECT @P1 = SUBSTRING(@DateResp,1,(@D1-1))
                                    SELECT @P2 = SUBSTRING(@DateResp,(@D1+1),(@D2-@D1-1))
                                    SELECT @P3 = SUBSTRING(@DateResp,(@D2+1),(@D3-@D2-1))
                                    SELECT @P4 = SUBSTRING(@DateResp,(@D3+1),(@D4-@D3-1))
                                    SELECT @P5 = SUBSTRING(@DateResp,(@D4+1),(@D5-@D4-1))
                                    SELECT @P6 = SUBSTRING(@DateResp,(@D5+1),(@D6-@D5-1))
                                    SELECT @DateResp = @P2 + '/' + @P1 + '/' + @P3 + ' ' + @P4 + ':' + @P5 + ':' + @P6
                                END
                            ELSE IF (@DateFs = 'm/d/y/h/m/s')
                                BEGIN
                                    SELECT @P1 = SUBSTRING(@DateResp,1,(@D1-1))
                                    SELECT @P2 = SUBSTRING(@DateResp,(@D1+1),(@D2-@D1-1))
                                    SELECT @P3 = SUBSTRING(@DateResp,(@D2+1),(@D3-@D2-1))
                                    SELECT @P4 = SUBSTRING(@DateResp,(@D3+1),(@D4-@D3-1))
                                    SELECT @P5 = SUBSTRING(@DateResp,(@D4+1),(@D5-@D4-1))
                                    SELECT @P6 = SUBSTRING(@DateResp,(@D5+1),(@D6-@D5-1))
                                    SELECT @DateResp = @P1 + '/' + @P2 + '/' + @P3 + ' ' + @P4 + ':' + @P5 + ':' + @P6
                                END
                            ELSE IF (@DateFs = 'y/m/d/h/m/s')
                                BEGIN
                                    SELECT @P1 = SUBSTRING(@DateResp,1,(@D1-1))
                                    SELECT @P2 = SUBSTRING(@DateResp,(@D1+1),(@D2-@D1-1))
                                    SELECT @P3 = SUBSTRING(@DateResp,(@D2+1),(@D3-@D2-1))
                                    SELECT @P4 = SUBSTRING(@DateResp,(@D3+1),(@D4-@D3-1))
                                    SELECT @P5 = SUBSTRING(@DateResp,(@D4+1),(@D5-@D4-1))
                                    SELECT @P6 = SUBSTRING(@DateResp,(@D5+1),(@D6-@D5-1))
                                    SELECT @DateResp = @P2 + '/' + @P3 + '/' + @P1 + ' ' + @P4 + ':' + @P5 + ':' + @P6
                                END

                            -- The SQL Server CONVERT function is set to use style 101 which is mm/dd/yyyy
                            -- Note that times without a date is always given the date 01/01/1900
                            SELECT @Response = 'CONVERT(DATETIME,''' +  @DateResp + ''',101)'
                        END
                END
            ELSE
                BEGIN
                    -- date formats y/m, m/y and y/d/m are not converted to dates, they remain strings
                    if @ResponseValue is NULL
                        BEGIN
                            IF (@ResponseStatus = 10)
                                BEGIN
                                    IF @SVMissing <> ''
                                        SELECT @Response = @SVMissing
                                    ELSE
                                        SELECT @Response = 'null'
                                END
                            ELSE IF (@ResponseStatus = -8)
                                BEGIN
                                    IF @SVNotApplicable <> ''
                                        SELECT @Response = @SVNotApplicable
                                    ELSE
                                        SELECT @Response = 'null'
                                END
                            ELSE IF (@ResponseStatus = -5)
                                BEGIN
                                    IF @SVUnobtainable <> ''
                                        SELECT @Response = @SVUnobtainable
                                    ELSE
                                        SELECT @Response = 'null'
                                END
                            ELSE
                                SELECT @Response = 'null'
                        END
                    ELSE
                        SELECT @Response = '''' + replace(@responsevalue,'''','''''') + ''''
                END
        END

    -- its Category, is it the code or the value that is required 
    IF (@ColumnType = 1)
        BEGIN
            IF @OutputCatValues = 1
                BEGIN
                    IF @ResponseValue IS NULL
                        BEGIN
                            IF (@ResponseStatus = 10)
                                BEGIN
                                    IF @SVMissing <> ''
                                        SELECT @Response = @SVMissing
                                    ELSE
                                        SELECT @Response = 'null'
                                END
                            ELSE IF (@ResponseStatus = -8)
                                BEGIN
                                    IF @SVNotApplicable <> ''
                                        SELECT @Response = @SVNotApplicable
                                    ELSE
                                        SELECT @Response = 'null'
                                END
                            ELSE IF (@ResponseStatus = -5)
                                BEGIN
                                    IF @SVUnobtainable <> ''
                                        SELECT @Response = @SVUnobtainable
                                    ELSE
                                        SELECT @Response = 'null'
                                END
                            ELSE
                                SELECT @Response = 'null'
                        END
                    ELSE
                        SELECT @Response = '''' + replace(@responsevalue,'''','''''') + ''''
                END
            ELSE
                BEGIN
                    IF @ValueCode IS NULL
                        BEGIN
                            IF (@ResponseStatus = 10)
                                BEGIN
                                    IF @SVMissing <> ''
                                        SELECT @Response = @SVMissing
                                    ELSE
                                        SELECT @Response = 'null'
                                END
                            ELSE IF (@ResponseStatus = -8)
                                BEGIN
                                    IF @SVNotApplicable <> ''
                                        SELECT @Response = @SVNotApplicable
                                    ELSE
                                        SELECT @Response = 'null'
                                END
                            ELSE IF (@ResponseStatus = -5)
                                BEGIN
                                    IF @SVUnobtainable <> ''
                                        SELECT @Response = @SVUnobtainable
                                    ELSE
                                        SELECT @Response = 'null'
                                END
                            ELSE
                                SELECT @Response = 'null'
                        END
                    ELSE
                        SELECT @Response = char(39) + @ValueCode + char(39)
                END
        END
    
    -- its IntegerData, Real or LabTest 
    IF (@ColumnType = 2) OR (@ColumnType = 3) OR (@ColumnType = 6)
        BEGIN
            IF @ResponseValue IS NULL
                BEGIN
                    IF (@ResponseStatus = 10)
                        BEGIN
                            IF @SVMissing <> ''
                                SELECT @Response = @SVMissing
                            ELSE
                                SELECT @Response = 'null'
                        END
                    ELSE IF (@ResponseStatus = -8)
                        BEGIN
                            IF @SVNotApplicable <> ''
                                SELECT @Response = @SVNotApplicable
                            ELSE
                                SELECT @Response = 'null'
                        END
                    ELSE IF (@ResponseStatus = -5)
                        BEGIN
                            IF @SVUnobtainable <> ''
                                SELECT @Response = @SVUnobtainable
                            ELSE
                                SELECT @Response = 'null'
                        END
                    ELSE
                        SELECT @Response = 'null'
                END
            ELSE
                SELECT @Response = @ResponseValue
        END

    SELECT         @CurrentRows = DataViewTrigCalc
    FROM           DataViewDetails
    WHERE          ClinicalTrialId = @ClinicalTrialId

  IF @OWQGroupID = 0 
      BEGIN
          IF @CurrentRows = 'I'
              BEGIN
                  SELECT @strSQL =           ' INSERT INTO ' + @TableName + ''
                  SELECT @strSQL = @strSQL + ' (ClinicalTrialId, Site, PersonId, VisitId, VisitCycleNumber, CRFPageId, '
                  SELECT @strSQL = @strSQL + ' CRFPageCycleNumber, [' + @ColumnName + ']) ' 
                  SELECT @strSQL = @strSQL + ' VALUES(' + CONVERT(VARCHAR(20),@ClinicalTrialId) + ', '
                  SELECT @strSQL = @strSQL + '' + char(39) + @TrialSite + char(39) + ','
                  SELECT @strSQL = @strSQL + '' + CONVERT(VARCHAR(20),@PersonId) + ','
                  SELECT @strSQL = @strSQL + '' + CONVERT(VARCHAR(20),@VisitId) + ','
                  SELECT @strSQL = @strSQL + '' + CONVERT(VARCHAR(20),@VisitCycleNumber) + ','
                  SELECT @strSQL = @strSQL + '' + CONVERT(VARCHAR(20),@CRFPageId) + ','
                  SELECT @strSQL = @strSQL + '' + CONVERT(VARCHAR(20),@CRFPageCycleNumber) + ','
                  SELECT @strSQL = @strSQL + '' + CONVERT(VARCHAR(512),@Response) + ')'
              END
          ELSE
              BEGIN
                  SELECT @strSQL =           ' UPDATE ' + @TableName + ''
                  SELECT @strSQL = @strSQL + ' SET [' + @ColumnName + '] = ' + CONVERT(VARCHAR(512),@Response)
                  SELECT @strSQL = @strSQL + ' WHERE ClinicalTrialId = ' + char(39) + CONVERT(VARCHAR(20),@ClinicalTrialId) + char(39) + ''
                  SELECT @strSQL = @strSQL + ' AND Site = ' + char(39) + @TrialSite + char(39)  
                  SELECT @strSQL = @strSQL + ' AND PersonId = ' + CONVERT(VARCHAR(20),@PersonId) + ''
                  SELECT @strSQL = @strSQL + ' AND VisitId = ' + CONVERT(VARCHAR(20),@VisitId) + ''  
                  SELECT @strSQL = @strSQL + ' AND VisitCycleNumber = ' + CONVERT(VARCHAR(20),@VisitCycleNumber) + ''
                  SELECT @strSQL = @strSQL + ' AND CRFPageId = ' + CONVERT(VARCHAR(20),@CRFPageId) + ''
                  SELECT @strSQL = @strSQL + ' AND CRFPageCycleNumber = ' + CONVERT(VARCHAR(20),@CRFPageCycleNumber) + ''
              END
      END
  ELSE
      BEGIN
          IF @CurrentRows = 'I'
              BEGIN
                  SELECT @strSQL =           ' INSERT INTO ' + @TableName + ''
                  SELECT @strSQL = @strSQL + ' (ClinicalTrialId, Site, PersonId, VisitId, VisitCycleNumber, CRFPageId, '
                  SELECT @strSQL = @strSQL + ' CRFPageCycleNumber, OwnerQGroupID, RepeatNumber, [' + @ColumnName + ']) ' 
                  SELECT @strSQL = @strSQL + ' VALUES(' + CONVERT(VARCHAR(20),@ClinicalTrialId) + ', '
                  SELECT @strSQL = @strSQL + '' + char(39) + @TrialSite + char(39) + ','
                  SELECT @strSQL = @strSQL + '' + CONVERT(VARCHAR(20),@PersonId) + ','
                  SELECT @strSQL = @strSQL + '' + CONVERT(VARCHAR(20),@VisitId) + ','
                  SELECT @strSQL = @strSQL + '' + CONVERT(VARCHAR(20),@VisitCycleNumber) + ','
                  SELECT @strSQL = @strSQL + '' + CONVERT(VARCHAR(20),@CRFPageId) + ','
                  SELECT @strSQL = @strSQL + '' + CONVERT(VARCHAR(20),@CRFPageCycleNumber) + ','
                  SELECT @strSQL = @strSQL + '' + CONVERT(VARCHAR(20),@OWQGroupID) + ','
                  SELECT @strSQL = @strSQL + '' + CONVERT(VARCHAR(20),@RepeatNumber) + ','
                  SELECT @strSQL = @strSQL + '' + CONVERT(VARCHAR(512),@Response) + ')'
              END
          ELSE
              BEGIN
                  SELECT @strSQL =           ' UPDATE ' + @TableName + ''
                  SELECT @strSQL = @strSQL + ' SET [' + @ColumnName + '] = ' + CONVERT(VARCHAR(512),@Response)
                  SELECT @strSQL = @strSQL + ' WHERE ClinicalTrialId = ' + char(39) + CONVERT(VARCHAR(20),@ClinicalTrialId) + char(39) + ''
                  SELECT @strSQL = @strSQL + ' AND Site = ' + char(39) + @TrialSite + char(39)  
                  SELECT @strSQL = @strSQL + ' AND PersonId = ' + CONVERT(VARCHAR(20),@PersonId) + ''
                  SELECT @strSQL = @strSQL + ' AND VisitId = ' + CONVERT(VARCHAR(20),@VisitId) + ''  
                  SELECT @strSQL = @strSQL + ' AND VisitCycleNumber = ' + CONVERT(VARCHAR(20),@VisitCycleNumber) + ''
                  SELECT @strSQL = @strSQL + ' AND CRFPageId = ' + CONVERT(VARCHAR(20),@CRFPageId) + ''
                  SELECT @strSQL = @strSQL + ' AND CRFPageCycleNumber = ' + CONVERT(VARCHAR(20),@CRFPageCycleNumber) + ''
                  SELECT @strSQL = @strSQL + ' AND OwnerQGroupID = ' + CONVERT(VARCHAR(20),@OWQGroupID) + ''
                  SELECT @strSQL = @strSQL + ' AND RepeatNumber = ' + CONVERT(VARCHAR(20),@RepeatNumber) + ''
              END
      END

      EXEC (@strSQL)

      SELECT @err = @@ERROR, @rows = @@ROWCOUNT

      IF @err <> 0
          BEGIN
              CLOSE Macro_DVROcursor
              DEALLOCATE Macro_DVROcursor
              RETURN
          END
      IF @rows <> 1
          BEGIN
              CLOSE Macro_DVROcursor
              DEALLOCATE Macro_DVROcursor
              RETURN
          END

--~~~~~~~~~~~~~~~~~~~~
 --cursor wrapper end
      	FETCH NEXT FROM Macro_DVROcursor INTO @ClinicalTrialId, @VisitId, @CRFPageId, @TrialSite, @PersonId,
                @VisitCycleNumber, @CRFPageCycleNumber, @ResponseValue, @ResponseStatus, @DataItemId, @ValueCode, @RepeatNumber
   END
  
CLOSE Macro_DVROcursor
DEALLOCATE Macro_DVROcursor

--~~~~~~~~~~~~~~~~~~~~
END

