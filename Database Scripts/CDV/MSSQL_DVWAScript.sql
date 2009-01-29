/*
------------------------------------------------------------------
Copyright:   InferMed Ltd. 2001. All Rights Reserved
File:        MACRO_DVWATRIGGER (DVWAScript.sql)
Author:      Richard Weare
Purpose:     Script to create a trigger that fills 'With Attributes' 
             dataviews when DataItemResponse is inserted into or 
             updated.
------------------------------------------------------------------
Mo Morris    10/5/2002, changes around handling single quotes in the ResponseValue field
ASH          8/8/2002,  Modified to account for RQGs
Mo Morris    28/8/2002, RQG changes checked
	TA 11/02/2004: changed tablename variable to 255 length
MLM 20/04/05: Modified for study-specific SDV.
------------------------------------------------------------------
*/

--IF EXISTS (SELECT name FROM sysobjects WHERE name = 'Macro_DVWAtrigger' AND type = 'TR')
--    DROP TRIGGER Macro_DVWAtrigger
--GO

CREATE TRIGGER dbo.Macro_DVWAtrigger 
ON             dbo.DataItemResponse
FOR            INSERT, UPDATE
AS
IF UPDATE(ClinicalTrialId)    OR 
   UPDATE(CRFPageCycleNumber) OR
   UPDATE(CRFPageId)          OR
   UPDATE(CTCGrade)           OR
   UPDATE(DataItemId)         OR
   UPDATE(LabResult)          OR
   UPDATE(PersonId)           OR
   UPDATE(ResponseStatus)     OR
   UPDATE(ResponseTimeStamp)  OR    
   UPDATE(ResponseValue)      OR
   UPDATE(TrialSite)          OR
   UPDATE(ValueCode)          OR
   UPDATE(VisitCycleNumber)   OR
   UPDATE(VisitId)          
BEGIN            
  SET NOCOUNT ON

  DECLARE        @strReturnDate	     VARCHAR(19),
                 @str_worker         VARCHAR(4),
                 @strZeroFiller      CHAR(1),

                 @TableName          VARCHAR(255),
                 @SeparateVisits     INT,
                 @strSQL             VARCHAR(500),
                 @err                INT,
                 @rows               INT,
                 
                 @ClinicalTrialId    INT,
                 @VisitId            INT,
                 @CRFPageId          INT,
                 @ResponseTimeStamp  DECIMAL(16,10),
                 @CalcDataItemCode   VARCHAR(15),
                 @CalcDataType       INT,
                 @TrialSite          VARCHAR(8),
                 @PersonId           INT,
                 @VisitCycleNumber   INT,
                 @CRFPageCycleNumber SMALLINT,
                 @ResponseValue      VARCHAR(255),
                 @DataItemId         INT,
                 @ValueCode          VARCHAR(15),
                 @LabResult          VARCHAR(6),
                 @Units              VARCHAR(15),
                 @CTCGrade           INT,
                 @strCTCGrade        VARCHAR(6),
                 @ResponseStatus     INT,
-- ASH 8/08/2002
                 @OwnerQGroupID      INT,
                 @RepeatNUmber       INT

  IF NOT EXISTS (SELECT * FROM sysobjects WHERE id = object_id(N'[dbo].[DataViewDetails]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
  RETURN

  IF NOT EXISTS (SELECT * FROM sysobjects WHERE id = object_id(N'[dbo].[DataViewTables]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
  RETURN

--~~~~~~~~~~~~~~~~~~~~
--cursor wrapper
--~~~~~~~~~~~~~~~~~~~~
  DECLARE         Macro_DVWAcursor CURSOR
  FOR             SELECT         ResponseValue,       ValueCode,  LabResult,        UnitofMeasurement,   CTCGrade,
                                 ClinicalTrialId,     VisitId,    CRFPageId,        ResponseTimeStamp,   DataItemId,
                                 TrialSite,           PersonId,   VisitCycleNumber, CRFPageCycleNumber,  ResponseStatus, RepeatNumber
	          FROM         INSERTED

  OPEN            Macro_DVWAcursor 

FETCH NEXT FROM Macro_DVWAcursor INTO @ResponseValue,   @ValueCode, @LabResult,        @Units,             @CTCGrade,
                                    @ClinicalTrialId, @VisitId,   @CRFPageId,        @ResponseTimeStamp, @DataItemId, 
                                    @TrialSite,       @PersonId,  @VisitCycleNumber, @CRFPageCycleNumber, @ResponseStatus, @RepeatNumber

    WHILE @@FETCH_STATUS = 0
    BEGIN

--~~~~~~~~~~~~~~~~~~~

-- Mo 10/5/2002, QUOTENAME call added to handle single quotes
-- SELECT @ResponseValue = ISNULL(@ResponseValue,'null')

IF @ResponseValue IS NULL
    SELECT @ResponseValue = 'null'
ELSE
    SELECT @ResponseValue = QUOTENAME(@ResponseValue,char(39))

SELECT @ValueCode =     ISNULL(@ValueCode,'null')
SELECT @LabResult =     ISNULL(@LabResult,'null')
SELECT @Units =         ISNULL(@Units,'null')

-- ASH 8/08/2002: SQL to get  OwnerQGroupID  
SELECT      @OwnerQGroupID = OwnerQGroupID 
FROM        CRFElement
WHERE       CRFElement.ClinicalTrialId = @ClinicalTrialId		
AND         CRFElement.DataItemId  = @DataItemId 
AND         CRFElement.CRFPageId = @CRFPageId

IF    @CTCGrade IS null
  SELECT @strCTCGrade =   'null'
ELSE
  SELECT @strCTCGrade =  CONVERT(VARCHAR(20),@CTCGrade)

  SELECT          @SeparateVisits = DataViewSeparateVisits 
  FROM            DataViewDetails
  WHERE           ClinicalTrialId = @ClinicalTrialId
    
  IF              @SeparateVisits = 1

    SELECT        @TableName = DataViewName
    FROM          DataViewTables
    WHERE         ClinicalTrialId = @ClinicalTrialId
    AND           VisitId = @VisitId
    AND           CRFPageId = @CRFPageId
    AND           DataViewType = 'WA'

  ELSE
	
    SELECT        @TableName = DataViewName
    FROM          DataViewTables
    WHERE         ClinicalTrialId = @ClinicalTrialId
    AND           CRFPageId = @CRFPageId
    AND           DataViewType = 'WA'

    IF @TableName IS NULL
      BEGIN 
        CLOSE Macro_DVWAcursor
        DEALLOCATE Macro_DVWAcursor
        RETURN
      END
    ELSE

-- Julian date in SQL Server is Jan 01 1900 but VB is 
--30-Dec-1899 so 2 days have been subtracted to realign.

SELECT @ResponseTimeStamp = @ResponseTimeStamp -2

--assign year
select @strReturnDate = CONVERT(CHAR(4),DATEPART(yyyy,@ResponseTimeStamp))

--assign month
SELECT @strZeroFiller = ''
IF DATEPART(mm,@ResponseTimeStamp) < 10
   SELECT @strZeroFiller = '0'
   SELECT @str_worker = RIGHT(@strZeroFiller + CONVERT(VARCHAR(2),DATEPART(mm,@ResponseTimeStamp)),2)
   SELECT @strReturnDate = @strReturnDate + '/' + @str_worker

--assign day
SELECT @strZeroFiller = ''
IF DATEPART(dd,@ResponseTimeStamp) < 10
   SELECT @strZeroFiller = '0'
   SELECT @str_worker = RIGHT(@strZeroFiller + CONVERT(VARCHAR(2),DATEPART(dd,@ResponseTimeStamp)),2)
   SELECT @strReturnDate = @strReturnDate + '/' + @str_worker 

--assign hour
SELECT @strZeroFiller = ''
IF DATEPART(hh,@ResponseTimeStamp) < 10
   SELECT @strZeroFiller = '0'
   SELECT @str_worker = RIGHT(@strZeroFiller + CONVERT(VARCHAR(2),DATEPART(hh,@ResponseTimeStamp)),2)
   SELECT @strReturnDate = @strReturnDate + ' ' + @str_worker

--assign minutes
SELECT @strZeroFiller = ''
IF DATEPART(mi,@ResponseTimeStamp) < 10
   SELECT @strZeroFiller = '0'
   SELECT @str_worker = RIGHT(@strZeroFiller + CONVERT(VARCHAR(2),DATEPART(mi,@ResponseTimeStamp)),2)
   SELECT @strReturnDate = @strReturnDate + ':' + @str_worker

--assign seconds
SELECT @strZeroFiller = ''
IF DATEPART(ss,@ResponseTimeStamp) < 10
    SELECT @strZeroFiller = '0'
    SELECT @str_worker = RIGHT(@strZeroFiller + CONVERT(VARCHAR(2),DATEPART(ss,@ResponseTimeStamp)),2)
    SELECT @strReturnDate = @strReturnDate + ':' + @str_worker

    SELECT        @CalcDataItemCode = DataItemCode,
                  @CalcDataType = DataType
    FROM          DataItem
    WHERE         DataItem.ClinicalTrialId = @ClinicalTrialId
    AND           DataItem.DataItemId = @DataItemId

    SELECT @strSQL =           'DELETE FROM [' + @TableName + ']'
    SELECT @strSQL = @strSQL + ' WHERE ClinicalTrialId = ' + CONVERT(VARCHAR(20),@ClinicalTrialId) + ''
    SELECT @strSQL = @strSQL + ' AND Site = ' + char(39) + @TrialSite + char(39) + ''
    SELECT @strSQL = @strSQL + ' AND PersonId = ' + CONVERT(VARCHAR(20),@PersonId) + ''
    SELECT @strSQL = @strSQL + ' AND VisitId = ' + CONVERT(VARCHAR(20),@VisitId) + ''
    SELECT @strSQL = @strSQL + ' AND VisitCycleNumber = ' + CONVERT(VARCHAR(20),@VisitCycleNumber) + ''
    SELECT @strSQL = @strSQL + ' AND CRFPageId = ' + CONVERT(VARCHAR(20),@CRFPageId) + ''
    SELECT @strSQL = @strSQL + ' AND CRFPageCycleNumber = ' + CONVERT(VARCHAR(20),@CRFPageCycleNumber) + ''
    SELECT @strSQL = @strSQL + ' AND DataItemCode = ' + char(39) + @CalcDataItemCode + char(39) + ''
    SELECT @strSQL = @strSQL + ' AND RepeatNumber= ' + CONVERT(VARCHAR(20),@RepeatNumber) + ''
-- ASH 8/08/2002, Added RepeatNumber


    EXEC (@strSQL)

    SELECT @err = @@ERROR, @rows = @@ROWCOUNT

    IF @err <> 0
      BEGIN 
        CLOSE Macro_DVWAcursor
        DEALLOCATE Macro_DVWAcursor
        RETURN
      END
    IF @rows > 1
      BEGIN 
        CLOSE Macro_DVWAcursor
        DEALLOCATE Macro_DVWAcursor
        RETURN
      END

    SELECT @strSQL =           ' INSERT INTO [' + @TableName + ']'
    SELECT @strSQL = @strSQL + ' VALUES(' + CONVERT(VARCHAR(20),@ClinicalTrialId) + ', '
    SELECT @strSQL = @strSQL + '' + char(39) + @TrialSite + char(39) + ','
    SELECT @strSQL = @strSQL + '' + CONVERT(VARCHAR(20),@PersonId) + ','
    SELECT @strSQL = @strSQL + '' + CONVERT(VARCHAR(20),@VisitId) + ','
    SELECT @strSQL = @strSQL + '' + CONVERT(VARCHAR(20),@VisitCycleNumber) + ','
    SELECT @strSQL = @strSQL + '' + CONVERT(VARCHAR(20),@CRFPageId) + ','
    SELECT @strSQL = @strSQL + '' + CONVERT(VARCHAR(20),@CRFPageCycleNumber) + ','
-- ASH 8/08/2002, Added OwnerQGroupID
    SELECT @strSQL = @strSQL + '' + CONVERT(VARCHAR(20),@OwnerQGroupID) + ','
    SELECT @strSQL = @strSQL + '' + char(39) + @CalcDataItemCode + char(39) +','
-- ASH 8/08/2002, Added RepeatNumber
    SELECT @strSQL = @strSQL + '' + CONVERT(VARCHAR(20),@RepeatNumber) + ','
    SELECT @strSQL = @strSQL + '' + char(39) + @strReturnDate + char(39) +','
    SELECT @strSQL = @strSQL + '' + CONVERT(VARCHAR(20),@CalcDataType) + ','
-- Mo 10/5/2002, Quotes added to ResponseValue earlier by call to QUOTENAME
--  SELECT @strSQL = @strSQL + '' + char(39) + @ResponseValue + char(39) +','
    SELECT @strSQL = @strSQL + '' + @ResponseValue +','
    SELECT @strSQL = @strSQL + '' + char(39) + @Units + char(39) +','
    SELECT @strSQL = @strSQL + '' + char(39) + @ValueCode + char(39) +','
    SELECT @strSQL = @strSQL + '' + CONVERT(VARCHAR(20),@ResponseStatus) + ','
    SELECT @strSQL = @strSQL + '' + @LabResult + ','
    SELECT @strSQL = @strSQL + '' + @strCTCGrade + ')'

    EXEC (@strSQL)

    SELECT @err = @@ERROR, @rows = @@ROWCOUNT

    IF @err <> 0
      BEGIN 
        CLOSE Macro_DVWAcursor
        DEALLOCATE Macro_DVWAcursor
        RETURN
      END
    IF @rows <> 1
      BEGIN 
        CLOSE Macro_DVWAcursor
        DEALLOCATE Macro_DVWAcursor
        RETURN
      END

--~~~~~~~~~~~~~~~~~~~~
--cursor wrapper end
FETCH NEXT FROM Macro_DVWAcursor INTO @ResponseValue,   @ValueCode, @LabResult,        @Units,             @CTCGrade,
                                    @ClinicalTrialId, @VisitId,   @CRFPageId,        @ResponseTimeStamp, @DataItemId, 
                                    @TrialSite,       @PersonId,  @VisitCycleNumber, @CRFPageCycleNumber, @ResponseStatus, @RepeatNumber
    END
  CLOSE Macro_DVWAcursor
  DEALLOCATE Macro_DVWAcursor
--~~~~~~~~~~~~~~~~~~~~
END