--//--SEPARATOR=--//--
CREATE TRIGGER SetDBTimeStamp_DIR ON DataItemResponse
FOR INSERT, UPDATE
AS
BEGIN
	DECLARE C1 CURSOR FOR SELECT ClinicalTrialID,TrialSite,PersonID,ResponseTaskID, RepeatNumber, DatabaseTimeStamp  from Inserted;
	DECLARE @ClinicalTrialID int
	DECLARE @TrialSite varchar(15)
	DECLARE @PersonID int
	DECLARE @ResponseTaskID int
	DECLARE @RepeatNumber smallint
	DECLARE @DatabaseTimestamp numeric(16,10)
	DECLARE @Timezone varchar(15)
	
	-- RS 03/08/2003: Get Offset from server, instead of DB
	-- Override timezone, calculate directly from server	
	SET @Timezone = datediff(minute,getdate(),getutcdate())

	OPEN C1;
	FETCH NEXT FROM C1 INTO @ClinicalTrialID, @TrialSite, @PersonID, @ResponseTaskID, @RepeatNumber, @DatabaseTimestamp
	WHILE (@@FETCH_STATUS <> -1)
	BEGIN
		PRINT 'ClinicalTrialID = ' + STR(@ClinicalTrialID)
		PRINT 'TrialSite = ' + @TrialSite
		PRINT 'PersonID = ' + STR(@PersonID)
		PRINT 'ResponseTaskID = ' + STR(@ResponseTaskID)
		PRINT 'RepeatNumber = ' + STR(@RepeatNumber)
		PRINT 'DatabaseTimestamp = '  + STR(@DatabaseTimestamp,16,10)
		IF @DatabaseTimestamp = 0 OR @DatabaseTimestamp IS NULL
		BEGIN
			PRINT 'Update DatabaseTimestamp'
			/* HAVE TO UPDATE VALUE */
			UPDATE DataItemResponse SET DatabaseTimeStamp=(select convert(numeric(19,10),getdate())+2 from macrocontrol), 
				DatabaseTimestamp_TZ= @Timezone
			WHERE ClinicalTrialID = @ClinicalTrialID
				AND TrialSite = @TrialSite
				AND PersonID = @PersonID
				AND ResponseTaskID = @ResponseTaskID
				AND RepeatNumber = @RepeatNumber
		END
		FETCH NEXT FROM C1 INTO @ClinicalTrialID, @TrialSite, @PersonID, @ResponseTaskID, @RepeatNumber, @DatabaseTimestamp
	END
	CLOSE C1
	DEALLOCATE C1
END
--//--
CREATE TRIGGER SetDBTimeStamp_DIRH ON DataItemResponseHistory
FOR INSERT, UPDATE
AS
BEGIN
	DECLARE C1 CURSOR FOR SELECT ClinicalTrialID,TrialSite,PersonID,ResponseTaskID,RepeatNumber,ResponseTimestamp,DatabaseTimeStamp  from Inserted;
	DECLARE @ClinicalTrialID int
	DECLARE @TrialSite varchar(15)
	DECLARE @PersonID int
	DECLARE @ResponseTaskID int
	DECLARE @DatabaseTimestamp numeric(16,10)
	DECLARE @ResponseTimestamp numeric(16,10)
	DECLARE @RepeatNumber smallint
	DECLARE @Timezone varchar(15)
	
	-- RS 03/08/2003: Get Offset from server, instead of DB
	-- Override timezone, calculate directly from server
	SET @Timezone = datediff(minute,getdate(),getutcdate())

	OPEN C1;
	FETCH NEXT FROM C1 INTO @ClinicalTrialID, @TrialSite, @PersonID, @ResponseTaskID, @RepeatNumber, @ResponseTimestamp, @DatabaseTimeStamp
	WHILE (@@FETCH_STATUS <> -1)
	BEGIN
		PRINT 'ClinicalTrialID = ' + STR(@ClinicalTrialID)
		PRINT 'TrialSite = ' + @TrialSite
		PRINT 'PersonID = ' + STR(@PersonID)
		PRINT 'ResponseTaskID = ' + STR(@ResponseTaskID)
		PRINT 'ResponseTimestamp = '  + STR(@ResponseTimestamp,16,10)
		PRINT 'RepeatNumber = ' + STR(@RepeatNumber)
		PRINT 'DatabaseTimestamp = '  + STR(@DatabaseTimestamp,16,10)
		IF @DatabaseTimestamp = 0 OR @DatabaseTimestamp IS NULL
		BEGIN
			PRINT 'Update DatabaseTimestamp'
			/* HAVE TO UPDATE VALUE */
			UPDATE DataItemResponseHistory SET DatabaseTimeStamp=(select convert(numeric(19,10),getdate())+2 from macrocontrol), 
				DatabaseTimestamp_TZ= @Timezone
			WHERE ClinicalTrialID = @ClinicalTrialID
				AND TrialSite = @TrialSite
				AND PersonID = @PersonID
				AND ResponseTaskID = @ResponseTaskID
				AND ResponseTimestamp = @ResponseTimestamp
				AND RepeatNumber = @RepeatNumber
		END
		FETCH NEXT FROM C1 INTO @ClinicalTrialID, @TrialSite, @PersonID, @ResponseTaskID, @RepeatNumber, @ResponseTimestamp, @DatabaseTimeStamp
	END
	CLOSE C1
	DEALLOCATE C1
END
--//--