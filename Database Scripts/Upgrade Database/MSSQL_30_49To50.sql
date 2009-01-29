--//--SEPARATOR=--//--
DROP TRIGGER SetDBTimeStamp_DIR
--//--
DROP TRIGGER SetDBTimeStamp_DIRH
--//--
GO
--//--
CREATE TRIGGER SetDBTimeStamp_DIR ON DataItemResponse FOR INSERT, UPDATE AS
BEGIN
	DECLARE C1 CURSOR FOR SELECT ClinicalTrialID,TrialSite,PersonID,ResponseTaskID,DatabaseTimeStamp from Inserted;
	DECLARE @ClinicalTrialID int
	DECLARE @TrialSite varchar(15)
	DECLARE @PersonID int
	DECLARE @ResponseTaskID int
	DECLARE @DatabaseTimeStamp numeric(16,10)
	DECLARE @Timezone varchar(15)
	
	-- Get Timezone value from MACRODBSettingsTable
	-- set @Timezone = (select settingvalue from macrodbsetting where settingsection='timezone' and settingkey='dbtz')
	-- PRINT 'Timezone = ' + @Timezone
	-- IF @Timezone IS NULL
	-- BEGIN
	-- 	SET @Timezone = '0'
	-- END

	-- RS 03/08/2003: Get Offset from server, instead of DB
	-- Override timezone, calculate directly from server
	SET @Timezone = datediff(minute,getdate(),getutcdate())

	OPEN C1;
	FETCH NEXT FROM C1 INTO @ClinicalTrialID, @TrialSite, @PersonID, @ResponseTaskID, @DatabaseTimestamp
	WHILE (@@FETCH_STATUS <> -1)
	BEGIN
		PRINT 'ClinicalTrialID = ' + STR(@ClinicalTrialID)
		PRINT 'TrialSite = ' + @TrialSite
		PRINT 'PersonID = ' + STR(@PersonID)
		PRINT 'ResponseTaskID = ' + STR(@ResponseTaskID)
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
		END
		FETCH NEXT FROM C1 INTO @ClinicalTrialID, @TrialSite, @PersonID, @ResponseTaskID, @DatabaseTimestamp
	END
	CLOSE C1
	DEALLOCATE C1
END
--//--
GO
--//-- 
CREATE TRIGGER SetDBTimeStamp_DIRH ON DataItemResponseHistory FOR INSERT, UPDATE AS
BEGIN
	DECLARE C1 CURSOR FOR SELECT ClinicalTrialID,TrialSite,PersonID,ResponseTaskID,DatabaseTimeStamp,ResponseTimestamp from Inserted;
	DECLARE @ClinicalTrialID int
	DECLARE @TrialSite varchar(15)
	DECLARE @PersonID int
	DECLARE @ResponseTaskID int
	DECLARE @DatabaseTimeStamp numeric(16,10)
	DECLARE @ResponseTimeStamp numeric(16,10)
	DECLARE @Timezone varchar(15)
	
	-- Get Timezone value from MACRODBSettingsTable
	-- set @Timezone = (select settingvalue from macrodbsetting where settingsection='timezone' and settingkey='dbtz')
	-- PRINT 'Timezone = ' + @Timezone
	-- IF @Timezone IS NULL
	-- BEGIN
	-- 	SET @Timezone = '0'
	-- END
	
	-- RS 03/08/2003: Get Offset from server, instead of DB
	-- Override timezone, calculate directly from server
	SET @Timezone = datediff(minute,getdate(),getutcdate())

	OPEN C1;
	FETCH NEXT FROM C1 INTO @ClinicalTrialID, @TrialSite, @PersonID, @ResponseTaskID, @DatabaseTimestamp, @ResponseTimestamp
	WHILE (@@FETCH_STATUS <> -1)
	BEGIN
		PRINT 'ClinicalTrialID = ' + STR(@ClinicalTrialID)
		PRINT 'TrialSite = ' + @TrialSite
		PRINT 'PersonID = ' + STR(@PersonID)
		PRINT 'ResponseTaskID = ' + STR(@ResponseTaskID)
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
		END
		FETCH NEXT FROM C1 INTO @ClinicalTrialID, @TrialSite, @PersonID, @ResponseTaskID, @DatabaseTimestamp, @ResponseTimestamp
	END
	CLOSE C1
	DEALLOCATE C1
END
--//--
GO
--//--
UPDATE MACROCONTROL SET BUILDSUBVERSION = '50';
--//--
