--//-- SEPARATOR=--//--
CREATE TRIGGER SetDBTimeStamp_DIR ON DataItemResponse
FOR INSERT, UPDATE
AS
BEGIN
	DECLARE C1 CURSOR FOR SELECT ClinicalTrialID,TrialSite,PersonID,ResponseTaskID,DatabaseTimeStamp  from Inserted;
	DECLARE @ClinicalTrialID int
	DECLARE @TrialSite varchar(15)
	DECLARE @PersonID int
	DECLARE @ResponseTaskID int
	DECLARE @DatabaseTimeStamp numeric(16,10)
	DECLARE @Timezone varchar(15)
	
	-- Get Timezone value from MACRODBSettingsTable
	set @Timezone = (select settingvalue from macrodbsetting where settingsection='timezone' and settingkey='dbtz')
	PRINT 'Timezone = ' + @Timezone
	IF @Timezone IS NULL
	BEGIN
		SET @Timezone = '0'
	END

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
CREATE TRIGGER SetDBTimeStamp_DIRH ON DataItemResponseHistory
FOR INSERT, UPDATE
AS
BEGIN
	DECLARE C1 CURSOR FOR SELECT ClinicalTrialID,TrialSite,PersonID,ResponseTaskID,DatabaseTimeStamp,ResponseTimestamp  from Inserted;
	DECLARE @ClinicalTrialID int
	DECLARE @TrialSite varchar(15)
	DECLARE @PersonID int
	DECLARE @ResponseTaskID int
	DECLARE @DatabaseTimeStamp numeric(16,10)
	DECLARE @ResponseTimeStamp numeric(16,10)
	DECLARE @Timezone varchar(15)
	
	-- Get Timezone value from MACRODBSettingsTable
	set @Timezone = (select settingvalue from macrodbsetting where settingsection='timezone' and settingkey='dbtz')
	PRINT 'Timezone = ' + @Timezone
	IF @Timezone IS NULL
	BEGIN
		SET @Timezone = '0'
	END
	
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
ALTER TABLE DATAITEMRESPONSE ADD USERNAMEFULL VARCHAR(100);
--//--
ALTER TABLE DATAITEMRESPONSEHISTORY ADD USERNAMEFULL VARCHAR(100);
--//--
INSERT INTO NewDBColumn (VersionMajor,VersionMinor,VersionRevision,TableName,ColumnName,ColumnOrder,DefaultValue,ChangeType,ColumnNumber) VALUES (3,0,32,'DATAITEMRESPONSE','USERNAMEFULL',null,'#NULL#','NEWCOLUMN',null);
--//--
INSERT INTO NewDBColumn (VersionMajor,VersionMinor,VersionRevision,TableName,ColumnName,ColumnOrder,DefaultValue,ChangeType,ColumnNumber) VALUES (3,0,32,'DATAITEMRESPONSEHISTORY','USERNAMEFULL',null,'#NULL#','NEWCOLUMN',null);
--//--
CREATE TABLE BATCHRESPONSEDATA ( BATCHRESPONSEID INTEGER, CLINICALTRIALID INTEGER, SITE VARCHAR(8), PERSONID INTEGER, SUBJECTLABEL VARCHAR(50), VISITID INTEGER, VISITCYCLENUMBER SMALLINT, VISITCYCLEDATE NUMERIC(16,10), CRFPAGEID INTEGER, CRFPAGECYCLENUMBER SMALLINT, CRFPAGECYCLEDATE NUMERIC(16,10), DATAITEMID INTEGER, REPEATNUMBER SMALLINT, RESPONSE VARCHAR(255), USERNAME VARCHAR(20), UPLOADMESSAGE VARCHAR(255), CONSTRAINT PKBATCHRESPONSEDATA PRIMARY KEY (BATCHRESPONSEID));
--//--
INSERT INTO MACROTable (TableName,SegmentId,STYDEF,PATRSP,LABDEF) VALUES ('BATCHRESPONSEDATA','',0,0,0);
--//--
UPDATE MACROControl SET BUILDSUBVERSION = '33';
--//--