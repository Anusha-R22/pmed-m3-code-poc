CREATE Table EFormQGroup(ClinicalTrialID Integer,VersionID SmallInt,CRFPageID Integer,QGroupID Integer,Border SmallInt,DisplayRows SmallInt,InitialRows SmallInt,MinRepeats SmallInt,MaxRepeats SmallInt, CONSTRAINT PKEFormQuestionGroup PRIMARY KEY (ClinicalTrialID,VersionID,CRFPageID,QGroupID));
GO
CREATE Table QGroup(ClinicalTrialID Integer,VersionID SmallInt,QGroupID Integer,QGroupCode VarChar(15),QGroupName VarChar(255),DisplayType SmallInt,CONSTRAINT PKQGroup PRIMARY KEY (ClinicalTrialID,VersionID,QGroupID));
GO
CREATE Table QGroupQuestion(ClinicalTrialID Integer,VersionID SmallInt,QGroupID Integer,DataItemID Integer,QOrder SmallInt,CONSTRAINT PKQGroupQuestion PRIMARY KEY (ClinicalTrialID,VersionID,QGroupID,DataItemID));
GO
CREATE Table QGroupInstance(ClinicalTrialID Integer,TrialSite VarChar(8),PersonID Integer,CRFPageTaskID Integer,QGroupID Integer,QGroupRows SmallInt,QGroupStatus SmallInt,LockStatus TinyInt,Changed SmallInt,ImportTimeStamp Numeric(16,10),CONSTRAINT PKQGroupInstance PRIMARY KEY (ClinicalTrialID,TrialSite,PersonID,CRFPageTaskID,QGroupID));
GO
INSERT INTO MACROTable (TableName,SegmentId,STYDEF,PATRSP,LABDEF) VALUES ('EFormQGroup','210',1,0,0);
INSERT INTO MACROTable (TableName,SegmentId,STYDEF,PATRSP,LABDEF) VALUES ('QGroupQuestion','220',1,0,0);
INSERT INTO MACROTable (TableName,SegmentId,STYDEF,PATRSP,LABDEF) VALUES ('QGroup','230',1,0,0);
INSERT INTO MACROTable (TableName,SegmentId,STYDEF,PATRSP,LABDEF) VALUES ('QGroupInstance','095',0,1,0);

ALTER Table CRFElement ADD OwnerQGroupID INTEGER;
GO
ALTER Table CRFElement ADD QGroupID INTEGER;
GO
ALTER Table CRFElement ADD QGroupFieldOrder SMALLINT;
GO
ALTER Table CRFElement ADD ShowStatusFlag SMALLINT;
GO
ALTER Table DataItemResponse ADD RepeatNumber SMALLINT;
GO
ALTER Table DataItemResponseHistory ADD RepeatNumber SMALLINT;
GO
ALTER Table MIMessage ADD MIMessageResponseCycle SMALLINT;
GO

UPDATE CRFElement SET OwnerQGroupID = 0 WHERE OwnerQGroupID IS NULL;
UPDATE CRFElement SET QGroupID = 0 WHERE QGroupId IS NULL;
UPDATE CRFElement SET QGroupFieldOrder = 0 WHERE QGroupFieldOrder IS NULL;
UPDATE CRFElement SET ShowStatusFlag = 1 WHERE ShowStatusFlag IS NULL;
UPDATE DataItemResponse SET RepeatNumber = 1 WHERE RepeatNumber IS NULL;
UPDATE DataItemResponseHistory SET RepeatNumber = 1 WHERE RepeatNumber IS NULL;
UPDATE MIMessage SET MIMessageResponseCycle = 1 WHERE MIMessageResponseCycle IS NULL;

ALTER TABLE DataItemResponse ALTER COLUMN RepeatNumber smallint NOT NULL;
GO
ALTER TABLE DataItemResponseHistory ALTER COLUMN RepeatNumber smallint NOT NULL;
GO
ALTER Table DataItemResponse DROP CONSTRAINT PKDataItemResponse;
GO
ALTER TAble DataItemResponseHistory DROP CONSTRAINT PKDataItemResponseHistory;
GO

ALTER Table DataItemResponse ADD CONSTRAINT PKDataItemResponse PRIMARY KEY(ClinicalTrialId,TrialSite,PersonId,ResponseTaskId,RepeatNumber);
GO
ALTER Table DataItemResponseHistory ADD CONSTRAINT PKDataItemResponseHistory PRIMARY KEY(ClinicalTrialId,TrialSite,PersonId,ResponseTaskId,ResponseTimeStamp,RepeatNumber);
GO

INSERT INTO NewDBColumn VALUES (3,0,3,'CRFElement','OwnerQGroupID',1,'0','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,3,'CRFElement','QGroupID',2,'0','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,3,'CRFElement','QGroupFieldOrder',3,'0','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,3,'CRFElement','ShowStatusFlag',4,'1','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,3,'DataItemResponse','RepeatNumber',null,'1','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,3,'DataItemResponseHistory','RepeatNumber',null,'1','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,3,'MIMessage','MIMessageResponseCycle',null,'1','NEWCOLUMN',null);

ALTER Table CRFElement ADD CaptionFontName VARCHAR(50);
GO
ALTER Table CRFElement ADD CaptionFontBold SMALLINT;
GO
ALTER Table CRFElement ADD CaptionFontItalic SMALLINT;
GO
ALTER Table CRFElement ADD CaptionFontSize SMALLINT;
GO
ALTER Table CRFElement ADD CaptionFontColour INTEGER;
GO

INSERT INTO NewDBColumn VALUES (3,0,4,'CRFElement','CaptionFontName',1,'#NULL#','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,4,'CRFElement','CaptionFontBold',2,'#NULL#','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,4,'CRFElement','CaptionFontItalic',3,'#NULL#','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,4,'CRFElement','CaptionFontSize',4,'#NULL#','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,4,'CRFElement','CaptionFontColour',5,'#NULL#','NEWCOLUMN',null);

UPDATE CRFElement SET CaptionFontName = FontName, CaptionFontBold = FontBold, CaptionFontItalic = FontItalic, CaptionFontSize = FontSize, CaptionFontColour = FontColour WHERE DataItemId > 0 OR ControlType = 16386;

ALTER Table DataItemResponse ADD ChangeCount SMALLINT;
GO
ALTER Table DataItemResponse ADD DiscrepancyStatus SMALLINT;
GO
ALTER Table DataItemResponse ADD SDVStatus SMALLINT;
GO
ALTER Table DataItemResponse ADD Notestatus SMALLINT;
GO

Update DataItemResponse set ChangeCount = HadValue;
Update DataItemResponse set DiscrepancyStatus = 0;
Update DataItemResponse set SDVStatus = 0;
Update DataItemResponse set NoteStatus = 0;

ALTER Table CRFPageInstance ADD DiscrepancyStatus SMALLINT;
GO
ALTER Table CRFPageInstance ADD SDVStatus SMALLINT;
GO
ALTER Table CRFPageInstance ADD Notestatus SMALLINT;
GO

Update CRFPageInstance set DiscrepancyStatus = 0;
Update CRFPageInstance set SDVStatus = 0;
Update CRFPageInstance set NoteStatus = 0;

ALTER Table VisitInstance ADD DiscrepancyStatus SMALLINT;
GO
ALTER Table VisitInstance ADD SDVStatus SMALLINT;
GO
ALTER Table VisitInstance ADD Notestatus SMALLINT;
GO

Update VisitInstance set DiscrepancyStatus = 0;
Update VisitInstance set SDVStatus = 0;
Update VisitInstance set NoteStatus = 0;

ALTER Table TrialSubject ADD DiscrepancyStatus SMALLINT;
GO
ALTER Table TrialSubject ADD SDVStatus SMALLINT;
GO
ALTER Table TrialSubject ADD Notestatus SMALLINT;
GO

Update TrialSubject set DiscrepancyStatus = 0;
Update TrialSubject set SDVStatus = 0;
Update TrialSubject set NoteStatus = 0;

INSERT INTO NewDBColumn VALUES (3,0,5,'DataItemResponse','ChangeCount',1,'0','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,5,'DataItemResponse','DiscrepancyStatus',2,'0','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,5,'DataItemResponse','SDVStatus',3,'0','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,5,'DataItemResponse','NoteStatus',4,'0','NEWCOLUMN',null);

INSERT INTO NewDBColumn VALUES (3,0,5,'CRFPageInstance','DiscrepancyStatus',1,'0','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,5,'CRFPageInstance','SDVStatus',2,'0','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,5,'CRFPageInstance','NoteStatus',3,'0','NEWCOLUMN',null);

INSERT INTO NewDBColumn VALUES (3,0,5,'VisitInstance','DiscrepancyStatus',1,'0','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,5,'VisitInstance','SDVStatus',2,'0','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,5,'VisitInstance','NoteStatus',3,'0','NEWCOLUMN',null);

INSERT INTO NewDBColumn VALUES (3,0,5,'TrialSubject','DiscrepancyStatus',1,'0','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,5,'TrialSubject','SDVStatus',2,'0','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,5,'TrialSubject','NoteStatus',3,'0','NEWCOLUMN',null);

ALTER Table DataItem ADD MACROOnly SMALLINT;
GO
Update DataItem SET MACROOnly = 0;
ALTER Table DataItem ADD Description VARCHAR(255);
GO
ALTER Table StudyDefinition ADD AREZZOUpdateStatus SMALLINT;
GO
Update StudyDefinition SET AREZZOUpdateStatus = 0 WHERE AREZZOUpdateStatus IS NULL;
ALTER Table StudyDefinition ALTER COLUMN AREZZOUpdateStatus SMALLINT NOT NULL;
GO

INSERT INTO NewDBColumn VALUES (3,0,7,'DataItem','MACROOnly',1,'0','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,7,'DataItem','Description',2,'#NULL#','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,7,'StudyDefinition','AREZZOUpdateStatus',2,'0','NEWCOLUMN',null);

CREATE Table StudyVersion(ClinicalTrialId INTEGER, StudyVersion SMALLINT, VersionTimeStamp NUMERIC(16,10),VersionDescription VARCHAR(255), VersionTimeStamp_TZ SMALLINT, CONSTRAINT PKStudyVersion PRIMARY KEY (ClinicalTrialId,StudyVersion));
GO
CREATE Table MACRODBSetting(SettingSection VARCHAR(15),SettingKey VARCHAR(15),SettingValue VARCHAR(15), CONSTRAINT PKMACRODBSetting PRIMARY KEY (SettingSection,SettingKey));
GO
INSERT INTO MACROTable (TableName,SegmentId,STYDEF,PATRSP,LABDEF) VALUES ('StudyVersion','',0,0,0);
INSERT INTO MACROTable (TableName,SegmentId,STYDEF,PATRSP,LABDEF) VALUES ('MACRODBSetting','',0,0,0);

ALTER Table Message ADD MessageReceivedTimeStamp NUMERIC(16,10);
GO
ALTER Table Site ADD SiteLocation SMALLINT;
GO
ALTER Table TrialSite ADD StudyVersion SMALLINT;
GO

INSERT INTO NewDBColumn VALUES (3,0,8,'Message','MessageReceivedTimeStamp',null,'#NULL#','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,8,'Site','SiteLocation',null,'#NULL#','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,8,'TrialSite','StudyVersion',null,'#NULL#','NEWCOLUMN',null);

INSERT INTO MACRODBSetting(SettingSection, SettingKey,SettingValue) VALUES ('datatransfer','dbtype','server');

ALTER Table CRFElement ADD ElementUse SMALLINT;
GO
ALTER Table StudyVisitCRFPage ADD EFormUse SMALLINT;
GO

UPDATE CRFElement SET ElementUse = 0 WHERE ElementUse IS NULL;
UPDATE StudyVisitCRFPage SET eFormUse = 0 WHERE eFormUse IS NULL;

ALTER TABLE CRFElement ALTER COLUMN ElementUse SMALLINT NOT NULL;
GO
ALTER TABLE StudyVisitCRFPage ALTER COLUMN eFormUse SMALLINT NOT NULL;
GO

INSERT INTO NewDBColumn VALUES (3,0,8,'CRFElement','ElementUse',null,'0','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,8,'StudyVisitCRFPage','eFormUse',null,'0','NEWCOLUMN',null);

Update TrialSubject set DiscrepancyStatus = ( SELECT MAX(CASE MIMessageStatus WHEN 0 THEN 30 WHEN 1 THEN 20 WHEN 2 THEN 10 ELSE 0 END) From MIMESSAGE, ClinicalTrial Where MIMessageType = 0 AND MIMessageHistory = 0 AND MIMessageTrialName = ClinicalTrial.ClinicalTrialName AND MIMessageSite = TrialSubject.TrialSite AND MIMessagePersonId = TrialSubject.PersonId AND ClinicalTrial.ClinicalTrialId = TrialSubject.ClinicalTrialId );
Update TrialSubject SET DiscrepancyStatus = 0 Where DiscrepancyStatus IS NULL;
Update VisitInstance set DiscrepancyStatus = ( SELECT MAX(CASE MIMessageStatus WHEN 0 THEN 30 WHEN 1 THEN 20 WHEN 2 THEN 10 ELSE 0 END) From MIMESSAGE, ClinicalTrial Where MIMessageType = 0 AND MIMessageHistory = 0 AND MIMessageTrialName = ClinicalTrial.ClinicalTrialName AND MIMessageSite = VisitInstance.TrialSite AND MIMessagePersonId = VisitInstance.PersonId AND MIMessageVisitId = VisitInstance.VisitId AND MIMessageVisitCycle = VisitInstance.VisitCycleNumber AND ClinicalTrial.ClinicalTrialId = VisitInstance.ClinicalTrialId );
Update VisitInstance SET DiscrepancyStatus = 0 Where DiscrepancyStatus IS NULL;
Update CRFPageInstance set DiscrepancyStatus = ( SELECT MAX(CASE MIMessageStatus WHEN 0 THEN 30 WHEN 1 THEN 20 WHEN 2 THEN 10 ELSE 0 END) From MIMESSAGE, ClinicalTrial Where MIMessageType = 0 AND MIMessageHistory = 0 AND MIMessageTrialName = ClinicalTrial.ClinicalTrialName AND MIMessageSite = CRFPageInstance.TrialSite AND MIMessagePersonId = CRFPageInstance.PersonId AND MIMessageCRFPageTaskId = CRFPageInstance.CRFPageTaskId AND ClinicalTrial.ClinicalTrialId = CRFPageInstance.ClinicalTrialId );
Update CRFPageInstance SET DiscrepancyStatus = 0 Where DiscrepancyStatus IS NULL;
Update DataItemResponse set DiscrepancyStatus = ( SELECT MAX(CASE MIMessageStatus WHEN 0 THEN 30 WHEN 1 THEN 20 WHEN 2 THEN 10 ELSE 0 END) From MIMESSAGE, ClinicalTrial Where MIMessageType = 0 AND MIMessageHistory = 0 AND MIMessageTrialName = ClinicalTrial.ClinicalTrialName AND MIMessageSite = DataItemResponse.TrialSite AND MIMessagePersonId = DataItemResponse.PersonId AND MIMessageResponseTaskId = DataItemResponse.ResponseTaskId AND MIMessageResponseCycle = DataItemResponse.RepeatNumber AND ClinicalTrial.ClinicalTrialId = DataItemResponse.ClinicalTrialId );
Update DataItemResponse SET DiscrepancyStatus = 0 Where DiscrepancyStatus IS NULL;

Update TrialSubject set SDVStatus = ( SELECT MAX(CASE MIMessageStatus WHEN 0 THEN 30 WHEN 2 THEN 20 ELSE 0 END) From MIMESSAGE, ClinicalTrial Where MIMessageType = 3 AND MIMessageHistory = 0 AND MIMessageTrialName = ClinicalTrial.ClinicalTrialName AND MIMessageSite = TrialSubject.TrialSite AND MIMessagePersonId = TrialSubject.PersonId AND ClinicalTrial.ClinicalTrialId = TrialSubject.ClinicalTrialId );
Update TrialSubject SET SDVStatus = 0 Where SDVStatus IS NULL;
Update VisitInstance set SDVStatus = ( SELECT MAX(CASE MIMessageStatus WHEN 0 THEN 30 WHEN 2 THEN 20 ELSE 0 END) From MIMESSAGE, ClinicalTrial Where MIMessageType = 3 AND MIMessageHistory = 0 AND MIMessageTrialName = ClinicalTrial.ClinicalTrialName AND MIMessageSite = VisitInstance.TrialSite AND MIMessagePersonId = VisitInstance.PersonId AND MIMessageVisitId = VisitInstance.VisitId AND MIMessageVisitCycle = VisitInstance.VisitCycleNumber AND ClinicalTrial.ClinicalTrialId = VisitInstance.ClinicalTrialId );
Update VisitInstance SET SDVStatus = 0 Where SDVStatus IS NULL;
Update CRFPageInstance set SDVStatus = ( SELECT MAX(CASE MIMessageStatus WHEN 0 THEN 30 WHEN 2 THEN 20 ELSE 0 END) From MIMESSAGE, ClinicalTrial Where MIMessageType = 3 AND MIMessageHistory = 0 AND MIMessageTrialName = ClinicalTrial.ClinicalTrialName AND MIMessageSite = CRFPageInstance.TrialSite AND MIMessagePersonId = CRFPageInstance.PersonId AND MIMessageCRFPageTaskId = CRFPageInstance.CRFPageTaskId AND ClinicalTrial.ClinicalTrialId = CRFPageInstance.ClinicalTrialId );
Update CRFPageInstance SET SDVStatus = 0 Where SDVStatus IS NULL;
Update DataItemResponse set SDVStatus = ( SELECT MAX(CASE MIMessageStatus WHEN 0 THEN 30 WHEN 2 THEN 20 ELSE 0 END) From MIMESSAGE, ClinicalTrial Where MIMessageType = 3 AND MIMessageHistory = 0 AND MIMessageTrialName = ClinicalTrial.ClinicalTrialName AND MIMessageSite = DataItemResponse.TrialSite AND MIMessagePersonId = DataItemResponse.PersonId AND MIMessageResponseTaskId = DataItemResponse.ResponseTaskId AND MIMessageResponseCycle = DataItemResponse.RepeatNumber AND ClinicalTrial.ClinicalTrialId = DataItemResponse.ClinicalTrialId );
Update DataItemResponse SET SDVStatus = 0 Where SDVStatus IS NULL;

Update DataItemResponse set NoteStatus = ( SELECT count(*) From MIMESSAGE, ClinicalTrial Where MIMessageType = 2 AND MIMessageTrialName = ClinicalTrial.ClinicalTrialName AND MIMessageSite = DataItemResponse.TrialSite AND MIMessagePersonId = DataItemResponse.PersonId AND MIMessageResponseTaskId = DataItemResponse.ResponseTaskId AND MIMessageResponseCycle = DataItemResponse.RepeatNumber AND ClinicalTrial.ClinicalTrialId = DataItemResponse.ClinicalTrialId );
Update DataItemResponse SET NoteStatus = 1 Where NoteStatus > 0;

ALTER TABLE DataItemResponse ALTER COLUMN ChangeCount SMALLINT NOT NULL;
GO
ALTER TABLE DataItemResponse ALTER COLUMN DiscrepancyStatus SMALLINT NOT NULL;
GO
ALTER TABLE DataItemResponse ALTER COLUMN SDVStatus SMALLINT NOT NULL;
GO
ALTER TABLE DataItemResponse ALTER COLUMN NoteStatus SMALLINT NOT NULL;
GO
ALTER TABLE CRFPageInstance ALTER COLUMN DiscrepancyStatus SMALLINT NOT NULL;
GO
ALTER TABLE CRFPageInstance ALTER COLUMN SDVStatus SMALLINT NOT NULL;
GO
ALTER TABLE CRFPageInstance ALTER COLUMN NoteStatus SMALLINT NOT NULL;
GO
ALTER TABLE VisitInstance ALTER COLUMN DiscrepancyStatus SMALLINT NOT NULL;
GO
ALTER TABLE VisitInstance ALTER COLUMN SDVStatus SMALLINT NOT NULL;
GO
ALTER TABLE VisitInstance ALTER COLUMN NoteStatus SMALLINT NOT NULL;
GO
ALTER TABLE TrialSubject ALTER COLUMN DiscrepancyStatus SMALLINT NOT NULL;
GO
ALTER TABLE TrialSubject ALTER COLUMN SDVStatus SMALLINT NOT NULL;
GO
ALTER TABLE TrialSubject ALTER COLUMN NoteStatus SMALLINT NOT NULL;
GO

ALTER Table CRFElement ADD DisplayLength SMALLINT;
GO
ALTER Table StudyVisit ADD Repeating SMALLINT;
GO

INSERT into NewDbColumn (VersionMajor,VersionMinor,VersionRevision,TableName,ColumnName,ColumnOrder,DefaultValue,ChangeType,ColumnNumber) values (3,0,10,'CRFElement','DisplayLength',1,'#NULL#','NEWCOLUMN',null);
INSERT into NewDBColumn (VersionMajor,VersionMinor,VersionRevision,TableName,ColumnName,ColumnOrder,DefaultValue,ChangeType,ColumnNumber) values (3,0,10,'StudyVisit','Repeating',1,'#NULL#','NEWCOLUMN',null);

ALTER Table DataItemResponse ADD ResponseTimestamp_TZ SMALLINT;
GO
ALTER Table DataItemResponse ADD ImportTimestamp_TZ SMALLINT;
GO
ALTER Table DataItemResponse ADD DatabaseTimestamp DECIMAL(16,10);
GO
ALTER Table QGroupInstance ADD ImportTimestamp_TZ SMALLINT;
GO
ALTER Table CRFPageInstance ADD ImportTimestamp_TZ SMALLINT;
GO
ALTER Table VisitInstance ADD ImportTimestamp_TZ SMALLINT;
GO
ALTER Table TrialSubject ADD ImportTimestamp_TZ SMALLINT;
GO

Update DataItemResponse Set DatabaseTimeStamp = 0 where DatabaseTimeStamp is null;

ALTER Table DataItemResponseHistory ADD ResponseTimestamp_TZ SMALLINT;
GO
ALTER Table DataItemResponseHistory ADD ImportTimestamp_TZ SMALLINT;
GO
ALTER Table DataItemResponseHistory ADD DatabaseTimestamp DECIMAL(16,10);
GO
ALTER Table DataItemResponse ADD DatabaseTimestamp_TZ SMALLINT;
GO

Update DataItemResponseHistory Set DatabaseTimeStamp = 0 where DatabaseTimeStamp is null;

ALTER Table TrialStatusHistory ADD StatusChangedTimeStamp_TZ SMALLINT;
GO
ALTER Table RSUniquenessCheck ADD CheckTimestamp_TZ SMALLINT;
GO
ALTER Table Message ADD MessageTimestamp_TZ SMALLINT;
GO
ALTER Table Message ADD MessageReceivedTimestamp_TZ SMALLINT;
GO
ALTER Table MIMessage ADD MIMessageCreated_TZ SMALLINT;
GO
ALTER Table MIMessage ADD MIMessageSent_TZ SMALLINT;
GO
ALTER Table MIMessage ADD MIMessageReceived_TZ SMALLINT;
GO
ALTER Table DataItemResponseHistory ADD DatabaseTimestamp_TZ SMALLINT;
GO

INSERT INTO NewDBColumn VALUES (3,0,13,'TrialStatusHistory','StatusChangedTimeStamp_TZ',null,'#NULL#','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,13,'StudyVersion','VersionTimestamp_TZ',null,'#NULL#','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,13,'RSUniquenessCheck','CheckTimestamp_TZ',null,'#NULL#','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,13,'Message','MessageReceivedTimestamp_TZ',null,'#NULL#','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,13,'Message','MessageTimestamp_TZ',null,'#NULL#','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,13,'MIMessage','MIMessageReceived_TZ',null,'#NULL#','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,13,'MIMessage','MIMessageSent_TZ',null,'#NULL#','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,13,'MIMessage','MIMessageCreated_TZ',null,'#NULL#','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,13,'DataItemResponseHistory','DatabaseTimestamp_TZ',null,'#NULL#','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,13,'DataItemResponseHistory','DatabaseTimestamp',null,'0','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,13,'DataItemResponseHistory','ImportTimestamp_TZ',null,'#NULL#','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,13,'DataItemResponseHistory','ResponseTimestamp_TZ',null,'#NULL#','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,13,'DataItemResponse','DatabaseTimestamp_TZ',null,'#NULL#','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,13,'DataItemResponse','DatabaseTimestamp',null,'0','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,13,'DataItemResponse','ImportTimestamp_TZ',null,'#NULL#','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,13,'DataItemResponse','ResponseTimestamp_TZ',null,'#NULL#','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,13,'QGroupInstance','ImportTimestamp_TZ',null,'#NULL#','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,13,'CRFPageInstance','ImportTimestamp_TZ',null,'#NULL#','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,13,'VisitInstance','ImportTimestamp_TZ',null,'#NULL#','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,13,'TrialSubject','ImportTimestamp_TZ',null,'#NULL#','NEWCOLUMN',null);

ALTER Table ReasonForChange DROP CONSTRAINT PKReasonForChange;
GO
ALTER Table ReasonForChange ADD ReasonType SMALLINT;
GO
UPDATE ReasonForChange SET ReasonType = 0 WHERE ReasonType IS NULL;
GO
INSERT INTO NewDBColumn VALUES (3,0,14,'ReasonForChange','ReasonType',1,'0','NEWCOLUMN',null);

CREATE Table MACROCountry(CountryId SMALLINT, CountryDescription VARCHAR(50), CONSTRAINT PKMACROCountry PRIMARY KEY (CountryId));
GO
CREATE Table MACROTimeZone(TZId SMALLINT, Description VARCHAR(100), OffsetMins SMALLINT, CONSTRAINT PKMACROTimeZone PRIMARY KEY (TZId));
GO

INSERT INTO MACROTable VALUES ('MACROTimeZone','',0,0,0);
INSERT INTO MACROTable VALUES ('MACROCountry','',0,0,0);

ALTER TABLE LogDetails ADD LogDateTime_TZ SMALLINT;
GO
ALTER TABLE LogDetails ADD Location VARCHAR(50);
GO
ALTER TABLE Site ADD SiteTimeZone SMALLINT;
GO
ALTER TABLE Site ADD SiteCountry SMALLINT;
GO
ALTER TABLE Site ADD SiteLocale SMALLINT;
GO

INSERT INTO NewDBColumn VALUES (3,0,18,'LogDetails','Location',null,'#NULL#','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,18,'LogDetails','LogDateTime_TZ',null,'#NULL#','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,18,'Site','SiteTimeZone',null,'#NULL#','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,18,'Site','SiteCountry',null,'#NULL#','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,18,'Site','SiteLocale',null,'#NULL#','NEWCOLUMN',null);

UPDATE LogDetails SET Location = 'Local' WHERE Location IS NULL;

INSERT INTO MACROCountry VALUES (2057,'United Kingdom');
INSERT INTO MACROCountry VALUES (2055,'Switzerland');
INSERT INTO MACROCountry VALUES (1053,'Sweden');
INSERT INTO MACROCountry VALUES (1045,'Poland');
INSERT INTO MACROCountry VALUES (1044,'Norway');
INSERT INTO MACROCountry VALUES (1043,'Netherlands');
INSERT INTO MACROCountry VALUES (1042,'Korea');
INSERT INTO MACROCountry VALUES (1041,'Japan');
INSERT INTO MACROCountry VALUES (1040,'Italy');
INSERT INTO MACROCountry VALUES (1039,'Ireland');
INSERT INTO MACROCountry VALUES (1038,'Hungary');
INSERT INTO MACROCountry VALUES (1036,'France');
INSERT INTO MACROCountry VALUES (1035,'Finland');
INSERT INTO MACROCountry VALUES (1034,'Spain');
INSERT INTO MACROCountry VALUES (1033,'United States');
INSERT INTO MACROCountry VALUES (1031,'Germany');
INSERT INTO MACROCountry VALUES (1030,'Denmark');
INSERT INTO MACROCountry VALUES (1026,'Bulgaria');

INSERT INTO MACROTimeZone VALUES (25,'(GMT-12:00) Eniwetok, Kwajalein',-720);
INSERT INTO MACROTimeZone VALUES (24,'(GMT-11:00) Midway Island, Samoa',-660);
INSERT INTO MACROTimeZone VALUES (23,'(GMT-10:00) Hawaii',-600);
INSERT INTO MACROTimeZone VALUES (22,'(GMT-09:00) Alaska',-540);
INSERT INTO MACROTimeZone VALUES (21,'(GMT-08:00) Pacific Time (US & Canada)',-480);
INSERT INTO MACROTimeZone VALUES (20,'(GMT-07:00) Mountain Time (US & Canada), Arizona',-420);
INSERT INTO MACROTimeZone VALUES (19,'(GMT-06:00) Central Time (US & Canada), Mexico City',-360);
INSERT INTO MACROTimeZone VALUES (18,'(GMT-05:00) Eastern Time (US & Canada)',-300);
INSERT INTO MACROTimeZone VALUES (17,'(GMT-04:00) Aantiago, Atlantic Time (Canada)',-240);
INSERT INTO MACROTimeZone VALUES (16,'(GMT-03:00) Greenland, Georgetown, Buenos Aires',-180);
INSERT INTO MACROTimeZone VALUES (15,'(GMT-02:00) Mid Atlantic',-120);
INSERT INTO MACROTimeZone VALUES (14,'(GMT-01:00) Cape Verde Is, Azores',-60);
INSERT INTO MACROTimeZone VALUES (13,'(GMT+12:00) Auckland, Wllington, Fiji',720);
INSERT INTO MACROTimeZone VALUES (12,'(GMT+11:00) Solomon Is, New Caledonia',660);
INSERT INTO MACROTimeZone VALUES (11,'(GMT+10:00) Brisbane, Canberra, Melbourne, Sydney',600);
INSERT INTO MACROTimeZone VALUES (10,'(GMT+09:00) Osaka, Tokyo, Seoul',540);
INSERT INTO MACROTimeZone VALUES (9,'(GMT+08:00) Beijing, Hong Kong, Singapore, Perth',480);
INSERT INTO MACROTimeZone VALUES (8,'(GMT+07:00) Bangkok, Hanoi, Jakarta',420);
INSERT INTO MACROTimeZone VALUES (7,'(GMT+06:00) Almaty, Novosibirsk',360);
INSERT INTO MACROTimeZone VALUES (6,'(GMT+05:00) Islamabad, Karachi, Tashkent',300);
INSERT INTO MACROTimeZone VALUES (5,'(GMT+04:00) Abu Dhabi, Muscat',240);
INSERT INTO MACROTimeZone VALUES (4,'(GMT+03:00) Moscow, St.Petersburg, Volgograd',180);
INSERT INTO MACROTimeZone VALUES (3,'(GMT+02:00) Ahtens, Istanbul, Minsk',120);
INSERT INTO MACROTimeZone VALUES (2,'(GMT+01:00) Amsterdam, Berlin, Bern, Rome, Stockholm, Vienna',60);
INSERT INTO MACROTimeZone VALUES (1,'(GMT) Greenwich Mean Time: Dublin, Edinburgh, Lisbon, London',0);

CREATE  INDEX IDX_DIRH_CHANGED ON DATAITEMRESPONSEHISTORY(CHANGED);
CREATE  INDEX IDX_DIRH_PERSONID ON DATAITEMRESPONSEHISTORY(PERSONID);
CREATE  INDEX IDX_DIRH_RESPONSESTATUS ON DATAITEMRESPONSEHISTORY(RESPONSESTATUS);
CREATE  INDEX IDX_DIRH_RESPONSETIMESTAMP ON DATAITEMRESPONSEHISTORY(RESPONSETIMESTAMP);
CREATE  INDEX IDX_DIR_PERSONID ON DATAITEMRESPONSE(PERSONID);
CREATE  INDEX IDX_DIR_RESPONSESTATUS ON DATAITEMRESPONSE(RESPONSESTATUS);
CREATE  INDEX IDX_DIR_RESPONSETIMESTAMP ON DATAITEMRESPONSE(RESPONSETIMESTAMP);

UPDATE MACROControl SET MACROVERSION = '3.0', BUILDSUBVERSION = '18';