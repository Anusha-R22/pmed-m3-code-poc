CREATE Table EFormQGroup(ClinicalTrialID Number(11),VersionID Number(6),CRFPageID Number(11),QGroupID Number(11),Border Number(6),DisplayRows Number(6),InitialRows Number(6),MinRepeats Number(6),MaxRepeats Number(6), CONSTRAINT PKEFormQuestionGroup PRIMARY KEY (ClinicalTrialID,VersionID,CRFPageID,QGroupID));
CREATE Table QGroup(ClinicalTrialID Number(11),VersionID Number(6),QGroupID Number(11),QGroupCode VarChar2(15),QGroupName VarChar2(255),DisplayType Number(6),CONSTRAINT PKQGroup PRIMARY KEY (ClinicalTrialID,VersionID,QGroupID));
CREATE Table QGroupQuestion(ClinicalTrialID Number(11),VersionID Number(6),QGroupID Number(11),DataItemID Number(11),QOrder Number(6),CONSTRAINT PKQGroupQuestion PRIMARY KEY (ClinicalTrialID,VersionID,QGroupID,DataItemID));
CREATE Table QGroupInstance(ClinicalTrialID Number(11),TrialSite VarChar2(8),PersonID Number(11),CRFPageTaskID Number(11),QGroupID Number(11),QGroupRows Number(6),QGroupStatus Number(6),LockStatus Number(3),Changed Number(6),ImportTimeStamp Number(16,10),CONSTRAINT PKQGroupInstance PRIMARY KEY (ClinicalTrialID,TrialSite,PersonID,CRFPageTaskID,QGroupID));

INSERT INTO MACROTable VALUES ('QGroupInstance','095',0,1,0);
INSERT INTO MACROTable VALUES ('QGroup','230',1,0,0);
INSERT INTO MACROTable VALUES ('QGroupQuestion','220',1,0,0);
INSERT INTO MACROTable VALUES ('EFormQGroup','210',1,0,0);

ALTER Table CRFElement ADD OwnerQGroupID NUMBER(11);
ALTER Table CRFElement ADD QGroupID NUMBER(11);
ALTER Table CRFElement ADD QGroupFieldOrder NUMBER(6);
ALTER Table CRFElement ADD ShowStatusFlag NUMBER(6);
ALTER Table DataItemResponse ADD RepeatNumber NUMBER(6);
ALTER Table DataItemResponseHistory ADD RepeatNumber NUMBER(6);
ALTER Table MIMessage ADD MIMessageResponseCycle NUMBER(6);

UPDATE CRFElement SET OwnerQGroupID = 0 WHERE OwnerQGroupID IS NULL;
UPDATE CRFElement SET QGroupID = 0 WHERE QGroupId IS NULL;
UPDATE CRFElement SET QGroupFieldOrder = 0 WHERE QGroupFieldOrder IS NULL;
UPDATE CRFElement SET ShowStatusFlag = 1 WHERE ShowStatusFlag IS NULL;
UPDATE DataItemResponse SET RepeatNumber = 1 WHERE RepeatNumber IS NULL;
UPDATE DataItemResponseHistory SET RepeatNumber = 1 WHERE RepeatNumber IS NULL;
UPDATE MIMessage SET MIMessageResponseCycle = 1 WHERE MIMessageResponseCycle IS NULL;

ALTER Table DataItemResponse DROP CONSTRAINT PKDataItemResponse;
ALTER TAble DataItemResponseHistory DROP CONSTRAINT PKDataItemResponseHistory;
ALTER Table DataItemResponse ADD CONSTRAINT PKDataItemResponse PRIMARY KEY(ClinicalTrialId,TrialSite,PersonId,ResponseTaskId,RepeatNumber);
ALTER Table DataItemResponseHistory ADD CONSTRAINT PKDataItemResponseHistory PRIMARY KEY(ClinicalTrialId,TrialSite,PersonId,ResponseTaskId,ResponseTimeStamp,RepeatNumber);

INSERT INTO NewDBColumn VALUES (3,0,3,'MIMessage','MIMessageResponseCycle',null,'1','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,3,'DataItemResponseHistory','RepeatNumber',null,'1','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,3,'DataItemResponse','RepeatNumber',null,'1','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,3,'CRFElement','ShowStatusFlag',4,'1','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,3,'CRFElement','QGroupFieldOrder',3,'0','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,3,'CRFElement','QGroupID',2,'0','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,3,'CRFElement','OwnerQGroupID',1,'0','NEWCOLUMN',null);

ALTER Table CRFElement ADD CaptionFontName VARCHAR2(50);
ALTER Table CRFElement ADD CaptionFontBold NUMBER(6);
ALTER Table CRFElement ADD CaptionFontItalic NUMBER(6);
ALTER Table CRFElement ADD CaptionFontSize NUMBER(6);
ALTER Table CRFElement ADD CaptionFontColour NUMBER(11);

INSERT INTO NewDBColumn VALUES (3,0,4,'CRFElement','CaptionFontName',1,'#NULL#','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,4,'CRFElement','CaptionFontBold',2,'#NULL#','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,4,'CRFElement','CaptionFontItalic',3,'#NULL#','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,4,'CRFElement','CaptionFontSize',4,'#NULL#','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,4,'CRFElement','CaptionFontColour',5,'#NULL#','NEWCOLUMN',null);

UPDATE CRFElement SET CaptionFontName = FontName, CaptionFontBold = FontBold, CaptionFontItalic = FontItalic, CaptionFontSize = FontSize, CaptionFontColour = FontColour WHERE DataItemId > 0 OR ControlType = 16386;

ALTER Table DataItemResponse ADD ChangeCount NUMBER (6);
ALTER Table DataItemResponse ADD DiscrepancyStatus NUMBER (6);
ALTER Table DataItemResponse ADD SDVStatus NUMBER (6);
ALTER Table DataItemResponse ADD Notestatus NUMBER (6);

Update DataItemResponse set NoteStatus = 0;
Update DataItemResponse set SDVStatus = 0;
Update DataItemResponse set DiscrepancyStatus = 0;
Update DataItemResponse set ChangeCount = HadValue;

ALTER Table CRFPageInstance ADD DiscrepancyStatus NUMBER (6);
ALTER Table CRFPageInstance ADD SDVStatus NUMBER (6);
ALTER Table CRFPageInstance ADD Notestatus NUMBER (6);

Update CRFPageInstance set NoteStatus = 0;
Update CRFPageInstance set SDVStatus = 0;
Update CRFPageInstance set DiscrepancyStatus = 0;

ALTER Table VisitInstance ADD DiscrepancyStatus NUMBER (6);
ALTER Table VisitInstance ADD SDVStatus NUMBER (6);
ALTER Table VisitInstance ADD Notestatus NUMBER (6);

Update VisitInstance set NoteStatus = 0;
Update VisitInstance set SDVStatus = 0;
Update VisitInstance set DiscrepancyStatus = 0;

ALTER Table TrialSubject ADD DiscrepancyStatus NUMBER (6);
ALTER Table TrialSubject ADD SDVStatus NUMBER (6);
ALTER Table TrialSubject ADD Notestatus NUMBER (6);

Update TrialSubject set NoteStatus = 0;
Update TrialSubject set SDVStatus = 0;
Update TrialSubject set DiscrepancyStatus = 0;

INSERT INTO NewDBColumn VALUES (3,0,5,'TrialSubject','NoteStatus',3,'0','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,5,'TrialSubject','SDVStatus',2,'0','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,5,'TrialSubject','DiscrepancyStatus',1,'0','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,5,'VisitInstance','NoteStatus',3,'0','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,5,'VisitInstance','SDVStatus',2,'0','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,5,'VisitInstance','DiscrepancyStatus',1,'0','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,5,'CRFPageInstance','NoteStatus',3,'0','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,5,'CRFPageInstance','SDVStatus',2,'0','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,5,'CRFPageInstance','DiscrepancyStatus',1,'0','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,5,'DataItemResponse','NoteStatus',4,'0','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,5,'DataItemResponse','SDVStatus',3,'0','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,5,'DataItemResponse','DiscrepancyStatus',2,'0','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,5,'DataItemResponse','ChangeCount',1,'0','NEWCOLUMN',null);

ALTER Table DataItem ADD MACROOnly NUMBER (6);
Update DataItem SET MACROOnly = 0;
ALTER Table DataItem ADD Description VARCHAR(255);

ALTER Table StudyDefinition ADD AREZZOUpdateStatus NUMBER (6);
Update StudyDefinition SET AREZZOUpdateStatus = 0 WHERE AREZZOUpdateStatus IS NULL;

INSERT INTO NewDBColumn VALUES (3,0,7,'StudyDefinition','AREZZOUpdateStatus',2,'0','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,7,'DataItem','Description',2,'#NULL#','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,7,'DataItem','MACROOnly',1,'0','NEWCOLUMN',null);

CREATE Table StudyVersion(ClinicalTrialId NUMBER(11) NOT NULL, StudyVersion NUMBER(6) NOT NULL, VersionTimeStamp NUMBER(16,10),VersionDescription VARCHAR2(255), VersionTimeStamp_TZ NUMBER(6), CONSTRAINT PKStudyVersion PRIMARY KEY (ClinicalTrialId,StudyVersion));
CREATE Table MACRODBSetting(SettingSection VARCHAR2(15) NOT NULL,SettingKey VARCHAR2(15) NOT NULL,SettingValue VARCHAR2(15), CONSTRAINT PKMACRODBSetting PRIMARY KEY (SettingSection,SettingKey));

INSERT INTO MACROTable VALUES ('MACRODBSetting','',0,0,0);
INSERT INTO MACROTable VALUES ('StudyVersion','',0,0,0);

ALTER Table Message ADD MessageReceivedTimeStamp NUMBER(16,10);
ALTER Table TrialSite ADD StudyVersion NUMBER(6);
ALTER Table Site ADD SiteLocation NUMBER(6);

INSERT INTO NewDBColumn VALUES (3,0,8,'TrialSite','StudyVersion',null,'#NULL#','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,8,'Site','SiteLocation',null,'#NULL#','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,8,'Message','MessageReceivedTimeStamp',null,'#NULL#','NEWCOLUMN',null);

INSERT INTO MACRODBSetting(SettingSection, SettingKey,SettingValue) VALUES ('datatransfer','dbtype','server');

ALTER Table StudyVisitCRFPage ADD EFormUse NUMBER (6);
ALTER Table CRFElement ADD ElementUse NUMBER (6);

UPDATE StudyVisitCRFPage SET eFormUse = 0 WHERE eFormUse IS NULL;
UPDATE CRFElement SET ElementUse = 0 WHERE ElementUse IS NULL;

ALTER TABLE StudyVisitCRFPage MODIFY eFormUse NUMBER (6) NOT NULL;
ALTER TABLE CRFElement MODIFY ElementUse NUMBER (6) NOT NULL;

INSERT INTO NewDBColumn VALUES (3,0,8,'StudyVisitCRFPage','eFormUse',null,'0','NEWCOLUMN',null);
INSERT INTO NewDBColumn VALUES (3,0,8,'CRFElement','ElementUse',null,'0','NEWCOLUMN',null);

Update TrialSubject set DiscrepancyStatus = ( SELECT MAX(DECODE(MIMessageStatus,0,30,1,20,2,10,0)) From MIMESSAGE, ClinicalTrial Where MIMessageType = 0 AND MIMessageHistory = 0 AND MIMessageTrialName = ClinicalTrial.ClinicalTrialName AND MIMessageSite = TrialSubject.TrialSite AND MIMessagePersonId = TrialSubject.PersonId AND ClinicalTrial.ClinicalTrialId = TrialSubject.ClinicalTrialId );
Update TrialSubject SET DiscrepancyStatus = 0 Where DiscrepancyStatus IS NULL;
Update VisitInstance set DiscrepancyStatus = ( SELECT MAX(DECODE(MIMessageStatus,0,30,1,20,2,10,0)) From MIMESSAGE, ClinicalTrial Where MIMessageType = 0 AND MIMessageHistory = 0 AND MIMessageTrialName = ClinicalTrial.ClinicalTrialName AND MIMessageSite = VisitInstance.TrialSite AND MIMessagePersonId = VisitInstance.PersonId AND MIMessageVisitId = VisitInstance.VisitId AND MIMessageVisitCycle = VisitInstance.VisitCycleNumber AND ClinicalTrial.ClinicalTrialId = VisitInstance.ClinicalTrialId );
Update VisitInstance SET DiscrepancyStatus = 0 Where DiscrepancyStatus IS NULL;
Update CRFPageInstance set DiscrepancyStatus = ( SELECT MAX(DECODE(MIMessageStatus,0,30,1,20,2,10,0)) From MIMESSAGE, ClinicalTrial Where MIMessageType = 0 AND MIMessageHistory = 0 AND MIMessageTrialName = ClinicalTrial.ClinicalTrialName AND MIMessageSite = CRFPageInstance.TrialSite AND MIMessagePersonId = CRFPageInstance.PersonId AND MIMessageCRFPageTaskId = CRFPageInstance.CRFPageTaskId AND ClinicalTrial.ClinicalTrialId = CRFPageInstance.ClinicalTrialId );
Update CRFPageInstance SET DiscrepancyStatus = 0 Where DiscrepancyStatus IS NULL;
Update DataItemResponse set DiscrepancyStatus = ( SELECT MAX(DECODE(MIMessageStatus,0,30,1,20,2,10,0)) From MIMESSAGE, ClinicalTrial Where MIMessageType = 0 AND MIMessageHistory = 0 AND MIMessageTrialName = ClinicalTrial.ClinicalTrialName AND MIMessageSite = DataItemResponse.TrialSite AND MIMessagePersonId = DataItemResponse.PersonId AND MIMessageResponseTaskId = DataItemResponse.ResponseTaskId AND MIMessageResponseCycle = DataItemResponse.RepeatNumber AND ClinicalTrial.ClinicalTrialId = DataItemResponse.ClinicalTrialId );
Update DataItemResponse SET DiscrepancyStatus = 0 Where DiscrepancyStatus IS NULL;
Update TrialSubject set SDVStatus = ( SELECT MAX(DECODE(MIMessageStatus,0,30,2,20,0)) From MIMESSAGE, ClinicalTrial Where MIMessageType = 3 AND MIMessageHistory = 0 AND MIMessageTrialName = ClinicalTrial.ClinicalTrialName AND MIMessageSite = TrialSubject.TrialSite AND MIMessagePersonId = TrialSubject.PersonId AND ClinicalTrial.ClinicalTrialId = TrialSubject.ClinicalTrialId );
Update TrialSubject SET SDVStatus = 0 Where SDVStatus IS NULL;
Update VisitInstance set SDVStatus = ( SELECT MAX(DECODE(MIMessageStatus,0,30,2,20,0)) From MIMESSAGE, ClinicalTrial Where MIMessageType = 3 AND MIMessageHistory = 0 AND MIMessageTrialName = ClinicalTrial.ClinicalTrialName AND MIMessageSite = VisitInstance.TrialSite AND MIMessagePersonId = VisitInstance.PersonId AND MIMessageVisitId = VisitInstance.VisitId AND MIMessageVisitCycle = VisitInstance.VisitCycleNumber AND ClinicalTrial.ClinicalTrialId = VisitInstance.ClinicalTrialId );
Update VisitInstance SET SDVStatus = 0 Where SDVStatus IS NULL;
Update CRFPageInstance set SDVStatus = ( SELECT MAX(DECODE(MIMessageStatus,0,30,2,20,0)) From MIMESSAGE, ClinicalTrial Where MIMessageType = 3 AND MIMessageHistory = 0 AND MIMessageTrialName = ClinicalTrial.ClinicalTrialName AND MIMessageSite = CRFPageInstance.TrialSite AND MIMessagePersonId = CRFPageInstance.PersonId AND MIMessageCRFPageTaskId = CRFPageInstance.CRFPageTaskId AND ClinicalTrial.ClinicalTrialId = CRFPageInstance.ClinicalTrialId );
Update CRFPageInstance SET SDVStatus = 0 Where SDVStatus IS NULL;
Update DataItemResponse set SDVStatus = ( SELECT MAX(DECODE(MIMessageStatus,0,30,2,20,0)) From MIMESSAGE, ClinicalTrial Where MIMessageType = 3 AND MIMessageHistory = 0 AND MIMessageTrialName = ClinicalTrial.ClinicalTrialName AND MIMessageSite = DataItemResponse.TrialSite AND MIMessagePersonId = DataItemResponse.PersonId AND MIMessageResponseTaskId = DataItemResponse.ResponseTaskId AND MIMessageResponseCycle = DataItemResponse.RepeatNumber AND ClinicalTrial.ClinicalTrialId = DataItemResponse.ClinicalTrialId );
Update DataItemResponse SET SDVStatus = 0 Where SDVStatus IS NULL;
Update DataItemResponse set NoteStatus = ( SELECT count(*) From MIMESSAGE, ClinicalTrial Where MIMessageType = 2 AND MIMessageTrialName = ClinicalTrial.ClinicalTrialName AND MIMessageSite = DataItemResponse.TrialSite AND MIMessagePersonId = DataItemResponse.PersonId AND MIMessageResponseTaskId = DataItemResponse.ResponseTaskId AND MIMessageResponseCycle = DataItemResponse.RepeatNumber AND ClinicalTrial.ClinicalTrialId = DataItemResponse.ClinicalTrialId );
Update DataItemResponse SET NoteStatus = 1 Where NoteStatus > 0;

ALTER TABLE TrialSubject MODIFY NoteStatus NUMBER (6) NOT NULL;
ALTER TABLE TrialSubject MODIFY SDVStatus NUMBER (6) NOT NULL;
ALTER TABLE TrialSubject MODIFY DiscrepancyStatus NUMBER (6) NOT NULL;
ALTER TABLE VisitInstance MODIFY NoteStatus NUMBER (6) NOT NULL;
ALTER TABLE VisitInstance MODIFY SDVStatus NUMBER (6) NOT NULL;
ALTER TABLE VisitInstance MODIFY DiscrepancyStatus NUMBER (6) NOT NULL;
ALTER TABLE CRFPageInstance MODIFY NoteStatus NUMBER (6) NOT NULL;
ALTER TABLE CRFPageInstance MODIFY SDVStatus NUMBER (6) NOT NULL;
ALTER TABLE CRFPageInstance MODIFY DiscrepancyStatus NUMBER (6) NOT NULL;
ALTER TABLE DataItemResponse MODIFY NoteStatus NUMBER (6) NOT NULL;
ALTER TABLE DataItemResponse MODIFY SDVStatus NUMBER (6) NOT NULL;
ALTER TABLE DataItemResponse MODIFY DiscrepancyStatus NUMBER (6) NOT NULL;
ALTER TABLE DataItemResponse MODIFY ChangeCount NUMBER (6) NOT NULL;

ALTER Table CRFElement ADD DisplayLength NUMBER (6);
ALTER Table StudyVisit ADD Repeating NUMBER (6);

INSERT into NewDBColumn (VersionMajor,VersionMinor,VersionRevision,TableName,ColumnName,ColumnOrder,DefaultValue,ChangeType,ColumnNumber) values (3,0,10,'StudyVisit','Repeating',1,'#NULL#','NEWCOLUMN',null);
INSERT into NewDbColumn (VersionMajor,VersionMinor,VersionRevision,TableName,ColumnName,ColumnOrder,DefaultValue,ChangeType,ColumnNumber) values (3,0,10,'CRFElement','DisplayLength',1,'#NULL#','NEWCOLUMN',null);

ALTER Table DataItemResponse ADD ResponseTimestamp_TZ NUMBER (6);
ALTER Table DataItemResponse ADD ImportTimestamp_TZ NUMBER (6);
ALTER Table DataItemResponse ADD DatabaseTimestamp DECIMAL(16,10);
ALTER Table QGroupInstance ADD ImportTimestamp_TZ NUMBER (6);
ALTER Table CRFPageInstance ADD ImportTimestamp_TZ NUMBER (6);
ALTER Table VisitInstance ADD ImportTimestamp_TZ NUMBER (6);
ALTER Table TrialSubject ADD ImportTimestamp_TZ NUMBER (6);

Update DataItemResponse Set DatabaseTimeStamp = 0 where DatabaseTimeStamp is null;

ALTER Table DataItemResponseHistory ADD ResponseTimestamp_TZ NUMBER (6);
ALTER Table DataItemResponseHistory ADD ImportTimestamp_TZ NUMBER (6);
ALTER Table DataItemResponseHistory ADD DatabaseTimestamp DECIMAL(16,10);
ALTER Table DataItemResponse ADD DatabaseTimestamp_TZ NUMBER (6);

Update DataItemResponseHistory Set DatabaseTimeStamp = 0 where DatabaseTimeStamp is null;

ALTER Table TrialStatusHistory ADD StatusChangedTimeStamp_TZ NUMBER (6);
ALTER Table RSUniquenessCheck ADD CheckTimestamp_TZ NUMBER (6);
ALTER Table Message ADD MessageTimestamp_TZ NUMBER (6);
ALTER Table Message ADD MessageReceivedTimestamp_TZ NUMBER (6);
ALTER Table MIMessage ADD MIMessageCreated_TZ NUMBER (6);
ALTER Table MIMessage ADD MIMessageSent_TZ NUMBER (6);
ALTER Table MIMessage ADD MIMessageReceived_TZ NUMBER (6);
ALTER Table DataItemResponseHistory ADD DatabaseTimestamp_TZ NUMBER (6);

INSERT INTO NewDBColumn VALUES (3,0,13,'TrialStatusHistory','StatusChangedTimeStamp_TZ',null,'#NULL#','NEWCOLUMN',null);
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
ALTER Table ReasonForChange ADD ReasonType NUMBER (6);
UPDATE ReasonForChange SET ReasonType = 0 WHERE ReasonType IS NULL;
INSERT INTO NewDBColumn VALUES (3,0,14,'ReasonForChange','ReasonType',1,'0','NEWCOLUMN',null);

CREATE Table MACROCountry(CountryId NUMBER(6) NOT NULL, CountryDescription VARCHAR2(50), CONSTRAINT PKMACROCountry PRIMARY KEY (CountryId));
CREATE Table MACROTimeZone(TZId NUMBER(6), Description VARCHAR2(100), OffsetMins NUMBER(6), CONSTRAINT PKMACROTimeZone PRIMARY KEY (TZId));

INSERT INTO MACROTable VALUES ('MACROTimeZone','',0,0,0);
INSERT INTO MACROTable VALUES ('MACROCountry','',0,0,0);

ALTER TABLE LogDetails ADD LogDateTime_TZ NUMBER (6);
ALTER TABLE LogDetails ADD Location VARCHAR2(50);
ALTER TABLE Site ADD SiteTimeZone NUMBER (6);
ALTER TABLE Site ADD SiteCountry NUMBER (6);
ALTER TABLE Site ADD SiteLocale NUMBER (6);

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

CREATE BITMAP INDEX IDX_DIRH_CHANGED ON DATAITEMRESPONSEHISTORY(CHANGED);
CREATE  INDEX IDX_DIRH_PERSONID ON DATAITEMRESPONSEHISTORY(PERSONID);
CREATE BITMAP INDEX IDX_DIRH_RESPONSESTATUS ON DATAITEMRESPONSEHISTORY(RESPONSESTATUS);
CREATE  INDEX IDX_DIRH_RESPONSETIMESTAMP ON DATAITEMRESPONSEHISTORY(RESPONSETIMESTAMP);
CREATE BITMAP INDEX IDX_DIR_CHANGED ON DATAITEMRESPONSE(CHANGED);
CREATE  INDEX IDX_DIR_PERSONID ON DATAITEMRESPONSE(PERSONID);
CREATE BITMAP INDEX IDX_DIR_RESPONSESTATUS ON DATAITEMRESPONSE(RESPONSESTATUS);
CREATE  INDEX IDX_DIR_RESPONSETIMESTAMP ON DATAITEMRESPONSE(RESPONSETIMESTAMP);

UPDATE MACROControl SET MACROVERSION = '3.0', BUILDSUBVERSION = '18';