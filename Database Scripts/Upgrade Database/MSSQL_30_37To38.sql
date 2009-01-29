--//--SEPARATOR=--//--
ALTER TABLE TRIALSUBJECT ADD SUBJECTTIMESTAMP DECIMAL(16,10);
--//--
ALTER TABLE TRIALSUBJECT ADD SUBJECTTIMESTAMP_TZ SMALLINT;
--//--
GO
--//--
INSERT INTO NewDBColumn (VersionMajor,VersionMinor,VersionRevision,TableName,ColumnName,ColumnOrder,DefaultValue,ChangeType,ColumnNumber) VALUES (3,0,38,'TRIALSUBJECT','SUBJECTTIMESTAMP',null,'#NULL#','NEWCOLUMN',null);
--//--
INSERT INTO NewDBColumn (VersionMajor,VersionMinor,VersionRevision,TableName,ColumnName,ColumnOrder,DefaultValue,ChangeType,ColumnNumber) VALUES (3,0,38,'TRIALSUBJECT','SUBJECTTIMESTAMP_TZ',null,'#NULL#','NEWCOLUMN',null);
--//--
DROP PROCEDURE SP_MACRO_IMP_TRIALSUBJECT
--//--
GO
--//--
CREATE PROCEDURE SP_MACRO_IMP_TRIALSUBJECT
	(
	@clinicaltrialid  int,
	@trialsite  varchar(8),
	@personid  int,
	@dateofbirth double precision, 
	@gender varchar(1), 
	@localidentifier1 varchar(50), 
	@localidentifier2 varchar(50), 
	@proformastate text, 
	@trialsubjectstatus smallint, 
	@changed smallint, 
	@lockstatus smallint, 
	@importtimestamp double precision, 
	@subjectgender smallint, 
	@registrationstatus smallint, 
	@discrepancystatus smallint, 
	@sdvstatus smallint, 
	@notestatus smallint, 
	@importtimestamp_tz smallint, 
	@sequenceid	double precision,
	@subjecttimestamp double precision,
	@subjecttimestamp_tz smallint 
	)
AS
 SET NOCOUNT ON
 
 DECLARE @ErrNumber Int
 DECLARE @RecCount Int

	-- FIND OUT IF ROW EXISTS
 SELECT @RecCount =  (Select count(*) from TrialSubject where
     CLINICALTRIALID = @clinicaltrialid
     AND TRIALSITE = @trialsite
     AND PERSONID = @personid)

 
 -- INSERT LOGIC CODE HERE TO DETERMINE WHETHER TO INSERT OR UPDATE
  
 if @RecCount=0
  BEGIN
           insert into trialsubject
          (clinicaltrialid, trialsite, personid, dateofbirth, gender, localidentifier1, localidentifier2, proformastate, trialsubjectstatus, changed, lockstatus, importtimestamp, subjectgender, registrationstatus, discrepancystatus, sdvstatus, notestatus, importtimestamp_tz, sequenceid, subjecttimestamp, subjecttimestamp_tz)
         values
          (@clinicaltrialid, @trialsite, @personid, @dateofbirth, @gender, @localidentifier1, @localidentifier2, @proformastate, @trialsubjectstatus, @changed, @lockstatus, @importtimestamp, @subjectgender, @registrationstatus, @discrepancystatus, @sdvstatus, @notestatus, @importtimestamp_tz, @sequenceid, @subjecttimestamp, @subjecttimestamp_tz)
  RETURN 1
  END
 ELSE
  BEGIN
  update trialsubject
             set clinicaltrialid = @clinicaltrialid,
                 trialsite = @trialsite,
                 personid = @personid,
                 dateofbirth = @dateofbirth,
                 gender = @gender,
                 localidentifier1 = @localidentifier1,
                 localidentifier2 = @localidentifier2,
                 proformastate = @proformastate,
                 trialsubjectstatus = @trialsubjectstatus,
                 changed = @changed,
                 lockstatus = @lockstatus,
                 importtimestamp = @importtimestamp,
                 subjectgender = @subjectgender,
                 registrationstatus = @registrationstatus,
                 -- FOLLOWING FIELDS ARE NOT UPDATED WHEN DOING AN IMPORT
                 -- discrepancystatus = @discrepancystatus,
                 -- sdvstatus = @sdvstatus,
                 -- notestatus = @notestatus,
                 importtimestamp_tz = @importtimestamp_tz,
                 sequenceid = @sequenceid,
                 subjecttimestamp = @subjecttimestamp,
                 subjecttimestamp_tz = @subjecttimestamp_tz
           where clinicaltrialid = @clinicaltrialid
             and trialsite = @trialsite
             and personid = @personid
  
  RETURN 2
 END
--//--
GO
--//--
UPDATE MACROControl SET BUILDSUBVERSION = '38';
--//--
