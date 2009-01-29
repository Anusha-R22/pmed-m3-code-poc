--//--SEPARATOR=--//--
DROP PROCEDURE SP_MACRO_IMP_DIR
--//--
DROP PROCEDURE SP_MACRO_IMP_DIRH
--//--
GO
--//--
CREATE PROCEDURE SP_MACRO_IMP_DIR
(
 @clinicaltrialid  int,
 @trialsite  varchar(8),
 @personid  int,
 @responsetaskid int,
 @visitid   int,
 @crfpageid  int,
 @crfelementid  smallint,
 @dataitemid  int,
 @visitcyclenumber smallint,
 @crfpagecyclenumber smallint,
 @crfpagetaskid  int,
 @responsevalue varchar(255),
 @responsetimestamp double precision,
 @valuecode  varchar(15),
 @username  Varchar(20),
 @unitofmeasurement Varchar(15),
 @comments  Varchar(255),
 @responsestatus Smallint,
 @changed  Smallint,
 @softwareversion Varchar(15),
 @reasonforchange Varchar(255),
 @lockstatus  Tinyint,
 @importtimestamp double precision,
 @validationid  Smallint,
 @validationmessage Varchar(2000),
 @overrulereason Varchar(255),
 @labresult  Varchar(1),
 @ctcgrade  Smallint,
 @clinicaltestdate double precision,
 @laboratorycode Varchar(15),
 @hadvalue  Smallint,
 @repeatnumber  Smallint,
 @changecount  Smallint,
 @discrepancystatus Smallint,
 @sdvstatus  Smallint,
 @notestatus  Smallint,
 @responsetimestamp_tz Smallint,
 @importtimestamp_tz Smallint,
 @databasetimestamp double precision,
 @databasetimestamp_tz Smallint,
 @sequenceid  double precision,
 @standardvalue  double precision,
 @usernamefull  Varchar(100)
) AS
BEGIN
 SET NOCOUNT ON
 
 DECLARE @ErrNumber Int
 DECLARE @RecCount Int
 
 SELECT @RecCount =  (Select count(*) from DataItemResponse where
     CLINICALTRIALID = @clinicaltrialid
     AND TRIALSITE = @trialsite
     AND PERSONID = @personid
     AND RESPONSETASKID = @responsetaskid
     AND  REPEATNUMBER = @repeatnumber)
 
 -- INSERT LOGIC CODE HERE TO DETERMINE WHETHER TO INSERT OR UPDATE
 -- PROBABLY INSERT, AND CHECK FOR ERROR
 
 if @RecCount=0
  BEGIN
  insert into DataItemResponse
    (  CLINICALTRIALID,TRIALSITE,PERSONID,RESPONSETASKID,VISITID,CRFPAGEID,CRFELEMENTID,
     DATAITEMID,VISITCYCLENUMBER,CRFPAGECYCLENUMBER,CRFPAGETASKID,RESPONSEVALUE,
     RESPONSETIMESTAMP,VALUECODE,USERNAME,UNITOFMEASUREMENT,COMMENTS,RESPONSESTATUS,
     CHANGED,SOFTWAREVERSION,REASONFORCHANGE,LOCKSTATUS,IMPORTTIMESTAMP,VALIDATIONID,
     VALIDATIONMESSAGE,OVERRULEREASON,LABRESULT,CTCGRADE,CLINICALTESTDATE,
     LABORATORYCODE,HADVALUE,REPEATNUMBER,CHANGECOUNT,DISCREPANCYSTATUS,
     SDVSTATUS,NOTESTATUS,RESPONSETIMESTAMP_TZ,IMPORTTIMESTAMP_TZ,DATABASETIMESTAMP,
     DATABASETIMESTAMP_TZ,SEQUENCEID,STANDARDVALUE,USERNAMEFULL
     )    VALUES
    (  @clinicaltrialid,@trialsite,@personid,@responsetaskid,@visitid,@crfpageid,@crfelementid,
     @dataitemid,@visitcyclenumber,@crfpagecyclenumber,@crfpagetaskid,@responsevalue,
     @responsetimestamp,@valuecode,@username,@unitofmeasurement,@comments,@responsestatus,
     @changed,@softwareversion,@reasonforchange,@lockstatus,@importtimestamp,@validationid,
     @validationmessage,@overrulereason,@labresult,@ctcgrade,@clinicaltestdate,@laboratorycode,
     @hadvalue,@repeatnumber,@changecount,@discrepancystatus,@sdvstatus,@notestatus,
     @responsetimestamp_tz,@importtimestamp_tz,@databasetimestamp,@databasetimestamp_tz,
     @sequenceid,@standardvalue,@usernamefull
    )
  return 1
  END
 ELSE
  BEGIN
  UPDATE DataItemResponse
  SET VISITID = @visitid,
   CRFPAGEID = @crfpageid,
   CRFELEMENTID = @crfelementid,
   DATAITEMID = @dataitemid,
   VISITCYCLENUMBER = @visitcyclenumber,
   CRFPAGECYCLENUMBER = @crfpagecyclenumber,
   CRFPAGETASKID = @crfpagetaskid,
   RESPONSEVALUE = @responsevalue,
   RESPONSETIMESTAMP = @responsetimestamp,
   VALUECODE = @valuecode,
   USERNAME = @username,
   UNITOFMEASUREMENT = @unitofmeasurement,
   COMMENTS = @comments,
   RESPONSESTATUS = @responsestatus,
   CHANGED = @changed,
   SOFTWAREVERSION = @softwareversion,
   REASONFORCHANGE = @reasonforchange,
   LOCKSTATUS = @lockstatus,
   IMPORTTIMESTAMP = @importtimestamp,
   VALIDATIONID = @validationid,
   VALIDATIONMESSAGE = @validationmessage,
   OVERRULEREASON = @overrulereason,
   LABRESULT = @labresult,
   CTCGRADE = @ctcgrade,
   CLINICALTESTDATE = @clinicaltestdate,
   LABORATORYCODE = @laboratorycode,
   HADVALUE = @hadvalue,
   CHANGECOUNT = @changecount,
   DISCREPANCYSTATUS = @discrepancystatus,
   SDVSTATUS = @sdvstatus,
   NOTESTATUS = @notestatus,
   RESPONSETIMESTAMP_TZ = @responsetimestamp_tz,
   IMPORTTIMESTAMP_TZ = @importtimestamp_tz,
   DATABASETIMESTAMP = @databasetimestamp,
   DATABASETIMESTAMP_TZ = @databasetimestamp_tz,
   SEQUENCEID = @sequenceid,
   STANDARDVALUE = @standardvalue,
   USERNAMEFULL = @usernamefull
  WHERE
   CLINICALTRIALID = @clinicaltrialid
  AND TRIALSITE = @trialsite
  AND PERSONID = @personid
  AND RESPONSETASKID = @responsetaskid
  AND  REPEATNUMBER = @repeatnumber
  
  RETURN 2
 END
 
END
--//--
GO
--//-- 
CREATE PROCEDURE SP_MACRO_IMP_DIRH
(
 @clinicaltrialid  int,
 @trialsite  varchar(8),
 @personid  int,
 @responsetaskid int,
 @responsetimestamp double precision,
 @visitid   int,
 @crfpageid  int,
 @crfelementid  smallint,
 @dataitemid  int,
 @visitcyclenumber smallint,
 @crfpagecyclenumber smallint,
 @crfpagetaskid  int,
 @responsevalue varchar(255),
 @valuecode  varchar(15),
 @username  Varchar(20),
 @unitofmeasurement Varchar(15),
 @comments  Varchar(255),
 @responsestatus Smallint,
 @changed  Smallint,
 @softwareversion Varchar(15),
 @reasonforchange Varchar(255),
 @lockstatus  Tinyint,
 @importtimestamp double precision,
 @validationid  Smallint,
 @validationmessage Varchar(2000),
 @overrulereason Varchar(255),
 @labresult  Varchar(1),
 @ctcgrade  Smallint,
 @clinicaltestdate double precision,
 @laboratorycode Varchar(15),
 @hadvalue  Smallint,
 @repeatnumber  Smallint,
 @responsetimestamp_tz Smallint,
 @importtimestamp_tz Smallint,
 @databasetimestamp double precision,
 @databasetimestamp_tz Smallint,
 @sequenceid  double precision,
 @standardvalue  double precision,
 @usernamefull  Varchar(100)
) AS
BEGIN
 SET NOCOUNT ON
 
 DECLARE @ErrNumber Int
 DECLARE @RecCount Int
 
 SELECT @RecCount =  (Select count(*) from DataItemResponseHistory where
     CLINICALTRIALID = @clinicaltrialid
     AND TRIALSITE = @trialsite
     AND PERSONID = @personid
     AND RESPONSETASKID = @responsetaskid
     AND RESPONSETIMESTAMP = @responsetimestamp
     AND  REPEATNUMBER = @repeatnumber)
 
 
 IF @RecCount=0
 BEGIN
  -- NO Record with this primary key found: Insert New
  insert into DataItemResponseHistory
    (  CLINICALTRIALID,TRIALSITE,PERSONID,RESPONSETASKID,VISITID,CRFPAGEID,CRFELEMENTID,
     DATAITEMID,VISITCYCLENUMBER,CRFPAGECYCLENUMBER,CRFPAGETASKID,RESPONSEVALUE,
     RESPONSETIMESTAMP,VALUECODE,USERNAME,UNITOFMEASUREMENT,COMMENTS,RESPONSESTATUS,
     CHANGED,SOFTWAREVERSION,REASONFORCHANGE,LOCKSTATUS,IMPORTTIMESTAMP,VALIDATIONID,
     VALIDATIONMESSAGE,OVERRULEREASON,LABRESULT,CTCGRADE,CLINICALTESTDATE,
     LABORATORYCODE,HADVALUE,REPEATNUMBER,RESPONSETIMESTAMP_TZ,IMPORTTIMESTAMP_TZ,DATABASETIMESTAMP,
     DATABASETIMESTAMP_TZ,SEQUENCEID,STANDARDVALUE,USERNAMEFULL
     )    VALUES
    (  @clinicaltrialid,@trialsite,@personid,@responsetaskid,@visitid,@crfpageid,@crfelementid,
     @dataitemid,@visitcyclenumber,@crfpagecyclenumber,@crfpagetaskid,@responsevalue,
     @responsetimestamp,@valuecode,@username,@unitofmeasurement,@comments,@responsestatus,
     @changed,@softwareversion,@reasonforchange,@lockstatus,@importtimestamp,@validationid,
     @validationmessage,@overrulereason,@labresult,@ctcgrade,@clinicaltestdate,@laboratorycode,
     @hadvalue,@repeatnumber,
     @responsetimestamp_tz,@importtimestamp_tz,@databasetimestamp,@databasetimestamp_tz,
     @sequenceid,@standardvalue,@usernamefull
    )
  return 1
 END
 ELSE
 BEGIN
  -- NO Record with this primary key found: Update Existing
  UPDATE DataItemResponse
  SET VISITID = @visitid,
   CRFPAGEID = @crfpageid,
   CRFELEMENTID = @crfelementid,
   DATAITEMID = @dataitemid,
   VISITCYCLENUMBER = @visitcyclenumber,
   CRFPAGECYCLENUMBER = @crfpagecyclenumber,
   CRFPAGETASKID = @crfpagetaskid,
   RESPONSEVALUE = @responsevalue,
   VALUECODE = @valuecode,
   USERNAME = @username,
   UNITOFMEASUREMENT = @unitofmeasurement,
   COMMENTS = @comments,
   RESPONSESTATUS = @responsestatus,
   CHANGED = @changed,
   SOFTWAREVERSION = @softwareversion,
   REASONFORCHANGE = @reasonforchange,
   LOCKSTATUS = @lockstatus,
   IMPORTTIMESTAMP = @importtimestamp,
   VALIDATIONID = @validationid,
   VALIDATIONMESSAGE = @validationmessage,
   OVERRULEREASON = @overrulereason,
   LABRESULT = @labresult,
   CTCGRADE = @ctcgrade,
   CLINICALTESTDATE = @clinicaltestdate,
   LABORATORYCODE = @laboratorycode,
   HADVALUE = @hadvalue,
   RESPONSETIMESTAMP_TZ = @responsetimestamp_tz,
   IMPORTTIMESTAMP_TZ = @importtimestamp_tz,
   DATABASETIMESTAMP = @databasetimestamp,
   DATABASETIMESTAMP_TZ = @databasetimestamp_tz,
   SEQUENCEID = @sequenceid,
   STANDARDVALUE = @standardvalue,
   USERNAMEFULL = @usernamefull
  WHERE
   CLINICALTRIALID = @clinicaltrialid
  AND TRIALSITE = @trialsite
  AND PERSONID = @personid
  AND RESPONSETASKID = @responsetaskid
  AND RESPONSETIMESTAMP = @responsetimestamp
  AND  REPEATNUMBER = @repeatnumber
  
  RETURN 2
 

 END
 
END
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
	@sequenceid	double precision
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
          (clinicaltrialid, trialsite, personid, dateofbirth, gender, localidentifier1, localidentifier2, proformastate, trialsubjectstatus, changed, lockstatus, importtimestamp, subjectgender, registrationstatus, discrepancystatus, sdvstatus, notestatus, importtimestamp_tz, sequenceid)
         values
          (@clinicaltrialid, @trialsite, @personid, @dateofbirth, @gender, @localidentifier1, @localidentifier2, @proformastate, @trialsubjectstatus, @changed, @lockstatus, @importtimestamp, @subjectgender, @registrationstatus, @discrepancystatus, @sdvstatus, @notestatus, @importtimestamp_tz, @sequenceid)
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
                 sequenceid = @sequenceid
           where clinicaltrialid = @clinicaltrialid
             and trialsite = @trialsite
             and personid = @personid
  
  RETURN 2
 END
--//--
GO
--//--
CREATE PROCEDURE SP_MACRO_IMP_VISITINSTANCE
	(
	@clinicaltrialid  int,
	@trialsite  varchar(8),
	@personid  int,
	@visittaskid int,
	@visitid int,
	@visitcyclenumber smallint,
	@visitdate double precision,
	@visitstatus smallint,
	@changed smallint,
	@lockstatus tinyint,
	@importtimestamp double precision,
	@discrepancystatus smallint,
	@sdvstatus smallint,
	@notestatus smallint,
	@importtimestamp_tz smallint,
	@sequenceid double precision
	)
AS
 SET NOCOUNT ON
 
 DECLARE @ErrNumber Int
 DECLARE @RecCount Int

	-- FIND OUT IF ROW EXISTS
 SELECT @RecCount =  (Select count(*) from VisitInstance where
     CLINICALTRIALID = @clinicaltrialid
     AND TRIALSITE = @trialsite
     AND PERSONID = @personid
     AND VISITTASKID = @visittaskid)

 
 -- INSERT LOGIC CODE HERE TO DETERMINE WHETHER TO INSERT OR UPDATE
  
 if @RecCount=0
  BEGIN
         insert into visitinstance
           (clinicaltrialid, trialsite, personid, visittaskid, visitid, visitcyclenumber, visitdate, visitstatus, changed, lockstatus, importtimestamp, discrepancystatus, sdvstatus, notestatus, importtimestamp_tz, sequenceid)
         values
           (@clinicaltrialid, @trialsite, @personid, @visittaskid, @visitid, @visitcyclenumber, @visitdate, @visitstatus, @changed, @lockstatus, @importtimestamp, @discrepancystatus, @sdvstatus, @notestatus, @importtimestamp_tz, @sequenceid)
  RETURN 1
  END
 ELSE
  BEGIN
          update visitinstance
             set clinicaltrialid = @clinicaltrialid,
                 trialsite = @trialsite,
                 personid = @personid,
                 visittaskid = @visittaskid,
                 visitid = @visitid,
                 visitcyclenumber = @visitcyclenumber,
                 visitdate = @visitdate,
                 visitstatus = @visitstatus,
                 changed = @changed,
                 lockstatus = @lockstatus,
                 importtimestamp = @importtimestamp,
                 discrepancystatus = @discrepancystatus,
                 sdvstatus = @sdvstatus,
                 notestatus = @notestatus,
                 importtimestamp_tz = @importtimestamp_tz,
                 sequenceid = @sequenceid
           where clinicaltrialid = @clinicaltrialid
             and trialsite = @trialsite
             and personid = @personid
             and visittaskid = @visittaskid
  
  RETURN 2
 END
--//--
GO
--//--
CREATE PROCEDURE SP_MACRO_IMP_CRFPAGEINSTANCE
	(
	@clinicaltrialid  int,
	@trialsite  varchar(8),
	@personid  int,	
	@crfpagetaskid	int,
	@visitid	int,
	@crfpageid	int,
	@visitcyclenumber	smallint,
	@crfpagecyclenumber	smallint,
	@crfpagedate	double precision,
	@crfpagestatus	smallint,
	@changed	smallint,
	@crfpageinstancelabel	varchar(255),
	@lockstatus	tinyint,
	@importtimestamp	double precision,
	@laboratorycode	varchar(15),
	@discrepancystatus	smallint,
	@sdvstatus	smallint,
	@notestatus	smallint,
	@importtimestamp_tz	smallint,
	@sequenceid	double precision
	)
AS
 SET NOCOUNT ON
 
 DECLARE @ErrNumber Int
 DECLARE @RecCount Int

	-- FIND OUT IF ROW EXISTS
 SELECT @RecCount =  (Select count(*) from CRFPageInstance where
     CLINICALTRIALID = @clinicaltrialid
     AND TRIALSITE = @trialsite
     AND PERSONID = @personid
     AND CRFPAGETASKID = @crfpagetaskid)

 
 -- INSERT LOGIC CODE HERE TO DETERMINE WHETHER TO INSERT OR UPDATE
  
 if @RecCount=0
  BEGIN
      insert into crfpageinstance
        (clinicaltrialid, trialsite, personid, crfpagetaskid, visitid, crfpageid, visitcyclenumber, crfpagecyclenumber, crfpagedate, crfpagestatus, changed, crfpageinstancelabel, lockstatus, importtimestamp, laboratorycode, discrepancystatus, sdvstatus, notestatus, importtimestamp_tz, sequenceid)
      values
        (@clinicaltrialid, @trialsite, @personid, @crfpagetaskid, @visitid, @crfpageid, @visitcyclenumber, @crfpagecyclenumber, @crfpagedate, @crfpagestatus, @changed, @crfpageinstancelabel, @lockstatus, @importtimestamp, @laboratorycode, @discrepancystatus, @sdvstatus, @notestatus, @importtimestamp_tz, @sequenceid)
  RETURN 1
  END
 ELSE
  BEGIN
      update crfpageinstance
         set visitid = @visitid,
             crfpageid = @crfpageid,
             visitcyclenumber = @visitcyclenumber,
             crfpagecyclenumber = @crfpagecyclenumber,
             crfpagedate = @crfpagedate,
             crfpagestatus = @crfpagestatus,
             changed = @changed,
             crfpageinstancelabel = @crfpageinstancelabel,
             lockstatus = @lockstatus,
             importtimestamp = @importtimestamp,
             laboratorycode = @laboratorycode,
             discrepancystatus = @discrepancystatus,
             sdvstatus = @sdvstatus,
             notestatus = @notestatus,
             importtimestamp_tz = @importtimestamp_tz,
             sequenceid = @sequenceid
       where clinicaltrialid = @clinicaltrialid
         and trialsite = @trialsite
         and personid = @personid
         and crfpagetaskid = @crfpagetaskid
  
  RETURN 2
 END
--//--
GO
--//--
UPDATE MACROCONTROL SET BUILDSUBVERSION = '37'
--//--

