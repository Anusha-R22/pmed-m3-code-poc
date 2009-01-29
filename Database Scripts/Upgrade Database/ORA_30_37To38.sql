--//-- SEPARATOR=--//--
ALTER TABLE TRIALSUBJECT ADD SUBJECTTIMESTAMP NUMBER(16,10);
--//--
ALTER TABLE TRIALSUBJECT ADD SUBJECTTIMESTAMP_TZ NUMBER(6);
--//--
INSERT INTO NewDBColumn (VersionMajor,VersionMinor,VersionRevision,TableName,ColumnName,ColumnOrder,DefaultValue,ChangeType,ColumnNumber) VALUES (3,0,38,'TRIALSUBJECT','SUBJECTTIMESTAMP',null,'#NULL#','NEWCOLUMN',null);
--//--
INSERT INTO NewDBColumn (VersionMajor,VersionMinor,VersionRevision,TableName,ColumnName,ColumnOrder,DefaultValue,ChangeType,ColumnNumber) VALUES (3,0,38,'TRIALSUBJECT','SUBJECTTIMESTAMP_TZ',null,'#NULL#','NEWCOLUMN',null);
--//--
create or replace procedure SP_MACRO_IMP_TRIALSUBJECT
--------------------------------------------------------------------------------
-- SP_MACRO_IMP_TRIALSUBJECT
--
-- 25/02/2003  Ronald Schravendeel, (C) Infermed 2003
--
-- Called by clsExchange.ImportPRD
--
-- Note: The return value is declared as the first parameter with added 'OUT'
-- This is needed to make the parameter list compatible/equivalent with the SQL
-- Server version of this stored procedure.
--
-- Revisions
-- NCJ 7 Mar 03 - Added SubjectTimeStamp fields
--------------------------------------------------------------------------------

(V_RETURNVALUE         OUT NUMBER,
 V_CLINICALTRIALID     NUMBER,
 V_TRIALSITE           VARCHAR2,
 V_PERSONID            NUMBER,
 V_DATEOFBIRTH         NUMBER,
 V_GENDER              VARCHAR2,
 V_LOCALIDENTIFIER1    VARCHAR2,
 V_LOCALIDENTIFIER2    VARCHAR2,
 V_PROFORMASTATE       LONG,
 V_TRIALSUBJECTSTATUS  NUMBER,
 V_CHANGED             NUMBER,
 V_LOCKSTATUS          NUMBER,
 V_IMPORTTIMESTAMP     NUMBER,
 V_SUBJECTGENDER       NUMBER,
 V_REGISTRATIONSTATUS  NUMBER,
 V_DISCREPANCYSTATUS   NUMBER,
 V_SDVSTATUS           NUMBER,
 V_NOTESTATUS          NUMBER,
 V_IMPORTTIMESTAMP_TZ  NUMBER,
 V_SEQUENCEID          NUMBER,
 V_SUBJECTTIMESTAMP    NUMBER,
 V_SUBJECTTIMESTAMP_TZ NUMBER
)
is
begin
  V_RETURNVALUE := 0;


    begin
         -- TRY INSERT FIRST
         insert into trialsubject
           (clinicaltrialid, trialsite, personid, dateofbirth, gender, localidentifier1, localidentifier2, proformastate, trialsubjectstatus, changed, lockstatus, importtimestamp, subjectgender, registrationstatus, discrepancystatus, sdvstatus, notestatus, importtimestamp_tz, sequenceid, subjecttimestamp, subjecttimestamp_tz)
         values
           (v_clinicaltrialid, v_trialsite, v_personid, v_dateofbirth, v_gender, v_localidentifier1, v_localidentifier2, v_proformastate, v_trialsubjectstatus, v_changed, v_lockstatus, v_importtimestamp, v_subjectgender, v_registrationstatus, v_discrepancystatus, v_sdvstatus, v_notestatus, v_importtimestamp_tz, v_sequenceid, v_subjecttimestamp, v_subjecttimestamp_tz);
          v_returnvalue := 1;
    exception
      -- AN ERROR OCCURRED DURING INSERT
      when dup_val_on_index then
          -- INSERT FAILED (PRIMARY KEY ALREADY PRESENT) DO UPDATE INSTEAD
          update trialsubject
             set clinicaltrialid = v_clinicaltrialid,
                 trialsite = v_trialsite,
                 personid = v_personid,
                 dateofbirth = v_dateofbirth,
                 gender = v_gender,
                 localidentifier1 = v_localidentifier1,
                 localidentifier2 = v_localidentifier2,
                 proformastate = v_proformastate,
                 trialsubjectstatus = v_trialsubjectstatus,
                 changed = v_changed,
                 lockstatus = v_lockstatus,
                 importtimestamp = v_importtimestamp,
                 subjectgender = v_subjectgender,
                 registrationstatus = v_registrationstatus,
                 -- FOLLOWING FIELDS ARE NOT UPDATED WHEN DOING AN IMPORT
                 -- discrepancystatus = v_discrepancystatus,
                 -- sdvstatus = v_sdvstatus,
                 -- notestatus = v_notestatus,
                 importtimestamp_tz = v_importtimestamp_tz,
                 sequenceid = v_sequenceid,
                 subjecttimestamp = v_subjecttimestamp,
                 subjecttimestamp_tz = v_subjecttimestamp_tz
           where clinicaltrialid = v_clinicaltrialid
             and trialsite = v_trialsite
             and personid = v_personid;
           v_returnvalue := 2;
      when others then
           -- OTHER ERROR OCCURRED
           v_returnvalue := 0;
      end;


end SP_MACRO_IMP_TRIALSUBJECT;
/
--//--

UPDATE MACROControl SET BUILDSUBVERSION = '38';
--//--
