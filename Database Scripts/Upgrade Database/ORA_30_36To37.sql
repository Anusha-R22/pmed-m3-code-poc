--//-- SEPARATOR=--//--
create or replace procedure SP_MACRO_IMP_DIR(V_RETURNVALUE OUT NUMBER,
                                             V_CLINICALTRIALID NUMBER,
                                             V_TRIALSITE VARCHAR2,
                                             V_PERSONID NUMBER,
                                             V_RESPONSETASKID NUMBER,
                                             V_VISITID NUMBER,
                                             V_CRFPAGEID NUMBER,
                                             V_CRFELEMENTID NUMBER,
                                             V_DATAITEMID NUMBER,
                                             V_VISITCYCLENUMBER NUMBER,
                                             V_CRFPAGECYCLENUMBER NUMBER,
                                             V_CRFPAGETASKID NUMBER,
                                             V_RESPONSEVALUE VARCHAR2,
                                             V_RESPONSETIMESTAMP NUMBER,
                                             V_VALUECODE VARCHAR2,
                                             V_USERNAME VARCHAR2,
                                             V_UNITOFMEASUREMENT VARCHAR2,
                                             V_COMMENTS VARCHAR2,
                                             V_RESPONSESTATUS NUMBER,
                                             V_CHANGED NUMBER,
                                             V_SOFTWAREVERSION VARCHAR2,
                                             V_REASONFORCHANGE VARCHAR2,
                                             V_LOCKSTATUS NUMBER,
                                             V_IMPORTTIMESTAMP NUMBER,
                                             V_VALIDATIONID NUMBER,
                                             V_VALIDATIONMESSAGE VARCHAR2,
                                             V_OVERRULEREASON VARCHAR2,
                                             V_LABRESULT VARCHAR2,
                                             V_CTCGRADE NUMBER,
                                             V_CLINICALTESTDATE NUMBER,
                                             V_LABORATORYCODE VARCHAR2,
                                             V_HADVALUE NUMBER,
                                             V_REPEATNUMBER NUMBER,
                                             V_CHANGECOUNT NUMBER,
                                             V_DISCREPANCYSTATUS NUMBER,
                                             V_SDVSTATUS NUMBER,
                                             V_NOTESTATUS NUMBER,
                                             V_RESPONSETIMESTAMP_TZ NUMBER,
                                             V_IMPORTTIMESTAMP_TZ NUMBER,
                                             V_DATABASETIMESTAMP NUMBER,
                                             V_DATABASETIMESTAMP_TZ NUMBER,
                                             V_SEQUENCEID NUMBER,
                                             V_STANDARDVALUE NUMBER,
                                             V_USERNAMEFULL VARCHAR2) is
begin

  -- EXCEPTION BLOCK: ATTEMPT INSERT, IF KEY VIOLATION, RECORD EXISTS, DO UPDATE INSTEAD
  begin
    -- STANDARD INSERT STATEMENT
    insert into dataitemresponse
      (clinicaltrialid, trialsite, personid, responsetaskid, visitid,
       crfpageid, crfelementid, dataitemid, visitcyclenumber,
       crfpagecyclenumber, crfpagetaskid, responsevalue, responsetimestamp,
       valuecode, username, unitofmeasurement, comments, responsestatus,
       changed, softwareversion, reasonforchange, lockstatus,
       importtimestamp, validationid, validationmessage, overrulereason,
       labresult, ctcgrade, clinicaltestdate, laboratorycode, hadvalue,
       repeatnumber, changecount, discrepancystatus, sdvstatus, notestatus,
       responsetimestamp_tz, importtimestamp_tz, databasetimestamp,
       databasetimestamp_tz, sequenceid, standardvalue, usernamefull)
    values
      (v_clinicaltrialid, v_trialsite, v_personid, v_responsetaskid,
       v_visitid, v_crfpageid, v_crfelementid, v_dataitemid,
       v_visitcyclenumber, v_crfpagecyclenumber, v_crfpagetaskid,
       v_responsevalue, v_responsetimestamp, v_valuecode, v_username,
       v_unitofmeasurement, v_comments, v_responsestatus, v_changed,
       v_softwareversion, v_reasonforchange, v_lockstatus, v_importtimestamp,
       v_validationid, v_validationmessage, v_overrulereason, v_labresult,
       v_ctcgrade, v_clinicaltestdate, v_laboratorycode, v_hadvalue,
       v_repeatnumber, v_changecount, v_discrepancystatus, v_sdvstatus,
       v_notestatus, v_responsetimestamp_tz, v_importtimestamp_tz,
       v_databasetimestamp, v_databasetimestamp_tz, v_sequenceid,
       v_standardvalue, v_usernamefull);
  
    v_returnvalue := 1;
  exception
    when dup_val_on_index then
      -- STANDARD UPDATE STATEMENT
      update dataitemresponse
         set clinicaltrialid = v_clinicaltrialid, trialsite = v_trialsite,
             personid = v_personid, responsetaskid = v_responsetaskid,
             visitid = v_visitid, crfpageid = v_crfpageid,
             crfelementid = v_crfelementid, dataitemid = v_dataitemid,
             visitcyclenumber = v_visitcyclenumber,
             crfpagecyclenumber = v_crfpagecyclenumber,
             crfpagetaskid = v_crfpagetaskid,
             responsevalue = v_responsevalue,
             responsetimestamp = v_responsetimestamp,
             valuecode = v_valuecode, username = v_username,
             unitofmeasurement = v_unitofmeasurement, comments = v_comments,
             responsestatus = v_responsestatus, changed = v_changed,
             softwareversion = v_softwareversion,
             reasonforchange = v_reasonforchange, lockstatus = v_lockstatus,
             importtimestamp = v_importtimestamp,
             validationid = v_validationid,
             validationmessage = v_validationmessage,
             overrulereason = v_overrulereason, labresult = v_labresult,
             ctcgrade = v_ctcgrade, clinicaltestdate = v_clinicaltestdate,
             laboratorycode = v_laboratorycode, hadvalue = v_hadvalue,
             repeatnumber = v_repeatnumber, changecount = v_changecount,
             discrepancystatus = v_discrepancystatus,
             sdvstatus = v_sdvstatus, notestatus = v_notestatus,
             responsetimestamp_tz = v_responsetimestamp_tz,
             importtimestamp_tz = v_importtimestamp_tz,
             databasetimestamp = v_databasetimestamp,
             databasetimestamp_tz = v_databasetimestamp_tz,
             sequenceid = v_sequenceid, standardvalue = v_standardvalue,
             usernamefull = v_usernamefull
       where clinicaltrialid = v_clinicaltrialid and
             trialsite = v_trialsite and personid = v_personid and
             responsetaskid = v_responsetaskid and
             repeatnumber = v_repeatnumber;
    
      v_returnvalue := 2;
    when others then
      v_returnvalue := 0;
  end;

end SP_MACRO_IMP_DIR;
/
--//--
create or replace procedure SP_MACRO_IMP_DIRH(V_RETURNVALUE OUT NUMBER,
                                              V_CLINICALTRIALID NUMBER,
                                              V_TRIALSITE VARCHAR2,
                                              V_PERSONID NUMBER,
                                              V_RESPONSETASKID NUMBER,
                                              V_RESPONSETIMESTAMP NUMBER,
                                              V_VISITID NUMBER,
                                              V_CRFPAGEID NUMBER,
                                              V_CRFELEMENTID NUMBER,
                                              V_DATAITEMID NUMBER,
                                              V_VISITCYCLENUMBER NUMBER,
                                              V_CRFPAGECYCLENUMBER NUMBER,
                                              V_CRFPAGETASKID NUMBER,
                                              V_RESPONSEVALUE VARCHAR2,
                                              V_VALUECODE VARCHAR2,
                                              V_USERNAME VARCHAR2,
                                              V_UNITOFMEASUREMENT VARCHAR2,
                                              V_COMMENTS VARCHAR2,
                                              V_RESPONSESTATUS NUMBER,
                                              V_CHANGED NUMBER,
                                              V_SOFTWAREVERSION VARCHAR2,
                                              V_REASONFORCHANGE VARCHAR2,
                                              V_LOCKSTATUS NUMBER,
                                              V_IMPORTTIMESTAMP NUMBER,
                                              V_VALIDATIONID NUMBER,
                                              V_VALIDATIONMESSAGE VARCHAR2,
                                              V_OVERRULEREASON VARCHAR2,
                                              V_LABRESULT VARCHAR2,
                                              V_CTCGRADE NUMBER,
                                              V_CLINICALTESTDATE NUMBER,
                                              V_LABORATORYCODE VARCHAR2,
                                              V_HADVALUE NUMBER,
                                              V_REPEATNUMBER NUMBER,
                                              V_RESPONSETIMESTAMP_TZ NUMBER,
                                              V_IMPORTTIMESTAMP_TZ NUMBER,
                                              V_DATABASETIMESTAMP NUMBER,
                                              V_DATABASETIMESTAMP_TZ NUMBER,
                                              V_SEQUENCEID NUMBER,
                                              V_STANDARDVALUE NUMBER,
                                              V_USERNAMEFULL VARCHAR2) is
begin

  -- EXCEPTION BLOCK: ATTEMPT INSERT, IF KEY VIOLATION, RECORD EXISTS, DO UPDATE INSTEAD
  begin
    -- STANDARD INSERT STATEMENT
    insert into dataitemresponsehistory
      (clinicaltrialid, trialsite, personid, responsetaskid, visitid,
       crfpageid, crfelementid, dataitemid, visitcyclenumber,
       crfpagecyclenumber, crfpagetaskid, responsevalue, responsetimestamp,
       valuecode, username, unitofmeasurement, comments, responsestatus,
       changed, softwareversion, reasonforchange, lockstatus,
       importtimestamp, validationid, validationmessage, overrulereason,
       labresult, ctcgrade, clinicaltestdate, laboratorycode, hadvalue,
       repeatnumber, responsetimestamp_tz, importtimestamp_tz,
       databasetimestamp, databasetimestamp_tz, sequenceid, standardvalue,
       usernamefull)
    values
      (v_clinicaltrialid, v_trialsite, v_personid, v_responsetaskid,
       v_visitid, v_crfpageid, v_crfelementid, v_dataitemid,
       v_visitcyclenumber, v_crfpagecyclenumber, v_crfpagetaskid,
       v_responsevalue, v_responsetimestamp, v_valuecode, v_username,
       v_unitofmeasurement, v_comments, v_responsestatus, v_changed,
       v_softwareversion, v_reasonforchange, v_lockstatus, v_importtimestamp,
       v_validationid, v_validationmessage, v_overrulereason, v_labresult,
       v_ctcgrade, v_clinicaltestdate, v_laboratorycode, v_hadvalue,
       v_repeatnumber, v_responsetimestamp_tz, v_importtimestamp_tz,
       v_databasetimestamp, v_databasetimestamp_tz, v_sequenceid,
       v_standardvalue, v_usernamefull);
  
    v_returnvalue := 1;
  exception
    when dup_val_on_index then
      -- STANDARD UPDATE STATEMENT
      update dataitemresponsehistory
         set clinicaltrialid = v_clinicaltrialid, trialsite = v_trialsite,
             personid = v_personid, responsetaskid = v_responsetaskid,
             visitid = v_visitid, crfpageid = v_crfpageid,
             crfelementid = v_crfelementid, dataitemid = v_dataitemid,
             visitcyclenumber = v_visitcyclenumber,
             crfpagecyclenumber = v_crfpagecyclenumber,
             crfpagetaskid = v_crfpagetaskid,
             responsevalue = v_responsevalue, valuecode = v_valuecode,
             username = v_username, unitofmeasurement = v_unitofmeasurement,
             comments = v_comments, responsestatus = v_responsestatus,
             changed = v_changed, softwareversion = v_softwareversion,
             reasonforchange = v_reasonforchange, lockstatus = v_lockstatus,
             importtimestamp = v_importtimestamp,
             validationid = v_validationid,
             validationmessage = v_validationmessage,
             overrulereason = v_overrulereason, labresult = v_labresult,
             ctcgrade = v_ctcgrade, clinicaltestdate = v_clinicaltestdate,
             laboratorycode = v_laboratorycode, hadvalue = v_hadvalue,
             repeatnumber = v_repeatnumber,
             responsetimestamp_tz = v_responsetimestamp_tz,
             importtimestamp_tz = v_importtimestamp_tz,
             databasetimestamp = v_databasetimestamp,
             databasetimestamp_tz = v_databasetimestamp_tz,
             sequenceid = v_sequenceid, standardvalue = v_standardvalue,
             usernamefull = v_usernamefull
       where clinicaltrialid = v_clinicaltrialid and
             trialsite = v_trialsite and personid = v_personid and
             responsetaskid = v_responsetaskid and
             responsetimestamp = v_responsetimestamp and
             repeatnumber = v_repeatnumber;
    
      v_returnvalue := 2;
    when others then
      v_returnvalue := 99;
  end;

end SP_MACRO_IMP_DIRH;
/
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
 V_SEQUENCEID          NUMBER
)
is
begin
  V_RETURNVALUE := 0;


    begin
         -- TRY INSERT FIRST
         insert into trialsubject
           (clinicaltrialid, trialsite, personid, dateofbirth, gender, localidentifier1, localidentifier2, proformastate, trialsubjectstatus, changed, lockstatus, importtimestamp, subjectgender, registrationstatus, discrepancystatus, sdvstatus, notestatus, importtimestamp_tz, sequenceid)
         values
           (v_clinicaltrialid, v_trialsite, v_personid, v_dateofbirth, v_gender, v_localidentifier1, v_localidentifier2, v_proformastate, v_trialsubjectstatus, v_changed, v_lockstatus, v_importtimestamp, v_subjectgender, v_registrationstatus, v_discrepancystatus, v_sdvstatus, v_notestatus, v_importtimestamp_tz, v_sequenceid);
          v_returnvalue := 1;
    exception
      -- AN ERROR OCCURED DURING INSERT
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
                 sequenceid = v_sequenceid
           where clinicaltrialid = v_clinicaltrialid
             and trialsite = v_trialsite
             and personid = v_personid;
           v_returnvalue := 2;
      when others then
           -- OTHER ERROR OCCURED
           v_returnvalue := 0;
      end;


end SP_MACRO_IMP_TRIALSUBJECT;
/
--//--
create or replace procedure SP_MACRO_IMP_VISITINSTANCE
(V_RETURNVALUE         OUT NUMBER,
 V_CLINICALTRIALID     NUMBER,
 V_TRIALSITE           VARCHAR2,
 V_PERSONID            NUMBER,
 V_VISITTASKID         NUMBER,
 V_VISITID             NUMBER,
 V_VISITCYCLENUMBER    NUMBER,
 V_VISITDATE           NUMBER,
 V_VISITSTATUS         NUMBER,
 V_CHANGED             NUMBER,
 V_LOCKSTATUS          NUMBER,
 V_IMPORTTIMESTAMP     NUMBER,
 V_DISCREPANCYSTATUS   NUMBER,
 V_SDVSTATUS           NUMBER,
 V_NOTESTATUS          NUMBER,
 V_IMPORTTIMESTAMP_TZ  NUMBER,
 V_SEQUENCEID          NUMBER
) is
begin
    begin
         -- TRY INSERT
         insert into visitinstance
           (clinicaltrialid, trialsite, personid, visittaskid, visitid, visitcyclenumber, visitdate, visitstatus, changed, lockstatus, importtimestamp, discrepancystatus, sdvstatus, notestatus, importtimestamp_tz, sequenceid)
         values
           (v_clinicaltrialid, v_trialsite, v_personid, v_visittaskid, v_visitid, v_visitcyclenumber, v_visitdate, v_visitstatus, v_changed, v_lockstatus, v_importtimestamp, v_discrepancystatus, v_sdvstatus, v_notestatus, v_importtimestamp_tz, v_sequenceid);
         v_returnvalue:=1;
  exception
    when dup_val_on_index then
          -- INSERT FAILED (DUPLICATE PRIMARY KEY)
          update visitinstance
             set clinicaltrialid = v_clinicaltrialid,
                 trialsite = v_trialsite,
                 personid = v_personid,
                 visittaskid = v_visittaskid,
                 visitid = v_visitid,
                 visitcyclenumber = v_visitcyclenumber,
                 visitdate = v_visitdate,
                 visitstatus = v_visitstatus,
                 changed = v_changed,
                 lockstatus = v_lockstatus,
                 importtimestamp = v_importtimestamp,
                 discrepancystatus = v_discrepancystatus,
                 sdvstatus = v_sdvstatus,
                 notestatus = v_notestatus,
                 importtimestamp_tz = v_importtimestamp_tz,
                 sequenceid = v_sequenceid
           where clinicaltrialid = v_clinicaltrialid
             and trialsite = v_trialsite
             and personid = v_personid
             and visittaskid = v_visittaskid;
           v_returnvalue:=2;
    when others then
       -- OTHER ERROR OCCURED
       v_returnvalue := 0;
   end;

end SP_MACRO_IMP_VISITINSTANCE;
/
--//--
create or replace procedure SP_MACRO_IMP_CRFPAGEINSTANCE
( V_RETURNVALUE           OUT NUMBER,
  V_CLINICALTRIALID       NUMBER,
  V_TRIALSITE             VARCHAR2,
  V_PERSONID              NUMBER,
  V_CRFPAGETASKID         NUMBER,
  V_VISITID               NUMBER,
  V_CRFPAGEID             NUMBER,
  V_VISITCYCLENUMBER      NUMBER,
  V_CRFPAGECYCLENUMBER    NUMBER,
  V_CRFPAGEDATE           NUMBER,
  V_CRFPAGESTATUS         NUMBER,
  V_CHANGED               NUMBER,
  V_CRFPAGEINSTANCELABEL  VARCHAR2,
  V_LOCKSTATUS            NUMBER,
  V_IMPORTTIMESTAMP       NUMBER,
  V_LABORATORYCODE        VARCHAR2,
  V_DISCREPANCYSTATUS     NUMBER,
  V_SDVSTATUS             NUMBER,
  V_NOTESTATUS            NUMBER,
  V_IMPORTTIMESTAMP_TZ    NUMBER,
  V_SEQUENCEID            NUMBER
) is
begin
    begin
      -- TRY INSERT
      insert into crfpageinstance
        (clinicaltrialid, trialsite, personid, crfpagetaskid, visitid, crfpageid, visitcyclenumber, crfpagecyclenumber, crfpagedate, crfpagestatus, changed, crfpageinstancelabel, lockstatus, importtimestamp, laboratorycode, discrepancystatus, sdvstatus, notestatus, importtimestamp_tz, sequenceid)
      values
        (v_clinicaltrialid, v_trialsite, v_personid, v_crfpagetaskid, v_visitid, v_crfpageid, v_visitcyclenumber, v_crfpagecyclenumber, v_crfpagedate, v_crfpagestatus, v_changed, v_crfpageinstancelabel, v_lockstatus, v_importtimestamp, v_laboratorycode, v_discrepancystatus, v_sdvstatus, v_notestatus, v_importtimestamp_tz, v_sequenceid);
      v_returnvalue:=1;
  exception
    when dup_val_on_index then
      -- INSERT FAILED: DUPLICATE KEY, DO UPDATE INSTEAD
      update crfpageinstance
         set clinicaltrialid = v_clinicaltrialid,
             trialsite = v_trialsite,
             personid = v_personid,
             crfpagetaskid = v_crfpagetaskid,
             visitid = v_visitid,
             crfpageid = v_crfpageid,
             visitcyclenumber = v_visitcyclenumber,
             crfpagecyclenumber = v_crfpagecyclenumber,
             crfpagedate = v_crfpagedate,
             crfpagestatus = v_crfpagestatus,
             changed = v_changed,
             crfpageinstancelabel = v_crfpageinstancelabel,
             lockstatus = v_lockstatus,
             importtimestamp = v_importtimestamp,
             laboratorycode = v_laboratorycode,
             discrepancystatus = v_discrepancystatus,
             sdvstatus = v_sdvstatus,
             notestatus = v_notestatus,
             importtimestamp_tz = v_importtimestamp_tz,
             sequenceid = v_sequenceid
       where clinicaltrialid = v_clinicaltrialid
         and trialsite = v_trialsite
         and personid = v_personid
         and crfpagetaskid = v_crfpagetaskid;
       v_returnvalue:=2;
    when others then
      -- OTHER ERROR
       v_returnvalue:=0;
  end;

end SP_MACRO_IMP_CRFPAGEINSTANCE;
/

--//--
UPDATE MACROCONTROL SET BUILDSUBVERSION = '37'
--//--