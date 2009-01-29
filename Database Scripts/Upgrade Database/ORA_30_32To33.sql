--//--SEPARATOR=--//--
CREATE OR REPLACE FUNCTION DBTIMESTAMP RETURN NUMBER
IS 
TS NUMBER(16,10); -- declare the return variable 
BEGIN 
-- The to_char function does not return milliseconds in oracle 8i
TS := to_char(sysdate,'J')-2452552+37533+(to_char(sysdate,'HH24')/24.0)+(to_char(sysdate,'MI')/1440.0)+(to_char(sysdate,'SS')/86400.0);
RETURN TS; 
EXCEPTION 
WHEN OTHERS THEN -- any error at all 
RETURN 0; 
END;
/
--//--
create or replace trigger SETDATABASETIMESTAMP BEFORE INSERT OR UPDATE OF DATABASETIMESTAMP, DATABASETIMESTAMP_TZ ON DATAITEMRESPONSE FOR EACH ROW
-----------------------------------------------------------------------------------
-- SETDATABASETIMESTAMP
--
-- RS 28/01/2003
--
-- Trigger that ensures that the databasetimestamp is set with the ACTUAL value
-- from the database
--
-----------------------------------------------------------------------------------
declare
lTimezone number;
begin
if :new.DatabaseTimestamp=0 or :new.DatabaseTimestamp is NULL then

   -- Retrieve the Timezone from the MACRODBSETTINGS table, use 0 if value not found
   begin
        select settingvalue into lTimezone from MACRODBSetting where settingsection='timezone' and settingkey='dbtz';  
        exception
          when no_data_found then
              lTimezone :=0;
   end;

   -- Set the values to be used during the insert/update
   :new.DatabaseTimestamp := DBTIMESTAMP();
   :new.DatabaseTimestamp_TZ := lTimezone;
              
end if;
end;
/
--//--
create or replace trigger SETDATABASETIMESTAMP_DIRH BEFORE INSERT OR UPDATE OF DATABASETIMESTAMP, DATABASETIMESTAMP_TZ ON DATAITEMRESPONSEHISTORY FOR EACH ROW
-----------------------------------------------------------------------------------
-- SETDATABASETIMESTAMP_DIRH
--
-- RS 28/01/2003
--
-- Trigger that ensures that the databasetimestamp is set with the ACTUAL value
-- from the database
--
-----------------------------------------------------------------------------------
declare
lTimezone number;
begin
if :new.DatabaseTimestamp=0 or :new.DatabaseTimestamp is NULL then

   -- Retrieve the Timezone from the MACRODBSETTINGS table, use 0 if value not found
   begin
        select settingvalue into lTimezone from MACRODBSetting where settingsection='timezone' and settingkey='dbtz';  
        exception
          when no_data_found then
              lTimezone :=0;
   end;

   -- Set the values to be used during the insert/update
   :new.DatabaseTimestamp := DBTIMESTAMP();
   :new.DatabaseTimestamp_TZ := lTimezone;
              
end if;
end;
/
--//--
ALTER TABLE DATAITEMRESPONSE ADD USERNAMEFULL VARCHAR2(100);
--//--
ALTER TABLE DATAITEMRESPONSEHISTORY ADD USERNAMEFULL VARCHAR2(100);
--//--
INSERT INTO NewDBColumn (VersionMajor,VersionMinor,VersionRevision,TableName,ColumnName,ColumnOrder,DefaultValue,ChangeType,ColumnNumber) VALUES (3,0,32,'DATAITEMRESPONSE','USERNAMEFULL',null,'#NULL#','NEWCOLUMN',null);
--//--
INSERT INTO NewDBColumn (VersionMajor,VersionMinor,VersionRevision,TableName,ColumnName,ColumnOrder,DefaultValue,ChangeType,ColumnNumber) VALUES (3,0,32,'DATAITEMRESPONSEHISTORY','USERNAMEFULL',null,'#NULL#','NEWCOLUMN',null);
--//--
CREATE TABLE BATCHRESPONSEDATA (BATCHRESPONSEID NUMBER(11), CLINICALTRIALID NUMBER(11), SITE VARCHAR2(8), PERSONID NUMBER(11), SUBJECTLABEL VARCHAR2(50), VISITID NUMBER(11), VISITCYCLENUMBER NUMBER(6), VISITCYCLEDATE NUMBER(16,10), CRFPAGEID NUMBER(11), CRFPAGECYCLENUMBER NUMBER(6), CRFPAGECYCLEDATE NUMBER(16,10), DATAITEMID NUMBER(11), REPEATNUMBER NUMBER(6), RESPONSE VARCHAR2(255), USERNAME VARCHAR2(20), UPLOADMESSAGE VARCHAR2(255), CONSTRAINT PKBATCHRESPONSEDATA PRIMARY KEY (BATCHRESPONSEID));
--//--
INSERT INTO MACROTable (TableName,SegmentId,STYDEF,PATRSP,LABDEF) VALUES ('BATCHRESPONSEDATA','',0,0,0);
--//--
UPDATE MACROControl SET BUILDSUBVERSION = '33';
--//--