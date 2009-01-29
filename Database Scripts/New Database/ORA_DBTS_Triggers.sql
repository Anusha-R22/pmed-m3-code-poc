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
