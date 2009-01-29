THIS IS HOW TO CREATE A SEQUENCE THAT HAS ALREADY STARTED ON THE PERSONID IN THE TRIALSUBJECT TABLE



--//--
DECLARE @MP VARCHAR(15)
SELECT @MP = CAST(COUNT(PERSONID)+1 AS VARCHAR(15)) FROM TRIALSUBJECT
EXECUTE SP_MACRO_SEQ_CREATE 'SEQ_PERSONID',@MP




--//--
declare
  mp number;
  ms varchar2(200);
begin

  select count(personid) into mp from trialsubject;
  mp:=mp+1;
  
  ms:= 'CREATE SEQUENCE SEQ_PERSONID INCREMENT BY 1 START WITH ';
  ms := ms || mp;
  ms := ms || ' MAXVALUE 999999999999999 MINVALUE ';
  ms := ms || mp;

  execute immediate ms;
end;
/