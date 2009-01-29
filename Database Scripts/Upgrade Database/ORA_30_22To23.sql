ALTER TABLE MIMESSAGE ADD MIMESSAGECRFPAGEID NUMBER(11);
ALTER TABLE MIMESSAGE ADD MIMESSAGECRFPAGECYCLE NUMBER(6);
ALTER TABLE MIMESSAGE ADD MIMESSAGEDATAITEMID NUMBER(11);

ALTER TABLE MIMESSAGE MODIFY MIMESSAGETEXT VARCHAR2(2000);

INSERT INTO NewDBColumn (VersionMajor,VersionMinor,VersionRevision,TableName,ColumnName,ColumnOrder,DefaultValue,ChangeType,ColumnNumber) VALUES (3,0,23,'MIMESSAGE','MIMESSAGECRFPAGEID',null,'#NULL#','NEWCOLUMN',null);
INSERT INTO NewDBColumn (VersionMajor,VersionMinor,VersionRevision,TableName,ColumnName,ColumnOrder,DefaultValue,ChangeType,ColumnNumber) VALUES (3,0,23,'MIMESSAGE','MIMESSAGECRFPAGECYCLE',null,'#NULL#','NEWCOLUMN',null);
INSERT INTO NewDBColumn (VersionMajor,VersionMinor,VersionRevision,TableName,ColumnName,ColumnOrder,DefaultValue,ChangeType,ColumnNumber) VALUES (3,0,23,'MIMESSAGE','MIMESSAGEDATAITEMID',null,'#NULL#','NEWCOLUMN',null);

UPDATE MIMESSAGE SET MIMESSAGEDATAITEMID = (SELECT DISTINCT DATAITEMRESPONSE.DATAITEMID  FROM DATAITEMRESPONSE,CLINICALTRIAL WHERE CLINICALTRIAL.CLINICALTRIALID = DATAITEMRESPONSE.CLINICALTRIALID AND CLINICALTRIAL.CLINICALTRIALNAME = MIMESSAGE.MIMESSAGETRIALNAME AND DATAITEMRESPONSE.TRIALSITE = MIMESSAGE.MIMESSAGESITE AND DATAITEMRESPONSE.PERSONID = MIMESSAGE.MIMESSAGEPERSONID AND DATAITEMRESPONSE.RESPONSETASKID = MIMESSAGE.MIMESSAGERESPONSETASKID);
UPDATE MIMESSAGE SET MIMESSAGECRFPAGEID = (SELECT DISTINCT DATAITEMRESPONSE.CRFPAGEID FROM DATAITEMRESPONSE,CLINICALTRIAL WHERE CLINICALTRIAL.CLINICALTRIALID = DATAITEMRESPONSE.CLINICALTRIALID AND CLINICALTRIAL.CLINICALTRIALNAME = MIMESSAGE.MIMESSAGETRIALNAME AND DATAITEMRESPONSE.TRIALSITE = MIMESSAGE.MIMESSAGESITE AND DATAITEMRESPONSE.PERSONID = MIMESSAGE.MIMESSAGEPERSONID AND DATAITEMRESPONSE.RESPONSETASKID = MIMESSAGE.MIMESSAGERESPONSETASKID);
UPDATE MIMESSAGE SET MIMESSAGECRFPAGECYCLE = (SELECT DISTINCT DATAITEMRESPONSE.CRFPAGECYCLENUMBER FROM DATAITEMRESPONSE,CLINICALTRIAL WHERE CLINICALTRIAL.CLINICALTRIALID = DATAITEMRESPONSE.CLINICALTRIALID AND CLINICALTRIAL.CLINICALTRIALNAME = MIMESSAGE.MIMESSAGETRIALNAME AND DATAITEMRESPONSE.TRIALSITE = MIMESSAGE.MIMESSAGESITE AND DATAITEMRESPONSE.PERSONID = MIMESSAGE.MIMESSAGEPERSONID AND DATAITEMRESPONSE.RESPONSETASKID = MIMESSAGE.MIMESSAGERESPONSETASKID);

UPDATE MACROControl SET BUILDSUBVERSION = '23';
