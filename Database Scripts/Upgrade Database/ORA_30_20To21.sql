ALTER TABLE TRIALSUBJECT ADD SEQUENCEID NUMBER(15);
ALTER TABLE VISITINSTANCE ADD SEQUENCEID NUMBER(15);
ALTER TABLE CRFPAGEINSTANCE ADD SEQUENCEID NUMBER(15);
ALTER TABLE DATAITEMRESPONSE ADD SEQUENCEID NUMBER(15);
ALTER TABLE DATAITEMRESPONSEHISTORY ADD SEQUENCEID NUMBER(15);
ALTER TABLE MIMESSAGE ADD SEQUENCEID NUMBER(15);

INSERT INTO NewDBColumn (VersionMajor,VersionMinor,VersionRevision,TableName,ColumnName,ColumnOrder,DefaultValue,ChangeType,ColumnNumber) VALUES (3,0,21,'TRIALSUBJECT','SEQUENCEID',null,'#NULL#','NEWCOLUMN',null);
INSERT INTO NewDBColumn (VersionMajor,VersionMinor,VersionRevision,TableName,ColumnName,ColumnOrder,DefaultValue,ChangeType,ColumnNumber) VALUES (3,0,21,'VISITINSTANCE','SEQUENCEID',null,'#NULL#','NEWCOLUMN',null);
INSERT INTO NewDBColumn (VersionMajor,VersionMinor,VersionRevision,TableName,ColumnName,ColumnOrder,DefaultValue,ChangeType,ColumnNumber) VALUES (3,0,21,'CRFPAGEINSTANCE','SEQUENCEID',null,'#NULL#','NEWCOLUMN',null);
INSERT INTO NewDBColumn (VersionMajor,VersionMinor,VersionRevision,TableName,ColumnName,ColumnOrder,DefaultValue,ChangeType,ColumnNumber) VALUES (3,0,21,'DATAITEMRESPONSE','SEQUENCEID',null,'#NULL#','NEWCOLUMN',null);
INSERT INTO NewDBColumn (VersionMajor,VersionMinor,VersionRevision,TableName,ColumnName,ColumnOrder,DefaultValue,ChangeType,ColumnNumber) VALUES (3,0,21,'DATAITEMRESPONSEHISTORY','SEQUENCEID',null,'#NULL#','NEWCOLUMN',null);
INSERT INTO NewDBColumn (VersionMajor,VersionMinor,VersionRevision,TableName,ColumnName,ColumnOrder,DefaultValue,ChangeType,ColumnNumber) VALUES (3,0,21,'MIMESSAGE','SEQUENCEID',null,'#NULL#','NEWCOLUMN',null);

CREATE SEQUENCE SEQ_SUBJECTDATA INCREMENT BY 1 START WITH 1 MAXVALUE 999999999999999 MINVALUE 1;

UPDATE MACROControl SET BUILDSUBVERSION = '21';