ALTER TABLE CRFELEMENT ADD DESCRIPTION VARCHAR2(255);
INSERT INTO NewDBColumn (VersionMajor,VersionMinor,VersionRevision,TableName,ColumnName,ColumnOrder,DefaultValue,ChangeType,ColumnNumber) VALUES (3,0,46,'CRFELEMENT','DESCRIPTION',null,'#NULL#','NEWCOLUMN',null);

UPDATE MACROCONTROL SET BUILDSUBVERSION = '46';