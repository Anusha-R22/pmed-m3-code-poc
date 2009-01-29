ALTER TABLE CRFELEMENT ADD HOTLINK VARCHAR2(2000);
INSERT INTO NewDBColumn (VersionMajor,VersionMinor,VersionRevision,TableName,ColumnName,ColumnOrder,DefaultValue,ChangeType,ColumnNumber) VALUES (3,0,20,'CRFELEMENT','HOTLINK',null,'#NULL#','NEWCOLUMN',null);
UPDATE MACROControl SET BUILDSUBVERSION = '20';