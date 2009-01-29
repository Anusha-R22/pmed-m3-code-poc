ALTER TABLE CRFPage ADD EFORMWIDTH NUMBER(11);
INSERT INTO NewDBColumn (VersionMajor,VersionMinor,VersionRevision,TableName,ColumnName,ColumnOrder,DefaultValue,ChangeType,ColumnNumber) VALUES (3,0,19,'CRFPAGE','EFORMWIDTH',null,'#NULL#','NEWCOLUMN',null);
UPDATE MACROControl SET BUILDSUBVERSION = '19';