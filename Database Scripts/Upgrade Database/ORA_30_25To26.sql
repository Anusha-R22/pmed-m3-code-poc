CREATE TABLE MACROUSERSETTINGS( USERNAME VARCHAR2(15) NOT NULL,  USERSETTING VARCHAR2(20) NOT NULL,  SETTINGVALUE VARCHAR2(50), CONSTRAINT PKMACROUSERSETTING PRIMARY KEY (USERNAME, USERSETTING, SETTINGVALUE));

INSERT INTO MACROTable (TableName,SegmentId,STYDEF,PATRSP,LABDEF) VALUES ('MACROUSERSETTINGS','',0,0,0);

UPDATE MACROControl SET BUILDSUBVERSION = '26';







