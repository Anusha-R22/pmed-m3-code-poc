
ALTER TABLE MACROUSERSETTINGS DROP CONSTRAINT PKMACROUSERSETTING;
GO
ALTER TABLE MACROUSERSETTINGS ALTER COLUMN USERNAME VARCHAR(20) NOT NULL;
GO
ALTER TABLE MACROUSERSETTINGS ADD CONSTRAINT PKMACROUSERSETTING PRIMARY KEY (USERNAME, USERSETTING, SETTINGVALUE);
GO
UPDATE MACROControl SET BUILDSUBVERSION = '43';