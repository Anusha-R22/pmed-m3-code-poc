ALTER TABLE MACROUSER ADD SYSADMIN SMALLINT;
GO
UPDATE MACROUSER SET SYSADMIN = 0;
UPDATE SECURITYCONTROL SET BUILDSUBVERSION = '26';