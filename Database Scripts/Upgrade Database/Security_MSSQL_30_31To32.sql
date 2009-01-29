INSERT INTO MACROFunction (FunctionCode,MACROFunction) VALUES ('F5032','View lock/freeze history');

INSERT INTO Rolefunction (RoleCode,FunctionCode) VALUES ('MACROUser','F5032');

Insert INTO FunctionModule (FunctionCode,MACROModule) VALUES ('F5032','DE');

INSERT INTO FUNCTIONMODULE (FUNCTIONCODE, MACROMODULE) VALUES ('F5026','DE');

UPDATE SECURITYCONTROL SET BUILDSUBVERSION = '32';