
INSERT INTO MACROFunction (FunctionCode,MACROFunction) VALUES ('F5033','Register subject');
INSERT INTO Rolefunction (RoleCode,FunctionCode) VALUES ('MACROUser','F5033');
Insert INTO FunctionModule (FunctionCode,MACROModule) VALUES ('F5033','DE');

UPDATE SECURITYCONTROL SET BUILDSUBVERSION = '46';