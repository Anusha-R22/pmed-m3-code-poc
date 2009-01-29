
INSERT INTO MACROFunction (FunctionCode,MACROFunction) VALUES ('F1010','Access Batch Validation');
INSERT INTO Rolefunction (RoleCode,FunctionCode) VALUES ('MACROUser','F1010');
Insert INTO FunctionModule (FunctionCode,MACROModule) VALUES ('F1010','BV');

UPDATE SECURITYCONTROL SET BUILDSUBVERSION = '37';
