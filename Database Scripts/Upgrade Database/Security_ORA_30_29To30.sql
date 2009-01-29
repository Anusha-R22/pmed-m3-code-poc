UPDATE FUNCTIONMODULE SET MACROMODULE = 'SM' WHERE MACROMODULE = 'EX';

INSERT INTO MACROFunction (FunctionCode,MACROFunction) VALUES ('F2015','Change password properties');
INSERT INTO MACROFunction (FunctionCode,MACROFunction) VALUES ('F4009','Export study definition');
INSERT INTO MACROFunction (FunctionCode,MACROFunction) VALUES ('F4010','Laboratory/site administration');
INSERT INTO MACROFunction (FunctionCode,MACROFunction) VALUES ('F4011','Export laboratory');
INSERT INTO MACROFunction (FunctionCode,MACROFunction) VALUES ('F4012','Import laboratory');
INSERT INTO MACROFunction (FunctionCode,MACROFunction) VALUES ('F4013','Distribute laboratory');
INSERT INTO MACROFunction (FunctionCode,MACROFunction) VALUES ('F5031','Transfer data');

INSERT INTO Rolefunction (RoleCode,FunctionCode) VALUES ('MACROUser','F2015');
INSERT INTO Rolefunction (RoleCode,FunctionCode) VALUES ('MACROUser','F4009');
INSERT INTO Rolefunction (RoleCode,FunctionCode) VALUES ('MACROUser','F4010');
INSERT INTO Rolefunction (RoleCode,FunctionCode) VALUES ('MACROUser','F4011');
INSERT INTO Rolefunction (RoleCode,FunctionCode) VALUES ('MACROUser','F4012');
INSERT INTO Rolefunction (RoleCode,FunctionCode) VALUES ('MACROUser','F4013');
INSERT INTO Rolefunction (RoleCode,FunctionCode) VALUES ('MACROUser','F5031');

Insert INTO FunctionModule (FunctionCode,MACROModule) VALUES ('F2015','SM');
Insert INTO FunctionModule (FunctionCode,MACROModule) VALUES ('F4009','SM');
Insert INTO FunctionModule (FunctionCode,MACROModule) VALUES ('F4010','SM');
Insert INTO FunctionModule (FunctionCode,MACROModule) VALUES ('F4011','SM');
Insert INTO FunctionModule (FunctionCode,MACROModule) VALUES ('F4012','SM');
Insert INTO FunctionModule (FunctionCode,MACROModule) VALUES ('F4013','SM');
Insert INTO FunctionModule (FunctionCode,MACROModule) VALUES ('F5031','DE');
Insert INTO FunctionModule (FunctionCode,MACROModule) VALUES ('F5031','DR');

DELETE FROM FUNCTIONMODULE WHERE FUNCTIONCODE = 'F1002';
DELETE FROM Rolefunction WHERE FUNCTIONCODE = 'F1002';
DELETE FROM MACROFunction WHERE FUNCTIONCODE = 'F1002';


UPDATE SECURITYCONTROL SET BUILDSUBVERSION = '30';