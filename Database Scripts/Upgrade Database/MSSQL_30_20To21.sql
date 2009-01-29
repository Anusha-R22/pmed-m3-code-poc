ALTER TABLE TRIALSUBJECT ADD SEQUENCEID NUMERIC(15);
ALTER TABLE VISITINSTANCE ADD SEQUENCEID NUMERIC(15);
ALTER TABLE CRFPAGEINSTANCE ADD SEQUENCEID NUMERIC(15);
ALTER TABLE DATAITEMRESPONSE ADD SEQUENCEID NUMERIC(15);
ALTER TABLE DATAITEMRESPONSEHISTORY ADD SEQUENCEID NUMERIC(15);
ALTER TABLE MIMESSAGE ADD SEQUENCEID NUMERIC(15);

INSERT INTO NewDBColumn (VersionMajor,VersionMinor,VersionRevision,TableName,ColumnName,ColumnOrder,DefaultValue,ChangeType,ColumnNumber) VALUES (3,0,21,'TRIALSUBJECT','SEQUENCEID',null,'#NULL#','NEWCOLUMN',null);
INSERT INTO NewDBColumn (VersionMajor,VersionMinor,VersionRevision,TableName,ColumnName,ColumnOrder,DefaultValue,ChangeType,ColumnNumber) VALUES (3,0,21,'VISITINSTANCE','SEQUENCEID',null,'#NULL#','NEWCOLUMN',null);
INSERT INTO NewDBColumn (VersionMajor,VersionMinor,VersionRevision,TableName,ColumnName,ColumnOrder,DefaultValue,ChangeType,ColumnNumber) VALUES (3,0,21,'CRFPAGEINSTANCE','SEQUENCEID',null,'#NULL#','NEWCOLUMN',null);
INSERT INTO NewDBColumn (VersionMajor,VersionMinor,VersionRevision,TableName,ColumnName,ColumnOrder,DefaultValue,ChangeType,ColumnNumber) VALUES (3,0,21,'DATAITEMRESPONSE','SEQUENCEID',null,'#NULL#','NEWCOLUMN',null);
INSERT INTO NewDBColumn (VersionMajor,VersionMinor,VersionRevision,TableName,ColumnName,ColumnOrder,DefaultValue,ChangeType,ColumnNumber) VALUES (3,0,21,'DATAITEMRESPONSEHISTORY','SEQUENCEID',null,'#NULL#','NEWCOLUMN',null);
INSERT INTO NewDBColumn (VersionMajor,VersionMinor,VersionRevision,TableName,ColumnName,ColumnOrder,DefaultValue,ChangeType,ColumnNumber) VALUES (3,0,21,'MIMESSAGE','SEQUENCEID',null,'#NULL#','NEWCOLUMN',null);

create procedure sp_MACRO_Seq_Create(@SequenceName as varchar(30), @Seed varchar(15)) as declare @ProcSQL varchar(500) 	execute ( 'Create table ' + @SequenceName + ' (SequenceNo numeric(15) identity(' + @seed + ',1), DummyCol smallint)') 	select @ProcSQL = 'CREATE procedure sp_MACRO_Seq_' + @SequenceName + ' (@NextVal numeric(15) Output) as ' + char(13) 	select @ProcSQL = @ProcSQL + 'insert into ' + @SequenceName + ' values (null) ' + char(13) 	select @ProcSQL = @ProcSQL + 'delete from ' + @SequenceName + char(13) 	select @ProcSQL = @ProcSQL + 'set @NextVal = @@IDENTITY' + char(13) 	execute (@ProcSQL); 
go
create procedure sp_MACRO_Seq_Drop(@SequenceName as varchar(30)) as     execute ('drop procedure sp_MACRO_seq_' + @SequenceName)     execute ('drop table ' + @SequenceName);
go
sp_MACRO_Seq_Create 'SEQ_SUBJECTDATA','1';
go

UPDATE MACROControl SET BUILDSUBVERSION = '21';