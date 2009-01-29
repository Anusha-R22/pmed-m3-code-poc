create procedure sp_MACRO_Seq_Create(@SequenceName as varchar(30), @Seed varchar(15)) as declare @ProcSQL varchar(500) 	execute ( 'Create table ' + @SequenceName + ' (SequenceNo numeric(15) identity(' + @Seed + ',1), DummyCol smallint)') 	select @ProcSQL = 'CREATE procedure sp_MACRO_Seq_' + @SequenceName + ' (@NextVal numeric(15) Output) as ' + char(13) 	select @ProcSQL = @ProcSQL + 'insert into ' + @SequenceName + ' values (null) ' + char(13) 	select @ProcSQL = @ProcSQL + 'delete from ' + @SequenceName + char(13) 	select @ProcSQL = @ProcSQL + 'set @NextVal = @@IDENTITY' + char(13) 	execute (@ProcSQL); 
go
create procedure sp_MACRO_Seq_Drop(@SequenceName as varchar(30)) as     execute ('drop procedure sp_MACRO_seq_' + @SequenceName)     execute ('drop table ' + @SequenceName);
go