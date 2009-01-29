DECLARE	@start varchar(16)
select @start = (select cast(isnull(max(messageid),0)+1 as varchar )from message)
EXEC sp_MACRO_Seq_Create @SequenceName = N'SEQ_MESSAGEID', @Seed = @start
GO