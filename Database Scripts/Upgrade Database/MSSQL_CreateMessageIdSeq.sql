-- creates a correct seeded messageid sequence for upgrading to patch 3.0.70b and 3.0.73
DECLARE	@start varchar(16)
select @start = (select cast(max(messageid)+1 as varchar )from message) 
EXEC sp_MACRO_Seq_Create @SequenceName = N'SEQ_MESSAGEID', @Seed = @start
GO
