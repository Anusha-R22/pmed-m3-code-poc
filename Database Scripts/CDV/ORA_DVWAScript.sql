CREATE OR REPLACE TRIGGER "MACRO_DVWATRIGGER" AFTER INSERT OR UPDATE OF "CLINICALTRIALID", "CRFPAGECYCLENUMBER", "CRFPAGEID", "CTCGRADE", "DATAITEMID", "LABRESULT", "PERSONID", "RESPONSESTATUS", "RESPONSETIMESTAMP", "RESPONSEVALUE", "TRIALSITE", "VALUECODE", "VISITCYCLENUMBER", "VISITID" ON "DATAITEMRESPONSE" FOR EACH ROW DECLARE

SQLString VARCHAR2(255);
TableName VARCHAR2(30);
CurrentRows VARCHAR2(1);
CalcDataItemCode VARCHAR2(15);
CalcDataType INTEGER;
ResponseTimeStamp VARCHAR2(19);
SeparateVisits NUMBER;

ResponseValue varchar2(257);
ValueCode varchar2(17);
LabResult varchar2(6);
CTCGrade varchar2(6);
Units varchar2(17);
OQGroupID NUMBER;

	FUNCTION fn_DateFromMS (p_MSDate IN NUMBER)
	 RETURN VARCHAR2
	IS
		v_WholeDays	NUMBER;
		v_PartDays	NUMBER (30,30);
		v_WholeDate	DATE;
		v_Result	VARCHAR2 (19);
	BEGIN
		v_WholeDays := FLOOR (p_MSDate);
		v_PartDays := p_MSDate - v_WholeDays;
		v_WholeDate := TO_DATE (TO_NUMBER (TO_CHAR (TO_DATE ('30-DEC-1899','dd-mon-yyyy'), 'J'))+v_WholeDays, 'J');
		v_Result := TO_CHAR (v_WholeDate + v_PartDays, 'YYYY-MM-DD HH24:MI:SS');
		RETURN (v_Result);
	END;

  Function SQLFromField(sValue in VARCHAR2)
   RETURN VARCHAR2
  IS
  BEGIN
    If sValue is null then
      RETURN('null');
    else
      /* Mo 10/5/2002, REPLACE call added to handle single quotes */
      RETURN(''''||REPLACE(sValue,CHR(39),CHR(39)||CHR(39))||'''');
    end if;
  end;

BEGIN

  responsevalue :=sqlfromfield(:New.ResponseValue);
  ValueCode :=sqlfromfield(:New.ValueCode);
  LabResult :=sqlfromfield(:New.LabResult);
  Units :=sqlfromfield(:New.UnitofMeasurement);

  if :New.CTCGrade is null then
    CTCGrade := 'null';
  else
    CTCGrade := :New.CTCGrade;
  end if;

  SELECT DataViewSeparateVisits into SeparateVisits
    FROM DataViewDetails
    WHERE ClinicalTrialId = :New.ClinicalTrialId;

  IF SeparateVisits = 1 THEN
    SELECT DataViewName INTO TableName 
      FROM DataViewTables 
      WHERE ClinicalTrialId = :New.ClinicalTrialId
      AND VisitId = :New.VisitId  
      AND CRFPageId = :New.CRFPageId
      AND DataviewType = 'WA';   
  ELSE
    SELECT DataViewName INTO TableName 
      FROM DataViewTables 
      WHERE ClinicalTrialId = :New.ClinicalTrialId  
      AND CRFPageId = :New.CRFPageId
      AND DataviewType = 'WA';  
  END IF;
            
  IF TableName Is Not Null Then

    ResponseTimeStamp := fn_DateFromMS(:New.ResponseTimeStamp);

    SELECT DataItemCode, DataType INTO CalcDataItemCode, CalcDataType
      FROM DataItem
      WHERE ClinicalTrialId = :New.ClinicalTrialId
      AND DataItemId = :New.DataItemId;

    SELECT OwnerQGroupID INTO OQGroupID
      FROM CRFElement
      WHERE ClinicalTrialId = :New.ClinicalTrialId
      AND DataItemId = :New.DataItemId
      AND CRFPageId = :New.CRFPageId;

    if UPDATING then  
      execute immediate 'DELETE FROM '||TableName
        ||' WHERE'
        ||' ClinicalTrialId = '||:New.ClinicalTrialId
        ||' AND Site = '''||:New.TrialSite||''''
        ||' AND PersonId = '||:New.PersonId
        ||' AND VisitId = '||:New.VisitId
        ||' AND VisitCycleNumber = '||:New.VisitCycleNumber
        ||' AND CRFPageId = '||:New.CRFPageId
        ||' AND CRFPageCycleNumber = '||:New.CRFPageCycleNumber
        ||' AND DataItemCode = '''||CalcDataItemCode||''''
        ||' AND RepeatNumber = '||:New.RepeatNumber; 
    end if;
        
    EXECUTE IMMEDIATE 'INSERT INTO '||TableName        
        ||' VALUES('
        ||:New.ClinicalTrialId||','''
        ||:New.TrialSite||''','
        ||:New.PersonId||','
        ||:New.VisitId||','
        ||:New.VisitCycleNumber||','
        ||:New.CRFPageId||','
        ||:New.CRFPageCycleNumber||','
        ||OQGroupID||','
        ||''''||CalcDataItemCode||''','
        ||:New.RepeatNumber||','
        ||''''||ResponseTimeStamp||''','
        ||CalcDataType||','
        ||ResponseValue||','
        ||Units||','
        ||ValueCode||','
        ||:New.ResponseStatus||','
        ||LabResult||','
        ||CTCGrade	      
        ||')';
       
  END IF;

  /* TA 08/03/2001: catch all exceptions so that data is always saved to DataItemResponse */
  EXCEPTION
    when others then
        null;

END;
