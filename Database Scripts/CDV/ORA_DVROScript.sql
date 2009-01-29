
/*---------------------------------------------------------------------------------------	      */
/*   Copyright:  InferMed Ltd. 2001. All Rights Reserved                                        */
/*   File:       MACRO_DVROTRIGGER (DVROScript.txt)                                             */
/*   Author:     Toby Aldridge, January 2001                                                    */
/*   Purpose:    Script to create a trigger that fills Response Only dataviews when	            */
/*               DataItemResponse is inserted into or updated                                   */
/*---------------------------------------------------------------------------------------	      */
/*   Revisions                                                                                  */
/*   Mo Morris 	22/4/2002, Changes around the option to use category codes or values.         */
/*                Changes stemming from the typing of dataview question columns.                */
/*   Mo Morris 	30/4/2002, changes around nulls and single quotes in response data            */
/*   Mo Morris 	21/6/2002, changes around Date Responses being placed in DateTime fields      */
/*   Mo Morris	3/12/2004 Bug 2446 - Adding Special Values facilities to CDV                  */
/*   Mo Morris	25/10/2006 Bug 2824 - Make Create Data Views Module comply with Partial Dates */
/*                When formating a date response the Partial Dates flag DataItemCase is checked */
/*                When the flag is 1 the date response is written away as a string              */
/*   Mo Morris	9/7/2007 - Bug 2936 - Make the Trigger correctly handle missing response data	*/
/*				 with a status of Success (0)													*/
/*---------------------------------------------------------------------------------------       */

CREATE OR REPLACE TRIGGER "MACRO_DVROTRIGGER" AFTER INSERT OR UPDATE OF "RESPONSEVALUE" ON "DATAITEMRESPONSE" FOR EACH ROW DECLARE

SQLString VARCHAR2(1000);
TableName VARCHAR2(30);
ColumnName VARCHAR2(30);
CurrentRows VARCHAR2(1);
SeparateVisits NUMBER;
OutputCatValues NUMBER;
ColumnType NUMBER;
Response VARCHAR2(255);
DateFS VARCHAR2(30);
DateResp VARCHAR2(30);
SVMissing VARCHAR2(2);
SVUnobtainable VARCHAR2(2);
SVNotApplicable VARCHAR2(2);
D1 NUMBER;
D2 NUMBER;
D3 NUMBER;
D4 NUMBER;
D5 NUMBER;
P1 VARCHAR2(4);
P2 VARCHAR2(4);
P3 VARCHAR2(4);
P4 VARCHAR2(4);
P5 VARCHAR2(4);
P6 VARCHAR2(4);
OWQGroupID NUMBER;
PartialDateFlag NUMBER;

BEGIN
    /*Extract OwnerQGroupID from table CRFELement */
    SELECT OwnerQGroupID into OWQGroupID
      FROM CRFELEMENT
      WHERE ClinicalTrialId = :New.ClinicalTrialId
      AND DataItemID = :New.DataItemId
      AND CRFPageId = :New.CRFPageId;

    SELECT OutputCategoryValues, DataViewSeparateVisits, SpecialValueMissing, SpecialValueUnobtainable, SpecialValueNotApplicable
      INTO  OutputCatValues, SeparateVisits, SVMissing, SVUnobtainable, SVNotApplicable
      FROM DataViewDetails
      WHERE ClinicalTrialId = :New.ClinicaltrialId;

    /* Extract the required TableName for the current response */
    IF SeparateVisits = 1 THEN
        SELECT DataViewName INTO TableName 
        FROM DataViewTables 
        WHERE ClinicalTrialId = :New.ClinicalTrialId
        AND VisitId = :New.VisitId
        AND CRFPageId = :New.CRFPageId
        AND QGroupID = OWQGroupID
        AND DataviewType = 'RO'; 
    ELSE
        SELECT DataViewName INTO TableName 
        FROM DataViewTables 
        WHERE ClinicalTrialId = :New.ClinicalTrialId  
        AND CRFPageId = :New.CRFPageId
        AND QGroupID = OWQGroupID
        AND DataviewType = 'RO';  
    END IF;
  
    IF TableName Is Not Null Then

        /* Extract ColumnName (DataItemCode) from table DataItem, for the current response */
        /* Extract ColumnType (DataType) from table DataItem, for the current response */
        /* Extract DateFS (DataItemFormat) from table DataItem, for the current response */
        /* Extract PartialDateFlag (DataItemCase) from table DataItem, for the current response */
        SELECT DataItemCode, DataType, DataItemFormat, DataItemCase
        INTO ColumnName, ColumnType, DateFS, PartialDateFlag
        FROM DataItem
        WHERE DataItem.ClinicalTrialId = :New.ClinicalTrialId
        AND DataITem.DataItemId = :New.DataItemId;

        IF PartialDateFlag Is Null THEN
            PartialDateFlag := 0;
        END IF;

        /* Format the ResponseValue ready for inserting into Data Views table DataViewName */
        /* test for a non-Null response */
        IF :New.ResponseValue Is Not Null THEN
            IF (ColumnType = 0) OR (ColumnType = 5) OR (ColumnType = 8) Then
                /* its a Text or Mulimedia or Thesaurus */
                Response := ''''||REPLACE(:New.ResponseValue,CHR(39),CHR(39)||CHR(39))||'''';
            ELSIF (ColumnType = 4) THEN
                /* its a Date question */
                /* Standardize the format string */
                DateFS := REPLACE(DateFS,'dd','d');
                DateFS := REPLACE(DateFS,'mm','m');
                DateFS := REPLACE(DateFS,'hh','h');
                DateFS := REPLACE(DateFS,'ss','s');
                DateFS := REPLACE(DateFS,'yyyy','y');
                DateFS := REPLACE(DateFS,':','/');
                DateFS := REPLACE(DateFS,'.','/');
                DateFS := REPLACE(DateFS,'-','/');
                DateFS := REPLACE(DateFS,' ','/');
                /* is it a date question that will be converted into a date field */
                IF (PartialDateFlag = 0) AND ((DateFS = 'd/m/y') OR (DateFS = 'm/d/y') OR (DateFS = 'y/m/d') OR (DateFS = 'h/m') OR (DateFS = 'h/m/s') OR (DateFS = 'd/m/y/h/m') OR (DateFS = 'm/d/y/h/m') OR (DateFS = 'y/m/d/h/m') OR (DateFS = 'd/m/y/h/m/s') OR (DateFS = 'm/d/y/h/m/s') OR (DateFS = 'y/m/d/h/m/s')) THEN
                    /* Standardize the separtors within the response */
                    DateResp := :New.ResponseValue;
                    DateResp := REPLACE(DateResp,':','/');
                    DateResp := REPLACE(DateResp,'.','/');
                    DateResp := REPLACE(DateResp,'-','/');
                    DateResp := REPLACE(DateResp,' ','/');
                    /* Using the standardized format string the response is separated out into its individual parts and put together as a Universal Date Format string */ 
                    D1 := INSTR(DateResp,'/',1,1);
                    D2 := INSTR(DateResp,'/',1,2);
                    D3 := INSTR (DateResp,'/',1,3);
                    D4 := INSTR(DateResp,'/',1,4);
                    D5 := INSTR(DateResp,'/',1,5);
                    /* using DateFS as a template build a Universal Date Format string */
                    IF (DateFS = 'd/m/y') THEN
                        P1 := SUBSTR(DateResp,1,(D1-1));
                        P2 := SUBSTR(DateResp,(D1+1),(D2-D1-1));
                        P3 := SUBSTR(DateResp,(D2+1),4);
                        DateResp := P2||'/'||P1||'/'||P3;
                    ELSIF (DateFS = 'm/d/y') THEN
                        P1 := SUBSTR(DateResp,1,(D1-1));
                        P2 := SUBSTR(DateResp,(D1+1),(D2-D1-1));
                        P3 := SUBSTR(DateResp,(D2+1),4);
                        DateResp := P1||'/'||P2||'/'||P3;
                    ELSIF (DateFS = 'y/m/d') THEN
                        P1 := SUBSTR(DateResp,1,(D1-1));
                        P2 := SUBSTR(DateResp,(D1+1),(D2-D1-1));
                        P3 := SUBSTR(DateResp,(D2+1),2);
                        DateResp := P2||'/'||P3||'/'||P1;
                    ELSIF (DateFS = 'h/m') THEN
                        P1 := SUBSTR(DateResp,1,(D1-1));
                        P2 := SUBSTR(DateResp,(D1+1),2);
                        DateResp := '01/01/1900 '||P1||':'||P2;
                    ELSIF (DateFS = 'h/m/s') THEN
                        P1 := SUBSTR(DateResp,1,(D1-1));
                        P2 := SUBSTR(DateResp,(D1+1),(D2-D1-1));
                        P3 := SUBSTR(DateResp,(D2+1),2);
                        DateResp := '01/01/1900 '||P1||':'||P2||':'||P3;
                    ELSIF (DateFS = 'd/m/y/h/m') THEN
                        P1 := SUBSTR(DateResp,1,(D1-1));
                        P2 := SUBSTR(DateResp,(D1+1),(D2-D1-1));
                        P3 := SUBSTR(DateResp,(D2+1),(D3-D2-1));
                        P4 := SUBSTR(DateResp,(D3+1),(D4-D3-1));
                        P5 := SUBSTR(DateResp,(D4+1),2);
                        DateResp := P2||'/'||P1||'/'||P3||' '||P4||':'||P5;
                    ELSIF (DateFS = 'm/d/y/h/m') THEN
                        P1 := SUBSTR(DateResp,1,(D1-1));
                        P2 := SUBSTR(DateResp,(D1+1),(D2-D1-1));
                        P3 := SUBSTR(DateResp,(D2+1),(D3-D2-1));
                        P4 := SUBSTR(DateResp,(D3+1),(D4-D3-1));
                        P5 := SUBSTR(DateResp,(D4+1),2);
                        DateResp := P1||'/'||P2||'/'||P3||' '||P4||':'||P5;
                    ELSIF (DateFS = 'y/m/d/h/m') THEN
                        P1 := SUBSTR(DateResp,1,(D1-1));
                        P2 := SUBSTR(DateResp,(D1+1),(D2-D1-1));
                        P3 := SUBSTR(DateResp,(D2+1),(D3-D2-1));
                        P4 := SUBSTR(DateResp,(D3+1),(D4-D3-1));
                        P5 := SUBSTR(DateResp,(D4+1),2);
                        DateResp := P2||'/'||P3||'/'||P1||' '||P4||':'||P5;
                    ELSIF (DateFS = 'd/m/y/h/m/s') THEN
                        P1 := SUBSTR(DateResp,1,(D1-1));
                        P2 := SUBSTR(DateResp,(D1+1),(D2-D1-1));
                        P3 := SUBSTR(DateResp,(D2+1),(D3-D2-1));
                        P4 := SUBSTR(DateResp,(D3+1),(D4-D3-1));
                        P5 := SUBSTR(DateResp,(D4+1),(D5-D4-1));
                        P6 := SUBSTR(DateResp,(D5+1),2);
                        DateResp := P2||'/'||P1||'/'||P3||' '||P4||':'||P5||':'||P6;
                    ELSIF (DateFS = 'm/d/y/h/m/s') THEN
                        P1 := SUBSTR(DateResp,1,(D1-1));
                        P2 := SUBSTR(DateResp,(D1+1),(D2-D1-1));
                        P3 := SUBSTR(DateResp,(D2+1),(D3-D2-1));
                        P4 := SUBSTR(DateResp,(D3+1),(D4-D3-1));
                        P5 := SUBSTR(DateResp,(D4+1),(D5-D4-1));
                        P6 := SUBSTR(DateResp,(D5+1),2);
                        DateResp := P1||'/'||P2||'/'||P3||' '||P4||':'||P5||':'||P6;
                    ELSE
                        /* 'y/m/d/h/m/s' */
                        P1 := SUBSTR(DateResp,1,(D1-1));
                        P2 := SUBSTR(DateResp,(D1+1),(D2-D1-1));
                        P3 := SUBSTR(DateResp,(D2+1),(D3-D2-1));
                        P4 := SUBSTR(DateResp,(D3+1),(D4-D3-1));
                        P5 := SUBSTR(DateResp,(D4+1),(D5-D4-1));
                        P6 := SUBSTR(DateResp,(D5+1),2);
                        DateResp := P2||'/'||P3||'/'||P1||' '||P4||':'||P5||':'||P6;
                    END IF;
                    /* Response := ''''||TO_DATE(DateResp,'mm/dd/yyyy hh24:mi:ss')||''''; */
                    Response := 'TO_DATE('''||DateResp||''',''mm/dd/yyyy hh24:mi:ss'')';
                ELSE
                    /* date formats y/m, m/y and y/d/m and partial dates are not converted */
                    Response := ''''||REPLACE(:New.ResponseValue,CHR(39),CHR(39)||CHR(39))||'''';    
                END IF;
            ELSIF (ColumnType = 1) Then
                /* its Category, is it the code or the value that is required */
                IF OutputCatValues = 1 THEN
                    Response := ''''||REPLACE(:New.ResponseValue,CHR(39),CHR(39)||CHR(39))||'''';
                ELSE
                    Response := ''''||:New.ValueCode||'''';
                END IF;
            ELSE
                /* its IntegerData, Real or LabTest */
                Response := :New.ResponseValue;
            END IF;
        ELSE
            /* its an empty response, check for any Special Values first */
            IF (SVMissing is null) AND (SVUnobtainable is null) AND (SVNotApplicable is null) THEN
                Response := 'Null';
            ELSE
                IF (ColumnType = 0) OR (ColumnType = 5) OR (ColumnType = 8) Then
                    /* its a Text or Mulimedia or Thesaurus */
                    IF (:New.ResponseStatus = 10) THEN 
                        Response := SVMissing;
                    ELSIF (:New.ResponseStatus = -8) THEN
                        Response := SVNotApplicable;
                    ELSIf (:New.ResponseStatus = -5) THEN
                        Response := SVUnobtainable;
                    ELSIf (:New.ResponseStatus = 0) THEN
                        Response := 'Null';
                    END IF;
                ELSIF (ColumnType = 4) THEN
                    /* its a Date question */
                    /* Standardize the format string */
                    DateFS := REPLACE(DateFS,'dd','d');
                    DateFS := REPLACE(DateFS,'mm','m');
                    DateFS := REPLACE(DateFS,'hh','h');
                    DateFS := REPLACE(DateFS,'ss','s');
                    DateFS := REPLACE(DateFS,'yyyy','y');
                    DateFS := REPLACE(DateFS,':','/');
                    DateFS := REPLACE(DateFS,'.','/');
                    DateFS := REPLACE(DateFS,'-','/');
                    DateFS := REPLACE(DateFS,' ','/');
                    /* is it a date question that will be converted into a date field */
                    IF (PartialDateFlag = 0) AND ((DateFS = 'd/m/y') OR (DateFS = 'm/d/y') OR (DateFS = 'y/m/d') OR (DateFS = 'h/m') OR (DateFS = 'h/m/s') OR (DateFS = 'd/m/y/h/m') OR (DateFS = 'm/d/y/h/m') OR (DateFS = 'y/m/d/h/m') OR (DateFS = 'd/m/y/h/m/s') OR (DateFS = 'm/d/y/h/m/s') OR (DateFS = 'y/m/d/h/m/s')) THEN
                        IF (:New.ResponseStatus = 10) THEN
                            /* convert Missing Special Value into a date */
                            IF (SVMissing = '-1') THEN
                                Response := 'TO_DATE('''||'12/29/1899'||''',''mm/dd/yyyy hh24:mi:ss'')';
                            ELSIF (SVMissing = '-2') THEN
                                Response := 'TO_DATE('''||'12/28/1899'||''',''mm/dd/yyyy hh24:mi:ss'')';
                            ELSIF (SVMissing = '-3') THEN
                                Response := 'TO_DATE('''||'12/27/1899'||''',''mm/dd/yyyy hh24:mi:ss'')';
                            ELSIF (SVMissing = '-4') THEN
                                Response := 'TO_DATE('''||'12/26/1899'||''',''mm/dd/yyyy hh24:mi:ss'')';
                            ELSIF (SVMissing = '-5') THEN
                                Response := 'TO_DATE('''||'12/25/1899'||''',''mm/dd/yyyy hh24:mi:ss'')';
                            ELSIF (SVMissing = '-6') THEN
                                Response := 'TO_DATE('''||'12/24/1899'||''',''mm/dd/yyyy hh24:mi:ss'')';
                            ELSIF (SVMissing = '-7') THEN
                                Response := 'TO_DATE('''||'12/23/1899'||''',''mm/dd/yyyy hh24:mi:ss'')';
                            ELSIF (SVMissing = '-8') THEN
                                Response := 'TO_DATE('''||'12/22/1899'||''',''mm/dd/yyyy hh24:mi:ss'')';
                            ELSIF (SVMissing = '-9') THEN
                                Response := 'TO_DATE('''||'12/21/1899'||''',''mm/dd/yyyy hh24:mi:ss'')';
                            END IF;
                        ELSIF (:New.ResponseStatus = -8) THEN
                            /* convert NotApplicable Special Value into a date */
                            IF (SVNotApplicable = '-1') THEN
                                Response := 'TO_DATE('''||'12/29/1899'||''',''mm/dd/yyyy hh24:mi:ss'')';
                            ELSIF (SVNotApplicable = '-2') THEN
                                Response := 'TO_DATE('''||'12/28/1899'||''',''mm/dd/yyyy hh24:mi:ss'')';
                            ELSIF (SVNotApplicable = '-3') THEN
                                Response := 'TO_DATE('''||'12/27/1899'||''',''mm/dd/yyyy hh24:mi:ss'')';
                            ELSIF (SVNotApplicable = '-4') THEN
                                Response := 'TO_DATE('''||'12/26/1899'||''',''mm/dd/yyyy hh24:mi:ss'')';
                            ELSIF (SVNotApplicable = '-5') THEN
                                Response := 'TO_DATE('''||'12/25/1899'||''',''mm/dd/yyyy hh24:mi:ss'')';
                            ELSIF (SVNotApplicable = '-6') THEN
                                Response := 'TO_DATE('''||'12/24/1899'||''',''mm/dd/yyyy hh24:mi:ss'')';
                            ELSIF (SVNotApplicable = '-7') THEN
                                Response := 'TO_DATE('''||'12/23/1899'||''',''mm/dd/yyyy hh24:mi:ss'')';
                            ELSIF (SVNotApplicable = '-8') THEN
                                Response := 'TO_DATE('''||'12/22/1899'||''',''mm/dd/yyyy hh24:mi:ss'')';
                            ELSIF (SVNotApplicable = '-9') THEN
                                Response := 'TO_DATE('''||'12/21/1899'||''',''mm/dd/yyyy hh24:mi:ss'')';
                            END IF;
                        ELSIf (:New.ResponseStatus = -5) THEN
                            /* convert Unobtainable Special Value into a date */
                            IF (SVUnobtainable = '-1') THEN
                                Response := 'TO_DATE('''||'12/29/1899'||''',''mm/dd/yyyy hh24:mi:ss'')';
                            ELSIF (SVUnobtainable = '-2') THEN
                                Response := 'TO_DATE('''||'12/28/1899'||''',''mm/dd/yyyy hh24:mi:ss'')';
                            ELSIF (SVUnobtainable = '-3') THEN
                                Response := 'TO_DATE('''||'12/27/1899'||''',''mm/dd/yyyy hh24:mi:ss'')';
                            ELSIF (SVUnobtainable = '-4') THEN
                                Response := 'TO_DATE('''||'12/26/1899'||''',''mm/dd/yyyy hh24:mi:ss'')';
                            ELSIF (SVUnobtainable = '-5') THEN
                                Response := 'TO_DATE('''||'12/25/1899'||''',''mm/dd/yyyy hh24:mi:ss'')';
                            ELSIF (SVUnobtainable = '-6') THEN
                                Response := 'TO_DATE('''||'12/24/1899'||''',''mm/dd/yyyy hh24:mi:ss'')';
                            ELSIF (SVUnobtainable = '-7') THEN
                                Response := 'TO_DATE('''||'12/23/1899'||''',''mm/dd/yyyy hh24:mi:ss'')';
                            ELSIF (SVUnobtainable = '-8') THEN
                                Response := 'TO_DATE('''||'12/22/1899'||''',''mm/dd/yyyy hh24:mi:ss'')';
                            ELSIF (SVUnobtainable = '-9') THEN
                                Response := 'TO_DATE('''||'12/21/1899'||''',''mm/dd/yyyy hh24:mi:ss'')';
                            END IF;
                        ELSIf (:New.ResponseStatus = 0) THEN
                            Response := 'Null';
                        END IF;
                    ELSE
                        /* date formats y/m, m/y and y/d/m and partial dates are held as strings */
                        IF (:New.ResponseStatus = 10) THEN 
                            Response := SVMissing;
                        ELSIF (:New.ResponseStatus = -8) THEN
                            Response := SVNotApplicable;
                        ELSIf (:New.ResponseStatus = -5) THEN
                            Response := SVUnobtainable;
                        ELSIf (:New.ResponseStatus = 0) THEN
                            Response := 'Null';
                        END IF;
                    END IF;
                ELSIF (ColumnType = 1) THEN
                    /* its Category */
                    IF (:New.ResponseStatus = 10) THEN 
                        Response := SVMissing;
                    ELSIF (:New.ResponseStatus = -8) THEN
                        Response := SVNotApplicable;
                    ELSIf (:New.ResponseStatus = -5) THEN
                        Response := SVUnobtainable;
                    ELSIf (:New.ResponseStatus = 0) THEN
                        Response := 'Null';
                    END IF;
                ELSE
                    /* its IntegerData, Real or LabTest */
                    IF (:New.ResponseStatus = 10) THEN 
                        Response := SVMissing;
                    ELSIF (:New.ResponseStatus = -8) THEN
                        Response := SVNotApplicable;
                    ELSIf (:New.ResponseStatus = -5) THEN
                        Response := SVUnobtainable;
                    ELSIf (:New.ResponseStatus = 0) THEN
                        Response := 'Null';
                    END IF;
                END IF;
            END IF;
        END IF;

        IF OWQGroupID = 0 THEN
            /* assess wether its going to be an Insert (I) or an Update (U) */
            execute immediate 'select decode(count(*),0,''I'',''U'') FROM '||TableName
            ||' WHERE '
            ||'ClinicalTrialId = '||:New.ClinicalTrialId
            ||' AND Site = '''||:New.TrialSite||''''
            ||' AND PersonId = '||:New.PersonId
            ||' AND VisitId = '||:New.VisitId
            ||' AND VisitCycleNumber = '||:New.VisitCycleNumber
            ||' AND CRFPageId = '||:New.CRFPageId
            ||' AND CRFPageCycleNumber = '||:New.CRFPageCycleNumber
            INTO CurrentRows;

            IF CurrentRows = 'I' Then
                SQLString :=
                'INSERT INTO '||TableName
                ||' ('
                ||'ClinicalTrialId,Site,PersonId,VisitId,VisitCycleNumber,CRFPageId,CRFPageCycleNumber,'
                ||ColumnName
                ||')'
                ||' VALUES('
                ||:New.ClinicalTrialId||','''
                ||:New.TrialSite||''','
                ||:New.PersonId||','
                ||:New.VisitId||','
                ||:New.VisitCycleNumber||','
                ||:New.CRFPageId||','
                ||:New.CRFPageCycleNumber||','
                ||Response
                ||')';
            ELSE
                SQLString :=
                'UPDATE '||TableName
                ||' SET '||ColumnName||' = '||Response
                ||' WHERE '
                ||'ClinicalTrialId = '||:New.ClinicalTrialId
                ||' AND Site = '''||:New.TrialSite||''''
                ||' AND PersonId = '||:New.PersonId
                ||' AND VisitId = '||:New.VisitId
                ||' AND VisitCycleNumber = '||:New.VisitCycleNumber
                ||' AND CRFPageId = '||:New.CRFPageId
                ||' AND CRFPageCycleNumber = '||:New.CRFPageCycleNumber;
            END IF;
        ELSE
            /* Its a Repeating Question Group Data View Table */
            /* assess wether its going to be an Insert (I) or an Update (U) */
            execute immediate 'select decode(count(*),0,''I'',''U'') FROM '||TableName
            ||' WHERE '
            ||'ClinicalTrialId = '||:New.ClinicalTrialId
            ||' AND Site = '''||:New.TrialSite||''''
            ||' AND PersonId = '||:New.PersonId
            ||' AND VisitId = '||:New.VisitId
            ||' AND VisitCycleNumber = '||:New.VisitCycleNumber
            ||' AND CRFPageId = '||:New.CRFPageId
            ||' AND CRFPageCycleNumber = '||:New.CRFPageCycleNumber
            ||' AND OwnerQGroupID = '||OWQGroupID
            ||' AND RepeatNumber = '||:New.RepeatNumber
            INTO CurrentRows;

            IF CurrentRows = 'I' Then
                SQLString :=
                'INSERT INTO '||TableName
                ||' ('
                ||'ClinicalTrialId,Site,PersonId,VisitId,VisitCycleNumber,CRFPageId,CRFPageCycleNumber,OwnerQGroupID,RepeatNumber,'
                ||ColumnName
                ||')'
                ||' VALUES('
                ||:New.ClinicalTrialId||','''
                ||:New.TrialSite||''','
                ||:New.PersonId||','
                ||:New.VisitId||','
                ||:New.VisitCycleNumber||','
                ||:New.CRFPageId||','
                ||:New.CRFPageCycleNumber||','
                ||OWQGroupID||','
                ||:New.RepeatNumber||','
                ||Response
                ||')';
            ELSE
                SQLString :=
                'UPDATE '||TableName
                ||' SET '||ColumnName||' = '||Response
                ||' WHERE '
                ||'ClinicalTrialId = '||:New.ClinicalTrialId
                ||' AND Site = '''||:New.TrialSite||''''
                ||' AND PersonId = '||:New.PersonId
                ||' AND VisitId = '||:New.VisitId
                ||' AND VisitCycleNumber = '||:New.VisitCycleNumber
                ||' AND CRFPageId = '||:New.CRFPageId
                ||' AND CRFPageCycleNumber = '||:New.CRFPageCycleNumber
                ||' AND OwnerQGroupID = '||OWQGroupID
                ||' AND RepeatNumber = '||:New.RepeatNumber;
            END IF;
        END IF;
    
    END IF;
	

/*	EXECUTE IMMEDIATE 'INSERT INTO DVRO_QMDemo_SubDetails_Fam (ClinicalTrialId,Site,PersonId,VisitId,VisitCycleNumber,CRFPageId,CRFPageCycleNumber,OwnerQGroupID,RepeatNumber,FamName) Values(1,''OraTown'',5,10034,1,10008,1,1,1,'''||REPLACE(SQLString,CHR(39),CHR(39)||CHR(39))||''')'; */

/* EXECUTE IMMEDIATE 'INSERT INTO DVRO_QMDemo_SubDetails_Fam (ClinicalTrialId,Site,PersonId,VisitId,VisitCycleNumber,CRFPageId,CRFPageCycleNumber,OwnerQGroupID,RepeatNumber,FamName) Values(1,''OraTown'',5,10034,1,10008,1,1,1,''failed'')'; */

   EXECUTE IMMEDIATE SQLString;

  /* TA 08/03/2001: catch all exceptions so that data is always saved to DataItemResponse */
  EXCEPTION
    when others then
        null;

END;