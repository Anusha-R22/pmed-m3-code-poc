
CREATE OR REPLACE PROCEDURE UPDATEDISCSTATUS
 (TRIALNAME in varchar2,SITE in varchar2, SUBJECTID in number,
		VISIT in number,VCYCLE in number,
		EFORMTASKID in number,
		RTASKID in number, RCYCLE in number)
AS
BEGIN
	DECLARE
		TRIALID number(6);

  BEGIN
		SELECT ClinicalTrialId into TRIALID FROM ClinicalTrial where Clinicaltrialname = TRIALNAME;

		Update TrialSubject set DiscrepancyStatus=(SELECT nvl(MAX(DECODE(MIMessageStatus,0,30,1,20,2,10,0)),0) FROM MIMESSAGE WHERE MIMessageType = 0 AND MIMessageHistory = 0 AND MIMessageTrialName = TRIALNAME AND MIMessageSite = SITE AND MIMessagePersonId = SUBJECTID) WHERE ClinicalTrialId = TRIALID AND TrialSite = SITE AND PersonId = SUBJECTID;
		Update VisitInstance set DiscrepancyStatus=(SELECT nvl(MAX(DECODE(MIMessageStatus,0,30,1,20,2,10,0)),0) FROM MIMESSAGE WHERE MIMessageType = 0 AND MIMessageHistory = 0 AND MIMessageTrialName = TRIALNAME AND MIMessageSite = SITE AND MIMessagePersonId = SUBJECTID AND MIMessageVisitId = VISIT AND MIMessageVisitCycle = VCYCLE) WHERE ClinicalTrialId = TRIALID AND TrialSite = SITE AND PersonId = SUBJECTID AND VisitId=VISIT AND VisitCycleNumber=VCYCLE;
		Update CRFPageInstance set DiscrepancyStatus=(SELECT nvl(MAX(DECODE(MIMessageStatus,0,30,1,20,2,10,0)),0) FROM MIMESSAGE WHERE MIMessageType = 0 AND MIMessageHistory = 0 AND MIMessageTrialName = TRIALNAME AND MIMessageSite = SITE AND MIMessagePersonId = SUBJECTID AND MIMessageVisitId = VISIT AND MIMessageVisitCycle = VCYCLE AND MIMessageCRFPageTaskId = EFORMTASKID) WHERE ClinicalTrialId = TRIALID AND TrialSite = SITE AND PersonId = SUBJECTID AND VisitId=VISIT AND VisitCycleNumber=VCYCLE AND CRFPageTaskId=EFORMTASKID;
		Update DataItemResponse set DiscrepancyStatus=(SELECT nvl(MAX(DECODE(MIMessageStatus,0,30,1,20,2,10,0)),0) FROM MIMESSAGE WHERE MIMessageType = 0 AND MIMessageHistory = 0 AND MIMessageTrialName = TRIALNAME AND MIMessageSite = SITE AND MIMessagePersonId = SUBJECTID AND MIMessageVisitId = VISIT AND MIMessageVisitCycle = VCYCLE AND MIMessageCRFPageTaskId = EFORMTASKID AND MIMESSAGERESPONSETASKID = RTASKID AND MIMessageResponseCycle = RCYCLE) WHERE ClinicalTrialId = TRIALID AND TrialSite = SITE AND PersonId = SUBJECTID AND VisitId=VISIT AND VisitCycleNumber=VCYCLE AND CRFPageTaskId=EFORMTASKID AND RESPONSETASKID=RTASKID AND RepeatNumber=RCYCLE;
	END;
END;
/
--//--

