-- Updates trialsubject.notestatus
UPDATE  a   
SET	notestatus = 
	CASE (SELECT COUNT(*)
		FROM  mimessage m, clinicaltrial ct 
		WHERE	 a.clinicaltrialid  = ct.clinicaltrialid
		AND	a.trialsite  = m.mimessagesite
		AND	a.personid  = m.mimessagepersonid
		AND	m.mimessagetrialname  = ct.clinicaltrialname
		AND	m.mimessagetype  = 2
		AND	m.mimessagescope  = 1) 
	WHEN 0 THEN 0 
	ELSE 1 
	END 
FROM  trialsubject a 	
GO
GO

-- Updates visitinstance.notestatus
UPDATE  a   
SET	notestatus = 
	CASE (SELECT COUNT(*)
			FROM  mimessage m, clinicaltrial ct 
			WHERE a.clinicaltrialid  = ct.clinicaltrialid
			AND	a.trialsite  = m.mimessagesite
			AND	a.personid  = m.mimessagepersonid
			AND	a.visitid  = m.mimessagevisitid
			AND	a.visitcyclenumber  = m.mimessagevisitcycle
			AND	m.mimessagetrialname  = ct.clinicaltrialname
			AND	m.mimessagetype  = 2
			AND	m.mimessagescope  = 2) 
		WHEN 0 THEN 0 
		ELSE 1 
	END 
FROM  visitinstance a 
GO
GO
 

-- Updates crfpageinstance.notestatus
UPDATE  a   
SET	notestatus = 
	CASE (SELECT COUNT(*)
		FROM  mimessage m, clinicaltrial ct 
		WHERE a.clinicaltrialid  = ct.clinicaltrialid
		AND	a.trialsite  = m.mimessagesite
		AND	a.personid  = m.mimessagepersonid
		AND	a.crfpagetaskid  = m.mimessagecrfpagetaskid
		AND	m.mimessagetrialname  = ct.clinicaltrialname
		AND	m.mimessagetype  = 2
		AND	m.mimessagescope  = 3) 
	WHEN 0 THEN 0 
	ELSE 1 
	END 
FROM  crfpageinstance a 
GO
GO

-- Updates dataitemresponse.notestatus
UPDATE  a   
SET	notestatus = 
CASE (SELECT COUNT(*)
		FROM  mimessage m,
		clinicaltrial ct 
		WHERE a.clinicaltrialid  = ct.clinicaltrialid
		AND	a.trialsite  = m.mimessagesite
		AND	a.personid  = m.mimessagepersonid
		AND	a.responsetaskid  = m.mimessageresponsetaskid
		AND	a.repeatnumber  = m.mimessageresponsecycle
		AND	m.mimessagetrialname  = ct.clinicaltrialname
		AND	m.mimessagetype  = 2
		AND	m.mimessagescope  = 4) 
	WHEN 0 THEN 0 
	ELSE 1 
END 
FROM  dataitemresponse a 
GO
GO


-- Updates trialsubject.discrepancystatus
UPDATE  a   
SET	discrepancystatus = ISNULL((SELECT MAX(CASE m.mimessagestatus 
											WHEN 0 THEN 30 
											WHEN 1 THEN 20 
											WHEN 2 THEN 10 
											END)
FROM  mimessage m, clinicaltrial ct 
WHERE a.clinicaltrialid  = ct.clinicaltrialid
AND	a.trialsite  = m.mimessagesite
AND	a.personid  = m.mimessagepersonid
AND	m.mimessagetrialname  = ct.clinicaltrialname
AND	m.mimessagetype  = 0
AND	m.mimessagehistory  = 0), 0) 
FROM  trialsubject a 
GO
GO


-- Updates visitinstance.discrepancystatus
UPDATE  a   
SET	discrepancystatus = ISNULL((SELECT MAX(CASE m.mimessagestatus 
											WHEN 0 THEN 30 
											WHEN 1 THEN 20 
											WHEN 2 THEN 10 
											END)
FROM  mimessage m, clinicaltrial ct 
WHERE a.clinicaltrialid  = ct.clinicaltrialid
AND	a.trialsite  = m.mimessagesite
AND	a.personid  = m.mimessagepersonid
AND	a.visitid  = m.mimessagevisitid
AND	a.visitcyclenumber  = m.mimessagevisitcycle
AND	m.mimessagetrialname  = ct.clinicaltrialname
AND	m.mimessagetype  = 0
AND	m.mimessagehistory  = 0), 0) 
FROM  visitinstance a 
GO
GO


-- Updates crfpageinstance.discrepancystatus
UPDATE  a   
SET	discrepancystatus = ISNULL((SELECT MAX(CASE m.mimessagestatus 
											WHEN 0 THEN 30 
											WHEN 1 THEN 20 
											WHEN 2 THEN 10 
											END)
FROM  mimessage m, clinicaltrial ct 
WHERE	 a.clinicaltrialid  = ct.clinicaltrialid
AND	a.trialsite  = m.mimessagesite
AND	a.personid  = m.mimessagepersonid
AND	a.crfpagetaskid  = m.mimessagecrfpagetaskid
AND	m.mimessagetrialname  = ct.clinicaltrialname
AND	m.mimessagetype  = 0
AND	m.mimessagehistory  = 0), 0) 
FROM  crfpageinstance a 
GO
GO


-- Updates dataitemresponse.discrepancystatus
UPDATE  a   
SET	discrepancystatus = ISNULL((SELECT MAX(CASE m.mimessagestatus 
											WHEN 0 THEN 30 
											WHEN 1 THEN 20 
											WHEN 2 THEN 10 
											END)
FROM  mimessage m,  clinicaltrial ct 
WHERE	 a.clinicaltrialid  = ct.clinicaltrialid
AND	a.trialsite  = m.mimessagesite
AND	a.personid  = m.mimessagepersonid
AND	a.responsetaskid  = m.mimessageresponsetaskid
AND	a.repeatnumber  = m.mimessageresponsecycle
AND	m.mimessagetrialname  = ct.clinicaltrialname
AND	m.mimessagetype  = 0
AND	m.mimessagehistory  = 0), 0) 
FROM  dataitemresponse a 
GO
GO


-- Updates trialsubject.sdvstatus
UPDATE  a   
SET	sdvstatus = ISNULL((SELECT MAX(CASE m.mimessagescope 
									WHEN 1 THEN 
										CASE m.mimessagestatus 
										WHEN 0 THEN 30 
										WHEN 1 THEN 40 
										WHEN 2 THEN 20 
										WHEN 3 THEN 10 
										END 
									ELSE 
										CASE m.mimessagestatus 
										WHEN 0 THEN 30 
										WHEN 1 THEN 40 
										ELSE 0 
										END 
									END)
FROM  mimessage m, clinicaltrial ct 
WHERE	 a.clinicaltrialid  = ct.clinicaltrialid
AND	a.trialsite  = m.mimessagesite
AND	a.personid  = m.mimessagepersonid
AND	m.mimessagetrialname  = ct.clinicaltrialname
AND	m.mimessagetype  = 3
AND	m.mimessagehistory  = 0), 0) 
FROM  trialsubject a 
GO
GO

-- Updates visitinstance.sdvstatus
UPDATE  a   
SET	sdvstatus = ISNULL((SELECT MAX(CASE m.mimessagescope 
									WHEN 2 THEN 
										CASE m.mimessagestatus 
										WHEN 0 THEN 30 
										WHEN 1 THEN 40 
										WHEN 2 THEN 20 
										WHEN 3 THEN 10 
										END 
									ELSE 
										CASE m.mimessagestatus 
										WHEN 0 THEN 30 
										WHEN 1 THEN 40 
										ELSE 0 
										END 
									END)
FROM  mimessage m, clinicaltrial ct 
WHERE	 a.clinicaltrialid  = ct.clinicaltrialid
AND	a.trialsite  = m.mimessagesite
AND	a.personid  = m.mimessagepersonid
AND	a.visitid  = m.mimessagevisitid
AND	a.visitcyclenumber  = m.mimessagevisitcycle
AND	m.mimessagetrialname  = ct.clinicaltrialname
AND	m.mimessagetype  = 3
AND	m.mimessagehistory  = 0), 0) 
FROM  visitinstance a 
GO
GO

-- Updates crfpageinstance.sdvstatus
UPDATE  a   
SET	sdvstatus = ISNULL((SELECT MAX(CASE m.mimessagescope 
									WHEN 3 THEN 
										CASE m.mimessagestatus 
										WHEN 0 THEN 30 
										WHEN 1 THEN 40 
										WHEN 2 THEN 20 
										WHEN 3 THEN 10 
										END 
									ELSE 
										CASE m.mimessagestatus 
										WHEN 0 THEN 30 
										WHEN 1 THEN 40 
										ELSE 0 
										END 
									END)
FROM  mimessage m, clinicaltrial ct 
WHERE	 a.clinicaltrialid  = ct.clinicaltrialid
AND	a.trialsite  = m.mimessagesite
AND	a.personid  = m.mimessagepersonid
AND	a.crfpagetaskid  = m.mimessagecrfpagetaskid
AND	m.mimessagetrialname  = ct.clinicaltrialname
AND	m.mimessagetype  = 3
AND	m.mimessagehistory  = 0), 0) 
FROM  crfpageinstance a 
GO
GO

-- Updates dataitemresponse.sdvstatus
UPDATE  a   
SET	sdvstatus = ISNULL((SELECT MAX(CASE m.mimessagestatus 
									WHEN 0 THEN 30 
									WHEN 1 THEN 40 
									WHEN 2 THEN 20 
									WHEN 3 THEN 10 
									END)
FROM  mimessage m, clinicaltrial ct 
WHERE	 a.clinicaltrialid  = ct.clinicaltrialid
AND	a.trialsite  = m.mimessagesite
AND	a.personid  = m.mimessagepersonid
AND	a.responsetaskid  = m.mimessageresponsetaskid
AND	a.repeatnumber  = m.mimessageresponsecycle
AND	m.mimessagetrialname  = ct.clinicaltrialname
AND	m.mimessagetype  = 3
AND	m.mimessagehistory  = 0), 0) 
FROM  dataitemresponse a 	
GO
GO

