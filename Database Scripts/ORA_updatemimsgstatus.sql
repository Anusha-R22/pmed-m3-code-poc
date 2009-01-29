
update trialsubject a
set notestatus = decode((select count(*)
from mimessage m, clinicaltrial ct
where a.clinicaltrialid = ct.clinicaltrialid
and a.trialsite = m.mimessagesite
and a.personid = m.mimessagepersonid
and m.mimessagetrialname = ct.clinicaltrialname
and m.mimessagetype = 2
and m.mimessagescope = 1), 0, 0, 1)
/

update visitinstance a
set notestatus = decode((select count(*)
from mimessage m, clinicaltrial ct
where a.clinicaltrialid = ct.clinicaltrialid
and a.trialsite = m.mimessagesite
and a.personid = m.mimessagepersonid
and a.visitid = m.mimessagevisitid
and a.visitcyclenumber = m.mimessagevisitcycle
and m.mimessagetrialname = ct.clinicaltrialname
and m.mimessagetype = 2
and m.mimessagescope = 2), 0, 0, 1)
/

update crfpageinstance a
set notestatus = decode((select count(*)
from mimessage m, clinicaltrial ct
where a.clinicaltrialid = ct.clinicaltrialid
and a.trialsite = m.mimessagesite
and a.personid = m.mimessagepersonid
and a.crfpagetaskid = m.mimessagecrfpagetaskid
and m.mimessagetrialname = ct.clinicaltrialname
and m.mimessagetype = 2
and m.mimessagescope = 3), 0, 0, 1)
/

update dataitemresponse a
set notestatus = decode((select count(*)
from mimessage m, clinicaltrial ct
where a.clinicaltrialid = ct.clinicaltrialid
and a.trialsite = m.mimessagesite
and a.personid = m.mimessagepersonid
and a.responsetaskid = m.mimessageresponsetaskid
and a.repeatnumber = m.mimessageresponsecycle
and m.mimessagetrialname = ct.clinicaltrialname
and m.mimessagetype = 2
and m.mimessagescope = 4), 0, 0, 1)
/

update trialsubject a
set discrepancystatus = nvl((select max(decode(m.mimessagestatus, 0, 30, 1, 20, 2, 10))
from mimessage m, clinicaltrial ct
where a.clinicaltrialid = ct.clinicaltrialid
and a.trialsite = m.mimessagesite
and a.personid = m.mimessagepersonid
and m.mimessagetrialname = ct.clinicaltrialname
and m.mimessagetype = 0
and m.mimessagehistory = 0),0)
/

update visitinstance a
set discrepancystatus = nvl((select max(decode(m.mimessagestatus, 0, 30, 1, 20, 2, 10))
from mimessage m, clinicaltrial ct
where a.clinicaltrialid = ct.clinicaltrialid
and a.trialsite = m.mimessagesite
and a.personid = m.mimessagepersonid
and a.visitid = m.mimessagevisitid
and a.visitcyclenumber = m.mimessagevisitcycle
and m.mimessagetrialname = ct.clinicaltrialname
and m.mimessagetype = 0
and m.mimessagehistory = 0),0)
/

update crfpageinstance a
set discrepancystatus = nvl((select max(decode(m.mimessagestatus, 0, 30, 1, 20, 2, 10))
from mimessage m, clinicaltrial ct
where a.clinicaltrialid = ct.clinicaltrialid
and a.trialsite = m.mimessagesite
and a.personid = m.mimessagepersonid
and a.crfpagetaskid = m.mimessagecrfpagetaskid
and m.mimessagetrialname = ct.clinicaltrialname
and m.mimessagetype = 0
and m.mimessagehistory = 0),0)
/

update dataitemresponse a
set discrepancystatus = nvl((select max(decode(m.mimessagestatus, 0, 30, 1, 20, 2, 10))
from mimessage m, clinicaltrial ct
where a.clinicaltrialid = ct.clinicaltrialid
and a.trialsite = m.mimessagesite
and a.personid = m.mimessagepersonid
and a.responsetaskid = m.mimessageresponsetaskid
and a.repeatnumber = m.mimessageresponsecycle
and m.mimessagetrialname = ct.clinicaltrialname
and m.mimessagetype = 0
and m.mimessagehistory = 0),0)
/

update trialsubject a
set sdvstatus = nvl((select max(decode(m.mimessagescope, 1, decode(m.mimessagestatus, 0, 30, 1, 40, 2, 20, 3, 10), decode(m.mimessagestatus, 0, 30, 1, 40, 0)))
from mimessage m, clinicaltrial ct
where a.clinicaltrialid = ct.clinicaltrialid
and a.trialsite = m.mimessagesite
and a.personid = m.mimessagepersonid
and m.mimessagetrialname = ct.clinicaltrialname
and m.mimessagetype = 3
and m.mimessagehistory = 0),0)
/

update visitinstance a
set sdvstatus = nvl((select max(decode(m.mimessagescope, 2, decode(m.mimessagestatus, 0, 30, 1, 40, 2, 20, 3, 10),  decode(m.mimessagestatus, 0, 30, 1, 40, 0)))
from mimessage m, clinicaltrial ct
where a.clinicaltrialid = ct.clinicaltrialid
and a.trialsite = m.mimessagesite
and a.personid = m.mimessagepersonid
and a.visitid = m.mimessagevisitid
and a.visitcyclenumber = m.mimessagevisitcycle
and m.mimessagetrialname = ct.clinicaltrialname
and m.mimessagetype = 3
and m.mimessagehistory = 0),0)
/

update crfpageinstance a
set sdvstatus = nvl((select max(decode(m.mimessagescope, 3, decode(m.mimessagestatus, 0, 30, 1, 40, 2, 20, 3, 10), decode(m.mimessagestatus, 0, 30, 1, 40, 0)))
from mimessage m, clinicaltrial ct
where a.clinicaltrialid = ct.clinicaltrialid
and a.trialsite = m.mimessagesite
and a.personid = m.mimessagepersonid
and a.crfpagetaskid = m.mimessagecrfpagetaskid
and m.mimessagetrialname = ct.clinicaltrialname
and m.mimessagetype = 3
and m.mimessagehistory = 0),0)
/

update dataitemresponse a
set sdvstatus = nvl((select max(decode(m.mimessagestatus, 0, 30, 1, 40, 2, 20, 3, 10))
from mimessage m, clinicaltrial ct
where a.clinicaltrialid = ct.clinicaltrialid
and a.trialsite = m.mimessagesite
and a.personid = m.mimessagepersonid
and a.responsetaskid = m.mimessageresponsetaskid
and a.repeatnumber = m.mimessageresponsecycle
and m.mimessagetrialname = ct.clinicaltrialname
and m.mimessagetype = 3
and m.mimessagehistory = 0),0)
/


