DECLARE
	vclinicaltrialid dataitemresponsehistory.clinicaltrialid%type := null;
	vtrialsite dataitemresponsehistory.trialsite%type;
	vpersonid dataitemresponsehistory.personid%type;
	vcrfpagetaskid dataitemresponsehistory.crfpagetaskid%type;
	vvisitid dataitemresponsehistory.visitid%type;
	vvisitcyclenumber dataitemresponsehistory.visitcyclenumber%type;
	CURSOR dirh IS
	select d5.clinicaltrialid, d5.trialsite, d5.personid, d5.responsetaskid, d5.responsetimestamp, 
	d5.repeatnumber, d5.crfpagetaskid, d5.visitid, d5.visitcyclenumber,
	d5.databasetimestamp, d5.sequenceid
	from (select clinicaltrialid, trialsite, personid, responsetaskid, repeatnumber
		from dataitemresponsehistory d2
		where clinicaltrialid = 18
		and exists (select null 
			from dataitemresponsehistory d1, dataitemresponse dir2
			where d1.clinicaltrialid = d2.clinicaltrialid
			and d1.trialsite = d2.trialsite
			and d1.personid = d2.personid
			and d1.responsetaskid = d2.responsetaskid
			and d1.repeatnumber = d2.repeatnumber
			and dir2.clinicaltrialid = d1.clinicaltrialid
			and dir2.trialsite = d1.trialsite
			and dir2.personid = d1.personid
			and dir2.responsetaskid = d1.responsetaskid
			and dir2.repeatnumber = d1.repeatnumber
			and d1.databasetimestamp > d2.databasetimestamp
			and d1.sequenceid < d2.sequenceid
			and dir2.responsetimestamp <> d1.responsetimestamp
			and dir2.responsetimestamp <> d2.responsetimestamp)
		group by clinicaltrialid, trialsite, personid, responsetaskid, repeatnumber
		union select clinicaltrialid, trialsite, personid, responsetaskid, repeatnumber
		from dataitemresponsehistory d4
		where clinicaltrialid = 18
		and sequenceid > (select d3.sequenceid
			from dataitemresponsehistory d3, dataitemresponse dir1
			where dir1.clinicaltrialid = d4.clinicaltrialid
			and dir1.trialsite = d4.trialsite
			and dir1.personid = d4.personid
			and dir1.responsetaskid = d4.responsetaskid
			and dir1.repeatnumber = d4.repeatnumber
			and d3.clinicaltrialid = d4.clinicaltrialid
			and d3.trialsite = d4.trialsite
			and d3.personid = d4.personid
			and d3.responsetaskid = d4.responsetaskid
			and d3.repeatnumber = d4.repeatnumber
			and d3.responsetimestamp = dir1.responsetimestamp)
		group by clinicaltrialid, trialsite, personid, responsetaskid, repeatnumber) qstoupdate,
	dataitemresponsehistory d5, dataitemresponse dir3
	where qstoupdate.clinicaltrialid = d5.clinicaltrialid
	and qstoupdate.trialsite = d5.trialsite
	and qstoupdate.personid = d5.personid
	and qstoupdate.responsetaskid = d5.responsetaskid
	and qstoupdate.repeatnumber = d5.repeatnumber
	and dir3.clinicaltrialid = d5.clinicaltrialid
	and dir3.trialsite = d5.trialsite
	and dir3.personid = d5.personid
	and dir3.responsetaskid = d5.responsetaskid
	and dir3.repeatnumber = d5.repeatnumber
	order by d5.clinicaltrialid, d5.trialsite, d5.personid, d5.responsetaskid, d5.repeatnumber,
	decode(d5.responsetimestamp, dir3.responsetimestamp, 1, 0), d5.databasetimestamp;
BEGIN
	FOR dirhrow IN dirh
	LOOP
		if vclinicaltrialid is null then
			vclinicaltrialid := dirhrow.clinicaltrialid;
			vtrialsite := dirhrow.trialsite;
			vpersonid := dirhrow.personid;
			vcrfpagetaskid := dirhrow.crfpagetaskid;
			vvisitid := dirhrow.visitid;
			vvisitcyclenumber := dirhrow.visitcyclenumber;
		end if;
		update dataitemresponsehistory set lockstatus = lockstatus
			where clinicaltrialid = dirhrow.clinicaltrialid
			and trialsite = dirhrow.trialsite
			and personid = dirhrow.personid
			and responsetaskid = dirhrow.responsetaskid
			and responsetimestamp = dirhrow.responsetimestamp
			and repeatnumber = dirhrow.repeatnumber;
		if not (vclinicaltrialid = dirhrow.clinicaltrialid
			and vtrialsite = dirhrow.trialsite
			and vpersonid = dirhrow.personid
			and vcrfpagetaskid = dirhrow.crfpagetaskid) then
			update crfpageinstance set lockstatus = lockstatus
				where clinicaltrialid = vclinicaltrialid
				and trialsite = vtrialsite
				and personid = vpersonid
				and crfpagetaskid = vcrfpagetaskid;
			if not (vclinicaltrialid = dirhrow.clinicaltrialid
				and vtrialsite = dirhrow.trialsite
				and vpersonid = dirhrow.personid
				and vvisitid = dirhrow.visitid
				and vvisitcyclenumber = dirhrow.visitcyclenumber) then
				update visitinstance set lockstatus = lockstatus
					where clinicaltrialid = vclinicaltrialid
					and trialsite = vtrialsite
					and personid = vpersonid
					and visitid = vvisitid
					and visitcyclenumber = vvisitcyclenumber;
			end if;
			vclinicaltrialid := dirhrow.clinicaltrialid;
			vtrialsite := dirhrow.trialsite;
			vpersonid := dirhrow.personid;
			vcrfpagetaskid := dirhrow.crfpagetaskid;
			vvisitid := dirhrow.visitid;
			vvisitcyclenumber := dirhrow.visitcyclenumber;
		end if;
	END LOOP;
	update crfpageinstance set lockstatus = lockstatus
		where clinicaltrialid = vclinicaltrialid
		and trialsite = vtrialsite
		and personid = vpersonid
		and crfpagetaskid = vcrfpagetaskid;
	update visitinstance set lockstatus = lockstatus
		where clinicaltrialid = vclinicaltrialid
		and trialsite = vtrialsite
		and personid = vpersonid
		and visitid = vvisitid
		and visitcyclenumber = vvisitcyclenumber;
END;