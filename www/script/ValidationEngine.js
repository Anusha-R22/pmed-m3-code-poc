
//
// Global scope is essential for the following declarations!
//
var	undefined;	// This must never be defined!
var	oForm;	// Form data structure object
var	bEvaluationError;	// Evaluation Error flag
// DPH 29/11/2002 
var gnCurrentRepeat=null; // holds current repeat for use in validation

// Enumerations for Data Types
var	etText=0;
var	etCategory=1;
var	etIntegerNumber=2;
var	etRealNumber=3;
var	etDateTime=4;
var	etMultimedia=5;
var	etLabTest=6;
var	etCatSelect=7;

//ic 19/08/2002
// Enumerations for statuses
var eComment=2;
var eNote=1;
var eCommentNote=3;

var eDiscrepancyType=0;
var eSDVType=1;
var eNoteType=2;

//case
var eUpperCase=1;
var eLowerCase=2;

//rfc
var eRFCValue=1;
var eRFCStatus=2;
var eRFCOverrule=3;

var	oLastFocusID=null;	// The field last used for entry. Used to allow validations to be performed on gotfocus, rather than lostfocus.
var	oFocusTarget=null;

var bFieldInstancing=false; // global to mark if a field is currently being instanced

var oTopImages=top.oImages;
var gbReValidationOk;

//dependency type enum
var eDepType = new Object();
eDepType.Validation="V";
eDepType.Derivation="D";
eDepType.Skip="S";

//validation enum
var eValidation = new Object();
eValidation.Warn="W";
eValidation.OKWarn="O";
eValidation.Reject="R";
eValidation.Inform="I";

//status enum
var eStatus = new Object();
eStatus.CancelledByUser=-20;
eStatus.Requested=-10;
eStatus.NotApplicable=-8;
eStatus.Unobtainable=-5;
eStatus.Success=0;
eStatus.Missing=10;
eStatus.Inform=20;
eStatus.OKWarning=25;
eStatus.Warning=30;
eStatus.InvalidData=40;

//lock/freeze enum
var eLock = new Object();
eLock.Locked=5;
eLock.Frozen=6;

//discrepancy status enum
var eDiscStatus = new Object();
eDiscStatus.Responded=20;
eDiscStatus.Raised=30;

//sdv status enum
var eSDVStatus = new Object();
eSDVStatus.Complete=20;
eSDVStatus.Planned=30;
eSDVStatus.Queried=40;

//--------------------------------------------------------------------------------

//
// Arezzo: date(years,months,days,hours,minutes,seconds)
//      date(years,months,days)
//      time(hours,minutes,seconds)
// Returns a datetime in millisecond format,or null if invalid.
// OC 30/11/2007 - Issue number 2962 - US Date Issue
// Changed to specify exact zero date to avoid  "roll-overs"

function jsdate(nYears,nMonths,nDays,nHours,nMinutes,nSeconds)
{
	//var	oDate=new Date(0);	// Use zero to avoid internal date "roll-overs"
	var oDate=new Date(1970,0,1); //OC -Issue 2962- Specify exact zero date regardless of timezone to avoid "roll-overs"
	oDate.setFullYear(nYears);
	oDate.setMonth(nMonths-1);
	oDate.setDate(nDays);
	oDate.setHours(nHours);
	oDate.setMinutes(nMinutes);
	oDate.setSeconds(nSeconds);
	if(oDate.getFullYear()==nYears
			&& oDate.getMonth()==nMonths-1
			&& oDate.getDate()==nDays
			&& oDate.getHours()==nHours
			&& oDate.getMinutes()==nMinutes
			&& oDate.getSeconds()==nSeconds)
		return oDate.getTime();
	bEvaluationError=true;
	return null;
}

//
// Arezzo: datenow
// Returns midnight of today
//
function jsdatenow()
{
	var	oDate=new Date();

	oDate.setHours(0);
	oDate.setMinutes(0);
	oDate.setSeconds(0);
	oDate.setMilliseconds(0);
	return oDate.getTime();
}

//
// Do a 'like' wise comparison on the two identified fields.
// Returns true if they are alike,else false
// Wildcards are "%" for 0+ wilds and "?" for exactly 1 wild.
//
function jslike(sStr,sPatt)
{
	var	sPattern='';
	var	nCharCount;

	sFixed=sStr.toLowerCase();
	var	sWild=sPatt.toLowerCase();

	for(nCharCount=0; nCharCount<sWild.length;++nCharCount)
	{
		switch(sWild.charAt(nCharCount))
		{
			case '?':	// 1 char
				sPattern+='.';
				break;
			case '%':	// 0+ chars
				sPattern+='.*';
				break;
			default:
				sPattern+='['+sWild.charAt(nCharCount)+']';
		}
	}
	var	oRegExp=new RegExp(sPattern);
	return oRegExp.test(sFixed);
}


//
// See if 1st value is between the other two
// Returns true if a is between b and c(inclusive),
//  false if it is not,or null if any is non-numeric.
//
function jsbetween(nTest,nLower,nUpper)
{
	return(nTest>=nLower)&&(nTest<=nUpper);
}

//
// Returns true if the supplied field ID has an answer supplied
//
function jsisknown(sFieldID,nRepeat)
{
	// Default Repeat No for non RQG questions
	nRepeat=DefaultRepeatNo(nRepeat,true);
	
	if(!IsFieldOnForm(sFieldID,nRepeat))
	{
		// to match arezzo
		//bEvaluationError=true;
		return false;
	}
	if(isNaN(nRepeat))
	{
		return false;
	}
	if((oForm.olQuestion[sFieldID].olRepeat[nRepeat].vValue==null)
		||(oForm.olQuestion[sFieldID].olRepeat[nRepeat].vValue==undefined))
	{
		return false;
	}
	switch(oForm.olQuestion[sFieldID].nType)
	{
		case etRealNumber:
		case etIntegerNumber:
			//ic 07/01/2004 use identity operator to compare for empty field but cater for 0
			return(!(oForm.olQuestion[sFieldID].olRepeat[nRepeat].vValue===""))
			break;
		default:
			return(oForm.olQuestion[sFieldID].olRepeat[nRepeat].vValue!="")
	}
}

//
// Arezzo: status(question)
// Returns one of: 'n_a','u_o','known','unknown'
//

function jsstatus(question,nRepeat)
{
	// Default Repeat No for non RQG questions
	nRepeat=DefaultRepeatNo(nRepeat,true);
	
	if(!IsFieldOnForm(sFieldID,nRepeat))
	{
		bEvaluationError=true;
		return 'unknown';
	}
	if(isNaN(nRepeat))
	{
		return 'unknown';
	}
	switch(oForm.olQuestion[sFieldID].olRepeat[nRepeat].nStatus)
	{
		case eStatus.Requested:
			return 'unknown';
		case eStatus.Unobtainable:
			return 'u_o';
		case eStatus.NotApplicable:
			return 'n_a';
		case eStatus.Success:
		case eStatus.Warning:
		case eStatus.OKWarning:
		case eStatus.Inform:
		case eStatus.Missing:
		default:
			return 'known';
	}
}

//
// Arezzo: datepart(timeunit,date)
// Returns an integer representing the value of the requested time unit within the
// date,e.g. jsdatepart('y',(new Date(1966,11,25)).getTime())==1966
//
function jsdatepart(sTimeunit,nMilliseconds)
{
	if(isNaN(nMilliseconds))
	{
		bEvaluationError=true;
		return Number.NaN;
	}

	var	oDate=new Date(nMilliseconds);
	switch(sTimeunit)
	{
		case 'y':
			return(oDate.getFullYear());
		case 'm':
			return(oDate.getMonth()+1);
		case 'd':
			return(oDate.getDate());
		case 'h':
			return(oDate.getHours());
		case 'mm':
			return(oDate.getMinutes());
		case 's':
			return(oDate.getSeconds());
		default:
			bEvaluationError=true;
			return(Number.NaN);
	}
}

//
// Arezzo: date_diff(timeunit,datetime,datetime)
// Returns the calendar difference between the datetimes in the unit specified,
// e.g. 'h',10:59,11:01 returns 1
//
function jsdatediff(sTimeunit,nMilliseconds1,nMilliseconds2)
{
	if(isNaN(nMilliseconds1)
		|| isNaN(nMilliseconds2))
	{
		bEvaluationError=true;
		return Number.NaN;
	}

	var	oDate1=new Date(0);
	var	oDate2=new Date(0);

	oDate1.setTime(nMilliseconds1);
	oDate2.setTime(nMilliseconds2);

	switch(sTimeunit)
	{
		case 'y':
			oDate1.setMonth(0);
			oDate2.setMonth(0);
		case 'm':
			oDate1.setDate(1);
			oDate2.setDate(1);
		case 'w':
		case 'd':
			oDate1.setHours(0);
			oDate2.setHours(0);
		case 'h':
			oDate1.setMinutes(0);
			oDate2.setMinutes(0);
		case 'mm':
			oDate1.setSeconds(0);
			oDate2.setSeconds(0);
			break;
		case 's':
			break;
		default:
			bEvaluationError=true;
			return Number.NaN;
	}

	// DPH 14/01/2003 - Use UTC dates to avoid daylight saving problem
	// Date.UTC() returns date in milliseconds
	// DPH 24/02/2003 - Force completely to UTC format
	var nUTCMs1=Date.UTC(oDate1.getUTCFullYear(),oDate1.getUTCMonth(),oDate1.getUTCDate(),oDate1.getUTCHours(),oDate1.getUTCMinutes(),oDate1.getUTCSeconds());
	var nUTCMs2=Date.UTC(oDate2.getUTCFullYear(),oDate2.getUTCMonth(),oDate2.getUTCDate(),oDate2.getUTCHours(),oDate2.getUTCMinutes(),oDate2.getUTCSeconds());
	return(jstimediff(sTimeunit,nUTCMs1,nUTCMs2,true));
	//return(jstimediff(sTimeunit,oDate1.getTime(),oDate2.getTime()));
}

//
// Arezzo: time_diff(timeunit,datetime,datetime)
// Returns the elapsed time difference between the datetimes in the unit specified,
// e.g. 'h',10:59,11:01 returns 0
//
// changed to take
function jstimediff(sTimeunit,nMilliseconds1,nMilliseconds2,bUTC)
{
	if(isNaN(nMilliseconds1)
		|| isNaN(nMilliseconds2))
	{
		bEvaluationError=true;
		return Number.NaN;
	}

	var	nDifference;
	var nTestMillisec;
	var	oDate=new Date(0);
	var	nDays;

	// if not already in UTC format
	bUTC=(bUTC==undefined)?false:bUTC;
	if(!bUTC)
	{
		var oLocal1 = new Date(nMilliseconds1);
		var oLocal2 = new Date(nMilliseconds2);
		// DPH 24/02/2003 - Force completely to UTC format
		// convert to UTC
		nMilliseconds1=Date.UTC(oLocal1.getUTCFullYear(),oLocal1.getUTCMonth(),oLocal1.getUTCDate(),oLocal1.getUTCHours(),oLocal1.getUTCMinutes(),oLocal1.getUTCSeconds());
		nMilliseconds2=Date.UTC(oLocal2.getUTCFullYear(),oLocal2.getUTCMonth(),oLocal2.getUTCDate(),oLocal2.getUTCHours(),oLocal2.getUTCMinutes(),oLocal2.getUTCSeconds());
	}
	// dph - Make daylight saving friendly
	var nAdjust=jsDetectDaylightSavings(nMilliseconds1,nMilliseconds2);
	nMilliseconds1=jsAdjustTime(nMilliseconds1,nAdjust);

	if(nMilliseconds1 > nMilliseconds2)
	{
		//swap the arguments round so that 2nd is aways bigger/later
		var	nTemp=nMilliseconds1;
		nMilliseconds1=nMilliseconds2;
		nMilliseconds2=nTemp;
	}

	nDifference=nMilliseconds2-nMilliseconds1;
	oDate.setTime(nMilliseconds2);

	switch(sTimeunit)
	{
		case 'y':	// since years can contain different numbers of milliseconds,difference must be computed by repeatedly subtracting years from the second date until it is before the first
			nDifference=0;
			do
			{
				nDifference++;
				oTestDate=new Date(oDate.getTime());
				oTestDate.setFullYear(oTestDate.getFullYear()-nDifference);
				if(oTestDate.getMonth()!=oDate.getMonth())
				{
					// The later argument was on a leap day,but the test date isn't.
					// JavaScript will have converted the date to 2 Mar... move it to 28 Feb
					oTestDate.setDate(28);
					oTestDate.setMonth(1);	//Feb=1
				}
				nTestMillisec=oTestDate.getTime();
				// if have adjusted date1 for daylight savings
				// check if also need to adjust test date
				var nAdjustTest=jsDetectDaylightSavings(nTestMillisec,nMilliseconds2);
				if((nAdjust!=0)&&(nAdjustTest!=0))
				{
					nTestMillisec=jsAdjustTime(nTestMillisec,nAdjustTest);
				}
			} while(nTestMillisec>=nMilliseconds1)
			return --nDifference;

		case 'm':	// similarly for months
			nDifference=0;
			do
			{
				++nDifference;
				nDays=0;
				do
				{
					oTestDate=new Date(oDate.getTime());
					oTestDate.setDate(oTestDate.getDate()-nDays);
					oTestDate.setFullYear(oTestDate.getFullYear()-Math.floor((nDifference+11-oTestDate.getMonth())/ 12));
					oTestDate.setMonth(11-(nDifference+11-oTestDate.getMonth())% 12);
				} while(oTestDate.getDate()!=oDate.getDate()-nDays++);
				nTestMillisec=oTestDate.getTime();
				// if have adjusted date1 for daylight savings
				// check if also need to adjust test date
				var nAdjustTest=jsDetectDaylightSavings(nTestMillisec,nMilliseconds2);
				if((nAdjust!=0)&&(nAdjustTest!=0))
				{
					nTestMillisec=jsAdjustTime(nTestMillisec,nAdjustTest);
				}
			} while(nTestMillisec>=nMilliseconds1)
			return --nDifference;

		//all other time units are a fixed number of milliseconds,so can simply divide...
		case 'w':
			nDifference/=7;
		case 'd':
			nDifference/=24;
		case 'h':
			nDifference/=60;
		case 'mm':
			nDifference/=60;
		case 's':
			nDifference/=1000;
			break;
		default:
			bEvaluationError=true;
			return Number.NaN;
	}
	return Math.floor(nDifference);
}

//
// Arezzo: datetime + interval timeunit
// NB: nInterval can be +ve or -ve.
// Partially rewritten by dph because of daylight saving issue
// new version adds interval to appropriate datepart rather than adding 
// millisecond intervals together (as can be 25 hours on daylight saving day)
//
function jsdateadd(sTimeunit,nInterval,nMilliseconds)
{
	if(isNaN(nMilliseconds)||isNaN(nInterval))
	{
		bEvaluationError=true;
		return Number.NaN;
	}

	var	nDays=0;
	var	oDate=new Date(0);
	var	oTestDate=new Date(0);
	var	oResult=new Date(0);

	oDate.setTime(nMilliseconds);
	oTestDate.setTime(nMilliseconds);
	oResult.setTime(nMilliseconds);

	switch(sTimeunit){
		case 'y':
			// dph - part rewritten as could infinite loop
			oResult.setTime(oDate.getTime());
			//set day to 1 to avoid 29 Feb problem when not leap year
			oResult.setDate(1);
			//oResult.setDate(oResult.getDate()-nDays);
			oResult.setFullYear(oResult.getFullYear()+nInterval);
			//store date without day update so can compare
			var	oTestDate=new Date(oResult);
			// set day and then compare months to make sure not gone to next one
			oResult.setDate(oDate.getDate());
			if(oResult.getMonth()!=oTestDate.getMonth())
			{
				//have rolled into next month due to day value
				//now need to set max day/month value
				oResult=new Date(oTestDate);
				oResult.setDate(jsMaxMonthDay(oTestDate.getMonth()));
			}
			break;
			//return oResult.getTime();
		case 'm':
			// dph - part rewritten as could infinite loop
			if(nInterval>=0)
			{
				nYearDiff=Math.floor((nInterval+oDate.getMonth())/ 12);
				nNewMonth=(nInterval+oDate.getMonth())% 12;
			}
			else
			{
				nYearDiff=-Math.floor((11-nInterval-oTestDate.getMonth())/ 12);
				nNewMonth=11-(11-nInterval-oTestDate.getMonth())% 12;
			}
			oResult.setTime(oDate.getTime());
			//set day to 1 to avoid 31 Feb problem
			oResult.setDate(1);
			//oResult.setDate(oResult.getDate());
			oResult.setFullYear(oResult.getFullYear()+nYearDiff);
			oResult.setMonth(nNewMonth);
			//store date without day update so can compare
			var	oTestDate=new Date(oResult);
			// set day and then compare months to make sure not gone to next one
			oResult.setDate(oDate.getDate());
			if(oResult.getMonth()!=oTestDate.getMonth())
			{
				//have rolled into next month due to day value
				//now need to set max day/month value
				oResult.setDate(1);
				oResult.setMonth(nNewMonth);
				oResult.setDate(jsMaxMonthDay(nNewMonth));
			}
			break;
			//return oResult.getTime();
		case 'w':
			// weeks to days
			nInterval*=7;
		case 'd':
			oResult.setDate(oDate.getDate()+nInterval);	
			break;
		case 'h':
			// time only - convert to UTC - set hours - convert back ...
			var oUTCDate=new Date(Date.UTC(oResult.getFullYear(),oResult.getMonth(),oResult.getDate(),oResult.getHours(),oResult.getMinutes(),oResult.getSeconds()));
			oUTCDate.setUTCHours(oUTCDate.getUTCHours()+nInterval);
			oResult=new Date(oUTCDate.getUTCFullYear(),oUTCDate.getUTCMonth(),oUTCDate.getUTCDate(),oUTCDate.getUTCHours(),oUTCDate.getUTCMinutes(),oUTCDate.getUTCSeconds());
			break;
		case 'mm':
			// time only - convert to UTC - set minutes - convert back ...
			var oUTCDate=new Date(Date.UTC(oResult.getFullYear(),oResult.getMonth(),oResult.getDate(),oResult.getHours(),oResult.getMinutes(),oResult.getSeconds()));
			oUTCDate.setUTCMinutes(oUTCDate.getUTCMinutes()+nInterval);
			oResult=new Date(oUTCDate.getUTCFullYear(),oUTCDate.getUTCMonth(),oUTCDate.getUTCDate(),oUTCDate.getUTCHours(),oUTCDate.getUTCMinutes(),oUTCDate.getUTCSeconds());
			break;
		case 's':
			// time only - convert to UTC - set seconds - convert back ...
			var oUTCDate=new Date(Date.UTC(oResult.getFullYear(),oResult.getMonth(),oResult.getDate(),oResult.getHours(),oResult.getMinutes(),oResult.getSeconds()));
			oUTCDate.setUTCSeconds(oUTCDate.getUTCSeconds()+nInterval);
			oResult=new Date(oUTCDate.getUTCFullYear(),oUTCDate.getUTCMonth(),oUTCDate.getUTCDate(),oUTCDate.getUTCHours(),oUTCDate.getUTCMinutes(),oUTCDate.getUTCSeconds());
			break;
		default:
			bEvaluationError=true;
			return Number.NaN;
	}
	return(oResult.getTime());
}

function jsMaxMonthDay(nMonth)
{
	var nDay=0;
	// zero based so add 1 to make clearer
	switch(nMonth+1)
	{
		case 1:
		case 3:
		case 5:
		case 7:
		case 8:
		case 10:
		case 12: 
			{
				// 31 days Jan/Mar/May/Jul/Aug/Oct/Dec 
				nDay=31;
				break;
			}
		case 4:
		case 6:
		case 9:
		case 11: 
			{
				// 30 days Apr/Jun/Sep/Nov 
				nDay=30;
				break; 
			}
		case 2: 
			{
				// 28 days for Feb 
				nDay=28; 
				break;
			}
		default:
			{
				nDay=1;
			}
	}
	return nDay;
}

//
// Arezzo: Question1 & question2
//
function jsconcat(sStr1,sStr2)
{
	return(""+sStr1+sStr2);
}

//
// Do an Arezzo comparison,based on the supplied comparator("==",">=",">",etc.)
//
function jsCompare(vInValue1,vInValue2,sComparator)
{
	if((vInValue1==undefined)||(vInValue2==undefined))
	{
		bEvaluationError=true;
		return null;
	}

	switch(typeof(vInValue1))
	{
		case "string":
			// All non-numerics are treated as lower-case quoted strings
			// quoted in " to allow for ' in data
			vValue1='"'+vInValue1.toLowerCase()+'"';
			vValue2='"'+(vInValue2.toString()).toLowerCase()+'"';
			break;
		case "object":
			vValue1=vInValue1.valueOf();
			vValue2=vInValue2.valueOf();
			break;
		default:
			vValue1=1 * vInValue1;
			vValue2=1 * vInValue2;
	}
	return(eval(vValue1+" "+sComparator+" "+vValue2));
}

//
// Do an arezzo "len"
//
function jslen(sValue)
{
	sValue=""+sValue;
	// dph/ic 16/02/2004 do not return length of an empty string to match windows
	if(sValue==="")
	{
		return("");
	}
	else
	{
		return(sValue.length);
	}
}

//
// Do an arezzo "Squareroot"
//
function jssquareroot(nValue)
{
	nValue=1 * nValue;
	return(Math.sqrt(nValue));
}

//
// Do an Arezzo "substring"
//
function jssubstring(sString,nStart,nLength)
{
	sString+="";
//  return sString.substr(nStart,nLength);	// Not a good idea - I.E. 4 does not work correctly with this
	if(nStart<0)
	{
		// dph corrected -ve jssubstring calculation +nstart as will be -ve
		//		also string.substr does not work properly in IE so use string.substring
		//return(sString.substr(sString.length-(nStart-1),nLength));
		var nStartPos=sString.length+nStart;
		if(nStartPos<0)
		{
			nStartPos=0;
		}
		var nEndPos=nStartPos+nLength;
		return(sString.substring(nStartPos,nEndPos));
	}
	else
	{
		return(sString.substr((nStart-1),nLength));
	}
}

//
// Do an Arezzo "="
//
function jseq(sFieldID1,sFieldID2)
{
	var bRet=(jsCompare(sFieldID1,sFieldID2,"=="));
	// to match arezzo
	if(bEvaluationError)
	{
		bRet=false;
		bEvaluationError=false;
	}
	return(bRet);
}

//
// Do an Arezzo "<>"
//
function jsne(sFieldID1,sFieldID2)
{
	var bRet=(jsCompare(sFieldID1,sFieldID2,"!="));
	// to match arezzo
	if(bEvaluationError)
	{
		bRet=false;
		bEvaluationError=false;
	}
	return(bRet);
}

//
// Do an Arezzo "<"
//
function jslt(sFieldID1,sFieldID2)
{
	//ic 15/06/2006 issue 2674
	if(sFieldID1==="") return(false);
	if(sFieldID2==="") return(false);
	
	sFieldID1*=1;
	sFieldID2*=1;
	var bRet=(jsCompare(sFieldID1,sFieldID2,"<"));
	// to match arezzo
	if(bEvaluationError)
	{
		bRet=false;
		bEvaluationError=false;
	}
	return(bRet);
}

//
// Do an Arezzo "<="
//
function jsle(sFieldID1,sFieldID2)
{
	//ic 15/06/2006 issue 2674
	if(sFieldID1==="") return(false);
	if(sFieldID2==="") return(false);
	
	sFieldID1*=1;
	sFieldID2*=1;
	var bRet=(jsCompare(sFieldID1,sFieldID2,"<="));
	// to match arezzo
	if(bEvaluationError)
	{
		bRet=false;
		bEvaluationError=false;
	}
	return(bRet);
}

//
// Do an Arezzo ">"
//
function jsgt(sFieldID1,sFieldID2)
{
	//ic 15/06/2006 issue 2674
	if(sFieldID1==="") return(false);
	if(sFieldID2==="") return(false);
	
	sFieldID1*=1;
	sFieldID2*=1;
	var bRet=(jsCompare(sFieldID1,sFieldID2,">"));
	// to match arezzo
	if(bEvaluationError)
	{
		bRet=false;
		bEvaluationError=false;
	}
	return(bRet);
}

//
// Do an Arezzo ">="
//
function jsge(sFieldID1,sFieldID2)
{
	//ic 15/06/2006 issue 2674
	if(sFieldID1==="") return(false);
	if(sFieldID2==="") return(false);
	
	sFieldID1*=1;
	sFieldID2*=1;
	var bRet=(jsCompare(sFieldID1,sFieldID2,">="));
	// to match arezzo
	if(bEvaluationError)
	{
		bRet=false;
		bEvaluationError=false;
	}
	return(bRet);
}

//
// New Arezzo functionality for RQGs
//
// Get a particular questions Repeat No 
//
function jsRepNo(sFieldID,sRepNo)
{
	var vRepNo=null;
	var vTempRep=null;
	// dph 10/03/2004 - make sure repeat number is numeric
	gnCurrentRepeat*=1;
	if(!IsFieldOnForm(sFieldID))
	{
		return Number.NaN;
	}
	var oFieldTemp=oForm.olQuestion[sFieldID];
	if((oFieldTemp.olRepeat!=null)&&(oFieldTemp.olRepeat!=undefined))
	{
		switch(sRepNo)
		{
			case "first":
			{
				vTempRep=0;
				break;
			}
			case "last":
			{
				vTempRep=(oFieldTemp.olRepeat.length)-1;
				if(vTempRep<0)
				{
					vTempRep=null;
				}
				break;
			}
			case "next":
			{
				if((gnCurrentRepeat==null)||(gnCurrentRepeat==undefined))
				{
					break;
				}
				var vTempRep=gnCurrentRepeat+1;
				break;
			}
			case "previous":
			{
				if((gnCurrentRepeat==null)||(gnCurrentRepeat==undefined))
				{
					break;
				}
				var vTempRep=gnCurrentRepeat-1;
				break;
			}
			case "this":
			{
				if((gnCurrentRepeat==null)||(gnCurrentRepeat==undefined))
				{
					break;
				}
				var vTempRep=gnCurrentRepeat;
				break;
			}
			default:
			{
				// if have been supplied with a number then zero base it
				// i.e. sRepNo - 1
				if(!isNaN(sRepNo))
				{
					vTempRep=sRepNo-1;
				}
				else
				{
					vRepNo=null;
				}
				break;
			}
		}
		// Check if RepNo & response exists
		if((vTempRep!=null)&&(vTempRep>=0))
		{
			if((oFieldTemp.olRepeat[vTempRep]!=null)&&(oFieldTemp.olRepeat[vTempRep]!=undefined))
			{
				vRepNo=vTempRep;
			}
		}
	}

	if(vRepNo==null)
	{
		vRepNo=Number.NaN;
	}
	else
	{
		// vRepNo + 1 to keep in line with windows numbering
		vRepNo++;
	}
	return vRepNo;
}

//
// Arezzo 'Count'
//
function jsQCount(sFieldID)
{
	if(!IsFieldOnForm(sFieldID))
	{
		bEvaluationError=true;
		return 0;
	}
	// return number of non-empty responses
	var nRows=oForm.olQuestion[sFieldID].olRepeat.length;
	var nCount=0;
	for(var i=0;i<nRows;i++)
	{
		if(!fnIsFieldEmpty(sFieldID,i))
		{
			nCount++;
		}
	}
	return nCount;
}

//
// Arezzo 'Max'
//
function jsQMax(sFieldID)
{
	if(!IsFieldOnForm(sFieldID))
	{
		bEvaluationError=true;
		return Number.NaN;
	}
	var vMax=null;
	var vVal;
	// Loop through responses and get Max
	var nRows=oForm.olQuestion[sFieldID].olRepeat.length;
	for(var i=0;i<nRows;i++)
	{
		vVal=oForm.olQuestion[sFieldID].olRepeat[i].get();
		//ic 09/01/2004 compare formatted value using identity to handle zeroes
		if(!isNaN(vVal))
		{
			if(!(oForm.olQuestion[sFieldID].olRepeat[i].getFormatted()===""))
			{
				if(vMax==null)
				{
					vMax=vVal;
				}
				else
				{
					//if(vMax<vVal)
					if(jslt(vMax,vVal))
					{
						vMax=vVal;
					}
				}
			}
		}
	}
	if(vMax==null)
	{
		vMax=Number.NaN;
	}
	return vMax;
}

//
// Arezzo 'Min'
//
function jsQMin(sFieldID)
{
	if(!IsFieldOnForm(sFieldID))
	{
		bEvaluationError=true;
		return Number.NaN;
	}
	var vMin=null;
	var vVal;
	// Loop through responses and get Min
	var nRows=oForm.olQuestion[sFieldID].olRepeat.length;
	for(var i=0;i<nRows;i++)
	{
		vVal=oForm.olQuestion[sFieldID].olRepeat[i].get();
		//ic 09/01/2004 compare formatted value using identity to handle zeroes
		if(!isNaN(vVal))
		{
			if(!(oForm.olQuestion[sFieldID].olRepeat[i].getFormatted()===""))
			{
				if(vMin==null)
				{
					vMin=vVal;
				}
				else
				{
					//if(vMin>vVal)
					if(jsgt(vMin,vVal))
					{
						vMin=vVal;
					}
				}
			}
		}
	}
	if((vMin==null)||(isNaN(vMin)))
	{
		vMin=Number.NaN;
	}
	return vMin;
}

//
// Arezzo 'Avg'
//
function jsQAvg(sFieldID)
{
	if(!IsFieldOnForm(sFieldID))
	{
		bEvaluationError=true;
		return Number.NaN;
	}
	var vCount=jsQCount(sFieldID);
	var vSum=jsQSum(sFieldID);
	var vAvg=null;
	if((!isNaN(vCount))&&(!isNaN(vSum)))
	{
		//ic/dph 09/02/2004 use jsDivide()
		vAvg=jsDivide(vSum,vCount);
	}
	if(vAvg==null)
	{
		vAvg=Number.NaN;
	}
	return vAvg;
}

//
// Arezzo 'Sum'
//
function jsQSum(sFieldID)
{
	if(!IsFieldOnForm(sFieldID))
	{
		bEvaluationError=true;
		return Number.NaN;
	}
	var vSum=null;
	// Loop through responses and get Sum
	var nRows=oForm.olQuestion[sFieldID].olRepeat.length;
	for(var i=0;i<nRows;i++)
	{
		vVal=oForm.olQuestion[sFieldID].olRepeat[i].get();
		if((!isNaN(vVal))&&(!(vVal==="")))
		{
			if(vSum==null)
			{
				vSum=vVal;
			}
			else
			{
				//ic/dph 09/02/2004 use jsAdd()
				vSum=jsadd(vSum,vVal);
			}
		}
	}	
	if(vSum==null)
	{
		vSum=Number.NaN;
	}
	return vSum;
}

//
// Arezzo 'True'
//
function jsTrue()
{
	return true;
}

//
// Arezzo 'False'
//
function jsFalse()
{
	return false;
}

// jsAuthoriser included for completeness - not active yet
function jsAuthoriser(sFieldID, nRepeat)
{
	// Default Repeat No for non RQG questions
	nRepeat=DefaultRepeatNo(nRepeat,true);
	
	var oFieldTemp=oForm.olQuestion[sFieldID];
	var oField=oFieldTemp.olRepeat[nRepeat];
	
	// if the field authorising from is not empty and has a 'saved' value
	if((!fnIsFieldEmpty(sFieldID,nRepeat))&&(oField.vDBFormatted!=""))
	{
		return oField.sUserFull;
	}
	else
	{
		return "";
	}
}

// jsReadDate - 
function jsReadDate(sDate, sFormat)
{
	var oDate=fnParseDate(sDate,sFormat);
	// will be null if fails
	return oDate;
}

// jsValidDate - returns true if valid else false
function jsValidDate(sDate, sFormat)
{
	var bValid=true;
	var oDate=fnParseDate(sDate,sFormat);
	if(oDate==null)
	{
		bValid=false;
	}
	return bValid;
}

// jsformatdate - 
function jsformatdate(dateDate, sFormat)
{
	if(dateDate==null)
	{
		return "";
	}
	var	iOffset;
	var	sFormat=sFormat.toUpperCase();
	var	sDate;
	var	dDate=dateDate;
	if(typeof(dDate)!="object")
	{
		dDate=new Date(dateDate);
	}

	iOffset=fnStrPos(sFormat,"DD",0);
	if(iOffset>=0)
	{
		sDate=dDate.getDate()+100+"X";
		sDate=sDate.substr(1,2);
		sFormat=sFormat.substr(0,iOffset)+sDate+sFormat.substring(iOffset+2,sFormat.length);
	}
	iOffset=fnStrPos(sFormat,"[^H].MM",2);
	if(iOffset>=0)
	{
		sDate=dDate.getMonth()+101+"X";
		sDate=sDate.substr(1,2);
		sFormat=sFormat.substr(0,iOffset)+sDate+sFormat.substring(iOffset+2,sFormat.length);
	}
	else
	{
		iOffset=fnStrPos(sFormat,"^MM",0);
		if(iOffset>=0)
		{
			sDate=dDate.getMonth()+101+"X";
			sDate=sDate.substr(1,2);
			sFormat=sFormat.substr(0,iOffset)+sDate+sFormat.substring(iOffset+2,sFormat.length);
		}
	}
	iOffset=fnStrPos(sFormat,"YYYY",0);
	if(iOffset>=0)
	{
		sDate=dDate.getYear();
		if(sDate<100)
		{
			sDate+=1900;
		}
		sDate=sDate+10000+"X";
		sDate=sDate.substr(1,4);
		sFormat=sFormat.substr(0,iOffset)+sDate+sFormat.substring(iOffset+4,sFormat.length);
	}
	iOffset=fnStrPos(sFormat,"YY",0);
	if(iOffset>=0)
	{
		sDate=dDate.getYear()+10000+"X";
		sDate=sDate.substr(3,2);
		sFormat=sFormat.substr(0,iOffset)+sDate+sFormat.substring(iOffset+2,sFormat.length);
	}
	iOffset=fnStrPos(sFormat,"H.MM",2);
	if(iOffset>=0)
	{
		sDate=dDate.getMinutes()+100+"X";
		sDate=sDate.substr(1,2);
		sFormat=sFormat.substr(0,iOffset)+sDate+sFormat.substring(iOffset+2,sFormat.length);
	}
	iOffset=fnStrPos(sFormat,"HH",0);
	if(iOffset>=0)
	{
		sDate=dDate.getHours()+100+"X";
		sDate=sDate.substr(1,2);
		sFormat=sFormat.substr(0,iOffset)+sDate+sFormat.substring(iOffset+2,sFormat.length);
	}
	iOffset=fnStrPos(sFormat,"SS",0);
	if(iOffset>=0)
	{
		sDate=dDate.getSeconds()+100+"X";
		sDate=sDate.substr(1,2);
		sFormat=sFormat.substr(0,iOffset)+sDate+sFormat.substring(iOffset+2,sFormat.length);
	}
	return sFormat;
}

// jsFullUserName - returns full user name of logged in user
function jsFullUserName()
{
	var sFullUserName=fnGetFormProperty("sFullUserName");
	if(sFullUserName==null)
	{
		sFullUserName="";
	}
	return sFullUserName;
}

// jsUserRole - returns user role of logged in user
function jsUserRole()
{
	var sUserRole=fnGetFormProperty("sUserRole");
	if(sUserRole==null)
	{
		sUserRole="";
	}
	return sUserRole;
}

// rounds lNumber to n decimal places
function jsround(lNumber,n) 
{
	if(fnDecimalPlaces(lNumber)>n)
	{
		lNumber=Math.round(lNumber*Math.pow(10,n))/Math.pow(10,n);
	}
	return lNumber;
}

//
// AREZZO function - calculates minimum of two given values
// only return a value if both values exist
//
function jsQMin2(vValue1,vValue2)
{
	//ic 09/01/2004 check for evaluation errors
	if(!bEvaluationError)
	{
		var vMin=null;
		var vVal;
		// check both values are numbers
		//ic 09/01/2004 compare vValue with identity operator to handle zeroes
		if((isNaN(vValue1))||(vValue1==="")||(isNaN(vValue2))||(vValue2===""))
		{
			bEvaluationError=true;
			vMin=Number.NaN;
			return vMin;
		}

		// vValue1 is (1st) min value
		vMin=vValue1;
	
		// now compare 2nd value
		//if (vMin>vValue2)
		if(jsgt(vMin,vValue2))
		{
			vMin=vValue2;
		}
	
		if((vMin==null)||(isNaN(vMin)))
		{
			vMin=Number.NaN;
		}
		return vMin;
	}
}

//
// calculates maximum of two given values
// only return a value if both values exist
//
function jsQMax2(vValue1,vValue2)
{
	//ic 09/01/2004 check for evaluation errors
	if(!bEvaluationError)
	{
		var vMax=null;
		var vVal;

		// check both values are numbers
		//ic 09/01/2004 compare vValue with identity operator to handle zeroes
		if((isNaN(vValue1))||(vValue1==="")||(isNaN(vValue2))||(vValue2===""))
		{
			bEvaluationError=true;
			vMax=Number.NaN;
			return vMax;
		}

		// vValue1 is (1st) max value
		vMax=vValue1;
	
		// now compare 2nd value
		//if (vMax<vValue2)
		if(jslt(vMax,vValue2))
		{
			vMax=vValue2;
		}
	
		if((vMax==null)||(isNaN(vMax)))
		{
			vMax=Number.NaN;
		}
		return vMax;
	}
}

//
// Javascript Date and time calculated from a given date and a given time
//
function jsDateAndTime(dateDate,dateTime)
{
	// quit if null dates
	if((dateDate==null)||(dateTime==null))
	{
		return null;
	}
	if(typeof(dateDate)!="object")
	{
		var dateDT = new Date(dateDate);
	}
	else
	{
		var dateDT=dateDate;
	}
	if(typeof(dateTime)!="object")
	{
		var dateT = new Date(dateTime);
	}
	else
	{
		var dateT=dateTime;
	}
	// set minutes section of date
	dateDT.setHours(dateT.getHours(),dateT.getMinutes(),dateT.getSeconds(),dateT.getMilliseconds());

	return dateDT;
}

//
// Javascript base 10 Logarithm of given number
//
function jsLog(nNumber)
{
	if(typeof(nNumber)!="number")
	{
		return "";
	}
	return (Math.LOG10E * Math.log(nNumber));
}

//
// javascript square root of passed number
//
function jssqrt(nNumber)
{
	if(typeof(nNumber)!="number")
	{
		return "";
	}
	return (Math.sqrt(nNumber));
}

//
// Calculate any/every Arezzo function
// eg. jseq( jsValueOf( "f_10018_10359", jsRepNo( "f_10018_10359", "1" ) , 1 ) )
//
function jsAnyEvery(sType,sFieldID,sJSComparison,vValue)
{
	// exit if question not on form
	if(oForm.olQuestion[sFieldID]==null)
	{
		return false;
	}
	var bResult;
	var nRepeats=oForm.olQuestion[sFieldID].olRepeat.length;
	var sEval;
	if(sType=="any")
	{
		// "any" - false unless get a true
		bResult=false;
	}
	else
	{
		// "every" - true unless get a false
		bResult=true;
	}
	for(var i=0;i<nRepeats;i++)
	{
		// only do comparison if field is not empty
		if(!fnIsFieldEmpty(sFieldID,i))
		{
			// of type - comparison ( jsValueOf( "f_10018_10359", jsRepNo( "f_10018_10359", "repeat" ) , value )
			sEval=sJSComparison
			sEval+="( jsValueOf( \""+sFieldID+"\", jsRepNo( \""+sFieldID+"\", ";
			sEval+="\""+(i+1)+"\" ) ) , ";
			// decide if need quotes for value
			if(typeof(vValue)=="string")
			{
				sEval+="\""+vValue+"\"";
			}
			else
			{
				sEval+=vValue;
			}
			sEval+=" ) "
			//evaluate
			bResult=eval(sEval);
			if(bEvaluationError)
			{
				return false;
			}
			// check type "any"/"every" and see if can quit
			if(sType=="any")
			{
				// "any" - if any expressions are true return true
				if(bResult)
				{
					return true;
				}
			}
			if(sType=="every")
			{
				// "every" - all expressions must be true to return true
				if(!bResult)
				{
					return false;
				}
			}
		}
	}
	return bResult;
}

//multiply two numbers, round to sum of decimal places
function jsMultiply(n1,n2)
{
	n1*=1;
	n2*=1;
	return (jsround((n1*n2),(fnDecimalPlaces(n1)+fnDecimalPlaces(n2))))
}

//divide one number by another, round to 14dp
function jsDivide(n1,n2)
{
	n1*=1;
	n2*=1;
	return (n1/n2);
}

// adds 2 numerics together, round to max decimal places
function jsadd(n1,n2)
{
	n1*=1;
	n2*=1;
	return (jsround((n1+n2),fnMax(fnDecimalPlaces(n1),fnDecimalPlaces(n2))));
}

//subtract 1 numeric from another, round to max decimal places
function jsSubtract(n1,n2)
{
	n1*=1;
	n2*=1;
	return (jsround((n1-n2),fnMax(fnDecimalPlaces(n1),fnDecimalPlaces(n2))));
}

//exponent of number
function jsExp(nNum,nExp)
{
	nNum*=1;
	nExp*=1;
	var nRtn=(Math.pow(nNum,nExp));;
	
	//for positive integers, round to the ((number of decimal places in nNum)*nExp) decimal places
	if((nExp>0)&&((nExp%1)==0))
	{
		nRtn=(jsround(nRtn,(jsMultiply(nExp,fnDecimalPlaces(nNum)))));
	}

	return nRtn;
}

//negative of number
function jsNeg(n)
{
	return (n<0)? n:jsMultiply(n,-1);
}

//case sensitive comparison of text (equivelent to arezzo '==')
function jsCaseEq(sText1,sText2)
{
	return (sText1==sText2);
}

//--------------------------------------------------------------------------------

// detect if there is a daylights savings difference between 2 dates
function jsDetectDaylightSavings(nMilliSec1,nMilliSec2)
{
	// if in timezone calculate value
	var lTimezone1=jsGetTimezone(nMilliSec1);
	var lTimezone2=jsGetTimezone(nMilliSec2);
	// return difference
	return lTimezone1-lTimezone2;
}
// get timezone value for a date
function jsGetTimezone(nMilliSec)
{
	var oLocal = new Date(nMilliSec);
	// if in timezone calculate value
	var lTimezone=oLocal.getTimezoneOffset();
	return lTimezone;
}
// Adjust time by nDiff minutes (for use when datediffing with daylight savings)
function jsAdjustTime(nMilliSec,nDiff)
{
	// add 1 hour
	var nAdjust=nMilliSec-(nDiff*60*1000);
	return nAdjust;
}

// Create an internal time to use within the Validation engine given hour, min & sec values
function jsTime(nHours,nMinutes,nSeconds)
{
	var	oDate=new Date(0);	// Use zero to avoid internal date "roll-overs"

	// default a times date to 01/01/1600
	oDate.setFullYear(1600);
	oDate.setMonth(0);
	oDate.setDate(1);

	oDate.setHours(nHours);
	oDate.setMinutes(nMinutes);
	oDate.setSeconds(nSeconds);
	if(oDate.getHours()==nHours
			&& oDate.getMinutes()==nMinutes
			&& oDate.getSeconds()==nSeconds)
		return oDate.getTime();
	bEvaluationError=true;
	return null;
}
// Create an internal time to use within the Validation engine for the current time
function jsTimeNow()
{
	var	oDate=new Date();
	// default a times date to 01/01/1600
	oDate.setFullYear(1600);
	oDate.setMonth(0);
	oDate.setDate(1);
	oDate.setMilliseconds(0);
	return oDate.getTime();
}

//return the number of digits following a decimal point
function fnDecimalPlaces(lNum)
{
	var sNum=lNum+"";
	var nAt=sNum.indexOf(".");
	return (nAt==-1)? 0:(sNum.length-(nAt+1));
}

//return the higher of 2 passed numbers
function fnMax(n1,n2)
{
	n1*=1;
	n2*=1;
	return (n1>n2)? n1:n2;
}

//
// Function to create a new Form object.
//
//ic 19/08/2002
//added sDisabledColour arguement to store disabled element bg colour
function fnInitialiseApplet(oImgFrame,sBlurColour,sFocusColour,oRadioForm,
					sDisabledColour,bUReadOnly,bVReadOnly,sLab,bNewForm)
{
	oForm=new Object();
	oForm.oOtherPage=oImgFrame;
	oForm.bNewForm=bNewForm;
	oForm.sBlurColour=((sBlurColour==null)||(sBlurColour=="")?"#FFFFFF":sBlurColour);
	oForm.sFocusColour=((sFocusColour==null)||(sFocusColour=="")?"#FFFF80":sFocusColour);
	oForm.sDisabledColour=((sDisabledColour==null)||(sDisabledColour=="")?"#EEEEEE":sDisabledColour);
	oForm.bUReadOnly=bUReadOnly; //is user eform read only
	oForm.bVReadOnly=bVReadOnly //is visit eform (if any) read only
	oForm.sLab=sLab; //lab chosen for eform, if any
	oForm.sDecimalPoint=fnDP(); //locale decimal point
	oForm.sThousandSeparator=fnTS(); //locale thousand separator
	oCurrentFieldID=null;	// Need to reset these two each page, or it all goes a bit mauve
	oLastFocusID=null;
	oFocusTarget=null;
	return true;
}
	
//
// Roung the value supplied to match the number of decimal places indicated
//
function fnRound(nValue,nDecimals)
{
	var	nExp=Math.pow(10,nDecimals);
	nValue=nValue * nExp+0.5;
	nValue=(Math.floor(nValue)/ nExp);
	if((nValue>=0)&&(nValue<1))
	{
		nValue=""+nValue;
		nValue=nValue.substr(1);
	}
		nValue=""+nValue;
	return nValue;
}

//
// Function to add the template of the supplied question to the question list
// for the current form.
//
//  sFieldID=field identifier
//  nType=field type (see enumerations "etXXX" at top of code)
//  nLength=field length
//  sFormat=input format(meaning depends on data type)
//  nCase=case sensitivity flag(not currently used)
//  sColour=normal colour of caption
//  sAuthorisation=role of used needed to authorise a change in the field (or blank)
//  bRequiresRFC=flag indicating if "reason for change" is needed on this field
//	nElementID, numeric element id
//  bRQG=Question belongs to Repeating Question Group
//
function fnCT(sFieldID,nType,nLength,sFormat,nCase,
				sColour,sAuthorisation,bRequiresRFC,nQuestionID,bRQG,
				sRQG, bMandatory,sCaptionText,bDisplayStatusIcon,bEform,
				bIsLabField,nDisplayLength,bOptional,sFontStyle,bHidden)
{
	if(oForm.olQuestion==null)
	{
		oForm.olQuestion=new Array();
	}

	if(oForm.olQuestion[sFieldID]==null)
	{
		oForm.olQuestion[sFieldID]=new Object();
	}

	var	oField=oForm.olQuestion[sFieldID]
	oField.sID=sFieldID;
	oField.nType=nType;
	oField.nCase=nCase;
	oField.nLength=nLength;
	if(nType==etRealNumber)
	{
		//Ensure the format is resonable (so "#9#9#.#9" becomes "#9999.99#")
		while(sFormat.search(/^(#*[0-9]+)#/)>=0)
		{
			sFormat=sFormat.replace(/^([^\.]*[0-9]+)#/,"$19")
		}
		while(sFormat.search(/\.([0-9]*)#([0-9]+)/)>=0)
		{
			sFormat=sFormat.replace(/\.([0-9]*)#([0-9]+)/,".$19$2")
		}
	}
	else if(nType==etIntegerNumber)
	{
		while(sFormat.search(/([0-9]+)#(#*)([0-9]+)/)>=0)
		{
			sFormat=sFormat.replace(/([0-9]+)#(#*)([0-9]+)/g,"$19$2$3")
		}
	}
	oField.sFormat=sFormat;
	oField.sColour=sColour;
	oField.sAuthorisation=sAuthorisation;
	oField.bRequiresRFC=bRequiresRFC;
	oField.nQuestionID=nQuestionID;
	oField.bRQG=bRQG;
	oField.sRQG=sRQG;
	oField.bMandatory=bMandatory;
	oField.sCaptionText=sCaptionText;
	oField.bDisplayStatusIcon=bDisplayStatusIcon;
	oField.bEform=bEform;
	oField.bIsLabField=bIsLabField;
	if(nDisplayLength==null)
	{
		// set to length of field
		oField.nDisplayLength=nLength;
	}
	else
	{
		oField.nDisplayLength=nDisplayLength;
	}
	oField.bOptional=bOptional;
	oField.sFontStyle=sFontStyle;
	oField.bHidden=bHidden;
	
	return true;
}

//
// Function to add the supplied question instance to the question repeat list 
// for the current form.
//
//  sFieldID=field identifier
//	nRepeatNo=Repeat Number of question
//  vValue=initial value
//  bEnabled=visibility status (not necessarily the same as the "skip" status)
//  nStatus=MACRO status of field (30,25,20,10,0,-5,-10,etc.)
//  oHandle=handle to HTML object containing this field
//  oCapHandle=handle to HTML object containing this field's caption
//  sImageName=name of this field's icon picture
//	nLockStatus, locked/frozen status, for display locked/frozen icon
//	nDiscrepancyStatus, discrepancy status, for displaying discrepancy icon
//	nSDVStatus, SDV status, for displaying SDV borderstyle
//	bNote, note status, for displaying note/comment icon
//	bComment, comment status, for displaying note/comment icon
//	sImageSName, name of the sdv image
//	oSelectNoteImage, name of the 'note/comment' image icon for select lists only 
//	nChanges, number of db changes fo field, for 'number of changes' icon
//	sImageCName, name of 'number of changes' icon for field
//  bHasResponse
//
function fnCI(sFieldID,nRepeatNo,vValue,bEnabled,nStatus,
				oHandle,oCapHandle,sImageName,nLockStatus,
				nDiscrepancyStatus,nSDVStatus,bNote,bComment,oRadio,
				sImageSName,sSelectNoteImage,nChanges,sImageCName,
				sComments,sRFC,sUserFull,sNRCTC,bReVal,sRFO,sValidationMessage)
{
	if(oForm.olQuestion==null)
	{
		oForm.olQuestion=new Array();
	}

	if(oForm.olQuestion[sFieldID]==null)
	{
		oForm.olQuestion[sFieldID]=new Object();
	}
	
	// Create Repeat Array if need be
	if(oForm.olQuestion[sFieldID].olRepeat==null)
	{
		oForm.olQuestion[sFieldID].olRepeat=new Array();
	}

	var	oField=oForm.olQuestion[sFieldID].olRepeat	// Object Short-cut
	// Get Next Available Repeat Number
	nRepeatCount=oField.length;
	oField[nRepeatCount]=new Object();

	oField=oField[nRepeatCount]
	oField.sID=sFieldID;
	oField.bEnabled=bEnabled; //current enable status - can change due to skips etc
	oField.bChangable=bEnabled;	// can field ever be changed - is it locked/frozen/hidden/no user change permission
	oField.nStatus=nStatus;
	oField.oHandle=oHandle;
	oField.oCapHandle=oCapHandle;
	oField.sImageName=sImageName;
	oField.nLockStatus=nLockStatus;
	oField.nDiscrepancyStatus=nDiscrepancyStatus;
	oField.bComment=bComment;
	oField.bNote=bNote;
	oField.nSDVStatus=nSDVStatus;
	oField.oRadio=oRadio;
	oField.sImageSName=sImageSName;
	oField.sSelectNoteImage=sSelectNoteImage;
	oField.nChanges=nChanges;
	oField.sImageCName=sImageCName;
	oField.sComments=sComments;
	oField.sRFC=sRFC
	oField.sUserFull=sUserFull;
	oField.sNRCTC=sNRCTC;
	oField.bReVal=bReVal;
	oField.sRFO=sRFO;
	oField.sValidationMessage=sValidationMessage;
	
	bFieldInstancing=true;
	
// Get a reference to Question detail
	var oFieldTemplate = oForm.olQuestion[sFieldID]

//
// Assign the 'get()','getFormatted()' and 'set()' functions for each data type.
//  get() returns the value of the field,unformatte-d and in its native data type;
//  getFormatted() returns a string formatted appropriately for the display type;
//  set() removes formatting and stores in an internal format:
//		Returns false if the value is invalid.
//
	switch(oFieldTemplate.nType)
	{
		case etIntegerNumber:
			oField.get=function()
			{
				return 1*this.vValue;
				//return 0+this.vValue;
			};
			oField.getErrMes=function()
			{
				return"integer number";
			};
			oField.getFormatted=function()
			{
				if((this.vValue==="")||(this.vValue==null)||(this.vValue==undefined))
				{
					return("");
				}
				var	vValue=Math.floor(this.vValue);
				var oFieldTemplate=oForm.olQuestion[this.sID]
				var	sFormat=oFieldTemplate.sFormat;
				var	nValSign=1;

				if(vValue<0)
				{
					nValSign=-1;
					vValue=Math.abs(vValue);
				}
				var	sFormatted=""+vValue;
				if(sFormat.search(/\-/)>=0)
				{
					sFormat=sFormat.replace(/\-/g,"");
				}

				if(sFormatted.length>(""+vValue).nLength)
				{
					//Integer too long for field
					sFormatted=sFormatted.substr(sFormatted.length-nLength,nLength)
				}
				var	nOffset;
				var	sCharType;
				while(sFormatted.length<sFormat.length)
				{
					sFormatted=" "+sFormatted;
				}
				for(nOffset=0; nOffset<sFormatted.length;++nOffset)
				{
					if((sFormatted.substr(nOffset,1)==" ")&&(sFormat.substr(nOffset,1)!="#"))
					{
						sFormatted=sFormatted.substr(0,nOffset)
									+"0"
									+sFormatted.substr(nOffset+1,sFormatted.length-1);
					}
				}
				while(sFormatted.substr(0,1)==" ")
				{
					sFormatted=sFormatted.substr(1,sFormatted.length-1);
				}
				if(nValSign<0)
				{
					sFormatted="-"+sFormatted;
				}

				return isNaN(sFormatted)? "" : sFormatted;
			};
			oField.set=function(vNewValue)
			{			
				//ic 28/05/2002
				if (!fnOnlyLegalChars(vNewValue))
				{
					return false;
				}
				if(vNewValue==undefined)
				{
					vNewValue=0;
				}
				//ic 25/02/2004 use identity
				else if((typeof(vNewValue)!="number")&&(vNewValue===""))
				{
					this.vValue=vNewValue;
					return true;
				}
				//dph 03/02/2004 - inline with windows - round to integer
				// removed non-integer check
				var	vNewValue=jsround(vNewValue,0);

				var oFieldTemplate=oForm.olQuestion[this.sID]
				var	sMax=oFieldTemplate.sFormat.replace(/[0-9#]/g,"9");
				sMax=sMax.replace(/\s*/,"");

				var	nMax=1*sMax;
				var	nMin=0;
				if(nMax<0)
				{
					nMin=nMax;	// Negative mask - sort out the min & max
					nMax=Math.abs(nMax);
				}
				if((vNewValue<=nMax)&&(vNewValue>=nMin))
				{
					this.vValue=vNewValue;
					return true;
				}
				return false;
			};
			oField.setRaw=function(vNewValue)
			{
				this.set(vNewValue);
			}
			oField.blankAll=function()
			{
				this.vValue="";
				this.vOldValue="";
				this.vStartValue="";
				this.vStartStatus=(oForm.olQuestion[this.sID].bOptional)? eStatus.Success:eStatus.Missing;
			}
			oField.blank=function()
			{
				this.vValue="";
			}
			break;
		case etRealNumber:
		case etLabTest:
			oField.get=function()
			{
				//return this.vValue;
				//ic 25/02/2004 use identity
				if(this.vValue==="")
				{
					//this.vValue=0;
					return this.vValue;
				}
				return 1 * this.vValue;
			};
			oField.getErrMes=function()
			{
				return"real number";
			};
			oField.getFormatted=function()
			{
				if((this.vValue==="")||(this.vValue==null)||(this.vValue==undefined))
				{
					return("");
				}
				var oFieldTemplate=oForm.olQuestion[this.sID]
				var	sFormat=oFieldTemplate.sFormat;
				var	sBits;
				var	sVInt;
				var	sVDec;
				var	sVDP;
				var	sFInt;
				var	sFDec;
				var	sFDP;
				var	sFormatted;
				var	rExp=/^([^.]*)(\.)(.*$)/;
				var	nValSign=1;
				var	vValue=this.vValue;
				if(vValue<0)
				{
					nValSign=-1;
					vValue=Math.abs(vValue);
				}
				if(sFormat.search(/\-/)>=0)
				{
					sFormat=sFormat.replace(/\-/g,"");
				}

				var	sBits=sFormat.match(rExp);	// Work out the format details for this field
				if(sBits!=null)
				{
					sFInt=sBits[1];
					sFDP=sBits[2];
					sFDec=sBits[3];
				}
				else
				{
					sFInt=sFormat;
					sFDP="";
					sFDec="";
				}
				sFormatted=fnRound(vValue,sFDec.length);
				sFormatted=sFormatted+"";
				sBits=sFormatted.match(rExp);
				if(sBits!=null)
				{
					sVInt=sBits[1];
					sVDP=sBits[2];
					sVDec=sBits[3];
				}
				else
				{
					sVInt=sFormatted;
					sVDP="";
					sVDec="";
				}
				var	nOffset;
				var	sCharType;
				// Format the whole-number part
				while(sVInt.length<sFInt.length)
				{
					sVInt=" "+sVInt;
				}
				while(sVDec.length<sFDec.length)
				{
					sVDec+=" ";
				}
				for(nOffset=0; nOffset<sFInt.length;++nOffset)
				{
					if((sVInt.substr(nOffset,1)==" ")&&(sFInt.substr(nOffset,1)!="#"))
					{
						sVInt=sVInt.substr(0,nOffset)
									+"0"
									+sVInt.substr(nOffset+1,sVInt.length-1);
					}
				}
				while(sVInt.substr(0,1)==" ")
				{
					sVInt=sVInt.substr(1,sVInt.length-1);
				}
				if(sFDec!="")
				{
					// Format the fractional part
					for(nOffset=0; nOffset<sFDec.length;++nOffset)
					{
						if((sVDec.substr(nOffset,1)==" ")&&(sFDec.substr(nOffset,1)!="#"))
						{
							sVDec=sVDec.substr(0,nOffset)
										+"0"
										+sVDec.substr(nOffset+1,sVDec.length-1);
						}
					}
					while(sVDec.substr(sVDec.length-1,1)==" ")
					{
						sVDec=sVDec.substr(0,sVDec.length-1);
					}
				}
				// Join them up
				//ic 25/02/2004 use identity
				sFormatted=""+sVInt+(sVDec==="" ? "" : sFDP)+sVDec;
				if(nValSign<0)
				{
					sFormatted="-"+sFormatted;
				}
				//ic 25/02/2004 use identity
				if(sFormatted==="")
				{
					sFormatted="0";
				}
				return sFormatted;
			};
			oField.set=function(vNewValue)
			{
				//ic 28/05/2002
				if (!fnOnlyLegalChars(vNewValue))
				{
					return false;
				}
				if(vNewValue==undefined)
				{
					vNewValue=0;
				}
				//ic 25/02/2004 use identity
				if((typeof(vNewValue)!="number")&&(vNewValue===""))
				{
					this.vValue=vNewValue;
					return true;
				}

				var oFieldTemplate=oForm.olQuestion[this.sID]
				var	sMax=oFieldTemplate.sFormat.replace(/[0-9#]/g,"9");
				sMax=sMax.replace(/\s*/,"");
				var	nMax=1*sMax;
				var	nMin=0;
				if(nMax<0)
				{
					nMin=nMax;	// Negative mask - sort out the min & max
					nMax=Math.abs(nMax);
				}
				if((vNewValue<=nMax)&&(vNewValue>=nMin))
				{
					this.vValue=vNewValue;
					// ic/dph 04/02/04 - store formatted number for calculations
					// as raw number may have many numbers after decimal point
					this.vValue=(this.getFormatted())*1;
					return true;
				}
				return false;
			};
			oField.setRaw=function(vNewValue)
			{
				this.set(vNewValue);
			}
			oField.blankAll=function()
			{
				this.vValue="";
				this.vOldValue="";
				this.vStartValue="";
				this.vStartStatus=(oForm.olQuestion[this.sID].bOptional)? eStatus.Success:eStatus.Missing;
			}
			oField.blank=function()
			{
				this.vValue="";
			}
			break;
		case etDateTime:
			oField.get=function()
			{
				if(this.vValue==null)
				{
					return(0);
				}
				else
				{
					return(this.vValue.valueOf());
				}
			};
			oField.getErrMes=function()
			{
				return"date/time";
			};
			oField.getFormatted=function()
			{
				if((isNaN(this.vValue))
					||(this.vValue==null)
					||(this.vValue==undefined)
					||(this.vValue==""))
				{
					if((isNaN(this.vValue))
						||(this.vValue==null)
						||(this.vValue==undefined)
						||(this.vValue==""))
					{
						return("");
					}
				}
				var	iOffset;
				var oFieldTemplate=oForm.olQuestion[this.sID]
				var sFormat=fnGetLocalFormatDate(oFieldTemplate.sFormat);
				sFormat=sFormat.toUpperCase();
				var	sDate;
				var	dDate=this.vValue;

				if(typeof(dDate)!="object")
				{
					//ic 21/02/2003 added 'new'
					dDate= new Date(this.vValue);
				}
				iOffset=fnStrPos(sFormat,"DD",0);
				if(iOffset>=0)
				{
					sDate=dDate.getDate()+100+"X";
					sDate=sDate.substr(1,2);
					sFormat=sFormat.substr(0,iOffset)+sDate+sFormat.substring(iOffset+2,sFormat.length);
				}
				iOffset=fnStrPos(sFormat,"[^H].MM",2);
				if(iOffset>=0)
				{
					sDate=dDate.getMonth()+101+"X";
					sDate=sDate.substr(1,2);
					sFormat=sFormat.substr(0,iOffset)+sDate+sFormat.substring(iOffset+2,sFormat.length);
				}
				else
				{
					iOffset=fnStrPos(sFormat,"^MM",0);
					if(iOffset>=0)
					{
						sDate=dDate.getMonth()+101+"X";
						sDate=sDate.substr(1,2);
						sFormat=sFormat.substr(0,iOffset)+sDate+sFormat.substring(iOffset+2,sFormat.length);
					}
				}
				iOffset=fnStrPos(sFormat,"YYYY",0);
				if(iOffset>=0)
				{
					sDate=dDate.getYear();
					if(sDate<100)
					{
						sDate+=1900;
					}
					sDate=sDate+10000+"X";
					sDate=sDate.substr(1,4);
					sFormat=sFormat.substr(0,iOffset)+sDate+sFormat.substring(iOffset+4,sFormat.length);
				}
				iOffset=fnStrPos(sFormat,"YY",0);
				if(iOffset>=0)
				{
					sDate=dDate.getYear()+10000+"X";
					sDate=sDate.substr(3,2);
					sFormat=sFormat.substr(0,iOffset)+sDate+sFormat.substring(iOffset+2,sFormat.length);
				}
				iOffset=fnStrPos(sFormat,"H.MM",2);
				if(iOffset>=0)
				{
					sDate=dDate.getMinutes()+100+"X";
					sDate=sDate.substr(1,2);
					sFormat=sFormat.substr(0,iOffset)+sDate+sFormat.substring(iOffset+2,sFormat.length);
				}
				iOffset=fnStrPos(sFormat,"HH",0);
				if(iOffset>=0)
				{
					sDate=dDate.getHours()+100+"X";
					sDate=sDate.substr(1,2);
					sFormat=sFormat.substr(0,iOffset)+sDate+sFormat.substring(iOffset+2,sFormat.length);
				}
				iOffset=fnStrPos(sFormat,"SS",0);
				if(iOffset>=0)
				{
					sDate=dDate.getSeconds()+100+"X";
					sDate=sDate.substr(1,2);
					sFormat=sFormat.substr(0,iOffset)+sDate+sFormat.substring(iOffset+2,sFormat.length);
				}
				return sFormat;
			};
			oField.set=function(vValue)
			{
				//ic 28/05/2002
				if (!fnOnlyLegalChars(vValue))
				{
					return false;
				}
				if((vValue==undefined)
					||(vValue==null)
					||(vValue==""))
				{
					this.vValue="";
					return(true);
				}
				else if(vValue.toUpperCase()=="T")
				{
					this.vValue=new Date();
					// Need to zero down all elements of the date which are NOT part of the mask.
					var oFieldTemplate=oForm.olQuestion[this.sID]
					var	sFormat=fnGetLocalFormatDate(oFieldTemplate.sFormat);
					sFormat=sFormat.toUpperCase();
					if(fnStrValue(vValue,sFormat,"YY",0,2,-1)<0)
					{
						//dph 10/03/2004 - if no year format must be a time
						// so default year to standard 1600
						this.vValue.setFullYear(1600);
					}
					if((fnStrValue(vValue,sFormat,"[^H].MM",2,2,-1)<0)
						&&(fnStrValue(vValue,sFormat,"^MM",0,2,-1)<0))
					{
						this.vValue.setMonth(0);
					}
					if(fnStrValue(vValue,sFormat,"DD",0,2,-1)<0)
					{
						this.vValue.setDate(1);
					}
					if(fnStrValue(vValue,sFormat,"HH",0,2,-1)<0)
					{
						this.vValue.setHours(0);
					}
					if(fnStrValue(vValue,sFormat,"H.MM",2,2,-1)<0)
					{
						this.vValue.setMinutes(0);
					}
					if(fnStrValue(vValue,sFormat,"SS",0,2,-1)<0)
					{
						this.vValue.setSeconds(0);
					}
					this.vValue.setMilliseconds(0);	// Always use zero MS
					return(true);
				}
				else
				{
					var oFieldTemplate=oForm.olQuestion[this.sID];
					var	sFormat;
					//use study def format for first set of instance
					if(bFieldInstancing)
					{
						sFormat=oFieldTemplate.sFormat;
					}
					else
					{
						sFormat=fnGetLocalFormatDate(oFieldTemplate.sFormat);
					}
					sFormat=sFormat.toUpperCase();
					var	oDate=fnParseDate(vValue,sFormat);
					if(oDate==null)
					{
						bEvaluationError=true;
						this.vValue=this.vOldValue;
						return false;
					}
					else
					{
						this.vValue=oDate;
						return true;
					}
				}
			};
			oField.setRaw=function(vNewValue)
			{
				if(vNewValue=="")
				{
					this.vValue="";
				}
				else
				{
					var	oDate=new Date(vNewValue);
					this.vValue=oDate;
				}
			}
			oField.blankAll=function()
			{
				this.vValue=null;
				this.vOldValue=null;
				this.vStartValue=null;
				this.vStartStatus=(oForm.olQuestion[this.sID].bOptional)? eStatus.Success:eStatus.Missing;
			}
				oField.blank=function()
		{
				this.vValue=null;
			}
			break;

		case etText:
		case etMultimedia:
		case etCategory:
		case etCatSelect:
			oField.get=function()
			{
				if((this.vValue==null)||(this.vValue==undefined))
				{
					this.vValue="";
				}
				return this.vValue;
			};
			oField.getErrMes=function()
			{
				return "value";
			};
			oField.getFormatted=function()
			{
				if (this.vValue==null)
				{
					return "";
				}
				else
				{
					// use nCase from field template object
					var oFieldTemplate=oForm.olQuestion[this.sID];
					var	nCase=oFieldTemplate.nCase;
					switch (nCase)
					{
						case eUpperCase:
							var sVal = this.vValue+"";
							return sVal.toUpperCase();
							break;
						case eLowerCase:
							var sVal = this.vValue+"";
							return sVal.toLowerCase();
							break;
						default:
							return this.vValue;
					}
				}
			};
			oField.set=function(vValue)
			{
				//ic 28/05/2002
				if (!fnOnlyLegalChars(vValue))
				{
					return false;
				}
				//ic 25/02/2004 use identity
				if ((vValue==undefined)||(vValue===""))
				{
					this.vValue="";
				}
				else
				{
					//handle any text format mask
					//ic 25/02/2004 force to string type
					var sF=oFieldTemplate.sFormat+"";
					var sCompareValue=vValue+"";
					var sP=/[^Aa9]/
					
					if ((sF!=null)&&(sF!=undefined)&&(sF!="")) //is there a format mask
					{
						if (sP.exec(sF)==null) //is format mask correct syntax
						{
							if (sF.length==sCompareValue.length) //is new value same length as format mask
							{
								//build regular expression for mask
								sF=sF.replace(/[Aa]/g,"[A-Za-z]");
								sF=sF.replace(/[9]/g,"[0-9]");									
								var sExp=new RegExp(sF);									
								if (sExp.exec(sCompareValue)==null) //does new value format match mask
								{										
									return false;
								}
							}
							else
							{
								return false;
							}
						}
					}
					this.vValue=vValue;
				}				
				return(true);
			};
			oField.setRaw=function(vNewValue)
			{
				this.set(vNewValue);
			}
			oField.blankAll=function()
			{
				this.vValue=null;
				this.vOldValue=null;
				this.vStartValue=null;
				this.vStartStatus=(oForm.olQuestion[this.sID].bOptional)? eStatus.Success:eStatus.Missing;
			}
			oField.blank=function()
			{
				this.vValue=null;
			}
			break;
		default:
	}
	oField.enterable=function()
	{
		//is field currently enterable
		//no if it has any derivations
		if ((oForm.olQuestion[this.sID].olDerivation!=undefined)&&(oForm.olQuestion[this.sID].olDerivation.length>0)) return false;
		//no if skipped, NEVER changeable (see assignment for description), locked or frozen
		return ((this.nStatus!=eStatus.NotApplicable)&&(this.bChangable)&&(!fnLockedOrFrozen("","",this)))
	}

	oField.set(vValue);
	oField.vOldValue=oField.vValue;
	var	sForm=oField.getFormatted();
	if(!oFieldTemplate.bRQG)
	{
		setJSValue(sFieldID,sForm,false,nRepeatNo);
		if(bEnabled)
		{
			setFieldEnabled(sFieldID,nRepeatNo,true);
		}
	}
	oField.vDBFormatted=sForm;
	oField.vDBValue=vValue;
	oField.nDBStatus=nStatus;

	bFieldInstancing=false;

	if(!oFieldTemplate.bRQG)
	{
		if ((oFieldTemplate.bEform&&!oForm.bUReadOnly)||(!oFieldTemplate.bEform&&!oForm.bVReadOnly)&&oEP.bChangeData)
		{
			//only adjust start value if not on readonly eform
			oField.vStartValue=fnGetStartValue(sFieldID,nRepeatNo);
			oField.vStartStatus=getFieldStatus(sFieldID,nRepeatNo);
		}
	}
	return true;
}

//
// Parse the user-entered input for correctness according to the supplied format.
//  Returns a date, if the input was acceptable, or null if it was not.
//  All missing elements are based on a default date of midnight on 1st January 1600.
//
//  Dates:
//
//      YMD
//      DMY
//      MDY
//      YM
//      MY
//
//  Times:
//
//      HMS
//      HM
//
//  Combinations:
//      YMD HMS
//      DMY HMS
//      MDY HMS
//      YMD HM
//      DMY HM
//      MDY HM
//
function fnParseDate(sInput,sFormat)
{
	// Get the basic format from the mask. see above list for allowable values.
	sFormat=sFormat.replace(/[Dd]+/g,"D");
	sFormat=sFormat.replace(/[Mm]+/g,"M");
	sFormat=sFormat.replace(/[Yy]+/g,"Y");
	sFormat=sFormat.replace(/[Hh]+/g,"H");
	sFormat=sFormat.replace(/[Ss]+/g,"S");
	// Get the format into a standard string, based on any of the following: "/.:-" and whitespace.
	sFormat=sFormat.replace(/[/\-:.\s]/g,"");
	var	lInput=sInput.split(/[/\-:.\s]/);
	// Set defaults for the offsets of the values (-1=missing)
	var	nElts=lInput.length;
	var	nDayO=-1;
	var	nMonthO=-1;
	var	nYearO=-1;
	var	nHourO=-1;
	var	nMinuteO=-1;
	var	nSecondO=-1;
	var	bError=false;
	switch (sFormat)
	{
		case "YMD":
			if(nElts==1)
			{
				// No separators
				if(lInput[0].length==6)
				{
					lInput[2]=lInput[0].substr(4);	// Day
					lInput[1]=lInput[0].substr(2,2);	// Month
					lInput[0]=lInput[0].substr(0,2);	// Year
				}
				else if(lInput[0].length==8)
				{
					lInput[2]=lInput[0].substr(6);	// Day
					lInput[1]=lInput[0].substr(4,2);	// Month
					lInput[0]=lInput[0].substr(0,4);	// Year
				}
				else
				{
					bError=true;	// Duff input
				}
			}
			else if(nElts!=3)
			{
				bError=true;	// Duff input
			}
			nDayO=2;
			nMonthO=1;
			nYearO=0;
			break;
		case "DMY":
			if(nElts==1)
			{
				// No separators
				if((lInput[0].length==6)||(lInput[0].length==8))
				{
					lInput[2]=lInput[0].substr(4);	// Year
					lInput[1]=lInput[0].substr(2,2);	// Month
					lInput[0]=lInput[0].substr(0,2);	// Day
				}
				else
				{
					bError=true;	// Duff input
				}
			}
			else if(nElts!=3)
			{
				bError=true;	// Duff input
			}
			nDayO=0;
			nMonthO=1;
			nYearO=2;
			break;
		case "MDY":
			if(nElts==1)
			{
				// No separators
				if((lInput[0].length==6)||(lInput[0].length==8))
				{
					lInput[2]=lInput[0].substr(4);	// Year
					lInput[1]=lInput[0].substr(2,2);	// Day
					lInput[0]=lInput[0].substr(0,2);	// Month
				}
				else
				{
					bError=true;	// Duff input
				}
			}
			else if(nElts!=3)
			{
				bError=true;	// Duff input
			}
			nDayO=1;
			nMonthO=0;
			nYearO=2;
			break;
		case "YM":
			if(nElts==1)
			{
				// No separators
				if(lInput[0].length==4)
				{
					lInput[1]=lInput[0].substr(2);	// Month
					lInput[0]=lInput[0].substr(0,2);	// Year
				}
				else if(lInput[0].length==6)
				{
					lInput[1]=lInput[0].substr(4);	// Month
					lInput[0]=lInput[0].substr(0,4);	// Year
				}
				else
				{
					bError=true;	// Duff input
				}
			}
			else if(nElts!=2)
			{
				bError=true;	// Duff input
			}
			nMonthO=1;
			nYearO=0;
			break;
		case "MY":
			if(nElts==1)
			{
				// No separators
				if((lInput[0].length==4)||(lInput[0].length==6))
				{
					lInput[1]=lInput[0].substr(2);	// Year
					lInput[0]=lInput[0].substr(0,2);	// Month
				}
				else
				{
					bError=true;	// Duff input
				}
			}
			else if(nElts!=2)
			{
				bError=true;	// Duff input
			}
			nMonthO=0;
			nYearO=1;
			break;
		case "HMS":
			if(nElts==1)
			{
				// No separators
				if(lInput[0].length==6)
				{
					lInput[2]=lInput[0].substr(4);	// Seconds
					lInput[1]=lInput[0].substr(2,2);	// Minutes
					lInput[0]=lInput[0].substr(0,2);	// Hours
				}
				else
				{
					bError=true;	// Duff input
				}
			}
			else if(nElts!=3)
			{
				bError=true;	// Duff input
			}
			nHourO=0;
			nMinuteO=1;
			nSecondO=2;
			break;
		case "HM":
			if(nElts==1)
			{
				// No separators
				if(lInput[0].length==4)
				{
					lInput[1]=lInput[0].substr(2,2);	// Minutes
					lInput[0]=lInput[0].substr(0,2);	// Hours
				}
				else
				{
					bError=true;	// Duff input
				}
			}
			else if(nElts!=2)
			{
				bError=true;	// Duff input
			}
			nHourO=0;
			nMinuteO=1;
			break;
		case "YMDHMS":
			if(nElts==2)
			{
				// A single separator - assume "YMD HMS" entered
				if (lInput[1].length==6)
				{
					lInput[5]=lInput[1].substr(4);	// Seconds
					lInput[4]=lInput[1].substr(2,2);	// Minutes
					lInput[3]=lInput[1].substr(0,2);	// Hours
				}
				else
				{
					bError=true;
				}
				if(lInput[0].length==6)
				{
					lInput[2]=lInput[0].substr(4);	// Day
					lInput[1]=lInput[0].substr(2,2);	// Month
					lInput[0]=lInput[0].substr(0,2);	// Year
				}
				else if(lInput[0].length==8)
				{
					lInput[2]=lInput[0].substr(6);	// Day
					lInput[1]=lInput[0].substr(4,2);	// Month
					lInput[0]=lInput[0].substr(0,4);	// Year
				}
				else
				{
					bError=true;	// Duff input
				}
			}
			else if(nElts!=6)
			{
				bError=true;	// Duff input
			}
			nSecondO=5;
			nMinuteO=4;
			nHourO=3
			nDayO=2;
			nMonthO=1;
			nYearO=0;
			break;
		case "DMYHMS":
			if(nElts==2)
			{
				// A single separator - assume "DMY HMS" entered
				if (lInput[1].length==6)
				{
					lInput[5]=lInput[1].substr(4);	// Seconds
					lInput[4]=lInput[1].substr(2,2);	// Minutes
					lInput[3]=lInput[1].substr(0,2);	// Hours
				}
				else
				{
					bError=true;
				}
				if((lInput[0].length==6)||(lInput[0].length==8))
				{
					lInput[2]=lInput[0].substr(4);	// Year
					lInput[1]=lInput[0].substr(2,2);	// Month
					lInput[0]=lInput[0].substr(0,2);	// Day
				}
				else
				{
					bError=true;	// Duff input
				}
			}
			else if(nElts!=6)
			{
				bError=true;	// Duff input
			}
			nSecondO=5;
			nMinuteO=4;
			nHourO=3
			nYearO=2;
			nMonthO=1;
			nDayO=0;
			break;
		case "MDYHMS":
			if(nElts==2)
			{
				// A single separator - assume "MDY HMS" entered
				if (lInput[1].length==6)
				{
					lInput[5]=lInput[1].substr(4);	// Seconds
					lInput[4]=lInput[1].substr(2,2);	// Minutes
					lInput[3]=lInput[1].substr(0,2);	// Hours
				}
				else
				{
					bError=true;
				}
				if((lInput[0].length==6)||(lInput[0].length==8))
				{
					lInput[2]=lInput[0].substr(4);	// Year
					lInput[1]=lInput[0].substr(2,2);	// Day
					lInput[0]=lInput[0].substr(0,2);	// Month
				}
				else
				{
					bError=true;	// Duff input
				}
			}
			else if(nElts!=6)
			{
				bError=true;	// Duff input
			}
			nSecondO=5;
			nMinuteO=4;
			nHourO=3
			nYearO=2;
			nDayO=1;
			nMonthO=0;
			break;
		case "YMDHM":
			if(nElts==2)
			{
				// A single separator - assume "YMD HM" entered
				if (lInput[1].length==4)
				{
					lInput[4]=lInput[1].substr(2,2);	// Minutes
					lInput[3]=lInput[1].substr(0,2);	// Hours
				}
				else
				{
					bError=true;
				}
				if(lInput[0].length==6)
				{
					lInput[2]=lInput[0].substr(4);	// Day
					lInput[1]=lInput[0].substr(2,2);	// Month
					lInput[0]=lInput[0].substr(0,2);	// Year
				}
				else if(lInput[0].length==8)
				{
					lInput[2]=lInput[0].substr(6);	// Day
					lInput[1]=lInput[0].substr(4,2);	// Month
					lInput[0]=lInput[0].substr(0,4);	// Year
				}
				else
				{
					bError=true;	// Duff input
				}
			}
			else if(nElts!=5)
			{
				bError=true;	// Duff input
			}
			nMinuteO=4;
			nHourO=3
			nDayO=2;
			nMonthO=1;
			nYearO=0;
			break;
		case "DMYHM":
			if(nElts==2)
			{
				// A single separator - assume "DMY HM" entered
				if (lInput[1].length==4)
				{
					lInput[4]=lInput[1].substr(2,2);	// Minutes
					lInput[3]=lInput[1].substr(0,2);	// Hours
				}
				else
				{
					bError=true;
				}
				if((lInput[0].length==6)||(lInput[0].length==8))
				{
					lInput[2]=lInput[0].substr(4);	// Year
					lInput[1]=lInput[0].substr(2,2);	// Month
					lInput[0]=lInput[0].substr(0,2);	// Day
				}
				else
				{
					bError=true;	// Duff input
				}
			}
			else if(nElts!=5)
			{
				bError=true;	// Duff input
			}
			nMinuteO=4;
			nHourO=3
			nYearO=2;
			nMonthO=1;
			nDayO=0;
			break;
		case "MDYHM":
			if(nElts==2)
			{
				// A single separator - assume "MDY HM" entered
				if (lInput[1].length==4)
				{
					lInput[4]=lInput[1].substr(2,2);	// Minutes
					lInput[3]=lInput[1].substr(0,2);	// Hours
				}
				else
				{
					bError=true;
				}
				if((lInput[0].length==6)||(lInput[0].length==8))
				{
					lInput[2]=lInput[0].substr(4);	// Year
					lInput[1]=lInput[0].substr(2,2);	// Day
					lInput[0]=lInput[0].substr(0,2);	// Month
				}
				else
				{
					bError=true;	// Duff input
				}
			}
			else if(nElts!=5)
			{
				bError=true;	// Duff input
			}
			nMinuteO=4;
			nHourO=3
			nYearO=2;
			nDayO=1;
			nMonthO=0;
			break;
	}
	// Now put it all together and see if we get a valid date...
	if(!bError)
	{
		var	oDate=new Date(0);
		var	nYear=1600;
		var	nMonth=0;
		var	nDay=1;
		var	nHour=0;
		var	nMinute=0;
		var	nSecond=0;
		if(nYearO>=0)
		{
			nYear=1*lInput[nYearO];
			if(nYear<100)
			{
				nYear=nYear+(nYear<50?2000:1900);	// Years less than 50 become  21st Century, rest are 20th
			}
		}
		oDate.setFullYear(nYear);
		if(nMonthO>=0)
		{
			nMonth=1*lInput[nMonthO]-1;
		}
		if(nDayO>=0)
		{
			nDay=1*lInput[nDayO];
		}
		oDate.setDate(nDay);
		// oddity with us installation requires setmonth after setting the day
		oDate.setMonth(nMonth);
		if(nHourO>=0)
		{
			nHour=1*lInput[nHourO];
		}
		oDate.setHours(nHour);
		if(nMinuteO>=0)
		{
			nMinute=1*lInput[nMinuteO];
		}
		oDate.setMinutes(nMinute);
		if(nSecondO>=0)
		{
			nSecond=1*lInput[nSecondO];
		}
		oDate.setSeconds(nSecond);
		// Check it all ties up - we get back what we put in
		if((oDate.getFullYear()!=nYear)
			|| (oDate.getMonth()!=nMonth)
			|| (oDate.getDate()!=nDay)
			|| (oDate.getHours()!=nHour)
			|| (oDate.getMinutes()!=nMinute)
			|| (oDate.getSeconds()!=nSecond))
		{
			bError=true;
		}
	}
	if(bError)
	{
		return null;
	}
	else
	return oDate;
}


//
// Function to assign text values associated with category question values.
//
function fnSetCategoryText(sFieldID,sCatValue,sCatText)
{
	if(oForm.olQuestion==null)
	{
		oForm.olQuestion=new Array();
	}

	if(oForm.olQuestion[sFieldID]==null)
	{
		oForm.olQuestion[sFieldID]=new Object();
	}

	if(oForm.olQuestion[sFieldID].olCatValue==null)
	{
		oForm.olQuestion[sFieldID].olCatValue=new Array();
	}

	if(oForm.olQuestion[sFieldID].olCatValue[sCatValue]==null)
	{
		oForm.olQuestion[sFieldID].olCatValue[sCatValue]=new Object();
	}
	oForm.olQuestion[sFieldID].olCatValue[sCatValue].sCatText=sCatText;
}

//
// Function to add the supplied validation rule to the identified question.
//
function fnSetValidationRule(sFieldID,sType,sExpression,sMessage,sNiceExpression)
{
	var	nValidationCount;

	//reset bEvaluationError in case a previous fnSetValidationRule() call set it to true
	bEvaluationError=false;

	if(sExpression=="\"\"")
	{
		// Not translated into javascript correctly so ignore
		return;
	}

	switch (sType)
	{
		case 0:
			sType=eValidation.Reject;	// Reject data if invalid
			break;
		case 2:
			sType=eValidation.Inform;	// Inform
			break;
		case 1:
		default:
			sType=eValidation.Warn;	// Warn if data invalid - allow over-rule
			break;
	}
	if(oForm.olQuestion==null)
	{
		oForm.olQuestion=new Array();
	}

	if(oForm.olQuestion[sFieldID]==null)
	{
		oForm.olQuestion[sFieldID]=new Object();
	}
	if(oForm.olQuestion[sFieldID].olValidation==null)
	{
		oForm.olQuestion[sFieldID].olValidation=new Array();
	}

	var	oField=oForm.olQuestion[sFieldID].olValidation	// Short-cut
	nValidationCount=oField.length;
	oField[nValidationCount]=new Object();

	oField=oField[nValidationCount]
	oField.nOrder=nValidationCount;
	oField.sExpression=sExpression;
	oField.sType=sType;
	oField.sMessage=sMessage;
	oField.sNiceExpression=sNiceExpression;

	// Add the field to the relevant dependency lists
	fnCreateDependencies(sFieldID,eDepType.Validation,sExpression);
}

//
// Apply all the skip and derivation rules we have. This is done as the final part of page initialisation.
// It is near impossible to know what order to do these in...
//
//ic 20/08/2002
//added calls to fnDisplayNoteStatus() & fnDisplaySDVStatus
function fnApplyRules()
{
	// set all status icons for questions on the eform
	fnSetAllStatuses();
	
	// perform all skips & derivs
	fnExecuteAllSkipsDerivs();
	
	// perform dependency deriv/skip conditions
	for(var	sFieldID in oForm.olQuestion)
	{
		fnEmptyDependentsStack();
		fnCalculateDependencies(oForm.olQuestion[sFieldID].sID,true,undefined,true);	// Calculate all dependencies
	}
	
	// Anything which is derived or is a multimedia type needs to be READ ONLY
	for(var	sFieldID in oForm.olQuestion)
	{
		for(var nRepeat=0;nRepeat<oForm.olQuestion[sFieldID].olRepeat.length;nRepeat++)
		{
			setFieldStatus(oForm.olQuestion[sFieldID].sID,oForm.olQuestion[sFieldID].olRepeat[nRepeat].nStatus,oForm.olQuestion[sFieldID].olRepeat[nRepeat].nLockStatus,oForm.olQuestion[sFieldID].olRepeat[nRepeat].nDiscrepancyStatus,oForm.olQuestion[sFieldID].olRepeat[nRepeat].nSDVStatus,nRepeat);	// Display the appropriate icon for this field
			fnDisplayNoteStatus(sFieldID,nRepeat);
			fnDisplaySDVStatus(sFieldID,nRepeat);
			fnDisplayChangeStatus(sFieldID,nRepeat);
		
			if(oForm.olQuestion[sFieldID].nType==etMultimedia)
			{
				setInputDisabled(sFieldID,nRepeat);
				if(getFieldStatus(sFieldID,nRepeat)==eStatus.Requested)
				{
					setFieldStatus(sFieldID,eStatus.Unobtainable,oForm.olQuestion[sFieldID].olRepeat[nRepeat].nLockStatus,oForm.olQuestion[sFieldID].olRepeat[nRepeat].nDiscrepancyStatus,oForm.olQuestion[sFieldID].olRepeat[nRepeat].nSDVStatus,nRepeat);	// Set to "Unobtainable"
				}
			}
			else if (!oForm.olQuestion[sFieldID].olRepeat[nRepeat].enterable())
			{
				setInputDisabled(sFieldID,nRepeat);
			}
		}
	}
	fnSetDisabledRadiosValue();
	fnInitialiseComments();
	fnSetNRCTCDiv();
	fnDisplayLab();
	fnRevalidatePage(true);
}
//ic 21/02/2003
//function adds comments to hidden fields
function fnInitialiseComments()
{
	var oAIField;
	for(var	nFieldID in oForm.olQuestion)
	{
		for(var nRepeat=0;nRepeat<oForm.olQuestion[nFieldID].olRepeat.length;nRepeat++)
		{
			oAIField=fnGetFieldProperty(nFieldID,"oAIHandle",nRepeat);
			oAIField.value=oAIField.value+sDel1+oForm.olQuestion[nFieldID].olRepeat[nRepeat].sComments;
		}
	}
}
//ic 13/02/2003
//function sets value for disabled radios
function fnSetDisabledRadiosValue()
{
	for(var	nFieldID in oForm.olQuestion)
	{
		for(var nRepeat=0;nRepeat<oForm.olQuestion[nFieldID].olRepeat.length;nRepeat++)
		{
			if((oForm.olQuestion[nFieldID].nType==etCategory)&&(!oForm.olQuestion[nFieldID].olRepeat[nRepeat].enterable()))
			{
				setJSValue(nFieldID,oForm.olQuestion[nFieldID].olRepeat[nRepeat].vValue,true,nRepeat,true);
			}
		}
	}
}
//dph 24/02/2003
//function sets CTC grade up if necessary
function fnSetNRCTCDiv()
{
	var sHTML="";
	var oHandle;
	for(var	nFieldID in oForm.olQuestion)
	{
		for(var nRepeat=0;nRepeat<oForm.olQuestion[nFieldID].olRepeat.length;nRepeat++)
		{
			// if not RQG draw CTC
			if(!oForm.olQuestion[nFieldID].bRQG)
			{
				fnDrawNRCTC(nFieldID,nRepeat);
			}
		}
	}
}
//
// function to draw NR/CTC table
function fnDrawNRCTC(sFieldID,nRepeat)
{
	if(oForm.olQuestion[sFieldID].olRepeat[nRepeat].sNRCTC!="")
	{
		sHTML="<table class='clsNRCTC'><tr><td><b>"+oForm.olQuestion[sFieldID].olRepeat[nRepeat].sNRCTC+"</b></td></tr></table>";
		if(oForm.olQuestion[sFieldID].olRepeat.length>1)
		{
			// array
			if((document.all[oForm.olQuestion[sFieldID].sID+"_tdCTC"][nRepeat])!=undefined)
			{
				(document.all[oForm.olQuestion[sFieldID].sID+"_tdCTC"][nRepeat]).innerHTML=sHTML;
			}
		}
		else
		{
			// normal
			if((document.all[oForm.olQuestion[sFieldID].sID+"_tdCTC"])!=undefined)
			{
				document.all[oForm.olQuestion[sFieldID].sID+"_tdCTC"].innerHTML=sHTML;
			}
		}
	}
}
//
// Function to add the supplied derivation rule to the identified question.
//
function fnSetDerivationRule(sFieldID,sExpression)
{
	var	nDerivationCount;
	var vValue;

	//reset bEvaluationError in case a previous fnSetDerivationRule() call set it to true
	bEvaluationError=false;

	if(oForm.olQuestion==null)
	{
		return false;
		oForm.olQuestion=new Array();
	}

	if(oForm.olQuestion[sFieldID]==null)
	{
		return false;
		oForm.olQuestion[sFieldID]=new Object();
	}
	if(oForm.olQuestion[sFieldID].olDerivation==null)
	{
		oForm.olQuestion[sFieldID].olDerivation=new Array();
	}

	nDerivationCount=oForm.olQuestion[sFieldID].olDerivation.length;
	oForm.olQuestion[sFieldID].olDerivation[nDerivationCount]=new Object();
	oForm.olQuestion[sFieldID].olDerivation[nDerivationCount].nOrder=nDerivationCount;
	oForm.olQuestion[sFieldID].olDerivation[nDerivationCount].sExpression=sExpression;

	// Add the field to the relevant dependency lists
	fnCreateDependencies(sFieldID,eDepType.Derivation,sExpression);
	for(var nRepeat=0;nRepeat<oForm.olQuestion[sFieldID].olRepeat.length;nRepeat++)
	{
		setInputDisabled(sFieldID,nRepeat);
	}
}

//
// Function to add the supplied skip rule to the identified question.
//
function fnSetSkipRule(sFieldID,sExpression,bRQGSkip)
{
	var	nSkipCount;
	
	//reset bEvaluationError in case a previous fnSetSkipRule() call set it to true
	bEvaluationError=false;
	
	if(sExpression=="\"\"")
	{
		// Not translated into javascript correctly so ignore
		return;
	}
	
	if(!bRQGSkip)
	{
		if(oForm.olQuestion==null)
		{
			return false;
			oForm.olQuestion=new Array();
		}

		if(oForm.olQuestion[sFieldID]==null)
		{
			return false;
			oForm.olQuestion[sFieldID]=new Object();
		}
		if(oForm.olQuestion[sFieldID].olSkip==null)
		{
			oForm.olQuestion[sFieldID].olSkip=new Array();
		}
		nSkipCount=oForm.olQuestion[sFieldID].olSkip.length;
		oForm.olQuestion[sFieldID].olSkip[nSkipCount]=new Object();
		oForm.olQuestion[sFieldID].olSkip[nSkipCount].nOrder=nSkipCount;
		oForm.olQuestion[sFieldID].olSkip[nSkipCount].sExpression=sExpression;

		// Add the field to the relevant dependency lists
		fnCreateDependencies(sFieldID,eDepType.Skip,sExpression);
	}
	else
	{
		if(aRQG[sFieldID]==null)
		{
			return false;
		}
		aRQG[sFieldID].setallskips(sExpression);		
	}

}

//
// Function to return the value of the supplied question code
//
function jsValueOf(sFieldID,nRepeat)
{
	// Default Repeat No for non RQG questions
	nRepeat=DefaultRepeatNo(nRepeat,true);
	
	if(!IsFieldOnForm(sFieldID,nRepeat))
	{
		bEvaluationError=true;
		return "";
	}
	if(!isNaN(nRepeat))
	{
		switch(oForm.olQuestion[sFieldID].nType)
		{
			case	etText:
			case	etCategory:
			case	etMultimedia:
			case	etLabTest:
			case	etCatSelect:
				return ""+oForm.olQuestion[sFieldID].olRepeat[nRepeat].get();	// Return as a string
				break;
			case	etIntegerNumber:
			case	etRealNumber:
				var	vValue=oForm.olQuestion[sFieldID].olRepeat[nRepeat].vValue;
				//ic 07/01/2004 compare vValue with "" using identity(===) operator as (0=="")
				//incorrectly evaluates to true
				if((vValue==="")||(vValue==null)||(vValue==undefined))
				{
					vValue=0;
					bEvaluationError=true;
				}
				//ic 09/01/2004 return vValue
				return 1*vValue;	// Return as a numeric
				break;
			case	etDateTime:
				var	vValue=oForm.olQuestion[sFieldID].olRepeat[nRepeat].get();
				if((vValue=="")||(vValue==null)||(vValue==undefined))
				{
					vValue=0;
					bEvaluationError=true;
				}
				return vValue;	// Return as a numeric
				break
		}
	}
	else
	{
		bEvaluationError=true;
		switch(oForm.olQuestion[sFieldID].nType)
		{
			case	etText:
			case	etCategory:
			case	etMultimedia:
			case	etLabTest:
			case	etCatSelect:
				return "";
				break;
			case	etIntegerNumber:
			case	etRealNumber:
				return 0;
				break;
			case	etDateTime:
				return 0; 
				break;
		}
	}
}

//
// Function to add the specifiec field to the relevant dependency list.
// Used only during initialisation.
// 
function fnAddDependency(sFieldID,sDependentField,sType)
{
	//dont make a field dependent on itself
	if(sFieldID==sDependentField)
	{
		return;
	}
	//if no questions exist, dont add dependency
	if(oForm.olQuestion==null)
	{
		return;
	}
	//if this question doesnt exist, dont add dependency
	if(oForm.olQuestion[sFieldID]==null)
	{
		return;
	}
	//create a new dependency list, if necessary
	if(oForm.olQuestion[sFieldID].olDependency==null)
	{
		oForm.olQuestion[sFieldID].olDependency=new Array();
	}
	nDependencyCount=oForm.olQuestion[sFieldID].olDependency.length;
	
	//check whether dependency is already added
	var	iIndexRef;
	for(nIndexRef=0; nIndexRef<nDependencyCount;++nIndexRef)
	{
		if((oForm.olQuestion[sFieldID].olDependency[nIndexRef].sDependencyField==sDependentField)
			&&(oForm.olQuestion[sFieldID].olDependency[nIndexRef].sType==sType))
		{
			//already added, only needs to be in list once
			return;
		}
	}
	//add dependency
	oForm.olQuestion[sFieldID].olDependency[nDependencyCount]=new Object();
	oForm.olQuestion[sFieldID].olDependency[nDependencyCount].nOrder=nDependencyCount;
	oForm.olQuestion[sFieldID].olDependency[nDependencyCount].sType=sType;
	oForm.olQuestion[sFieldID].olDependency[nDependencyCount].sDependencyField=sDependentField;
}

//
// Function to add the identified field to the dependency list for all fields referred to
// in the supplied expression,using the supplied dependency type - V(alidation), D(erivation), S(kip)
// ic 22/04/2005 changed to pattern matching
function fnCreateDependencies(sFieldID,sType,sExpression)
{
	//get an array of field codes found in this expression
	//eg in the expression 'jseq( jsValueOf( "f_10061_10095", jsRepNo( "f_10061_10095", "1" ) ), 1 )'
	//the field code 'f_10061_10095' will be matched and returned twice in an array
	var sPattern = /f_\d{5}_\d{5}/g
	var aDependentFields=sExpression.match(sPattern);
	
	//if any are found
	if(aDependentFields!=null)
	{
		//loop through adding them to the dependency list. fnAddDependency function
		//will handle duplicates
		for(n=0;n<aDependentFields.length;n++)
		{
			fnAddDependency(aDependentFields[n],sFieldID,sType);
		}
	}
}

//
// Function to perform all relevant page-level validations and jump to the specified URL
//
function fnFinalisePage(oMainWindow, oSubmitForm, bF7, bF6)
{
	if((!fnAllMandQsComplete())&&(!bF7))
	{
		if((fnUserHasChangedEForm()||fnDerivationHasChangedEForm()||bF6))
		{
			//inform user of missing mandatory questions if they are saving 
			//and (they have changed data) or (a derivation has changed data) 
			//or (it is an explicit save and return)
			//user is not informed of missing mandatory questions during a save
			//if (this is a new eform) and (user has clicked a navigation button 
			//other than 'save and return')
			if(!confirm("Some mandatory questions are blank. Are you sure you wish to leave this eForm?"))
			{
				return false;
			}
		}
	}
	// Disable input for all fields
	var	sFieldID;
	for(sFieldID in oForm.olQuestion)
	{
		for(nRepeat in oForm.olQuestion[sFieldID].olRepeat)
		{	
			setInputDone(sFieldID,nRepeat);
		}
	}

	return true;	// form saved	
}

//function decides if the passed question requires a password authentication,
//gets one if it does
function fnPassword(sFieldID,nRepeat)
{
	var oTField=oForm.olQuestion[sFieldID];
	var oField=oForm.olQuestion[sFieldID].olRepeat[nRepeat];
	var bPassword=false;
	var bOk=true;
	
	if((oTField.sAuthorisation!=null)&&(oTField.sAuthorisation!=""))
	{
		//only require password for enterable fields, not skipped or derived
		if(oField.enterable())
		{
			switch(oTField.nType)
			{
				case etDateTime:
					//compare the formatted value for dates (eg 01/01/2004)
					if(oField.getFormatted()!=oField.vDBValue) bPassword=true;
					break;
				default:
					if(oField.vDBValue==="")
					{
						//use identity operator for comparison with an empty string
						if(!(oField.vValue===oField.vDBValue)) bPassword=true;
					}
					else
					{
						//otherwise use the equality operator
						if(oField.vValue!=oField.vDBValue) bPassword=true;
					}
					break;
			}
		}
		if(bPassword) bOk=fnGetPassword(sFieldID,nRepeat,oTField.sAuthorisation);
	}
	
	return bOk;
}

//function displays the password authentication dialog, updates the password 
//and username in the ai field, returns success status
function fnGetPassword(sFieldID,nRepeat,sAuthorisation)
{
	var sPassword;
	var bOk=true;
	var sCaption=fnGetFieldProperty(sFieldID,"sCaptionText");
	
	//get the authorisation password
	sPassword=window.showModalDialog('passwordInput.asp?name='+sCaption+'&role='+sAuthorisation,'',
	'dialogHeight:300px; dialogWidth:500px; center:yes; status:0; dependent,scrollbars');
	
	if ((sPassword!=undefined)&&(sPassword!=""))
	{	
		fnReplaceAIBlock(sFieldID,"p",sPassword,"","",nRepeat,"")
	}
	else
	{
		//if no valid password was returned, flag to calling function
		bOk = false;
	}
	
	return bOk;
}

//function decides if the passed question requires a rfc, gets one
//if it does. for nType see enumeration at top of file
function fnRFC(sFieldID,nRepeat,nType)
{
	var oTField=oForm.olQuestion[sFieldID];
	var oField=oForm.olQuestion[sFieldID].olRepeat[nRepeat];
	var bRFC=false;
	var bOk=true;
	
	//question must require a rfc, and must have been saved at least once
	if((oTField.bRequiresRFC)&&(oField.nChanges>0))
	{
		//type of change: value,status,overrule reason
		switch(nType)
		{
			case eRFCValue:
				//only require rfc for enterable fields, not skipped or derived
				//as they will get an automatic rfc
				if(oField.enterable())
				{
					switch(oTField.nType)
					{
						case etDateTime:
							//compare the formatted value for dates (eg 01/01/2004)
							if(oField.getFormatted()!=oField.vDBValue) bRFC=true;
							break;
						default:
							if(oField.vDBValue==="")
							{
								//use identity operator for comparison with an empty string
								if(!(oField.vValue===oField.vDBValue)) bRFC=true;
							}
							else
							{
								//otherwise use the equality operator
								if(oField.vValue!=oField.vDBValue) bRFC=true;
							}
							break;
					}
				}
				break;
			case eRFCStatus:
				//compare current field status with db status
				//status rfc is only done when the user manually changes between overrule/warning
				if((getFieldStatus(sFieldID,nRepeat)!=oField.nDBStatus)) bRFC=true;
				break;
			case eRFCOverrule:
				//comparison of overrule reason will already have been done
				bRFC=true;
				break;
			default:
		}
		if(bRFC) bOk=fnGetRFC(sFieldID,nRepeat);
	}
	
	return bOk;
}

//function displays the rfc dialog, updates the rfc in the ai field,
//returns success status
function fnGetRFC(sFieldID,nRepeat)
{
	var sRFC;
	var bOk=true;
	var sDb=fnGetFormProperty("sDatabase");
	var sSt=fnGetFormProperty("sStudyId");
	var sCaption=fnGetFieldProperty(sFieldID,"sCaptionText");
	
	//display dialog to get user entered rfc
	sRFC=window.showModalDialog('../sites/'+sDb+'/'+sSt+'/rfc.html?name='+sCaption,'',
	'dialogHeight:300px; dialogWidth:500px; status:0; center:yes; dependent,scrollbars');
			
	if ((sRFC!=undefined)&&(!(sRFC=="")))
	{
		//if the rfc is ok, update it in the ai field and jsve
		fnReplaceAIBlock(sFieldID,"r","","","",nRepeat,sRFC);
		oForm.olQuestion[sFieldID].olRepeat[nRepeat].sRFC=sRFC;
	}
	else
	{
		bOk=false;
	}
	
	return bOk;
}

// DPH 26/02/2003
// ic 03/03/2004 revalidate on loading and saving
function fnRevalidatePage(bDealWithRequested)
{
	var	sFieldID;
	var nRepeat;
	
	if(fnDataIsReadOnly()) return true;
	
	gbReValidationOk=true;
	
	for(sFieldID in oForm.olQuestion)
	{
		for(nRepeat in oForm.olQuestion[sFieldID].olRepeat)
		{
			if(oForm.olQuestion[sFieldID].olRepeat[nRepeat].nStatus!=eStatus.Requested)
			{
				//only validate fields that are not 'requested'
				var oFieldTemp=oForm.olQuestion[sFieldID];
				var oField=oForm.olQuestion[sFieldID].olRepeat[nRepeat];
				if(oField.bReVal)
				{
					if(!(fnReValidateField(sFieldID,nRepeat)))
					{
						// send undefined value to set so blanks field appropriately
						oField.set();
						setJSValue(sFieldID,"",false,nRepeat);
						// set icon
						var nDefStatus=fnGetDefaultStatus(sFieldID,nRepeat);
						setFieldStatus(sFieldID,nDefStatus,oField.nLockStatus,oField.nDiscrepancyStatus,oField.nSDVStatus,nRepeat);	// Put the icon back too
					}			
				}
				//check if need to change optional status
				//must not be new form / optional / and be empty / & not set to success / not 'not applicable'
				if((!oForm.bNewForm)&&(oFieldTemp.bOptional)&&(fnIsFieldEmpty(sFieldID,nRepeat))&&(oField.nDBStatus!=eStatus.Success)&&(oField.nDBStatus!=eStatus.NotApplicable)&&(oField.nDBStatus!=eStatus.Unobtainable))
				{
					//set status to success
					setFieldStatus(sFieldID,eStatus.Success,oField.nLockStatus,oField.nDiscrepancyStatus,oField.nSDVStatus,nRepeat);
				}
				//must not be new form / not optional / and be empty / & set to success / not 'not applicable'
				if((!oForm.bNewForm)&&(!oFieldTemp.bOptional)&&(fnIsFieldEmpty(sFieldID,nRepeat))&&(oField.nDBStatus==eStatus.Success)&&(oField.nStatus!=eStatus.Unobtainable))
				{
					//set status to missing
					setFieldStatus(sFieldID,eStatus.Missing,oField.nLockStatus,oField.nDiscrepancyStatus,oField.nSDVStatus,nRepeat);
				}
			}
			else
			{
			    if (bDealWithRequested)
			    {
			        //ic 28/02/2008 issue 2998 - set timestamps on opening eform for requested questions
			        fnReplaceAIBlock(sFieldID,"t","","","",nRepeat,"")
			    }
			}
		}
	}
	return gbReValidationOk;
}

//ic 03/11/2004
//revalidates on opening and before saving eform
//see also fnValidateField(), fnShowValidationDialog()
function fnReValidateField(sFieldID,nRepeat)
{
	//check field is checkable
	if(oForm.olQuestion==null) return false;
	if(oForm.olQuestion[sFieldID]==null) return false;
	if(oForm.olQuestion[sFieldID].olRepeat[nRepeat]==null) return false;

	//initialise vars, get additional info
	var oTField=oForm.olQuestion[sFieldID];
	var oField=oForm.olQuestion[sFieldID].olRepeat[nRepeat];
	var	nNewStatus=fnGetDefaultStatus(sFieldID,nRepeat);
	var oAIField = fnGetFieldProperty(sFieldID,"oAIHandle",nRepeat);
	var aAdd=oAIField.value.split(sDel1);
	var	bWarn=true;
	var bReject=false;
	var bChanged;
								
	if(oField.getFormatted()==="")
	{
		//dont validate empty values, but check for studydef changes that
		//effect statuses
		switch(oField.nStatus)
		{
			case eStatus.Missing:
			case eStatus.Success:
				//if ok/missing, check for changed 'optional' property
				nNewStatus=(oTField.bOptional)? eStatus.Success:eStatus.Missing;
				break;
			case eStatus.NotApplicable:
				//if not applicable, check for removed skip condition
				if((oTField.olSkip==null)||(oTField.olSkip==undefined))
				{
					//re-enable field
					setFieldEnabled(sFieldID,nRepeat,true);
				
					if(!oTField.bRQG)
					{
						nNewStatus=(oTField.bOptional)? eStatus.Success:eStatus.Missing;
					}
					else
					{
						if(aRQG[oTField.sRQG].nStatus!=eStatus.NotApplicable)
						{
							nNewStatus=(oTField.bOptional)? eStatus.Success:eStatus.Missing;
						}
					}
				}
				else
				{
					//otherwise execute the skip
					fnExecSkipOrDeriv(sFieldID,nRepeat,eDepType.Skip,true,true);
					nNewStatus=oField.nStatus;
				}
				break
			default:
		}
	}
	else
	{
		//loop through this fields validation conditions
		for(iValidation in oTField.olValidation)
		{
			bChanged=false;
			bEvaluationError=false;
			
			//set global gnCurrentRepeat
			gnCurrentRepeat=nRepeat;
			//evaluate. true means validation failed
			bWarn=eval(oTField.olValidation[iValidation].sExpression);
			gnCurrentRepeat=null;

			if((bWarn==true)&&(!bEvaluationError))
			{
				//validation failed
				//has this failure been previously handled? compare the status expected
				//from the validation failure with the current status of the field. if they
				//match and the failed validation message matches the current one, the failure
				//has been handled so ignore it
				gnCurrentRepeat=nRepeat;
				switch(oField.nStatus)
				{
					case eStatus.Inform:
						//inform
						nNewStatus=oField.nStatus;
						break;
					case eStatus.Warning:
					case eStatus.OKWarning:
						//warning/okwarning
						if((oTField.olValidation[iValidation].sType!=eValidation.Warn)||(eval(oTField.olValidation[iValidation].sMessage)!=oField.sValidationMessage))
						{
							//if either the validation type or validation message are different
							//this validation MUST be outstanding
							bChanged=true;
						}
						else
						{
							nNewStatus=oField.nStatus;
						}
						break;
					default:
						//the current question status is not inform/warning/okwarning
						//this validation MUST be outstanding
						bChanged=true;
				}
				gnCurrentRepeat=null;
				
				
				if(bChanged)
				{				
					//revalidation has changed something
					gbReValidationOk=false;
					
					if(oTField.olValidation[iValidation].sType==eValidation.Inform)
					{
						//new status is Inform - dont show a dialog
						nNewStatus=eStatus.Inform;
					}
					else
					{
						var	sResponse="";
						if((!oField.enterable())||(oTField.bHidden))
						{
							//dont show revalidation dialog for not enterable/hidden questions, just revalidate
							sResponse=oTField.olValidation[iValidation].sType;
						}
						else
						{

							//validation type: R(0)=rejectif, W(1)=warnif, I(2)=informif
							sResponse=fnValidationDialog(sFieldID,oTField.olValidation[iValidation].sType,
								oTField.olValidation[iValidation].sMessage,
								oTField.olValidation[iValidation].sNiceExpression,nRepeat,"");
						}
						
						//handle the return value
						switch(sResponse.substr(0,1))
						{
							case eValidation.OKWarn:
								//overruled
								fnSeteFormToChanged();
								fnSetRFO(sFieldID,nRepeat,sResponse.substr(1));
								fnReplaceAIBlock(sFieldID,"o","",sResponse.substr(1),"",nRepeat);
								nNewStatus=eStatus.OKWarning;
								break;
							case eValidation.Reject:
								//rejected
								// MLM 29/05/2008: Issue 3038: Only clear out value if really a 'reject'
								if("" + oTField.olValidation[iValidation].sType == eValidation.Reject)
								{
									bReject=true;
									nNewStatus=eStatus.InvalidData;
									break;
								}
							default:
							case eValidation.Warn:
								//okayed
								fnSeteFormToChanged();
								fnSetRFO(sFieldID,nRepeat,"");
								fnReplaceAIBlock(sFieldID,"o","","","",nRepeat);
								nNewStatus=eStatus.Warning;
								break;
						}
					}
				}
				
				//dont remember the validation message for rejected values as these cant stick
				if(!bReject)
				{
					//set the stored validation message to the new validation message
					gnCurrentRepeat=nRepeat;
					oField.sValidationMessage=eval(oTField.olValidation[iValidation].sMessage);
					gnCurrentRepeat=null;
				}
				//dont run any more validations if this one failed
				break;
			}
		}
	}
	
	//check the length of derived text values, reject if too long
	if((oTField.olDerivation!=null)&&(oTField.nType==etText)&&(!bReject)&&(oField.vValue!=null)&&(oField.vValue!=undefined))
	{
		if(oField.vValue.length>oTField.nLength)
		{
			if(bChanged)
			{
				alert("The value for question '"+oForm.olQuestion[sFieldID].sCaptionText+"' has been rejected because:\n\nQuestion responses may not be longer than "+oForm.olQuestion[sFieldID].nLength+" characters.");
			}
			bReject=true;
		}
	}
	
	//exit if new value was rejected
	if(bReject) return(false);
	
	//set field to new status
	setFieldStatus(sFieldID,nNewStatus,oForm.olQuestion[sFieldID].olRepeat[nRepeat].nLockStatus,oForm.olQuestion[sFieldID].olRepeat[nRepeat].nDiscrepancyStatus,oForm.olQuestion[sFieldID].olRepeat[nRepeat].nSDVStatus,nRepeat);
	
	//clear any overrule reason if the question status is not okwarning
	if(nNewStatus!=eStatus.OKWarning)
	{
		fnSetRFO(sFieldID,nRepeat,"");
		fnReplaceAIBlock(sFieldID,"o","","","",nRepeat);
	}
	
	return(true);
}

//
// Function to analyse the supplied format string,
//  pick out the part matching the expression
//  & return its numeric value
//
//  sFormat=format string(e.g.,"DD/MM/YYYY HH:MM:SS")
//  sMatch=match pattern(e.g.,"[^H].MM" for months)
//  nOffset=offset in matched text for value(starting at zero)
//  nLength=length of value in matched text
//  nDefault=default value to be use in the case of the element not being found
//
function fnStrValue(vValue,sFormat,sMatch,nOffset,nLength,nDefault)
{
	var	sFound=eval('"'+sFormat+'".search(/'+sMatch+'/)');
	var	sResult;
	var	nResult;
	if(sFound<0)
	{
		nResult=nDefault;
	}
	else
	{
		var	sResult=vValue.substr(sFound+nOffset,nLength);
		nResult=1 * sResult;
	}
	return nResult;
}

//
// Function to analyse the supplied format string,
//  pick out the part matching the expression
//  & return its character position(or -1 if not found)
//
//  sFormat=format string(e.g.,"DD/MM/YYYY HH:MM:SS")
//  sMatch=match pattern(e.g.,"[^H].MM" for months)
//  nOffset=offset in matched text for value(starting at zero)
//
function fnStrPos(sFormat,sMatch,nOffset)
{
	var	sFound=eval('"'+sFormat+'".search(/'+sMatch+'/)');
	if(sFound<0)
	{
		return(-1);
	}
	else
	{
		return(sFound+nOffset);
	}
}

//
// Return the value assigned to the specified field,
// or NULL if the field is defined but unassigned,
// or UNDEFINED if the field is not known about.
//
function fnGetFieldValue(sFieldID,nRepeat)
{
	if(oForm.olQuestion==null)
	{
		return undefined;
	}

	if(oForm.olQuestion[sFieldID]==null)
	{
		return undefined;
	}

	if(oForm.olQuestion[sFieldID].olRepeat[nRepeat]==null)
	{
		return undefined;
	}

	return(oForm.olQuestion[sFieldID].olRepeat[nRepeat].getFormatted());
}

function fnGetDefaultStatus(sFieldID,nRepeat)
{
	//ic 08/01/2004 use identity operator to compare with ""
	if(getJSValue(sFieldID,nRepeat)!=null)
	{
		if(!(getJSValue(sFieldID,nRepeat)===""))
		{
			return 0;	// Data entered - assume OK
		}
	}
	
	nNewStatus=getFieldStatus(sFieldID,nRepeat);

	if(nNewStatus!=eStatus.Requested)
	{
		if(nNewStatus!=eStatus.Unobtainable)
		{
			nNewStatus=eStatus.Missing;	// Can't go back to "fresh" status
		}
	}
	else
	{
		nNewStatus=eStatus.Requested;
	}
	return nNewStatus;
}

//
// Derive all fields and skips dependent on sFieldID,
// taking care not to re-evaluate them more than once(i.e.,avoid cyclic dependencies)
//
function fnCalculateDependencies(sFieldID,bNoRefresh,nRepeat,bInitialise)
{
	var	oDep;
	var	iDep;

	// Loop through dependent skip conditions & derivations and add them to the processing stack
	// do 'skips' firstly then 'derivations' after!!!
	for(var	iDep in oForm.olQuestion[sFieldID].olDependency)
	{
		oDep=oForm.olQuestion[sFieldID].olDependency[iDep];
		if(oDep.sType==eDepType.Skip)
		{
			// Skip or Derivation
			fnPushDependent(oDep.sDependencyField,oDep.sType);
		}
	}
	for(var	iDep in oForm.olQuestion[sFieldID].olDependency)
	{
		oDep=oForm.olQuestion[sFieldID].olDependency[iDep];
		if(oDep.sType==eDepType.Derivation)
		{
			// Skip or Derivation
			fnPushDependent(oDep.sDependencyField,oDep.sType);
		}
	}
		
	var	sDepType;
	var	sDepField;
	var	sDep;
	var	oDeriv;
	var	vValue;
	var	nNewStatus;

	while((sDep=fnPopDependent())!='')
	{
		sDepField=sDep.substr(1);
		sDepType=sDep.substr(0,1);
	
		// check for question 
		if((oForm.olQuestion[sDepField]!=null)&&(oForm.olQuestion[sDepField]!=undefined))
		{
			// question
			for(var nDepRepeat=0;nDepRepeat<oForm.olQuestion[sDepField].olRepeat.length;nDepRepeat++)
			{
				//if(oForm.olQuestion[sDepField].olRepeat[nDepRepeat].bChangable)
				if(!fnLockedOrFrozen(sDepField,nDepRepeat)&&oForm.olQuestion[sDepField].olRepeat[nDepRepeat].bChangable)		
				{
					// execute skip or derivation
					fnExecSkipOrDeriv(sDepField,nDepRepeat,sDepType,bNoRefresh,bInitialise);	
				}		
			} // end for
			
			// Move Calculate dependencies - no repeat
			fnCalculateDependencies(sDepField,bNoRefresh,undefined,bInitialise);
		}		
	}
}

//
// Pass in field info and execute skip or derivation appropriately
//
function fnExecSkipOrDeriv(sDepField,nDepRepeat,sDepType,bNoRefresh,bInitialise)
{
	var	oDeriv;
	var	vValue;
	var	nNewStatus;
	var iSkip;
	var vOldValue;
		
	if(sDepType==eDepType.Skip)
	{
		// Deal with skips
		var	bSkip=false;
		for(iSkip in oForm.olQuestion[sDepField].olSkip)
		{
			bEvaluationError=false;
			// Set global gnCurrentRepeat
			gnCurrentRepeat=nDepRepeat;
			var	vValue=eval(oForm.olQuestion[sDepField].olSkip[iSkip].sExpression)
			gnCurrentRepeat=null;
			if((!vValue)||(bEvaluationError))
			{
				bSkip=true;
			}
		}
		if(bSkip)
		{
			// disable field & set to skipped status
			setFieldDisabled(sDepField,bNoRefresh,nDepRepeat);
			nNewStatus=eStatus.NotApplicable;
			setFieldStatus(sDepField,nNewStatus,oForm.olQuestion[sDepField].olRepeat[nDepRepeat].nLockStatus,oForm.olQuestion[sDepField].olRepeat[nDepRepeat].nDiscrepancyStatus,oForm.olQuestion[sDepField].olRepeat[nDepRepeat].nSDVStatus,nDepRepeat);
		}
		else
		{
			// if not a derivation AND not a multimedia question
			if((oForm.olQuestion[sDepField].olDerivation==null)&&(oForm.olQuestion[sDepField].nType!=etMultimedia))
			{
				// set field to enabled
				setFieldEnabled(sDepField,nDepRepeat,bNoRefresh);
				
				// set field status
				nNewStatus=getFieldStatus(sDepField,nDepRepeat);
				
				// check if an optional question
				// must be optional / and be empty / & not set to success/not applicable
				if((oForm.olQuestion[sDepField].bOptional)&&(fnIsFieldEmpty(sDepField,nDepRepeat)))
				{
					// then set to OK
					nNewStatus=eStatus.Success;
				}

				setFieldStatus(sDepField,nNewStatus,oForm.olQuestion[sDepField].olRepeat[nDepRepeat].nLockStatus,oForm.olQuestion[sDepField].olRepeat[nDepRepeat].nDiscrepancyStatus,oForm.olQuestion[sDepField].olRepeat[nDepRepeat].nSDVStatus,nDepRepeat);
			}
			else
			{
				if(oForm.olQuestion[sDepField].olDerivation!=null)
				{
					// dph/ic 15/12/2003 only set derivation status for empty questions
					if(fnIsFieldEmpty(sDepField,nDepRepeat))
					{
						// check if an optional question
						if(oForm.olQuestion[sDepField].bOptional)
						{
							// then set to OK
							nNewStatus=eStatus.Success;
						}
						else
						{
							// then set to missing
							nNewStatus=eStatus.Missing;
						}

						setFieldStatus(sDepField,nNewStatus,oForm.olQuestion[sDepField].olRepeat[nDepRepeat].nLockStatus,oForm.olQuestion[sDepField].olRepeat[nDepRepeat].nDiscrepancyStatus,oForm.olQuestion[sDepField].olRepeat[nDepRepeat].nSDVStatus,nDepRepeat);
					}
				}

				// Check derivation for field
				fnPushDependent(sDepField,eDepType.Derivation);
			}
		}
	}
	else
	{			
		// Deal with derivations
		// Don't set derivations on not applicable questions
		if(oForm.olQuestion[sDepField].olRepeat[nDepRepeat].nStatus!=eStatus.NotApplicable)
		{		
			for(oDeriv in oForm.olQuestion[sDepField].olDerivation)
			{
				vOldValue=oForm.olQuestion[sDepField].olRepeat[nDepRepeat].getFormatted();
				// evaluate derivation
				bEvaluationError=false;
				// Set global gnCurrentRepeat
				gnCurrentRepeat=nDepRepeat;
				var	vValue=eval(oForm.olQuestion[sDepField].olDerivation[oDeriv].sExpression);
				gnCurrentRepeat=null;
				// if evaluation has failed default value
				if((bEvaluationError)||((typeof(vValue)=="number")&&(isNaN(vValue))))
				{
					vValue="";
				}
				//if changed from DB value then set eForm to changed
				//ic 06/01/2004 modified this condition and moved setRaw() call - dates were not being compared correctly
				oForm.olQuestion[sDepField].olRepeat[nDepRepeat].setRaw(vValue);
				// dph 03/02/2004 - retrieve vValue from field (as may have changed during setraw)
				vValue=oForm.olQuestion[sDepField].olRepeat[nDepRepeat].get();
				if(!(oForm.olQuestion[sDepField].olRepeat[nDepRepeat].getFormatted()===vOldValue))
				{
					//only do this if the derived value is different from the previous value	
					if((oForm.olQuestion[sDepField].olRepeat[nDepRepeat].vDBValue!=oForm.olQuestion[sDepField].olRepeat[nDepRepeat].getFormatted())&&(oForm.olQuestion[sDepField].olRepeat[nDepRepeat].vDBValue!=vValue))
					{
						//28/01/2004 ic/dph only flag derivations in rgqs if they are in rows below min repeats
						if(oForm.olQuestion[sDepField].bRQG)
						{
							var oRQG=aRQG[oForm.olQuestion[sDepField].sRQG];
							if((nDepRepeat+1)<=oRQG.nMinRepeats)
							{					
								fnSetDeriveFormToChanged();
							}
						}
						else
						{					
							fnSetDeriveFormToChanged();
						}
					}
					
					// if question is a radio one do not refresh...
					if(oForm.olQuestion[sDepField].nType==etCategory)
					{
						setJSValue(sDepField,oForm.olQuestion[sDepField].olRepeat[nDepRepeat].getFormatted(),true,nDepRepeat);
					}
					else
					{
						setJSValue(sDepField,oForm.olQuestion[sDepField].olRepeat[nDepRepeat].getFormatted(),false,nDepRepeat);
					}
					// set status of derived field 
					// added null in check for derived radio buttons				
					var vJSValue=getJSValue(sDepField,nDepRepeat);
					//ic 09/01/2003 compare vJSValue using identity operator to handle zeroes
					if((vJSValue==="")||(vJSValue===null))
					{
						// ic/dph 15/12/2003 check if an optional question
						if(oForm.olQuestion[sDepField].bOptional)
						{
							// set status to ok
							setFieldStatus(sDepField,eStatus.Success,oForm.olQuestion[sDepField].olRepeat[nDepRepeat].nLockStatus,oForm.olQuestion[sDepField].olRepeat[nDepRepeat].nDiscrepancyStatus,oForm.olQuestion[sDepField].olRepeat[nDepRepeat].nSDVStatus,nDepRepeat);
						}
						else
						{
							// set status to missing
							setFieldStatus(sDepField,eStatus.Missing,oForm.olQuestion[sDepField].olRepeat[nDepRepeat].nLockStatus,oForm.olQuestion[sDepField].olRepeat[nDepRepeat].nDiscrepancyStatus,oForm.olQuestion[sDepField].olRepeat[nDepRepeat].nSDVStatus,nDepRepeat);
						}
					}
					else
					{
						// validate derived field
						// if loading eForm don't show validation dialog
						//ic 03/11/2004 new fnValidateField() function
						if(!fnValidateField(sDepField,nDepRepeat,bInitialise))
						{
							// clear field as a reject
							// send undefined value to set so blanks field appropriately
							oForm.olQuestion[sDepField].olRepeat[nDepRepeat].set();
							if(oForm.olQuestion[sDepField].nType==etCategory)
							{
								setJSValue(sDepField,"",true,nDepRepeat);
							}
							else
							{
								setJSValue(sDepField,"",false,nDepRepeat);
							}
							// set icon
							setFieldStatus(sDepField,fnGetDefaultStatus(sDepField,nDepRepeat),
										oForm.olQuestion[sDepField].olRepeat[nDepRepeat].nLockStatus,
										oForm.olQuestion[sDepField].olRepeat[nDepRepeat].nDiscrepancyStatus,
										oForm.olQuestion[sDepField].olRepeat[nDepRepeat].nSDVStatus,
										nDepRepeat);	// Put the icon back too
						}
					}
				}
			}
		}
	}
}

//
// Function to reset the dependents stack.
// This stack is actually two lists operating more as a cyclic buffer than a conventional stack(FIFO rather than LIFO).
//
var	olDepFIFO=new Array();	// These need to have global scope
var	olDepGetInd;
var	olDepPutInd;
function fnEmptyDependentsStack()
{
	olDepFIFO=new Array();
	olDepGetInd=-1;
	olDepPutInd=-1
}

//
// Function to add a dependency to the dependency stack.
// Returns true if the dependency was added,otherwise false.
//
function fnPushDependent(sFieldID,sType)
{
	var	sTag=sType+sFieldID;
	for(var	iInd=0; iInd<=olDepPutInd;++iInd)
	{
		if(olDepFIFO[iInd]==sTag)
		{
			return(false);
		}
	}
	olDepFIFO [++olDepPutInd]=sTag;
	return(true);
}

//
// Function to retrieve the next value from the dependency stack.
// Returns the dependency type followed by the dependent field,all as a single string.
//
function fnPopDependent()
{
	if(olDepGetInd>=olDepPutInd)
	{
		return('');
	}
	else
	{
		return(olDepFIFO[++olDepGetInd]);	// Keep the FIFO populated,just point past dead data.
	}
}

//
// The specified field has lost the input focus.
// Perform all required validations, derivations and skips.
// Returns TRUE if focus is OK to be shifted, else FALSE.
//
//function fnLostFocus(sFieldID,sColour)
function fnLostFocus(oControl,sColour)
{
	var	vResult;
	//sFocusTarget=sFieldID;
	oFocusTarget=oControl;
	var sFieldID=oControl.name;
	if(oControl.name==undefined)
	{
		// special case Radio control
		if((oControl.oValueRef!=undefined)&&(oControl.oValueRef!=null))
		{
			oControl=oControl.oValueRef;
			sFieldID=oControl.name;
		}
	}
	var nRepeat=DefaultRepeatNo(oControl.idx);
	vResult=fnLeaveField(oControl,sColour);
	oForm.olQuestion[sFieldID].olRepeat[nRepeat].vStartValue=fnGetStartValue(sFieldID,nRepeat);
	oForm.olQuestion[sFieldID].olRepeat[nRepeat].vStartStatus=getFieldStatus(sFieldID,nRepeat);
	return vResult;
}

//
// Same as for fnLeaveField, except does not validate for category and select types
// (because they do in "on change" rather than when the field is left).
// Called by the menu buttons.
//
//function fnLeaveFieldExceptCats(sFieldID,sColour)
function fnLeaveFieldExceptCats(oControl,sColour)
{
	if(oControl==null)
	{
		return(true);
	}
	var sFieldID=oControl.name;
	var nRepeat=DefaultRepeatNo(oControl.idx);
	if(sFieldID=="")
	{
		return(true);
	}
	if((oForm.olQuestion[sFieldID].nType!=etCategory)
		&&(oForm.olQuestion[sFieldID].nType!=etCatSelect))
	{
		return fnLeaveField(oControl ,sColour);
	}
	else
	{
		return true;
	}
}

//
// Function to perform all validations, skips & derivations for the specified field.
//
// The field may or may not have the focus, and whatever the result of the processing,
// no focus shifts will be made.
//
// Returns TRUE if all is okay, or false if an error occurs (validation fails, or RFC declined, etc.).
// the colour is applied to the field's background only if TRUE is returned.
//
//function fnLeaveField(sFieldID,sColour)
function fnLeaveField(oControl,sColour)
{
	if(oControl==null)
	{
		return(true);
	}
	var sFieldID=oControl.name;
	var nRepeat=DefaultRepeatNo(oControl.idx);
	var	nInitialFieldStatus=getFieldStatus(sFieldID,nRepeat);
	var	oFieldTemp=oForm.olQuestion[sFieldID];
	var	oField=oForm.olQuestion[sFieldID].olRepeat[nRepeat];
	if(!oField.enterable()){return true;}
	var	vNewValue=getJSValue(sFieldID,nRepeat);
	if ((oForm.olQuestion[sFieldID].bIsLabField)&&(oForm.sLab==""))
	{
		fnChooseLab();
	}
	if(sColour==undefined)
	{
		sColour=oForm.sBlurColour;
	}

	//ic 07/01/2004 handle zeroes correctly
	var bExit=false;
	if(oField.vStartValue==="")
	{
		//if startvalue is "" (not 0) we must compare using identity(===) operator
		//because the equality operator(==) matches 0(zero) with ""(empty string)
		bExit=(vNewValue===oField.vStartValue);
	}
	else
	{
		//if startvalue is not "", use equality operator because identity(===) operator
		//will not match 0(zero) with 000(zero zero zero)
		bExit=(vNewValue==oField.vStartValue);
	}
	if(bExit)
	{
		//no data was changed - do nothing special
		if((oFieldTemp.nType!=etCategory)&&(oFieldTemp.nType!=etCatSelect))
		{
			setFieldBGColour(sFieldID,sColour,nRepeat);
		}
		return(true);	
	}
	
	//try to set the new value using the set member function. if it succeeds then the new value
	//is in an acceptable format and value
	if(!oField.set(vNewValue))
	{
		//new value is rejected, display a rejected dialog
		fnValidationDialog(sFieldID,eValidation.Reject,"String('The value you have entered has been rejected because it is not a valid "+oForm.olQuestion[sFieldID].olRepeat[nRepeat].getErrMes()+", it may be out of range, the format you have used may not be valid, or it may contain illegal characters.')","Please try again",nRepeat,"");
		//set the screen value back to the previous good value
		setJSValue(sFieldID,oField.getFormatted(),false,nRepeat);
		//highlight the value
		fnHighlight(oFieldTemp,oField);
		return(false);
	}
	
	//get the new value with formatting applied
	vNewValue=oField.getFormatted();
	var	bReturnValue=true;
	
	//set the screen value to the new value with formatting applied
	setJSValue(sFieldID,oField.getFormatted(),true,nRepeat);
	
	//get the current rfo, so we can revert back if something goes wrong
	var sRFOBefore = oField.sRFO
	
	//run validations on new value
	bEvaluationError=false;
	if(!fnValidateField(sFieldID,nRepeat,false))
	{
		//validation failed, set back to the previous good value
		oField.set(oField.vStartValue);
		setJSValue(sFieldID,oField.getFormatted(),false,nRepeat);
		//highlight the value
		fnHighlight(oFieldTemp,oField);
		return(false);
	}

	//see if password authorisation is required
	if(!fnPassword(sFieldID,nRepeat))
	{
		// No password entered, set back to the previous good value
		oField.set(oField.vStartValue);
		setFieldStatus(sFieldID,nInitialFieldStatus,oForm.olQuestion[sFieldID].olRepeat[nRepeat].nLockStatus,
			oForm.olQuestion[sFieldID].olRepeat[nRepeat].nDiscrepancyStatus,
			oForm.olQuestion[sFieldID].olRepeat[nRepeat].nSDVStatus,nRepeat);
		setJSValue(sFieldID,oField.getFormatted(),true,nRepeat);
		fnSetRFO(sFieldID,nRepeat,sRFOBefore);
		//highlight the value
		fnHighlight(oFieldTemp,oField);
		return(false);
	}
	
	//see if reason for change is required
	if(!fnRFC(sFieldID,nRepeat,eRFCValue))
	{
		//no reason for change entered, set back to the previous good value
		oField.set(oField.vStartValue);
		setFieldStatus(sFieldID,nInitialFieldStatus,oForm.olQuestion[sFieldID].olRepeat[nRepeat].nLockStatus,
			oForm.olQuestion[sFieldID].olRepeat[nRepeat].nDiscrepancyStatus,
			oForm.olQuestion[sFieldID].olRepeat[nRepeat].nSDVStatus,nRepeat);
		setJSValue(sFieldID,oField.getFormatted(),true,nRepeat);
		fnSetRFO(sFieldID,nRepeat,sRFOBefore);
		//highlight the value
		fnHighlight(oFieldTemp,oField);
		return(false);
	}
	
	//set the background colour to white, only for non-category questions
	if((oFieldTemp.nType!=etCategory)&&(oFieldTemp.nType!=etCatSelect))
	{
		setFieldBGColour(sFieldID,sColour,nRepeat);
	}
	
	//rqg new row check - if validated OK!
	if((oFieldTemp.bRQG)&&(bReturnValue))
	{
		var bNewRQGRow=fnRQGNewRowCheck(sFieldID);
		if(bNewRQGRow)
		{
			//put the focus back where we came from
			setFieldFocus(oControl.name,DefaultRepeatNo(oControl.idx));
			//must set startvalue as this is not done in the calling function in this case
			oField.vStartValue=vNewValue;
			bReturnValue=false;
		}
	}
	
	//calculate skips and derivations based on the new value
	fnEmptyDependentsStack();
	fnCalculateDependencies(sFieldID,false,nRepeat,false);
	
	//set eform to changed
	fnSeteFormToChanged();
	
	return bReturnValue;
}
//highlight the text in the passed field
function fnHighlight(oTField,oField)
{
	switch(oTField.nType)
	{
		case	etText:
		case	etIntegerNumber:
		case	etRealNumber:
		case	etDateTime:
					oField.oHandle.select();
				break;
		case	etCategory:
					oField.oRadio.colour();
				break;
		default:
				break;
	}
}
//
// A field has got focus.
// Validate the field we have come from (if relevant) before proceeding, set its colour and 'Starting value'.
//
// If we are not truly shifting focus (e.g., we are returning here after "failing" to move to another field)
// then do nothing other than ensuring the display is correct.
//
//function fnGotFocus(sFieldID)
function fnGotFocus(oControl)
{
	var sFieldID=oControl.name;
	if(oControl.name==undefined)
	{
		// special case Radio control
		if((oControl.oValueRef!=undefined)&&(oControl.oValueRef!=null))
		{
			oControl=oControl.oValueRef;
			sFieldID=oControl.name;
		}
	}
	var nRepeat=DefaultRepeatNo(oControl.idx);
	var	oFieldTemp=oForm.olQuestion[sFieldID];
	var	oField=oForm.olQuestion[sFieldID].olRepeat[nRepeat];
	var sFocusTarget;
	if(oFocusTarget==null)
	{
		sFocusTarget=undefined;
	}
	else
	{	
		sFocusTarget=oFocusTarget.name;
	}
	// not on button next
	bBtnNext=false;
	// Deal with the field we are leaving first
	if (sFieldID==sFocusTarget)
	{
		if(nRepeat==DefaultRepeatNo(oFocusTarget.idx))
		{
			// Not a true focus shift - more of a "return". Do no checking, etc.
			setFieldBGColour(sFieldID,oForm.sFocusColour,nRepeat);
			//sFocusTarget="";
			oFocusTarget=null;
			return true;
		}
	}
	//if((sLastFocusID!=null)&&(sLastFocusID!=""))
	if((oLastFocusID!=null))
	{
		if((oForm.olQuestion[oLastFocusID.name].nType!=etCategory)&&(oForm.olQuestion[oLastFocusID.name].nType!=etCatSelect))
		{
			if(!fnLeaveField(oLastFocusID))	// We are moving *from* a field which needs validating
			{
				setFieldFocus(oLastFocusID.name,DefaultRepeatNo(oLastFocusID.idx));	// Put the focus back where we came from
				return true;
			}
		}
		if((oLastFocusID.name!=sFieldID)||((oLastFocusID.name==sFieldID)&&(DefaultRepeatNo(oLastFocusID.idx)!=nRepeat)))
		{
			if(oForm.olQuestion[oLastFocusID.name].nType==etCategory)
			{
				oForm.olQuestion[oLastFocusID.name].olRepeat[DefaultRepeatNo(oLastFocusID.idx)].oRadio.blur();	// Make it look properly blurred
			}
			setFieldBGColour(oLastFocusID.name,oForm.sBlurColour,DefaultRepeatNo(oLastFocusID.idx));
		}
	}

	setFieldBGColour(sFieldID,oForm.sFocusColour,nRepeat);
	oField.vStartValue=fnGetStartValue(sFieldID,nRepeat);
	fnHighlight(oFieldTemp,oField);
	oCurrentFieldID=oControl;
	oLastFocusID=oControl;
	oFocusTarget=null;
	oField.vOldValue=oField.vValue;	// Note the entered value
	return true;
}

//
// Caluculate a valid "start" value for the supplied field.
// This will be the current HTML-input field's formatted value,if possible,
// otherwise the last known good value for the field.
//
function fnGetStartValue(sFieldID,nRepeat)
{
	var	vInitialValue=oForm.olQuestion[sFieldID].olRepeat[nRepeat].vValue;
	var	vResult;
	if(oForm.olQuestion[sFieldID].olRepeat[nRepeat].set(getJSValue(sFieldID,nRepeat)))
	{
		vResult=oForm.olQuestion[sFieldID].olRepeat[nRepeat].getFormatted();	// Good value - use a pretty version of it!
		if(!fnLockedOrFrozen(sFieldID,nRepeat)&&oForm.olQuestion[sFieldID].olRepeat[nRepeat].bChangable)
		{
			oForm.olQuestion[sFieldID].olRepeat[nRepeat].vValue=vInitialValue;
		}
		return(vResult);
	}
	oForm.olQuestion[sFieldID].olRepeat[nRepeat].blank();	// Duff value - bin it!
	oForm.olQuestion[sFieldID].olRepeat[nRepeat].vValue=vInitialValue;
	fnValidationDialog(sFieldID,eValidation.Reject,"String('The value you have entered has been rejected because it is not a valid "+oForm.olQuestion[sFieldID].olRepeat[nRepeat].getErrMes()+", it may be out of range, the format you have used may not be valid, or it may contain illegal characters.')","Please try again",nRepeat,"");
	return(oForm.olQuestion[sFieldID].olRepeat[nRepeat].getFormatted());
}

//ic 23/02/01
//takes an object name and colour and sets the object background colour to
//the supplied colour. function called by 'onblur' and 'onfocus' events
function setFieldBGColour(fieldID,colour,nRepeat)
{
	nRepeat=DefaultRepeatNo(nRepeat);
	var	oFieldTemp=oForm.olQuestion[fieldID];
	var oField=oFieldTemp.olRepeat[nRepeat];

	if(oField.bEnabled)
	{
		switch(oFieldTemp.nType)
		{
			case etCategory:
				oField.oRadio.colour("000000",colour,colour);
				break;
			case etCatSelect:
				for(var	i=0;i<oField.oHandle.length;i++)
				{
					oField.oHandle[i].style.color="000000";
					oField.oHandle[i].style.backgroundColor=colour;
				}
				break;
			default:
				oField.oHandle.style.backgroundColor=colour;
				oField.oHandle.style.color="000000";
		}
	}
	else
	{
		switch(oFieldTemp.nType)
		{
			case etCategory:
				oField.oRadio.colour("888888",null,null);
				break;
			case etCatSelect:
				for(var	i=0;i<oField.oHandle.length;i++)
				{
					oField.oHandle[i].style.color="888888";
				}
				break;
			default:
				oField.oHandle.style.color="888888";
		}
	}
}

//ic 04/04/01
//takes an object name and a value and updates the object with the new
//value,can handle any type of UI control: radio,text,file,select
//ic new bForce arg= force setvalue regardless of enabled state
function setJSValue(sFieldID,value,bNoRefresh,nRepeat,bForce)
{
	nRepeat=DefaultRepeatNo(nRepeat);
	var	oFieldTemp=oForm.olQuestion[sFieldID];
	var oField=oFieldTemp.olRepeat[nRepeat];
	var bForce=(bForce!=undefined)? bForce:false;
	if (oFieldTemp.bHidden)
	{
		oField.oHandle.value=fnConvertToLocale(oForm.olQuestion[sFieldID].nType,value);
	}
	else
	{
		switch(oForm.olQuestion[sFieldID].nType)
		{
			case etCategory:
				oForm.olQuestion[sFieldID].olRepeat[nRepeat].oRadio.setValue(value,bNoRefresh,bForce);
				break;
			case etCatSelect:
				for(var	i=0; i<oField.oHandle.length;++i)
				{
					if(oField.oHandle[i].value==value )
					{
						oField.oHandle.selectedIndex=i;
					}
				}
				break;
			default:
			oField.oHandle.value=fnConvertToLocale(oForm.olQuestion[sFieldID].nType,value);
		}
	}
}

//
// Return the text associated with the selected value of the category question.
// If used with a non-category question will return some other value
//(currently the entered value,but this may change).
//
function fnGetFieldText(sFieldID,nRepeat)
{
	if(sFieldID=="")
	{
		return "";
	}
	var	sCatValue=oForm.olQuestion[sFieldID].olRepeat[nRepeat].get();
	switch(oForm.olQuestion[sFieldID].nType)
	{
		case etCategory:
			return oForm.olQuestion[sFieldID].olRepeat[nRepeat].oRadio.getText();
			break;
		case etCatSelect:
			return(sCatValue==""
					? ""
					: oForm.olQuestion[sFieldID].olCatValue[sCatValue].sCatText);
			break;
		default:
			return oForm.olQuestion[sFieldID].olRepeat[nRepeat].getFormatted();
	}
}

//ic 04/04/01
//takes an object name and a value and returns the object value,
//can handle any type of UI control: radio,text,file,select
function getJSValue(sFieldID,nRepeat)
{
	// Default Repeat No for non RQG questions
	nRepeat=DefaultRepeatNo(nRepeat);
	var	oField=oForm.olQuestion[sFieldID].olRepeat[nRepeat];
	if (oForm.olQuestion[sFieldID].bHidden)
	{
		return fnConvertFromLocale(oForm.olQuestion[sFieldID].nType,oField.oHandle.value);
	}
	else
	{
		switch(oForm.olQuestion[sFieldID].nType)
		{
			case etCategory:
				return oForm.olQuestion[sFieldID].olRepeat[nRepeat].oRadio.getValue();
				break;
			case etCatSelect:
				return oField.oHandle[oField.oHandle.selectedIndex].value;
				break;
			default:
				return fnConvertFromLocale(oForm.olQuestion[sFieldID].nType,oField.oHandle.value);
		}
	}
}


function getFieldStatus(sFieldID,nRepeat)
{
	return oForm.olQuestion[sFieldID].olRepeat[nRepeat].nStatus;
}

//ic 04/04/01
//takes an object name and status and sets field object enabled or disabled
//depending on the status value passed,updates the field status image,whose
//name is the object name with '_img' appended,with the new status image
//
//ic 19/08/2002
//added hierarchical switch to display locked/frozen status
//added hierarchical switch to display discrepancy status
function setFieldStatus(sFieldID,nStatus,nLockStatus,nDiscrepancyStatus,nSDVStatus,nRepeat)
{
	//var	sName="../img/";
	var	sName="";
	var	bUnob=false;
	
	nRepeat=DefaultRepeatNo(nRepeat);

	if(!oForm.olQuestion[sFieldID].bHidden)
	{	
		switch (nLockStatus)
		{
			case eLock.Locked:
				sName+="ico_locked";
				break;
			case eLock.Frozen:
				sName+="ico_frozen";
				break;
			default:
				switch (nDiscrepancyStatus)
				{
					case eDiscStatus.Raised:
						sName+="ico_disc_raise";
						break;
					case eDiscStatus.Responded:
						sName+="ico_disc_resp";

						break;
					default:
						switch(nStatus)
						{
							case eStatus.Warning:
								sName+="ico_warn";
								break;
							case eStatus.OKWarning:
								sName+="ico_ok_warn";
								break;
							case eStatus.Inform:
								sName+=(oEP.bMonitor)? "ico_inform":"ico_ok";
								break;
							case eStatus.Missing:
								sName+="ico_missing";
								break;
							case eStatus.Unobtainable:
								sName+="ico_uo";
								break;
							case eStatus.NotApplicable:
								sName+="ico_na";
								break;
							case eStatus.Requested:
								sName+="blank_status";
								break;
							case eStatus.Success:
								sName+="ico_ok";
								break;
							default:
								sName+="blank_status";
						}
				}
		}
		//do this outside the case, otherwise questions with discrepancies
		//will never get marked as unobtainable
		if(nStatus==eStatus.Unobtainable) bUnob=true;

		fnReplaceAIBlock(sFieldID,"u","","",bUnob,nRepeat);
		// has it more than 1 repeat then use diff object
		if(oForm.olQuestion[sFieldID].olRepeat.length>1)
		{
			// array
			var oImageHandle=oForm.oOtherPage[oForm.olQuestion[sFieldID].olRepeat[nRepeat].sImageName][nRepeat];
		}
		else
		{
			// standard
			var oImageHandle=oForm.oOtherPage[oForm.olQuestion[sFieldID].olRepeat[nRepeat].sImageName];
		}
	
		if(oImageHandle!=undefined)
		{
			// DPH Changed
			// If need to display
			if(oForm.olQuestion[sFieldID].bDisplayStatusIcon)
			{
				oImageHandle.src=oTopImages[sName].src;
			}
		}
	}
	oForm.olQuestion[sFieldID].olRepeat[nRepeat].nStatus=nStatus;
}

//
// ic 23/02/01
// Sets the focus of the browser to the specified object,
// highlighting any text if applicable.
// Can handle any type of UI control.
//
function setFieldFocus(sFieldID,nRepeat)
{
	if(((sFieldID=="")||(sFieldID==null)||(sFieldID==undefined))||((nRepeat==="")||(nRepeat==null)||(nRepeat==undefined)))
	{
		return;
	}
	if(fnLoading())
	{
		return;
	}
	var	oFieldTemp=oForm.olQuestion[sFieldID];
	var	oField=oForm.olQuestion[sFieldID].olRepeat[nRepeat];

	oFocusTarget=oField.oHandle;
	if(oFieldTemp.nType==etCategory)
	{
		oField.oRadio.focus();
	}
	else
	{
		oField.oHandle.focus();
	}
}

//ic 23/02/01
//takes an object ID and disables the object,can handle any type of UI
//control. Does not affect the caption or field value.
//
//ic 19/08/2002
//added setFieldBGColour() call to change element bg colour when disabled
function setInputDisabled(sFieldID,nRepeat)
{
	var	oFieldTemp=oForm.olQuestion[sFieldID];
	var	oField=oForm.olQuestion[sFieldID].olRepeat[nRepeat];
	setFieldBGColour(sFieldID,"FFFFFF",nRepeat);
	switch(oFieldTemp.nType)
	{
		case etCategory:
			oField.oRadio.setIndex(null,false);
			oField.oRadio.enable(false,"888888",null,null);
			break;
		default:	// including etCatSelect
			oField.oHandle.disabled=true;
			oField.oHandle.readOnly=true;
			
			setFieldBGColour(sFieldID,oForm.sDisabledColour,nRepeat)
	}
}

//ic 23/02/01
//takes an object ID and disables the object,can handle any type of UI
//control. Does not affect the caption or field value.
function setInputDone(sFieldID,nRepeat)
{
	var	oFieldTemp=oForm.olQuestion[sFieldID];
	var	oField=oForm.olQuestion[sFieldID].olRepeat[nRepeat];
	switch(oFieldTemp.nType)
	{
		case etCategory:
			break;
		default:	// including etCatSelect
			oField.oHandle.disabled=false;
			oField.oHandle.readOnly=true;
	}
	setFieldBGColour(sFieldID,"FFFFFF",nRepeat);
}

//ic 23/02/01
//takes an object ID and disables the object,can handle any type of UI
//control. Also updates the caption colour and blanks the input field.
function setFieldDisabled(sFieldID,bNoRefresh,nRepeat)
{
	var	oFieldTemp=oForm.olQuestion[sFieldID];
	var	oField=oForm.olQuestion[sFieldID].olRepeat[nRepeat];
	setFieldStatus(sFieldID,eStatus.NotApplicable,oForm.olQuestion[sFieldID].olRepeat[nRepeat].nLockStatus,oForm.olQuestion[sFieldID].olRepeat[nRepeat].nDiscrepancyStatus,oForm.olQuestion[sFieldID].olRepeat[nRepeat].nSDVStatus,nRepeat);	// Mark as Not Applicable
	setInputDisabled(sFieldID,nRepeat);
	if(oField.oCapHandle!=null)
	{
		var	sTmp=oField.oCapHandle.innerHTML;
		var	oTmp=oForm.oOtherPage[oField.sImageName].src;
		sTmp=sTmp.replace(/color=#....../g,"color=#888888");
		oField.oCapHandle.innerHTML=sTmp;
		oForm.oOtherPage[oField.sImageName].src=oTmp;
	}
	oField.blankAll();
	setJSValue(sFieldID,"",bNoRefresh,nRepeat);
	oField.bEnabled=false;
}

//takes an object ID and enables the object,can handle any type of UI
//control
//
//ic 19/08/2002
//added setFieldBGColour() call to change element bg colour when enabled
function setFieldEnabled(sFieldID,nRepeat,bNoRefresh)
{
	nRepeat=DefaultRepeatNo(nRepeat);
	var	oFieldTemp=oForm.olQuestion[sFieldID];
	var oField=oFieldTemp.olRepeat[nRepeat];
	if(!oField.bEnabled)
	{
		setFieldStatus(sFieldID,eStatus.Missing,oField.nLockStatus,oField.nDiscrepancyStatus,oField.nSDVStatus,nRepeat);	// Mark as Missing if we are enabling it
	}

	oField.oHandle.disabled=false;
	if (oFieldTemp.nType==etCategory)
	{
		oField.oRadio.enable(true,"000000",null,null);
	}
	if(oField.oHandle[0]==null)
	{
		oField.oHandle.disabled=false;
		oField.oHandle.readOnly=false;
	}
	else
	{
		for(var	i=0; i<oField.oHandle.length;++i)
		{
			oField.oHandle[i].disabled=false;
			oField.oHandle[i].readOnly=false;
		}
		oField.oHandle.selectedIndex=0;
	}
	if(oField.oCapHandle!=null)
	{
		var	sTmp=oField.oCapHandle.innerHTML;
		var	oTmp=oForm.oOtherPage[oField.sImageName].src;
		sTmp=sTmp.replace(/color=#....../g,"color=#"+oField.sColour);
		oField.oCapHandle.innerHTML=sTmp;
		oForm.oOtherPage[oField.sImageName].src=oTmp;
	}
	oField.bEnabled=true;
	// Only setjsvalue for category select control
	// was causing RQG to get stuck in eternal loop on skips/vals
	// due to focussing on radio button/leaving previous field & would overflow stack space
	if(oFieldTemp.nType==etCatSelect)
	{
		setJSValue(sFieldID,oField.getFormatted(),bNoRefresh,nRepeat);
	}
	
	setFieldBGColour(sFieldID,oForm.sBlurColour,nRepeat)
}

var	oCurrentFieldID="";	// ID of the field currently with focus - needed for following function

//
// Function to clear the value for the passed field
//
function fnClearField(sFieldID,nRepeat)
{
	if((sFieldID!=null)&&(sFieldID!=""))
	{
		var	nInitialFieldStatus=getFieldStatus(sFieldID,nRepeat);
		var	oFieldTemp=oForm.olQuestion[sFieldID];
		var	oField=oForm.olQuestion[sFieldID].olRepeat[nRepeat];
		var	vPreviousValue=oField.get();
		var	vNewValue=getJSValue(sFieldID,nRepeat);
		oField.blank();
		setJSValue(sFieldID,"",false,nRepeat);
		switch(oFieldTemp.nType)
		{
		case etCategory:
			oField.oRadio.setIndex(null);
			break;
		case etCatSelect:
			oField.oHandle.selectedIndex=0;
			break;
		default:
			oField.oHandle.value="";
		}
			
		if ((oCurrentFieldID!=null)&&(sFieldID==oCurrentFieldID.name)&&(nRepeat==DefaultRepeatNo(oCurrentFieldID.idx)))
		{	
			//clearfield is current field	
			// Derive all fields dependent on this one
			if((oFieldTemp.nType==etCategory)||(oFieldTemp.nType==etCatSelect))
			{
				if(fnLeaveField(oCurrentFieldID))
				{
					setFieldStatus(sFieldID,eStatus.Missing,oForm.olQuestion[sFieldID].olRepeat[nRepeat].nLockStatus,oForm.olQuestion[sFieldID].olRepeat[nRepeat].nDiscrepancyStatus,oForm.olQuestion[sFieldID].olRepeat[nRepeat].nSDVStatus,nRepeat);
				}
				else
				{
					// Restore to previous state
					setFieldStatus(sFieldID,nInitialFieldStatus,oForm.olQuestion[sFieldID].olRepeat[nRepeat].nLockStatus,oForm.olQuestion[sFieldID].olRepeat[nRepeat].nDiscrepancyStatus,oForm.olQuestion[sFieldID].olRepeat[nRepeat].nSDVStatus,nRepeat);
					oField.setRaw(vPreviousValue);
					setJSValue(sFieldID,vNewValue,false,nRepeat);
					fnEmptyDependentsStack();
					// DPH 3rd param set to false
					fnCalculateDependencies(sFieldID,false,nRepeat,false);
					//fnDisplayStatusCaption(sFieldID,getFieldStatus(sFieldID,nRepeat),nRepeat);
				}
			}
			setFieldFocus(sFieldID,nRepeat);
		}
		else
		{		
			//clearfield isnt current field
			var oFieldHandle=(oFieldTemp.nType==etCategory)? oField.oRadio.oDataLocation:oField.oHandle;
			if(fnLeaveField(oFieldHandle))
			{
				setFieldStatus(sFieldID,eStatus.Missing,oForm.olQuestion[sFieldID].olRepeat[nRepeat].nLockStatus,oForm.olQuestion[sFieldID].olRepeat[nRepeat].nDiscrepancyStatus,oForm.olQuestion[sFieldID].olRepeat[nRepeat].nSDVStatus,nRepeat);
			}
			else
			{
				// Restore to previous state
				setFieldStatus(sFieldID,nInitialFieldStatus,oForm.olQuestion[sFieldID].olRepeat[nRepeat].nLockStatus,oForm.olQuestion[sFieldID].olRepeat[nRepeat].nDiscrepancyStatus,oForm.olQuestion[sFieldID].olRepeat[nRepeat].nSDVStatus,nRepeat);
				oField.setRaw(vPreviousValue);
				setJSValue(sFieldID,vNewValue,false,nRepeat);
				fnEmptyDependentsStack();
				// DPH 3rd param set to false
				fnCalculateDependencies(sFieldID,false,nRepeat,false);
			}
		}
	}
}

//
// Function to request a change to the field's status.
// RFC, etc may be asked for and potentially cancelled.
// Only works when setting to and from missing/requested/unobtainable.
//
function setFieldNewStatus(sFieldID,nNewStatus,nRepeat)
{
	var	nInitialFieldStatus=getFieldStatus(sFieldID,nRepeat);
	var	oFieldTemp=oForm.olQuestion[sFieldID];
	var	oField=oForm.olQuestion[sFieldID].olRepeat[nRepeat];
	
	// change status firstly - reset later if not appropriate (no rfc entered)
	setFieldStatus(sFieldID,nNewStatus,oForm.olQuestion[sFieldID].olRepeat[nRepeat].nLockStatus,oForm.olQuestion[sFieldID].olRepeat[nRepeat].nDiscrepancyStatus,oForm.olQuestion[sFieldID].olRepeat[nRepeat].nSDVStatus,nRepeat);

	// See if reason for change is required
	if(!fnRFC(sFieldID,nRepeat,eRFCStatus))
	{
		// reset field status if no RFC collected
		setFieldStatus(sFieldID,nInitialFieldStatus,oForm.olQuestion[sFieldID].olRepeat[nRepeat].nLockStatus,oForm.olQuestion[sFieldID].olRepeat[nRepeat].nDiscrepancyStatus,oForm.olQuestion[sFieldID].olRepeat[nRepeat].nSDVStatus,nRepeat);	// Put the icon back too
	}
	else
	{
		// mark eform as changed
		fnSeteFormToChanged();
	}	
}

//
// Function to set a property of the specified field.
// A property is any value identified by any valid(non-null)string and can be used for whatever is needed.
//
// Properties are never used within the JSVE.
//
function fnFd(sFieldID,sPropertyID,sPropertyValue,nRepeat)
{
	if(oForm.olQuestion==null)
	{
		oForm.olQuestion=new Array();
	}

	if(oForm.olQuestion[sFieldID]==null)
	{
		oForm.olQuestion[sFieldID]=new Object();
	}

	if(nRepeat==undefined)
	{
		if(oForm.olQuestion[sFieldID].olProperty==null)
		{
			oForm.olQuestion[sFieldID].olProperty=new Array();
		}
	
		oForm.olQuestion[sFieldID].olProperty[sPropertyID]=sPropertyValue;
	}
	else
	{
		if(oForm.olQuestion[sFieldID].olRepeat==null)
		{
			oForm.olQuestion[sFieldID].olRepeat=new Array();
		}
		if(oForm.olQuestion[sFieldID].olRepeat[nRepeat]==null)
		{
			oForm.olQuestion[sFieldID].olRepeat[nRepeat]=new Object();
		}
		
		if(oForm.olQuestion[sFieldID].olRepeat[nRepeat].olProperty==null)
		{
			oForm.olQuestion[sFieldID].olRepeat[nRepeat].olProperty=new Array();
		}
	
		oForm.olQuestion[sFieldID].olRepeat[nRepeat].olProperty[sPropertyID]=sPropertyValue;
	}
}

//
// Function to get a property of the specified field.
// Returns NULL if not found.
//
function fnGetFieldProperty(sFieldID,sPropertyID,nRepeat)
{
	if(oForm.olQuestion==null)
	{
		return(null);
	}

	if(oForm.olQuestion[sFieldID]==null)
	{
		return(null);
	}

	var	sPropertyValue;
	if(nRepeat==undefined)
	{
		if(oForm.olQuestion[sFieldID].olProperty==null)
		{
			return(null);
		}
		sPropertyValue=oForm.olQuestion[sFieldID].olProperty[sPropertyID];
	}
	else
	{
		if(oForm.olQuestion[sFieldID].olRepeat[nRepeat].olProperty==null)
		{
			return(null);
		}
		sPropertyValue=oForm.olQuestion[sFieldID].olRepeat[nRepeat].olProperty[sPropertyID];		
	}
		
	return((sPropertyValue==undefined)?null:sPropertyValue);
}

//
// Function to set a property of the form.
// These are similar to field properties,but are global and not tied to any particular question.
//
function fnFm(sPropertyID,sPropertyValue)
{
	if(oForm.olProperty==null)
	{
		oForm.olProperty=new Array();
	}
	oForm.olProperty[sPropertyID]=sPropertyValue;
}

//
// Function to get a property of the form.
// Returns NULL if not found.
//
function fnGetFormProperty(sPropertyID)
{
	if(oForm.olProperty==null)
	{
		return(null);
	}
	var	sPropertyValue;
	return(((sPropertyValue=oForm.olProperty[sPropertyID])==undefined)? null : sPropertyValue);
}

//
// ic 28/05/2002
// Function checks for MACRO illegal chars.
// Returns false if any illegal chars are found
//
function fnOnlyLegalChars(vValue)
{
	var	sIllegalChars = /[`|~"]/;
	if (sIllegalChars.exec(vValue)!=null)
	{
		return false;
	}
	else
	{
		return true;
	}
}

function fnSetDiscrepancyStatus(sFieldID,nDiscrepancyStatus,nRepeat)
{
	var oField=oForm.olQuestion[sFieldID].olRepeat[nRepeat];

	oField.nDiscrepancyStatus=nDiscrepancyStatus;
	setFieldStatus(sFieldID,oField.nStatus,oField.nLockStatus,oField.nDiscrepancyStatus,oField.nSDVStatus,nRepeat);
}

//ic 20/08/2002
//function updates sdv status for field and refreshes icon
function fnSetSDVStatus(sFieldID,nSDVStatus,nRepeat)
{
	var oField=oForm.olQuestion[sFieldID].olRepeat[nRepeat];

	oField.nSDVStatus=nSDVStatus;
	fnDisplaySDVStatus(sFieldID,nRepeat);
}

//ic 20/08/2002
//function updates sdv status icon
function fnDisplaySDVStatus(sFieldID,nRepeat)
{
	// quit if do not display icons
	if((!oForm.olQuestion[sFieldID].bDisplayStatusIcon)||(oForm.olQuestion[sFieldID].bHidden))
	{
		return;
	}

	var oFieldTemp=oForm.olQuestion[sFieldID];
	var oField=oForm.olQuestion[sFieldID].olRepeat[nRepeat];
	var nSDVStatus=oField.nSDVStatus;
	var sImage="";
	

	switch (nSDVStatus)
	{
		case eSDVStatus.Queried:
			sImage+="ico_sdv_query";
			break;
		case eSDVStatus.Planned:
			sImage+="ico_sdv_plan";
			break;
		case eSDVStatus.Complete:
			sImage+="ico_sdv_done";
			break;
		default:
			sImage+="blank";
			break;
	}
	// has it more than 1 repeat then use diff object
	if(oForm.olQuestion[sFieldID].olRepeat.length>1)
	{
		// array
		if((oForm.oOtherPage[oField.sImageSName][nRepeat])!=undefined)
		{
			(oForm.oOtherPage[oField.sImageSName][nRepeat]).src=oTopImages[sImage].src;
		}
	}
	else
	{
		// standard
		if(oForm.oOtherPage[oField.sImageSName]!=undefined)
		{
			oForm.oOtherPage[oField.sImageSName].src=oTopImages[sImage].src;
		}
	}
}

//ic 20/08/2002
//function updates note/comment status for field and refreshes icon
function fnSetNoteStatus(sFieldID,sType,bPresent,nRepeat)
{
	var oField=oForm.olQuestion[sFieldID].olRepeat[nRepeat];
	
	switch (sType)
	{
		case "n":
			oField.bNote=bPresent;
			break;
		default:
			oField.bComment=bPresent;
	} 
	
	fnDisplayNoteStatus(sFieldID,nRepeat);
}

//ic 20/08/2002
//function adds/removes a note/comment icon for a field
function fnDisplayNoteStatus(sFieldID,nRepeat)
{
	// quit if do not display icons
	if((!oForm.olQuestion[sFieldID].bDisplayStatusIcon)||(oForm.olQuestion[sFieldID].bHidden))
	{
		return;
	}
	var oFieldTemp=oForm.olQuestion[sFieldID];
	var oField=oForm.olQuestion[sFieldID].olRepeat[nRepeat];
	var sImg;
	var sAlign="";
	var nCommentNote=0;
	
	if ((oField.bComment)||(oField.bNote))
	{
		if ((oField.bComment)&&(oField.bNote))
		{
			sImg="ico_note_comment";
			nCommentNote=eCommentNote;
		}
		else if (oField.bComment)
		{
			sImg="ico_comment";
			sAlign=" top";
			nCommentNote=eComment;
		}
		else
		{
			sImg="ico_note";
			sAlign=" bottom"
			nCommentNote=eNote;
		}

		switch(oFieldTemp.nType)
		{
			case etCategory:
				//0=none,1=note,2=comment,3=both
				oField.oRadio.RadioNoteStatus(nCommentNote);
				break;
			case etCatSelect:
				// DPH - img
				// has it more than 1 repeat then use diff object
				if(oForm.olQuestion[sFieldID].olRepeat.length>1)
				{
					// array
					(oForm.oOtherPage[oField.sSelectNoteImage][nRepeat]).src=oTopImages[sImg].src;
				}
				else
				{
					// standard
					oForm.oOtherPage[oField.sSelectNoteImage].src=oTopImages[sImg].src;
				}
				break;
			case etText:
				// DPH 28/06/2005 - Bug 2412 place note / comment at current right hand side of textbox
				oField.oHandle.style.backgroundImage="url("+oTopImages[sImg].src+")";
				oField.oHandle.style.backgroundRepeat="no-repeat";
				var lTextBoxWidth=oField.oHandle.clientWidth;
				var lScrollWidth=oField.oHandle.scrollWidth;
				// if scrollable then calculate image position
				if(lScrollWidth>lTextBoxWidth)
				{
					var nIconSize=8;
					if(nCommentNote==eCommentNote)
					{
						nIconSize = 16;
					}
					var lRightPos=oField.oHandle.scrollLeft+lTextBoxWidth-nIconSize;
					oField.oHandle.style.backgroundPosition=sAlign;
					oField.oHandle.style.backgroundPositionX=lRightPos;
				}
				else
				{
					oField.oHandle.style.backgroundPosition="right"+sAlign;
				}
				break;
			default:
				oField.oHandle.style.backgroundImage="url("+oTopImages[sImg].src+")";
				oField.oHandle.style.backgroundRepeat="no-repeat";
				oField.oHandle.style.backgroundPosition="right"+sAlign;
		}
	}
	else
	{
		switch(oFieldTemp.nType)
		{
			case etCategory:
				//0=none,1=note,2=comment,3=both
				oField.oRadio.RadioNoteStatus(nCommentNote);
				break;
			case etCatSelect:
				// dph - img
				// has it more than 1 repeat then use diff object
				if(oForm.olQuestion[sFieldID].olRepeat.length>1)
				{
					// array
					(oForm.oOtherPage[oField.sSelectNoteImage][nRepeat]).src=oTopImages["blank"].src;				
				}
				else
				{
					// standard
					oForm.oOtherPage[oField.sSelectNoteImage].src=oTopImages["blank"].src;				
				}
				break;
			default:
				oField.oHandle.style.backgroundImage="none";
		}
	}
}

//ic 22/08/2002
//function displays a fields 'change counter' icon
function fnDisplayChangeStatus(sFieldID,nRepeat)
{
	// quit if do not display icons
	if((!oForm.olQuestion[sFieldID].bDisplayStatusIcon)||(oForm.olQuestion[sFieldID].bHidden))
	{
		return;
	}

	var oField=oForm.olQuestion[sFieldID].olRepeat[nRepeat];
	var nChanges=oField.nChanges;
	// dph - img
	var sImage="";
	
	switch (nChanges)
	{
		case 0:
		case 1:
			sImage+="blank_change"
			break;
		case 2:
			sImage+="ico_change1"
			break;
		case 3:
			sImage+="ico_change2"
			break;
		default:
			sImage+="ico_change3"
	}

	// has it more than 1 repeat then use diff object
	if(oForm.olQuestion[sFieldID].olRepeat.length>1)
	{
		// array
		if((oForm.oOtherPage[oField.sImageCName][nRepeat])!=undefined)
		{
			(oForm.oOtherPage[oField.sImageCName][nRepeat]).src=oTopImages[sImage].src;
		}
	}
	else
	{
		// standard
		if(oForm.oOtherPage[oField.sImageCName]!=undefined)
		{
			// dph - img
			//oForm.oOtherPage[oField.sImageCName].src=sImage+".gif";
			oForm.oOtherPage[oField.sImageCName].src=oTopImages[sImage].src;
		}
	}
	
}

// Defaults nRepeat to 0 for non repeating questions
function DefaultRepeatNo(nRepeat,bOneBased)
{
	if(nRepeat==undefined)
	{
		nRepeat=0;
	}
	else
	{
		// Adjust if coming from Arezzo (one based)
		if(bOneBased==true)
		{
			nRepeat=nRepeat-1;
		}
	}
	return nRepeat;
}

// DPH 21/10/2002
// Stores AI info in question for RQG
function fnStoreAIInfo(sFieldID,nRepeat,sValue)
{
	if(oForm.olQuestion==null)
	{
		oForm.olQuestion=new Array();
	}
	if(oForm.olQuestion[sFieldID]==null)
	{
		oForm.olQuestion[sFieldID]=new Object();
	}
	if(!oForm.olQuestion[sFieldID].bRQG)
	{
		return;
	}
	if(oForm.olQuestion[sFieldID].olRepeat==null)
	{
		oForm.olQuestion[sFieldID].olRepeat=new Array();
	}
	if(oForm.olQuestion[sFieldID].olRepeat[nRepeat]==null)
	{
		oForm.olQuestion[sFieldID].olRepeat[nRepeat]=new Object();
	}

	oForm.olQuestion[sFieldID].olRepeat[nRepeat].AIValue=sValue;
}

// ApplyRules (set icons) for RQG Refresh
function fnApplyRulesForRQG(sRQGID) 
{
	var oRQG=aRQG[sRQGID];
	var nRQGQuestions=oRQG.slQuestion.length;

	// set all status icons for questions on the eform
	fnSetAllStatuses();
	
	// perform all skips & derivs
	fnExecuteRQGSkipsDerivs(oRQG);

	// complete any required dependent skips / derivations
	for(var	i=0;i<nRQGQuestions;i++)
	{
		sFieldID=oRQG.slQuestion[i];
		fnEmptyDependentsStack();
		fnCalculateDependencies(oForm.olQuestion[sFieldID].sID,true,undefined,true);	// Calculate all dependencies
	}

	// make fields read-only
	for(var	i=0;i<nRQGQuestions;i++)
	{
		sFieldID=oRQG.slQuestion[i];
		var nNoRepeats=oForm.olQuestion[sFieldID].olRepeat.length;
		for(var nRepeat=0;nRepeat<nNoRepeats;nRepeat++)
		{
			setFieldStatus(oForm.olQuestion[sFieldID].sID,oForm.olQuestion[sFieldID].olRepeat[nRepeat].nStatus,oForm.olQuestion[sFieldID].olRepeat[nRepeat].nLockStatus,oForm.olQuestion[sFieldID].olRepeat[nRepeat].nDiscrepancyStatus,oForm.olQuestion[sFieldID].olRepeat[nRepeat].nSDVStatus,nRepeat);	// Display the appropriate icon for this field
			fnDisplayNoteStatus(sFieldID,nRepeat);
			fnDisplaySDVStatus(sFieldID,nRepeat);
			fnDisplayChangeStatus(sFieldID,nRepeat);
		
			if(oForm.olQuestion[sFieldID].nType==etMultimedia)
			{
				setInputDisabled(sFieldID,nRepeat);
				if(getFieldStatus(sFieldID,nRepeat)==eStatus.Requested)
				{
					setFieldStatus(sFieldID,eStatus.Unobtainable,oForm.olQuestion[sFieldID].olRepeat[nRepeat].nLockStatus,oForm.olQuestion[sFieldID].olRepeat[nRepeat].nDiscrepancyStatus,oForm.olQuestion[sFieldID].olRepeat[nRepeat].nSDVStatus,nRepeat);	// Set to "Unobtainable"
				}
			}
			else if(((oForm.olQuestion[sFieldID].olDerivation!=undefined)&&(oForm.olQuestion[sFieldID].olDerivation.length>0))||(!oForm.olQuestion[sFieldID].olRepeat[nRepeat].bEnabled))
			{
				setInputDisabled(sFieldID,nRepeat);
			}
		}
	}
}

// Set all status icons for fields on the form
function fnSetAllStatuses()
{
	var	sFieldID;
	var nRepeat;
	var nNewStatus;

	// loop through all questions
	for(sFieldID in oForm.olQuestion)
	{
		// loop through all repeats
		for(nRepeat=0;nRepeat<oForm.olQuestion[sFieldID].olRepeat.length;nRepeat++)
		{
			// if data cannot be changed
			if(!fnLockedOrFrozen(sFieldID,nRepeat)&&oForm.olQuestion[sFieldID].olRepeat[nRepeat].bChangable)		
			{
				// get new status for the question
				nNewStatus=fnGetDefaultStatus(sFieldID,nRepeat);
				
				// if not derivation or a multimedia question (icon set elsewhere)
				if((oForm.olQuestion[sFieldID].olDerivation==null)&&(oForm.olQuestion[sFieldID].nType!=etMultimedia))
				{
					nNewStatus=getFieldStatus(sFieldID,nRepeat);
				}
				else
				{
					// don't want to set icon here
					return;
				}
				
				// Set field status icon
				setFieldStatus(sFieldID,nNewStatus,oForm.olQuestion[sFieldID].olRepeat[nRepeat].nLockStatus,oForm.olQuestion[sFieldID].olRepeat[nRepeat].nDiscrepancyStatus,oForm.olQuestion[sFieldID].olRepeat[nRepeat].nSDVStatus,nRepeat);			
			}
		}
	}
}

//
// loop through all questions and execute derivations &/or skips
// 
function fnExecuteAllSkipsDerivs()
{
	var	sFieldID;
	var nRepeat;
	var nNewStatus;
	
	// loop through all questions
	for(sFieldID in oForm.olQuestion)
	{
		// loop through all repeats
		for(nRepeat=0;nRepeat<oForm.olQuestion[sFieldID].olRepeat.length;nRepeat++)
		{
			fnPerformSkipDeriv(sFieldID,nRepeat);
		}
	}
}

//
// loop through all questions and execute derivations &/or skips for a RQG
// 
function fnExecuteRQGSkipsDerivs(oRQG)
{
	var nRQGQuestions=oRQG.slQuestion.length;

	// complete any required dependent skips / derivations
	for(var	i=0;i<nRQGQuestions;i++)
	{
		sFieldID=oRQG.slQuestion[i];
		for(var nRepeat=0;nRepeat<oForm.olQuestion[sFieldID].olRepeat.length;nRepeat++)
		{
			fnPerformSkipDeriv(sFieldID,nRepeat);
		}
	}

}

//
// Actually perform the skip or derivation on a question
//
function fnPerformSkipDeriv(sFieldID,nRepeat)
{
	// if data cannot be changed (locked/frozen)
	if(!fnLockedOrFrozen(sFieldID,nRepeat)&&oForm.olQuestion[sFieldID].olRepeat[nRepeat].bChangable)		
	{
		// if has derivation then execute it
		if((oForm.olQuestion[sFieldID].olDerivation!=null))
		{
			fnExecSkipOrDeriv(sFieldID,nRepeat,eDepType.Derivation,true,true);
		}

		// if has skip then execute it
		if((oForm.olQuestion[sFieldID].olSkip!=null))
		{
			fnExecSkipOrDeriv(sFieldID,nRepeat,eDepType.Skip,true,false);
		}
				
	}
}

function fnGetQuestion(sFieldID)
{
	return oForm.olQuestion[sFieldID];
}
function fnSetRFC(sFieldID,nRepeat,sRFC)
{
	oForm.olQuestion[sFieldID].olRepeat[nRepeat].sRFC=sRFC;
}
function fnSetRFO(sFieldID,nRepeat,sRFO)
{
	oForm.olQuestion[sFieldID].olRepeat[nRepeat].sRFO=sRFO;
}
function fnSetComments(sFieldID,nRepeat,sComments)
{
	fnSeteFormToChanged();
	oForm.olQuestion[sFieldID].olRepeat[nRepeat].sComments=sComments;
}
function fnSetFullName(sFieldID,nRepeat,sFullName)
{
	oForm.olQuestion[sFieldID].olRepeat[nRepeat].sUserFull=sFullName;
}

function fnIsFieldEmpty(sFieldID,nRepeat)
{
	var	oFieldTemp=oForm.olQuestion[sFieldID];
	var bBlank=true;
	if(oFieldTemp.olRepeat[nRepeat]!=null)
	{
		var oFieldIns=oFieldTemp.olRepeat[nRepeat];
		// Check if Blank ("")
		//ic 09/01/2004 compare value using identity operator to handle zeroes
		if(!(oFieldIns.getFormatted()===""))
		{
			bBlank=false;
		}
	}
	if(bBlank)
	{
		return true;
	}
	else
	{
		return false;
	}
}

// Check if field is on form for Arezzo JS functions - return true or false
function IsFieldOnForm(sFieldID,nRepeat)
{
	if(oForm.olQuestion[sFieldID]==null)
	{
		return false;
	}
	if(oForm.olQuestion[sFieldID].olRepeat==null)
	{
		return false;
	}
	if(nRepeat==undefined)
	{
		// only checking question
		return true;
	}
	if(oForm.olQuestion[sFieldID].olRepeat[nRepeat]==null)
	{
		return false;
	}
	return true;
}

// returns the localdateformat (if one is set - else use existing)
function fnGetLocalFormatDate(sFormat)
{
	var sLocalFormat=fnGetFormProperty("sLocalDate");
	if((sLocalFormat==null)||(sLocalFormat==""))
	{
		// send back original format
		return sFormat;
	}
	// if format not defined (but should be)
	if((sFormat==null)||(sFormat==""))
	{
		return "";
	}
	
	var nDateFormat=0; // 1 - datetime 2 - dateonly 3 - time only
	var sDateFormat="";
	var sOrigFormat=sFormat;

	sFormat=sFormat.replace(/[Dd]+/g,"D");
	sFormat=sFormat.replace(/[Mm]+/g,"M");
	sFormat=sFormat.replace(/[Yy]+/g,"Y");
	sFormat=sFormat.replace(/[Hh]+/g,"H");
	sFormat=sFormat.replace(/[Ss]+/g,"S");

	// if format has a year as part of it is a date
	if(fnStrPos(sFormat,"Y",0)>=0)
	{
		// if format has hours then is a datetime
		if(fnStrPos(sFormat,"H",0)>=0)
		{
			nDateFormat=1;
		}
		else
		{
			// is just a date
			nDateFormat=2;
		}
	}
	else
	{
		// if format has hours then is a time
		if(fnStrPos(sFormat,"H",0)>=0)
		{
			nDateFormat=3;
		}
	}
	
	switch(nDateFormat)
	{
		case 1:
			{
				// datetime - check if seconds should be added
				if(fnStrPos(sFormat,"S",0)>=0)
				{
					sDateFormat=sLocalFormat+" hh:mm:ss";
				}
				else
				{
					sDateFormat=sLocalFormat+" hh:mm";
				}
				break;
			}
		case 2:
			{
				// date
				sDateFormat=sLocalFormat;
				break;
			}
		case 3:
			{
				// time - check if seconds should be added
				if(fnStrPos(sFormat,"S",0)>=0)
				{
					sDateFormat="hh:mm:ss";
				}
				else
				{
					sDateFormat="hh:mm";
				}
				break;
			}
		default:
			{
				sDateFormat=sLocalFormat;
				break;
			}
	}

	return sDateFormat;
}

///////////////////////////////////////////////////////////
//
//public functions and methods
//
///////////////////////////////////////////////////////////
//ic, display current lab
function fnDisplayLab()
{
	if(oForm.sLab=="")
	{
		if (window.document.all["tdlab1"]!=undefined)
		{
			window.document.all["tdlab1"].innerHTML="Laboratory:&nbsp;None selected";
			window.document.all["tdlab2"].innerHTML="<a href='javascript:fnChooseLab();'>Choose laboratory</a>";
		}
	}
	else
	{
		window.document.all["tdlab1"].innerHTML="Laboratory:&nbsp;"+oForm.sLab;
		window.document.all["tdlab2"].innerHTML="<a href='javascript:fnChooseLab();'>Change laboratory</a>";
	}
}
//ic, prompt to choose lab
function fnChooseLab()
{
	var sRtn=window.showModalDialog('LabInput.asp?site='+fnGetFormProperty("sSite"),'','dialogHeight:300px; dialogWidth:500px; center:yes; status:0; dependent,scrollbars');		
	if ((sRtn!="")&&(sRtn!=undefined)&&(sRtn!=oForm.sLab))
	{
		if (oForm.sLab!="")
		{
			if (!confirm("Changing this form's laboratory will cause all laboratory results\non the form to be revalidated on saving.\nAre you sure you wish to change the laboratory?"))
			{
				return;
			}
		}
		oForm.sLab=sRtn;
		o2.labcode.value=sRtn;
		fnDisplayLab();
	}
}
//ic, is parent form of passed field read only 
function fnFieldEformIsReadOnly(sFieldID)
{
	return (oForm.olQuestion[sFieldID].bEform)? oForm.bUReadOnly:oForm.bVReadOnly;
}
//ic 21/08/2003 is passed field on user eform (else assume on visit eform)
function fnIsOnUserEform(sFieldID)
{
	return (oForm.olQuestion[sFieldID].bEform)
}
//ic, is passed field locked or frozen
function fnLockedOrFrozen(sFieldID,nRepeat,oField)
{
	if (oField==undefined)
	{
		var oChkField=oForm.olQuestion[sFieldID].olRepeat[nRepeat];
	}
	else
	{
		var oChkField=oField;
	}
	return ((oChkField.nLockStatus==eLock.Locked)||(oChkField.nLockStatus==eLock.Frozen));
}
//MLM 25/09/02: Added. Take a FieldID and return its display value
function fnGetFormatted(sFieldID,nRepeat)
{
	return oForm.olQuestion[sFieldID].olRepeat[nRepeat].getFormatted();
}
//ic 15/10/2002
//function returns element id
function fnGetQuestionID(sFieldID)
{
	return oForm.olQuestion[sFieldID].nQuestionID;
}
// MLM 25/09/02: Added. Take a FieldID and return whether it's Enterable
// ic 23/10/2002 check element exists before reading property
// ic 21/01/2003 changed to use field enterable() method
function fnEnterable(sFieldID,nRepeat)
{
//	return oForm.olQuestion[sFieldID].enterable();
	if (!IsFieldOnForm(sFieldID,nRepeat)) return false;
	return oForm.olQuestion[sFieldID].olRepeat[nRepeat].enterable();
}

// dph 05/02/2003 - Gets discrepancy status
function fnGetDiscrepancyStatus(sFieldID,nRepeat)
{
	var oField=oForm.olQuestion[sFieldID].olRepeat[nRepeat];

	return oField.nDiscrepancyStatus;
}

// dph 05/02/2003 - Gets SDV status
function fnGetSDVStatus(sFieldID,nRepeat)
{
	var oField=oForm.olQuestion[sFieldID].olRepeat[nRepeat];

	return oField.nSDVStatus;
}

//dph 05/02/2003
//function gets note/comment status for field 
function fnGetNoteStatus(sFieldID,nRepeat)
{
	var oField=oForm.olQuestion[sFieldID].olRepeat[nRepeat];
	
	return oField.bNote;
}

// dph 05/02/2003 - Get Response Status
function fnGetResponseStatus(sFieldID,nRepeat)
{
	return getFieldStatus(sFieldID,nRepeat);
}

//ic 02/11/2004
//run the validations for a specified field. this function is called from the initialisation
//routine and after a field has changed (leaving a field) 
//see also fnReValidateField(), fnShowValidationDialog()  
//a noteable difference between this and those functions is that no rfc is requested here
function fnValidateField(sFieldID,nRepeat,bInitialise)
{
	//check field is checkable
	if(oForm.olQuestion==null) return null;
	if(oForm.olQuestion[sFieldID]==null) return null;
	if(oForm.olQuestion[sFieldID].olRepeat[nRepeat]==null) return null;
	
	//initialise vars, get additional info
	var	bWarn=true;
	var	nNewStatus=fnGetDefaultStatus(sFieldID,nRepeat);	
	var bReject=false;
	var oAIField = fnGetFieldProperty(sFieldID,"oAIHandle",nRepeat);
	var aAdd=oAIField.value.split(sDel1);
	var sRFO=oForm.olQuestion[sFieldID].olRepeat[nRepeat].sRFO;

    // dph 07/03/2006 - bug2868. check if field is 'empty' - if so do not validate
    var bEmpty=false;
    // calculate if field is empty
    if(fnIsFieldEmpty(sFieldID,nRepeat))
    {
        bEmpty=true;
    }
    if(!bEmpty)
    {
	    //loop through this fields validation conditions
	    for(iValidation in oForm.olQuestion[sFieldID].olValidation)
	    {
		    //set global gnCurrentRepeat
		    gnCurrentRepeat=nRepeat;
		    //evaluate. true means validation failed
		    bWarn=eval(oForm.olQuestion[sFieldID].olValidation[iValidation].sExpression);
		    gnCurrentRepeat=null;

		    if((bWarn!=false)&&(!bEvaluationError))
		    {
			    //validation failed
			    if(bInitialise)
			    {
				    //this is during eform initialisation - dont show dialogs
				    //deal with derivation which has failed validation but not rejected
				    if(getFieldStatus(sFieldID,nRepeat)==eStatus.OKWarning)
				    {
					    //overruled - leave as is
					    nNewStatus=getFieldStatus(sFieldID,nRepeat);
				    }
				    else
				    {
					    //deal with other types
					    switch(oForm.olQuestion[sFieldID].olValidation[iValidation].sType)
					    {
						    case eValidation.Reject:
							    // Rejected value OK
							    bReject=true;
							    nNewStatus=eStatus.InvalidData;
							    break;
						    case eValidation.Inform:
							    //inform
							    nNewStatus=eStatus.Inform;
							    break;
						    case eValidation.Warn:
						    default:
							    nNewStatus=eStatus.Warning;
							    break;			
					    }
				    }	
			    }
			    else
			    {
				    if(oForm.olQuestion[sFieldID].olValidation[iValidation].sType==eValidation.Inform)
				    {
					    //new status is Inform - dont show a dialog
					    nNewStatus=eStatus.Inform;
				    }
				    else
				    {
					    //validation type: R(0)=rejectif, W(1)=warnif, I(2)=informif
					    var	sResponse=fnValidationDialog(sFieldID,oForm.olQuestion[sFieldID].olValidation[iValidation].sType,
						    oForm.olQuestion[sFieldID].olValidation[iValidation].sMessage,
						    oForm.olQuestion[sFieldID].olValidation[iValidation].sNiceExpression,nRepeat,sRFO);
    						
    						
					    switch(sResponse.substr(0,1))
					    {
						    case eValidation.OKWarn:
							    //O: overruled
							    if((sRFO=="")||(sRFO!=sResponse.substr(1)))
							    {

								    fnSetRFO(sFieldID,nRepeat,sResponse.substr(1));
								    fnReplaceAIBlock(sFieldID,"o","",sResponse.substr(1),"",nRepeat);
								    nNewStatus=eStatus.OKWarning;
							    }
							    else
							    {
								    nNewStatus=eStatus.OKWarning;
							    }
    							
							    break;
						    case eValidation.Reject:
							    //R: rejected
							    bReject=true;
							    nNewStatus=eStatus.InvalidData;
							    break;
						    default:
						    case eValidation.Warn:
							    //W: okayed warning
							    if(sRFO!="")
							    {
								    fnSetRFO(sFieldID,nRepeat,"");
								    fnReplaceAIBlock(sFieldID,"o","","","",nRepeat);
								    nNewStatus=eStatus.Warning;
							    }
							    else
							    {
								    nNewStatus=eStatus.Warning;
							    }
							    break;
					    }
				    }
			    }

			    //dont remember the validation message for rejected values as these cant stick
			    if(!bReject)
			    {
				    //set the stored validation message to the new validation message
				    gnCurrentRepeat=nRepeat;
				    oForm.olQuestion[sFieldID].olRepeat[nRepeat].sValidationMessage=eval(oForm.olQuestion[sFieldID].olValidation[iValidation].sMessage);
				    gnCurrentRepeat=null;
			    }
			    //dont run any more validations if this one failed
			    break;
		    }
	    }
	    // dph/ic 16/02/2004 - check derived text field length & reject if necessary
	    if((oForm.olQuestion[sFieldID].olDerivation!=null)&&(oForm.olQuestion[sFieldID].nType==etText)&&(!bReject)&&(oForm.olQuestion[sFieldID].olRepeat[nRepeat].vValue!=null)&&(oForm.olQuestion[sFieldID].olRepeat[nRepeat].vValue!=undefined))
	    {
		    if(oForm.olQuestion[sFieldID].olRepeat[nRepeat].vValue.length>oForm.olQuestion[sFieldID].nLength)
		    {
			    if(bInitialise==false)
			    {
				    alert("The value for question '"+oForm.olQuestion[sFieldID].sCaptionText+"' has been rejected because:\n\nQuestion responses may not be longer than "+oForm.olQuestion[sFieldID].nLength+" characters.");
			    }
			    bReject=true;
		    }
	    }
	    if(bReject)
	    {
		    return(false);
	    }
    }
	setFieldStatus(sFieldID,nNewStatus,oForm.olQuestion[sFieldID].olRepeat[nRepeat].nLockStatus,oForm.olQuestion[sFieldID].olRepeat[nRepeat].nDiscrepancyStatus,oForm.olQuestion[sFieldID].olRepeat[nRepeat].nSDVStatus,nRepeat);
	fnReplaceAIBlock(sFieldID,"t","","",false,nRepeat);
	return(true);
}

//ic 02/11/2004
//view the validations for a specified field. this function is called from the 
//popup menu on an eform 
//see also fnReValidateField(), fnValidateField() 
function fnShowValidationDialog(sFieldID,nRepeat)
{
	if(oForm.olQuestion==null)
	{
		return null;
	}
	if(oForm.olQuestion[sFieldID]==null)
	{
		return null;
	}
	if(oForm.olQuestion[sFieldID].olRepeat[nRepeat]==null)
	{
		return null;
	}
	
	var	bWarn=true;
	var	nNewStatus=fnGetDefaultStatus(sFieldID,nRepeat);	// Work out the default status
	var bReject=false;

	var oAIField = fnGetFieldProperty(sFieldID,"oAIHandle",nRepeat);
	//get the info field value from the form, split on major delimiter
	var aAdd=oAIField.value.split(sDel1);
	var sRFO="";
				
	sRFO=oForm.olQuestion[sFieldID].olRepeat[nRepeat].sRFO;

	for(iValidation in oForm.olQuestion[sFieldID].olValidation)
	{
		// Evaluate the validation condition. False means it passed, else it didn't. (it is "warn if true")
		// Set global gnCurrentRepeat
		gnCurrentRepeat=nRepeat;
		bWarn=eval(oForm.olQuestion[sFieldID].olValidation[iValidation].sExpression);
		gnCurrentRepeat=null;

		if((bWarn!=false)&&(!bEvaluationError))
		{
			// The validation failed - find out what to do now.
			
			//validation type: R(0)=rejectif, W(1)=warnif, I(2)=informif
			var	sResponse=fnValidationDialog(sFieldID,oForm.olQuestion[sFieldID].olValidation[iValidation].sType,
						oForm.olQuestion[sFieldID].olValidation[iValidation].sMessage,
						oForm.olQuestion[sFieldID].olValidation[iValidation].sNiceExpression,nRepeat,sRFO);
			switch(sResponse.substr(0,1))
			{
				//Inform = 20 OKWarning = 25 Warning = 30 InvalidData = 40
				case eValidation.OKWarn:
					// Over-ruled. Continue through all other warnings.
					// if RFO has been set/updated set eForm to changed
					if((sRFO=="")||(sRFO!=sResponse.substr(1)))
					{
						if(fnRFC(sFieldID,nRepeat,eRFCOverrule))
						{
							fnSeteFormToChanged();
							fnSetRFO(sFieldID,nRepeat,sResponse.substr(1));
							fnReplaceAIBlock(sFieldID,"o","",sResponse.substr(1),"",nRepeat);
							nNewStatus=eStatus.OKWarning;
						}
						else
						{
							if(sRFO=="")
							{
								nNewStatus=eStatus.Warning;
							}
							else
							{
								nNewStatus=eStatus.OKWarning;
							}
						}
					}
					else
					{
						nNewStatus=eStatus.OKWarning;
					}
					
					break;
				case eValidation.Inform:
					// Inform status
					nNewStatus=eStatus.Inform;
					break;
				case eValidation.Reject:
					// Rejected value OK
					bReject=true;
					nNewStatus=eStatus.InvalidData;
					break;
				default:
				case eValidation.Warn:
					// Pressed "OK" to warning. Continue through all other warnings.
					// if RFO has been removed set eForm to changed
					if(sRFO!="")
					{
						if(fnRFC(sFieldID,nRepeat,eRFCOverrule))
						{
							fnSeteFormToChanged();
							fnSetRFO(sFieldID,nRepeat,"");
							fnReplaceAIBlock(sFieldID,"o","","","",nRepeat);
							nNewStatus=eStatus.Warning;
						}
						else
						{
							nNewStatus=eStatus.OKWarning;
						}
					}
					else
					{
						nNewStatus=eStatus.Warning;
					}
					break;
			}

			//dont remember the validation message for rejected values as these cant stick
			if(!bReject)
			{
				//set the stored validation message to the new validation message
				gnCurrentRepeat=nRepeat;
				oForm.olQuestion[sFieldID].olRepeat[nRepeat].sValidationMessage=eval(oForm.olQuestion[sFieldID].olValidation[iValidation].sMessage);
				gnCurrentRepeat=null;
			}
			//dont run any more validations if this one failed
			break;
		}
	}
	// dph/ic 16/02/2004 - check derived text field length & reject if necessary
	if((oForm.olQuestion[sFieldID].olDerivation!=null)&&(oForm.olQuestion[sFieldID].nType==etText)&&(!bReject)&&(oForm.olQuestion[sFieldID].olRepeat[nRepeat].vValue!=null)&&(oForm.olQuestion[sFieldID].olRepeat[nRepeat].vValue!=undefined))
	{
		if(oForm.olQuestion[sFieldID].olRepeat[nRepeat].vValue.length>oForm.olQuestion[sFieldID].nLength)
		{
			if(bInitialise==false)
			{
				alert("The value for question '"+oForm.olQuestion[sFieldID].sCaptionText+"' has been rejected because:\n\nQuestion responses may not be longer than "+oForm.olQuestion[sFieldID].nLength+" characters.");
			}
			bReject=true;
		}
	}
	if(bReject)
	{
		return(false);
	}

	setFieldStatus(sFieldID,nNewStatus,oForm.olQuestion[sFieldID].olRepeat[nRepeat].nLockStatus,oForm.olQuestion[sFieldID].olRepeat[nRepeat].nDiscrepancyStatus,oForm.olQuestion[sFieldID].olRepeat[nRepeat].nSDVStatus,nRepeat);
	fnReplaceAIBlock(sFieldID,"t","","",false,nRepeat);
	return(true);
}

// Show the warning dialog
//  sType=reject,warn,inform (see eValidation enum)
//  sMessage=the warning message
//  sExpression=the "nice looking" version of the AREZZO expression - the one entered in SD.
//  sRFO=Reason for Overrule (if set)
// Returns:
//  "W" - okayed warning
//  "O" - overruled warning
//  "R" - rejected value
//	"I" - Inform
function fnValidationDialog(sFieldID,sType,sMessage,sExpression,nRepeat,sRFO)
{
	//get an evaluation of the message
	gnCurrentRepeat=nRepeat;
	sMessage=eval(sMessage);
	gnCurrentRepeat=null;
	//check the passed rfo parameter
	if((sRFO==undefined)||(sRFO==null))
	{
		sRFO="";
	}
	//ensure string type
	sType=""+sType;
	
	//get database and site parameters	
	var sDb=fnGetFormProperty("sDatabase");
	var sSt=fnGetFormProperty("sStudyId");	

	//create array of parameters to pass to rejectwarninform page
	var aArg=new Array();
	aArg[0]=fnGetFieldProperty(sFieldID,'sCaptionText');
	aArg[1]=sType;
	aArg[2]=sMessage;
	aArg[3]=sExpression;
	aArg[5]=sRFO;
	aArg[8]=fnOverruleWarnings(sFieldID);
	//call rejectwarninform page
	var sResult=window.showModalDialog('../sites/'+sDb+'/'+sSt+'/RejectWarnInform.htm',aArg,'dialogHeight:300px; dialogWidth:500px; center:yes; status:0; dependent,scrollbars');
	
	//return result
	return(sResult);
}

//ic 14/01/2004 check all mandatory questions are complete
function fnAllMandQsComplete()
{
	return (fnMandQsComplete()&&fnRQGMandQsComplete());
}
//ic 14/01/2004 check individual question is complete
function fnMandQComplete(sFieldID,nRepeat)
{
	return(!((fnIsFieldEmpty(sFieldID,nRepeat))&&(getFieldStatus(sFieldID,nRepeat)!=eStatus.NotApplicable)&&(getFieldStatus(sFieldID,nRepeat)!=eStatus.Unobtainable)));
}
//ic 14/01/2004 check if non RQG mandatory questions are complete
function fnMandQsComplete()
{
	var bMandOK=true;
	var	sFieldID;
	for(sFieldID in oForm.olQuestion)
	{
		//only check mandatory questions that do not belong to RQGs
		if((oForm.olQuestion[sFieldID].bMandatory)&&(!oForm.olQuestion[sFieldID].bRQG))
		{
			if(!fnMandQComplete(sFieldID,0))
			{
				bMandOK=false;
			}
		}
	}
	return (bMandOK);
}
//ic 14/01/2004 check if RQG mandatory questions are complete
function fnRQGMandQsComplete()
{
	var bMandOK=true;
	var sRQGID;
	var oRQG=new Object();
	var oField=new Object();
	var nRepeat;
	var nField;
	var nCheckToRepeat=0;
	var nCreatedRepeats=0;
	
	if(aRQG!=null)
	{
		for(sRQGID in aRQG)
		{
			oRQG=aRQG[sRQGID];
			if(oRQG.bMandatory)
			{
				//get the number of repeats that have actually been initialised in the jsve
				nCreatedRepeats=oForm.olQuestion[oRQG.slQuestion[0]].olRepeat.length;
				
				if(nCreatedRepeats<oRQG.nMinRepeats)
				{
					//the minimum number of rows hasnt been created so cannot have been completed
					bMandOK=false;
				}
				else
				{
					//find the highest repeat containing user entered, non-derived data
					for(nRepeat=0;nRepeat<nCreatedRepeats;nRepeat++)
					{
						for(nField=0;nField<oRQG.slQuestion.length;nField++)
						{
							oField=oForm.olQuestion[oRQG.slQuestion[nField]].olRepeat[nRepeat];
							if((oField.olDerivation==undefined)&&(!(oField.getFormatted()==="")))
							{
								nCheckToRepeat=nRepeat;
							}
						}
					}
				
					if(nCheckToRepeat<(oRQG.nMinRepeats-1))
					{
						//check at least the min number of repeats
						nCheckToRepeat=(oRQG.nMinRepeats-1);
					}
				
					//now check that these repeats have all mandatory questions completed
					for(nRepeat=0;nRepeat<=nCheckToRepeat;nRepeat++)
					{
						for(nField=0;nField<oRQG.slQuestion.length;nField++)
						{
							if(oForm.olQuestion[oRQG.slQuestion[nField]].bMandatory)
							{
								if(!fnMandQComplete(oRQG.slQuestion[nField],nRepeat))
								{
									bMandOK=false;
								}
							}
						}
					}
				}
			}
		}
	}
	
	return (bMandOK);
}


// Return Reason For Overrule
function fnGetRFO(sFieldID,nRepeat)
{
	var sRFO="";
	var oAIField = fnGetFieldProperty(sFieldID,"oAIHandle",nRepeat);
	//get the info field value from the form, split on major delimiter
	if(oAIField!=null)
	{
		var aAdd=oAIField.value.split(sDel1);
		var aItm;
					
		//extract the existing type block (if any) from the info field
		for (var n=0;n<aAdd.length;n++)
		{
			//split block on minor delimeter, check first item denoting 'type'
			// looking for 'o'
			aItm = aAdd[n].split(sDel2);
			if(aItm[0]=="o")
			{
				//get reason for overrule
				sRFO=aItm[1];
			}
		}
	}	
	return sRFO;
}

// Return Status string of a response (only needed for warning/inform)
function fnGetStatusString(nStatus)
{
	var sStatus="";
	switch(nStatus)
	{
		//Inform = 20 OKWarning = 25 Warning = 30 InvalidData = 40
		case 20:
					sStatus="Inform";
					break;
		case eStatus.OKWarning:
					sStatus="OK Warning";
					break;
		case eStatus.Warning:
					sStatus="Warning";
					break;
		default:
					sStatus="UNKNOWN";
					break;
	}
	return sStatus;
}

// Return validation message (if there is one)
function fnGetValidationMessage(sFieldID,nRepeat,nStatus)
{
	var sValMessage="";
	// if not inform/warning then return
	if((nStatus==eStatus.Inform)||(nStatus==eStatus.OKWarning)||(nStatus==eStatus.Warning))
	{
		var	bWarn=true;

		for(iValidation in oForm.olQuestion[sFieldID].olValidation)
		{
			// Evaluate the validation condition. False means it passed, else it didn't. (it is "warn if true")
			// Set global gnCurrentRepeat
			gnCurrentRepeat=nRepeat;
			bWarn=eval(oForm.olQuestion[sFieldID].olValidation[iValidation].sExpression);
			gnCurrentRepeat=null;
			if((bWarn!=false)&&(!bEvaluationError))
			{
				var sMessage=oForm.olQuestion[sFieldID].olValidation[iValidation].sMessage
				// Evaluate message firstly
				gnCurrentRepeat=nRepeat;
				sValMessage=eval(sMessage);
				gnCurrentRepeat=null;
				return sValMessage;
			}
		}
	}
	return sValMessage;
}
//decimal delimiter
function fnDP()
{
	var n=(1/2)
	n=n.toLocaleString();
	return n.substr(1,1);
}
//thousands delimiter
function fnTS()
{
	var n=1000;
	n=n.toLocaleString();
	return n.substr(1,1);
}
//covert numbers from standard to locale specific
function fnConvertToLocale(nType,vValue)
{
	switch(nType)
	{
		case etIntegerNumber:
		case etRealNumber:
		case etLabTest:
			if((vValue!="")&&(vValue!=null)&&(vValue!=undefined))
			{
				var sNonLocaleValue=vValue;
				var sLocaleValue="";
				for (var n=0;n<sNonLocaleValue.length;n++)
				{
					if(sNonLocaleValue.substr(n,1)==".")
					{
						sLocaleValue+=oForm.sDecimalPoint
					}
					else if (sNonLocaleValue.substr(n,1)==",")
					{
						sLocaleValue+=oForm.sThousandSeparator
					}
					else
					{
						sLocaleValue+=sNonLocaleValue.substr(n,1);
					}
				}
				return sLocaleValue;
			}
			else
			{
				return vValue;
			}
			break;
		default:
			break;
	}
	return vValue;
}
//covert numbers from locale specific to standard
function fnConvertFromLocale(nType,vValue)
{
	switch(nType)
	{
		case etIntegerNumber:
		case etRealNumber:
		case etLabTest:
			if((vValue!="")&&(vValue!=null)&&(vValue!=undefined))
			{
				var sLocaleNewValue=vValue+"";
				var sNonLocaleNewValue="";
				for (var n=0;n<sLocaleNewValue.length;n++)
				{
					if(sLocaleNewValue.substr(n,1)==oForm.sDecimalPoint)
					{
						sNonLocaleNewValue+="."
					}
					else if (sLocaleNewValue.substr(n,1)==oForm.sThousandSeparator)
					{
						sNonLocaleNewValue+=","
					}
					else
					{
						sNonLocaleNewValue+=sLocaleNewValue.substr(n,1);
					}
				}
				return sNonLocaleNewValue*1;
			}
			else
			{
				return vValue;
			}
			break;
		default:
			return vValue;
			break;
	}
}

//hide all select lists that overlap the popup menu (except the passed one)
function fnHideSelects(b,sID,nRpt,bHideAll)
{
	var sFieldID;
	var nRepeat;
	var oFieldTemp;
	var oField;
	var oDiv;
	var nFLeft
	var nFTop;
	var nFRight;
	var nFBottom;
	oMenu=document.all["divPopMenu"];
	nLeft=oMenu.style.pixelLeft;
	nTop=oMenu.style.pixelTop;
	nRight=nLeft+oMenu.clientWidth;
	nBottom=nTop+oMenu.clientHeight;
	bHideAll=(bHideAll==undefined)? false:bHideAll;
	
	if(sID!=undefined)
	{
		//the field that is popping
		var oSelFieldTemp=oForm.olQuestion[sID];
		var oSelField=oSelFieldTemp.olRepeat[nRpt];
		if(oSelFieldTemp.bRQG)
		{
			var oSelRQG=aRQG[oSelFieldTemp.sRQG];
			var nSelCol=fnGetRQGCol(oSelRQG,sID)
		}
	}
	
	for(sFieldID in oForm.olQuestion)
	{
		oFieldTemp=oForm.olQuestion[sFieldID];
		if((oFieldTemp.nType==etCatSelect)&&(!oFieldTemp.bHidden))
		{
			for(nRepeat in oFieldTemp.olRepeat)
			{
				oField=oForm.olQuestion[sFieldID].olRepeat[nRepeat];
				if(b)
				{
					if((sFieldID==sID)&&(nRepeat==nRpt))
					{
						//dont hide the question we're popping a menu for
					}
					else
					{
						//hide
						if(oFieldTemp.bRQG)
						{
							//field belongs to rqg
							if(oFieldTemp.sRQG==oSelFieldTemp.sRQG)
							{
								//this field AND popping field belong to same RQG - only hide selects
								//that are below and to the left of popping field
								if(((nRpt<=nRepeat)&&(nSelCol<fnGetRQGCol(oSelRQG,sFieldID)))||(bHideAll))
								{
									//possible overlap
									oForm.olQuestion[sFieldID].olRepeat[nRepeat].oHandle.style.visibility='hidden';
								}
							}
							else
							{
								//this field doesnt belong to same RQG as popping field OR popping
								//field doesnt belong to RQG at all - check for popup menu overlap
								//of RQG container
								oDiv=eval(oFieldTemp.sRQG+"_RQGDiv");
								
								//work out if popup menu overlaps this select control
								nFLeft=oDiv.offsetLeft;
								nFTop=oDiv.offsetTop;
								nFRight=nFLeft+oDiv.clientWidth-40;
								nFBottom=nFTop+oDiv.clientHeight;
					
								if((nFBottom>nTop)&&(nFTop<nBottom)&&(nFRight>nLeft)&&(nFLeft<nRight))
								{
									//overlap
									oForm.olQuestion[sFieldID].olRepeat[nRepeat].oHandle.style.visibility='hidden';
								}	
							}
						}
						else
						{
							//field doesnt belong to rqg
							oDiv=oField.oHandle.parentElement.parentElement.parentElement.parentElement.parentElement;
							
							//work out if popup menu overlaps this select control
							nFLeft=oDiv.offsetLeft;
							nFTop=oDiv.offsetTop;
							nFRight=nFLeft+oDiv.clientWidth-40;
							nFBottom=nFTop+oDiv.clientHeight;
					
							if((nFBottom>nTop)&&(nFTop<nBottom)&&(nFRight>nLeft)&&(nFLeft<nRight))
							{
								//overlap
								oForm.olQuestion[sFieldID].olRepeat[nRepeat].oHandle.style.visibility='hidden';
							}	
						}
					}
				}
				else
				{
					//show
					oField.oHandle.style.visibility='visible';
				}
			}
		}
	}
}

//returns the rqg column index for a passed field id
function fnGetRQGCol(oRQG,sID)
{
	for(var n=0;n<oRQG.slQuestion.length;n++)
	{
		if(oRQG.slQuestion[n]==sID)
		{
			return n
		}
	}
}

//returns a list of fields whose values do not match the stored database value
function fnGetChangedList()
{
	var sList="";
	
	for(var	sFieldID in oForm.olQuestion)
	{
		for(var nRepeat=0;nRepeat<oForm.olQuestion[sFieldID].olRepeat.length;nRepeat++)
		{
			if(oForm.olQuestion[sFieldID].olRepeat[nRepeat].vDBValue!=oForm.olQuestion[sFieldID].olRepeat[nRepeat].getFormatted())
			{
				sList+=sFieldID+","+oForm.olQuestion[sFieldID].olRepeat[nRepeat].vDBValue+","+oForm.olQuestion[sFieldID].olRepeat[nRepeat].getFormatted()+"\n";
			}
		}
	}
	return sList;			
}

//function returns boolean: is the eform new and unsaved.
//if so, we will need to ignore the 'changed' flags and save anyway
function fnNewEForm()
{
	return (oForm.bNewForm); 
}
//function calls initial focus function after a slight pause.
//the pause is required and allows the page to settle. the call
//to set focus has no effect without a pause
function fnInitialFocus(sFieldID,nRepeat)
{
	window.setTimeout('fnSetInitialFocus("'+sFieldID+'",'+nRepeat+');',200);
}
//function sets initial focus to passed field
function fnSetInitialFocus(sFieldID,nRepeat)
{
	if(fnEnterable(sFieldID,nRepeat)) 
	{
		if(oForm.olQuestion[sFieldID].nType!=etCategory)
		{
			fnGotFocus(oForm.olQuestion[sFieldID].olRepeat[nRepeat].oHandle);
		}
		setFieldFocus(sFieldID,nRepeat);	
	}
}
// function to collect scroll event for text boxes with (potential) note / comments
function fnDisplayNoteStatusScroll(oControl)
{
	var sFieldID=oControl.name;
	var nRepeat=DefaultRepeatNo(oControl.idx);
	var oField=oForm.olQuestion[sFieldID].olRepeat[nRepeat];
	// display if comment or note exists
	if ((oField.bComment)||(oField.bNote))
	{
		fnDisplayNoteStatus(sFieldID,nRepeat);
	}
}

//ic 28/02/2008 issue 2996 - add an inactive category value to a select list and select it
function fnAddInactiveCategoryValue(sFieldID,nRepeat,sCode,sValue)
{
    var	oFieldTemp=oForm.olQuestion[sFieldID];
    var	oField=oForm.olQuestion[sFieldID].olRepeat[nRepeat];
    
    switch(oFieldTemp.nType)
	{
	case etCategory:
		//radio
		break;
	case etCatSelect:
	    //dropdown
	    var op = o1.document.createElement("OPTION");
	    op.text = sValue;
	    op.value = sCode;
	    op.selected = true;
		oField.oHandle.options.add(op);
		for(var	i=0; i<oField.oHandle.length;++i)
		{
			if(oField.oHandle[i].value==sCode )
			{
				oField.oHandle.selectedIndex=i;
			}
		}
		break;
	default:
	}
}