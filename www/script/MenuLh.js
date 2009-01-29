///////////////////////////////////////////////////////////////////////////////////////////
//
//	(c) InferMed 2003
//
//	MenuLh event handlers
//
///////////////////////////////////////////////////////////////////////////////////////////

var oMain=window.parent.window.frames[4];
var sDel1="|";
var sDel2="`";
var undefined;
var bShownGenHTMLMessage=false;

function fnSetVisible(sID,bVisible)
{
	document.all[sID].style.visibility=(bVisible)? "visible":"hidden";
}
function fnSetTop(sID,nTop)
{
	document.all[sID].style.top=nTop;
}	
//function toggles search panes visible/hidden
function fnToggleSearch(oDiv)
{
	if (oDiv.isExpanded)
	{
		var oSelDiv=fnGetSelectedButton("divSearchBtn")
		if (oSelDiv!=undefined) fnSetButtonSelected(oSelDiv)
	}
	else
	{
		fnHideSearch();
	}
}
function fnHideSearch()
{
	fnSetVisible("searchDiv0",0); //refresh button,study,site,subject label
	fnSetVisible("searchDiv1",0); //visit,eform,question
	fnSetVisible("searchDiv2",0); //question status checkboxes
	fnSetVisible("searchDiv3",0); //before,after,time label
	fnSetVisible("searchDiv4",0); //raised,responded,closed
	fnSetVisible("searchDiv5",0); //planned,done,queried,cancelled
	fnSetVisible("searchDiv6",0); //public,private
	fnSetVisible("searchDiv8",0); //subject,visit,eform,question
	fnSetVisible("searchDiv9",0); //disc,sdv,note
	fnSetVisible("searchDiv10",0); //locked,frozen
	fnSetVisible("searchDiv11",0); //comment
	fnSetTop("searchDiv1",0);
	fnSetTop("searchDiv2",0)
	fnSetTop("searchDiv3",0)
	fnSetTop("searchDiv4",0)
	fnSetTop("searchDiv5",0)
	fnSetTop("searchDiv6",0)
	fnSetTop("searchDiv8",0)
	fnSetTop("searchDiv9",0)
	fnSetTop("searchDiv10",0)
	fnSetTop("searchDiv11",0)
	
}
//function called for search type button clicks. shows/hides search options
function fnDoButtonClick(ndx)
{
	var nSearchIndex;
	var nHeight;
	
	fnHideSearch();
	switch (ndx)
	{
		case 0:	//Disc
			fnSetTop("searchDiv1",136);
			fnSetTop("searchDiv3",235);
			fnSetTop("searchDiv4",290);
			fnSetVisible("searchDiv0",1);
			fnSetVisible("searchDiv1",1);
			fnSetVisible("searchDiv3",1);
			fnSetVisible("searchDiv4",1);
			nHeight=345;
			break;
		case 1: //SDV
			fnSetTop("searchDiv1",136);		
			fnSetTop("searchDiv3",235);
			fnSetTop("searchDiv5",290);
			fnSetTop("searchDiv8",340);																																																	 
			fnSetVisible("searchDiv0",1);
			fnSetVisible("searchDiv1",1);
			fnSetVisible("searchDiv3",1);
			fnSetVisible("searchDiv5",1);
			fnSetVisible("searchDiv8",1);
			nHeight=395;
			break;
		case 2: //Notes	
			fnSetTop("searchDiv1",136);
			fnSetTop("searchDiv3",235);
			fnSetTop("searchDiv6",290);																																																		
			fnSetVisible("searchDiv0",1);
			fnSetVisible("searchDiv1",1);
			fnSetVisible("searchDiv3",1);
			fnSetVisible("searchDiv6",1);
			nHeight=325;
			break;
		case 3: //Subject
			fnSetVisible("searchDiv0",1);
			nHeight=140;
			break;
		case 4: //Data
			fnSetTop("searchDiv2",136);
			fnSetTop("searchDiv9",261);
			fnSetTop("searchDiv10",357);
			fnSetTop("searchDiv1",386);
			fnSetTop("searchDiv3",485);
			fnSetTop("searchDiv11",332);
			fnSetVisible("searchDiv0",1);
			fnSetVisible("searchDiv2",1);  
			fnSetVisible("searchDiv9",1);  
			fnSetVisible("searchDiv10",1); 
			fnSetVisible("searchDiv1",1);  
			fnSetVisible("searchDiv3",1);
			fnSetVisible("searchDiv11",1);
			nHeight=540;
			break;
		case 5: // Audit
			fnSetTop("searchDiv2",136);
			fnSetTop("searchDiv11",261);
			fnSetTop("searchDiv10",287);
			fnSetTop("searchDiv1",315);
			fnSetTop("searchDiv3",414);
			fnSetVisible("searchDiv0",1);
			fnSetVisible("searchDiv2",1);  
			fnSetVisible("searchDiv11",1);  
			fnSetVisible("searchDiv10",1); 
			fnSetVisible("searchDiv1",1);  
			fnSetVisible("searchDiv3",1);
			nHeight=469;
			break;
	}
	for (var n=0;n<document.all["divMenu"].length-1;n++)
	{
		if(document.all["divMenu"][n].name=="search")
		{		
			nSearchIndex=n
		}
	}
	document.all["divMenuPane"][nSearchIndex].style.height=nHeight;
	fnSpaceMenus();
}
//function builds a url from the search criteria
function fnSearch(bMonitor)
{
	var sSt=fltSt.value;
	var sSi=fltSi.value;
	var sVi=fltVi.value;
	var sEf=fltEf.value;
	var sQu=fltQu.value;
	var sSj=fltLb.value;
	var sUs=fltUs.value;
	var sDi=fltDi.value;
	var sSd=fltSD.value;
	var sNo=fltNo.value;
	var sSs= (fltSs[0].checked)?"1":"0";
	    sSs+=(fltSs[1].checked)?"1":"0";
	    sSs+=(fltSs[2].checked)?"1":"0";
	    sSs+=(fltSs[3].checked)?"1":"0";
	    sSs+=(fltSs[4].checked)?"1":"0";
	    sSs+=(fltSs[5].checked)?"1":"0";
	    sSs+=(fltSs[6].checked)?"1":"0";
	var sLk=(fltLk[0].checked)?"1":"0";
		sLk+=(fltLk[1].checked)?"1":"0";
	var sB4=fltB4[0].checked;
	var sTm=fltTm.value;
	var sDSs= (fltDSs[0].checked)?"1":"0";
	    sDSs+=(fltDSs[1].checked)?"1":"0";
	    sDSs+=(fltDSs[2].checked)?"1":"0";
	var sSSs= (fltSSs[0].checked)?"1":"0";
		sSSs+=(fltSSs[1].checked)?"1":"0";
		sSSs+=(fltSSs[2].checked)?"1":"0";
		sSSs+=(fltSSs[3].checked)?"1":"0";
	var sNSs= (fltNSs[0].checked)?"1":"0";
		sNSs+=(fltNSs[1].checked)?"1":"0";	
	var sCm=fltCm.value;
	var sObj= (fltObj[0].checked)?"1":"0";
		sObj+=(fltObj[1].checked)?"1":"0";
		sObj+=(fltObj[2].checked)?"1":"0";
		sObj+=(fltObj[3].checked)?"1":"0";
		var oSelDiv=fnGetSelectedButton("divSearchBtn");
	if (!fnOnlyLegalChars(sSj))
	{
		alert("Search string cannot contain the following characters: `|~\"%");
		return;
	}
	if (!fnOnlyLegalChars(sTm))
	{
		alert("Time string cannot contain the following characters: `|~\"%");
		return;
	}
	sSj=escape(sSj);
	sTm=escape(sTm);
	
	switch (oSelDiv.ndx)
	{
		case 0:
			if (sDSs=="000")
			{	
				alert("Please select a discrepancy search status")
				return;
			}
			fnDiscUrl(sSt,sSi,sVi,sEf,sQu,sSj,sUs,sB4,sTm,sDSs);
			break;
		case 1:
			if (sSSs=="0000")
			{	
				alert("Please select an SDV search status")
				return;
			}
			if (sObj=="0000")
			{	
				alert("Please select an SDV scope status")
				return;
			}
			fnSDVUrl(sSt,sSi,sVi,sEf,sQu,sSj,sUs,sB4,sTm,sSSs,sObj);
			break;
		case 2:
			if (sNSs=="00")
			{	
				alert("Please select a note search status")
				return;
			}
			fnNoteUrl(sSt,sSi,sVi,sEf,sQu,sSj,sUs,sB4,sTm,sNSs);
			break;
		case 3:
			fnSubjectListUrl(sSt,sSi,sSj)
			break;
		case 4:
			if (sSs=="0000000")
			{	
				alert("Please select a search status")
				return;
			}
			if ((sSj=="")&&(!bMonitor))
			{
				alert("Please enter a subject label in the search criteria");
				fltLb.focus();
				return;
			}
			fnBrowserUrl(sSt,sSi,sVi,sEf,sQu,sSj,sUs,sSs,sLk,sB4,sTm,sCm,sDi,sSd,sNo,"1");
			break;
		case 5:
			if (sSs=="0000000")
			{	
				alert("Please select a search status")
				return;
			}
			if ((sSj=="")&&(!bMonitor))
			{
				alert("Please enter a subject label in the search criteria");
				fltLb.focus();
				return;
			}
			fnBrowserUrl(sSt,sSi,sVi,sEf,sQu,sSj,sUs,sSs,sLk,sB4,sTm,sCm,"","","","2");
			break;
	}
}
function fnOnlyLegalChars(vValue)
{
	var	sIllegalChars = /[`|~"%]/;
	if (sIllegalChars.exec(vValue)!=null)
	{
		return false;
	}
	else
	{
		return true;
	}
}
function fnOnlyNumeric(sStr)
{
	var sPattern = /[^0-9]/;
	return sPattern.exec(sStr);
}
//function fnLoadSelect(sName,sList,bClear,bCodeAndText)
//{
//	var oSelect=window.document.all[sName];
//	
//	if (sList!=undefined)
//	{
//		// ATN 23/1/2003 - use regular expressions to convert lists into SELECT OPTIONS
//		sOptions = sList.replace(/\|/g, '</option><option value=');
//		sOptions = sOptions.replace(/\`\d+/g, '');
//		sOptions =  sOptions.replace(/`/g, '>');
//
//		// ATN 23/1/2003 - replace HTML directly
//		if (bClear)
//		{
//			oSelect.innerHTML = '';
//			oSelect.outerHTML = oSelect.outerHTML.replace(/\<\/SELECT\>/i,'<option value=')  +  sOptions + "</select>";
//		}
//		else
//		{
//			oSelect.outerHTML = oSelect.outerHTML.replace(/\<\/SELECT\>/i,'<option value=') + oSelect.innerHTML +  sOptions  + "</select>";
//		}
//	}
//	else
//	{
//		oSelect.innerHTML = '';
//		alert("Menu select list file missing for '" + oSelect.name + "'. Please generate HTML in MACRO System Admin")
//	}
//}
function fnLoadSelect(sName,sList,bInit)
{
	var oS;
	oS=fnPopSelect("fltSt");
	var sStudy=oS.getText();
	var sMessage="";
	sMessage+="MACRO was unable to load 'Search' select list options for study '"+sStudy+"' into the left-hand menu pane.\n\n";
	sMessage+="The files required for this operation may not have been successfully generated through MACRO System Admin,\n"; 
	sMessage+="or old versions may be cached on your server.\nSee documentation on configuring your MACRO server for further ";
	sMessage+="information about this problem."

	oS=fnPopSelect(sName);	
	if (oS!=null)
	{
		oS.populate(sList,bInit);
		if (sList==undefined)
		{
			if(sStudy==="")
			{
				//no studies
			}
			else
			{
				if(!bShownGenHTMLMessage)
				{
					alert(sMessage)
					bShownGenHTMLMessage=true;
				}
			}
		}
	}
	else
	{
		alert("Select list "+sName+" not found");
	}
}
//function reloads study specific selects with newly selected studys items
function fnLoadStudy(lId)
{
	fnLoadSelect("fltVi",lstVisits[lId],0);
	fnLoadSelect("fltEf",lstEForms[lId],0);
	fnLoadSelect("fltQu",lstQuestions[lId],1);
}
function fnPageLoadedWin()
{
	fnInitialiseMenu("divMenu");
	fnInitialiseButton("divSearchBtn","clsHoverButton clsHoverButtonActive","clsHoverButton clsHoverButtonInactive","clsHoverButton clsHoverButtonSelected",true);
	fnSpaceMenus();
}
//check/uncheck option checkboxes - optFunctionKeys,optSymbols,optDateFormat,optSplitScreen,optSameForm
function fnSetOptionChecked(sName,bChecked)
{
	if (document.Form1.all[sName]!=undefined)
	{
		document.Form1.all[sName].checked=bChecked;
	}
}
//reloads subject quicklist - passed delimited list of delimited subjects:
//studyid|studyname|site|subjectid|subjectlabel`studyid|studyname|site|subjectid|subjectlabel
function fnReloadQuickList(sList)
{
	if((sList=="")||(sList==undefined)) return;
	var aSubjects=sList.split(sDel2);
	var aSubject;
	var sLabel;
	var osHTML=new Array();

	osHTML.push("<table id='tsubject' style='cursor:hand;' onmouseover='fnOnMouseOver(this,0);' onmouseout='fnOnMouseOut(this)' width='100%' class='clsTableText'>");
    for(var n=0;n<aSubjects.length;n++)
    {
		aSubject=aSubjects[n].split(sDel1);
        sLabel=(aSubject[4]!="")?aSubject[4]:"("+aSubject[3]+")";
        if (n%2)
        {
			osHTML.push("<tr height='10'>");
		}
		else
		{
			osHTML.push("<tr class='clsTableTextS' height='10'>");
		}
		osHTML.push("<a style='text-decoration:none;' href='javascript:fnScheduleUrl(\""+aSubject[0]+"\",\""+aSubject[2]+"\",\""+aSubject[3]+"\");'>");
		osHTML.push("<td>");
		osHTML.push(aSubject[1]+"/"+aSubject[2]+"/"+sLabel);
		osHTML.push("</td>");
		osHTML.push("</a>");
		osHTML.push("</tr>");            

	}
    osHTML.push("</table>");
	document.all["divQL"].innerHTML=osHTML.join('');
	//fnInitZOrder must be called to sort select list layers after innerhtml change
	fnInitZOrder();
}
function fnAddToQuickList(sItem)
{
	if((lstSubjects=="")||(lstSubjects==undefined))
	{
		lstSubjects=sItem;
	}
	else
	{
		lstSubjects+=sDel2+sItem;
	}		
	fnReloadQuickList(lstSubjects);
}
//sets date format label
function fnSetDateFormatLabel(sLabel)
{
	sLabel=(sLabel!="")?"("+sLabel+")":"";
	document.all["divDateLabel"].innerHTML=sLabel;
	fnInitZOrder();
}
//displays update password window
function fnChangePassword(sUser)
{
	var sRtn=window.showModalDialog('DialogChangePassword.asp?exp=&name='+sUser,'','dialogHeight:300px; dialogWidth:500px; center:yes; status:0; dependent,scrollbars');
}
//displays db lock admin window
function fnReleaseLocks()
{
	var sRtn=window.showModalDialog('DialogLockAdmin.asp','','dialogHeight:400px; dialogWidth:700px; center:yes; status:0; dependent,scrollbars');
}
//function updates task list item counter text
//sName arg is <td> id tag without 'td' prefix
function fnSetTaskListItemCounter(sName,sNum,bZ)
{
	bZ=(bZ!=undefined)?bZ:false;
	var oTD=window.document.all["td"+sName]
	if (oTD==undefined) return;
	var str=oTD.innerHTML;
	var sN = /\s\([0-9]*\)/
		
	if (sN.exec(str)!=null);
	{
		str = str.replace(sN," ("+sNum+")");
	}
	oTD.innerHTML=str;
	if (bZ) fnInitZOrder();
}
//function enables/disables tasklist items
//sName arg is <td> id tag without 'td' prefix
var aJS = new Array();
function fnEnableTaskListItem(sName,bEnable)
{
	var oTD=window.document.all["td"+sName]
	if (oTD==undefined) return;
	var str=oTD.innerHTML;
	var sStartA = /<[Aa][^>]*>/
	var sEndA = /<\/[Aa]>/
	
	if (bEnable)
	{
		if ((sStartA.exec(str)==null)&&(aJS[sName]!=undefined))
		{
			oTD.innerHTML=aJS[sName]+str+"</a>";
		}
	}
	else
	{
		if (sStartA.exec(str)!=null)
		{
			aJS[sName]=sStartA.exec(str);
			str = str.replace(sStartA,"");
			str = str.replace(sEndA,"");
			oTD.innerHTML=str;
		}
	}
	//ic 08/12/2003 performance fix - this call may need to be reinstated
	//if layer problems occur
	//fnInitZOrder();
}
function fnShowLegend(bSymbol,bFunction)
{
	var nS=(bSymbol)?70:0;
	var nF=(bFunction)?32:0;
	window.parent.document.all['f1'].rows='27,*,'+nS+','+nF;
}

function fnRegisterSubject()
{
	var sState=oMain.fnGetAppState(0)
	if ((sState!=undefined)&&((sState.substring(0,1)=="4")||(sState.substring(0,1)=="5")))
	{
		//schedule(4) or eform(5) is loaded
		if (confirm("This subject previously failed to register successfully.\nWould you like to try registration again?"))
		{
			var sRtn=window.showModalDialog('DialogRegister.asp?state='+sState,'','dialogHeight:300px; dialogWidth:500px; center:yes; status:0; dependent,scrollbars');
			fnEnableTaskListItem("RS",false);
		}
	}
}

//try for 10 seconds to initialise main frame in case lh frame loaded before it did
var nTries=0;
function fnInitMain(bSplit)
{
	if((oMain.fnShowWin==undefined)||(oMain.fnShowWin==null))
	{
		if(nTries<10)
		{
			nTries++
			window.setTimeout("fnInitMain("+bSplit+");",1000)
		}
	}
	else
	{	
		oMain.fnShowWin(bSplit); 
		oMain.fnSetAppState();
		oMain.fnResize();
	}
}