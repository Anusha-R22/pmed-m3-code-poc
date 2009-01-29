var oU;
var oS;
var sDel1="`";
var sDel2="|";
var gsETitle="";
var gsVTitle="";

function fnExpand(sHTML)
{
	var sDELIMITER = '\\|';
	var sIMAGE_START = '!F';
	var sIMAGE_PLANNED_SDV_START = "!SP";
	var sIMAGE_QUERIED_SDV_START = "!SQ";
	var sIMAGE_DONE_SDV_START = "!SD";
	var sIMAGE_END = "F->";
	var sTABLE_START = "!T";
	var sTABLE_END = "T->";
	var sCELL_START = "<!--C";
	var sCELL_END = "-->";
	var sTERM = "([^!]*)";
	var sEFORM_HEADER_START = "<!--E";
	var sBLANK_EFORM_HEADER = "<!--B-->";
	var sEFORM_HEADER_END = "-->";
	var sVISIT_START = "<!--V";
	var sVISIT_END = "-->";
	var sVISIT_TITLE_START = "<!--A";
	var sVISIT_TITLE_END = "-->";
	var sV = "<!--D-->";
	
	sHTML = sHTML.replace(new RegExp(sV,'g'),'<td valign=top align=center>');
	
	sHTML = sHTML.replace(new RegExp(sVISIT_TITLE_START + sTERM + sDELIMITER + sTERM + sDELIMITER + sTERM + sVISIT_TITLE_END,'g'),'<a onMouseup=\'javascript:$1\'><td style="cursor:default" title="$2">&nbsp;$3&nbsp;</td></a>');
	sHTML = sHTML.replace(new RegExp(sVISIT_START + sTERM + sDELIMITER + sTERM + sDELIMITER + sTERM + sVISIT_END,'g'),'<table><tr><td align=center class="clsScheduleBorder clsScheduleVisitDateText" height=20 title="Please open a form to edit the visit date">$1</td></tr><tr><a onMouseup=\'javascript:$2\'><td align=center style="cursor:default">$3</td></a></tr></table>');
	
	sHTML = sHTML.replace(new RegExp(sEFORM_HEADER_START + sTERM + sDELIMITER + sTERM + sEFORM_HEADER_END,'g'),'<tr valign=top align=center><td class=clsScheduleBorder title="$1"><div class="clsScheduleEformBorder clsScheduleEformText">$2</div></td>');
	sHTML = sHTML.replace(new RegExp(sBLANK_EFORM_HEADER,'g'),'<tr valign=top align=center><td class=clsScheduleBorder></td>');
	
	sHTML = sHTML.replace(new RegExp(sIMAGE_PLANNED_SDV_START + sTERM + sDELIMITER + sTERM + sIMAGE_END,'g'),'<table cellpadding=0 cellspacing=0><tr><td><img border=0 style="cursor:hand" src="../img/$1.gif" alt="$2"></td></tr><tr><td><img border=0 src="../img/icof_sdv_plan.gif"></td></tr></table>');
	sHTML = sHTML.replace(new RegExp(sIMAGE_QUERIED_SDV_START + sTERM + sDELIMITER + sTERM + sIMAGE_END,'g'),'<table cellpadding=0 cellspacing=0><tr><td><img border=0 style="cursor:hand" src="../img/$1.gif" alt="$2"></td></tr><tr><td><img border=0 src="../img/icof_sdv_query.gif"></td></tr></table>');
	sHTML = sHTML.replace(new RegExp(sIMAGE_DONE_SDV_START + sTERM + sDELIMITER + sTERM + sIMAGE_END,'g'),'<table cellpadding=0 cellspacing=0><tr><td><img border=0 style="cursor:hand" src="../img/$1.gif" alt="$2"></td></tr><tr><td><img border=0 src="../img/icof_sdv_done.gif"></td></tr></table>');
	
	sHTML = sHTML.replace(new RegExp(sIMAGE_START + sTERM + sDELIMITER + sTERM + sIMAGE_END,'g'),'<img border=0 style="cursor:hand" src="../img/$1.gif" alt="$2">');
	
	sHTML=sHTML.replace(new RegExp(sTABLE_START + sTERM + sDELIMITER + sTERM + sDELIMITER + sTERM + sDELIMITER + sTERM + sTABLE_END,'g'),'<table><tr><a onMouseup=\'javascript:$1\'><td>$2</td></a></tr></table>$3<br>$4<br>');

	sHTML = sHTML.replace(new RegExp(sCELL_START + sTERM + sDELIMITER + sTERM + sCELL_END,'g'),'<td align=center valign=top class=clsScheduleEformLabelText bgcolor=$1>$2</td>');
	
	return sHTML;	
}

var scrollLeft=0, scrollTop=0;

function ShowHeaders(){
	if(scrollLeft==document.body.scrollLeft&&scrollTop==document.body.scrollTop){
		document.all.top.style.visibility='visible';
		rowhead.style.visibility='visible';
		colhead.style.visibility='visible';
		document.all.blank.style.visibility='visible';
	}
	scrollLeft=document.body.scrollLeft;
	scrollTop=document.body.scrollTop;
}

function SizeHeaders(){
var i, j, nHeightCum=0;
var oMain=document.all.main;
var nMainLength=oMain.rows.length;
var oCol=document.all.colhead;
var oRow=document.all.rowhead;
var nLength,oStyle,oCell,nTop,nHeight,osHTML;

	for(i=0;i<nMainLength;i++)
	{
		if(oMain.rows(i).id=='head')
		{
			nLength=oMain.rows(i).cells.length;
			nTop=oCol.rows(i).cells(0).style.offsetTop;
			nHeight=oCol.rows(i).cells(0).style.clientHeight;
			
			for(j=0;j<nLength;j++)
			{
				oStyle=oCol.rows(i).cells(j).style
				oCell=oMain.rows(i).cells(j)
			
				oStyle.position='absolute';
				oStyle.pixelLeft=oCell.offsetLeft;
				oStyle.pixelTop=oCell.offsetTop;
				oStyle.pixelHeight=oCell.clientHeight;
				oStyle.pixelWidth=oCell.clientWidth;
			}
			nHeightCum+=oCol.rows(i).cells(0).style.pixelHeight;
		}
		oStyle=oRow.rows(i).cells(0).style;
		oCell=oMain.rows(i).cells(0);
		
		oStyle.position='absolute';
		oStyle.pixelLeft=oCell.offsetLeft;
		oStyle.pixelTop=oCell.offsetTop;
		oStyle.pixelHeight=oCell.clientHeight;
		oStyle.pixelWidth=oCell.clientWidth;
	}
	document.all.blank.style.pixelWidth=document.all.rowhead.rows(0).cells(0).style.pixelWidth;
	document.all.blank.style.pixelHeight=nHeightCum;
	PositionHeaders();	
}

function PositionHeaders(){
	rowhead.style.pixelLeft=document.body.scrollLeft + document.all.main.offsetLeft;
	colhead.style.pixelTop = document.body.scrollTop + document.all.main.offsetTop;
	document.all.top.style.pixelLeft=document.all.rowhead.style.pixelLeft;
	document.all.top.style.pixelTop=document.body.scrollTop;
	document.all.blank.style.pixelLeft=document.all.rowhead.style.pixelLeft;
	document.all.blank.style.pixelTop=document.all.colhead.style.pixelTop;

	document.all.top.style.visibility='hidden';
	rowhead.style.visibility='hidden';
	colhead.style.visibility='hidden';
	document.all.blank.style.visibility='hidden';
}

function DrawHeaders(){
var i,j,oRow;
var osCol=new Array();
var osRow=new Array();
var oMain=document.all.main;
var nMainLength=oMain.rows.length;
var sStart=/<[TR][^>]*>/;
	
	osRow.push(document.all.rowhead.outerHTML.replace(/<\/table>/i,''));
	osCol.push(document.all.colhead.outerHTML.replace(/<\/table>/i,''));
	
	osCol.push(oMain.rows(0).outerHTML);
	osCol.push(oMain.rows(1).outerHTML);
	osCol.push(oMain.rows(2).outerHTML);
	osCol.push(oMain.rows(3).outerHTML);
	
	for(i=0;i<nMainLength;i++)
	{
		oRow=oMain.rows(i);			
		osRow.push(sStart.exec(oRow.outerHTML));
		osRow.push(oRow.cells(0).outerHTML);
		osRow.push('</tr>');
	}
	
	osCol.push('</table>');
	osRow.push('</table>');
	
	colhead.outerHTML=osCol.join('');
	colhead.style.pixelWidth=oMain.clientWidth;
	
	rowhead.outerHTML=osRow.join('');
	rowhead.style.pixelWidth=oMain.rows(0).cells(0).clientWidth;
	rowhead.style.pixelHeight=oMain.clientHeight;
	rowhead.style.pixelTop=document.all.top.clientHeight;
	
	document.all.top.style.position='absolute';
	document.all.blank.style.position='absolute';
	document.all.top.style.zIndex=101;

	SizeHeaders();
	ShowHeaders();
	window.setInterval('ShowHeaders();',200);
}

function fnClose()
{
	window.parent.fnHomeUrl();
}
function fnInitUser(bViewData,bChangeData,bCreateSDV,bLockData,
					bFreezeData,bUnFreezeData)
{
	oU=new Object();
	oU.bViewData=bViewData;
	oU.bChangeData=bChangeData;
	oU.bCreateSDV=bCreateSDV;
	oU.bLockData=bLockData;
	oU.bFreezeData=bFreezeData;
	oU.bUnFreezeData=bUnFreezeData;
}
function fnInitSubject(sStudy,sSite,sSubject,sLabel)
{
	oS=new Object();
	oS.sStudy=sStudy;
	oS.sSite=sSite;
	oS.sSubject=sSubject;
	oS.Label=sLabel;
}

/*<!--r-->*/

window.onresize=SizeHeaders;

function fnS(nBtn,nStatus,bLocked,bFrozen,bCanUnFreeze,bHasSDV)
{
	var sHtml="";
	sHtml+="<table width='130'>";
	var bUnLocked=((!bLocked)&&(!bFrozen));
	var nLFScope=1;
	var nLFAction=0;
    var bSep=false;
    
    if (nBtn==1) return;
    
	if (oU.bLockData)
	{
		if(bUnLocked)
		{
			//Lock
			var sCompletid=oS.sStudy+sDel1+oS.sSite+sDel1+oS.sSubject+sDel1+sDel1+sDel1+sDel1;
			nLFAction=1;
			sCompletid+=sDel1+nLFScope+sDel1+nLFAction;
			sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>";
			sHtml+="<a href='javascript:fnPopMenuClick(\""+sCompletid+"\",5);'>Lock</a>";
			sHtml+="</td></tr>";
		}
		else
		{
			sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>Lock</td></tr>";
		}
		bSep=true;
	}
	if (oU.bLockData)
	{
		if(bLocked)
		{
			//Unlock
			var sCompletid=oS.sStudy+sDel1+oS.sSite+sDel1+oS.sSubject+sDel1+sDel1+sDel1+sDel1;
			nLFAction=0;
			sCompletid+=sDel1+nLFScope+sDel1+nLFAction;
			sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>";
			sHtml+="<a href='javascript:fnPopMenuClick(\""+sCompletid+"\",5);'>Unlock</a>";
			sHtml+="</td></tr>";
		}
		else
		{
			sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>Unlock</td></tr>";
		}
		bSep=true;
	}
	if (oU.bFreezeData)
	{
		if(!bFrozen)
		{
			//Freeze
			var sCompletid=oS.sStudy+sDel1+oS.sSite+sDel1+oS.sSubject+sDel1+sDel1+sDel1+sDel1;
			nLFAction=2;
			sCompletid+=sDel1+nLFScope+sDel1+nLFAction;
			sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>";
			sHtml+="<a href='javascript:fnPopMenuClick(\""+sCompletid+"\",5);'>Freeze</a>";
			sHtml+="</td></tr>";
		}
		else
		{
			sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>Freeze</td></tr>";
		}
		bSep=true;
	}
	if (oU.bUnFreezeData)
	{
		if(bCanUnFreeze)
		{
			//UnFreeze
			var sCompletid=oS.sStudy+sDel1+oS.sSite+sDel1+oS.sSubject+sDel1+sDel1+sDel1+sDel1;
			nLFAction=3;
			sCompletid+=sDel1+nLFScope+sDel1+nLFAction;
			sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>";
			sHtml+="<a href='javascript:fnPopMenuClick(\""+sCompletid+"\",5);'>UnFreeze</a>";
			sHtml+="</td></tr>";
		}
		else
		{
			sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>UnFreeze</td></tr>";
		}
		bSep=true;
	}
	if(bSep)
	{	
		sHtml+="<tr height='1'><td align=left width='98%' bgcolor='#C0C0C0'></td></tr>";
	}
	if (oU.bChangeData)
	{
		var sCompletid=oS.sStudy+sDel1+oS.sSite+sDel1+oS.sSubject+sDel1+sDel1+sDel1+sDel1+sDel1+"s";
		sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>";
		sHtml+=(((nStatus=="10")||(nStatus=="-10"))&&(!oS.bReadOnly)&&(bUnLocked))? "<a href='javascript:fnPopMenuClick(\""+sCompletid+"\",0);'>Unobtainable</a>":"Unobtainable";
		sHtml+="</td></tr>";
			
		sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>";
		sHtml+=((nStatus=="-5")&&(!oS.bReadOnly)&&(bUnLocked))? "<a href='javascript:fnPopMenuClick(\""+sCompletid+"\",1);'>Missing</a>":"Missing";
		sHtml+="</td></tr>";
		sHtml+="<tr height='1'><td align=left width='98%' bgcolor='#C0C0C0'></td></tr>";
	}
    if (oU.bCreateSDV)
	{
		sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>";
		if(bHasSDV)
		{
			var sCompletid=oS.sStudy+sDel1+oS.sSite+sDel1+oS.sSubject+sDel1+sDel1+sDel1+sDel1+sDel1+"s";
			sHtml+="<a href='javascript:fnPopMenuClick(\""+sCompletid+"\",7);'>Edit Subject SDV...</a>";
		}
		else
		{
			var sCompletid=oS.sStudy+sDel1+oS.sSite+sDel1+oS.sSubject+sDel1+sDel1+sDel1+sDel1;
			sHtml+="<a href='javascript:fnPopMenuClick(\""+sCompletid+"\",4);'>Create Subject SDV...</a>";
		}
		sHtml+="</td></tr>";
    }
	sHtml+="</table>";
	document.all['divPopMenu'].innerHTML=sHtml;
	fnPopMenuShow(document.all['divPopMenu'])
}

function fnV(nBtn,sViId,sViCycle,sViTitle,nStatus,bLocked,bFrozen,bCanUnFreeze,bHasSDV)
{
	gsVTitle=sViTitle;
	var bUnLocked=((!bLocked)&&(!bFrozen));
	var nLFScope=2;
	var nLFAction=0;
	var sHtml="";
	sHtml+="<table width='130'>";
    var bSep=false;
    
    if (nBtn==1) return;
    
	if (oU.bLockData)
	{
		if(bUnLocked)
		{
			//Lock
			var sCompletid=oS.sStudy+sDel1+oS.sSite+sDel1+oS.sSubject+sDel1+sViId+sDel1+sViCycle+sDel1+sDel1;
			nLFAction=1;
			sCompletid+=sDel1+nLFScope+sDel1+nLFAction;
			sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>";
			sHtml+="<a href='javascript:fnPopMenuClick(\""+sCompletid+"\",5);'>Lock</a>";
			sHtml+="</td></tr>";
		}
		else
		{
			sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>Lock</td></tr>";
		}
		bSep=true;
	}
	if (oU.bLockData)
	{
		if(bLocked)
		{
			//Unlock
			var sCompletid=oS.sStudy+sDel1+oS.sSite+sDel1+oS.sSubject+sDel1+sViId+sDel1+sViCycle+sDel1+sDel1;
			nLFAction=0;
			sCompletid+=sDel1+nLFScope+sDel1+nLFAction;
			sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>";
			sHtml+="<a href='javascript:fnPopMenuClick(\""+sCompletid+"\",5);'>Unlock</a>";
			sHtml+="</td></tr>";
		}
		else
		{
			sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>Unlock</td></tr>";
		}
		bSep=true;
	}
	if (oU.bFreezeData)
	{
		if(!bFrozen)
		{
			//Freeze
			var sCompletid=oS.sStudy+sDel1+oS.sSite+sDel1+oS.sSubject+sDel1+sViId+sDel1+sViCycle+sDel1+sDel1;
			nLFAction=2;
			sCompletid+=sDel1+nLFScope+sDel1+nLFAction;
			sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>";
			sHtml+="<a href='javascript:fnPopMenuClick(\""+sCompletid+"\",5);'>Freeze</a>";
			sHtml+="</td></tr>";
		}
		else
		{
			sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>Freeze</td></tr>";
		}
		bSep=true;
	}
	if (oU.bUnFreezeData)
	{
		if(bCanUnFreeze)
		{
			//UnFreeze
			var sCompletid=oS.sStudy+sDel1+oS.sSite+sDel1+oS.sSubject+sDel1+sViId+sDel1+sViCycle+sDel1+sDel1;
			nLFAction=3;
			sCompletid+=sDel1+nLFScope+sDel1+nLFAction;
			sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>";
			sHtml+="<a href='javascript:fnPopMenuClick(\""+sCompletid+"\",5);'>UnFreeze</a>";
			sHtml+="</td></tr>";
		}
		else
		{
			sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>UnFreeze</td></tr>";
		}
		bSep=true;
	}
	if(bSep)
	{
		sHtml+="<tr height='1'><td align=left width='98%' bgcolor='#C0C0C0'></td></tr>";
	}
	if (oU.bChangeData)
	{
		var sCompletid=oS.sStudy+sDel1+oS.sSite+sDel1+oS.sSubject+sDel1+sViId+sDel1+sViCycle+sDel1+sDel1+sDel1+"v";
		sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>";
		sHtml+=(((nStatus=="10")||(nStatus=="-10"))&&(!oS.bReadOnly)&&(bUnLocked))? "<a href='javascript:fnPopMenuClick(\""+sCompletid+"\",0);'>Unobtainable</a>":"Unobtainable";
		sHtml+="</td></tr>";
			
		sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>";
		sHtml+=((nStatus=="-5")&&(!oS.bReadOnly)&&(bUnLocked))? "<a href='javascript:fnPopMenuClick(\""+sCompletid+"\",1);'>Missing</a>":"Missing";
		sHtml+="</td></tr>";
		sHtml+="<tr height='1'><td align=left width='98%' bgcolor='#C0C0C0'></td></tr>";
	}
	if (oU.bCreateSDV)
	{
		var sCompletid=oS.sStudy+sDel1+oS.sSite+sDel1+oS.sSubject+sDel1+sViId+sDel1+sViCycle;
		sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>";
		if(bHasSDV)
		{
			sCompletid+=sDel1+sDel1+sDel1+"v";
			sHtml+="<a href='javascript:fnPopMenuClick(\""+sCompletid+"\",7);'>Edit Visit SDV...</a>";
		}
		else
		{
			sHtml+="<a href='javascript:fnPopMenuClick(\""+sCompletid+"\",3);'>Create Visit SDV...</a>";
		}
		sHtml+="</td></tr>";
    }
    sHtml+="</table>";
	document.all['divPopMenu'].innerHTML=sHtml;
	fnPopMenuShow(document.all['divPopMenu'])
}

function fnE(nBtn,sEfTskId,sEfId,sEfCycle,sEfTitle,sViId,sViCycle,sViTitle,nStatus,bLocked,bFrozen,bCanUnFreeze,bSetPlannedToDone,bSHasSDV,bVHasSDV,bEHasSDV)
{
	gsETitle=sEfTitle;
	gsVTitle=sViTitle;

	if (nBtn==1)
    {
		//if(CanOpeneForm(nStatus,bLocked,bFrozen))
		if (oU.bViewData)
		{
			if((nStatus==-10)&&(bLocked||bFrozen))
			{
				alert("This locked eForm contains no data and cannot be opened");
			}
			else if((nStatus==-10)&&(!oU.bChangeData))
			{
				alert("You may not enter new data for this subject");
			}
			else
			{
				window.parent.fnEformUrl(oS.sStudy,oS.sSite,oS.sSubject,sEfTskId,"",undefined,window.sWinState);
			}
		}
		else
		{
			alert("You do not have permission to view data");
		}
    }
    else
    {
		var bUnLocked=((!bLocked)&&(!bFrozen));
		var nLFScope=3;
		var nLFAction=0;
		var sHtml="";
        sHtml+="<table width='130'>";
        var bSep=false;
        
        if (oU.bViewData)
        {
			if(((nStatus==-10)&&(bLocked||bFrozen))||((nStatus==-10)&&(!oU.bChangeData)))
			{
				sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>Open...</td></tr>";
				sHtml+="<tr><td align=left><hr width='98%' size='1' color='#C0C0C0'></hr></td></tr>";
			}
			else
			{
				sHtml+="<tr height='15'><td class='clsPopMenuLinkText'><a href='javascript:fnE(1,\""+sEfTskId+"\");'><b>Open...</b></a></td></tr>";
				sHtml+="<tr height='1'><td align=left width='98%' bgcolor='#C0C0C0'></td></tr>";
			}
		}
		if (oU.bLockData)
		{
			if(bUnLocked)
			{
				//Lock
				var sCompletid=oS.sStudy+sDel1+oS.sSite+sDel1+oS.sSubject+sDel1+sViId+sDel1+sViCycle+sDel1+sEfId+sDel1+sEfCycle;
				nLFAction=1;
				sCompletid+=sDel1+nLFScope+sDel1+nLFAction;
				sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>";
				sHtml+="<a href='javascript:fnPopMenuClick(\""+sCompletid+"\",5);'>Lock</a>";
				sHtml+="</td></tr>";
			}
			else
			{
				sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>Lock</td></tr>";
			}
			bSep=true;
		}
		if (oU.bLockData)
		{
			if(bLocked)
			{
				//Unlock
				var sCompletid=oS.sStudy+sDel1+oS.sSite+sDel1+oS.sSubject+sDel1+sViId+sDel1+sViCycle+sDel1+sEfId+sDel1+sEfCycle;
				nLFAction=0;
				sCompletid+=sDel1+nLFScope+sDel1+nLFAction;
				sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>";
				sHtml+="<a href='javascript:fnPopMenuClick(\""+sCompletid+"\",5);'>Unlock</a>";
				sHtml+="</td></tr>";
			}
			else
			{
				sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>Unlock</td></tr>";
			}
			bSep=true;
		}
		if (oU.bFreezeData)
		{
			if(!bFrozen)
			{
				//Freeze
				var sCompletid=oS.sStudy+sDel1+oS.sSite+sDel1+oS.sSubject+sDel1+sViId+sDel1+sViCycle+sDel1+sEfId+sDel1+sEfCycle;
				nLFAction=2;
				sCompletid+=sDel1+nLFScope+sDel1+nLFAction;
				sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>";
				sHtml+="<a href='javascript:fnPopMenuClick(\""+sCompletid+"\",5);'>Freeze</a>";
				sHtml+="</td></tr>";
			}
			else
			{
				sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>Freeze</td></tr>";
			}
			bSep=true;
		}
		if (oU.bUnFreezeData)
		{
			if(bCanUnFreeze)
			{
				//UnFreeze
				var sCompletid=oS.sStudy+sDel1+oS.sSite+sDel1+oS.sSubject+sDel1+sViId+sDel1+sViCycle+sDel1+sEfId+sDel1+sEfCycle;
				nLFAction=3;
				sCompletid+=sDel1+nLFScope+sDel1+nLFAction;
				sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>";
				sHtml+="<a href='javascript:fnPopMenuClick(\""+sCompletid+"\",5);'>UnFreeze</a>";
				sHtml+="</td></tr>";
			}
			else
			{
				sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>UnFreeze</td></tr>";
			}
			bSep=true;
		}
		if(bSep)
		{
			sHtml+="<tr height='1'><td align=left width='98%' bgcolor='#C0C0C0'></td></tr>";
		}
		if (oU.bChangeData)
		{
			var sCompletid=oS.sStudy+sDel1+oS.sSite+sDel1+oS.sSubject+sDel1+sViId+sDel1+sViCycle+sDel1+sDel1+sEfTskId+sDel1+"e";
			sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>";
			sHtml+=(((nStatus=="10")||(nStatus=="-10"))&&(!oS.bReadOnly)&&(bUnLocked))? "<a href='javascript:fnPopMenuClick(\""+sCompletid+"\",0);'>Unobtainable</a>":"Unobtainable";
			sHtml+="</td></tr>";
			
			sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>";
			sHtml+=((nStatus=="-5")&&(!oS.bReadOnly)&&(bUnLocked))? "<a href='javascript:fnPopMenuClick(\""+sCompletid+"\",1);'>Missing</a>":"Missing";
			sHtml+="</td></tr>";
			sHtml+="<tr height='1'><td align=left width='98%' bgcolor='#C0C0C0'></td></tr>";
		}
		if (oU.bCreateSDV)
		{
			//if (bUnLocked)
			//{
				var sCompletid=oS.sStudy+sDel1+oS.sSite+sDel1+oS.sSubject+sDel1+sViId+sDel1+sViCycle+sDel1+sDel1+sDel1+sDel1+sEfTskId;
				sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>";
				if(bEHasSDV)
				{
					sCompletid=oS.sStudy+sDel1+oS.sSite+sDel1+oS.sSubject+sDel1+sDel1+sDel1+sEfId+sDel1+sEfCycle+sDel1+"e";
					sHtml+="<a href='javascript:fnPopMenuClick(\""+sCompletid+"\",7);'>Edit EForm SDV...</a>";
				}
				else
				{
					sHtml+="<a href='javascript:fnPopMenuClick(\""+sCompletid+"\",2);'>Create EForm SDV...</a>";
				}
				sHtml+="</td></tr>";
			//}
			//else
			//{
			//	sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>Create EForm SDV...</td></tr>";
			//}
			
			if(bShowSDVScheduleMenu)
			{
				sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>";
				if(bVHasSDV)
				{
					sCompletid=oS.sStudy+sDel1+oS.sSite+sDel1+oS.sSubject+sDel1+sViId+sDel1+sViCycle+sDel1+sDel1+sDel1+"v";
					sHtml+="<a href='javascript:fnPopMenuClick(\""+sCompletid+"\",7);'>Edit Visit SDV...</a>";
				}
				else
				{
					sHtml+="<a href='javascript:fnPopMenuClick(\""+sCompletid+"\",3);'>Create Visit SDV...</a>";
				}
				sHtml+="</td></tr>";
				
				sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>";
				if(bSHasSDV)
				{
					sCompletid=oS.sStudy+sDel1+oS.sSite+sDel1+oS.sSubject+sDel1+sDel1+sDel1+sDel1+sDel1+"s";
					sHtml+="<a href='javascript:fnPopMenuClick(\""+sCompletid+"\",7);'>Edit Subject SDV...</a>";
				}
				else
				{
					sHtml+="<a href='javascript:fnPopMenuClick(\""+sCompletid+"\",4);'>Create Subject SDV...</a>";
				}
				sHtml+="</td></tr>";
			}
		}
		sHtml+="<tr height='1'><td align=left width='98%' bgcolor='#C0C0C0'></td></tr>";
		
		if(bSetPlannedToDone)
		{
			var sCompletid=oS.sStudy+sDel1+oS.sSite+sDel1+oS.sSubject+sDel1+sDel1+sDel1+sDel1+sDel1+sDel1+sEfTskId;
			sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>";
			sHtml+="<a href='javascript:fnPopMenuClick(\""+sCompletid+"\",6);'>Change all Planned question SDVs to Done</a>";
			sHtml+="</td></tr>";
		}
		else
		{
			sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>Change all Planned question SDVs to Done</td></tr>";
		}
		
        sHtml+="</table>";
        
        document.all['divPopMenu'].innerHTML=sHtml;
        fnPopMenuShow(document.all['divPopMenu']);
    }
}
                        
function fnPopMenuClick(id,nAct,sTitle)
{
	switch (nAct)
	{
		case 0:
			//unobtainable
			var sText="";
			switch (id.charAt(id.length-1))
			{
				case "e":
					sText="eForm";
					break;
				case "v":
					sText="Visit";
					break;
				case "s":
					sText="Subject";
					break;
			}
			if (confirm("Are you sure you want to change this " + sText + " to unobtainable?"))
			{
				sRtn="u";
			}
			break;
		case 1:
			//missing
			var sText="";
			switch (id.charAt(id.length-1))
			{
				case "e":
					sText="eForm";
					break;
				case "v":
					sText="Visit";
					break;
				case "s":
					sText="Subject";
					break;
			}
			if (confirm("Are you sure you want to change this " + sText + " to missing?"))
			{
				sRtn="m";
			}
			break;
		case 2:
		case 3:
		case 4:
			var aArgs=new Array();
			var sType="";
			switch (nAct)
			{
				case 2: sType="e";aArgs[0]=gsETitle;break;
				case 3: sType="v";aArgs[0]=gsVTitle;break;
				default: sType="s";aArgs[0]=oS.Label+"("+oS.sSubject+")";break;
			}
			aArgs[1]="";
			
			//eform sdv
			var sRtn=window.showModalDialog('DialogNewSDV.htm',aArgs,'dialogHeight:300px;dialogWidth:500px;center:yes;status:0;dependent,scrollbars');
			sRtn+=((sRtn!=undefined)&&(sRtn!=""))? sDel1+sType:"";
			break;
		case 5:
			//subject/visit/eForm lock
			var sRtn="l";
			break;
		case 6:
			if (confirm("Are you sure you want to change all planned SDVs to done?"))
			{
				sRtn="d";
			}
			break;
		case 7:
			//edit sdv
			var aParams = id.split(sDel1);
			var sScope="";
			switch(aParams[7])
			{
				case "s": sScope="1000";break;
				case "v": sScope="0100";break;
				case "e": sScope="0010";break;
			}
			window.parent.fnMIMessageUrl("1",aParams[0],aParams[1],aParams[3],aParams[5],'',aParams[2],'','false','','1111',sScope,'0',undefined,aParams[4],aParams[6],"");
	}
	if ((sRtn!=undefined)&&(sRtn!=""))
	{
		window.document.Form1.SchedUpdate.value=sRtn;
		window.document.Form1.SchedIdentifier.value=id;
		window.document.Form1.submit();
	}	
}

// Hide the loader
function fnHideLoader()
{
	document.all.divMsgBox.style.visibility='hidden';
}

function fnPopMenuHide()
{
	document.all["divPopMenu"].style.visibility='hidden'
}

function fnPopMenuShow(oDiv)
{
	var nPopHeight=oDiv.clientHeight;
	var nPopWidth=oDiv.clientWidth;
	var nPopYPosition=window.event.clientY;
	var nPopXPosition=window.event.clientX;
	var nVisibleScreenHeight=document.body.clientHeight;
	var nVisibleScreenWidth=document.body.clientWidth;
	var nX=0;
	var nY=0;
	
	if((nVisibleScreenHeight-(nPopYPosition+nPopHeight)>0)||((nPopYPosition-nPopHeight)<0))
	{
		//if space below or not space above, display below
		nY=nPopYPosition;
	}
	else
	{
		//display above
		nY=nPopYPosition-nPopHeight;
	}
	if((nVisibleScreenWidth-(nPopXPosition+nPopWidth)>0)||((nPopXPosition-nPopWidth)<0))
	{
		//if space right or not space left, display right
		nX=nPopXPosition;
	}
	else
	{
		//display left
		nX=nPopXPosition-nPopWidth;
	}
    oDiv.style.pixelLeft=document.body.scrollLeft+nX;
    oDiv.style.pixelTop=document.body.scrollTop+nY;
    oDiv.style.visibility='visible';
}
document.onclick=fnPopMenuHide;
/*<!--r-->*/