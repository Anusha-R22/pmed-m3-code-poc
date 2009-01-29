//function initialises menus
function fnInitialiseMenu()
{
	for (var n=0;n<document.all["divMenu"].length;n++)
	{
		document.all["divMenu"][n].ndx=n;
		document.all["divMenu"][n].isExpanded=false;
		document.all["divMenuPane"][n].style.visibility="hidden";
	}
}
//function sets passed menu to hovered style
function fnSetMenuHover(oDiv)
{
	if ((oDiv.isExpanded==undefined)||(oDiv.isExpanded==false))
	{
		oDiv.className="clsMenuHeader clsMenuHeaderActive";
		document.all[oDiv.id+"Img"][oDiv.ndx].src="../img/exp_active.gif";
	}
}
//function sets menu back to unhovered style
function fnSetMenuUnHover(oDiv)
{
	if ((oDiv.isExpanded==undefined)||(oDiv.isExpanded==false))
	{
		oDiv.className="clsMenuHeader clsMenuHeaderInactive";
		document.all[oDiv.id+"Img"][oDiv.ndx].src="../img/exp_inactive.gif";
	}
}
//function spaces menus based on whether they are expanded or collapsed
function fnSpaceMenus()
{
	var oDiv;
	var oDivDisplay;
	var oPrevDiv;
	var oPrevDivDisplay;
	var nTop=0;
	
	//get top of first menu header
	nTop+=document.all["divMenu"][0].offsetTop;
	//loop through menu headers, missing first and last
	for (n=1;n<=document.all["divMenu"].length-1;n++)
	{
		//get current menu header and previous menu header
		oDiv=document.all["divMenu"][n];
		oDivDisplay=document.all["divMenuPane"][n];
		oPrevDiv=document.all["divMenu"][(n-1)];
		oPrevDivDisplay=document.all["divMenuPane"][(n-1)];
		//get bottom of previous menu header
		nTop+=oPrevDiv.offsetHeight;
		if ((oPrevDiv.isExpanded!=undefined)&&(oPrevDiv.isExpanded!=false))
		{
			//if previous menu header is expanded, add height of display pane
			nTop+=oPrevDivDisplay.offsetHeight;
		}
		//add height spacer
		nTop+=10;
		//set current menu header top to calculated position
		oDiv.style.top=nTop;
		if (oDiv.isExpanded)
		{
			//set current display pane top to calculated position + height of menu header
			oDivDisplay.style.top=oDivDisplay.style.top=nTop+oDiv.offsetHeight
		}
		else
		{
			//shift display pane to top of page to avoid scrollbars appearing when pane is hidden
			oDivDisplay.style.top=oDivDisplay.style.top=0;
			if (oDiv.name=="search") 
			{
				//search pane, special case
				fnHideSearch();
				oDivDisplay.style.height=10;
			}
		}
	}	
}
//function expands menu (shows display div)
function fnExpandMenu(oDiv)
{
	oDiv.className="clsMenuHeader clsMenuHeaderActive";
	document.all["divMenuImg"][oDiv.ndx].src="../img/col_active.gif";
	document.all["divMenuPane"][oDiv.ndx].style.visibility="visible";
	oDiv.isExpanded=true;
	fnSpaceMenus();
}
//function collapses menu (hides display div)
function fnCollapseMenu(oDiv)
{
	document.all["divMenuPane"][oDiv.ndx].style.visibility="hidden";
	oDiv.className="clsMenuHeader clsMenuHeaderInactive";
	document.all["divMenuImg"][oDiv.ndx].src="../img/exp_inactive.gif";
	oDiv.isExpanded=false;
	fnSetMenuHover(oDiv);
	fnSpaceMenus();
}
//function toggles menus between collapsed and expanded
function fnToggleMenu(oDiv)
{
	if ((oDiv.isExpanded==undefined)||(oDiv.isExpanded==false))
	{
		fnExpandMenu(oDiv);
	}
	else if (oDiv.isExpanded==true)
	{
		fnCollapseMenu(oDiv);
	}
}